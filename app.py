from __future__ import annotations

import io
import os
RESEND_API_KEY = os.environ.get("RESEND_API_KEY")
RESEND_FROM = "Webull Event <noreply@eventflowpro.fun>"
import re
import secrets
import sqlite3
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from typing import Optional
from functools import wraps
import resend
from werkzeug.security import generate_password_hash, check_password_hash

import pandas as pd
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    url_for,
    session,
)

# ============================================================
# 讲座签到系统:二维码签到版完整 app.py
# ============================================================
# 功能:
# 1. 后台上传 Excel 名单
# 2. 普通手机/邮箱签到
# 3. Webull 注册开户状态记录
# 4. 未匹配用户前台现场登记
# 5. 后台手动补签
# 6. 后台临时访客登记 + 补签
# 7. 每位名单用户生成唯一二维码 token
# 8. 到场后工作人员扫码，匹配 token 后直接签到
# 9. 防止重复扫码 / 重复签到
# 10. 导出签到记录与二维码链接表
#
# requirements.txt 建议:
# flask
# pandas
# openpyxl
# gunicorn
# qrcode[pil]
#
# Render Start Command:
# gunicorn app:app
# ============================================================

try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "checkin_mini.db"

app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET_KEY", "replace-this-with-a-random-secret-key")
ADMIN_REGISTER_CODE = os.environ.get("ADMIN_REGISTER_CODE")

# 工作人员扫码口令:上线前请改成你们内部口令。
# 之后也可以改成环境变量:os.environ.get("STAFF_SCAN_PASSWORD")
STAFF_SCAN_PASSWORD = "webull-staff-2026"

# -----------------------------
# Email configuration
# -----------------------------
# For safety, do not hard-code your Gmail App Password here.
# Set EMAIL_PASS in Render -> Settings -> Environment.


# -----------------------------
# 数据库
# -----------------------------
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db() -> None:
    conn = get_conn()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS registrants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id TEXT NOT NULL,
            name TEXT,
            phone TEXT,
            email TEXT,
            organization TEXT,
            source TEXT DEFAULT 'imported',
            has_webull_account TEXT,
            webull_opened TEXT,
            checkin_token TEXT UNIQUE,
            qr_sent_at TEXT,
            qr_used_at TEXT,
            raw_json TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    # 兼容旧数据库结构:缺字段则自动补字段
    cur.execute("PRAGMA table_info(registrants)")
    registrant_cols = [row[1] for row in cur.fetchall()]
    migrations = {
        "source": "ALTER TABLE registrants ADD COLUMN source TEXT DEFAULT 'imported'",
        "has_webull_account": "ALTER TABLE registrants ADD COLUMN has_webull_account TEXT",
        "webull_opened": "ALTER TABLE registrants ADD COLUMN webull_opened TEXT",
        "checkin_token": "ALTER TABLE registrants ADD COLUMN checkin_token TEXT UNIQUE",
        "qr_sent_at": "ALTER TABLE registrants ADD COLUMN qr_sent_at TEXT",
        "qr_used_at": "ALTER TABLE registrants ADD COLUMN qr_used_at TEXT",
    }
    for col, sql in migrations.items():
        if col not in registrant_cols:
            cur.execute(sql)

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS checkins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id TEXT NOT NULL,
            registrant_id INTEGER,
            submitted_phone TEXT,
            submitted_email TEXT,
            has_webull_account TEXT,
            checkin_method TEXT DEFAULT 'form',
            status TEXT NOT NULL,
            message TEXT,
            ip TEXT,
            user_agent TEXT,
            checked_in_at TEXT NOT NULL,
            UNIQUE(event_id, registrant_id)
        )
        """
    )

    cur.execute("PRAGMA table_info(checkins)")
    checkin_cols = [row[1] for row in cur.fetchall()]
    checkin_migrations = {
        "has_webull_account": "ALTER TABLE checkins ADD COLUMN has_webull_account TEXT",
        "checkin_method": "ALTER TABLE checkins ADD COLUMN checkin_method TEXT DEFAULT 'form'",
    }
    for col, sql in checkin_migrations.items():
        if col not in checkin_cols:
            cur.execute(sql)

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS admin_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )

    conn.commit()
    conn.close()


@app.before_request
def ensure_db() -> None:
    init_db()


# -----------------------------
# 工具函数
# -----------------------------
def normalize_phone(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\D", "", str(value).strip())


def normalize_email(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def normalize_yes_no(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower()
    if s in {"yes", "y", "是", "已", "1", "true"}:
        return "是"
    if s in {"no", "n", "否", "未", "0", "false"}:
        return "否"
    return ""


def pick_column(columns: list[str], candidates: list[str]) -> Optional[str]:
    low_map = {str(c).strip().lower(): c for c in columns}
    for cand in candidates:
        if cand.lower() in low_map:
            return low_map[cand.lower()]
    for col in columns:
        low = str(col).strip().lower()
        if any(c.lower() in low for c in candidates):
            return col
    return None


def make_msg_html(msg_text: str, msg_type: str = "ok") -> str:
    if not msg_text:
        return ""
    css = "ok" if msg_type == "ok" else "err"
    return f'<div class="{css}" style="margin-bottom:16px;">{msg_text}</div>'


def get_base_url() -> str:
    # Render / 反向代理场景下优先使用用户实际访问的 host
    return request.host_url.rstrip("/")


def is_safe_next_url(next_url: str) -> bool:
    if not next_url:
        return False
    return next_url.startswith("/") and not next_url.startswith("//")


def is_admin_logged_in() -> bool:
    return session.get("admin_logged_in") is True


def admin_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not is_admin_logged_in():
            next_url = request.full_path if request.query_string else request.path
            return redirect(url_for("admin_login", next=next_url))
        return func(*args, **kwargs)
    return wrapper


import requests
import os

RESEND_API_KEY = os.getenv("RESEND_API_KEY")

def send_qr_email(to_email, name, event_id, qr_link, qr_img):
    url = "https://api.resend.com/emails"

    headers = {
        "Authorization": f"Bearer {RESEND_API_KEY}",
        "Content-Type": "application/json"
    }

    if not RESEND_API_KEY:
        raise RuntimeError("RESEND_API_KEY is not set in environment variables.")

    data = {
        "from": RESEND_FROM,
        "to": [to_email],
        "subject": f"签到二维码 - 活动 {event_id}",
        "html": f"""
        <h2>你好 {name or 'Guest'}</h2>
        <p>感谢报名参加本次活动。请在到场时向工作人员出示下方二维码，工作人员扫码后即可完成签到。</p>
        <p><img src="{qr_img}" width="220" style="width:220px;height:220px;"/></p>
        <p>如果二维码无法显示，也可以打开以下链接：</p>
        <p><a href="{qr_link}">{qr_link}</a></p>
        <p>Webull Event Team</p>
        """
    }

    response = requests.post(url, json=data, headers=headers, timeout=20)
    print("Resend response:", response.status_code, response.text)
    if response.status_code >= 400:
        raise RuntimeError(f"Resend error {response.status_code}: {response.text}")


def generate_unique_token() -> str:
    conn = get_conn()
    cur = conn.cursor()
    while True:
        token = secrets.token_urlsafe(32)
        cur.execute("SELECT id FROM registrants WHERE checkin_token = ? LIMIT 1", (token,))
        if cur.fetchone() is None:
            conn.close()
            return token


def ensure_tokens_for_event(event_id: str) -> int:
    """为指定活动所有未生成 token 的名单用户生成唯一 token。返回新增 token 数。"""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id FROM registrants WHERE event_id = ? AND (checkin_token IS NULL OR checkin_token = '')",
        (event_id,),
    )
    ids = [row["id"] for row in cur.fetchall()]

    count = 0
    for rid in ids:
        token = secrets.token_urlsafe(32)
        # 极小概率冲突，冲突则重试
        while True:
            cur.execute("SELECT id FROM registrants WHERE checkin_token = ? LIMIT 1", (token,))
            if cur.fetchone() is None:
                break
            token = secrets.token_urlsafe(32)
        cur.execute("UPDATE registrants SET checkin_token = ? WHERE id = ?", (token, rid))
        count += 1

    conn.commit()
    conn.close()
    return count


# -----------------------------
# 数据操作
# -----------------------------
def import_excel_to_db(file_bytes: bytes, event_id: str) -> int:
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    name_col = pick_column(df.columns.tolist(), ["姓名", "名字", "name"])
    phone_col = pick_column(df.columns.tolist(), ["手机号", "手机", "电话", "phone", "mobile"])
    email_col = pick_column(df.columns.tolist(), ["邮箱", "email", "mail", "电子邮箱"])
    org_col = pick_column(df.columns.tolist(), ["单位", "机构", "公司", "organization"])
    has_webull_col = pick_column(df.columns.tolist(), ["是否已有webull账户", "是否已注册webull", "has_webull_account"])
    opened_col = pick_column(df.columns.tolist(), ["是否已完成webull注册开户", "是否已开户", "webull_opened"])

    if phone_col is None and email_col is None:
        raise ValueError("Excel 中至少要包含手机号或邮箱字段。")

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM checkins WHERE event_id = ?", (event_id,))
    cur.execute("DELETE FROM registrants WHERE event_id = ?", (event_id,))

    count = 0
    for _, row in df.iterrows():
        name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
        phone = normalize_phone(row[phone_col]) if phone_col else ""
        email = normalize_email(row[email_col]) if email_col else ""
        organization = str(row[org_col]).strip() if org_col and pd.notna(row[org_col]) else ""
        has_webull_account = normalize_yes_no(row[has_webull_col]) if has_webull_col else ""
        webull_opened = normalize_yes_no(row[opened_col]) if opened_col else ""

        if not phone and not email:
            continue

        token = secrets.token_urlsafe(32)
        cur.execute(
            """
            INSERT INTO registrants
            (event_id, name, phone, email, organization, source,
             has_webull_account, webull_opened, checkin_token, raw_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                event_id,
                name,
                phone,
                email,
                organization,
                "imported",
                has_webull_account,
                webull_opened,
                token,
                row.to_json(force_ascii=False),
                now_str(),
            ),
        )
        count += 1

    conn.commit()
    conn.close()
    return count


def find_registrant(event_id: str, phone: str, email: str):
    conn = get_conn()
    cur = conn.cursor()
    row = None
    if phone:
        cur.execute(
            "SELECT * FROM registrants WHERE event_id = ? AND phone = ? LIMIT 1",
            (event_id, phone),
        )
        row = cur.fetchone()
    if row is None and email:
        cur.execute(
            "SELECT * FROM registrants WHERE event_id = ? AND email = ? LIMIT 1",
            (event_id, email),
        )
        row = cur.fetchone()
    conn.close()
    return row


def find_registrant_by_token(token: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM registrants WHERE checkin_token = ? LIMIT 1", (token,))
    row = cur.fetchone()
    conn.close()
    return row


def create_walkin_registrant(
    event_id: str,
    name: str,
    phone: str,
    email: str,
    organization: str,
    has_webull_account: str,
    webull_opened: str,
):
    token = generate_unique_token()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO registrants
        (event_id, name, phone, email, organization, source,
         has_webull_account, webull_opened, checkin_token, raw_json, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            event_id,
            name,
            phone,
            email,
            organization,
            "walkin",
            has_webull_account,
            webull_opened,
            token,
            "",
            now_str(),
        ),
    )
    registrant_id = cur.lastrowid
    conn.commit()
    cur.execute("SELECT * FROM registrants WHERE id = ?", (registrant_id,))
    row = cur.fetchone()
    conn.close()
    return row


def insert_checkin(
    event_id: str,
    registrant_id: Optional[int],
    submitted_phone: str,
    submitted_email: str,
    has_webull_account: str,
    checkin_method: str,
    status: str,
    message: str,
    ip: str,
    user_agent: str,
) -> tuple[bool, str]:
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """
            INSERT INTO checkins (
                event_id, registrant_id, submitted_phone, submitted_email,
                has_webull_account, checkin_method, status, message,
                ip, user_agent, checked_in_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                event_id,
                registrant_id,
                submitted_phone,
                submitted_email,
                has_webull_account,
                checkin_method,
                status,
                message,
                ip,
                user_agent,
                now_str(),
            ),
        )
        conn.commit()
        return True, "ok"
    except sqlite3.IntegrityError:
        return False, "该参会者已签到，请勿重复提交。"
    finally:
        conn.close()


# -----------------------------
# HTML
# -----------------------------
BASE_HTML = """
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>{{ title }}</title>
<style>
:root {
  --bg:#f5f7fb; --card:#ffffff; --line:#e5e7eb; --text:#111827; --muted:#6b7280;
  --brand:#1677ff; --brand2:#0f5ed7; --ok:#16a34a; --okbg:#dcfce7; --err:#dc2626; --errbg:#fee2e2;
}
* { box-sizing:border-box; }
body { margin:0; background:var(--bg); color:var(--text); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif; }
.page { max-width:1100px; margin:0 auto; padding:24px 16px 40px; }
.card { background:var(--card); border-radius:18px; box-shadow:0 8px 28px rgba(0,0,0,.06); padding:20px; }
.topbar { display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; margin-bottom:18px; }
.brand { font-size:28px; font-weight:800; }
.sub { color:var(--muted); font-size:14px; }
a { color:var(--brand); text-decoration:none; }
.grid { display:grid; gap:16px; }
.grid-2 { grid-template-columns: 1fr 1fr; }
@media (max-width: 860px) { .grid-2 { grid-template-columns: 1fr; } }
input, button, select { width:100%; border:1px solid #dbe3ef; border-radius:14px; padding:13px 14px; font-size:16px; background:#fff; }
button { background:linear-gradient(135deg,var(--brand),var(--brand2)); color:#fff; border:none; font-weight:700; cursor:pointer; }
button:hover { opacity:.95; }
.table-wrap { overflow:auto; }
table { width:100%; border-collapse:collapse; font-size:14px; }
th, td { border-bottom:1px solid var(--line); padding:12px 10px; text-align:left; white-space:nowrap; }
th { background:#f9fafb; position:sticky; top:0; }
.stat-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }
@media (max-width: 900px) { .stat-grid { grid-template-columns:repeat(2,1fr); } }
@media (max-width: 520px) { .stat-grid { grid-template-columns:1fr; } }
.stat { padding:16px; border-radius:16px; background:#f8fbff; border:1px solid #e6eefc; }
.stat .k { color:var(--muted); font-size:13px; }
.stat .v { font-size:28px; font-weight:800; margin-top:6px; }
.ok { background:var(--okbg); color:#166534; border:1px solid #86efac; padding:12px 14px; border-radius:14px; }
.err { background:var(--errbg); color:#991b1b; border:1px solid #fca5a5; padding:12px 14px; border-radius:14px; }
.mini-phone { max-width:430px; margin:0 auto; min-height:760px; background:linear-gradient(180deg,#eef5ff,#f8fbff); border-radius:28px; padding:18px; box-shadow:0 18px 38px rgba(0,0,0,.08); }
.wx-header { padding:16px 8px 18px; text-align:center; }
.wx-title { font-size:22px; font-weight:800; }
.wx-sub { color:var(--muted); font-size:13px; margin-top:6px; }
.wx-card { background:#fff; border-radius:20px; padding:18px; box-shadow:0 8px 24px rgba(0,0,0,.06); margin-bottom:14px; }
.wx-btn { border-radius:16px; padding:14px; font-size:17px; }
.badge { display:inline-block; padding:4px 10px; border-radius:999px; background:#e8f1ff; color:var(--brand2); font-size:12px; font-weight:700; }
.radio-group { display:grid; gap:10px; }
.radio-row { display:flex; gap:16px; align-items:center; flex-wrap:wrap; }
.radio-row label { display:flex; gap:6px; align-items:center; font-size:15px; color:var(--text); }
.radio-row input[type="radio"] { width:auto; }
.helper { font-size:13px; color:var(--muted); margin-top:-8px; }
.success-box { background:#ecfdf5; border:2px solid #22c55e; border-radius:18px; padding:18px 16px; text-align:center; margin-bottom:12px; }
.success-icon { font-size:44px; line-height:1; margin-bottom:8px; }
.success-title { font-size:24px; font-weight:800; color:#166534; }
</style>
</head>
<body><div class="page">{{ body|safe }}</div></body>
</html>
"""


def page(title: str, body: str):
    return render_template_string(BASE_HTML, title=title, body=body)


def success_html(title: str, name: str = "", extra: str = "") -> str:
    return f"""
    <div class="success-box">
      <div class="success-icon">✅</div>
      <div class="success-title">{title}</div>
      {f'<div style="margin-top:10px;font-size:16px;color:#14532d;">{name}</div>' if name else ''}
      {extra}
      <div style="margin-top:10px;font-size:13px;color:#15803d;">系统已记录本次签到，请勿重复提交。</div>
    </div>
    """


# -----------------------------
# 后台首页
# -----------------------------
@app.route("/")
def root():
    return redirect(url_for("admin"))


@app.route("/admin/register", methods=["GET", "POST"])
def admin_register():
    msg = ""

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        register_code = request.form.get("register_code", "").strip()

        if not username or not password:
            msg = '<div class="err">请填写用户名和密码。</div>'
        elif len(password) < 6:
            msg = '<div class="err">密码至少需要 6 位。</div>'
        elif not ADMIN_REGISTER_CODE:
            msg = '<div class="err">系统尚未设置 ADMIN_REGISTER_CODE。请先在 Render 的 Environment 中添加管理员注册码。</div>'
        elif register_code != ADMIN_REGISTER_CODE:
            msg = '<div class="err">管理员注册码错误。</div>'
        else:
            try:
                conn = get_conn()
                cur = conn.cursor()
                cur.execute(
                    """
                    INSERT INTO admin_users (username, password_hash, created_at)
                    VALUES (?, ?, ?)
                    """,
                    (username, generate_password_hash(password), now_str()),
                )
                conn.commit()
                conn.close()
                return redirect(url_for("admin_login", msg="注册成功，请登录。"))
            except sqlite3.IntegrityError:
                msg = '<div class="err">该用户名已经存在，请换一个用户名。</div>'

    body = f"""
    <div class="mini-phone">
      <div class="wx-header">
        <div class="badge">后台管理</div>
        <div class="wx-title">注册管理员账号</div>
        <div class="wx-sub">注册后才能进入讲座签到后台</div>
      </div>
      <div class="wx-card">
        {msg}
        <form method="post" class="grid">
          <input name="username" placeholder="管理员用户名" required>
          <input name="password" type="password" placeholder="管理员密码，至少 6 位" required>
          <input name="register_code" type="password" placeholder="管理员注册码" required>
          <button class="wx-btn" type="submit">注册管理员</button>
        </form>
        <p class="sub" style="margin-top:14px;">已有账号？<a href="/admin/login">去登录</a></p>
      </div>
    </div>
    """
    return page("注册管理员", body)


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    msg_text = request.args.get("msg", "").strip()
    msg = f'<div class="ok">{msg_text}</div>' if msg_text else ""
    next_url = request.args.get("next", "/admin").strip() or "/admin"
    if not is_safe_next_url(next_url):
        next_url = "/admin"

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        next_url = request.form.get("next", "/admin").strip() or "/admin"
        if not is_safe_next_url(next_url):
            next_url = "/admin"

        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT * FROM admin_users WHERE username = ? LIMIT 1", (username,))
        user = cur.fetchone()
        conn.close()

        if not user or not check_password_hash(user["password_hash"], password):
            msg = '<div class="err">用户名或密码错误。</div>'
        else:
            session["admin_logged_in"] = True
            session["admin_username"] = username
            return redirect(next_url)

    body = f"""
    <div class="mini-phone">
      <div class="wx-header">
        <div class="badge">后台管理</div>
        <div class="wx-title">管理员登录</div>
        <div class="wx-sub">登录后进入讲座签到后台</div>
      </div>
      <div class="wx-card">
        {msg}
        <form method="post" class="grid">
          <input type="hidden" name="next" value="{next_url}">
          <input name="username" placeholder="管理员用户名" required>
          <input name="password" type="password" placeholder="管理员密码" required>
          <button class="wx-btn" type="submit">登录</button>
        </form>
        <p class="sub" style="margin-top:14px;">还没有管理员账号？<a href="/admin/register">注册管理员</a></p>
      </div>
    </div>
    """
    return page("管理员登录", body)


@app.route("/admin/logout")
def admin_logout():
    session.pop("admin_logged_in", None)
    session.pop("admin_username", None)
    return redirect(url_for("admin_login", msg="已退出登录。"))


@app.route("/admin", methods=["GET", "POST"])
@admin_required
def admin():
    msg_text = request.args.get("msg", "").strip()
    msg_type = request.args.get("msg_type", "ok").strip()
    msg_html = make_msg_html(msg_text, msg_type)

    if request.method == "POST":
        event_id = request.form.get("event_id", "").strip()
        file = request.files.get("file")
        if not event_id:
            return redirect(url_for("admin", msg="请先填写活动 ID。", msg_type="err"))
        if not file or not file.filename:
            return redirect(url_for("admin", msg="请上传 Excel 名单。", msg_type="err"))
        try:
            count = import_excel_to_db(file.read(), event_id)
            return redirect(url_for("admin", msg=f"导入成功，共导入 {count} 条名单，并已自动生成二维码 token。", msg_type="ok"))
        except Exception as e:
            return redirect(url_for("admin", msg=f"导入失败:{e}", msg_type="err"))

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT r.event_id,
               COUNT(DISTINCT r.id) AS total_registrants,
               COUNT(DISTINCT CASE WHEN c.status='success' THEN c.registrant_id END) AS total_checked_in,
               COUNT(DISTINCT CASE WHEN c.status='failed' THEN c.id END) AS failed_attempts,
               COUNT(DISTINCT CASE WHEN r.checkin_token IS NOT NULL AND r.checkin_token != '' THEN r.id END) AS qr_ready
        FROM registrants r
        LEFT JOIN checkins c ON r.event_id = c.event_id
        GROUP BY r.event_id
        ORDER BY r.event_id
        """
    )
    rows = cur.fetchall()
    cur.execute("SELECT COUNT(*) AS c FROM registrants")
    total_reg = cur.fetchone()["c"]
    cur.execute("SELECT COUNT(*) AS c FROM checkins WHERE status='success'")
    total_success = cur.fetchone()["c"]
    cur.execute("SELECT COUNT(*) AS c FROM checkins WHERE status='failed'")
    total_failed = cur.fetchone()["c"]
    conn.close()

    table_rows = "".join(
        f"<tr>"
        f"<td>{r['event_id']}</td>"
        f"<td>{r['total_registrants']}</td>"
        f"<td>{r['total_checked_in']}</td>"
        f"<td>{r['failed_attempts']}</td>"
        f"<td>{r['qr_ready']}</td>"
        f"<td>"
        f"<a href='/m/checkin?event_id={r['event_id']}'>手填签到页</a> | "
        f"<a href='/admin/generate_qr_tokens?event_id={r['event_id']}'>生成二维码Token</a> | "
        f"<a href='/admin/export_qr_links?event_id={r['event_id']}'>导出二维码链接</a> | "
        f"<a href='/admin/send_qr_emails?event_id={r['event_id']}' onclick=\"return confirm('确定向该活动所有有邮箱的用户发送二维码邮件？');\">发送二维码邮件</a> | "
        f"<a href='/admin/records?event_id={r['event_id']}'>签到记录</a> | "
        f"<a href='/admin/export?event_id={r['event_id']}'>导出记录</a> | "
        f"<a href='/admin/delete?event_id={r['event_id']}' onclick=\"return confirm('确定删除该活动所有数据？');\">删除</a>"
        f"</td></tr>"
        for r in rows
    )

    body = f"""
    <div class="topbar">
      <div>
        <div class="brand">讲座签到后台</div>
        <div class="sub">二维码签到 + 手填签到 + 现场补签</div>
      </div>
      <div>
        <span class="sub">当前管理员:{session.get("admin_username", "")}</span>
        | <a href="/staff/scan" style="font-weight:800;">现场扫码模式</a>
        | <a href="/admin/logout">退出登录</a>
      </div>
    </div>

    <div class="stat-grid" style="margin-bottom:16px;">
      <div class="stat"><div class="k">名单总人数</div><div class="v">{total_reg}</div></div>
      <div class="stat"><div class="k">成功签到</div><div class="v">{total_success}</div></div>
      <div class="stat"><div class="k">未匹配尝试</div><div class="v">{total_failed}</div></div>
      <div class="stat"><div class="k">活动数量</div><div class="v">{len(rows)}</div></div>
    </div>

    {msg_html}

    <div class="grid grid-2">
      <div class="card">
        <h3 style="margin-top:0;">上传活动名单</h3>
        <form method="post" enctype="multipart/form-data" class="grid">
          <input name="event_id" placeholder="活动 ID，例如 lecture_001" required>
          <input type="file" name="file" accept=".xlsx,.xls" required>
          <button type="submit">导入名单并生成二维码 token</button>
        </form>
        <p class="sub">Excel 至少应包含手机号或邮箱字段。若含邮箱，可用于后续邮件发送二维码。</p>
      </div>

      <div class="card">
        <h3 style="margin-top:0;">现场手动补签</h3>
        <form method="post" action="/admin/manual_checkin" class="grid">
          <input name="event_id" placeholder="活动 ID，例如 lecture_001" required>
          <input name="phone" placeholder="手机号（优先匹配）">
          <input name="email" placeholder="邮箱（手机号不便时可填）">
          <div class="radio-group">
            <div class="helper">该来宾是否已经注册开户 Webull 账户？</div>
            <div class="radio-row">
              <label><input type="radio" name="has_webull_account" value="是" required> 是</label>
              <label><input type="radio" name="has_webull_account" value="否"> 否</label>
            </div>
          </div>
          <button type="submit">手动补签</button>
        </form>
      </div>
    </div>

    <div class="card" style="margin-top:16px;">
      <h3 style="margin-top:0;">临时访客登记 + 补签</h3>
      <form method="post" action="/admin/walkin_checkin" class="grid grid-2">
        <input name="event_id" placeholder="活动 ID，例如 lecture_001" required>
        <input name="name" placeholder="访客姓名" required>
        <input name="phone" placeholder="手机号（建议填写）">
        <input name="email" placeholder="邮箱（手机号不便时可填）">
        <input name="organization" placeholder="单位 / 公司 / 机构（可选）">
        <div class="radio-group">
          <div class="helper">是否已有 Webull 账户？</div>
          <div class="radio-row">
            <label><input type="radio" name="has_webull_account" value="是" required> 是</label>
            <label><input type="radio" name="has_webull_account" value="否"> 否</label>
          </div>
        </div>
        <div class="radio-group">
          <div class="helper">是否已完成 Webull 注册开户？</div>
          <div class="radio-row">
            <label><input type="radio" name="webull_opened" value="是" required> 是</label>
            <label><input type="radio" name="webull_opened" value="否"> 否</label>
          </div>
        </div>
        <button type="submit">登记并补签</button>
      </form>
    </div>

    <div class="card" style="margin-top:16px;">
      <h3 style="margin-top:0;">二维码签到说明</h3>
      <p>导入名单后，系统会为每位用户生成唯一二维码 token。</p>
      <p>点击“导出二维码链接”，即可得到每位用户的专属签到链接。你可以把该链接或二维码发到用户邮箱。</p>
      <p>现场必须由工作人员进入“现场扫码模式”后扫描，用户自己扫描二维码不会完成签到。</p>
      <p><a href="/staff/scan" style="font-weight:800;">进入现场扫码模式</a></p>
    </div>

    <div class="card" style="margin-top:16px;">
      <h3 style="margin-top:0;">活动列表</h3>
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>活动 ID</th><th>名单人数</th><th>成功签到</th><th>未匹配尝试</th><th>二维码数</th><th>操作</th></tr>
          </thead>
          <tbody>{table_rows or '<tr><td colspan="6">暂无活动</td></tr>'}</tbody>
        </table>
      </div>
    </div>
    """
    return page("讲座签到后台", body)


@app.route("/admin/send_one_email")
@admin_required
def send_one_email():
    rid = request.args.get("id")

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, name, email, checkin_token, event_id FROM registrants WHERE id=?",
        (rid,)
    )
    r = cur.fetchone()
    conn.close()

    if not r:
        return redirect(url_for("admin", msg="用户不存在", msg_type="err"))

    if not r["email"]:
        return redirect(url_for("admin", msg="该用户没有邮箱，无法发送", msg_type="err"))

    base_url = get_base_url()
    qr_link = f"{base_url}/qr_checkin?token={r['checkin_token']}"
    qr_img = f"{base_url}/qr_image?token={r['checkin_token']}"

    try:
        send_qr_email(r["email"], r["name"], r["event_id"], qr_link, qr_img)

        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "UPDATE registrants SET qr_sent_at = ? WHERE id = ?",
            (now_str(), r["id"])
        )
        conn.commit()
        conn.close()

        return redirect(url_for("admin", msg=f"邮件发送成功: {r['email']}", msg_type="ok"))

    except Exception as e:
        return redirect(url_for("admin", msg=f"邮件发送失败: {e}", msg_type="err"))
# -----------------------------
# 工作人员扫码入口与权限控制
# -----------------------------
def staff_is_logged_in() -> bool:
    return session.get("staff_logged_in") is True


@app.route("/staff/login", methods=["GET", "POST"])
def staff_login():
    msg = ""
    next_url = request.args.get("next", "/staff/scan")

    if request.method == "POST":
        password = request.form.get("password", "").strip()
        next_url = request.form.get("next", "/staff/scan").strip() or "/staff/scan"
        if password == STAFF_SCAN_PASSWORD:
            session["staff_logged_in"] = True
            return redirect(next_url)
        msg = '<div class="err">工作人员口令错误，请重试。</div>'

    body = f"""
    <div class="mini-phone">
      <div class="wx-header">
        <div class="badge">工作人员入口</div>
        <div class="wx-title">扫码权限验证</div>
        <div class="wx-sub">只有工作人员登录后，扫码才会完成签到</div>
      </div>
      <div class="wx-card">
        {msg}
        <form method="post" class="grid">
          <input type="hidden" name="next" value="{next_url}">
          <input type="password" name="password" placeholder="请输入工作人员口令" required>
          <button class="wx-btn" type="submit">进入现场扫码模式</button>
        </form>
      </div>
    </div>
    """
    return page("工作人员登录", body)


@app.route("/staff/logout")
def staff_logout():
    session.pop("staff_logged_in", None)
    return redirect(url_for("staff_login", msg="已退出"))


@app.route("/staff/scan")
def staff_scan():
    if not staff_is_logged_in():
        return redirect(url_for("staff_login", next="/staff/scan"))

    body = """
    <div class="mini-phone">
      <div class="wx-header">
        <div class="badge">现场工作人员</div>
        <div class="wx-title">现场扫码模式</div>
        <div class="wx-sub">扫描参会者邮件中的个人二维码</div>
      </div>

      <div class="wx-card">
        <div id="reader" style="width:100%; min-height:280px;"></div>
        <div id="scan-result" class="sub" style="margin-top:12px;">请允许浏览器访问摄像头，然后对准参会者二维码。</div>
        <button id="switch-camera-btn" class="wx-btn" type="button" style="margin-top:12px;">切换摄像头</button>
      </div>

      <div class="wx-card">
        <h3 style="margin:0 0 8px;">工作人员说明</h3>
        <ul style="padding-left:18px; margin:8px 0; line-height:1.8;">
          <li>只有从这个页面扫码，系统才会确认签到。</li>
          <li>二维码内容应为 <code>/qr_checkin?token=...</code>。</li>
          <li>扫码成功后会跳转到核验页，显示姓名、手机号、邮箱等信息。</li>
          <li>若摄像头不可用，可用手机系统相机扫码；但必须先在本浏览器登录工作人员口令。</li>
        </ul>
        <p><a href="/staff/logout">退出工作人员模式</a></p>
      </div>
    </div>

    <script src="https://unpkg.com/html5-qrcode" type="text/javascript"></script>
    <script>
      function extractToken(decodedText) {
        try {
          const url = new URL(decodedText);
          return url.searchParams.get("token");
        } catch (e) {
          const match = decodedText.match(/token=([^&]+)/);
          return match ? decodeURIComponent(match[1]) : null;
        }
      }

      function onScanSuccess(decodedText, decodedResult) {
        const resultBox = document.getElementById("scan-result");
        const token = extractToken(decodedText);
        if (!token) {
          resultBox.innerHTML = "<span style='color:#dc2626;font-weight:700;'>二维码格式不正确，未找到 token。</span>";
          return;
        }
        resultBox.innerHTML = "<span style='color:#166534;font-weight:700;'>扫码成功，正在核验签到...</span>";
        window.location.href = "/qr_checkin?token=" + encodeURIComponent(token);
      }

      function onScanFailure(error) {
        // 连续扫描时会频繁触发，不需要显示错误
      }

      let html5QrCode = new Html5Qrcode("reader");
      let cameras = [];
      let currentCameraIndex = 0;
      let isScanning = false;

      function pickBackCameraIndex(devices) {
        // 尽量优先选后置摄像头。不同手机浏览器命名不一样，所以做多关键词匹配。
        const keywords = ["back", "rear", "environment", "后置", "背面"];
        const idx = devices.findIndex(d => {
          const label = (d.label || "").toLowerCase();
          return keywords.some(k => label.includes(k));
        });
        return idx >= 0 ? idx : 0;
      }

      function startCamera(index) {
        if (!cameras.length) return;
        const cameraId = cameras[index].id;
        const resultBox = document.getElementById("scan-result");
        resultBox.innerHTML = "正在启动摄像头...";

        html5QrCode.start(
          cameraId,
          { fps: 10, qrbox: { width: 240, height: 240 } },
          onScanSuccess,
          onScanFailure
        ).then(() => {
          isScanning = true;
          const label = cameras[index].label || `摄像头 ${index + 1}`;
          resultBox.innerHTML = `当前摄像头:${label}`;
        }).catch(err => {
          resultBox.innerHTML = "<span style='color:#dc2626;'>无法启动摄像头，请检查浏览器权限。</span>";
        });
      }

      function switchCamera() {
        if (!cameras.length) return;
        currentCameraIndex = (currentCameraIndex + 1) % cameras.length;
        if (isScanning) {
          html5QrCode.stop().then(() => {
            isScanning = false;
            startCamera(currentCameraIndex);
          }).catch(() => {
            startCamera(currentCameraIndex);
          });
        } else {
          startCamera(currentCameraIndex);
        }
      }

      document.getElementById("switch-camera-btn").addEventListener("click", switchCamera);

      Html5Qrcode.getCameras().then(devices => {
        cameras = devices || [];
        if (cameras.length) {
          currentCameraIndex = pickBackCameraIndex(cameras);
          startCamera(currentCameraIndex);
        } else {
          document.getElementById("scan-result").innerHTML = "<span style='color:#dc2626;'>未检测到摄像头。</span>";
        }
      }).catch(err => {
        document.getElementById("scan-result").innerHTML = "<span style='color:#dc2626;'>无法启动摄像头，请检查浏览器权限或改用手机系统相机。</span>";
      });
    </script>
    """
    return page("现场扫码模式", body)


# -----------------------------
# 二维码 token 与导出
# -----------------------------
@app.route("/admin/generate_qr_tokens")
@admin_required
def admin_generate_qr_tokens():
    event_id = request.args.get("event_id", "").strip()
    if not event_id:
        return redirect(url_for("admin", msg="缺少活动 ID。", msg_type="err"))
    count = ensure_tokens_for_event(event_id)
    return redirect(url_for("admin", msg=f"已为活动 {event_id} 新增生成 {count} 个二维码 token。", msg_type="ok"))


@app.route("/admin/export_qr_links")
@admin_required
def admin_export_qr_links():
    event_id = request.args.get("event_id", "").strip()
    if not event_id:
        return redirect(url_for("admin", msg="缺少活动 ID。", msg_type="err"))

    ensure_tokens_for_event(event_id)
    base_url = get_base_url()

    conn = get_conn()
    query = """
        SELECT id, event_id, name, phone, email, organization, source,
               has_webull_account, webull_opened, checkin_token
        FROM registrants
        WHERE event_id = ?
        ORDER BY id
    """
    df = pd.read_sql_query(query, conn, params=(event_id,))
    conn.close()

    if not df.empty:
        df["二维码签到链接"] = df["checkin_token"].apply(lambda t: f"{base_url}/qr_checkin?token={t}")
        df["二维码图片链接"] = df["checkin_token"].apply(lambda t: f"{base_url}/qr_image?token={t}")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="二维码链接")
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{event_id}_二维码链接.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )



@app.route("/admin/send_qr_emails")
@admin_required
def admin_send_qr_emails():
    event_id = request.args.get("event_id", "").strip()
    if not event_id:
        return redirect(url_for("admin", msg="缺少活动 ID。", msg_type="err"))

    ensure_tokens_for_event(event_id)
    base_url = get_base_url()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, event_id, name, email, checkin_token
        FROM registrants
        WHERE event_id = ?
          AND email IS NOT NULL
          AND email != ''
        ORDER BY id
        """,
        (event_id,),
    )
    rows = cur.fetchall()
    conn.close()

    sent = 0
    failed = 0

    for r in rows:
        try:
            qr_link = f"{base_url}/qr_checkin?token={r['checkin_token']}"
            qr_image_url = f"{base_url}/qr_image?token={r['checkin_token']}"

            send_qr_email(
                to_email=r["email"],
                name=r["name"],
                event_id=event_id,
                qr_link=qr_link,
                qr_image_url=qr_image_url,
            )

            conn = get_conn()
            cur = conn.cursor()
            cur.execute(
                "UPDATE registrants SET qr_sent_at = ? WHERE id = ?",
                (now_str(), r["id"]),
            )
            conn.commit()
            conn.close()

            sent += 1
        except Exception:
            failed += 1

    return redirect(
        url_for(
            "admin",
            msg=f"二维码邮件发送完成: 成功 {sent} 封, 失败 {failed} 封。",
            msg_type="ok" if failed == 0 else "err",
        )
    )


@app.route("/qr_image")
def qr_image():
    token = request.args.get("token", "").strip()
    if not token:
        return "Missing token", 400

    try:
        import qrcode
    except ImportError:
        return "qrcode package not installed. Please add qrcode[pil] to requirements.txt", 500

    url = f"{get_base_url()}/qr_checkin?token={token}"
    img = qrcode.make(url)
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    return send_file(buffer, mimetype="image/png")


@app.route("/qr_checkin")
def qr_checkin():
    if not staff_is_logged_in():
        next_url = request.full_path if request.query_string else request.path
        body = f"""
        <div class="mini-phone">
          <div class="wx-header">
            <div class="badge">工作人员权限</div>
            <div class="wx-title">需要工作人员确认</div>
            <div class="wx-sub">该二维码只能由工作人员扫码确认签到</div>
          </div>
          <div class="wx-card">
            <div class="err">你当前不是工作人员扫码模式，因此不会完成签到。</div>
            <p class="sub">请工作人员点击下方按钮登录后，再扫描或打开该二维码链接。</p>
            <a href="/staff/login?next={next_url}"><button class="wx-btn" type="button">工作人员登录并继续核验</button></a>
          </div>
        </div>
        """
        return page("需要工作人员权限", body)

    token = request.args.get("token", "").strip()
    checked_time = now_str()

    if not token:
        body = """
        <div class="mini-phone">
          <div class="wx-header">
            <div class="badge">工作人员扫码</div>
            <div class="wx-title">二维码核验</div>
          </div>
          <div class="wx-card">
            <div class="err">二维码无效: 缺少 token。请让参会者重新出示二维码，或联系后台工作人员。</div>
          </div>
        </div>
        """
        return page("二维码签到", body)

    registrant = find_registrant_by_token(token)
    if registrant is None:
        body = """
        <div class="mini-phone">
          <div class="wx-header">
            <div class="badge">工作人员扫码</div>
            <div class="wx-title">二维码核验</div>
          </div>
          <div class="wx-card">
            <div class="err">二维码无效或已不存在。请核对该二维码是否来自本系统，或改用后台手动补签。</div>
          </div>
        </div>
        """
        return page("二维码签到", body)

    event_id = registrant["event_id"]
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
    ua = request.headers.get("User-Agent", "")
    has_webull_account = registrant["has_webull_account"] or ""

    ok, msg = insert_checkin(
        event_id=event_id,
        registrant_id=registrant["id"],
        submitted_phone=registrant["phone"] or "",
        submitted_email=registrant["email"] or "",
        has_webull_account=has_webull_account,
        checkin_method="qr_staff_scan",
        status="success",
        message="工作人员扫码二维码签到",
        ip=ip,
        user_agent=ua,
    )

    if ok:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("UPDATE registrants SET qr_used_at = ? WHERE id = ?", (checked_time, registrant["id"]))
        conn.commit()
        conn.close()

        result = f"""
        <div class="success-box">
          <div class="success-icon">✅</div>
          <div class="success-title">签到成功</div>
          <div style="margin-top:10px;font-size:16px;color:#14532d;">欢迎你，<strong>{registrant['name'] or '参会者'}</strong></div>
          <div style="margin-top:8px;font-size:14px;color:#166534;">本次为首次签到，已写入后台记录。</div>
        </div>
        """
    else:
        result = f"""
        <div class="success-box">
          <div class="success-icon">✅</div>
          <div class="success-title">核验成功</div>
          <div style="margin-top:10px;font-size:16px;color:#14532d;">欢迎你，<strong>{registrant['name'] or '参会者'}</strong></div>
          <div style="margin-top:8px;font-size:14px;color:#166534;">该二维码此前已完成签到，本次不重复记录。</div>
          <div style="margin-top:4px;font-size:14px;color:#166534;">后台只保留第一次签到记录。</div>
        </div>
        """

    source_label = "临时访客" if registrant["source"] == "walkin" else "预登记名单"

    body = f"""
    <div class="mini-phone">
      <div class="wx-header">
        <div class="badge">工作人员扫码</div>
        <div class="wx-title">二维码核验签到</div>
        <div class="wx-sub">活动 ID: {event_id}</div>
      </div>

      <div class="wx-card">
        {result}
        <table style="width:100%;border-collapse:collapse;margin-top:14px;font-size:14px;">
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">姓名</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{registrant['name'] or ''}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">手机号</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{registrant['phone'] or ''}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">邮箱</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{registrant['email'] or ''}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">单位</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{registrant['organization'] or ''}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">来源</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{source_label}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">是否已有 Webull 账户</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{has_webull_account or ''}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">是否完成开户</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{registrant['webull_opened'] or ''}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">扫码时间</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">{checked_time}</td></tr>
          <tr><th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">签到方式</th><td style="padding:10px;border-bottom:1px solid #e5e7eb;">工作人员扫码二维码</td></tr>
        </table>
      </div>

      <div class="wx-card">
        <a href="/staff/scan"><button class="wx-btn" type="button">继续扫描下一位</button></a>
      </div>

      <div class="wx-card">
        <h3 style="margin:0 0 8px;">工作人员提示</h3>
        <ul style="padding-left:18px; margin:8px 0; line-height:1.8;">
          <li>请核对姓名、手机号或邮箱是否与到场人员一致。</li>
          <li>重复扫描会显示核验成功，但后台只保留第一次签到记录。</li>
          <li>若二维码无效，请回后台使用“手动补签”或“临时访客登记”。</li>
        </ul>
      </div>
    </div>
    """
    return page("二维码签到", body)




# -----------------------------
# 签到记录、导出、删除
# -----------------------------
@app.route("/admin/records")
@admin_required
def admin_records():
    event_id = request.args.get("event_id", "").strip()
    msg_text = request.args.get("msg", "").strip()
    msg_type = request.args.get("msg_type", "ok").strip()
    msg_html = make_msg_html(msg_text, msg_type)

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT
            r.id AS registrant_id,
            r.event_id,
            r.name,
            r.phone,
            r.email,
            r.organization,
            r.source,
            r.has_webull_account,
            r.webull_opened,
            r.qr_sent_at,
            r.qr_used_at,
            r.checkin_token,
            c.checked_in_at,
            c.status,
            c.checkin_method,
            c.message
        FROM registrants r
        LEFT JOIN checkins c
            ON r.id = c.registrant_id
           AND r.event_id = c.event_id
        WHERE r.event_id = ?
        ORDER BY r.id ASC
        """,
        (event_id,),
    )
    rows = cur.fetchall()
    conn.close()

    table_rows = "".join(
        f"<tr>"
        f"<td>{x['registrant_id']}</td>"
        f"<td>{x['name'] or ''}</td>"
        f"<td>{x['phone'] or ''}</td>"
        f"<td>{x['email'] or ''}</td>"
        f"<td>{x['organization'] or ''}</td>"
        f"<td>{x['source'] or ''}</td>"
        f"<td>{x['has_webull_account'] or ''}</td>"
        f"<td>{x['webull_opened'] or ''}</td>"
        f"<td>{x['checked_in_at'] or '未签到'}</td>"
        f"<td>{x['checkin_method'] or ''}</td>"
        f"<td>{x['qr_sent_at'] or ''}</td>"
        f"<td>"
        f"<a href='/admin/send_one_email?id={x['registrant_id']}'>发送邮件</a> | "
        f"<a href='/qr_image?token={x['checkin_token']}' target='_blank'>二维码</a>"
        f"</td>"
        f"</tr>"
        for x in rows
    )

    body = f"""
    <div class="topbar">
      <div>
        <div class="brand" style="font-size:24px;">活动名单与签到记录</div>
        <div class="sub">活动 ID: {event_id}</div>
      </div>
      <div>
        <a href="/admin">返回后台</a> |
        <a href="/admin/export?event_id={event_id}">导出签到记录</a> |
        <a href="/admin/export_qr_links?event_id={event_id}">导出二维码链接</a>
      </div>
    </div>

    {msg_html}

    <div class="card">
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>姓名</th>
              <th>手机号</th>
              <th>邮箱</th>
              <th>单位</th>
              <th>来源</th>
              <th>是否已有Webull账户</th>
              <th>是否已完成开户</th>
              <th>签到时间</th>
              <th>签到方式</th>
              <th>邮件发送时间</th>
              <th>操作</th>
            </tr>
          </thead>
          <tbody>{table_rows or '<tr><td colspan="12">暂无名单</td></tr>'}</tbody>
        </table>
      </div>
    </div>
    """
    return page("签到记录", body)


@app.route("/admin/export")
@admin_required
def admin_export():
    event_id = request.args.get("event_id", "").strip()
    conn = get_conn()
    query = """
        SELECT c.event_id AS 活动ID,
               c.checked_in_at AS 签到时间,
               c.status AS 状态,
               c.checkin_method AS 签到方式,
               c.message AS 说明,
               r.name AS 姓名,
               r.organization AS 单位,
               r.source AS 来源,
               r.phone AS 名单手机号,
               r.email AS 名单邮箱,
               c.submitted_phone AS 提交手机号,
               c.submitted_email AS 提交邮箱,
               c.has_webull_account AS 是否已有Webull账户,
               r.webull_opened AS 是否已完成Webull注册开户,
               r.qr_used_at AS 二维码使用时间
        FROM checkins c
        LEFT JOIN registrants r ON c.registrant_id = r.id
        WHERE c.event_id = ?
        ORDER BY c.checked_in_at DESC
    """
    df = pd.read_sql_query(query, conn, params=(event_id,))
    conn.close()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="签到记录")
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{event_id}_签到记录.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/delete")
@admin_required
def admin_delete():
    event_id = request.args.get("event_id", "").strip()
    if not event_id:
        return redirect(url_for("admin"))
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM checkins WHERE event_id = ?", (event_id,))
    cur.execute("DELETE FROM registrants WHERE event_id = ?", (event_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("admin", msg=f"已删除活动:{event_id}", msg_type="ok"))


# -----------------------------
# 手填签到、手动补签、临时访客
# -----------------------------
@app.route("/admin/manual_checkin", methods=["POST"])
@admin_required
def admin_manual_checkin():
    event_id = request.form.get("event_id", "").strip()
    phone = normalize_phone(request.form.get("phone", ""))
    email = normalize_email(request.form.get("email", ""))
    has_webull_account = normalize_yes_no(request.form.get("has_webull_account", ""))

    if not event_id:
        return redirect(url_for("admin", msg="请先填写活动ID。", msg_type="err"))
    if not phone and not email:
        return redirect(url_for("admin", msg="请至少填写手机号或邮箱。", msg_type="err"))

    registrant = find_registrant(event_id, phone, email)
    if registrant is None:
        return redirect(url_for("admin", msg=f"未在活动 {event_id} 的名单中找到该手机号/邮箱。", msg_type="err"))

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE registrants SET has_webull_account = COALESCE(?, has_webull_account) WHERE id = ?",
        (has_webull_account or None, registrant["id"]),
    )
    conn.commit()
    conn.close()

    ok, insert_msg = insert_checkin(
        event_id, registrant["id"], phone, email, has_webull_account,
        "manual", "success", "后台手动补签", "admin-manual", "admin-manual"
    )
    if not ok:
        return redirect(url_for("admin_records", event_id=event_id, msg=insert_msg, msg_type="err"))
    return redirect(url_for("admin_records", event_id=event_id, msg=f"手动补签成功:{registrant['name'] or '该参会者'}", msg_type="ok"))


@app.route("/admin/walkin_checkin", methods=["POST"])
@admin_required
def admin_walkin_checkin():
    event_id = request.form.get("event_id", "").strip()
    name = request.form.get("name", "").strip()
    phone = normalize_phone(request.form.get("phone", ""))
    email = normalize_email(request.form.get("email", ""))
    organization = request.form.get("organization", "").strip()
    has_webull_account = normalize_yes_no(request.form.get("has_webull_account", ""))
    webull_opened = normalize_yes_no(request.form.get("webull_opened", ""))

    if not event_id:
        return redirect(url_for("admin", msg="请先填写活动ID。", msg_type="err"))
    if not name:
        return redirect(url_for("admin", msg="请填写临时访客姓名。", msg_type="err"))
    if not phone and not email:
        return redirect(url_for("admin", msg="临时访客至少需要手机号或邮箱其一。", msg_type="err"))

    existing = find_registrant(event_id, phone, email)
    if existing is not None:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "UPDATE registrants SET has_webull_account = COALESCE(?, has_webull_account), webull_opened = COALESCE(?, webull_opened) WHERE id = ?",
            (has_webull_account or None, webull_opened or None, existing["id"]),
        )
        conn.commit()
        conn.close()
        ok, insert_msg = insert_checkin(event_id, existing["id"], phone, email, has_webull_account, "walkin_admin", "success", "后台补签（已存在名单）", "admin-walkin", "admin-walkin")
        if not ok:
            return redirect(url_for("admin_records", event_id=event_id, msg=insert_msg, msg_type="err"))
        return redirect(url_for("admin_records", event_id=event_id, msg=f"补签成功:{existing['name'] or name}", msg_type="ok"))

    walkin = create_walkin_registrant(event_id, name, phone, email, organization, has_webull_account, webull_opened)
    ok, insert_msg = insert_checkin(event_id, walkin["id"], phone, email, has_webull_account, "walkin_admin", "success", "临时访客现场登记并签到", "admin-walkin", "admin-walkin")
    if not ok:
        return redirect(url_for("admin_records", event_id=event_id, msg=insert_msg, msg_type="err"))
    return redirect(url_for("admin_records", event_id=event_id, msg=f"临时访客登记并补签成功:{name}", msg_type="ok"))


@app.route("/m/checkin", methods=["GET", "POST"])
def mobile_checkin():
    event_id = request.args.get("event_id", "").strip() or request.form.get("event_id", "").strip()
    result = ""
    show_register_form = False

    if request.method == "POST":
        action = request.form.get("action", "checkin").strip()
        if action == "checkin":
            phone = normalize_phone(request.form.get("phone", ""))
            email = normalize_email(request.form.get("email", ""))
            has_webull_account = normalize_yes_no(request.form.get("has_webull_account", ""))
            if not event_id:
                result = '<div class="err">缺少活动 ID，请联系工作人员。</div>'
            elif not phone and not email:
                result = '<div class="err">请输入手机号或邮箱。</div>'
            elif not has_webull_account:
                result = '<div class="err">请选择你是否已经注册开户 Webull 账户。</div>'
            else:
                registrant = find_registrant(event_id, phone, email)
                ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
                ua = request.headers.get("User-Agent", "")
                if registrant is None:
                    insert_checkin(event_id, None, phone, email, has_webull_account, "form", "failed", "名单未匹配", ip, ua)
                    result = '<div class="err">未在名单中匹配到该手机号或邮箱。你可以填写下方登记信息完成现场登记。</div>'
                    show_register_form = True
                else:
                    conn = get_conn()
                    cur = conn.cursor()
                    cur.execute("UPDATE registrants SET has_webull_account = COALESCE(?, has_webull_account) WHERE id = ?", (has_webull_account or None, registrant["id"]))
                    conn.commit()
                    conn.close()
                    ok, msg = insert_checkin(event_id, registrant["id"], phone, email, has_webull_account, "form", "success", "签到成功", ip, ua)
                    if ok:
                        result = success_html("签到成功", f"欢迎你，<strong>{registrant['name'] or '参会者'}</strong>", f'<div style="margin-top:8px;font-size:14px;color:#166534;">是否已有 Webull 账户:{has_webull_account}</div>')
                    else:
                        result = f'<div class="err">{msg}</div>'

        elif action == "register_walkin":
            name = request.form.get("name", "").strip()
            phone = normalize_phone(request.form.get("phone", ""))
            email = normalize_email(request.form.get("email", ""))
            organization = request.form.get("organization", "").strip()
            has_webull_account = normalize_yes_no(request.form.get("has_webull_account", ""))
            webull_opened = normalize_yes_no(request.form.get("webull_opened", ""))
            if not event_id:
                result = '<div class="err">缺少活动 ID，请联系工作人员。</div>'
                show_register_form = True
            elif not name:
                result = '<div class="err">请填写姓名。</div>'
                show_register_form = True
            elif not phone and not email:
                result = '<div class="err">请至少填写手机号或邮箱其中一项。</div>'
                show_register_form = True
            elif not webull_opened:
                result = '<div class="err">请选择是否已完成 Webull 注册开户。</div>'
                show_register_form = True
            else:
                ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
                ua = request.headers.get("User-Agent", "")
                existing = find_registrant(event_id, phone, email)
                if existing is not None:
                    target = existing
                    conn = get_conn()
                    cur = conn.cursor()
                    cur.execute("UPDATE registrants SET has_webull_account = COALESCE(?, has_webull_account), webull_opened = COALESCE(?, webull_opened), organization = COALESCE(NULLIF(?, ''), organization) WHERE id = ?", (has_webull_account or None, webull_opened or None, organization, target["id"]))
                    conn.commit()
                    conn.close()
                    message = "前台补登记后签到（已存在名单）"
                else:
                    target = create_walkin_registrant(event_id, name, phone, email, organization, has_webull_account, webull_opened)
                    message = "前台现场登记并签到"
                ok, msg = insert_checkin(event_id, target["id"], phone, email, has_webull_account, "walkin_form", "success", message, ip, ua)
                if ok:
                    result = success_html("登记并签到成功", f"欢迎你，<strong>{name}</strong>", f'<div style="margin-top:8px;font-size:14px;color:#166534;">是否已完成 Webull 注册开户:{webull_opened}</div>')
                    show_register_form = False
                else:
                    result = f'<div class="err">{msg}</div>'
                    show_register_form = True

    register_block = ""
    if show_register_form:
        register_block = f"""
        <div class="wx-card">
          <h3 style="margin:0 0 8px;">未匹配？请现场登记</h3>
          <p class="sub" style="margin-top:0;">当前活动将自动沿用，无需填写活动 ID。</p>
          <form method="post" class="grid" style="margin-top:12px;">
            <input type="hidden" name="action" value="register_walkin">
            <input type="hidden" name="event_id" value="{event_id}">
            <input name="name" placeholder="姓名" required>
            <input name="phone" placeholder="手机号（建议填写）">
            <input name="email" placeholder="邮箱（手机号不便时可填）">
            <input name="organization" placeholder="单位 / 公司 / 机构（可选）">
            <div class="radio-group"><div class="helper">是否已有 Webull 账户？</div><div class="radio-row"><label><input type="radio" name="has_webull_account" value="是" required> 是</label><label><input type="radio" name="has_webull_account" value="否"> 否</label></div></div>
            <div class="radio-group"><div class="helper">是否已完成 Webull 注册开户？</div><div class="radio-row"><label><input type="radio" name="webull_opened" value="是" required> 是</label><label><input type="radio" name="webull_opened" value="否"> 否</label></div></div>
            <button class="wx-btn" type="submit">登记并签到</button>
          </form>
        </div>
        """

    body = f"""
    <div class="mini-phone">
      <div class="wx-header"><div class="badge">扫码签到</div><div class="wx-title">讲座签到</div><div class="wx-sub">活动 ID:{event_id or '未提供'}</div></div>
      <div class="wx-card">
        <h3 style="margin:0 0 8px;">填写签到信息</h3>
        <p class="sub" style="margin-top:0;">请输入报名时填写的手机号，若手机号不便输入，也可填写邮箱。</p>
        {result}
        <form method="post" class="grid" style="margin-top:12px;">
          <input type="hidden" name="action" value="checkin">
          <input type="hidden" name="event_id" value="{event_id}">
          <input name="phone" placeholder="请输入手机号">
          <input name="email" placeholder="或输入邮箱（可选）">
          <div class="radio-group"><div class="helper">您是否已经注册开户 Webull 账户？</div><div class="radio-row"><label><input type="radio" name="has_webull_account" value="是" required> 是</label><label><input type="radio" name="has_webull_account" value="否"> 否</label></div></div>
          <button class="wx-btn" type="submit">立即签到</button>
        </form>
      </div>
      {register_block}
      <div class="wx-card"><h3 style="margin:0 0 8px;">签到说明</h3><ul style="padding-left:18px; margin:8px 0; line-height:1.8;"><li>系统优先使用手机号匹配名单。</li><li>若手机号未匹配，再使用邮箱匹配。</li><li>已签到用户不能重复签到。</li><li>若提示未匹配，可直接填写现场登记表单。</li></ul></div>
    </div>
    """
    return page("讲座签到", body)


# -----------------------------
# API
# -----------------------------
@app.route("/api/checkin", methods=["POST"])
def api_checkin():
    data = request.get_json(silent=True) or {}
    event_id = str(data.get("event_id", "")).strip()
    phone = normalize_phone(data.get("phone", ""))
    email = normalize_email(data.get("email", ""))
    has_webull_account = normalize_yes_no(data.get("has_webull_account", ""))
    if not event_id:
        return jsonify({"ok": False, "message": "缺少 event_id"}), 400
    if not phone and not email:
        return jsonify({"ok": False, "message": "请填写手机号或邮箱"}), 400
    if not has_webull_account:
        return jsonify({"ok": False, "message": "请填写是否已有 Webull 账户"}), 400
    registrant = find_registrant(event_id, phone, email)
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
    ua = request.headers.get("User-Agent", "")
    if registrant is None:
        insert_checkin(event_id, None, phone, email, has_webull_account, "api", "failed", "名单未匹配", ip, ua)
        return jsonify({"ok": False, "message": "未在名单中找到该用户，请转入现场登记"}), 404
    ok, msg = insert_checkin(event_id, registrant["id"], phone, email, has_webull_account, "api", "success", "签到成功", ip, ua)
    if not ok:
        return jsonify({"ok": False, "message": msg}), 409
    return jsonify({"ok": True, "message": "签到成功", "name": registrant["name"], "organization": registrant["organization"], "has_webull_account": has_webull_account})


application = app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False, use_reloader=False)
