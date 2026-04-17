from __future__ import annotations

import io
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    url_for,
)

# ============================================================
# 讲座签到系统（标准单文件版）
# 适用：
# 1. 本地开发
# 2. Render / Gunicorn 部署
# ============================================================

# -----------------------------
# 路径与应用
# -----------------------------
try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "checkin_mini.db"

app = Flask(__name__)
app.secret_key = "replace-this-with-a-random-secret-key"


# -----------------------------
# 数据库
# -----------------------------
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


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
            raw_json TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    # 兼容旧表：若没有 source 字段，则补上
    cur.execute("PRAGMA table_info(registrants)")
    cols = [row[1] for row in cur.fetchall()]
    if "source" not in cols:
        cur.execute("ALTER TABLE registrants ADD COLUMN source TEXT DEFAULT 'imported'")

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS checkins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id TEXT NOT NULL,
            registrant_id INTEGER,
            submitted_phone TEXT,
            submitted_email TEXT,
            status TEXT NOT NULL,
            message TEXT,
            ip TEXT,
            user_agent TEXT,
            checked_in_at TEXT NOT NULL,
            UNIQUE(event_id, registrant_id)
        )
        """
    )

    conn.commit()
    conn.close()


@app.before_request
def ensure_db() -> None:
    # 避免 Render / Gunicorn 启动阶段直接 crash
    init_db()


# -----------------------------
# 工具函数
# -----------------------------
def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_phone(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\D", "", str(value).strip())


def normalize_email(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


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


def import_excel_to_db(file_bytes: bytes, event_id: str) -> int:
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    name_col = pick_column(df.columns.tolist(), ["姓名", "名字", "name"])
    phone_col = pick_column(df.columns.tolist(), ["手机号", "手机", "电话", "phone", "mobile"])
    email_col = pick_column(df.columns.tolist(), ["邮箱", "email", "mail", "电子邮箱"])
    org_col = pick_column(df.columns.tolist(), ["单位", "机构", "公司", "organization"])

    if phone_col is None and email_col is None:
        raise ValueError("Excel 中至少要包含手机号或邮箱字段。")

    conn = get_conn()
    cur = conn.cursor()

    # 同一活动重新导入时，覆盖旧名单与旧签到
    cur.execute("DELETE FROM checkins WHERE event_id = ?", (event_id,))
    cur.execute("DELETE FROM registrants WHERE event_id = ?", (event_id,))

    count = 0
    for _, row in df.iterrows():
        name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
        phone = normalize_phone(row[phone_col]) if phone_col else ""
        email = normalize_email(row[email_col]) if email_col else ""
        organization = str(row[org_col]).strip() if org_col and pd.notna(row[org_col]) else ""

        if not phone and not email:
            continue

        cur.execute(
            """
            INSERT INTO registrants
            (event_id, name, phone, email, organization, source, raw_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                event_id,
                name,
                phone,
                email,
                organization,
                "imported",
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


def create_walkin_registrant(
    event_id: str,
    name: str,
    phone: str,
    email: str,
    organization: str,
):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO registrants
        (event_id, name, phone, email, organization, source, raw_json, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            event_id,
            name,
            phone,
            email,
            organization,
            "walkin",
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
                status, message, ip, user_agent, checked_in_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                event_id,
                registrant_id,
                submitted_phone,
                submitted_email,
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
# HTML 基础模板
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
  --bg:#f5f7fb;
  --card:#ffffff;
  --line:#e5e7eb;
  --text:#111827;
  --muted:#6b7280;
  --brand:#1677ff;
  --brand2:#0f5ed7;
  --ok:#16a34a;
  --okbg:#dcfce7;
  --err:#dc2626;
  --errbg:#fee2e2;
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
input, button, select {
  width:100%; border:1px solid #dbe3ef; border-radius:14px; padding:13px 14px; font-size:16px; background:#fff;
}
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
.mini-phone {
  max-width:430px; margin:0 auto; min-height:760px; background:linear-gradient(180deg,#eef5ff,#f8fbff);
  border-radius:28px; padding:18px; box-shadow:0 18px 38px rgba(0,0,0,.08);
}
.wx-header { padding:16px 8px 18px; text-align:center; }
.wx-title { font-size:22px; font-weight:800; }
.wx-sub { color:var(--muted); font-size:13px; margin-top:6px; }
.wx-card { background:#fff; border-radius:20px; padding:18px; box-shadow:0 8px 24px rgba(0,0,0,.06); margin-bottom:14px; }
.wx-btn { border-radius:16px; padding:14px; font-size:17px; }
.badge { display:inline-block; padding:4px 10px; border-radius:999px; background:#e8f1ff; color:var(--brand2); font-size:12px; font-weight:700; }
</style>
</head>
<body>
  <div class="page">{{ body|safe }}</div>
</body>
</html>
"""


def page(title: str, body: str):
    return render_template_string(BASE_HTML, title=title, body=body)


def make_msg_html(msg_text: str, msg_type: str = "ok") -> str:
    if not msg_text:
        return ""
    css = "ok" if msg_type == "ok" else "err"
    return f'<div class="{css}" style="margin-bottom:16px;">{msg_text}</div>'


# -----------------------------
# 路由：后台
# -----------------------------
@app.route("/")
def root():
    return redirect(url_for("admin"))


@app.route("/admin", methods=["GET", "POST"])
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
            return redirect(
                url_for(
                    "admin",
                    msg=f"导入成功，共导入 {count} 条名单。移动端签到链接：/m/checkin?event_id={event_id}",
                    msg_type="ok",
                )
            )
        except Exception as e:
            return redirect(url_for("admin", msg=f"导入失败：{e}", msg_type="err"))

    conn = get_conn()
    cur = conn.cursor()

    cur.execute(
        """
        SELECT r.event_id,
               COUNT(DISTINCT r.id) AS total_registrants,
               COUNT(DISTINCT CASE WHEN c.status='success' THEN c.registrant_id END) AS total_checked_in,
               COUNT(DISTINCT CASE WHEN c.status='failed' THEN c.id END) AS failed_attempts
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
        f"<td>"
        f"<a href='/m/checkin?event_id={r['event_id']}'>签到页</a> | "
        f"<a href='/admin/records?event_id={r['event_id']}'>签到记录</a> | "
        f"<a href='/admin/export?event_id={r['event_id']}'>导出</a> | "
        f"<a href='/admin/delete?event_id={r['event_id']}' onclick=\"return confirm('确定删除该活动所有数据？');\">删除</a>"
        f"</td>"
        f"</tr>"
        for r in rows
    )

    body = f"""
    <div class="topbar">
      <div>
        <div class="brand">讲座签到后台</div>
        <div class="sub">微信小程序式前台 + 后台管理界面</div>
      </div>
      <div><a href="/m/checkin?event_id=lecture_001">打开示例签到页</a></div>
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
          <button type="submit">导入名单</button>
        </form>
        <p class="sub">Excel 至少应包含手机号或邮箱字段。可额外包含姓名、单位等字段。</p>
      </div>

      <div class="card">
        <h3 style="margin-top:0;">现场手动补签</h3>
        <form method="post" action="/admin/manual_checkin" class="grid">
          <input name="event_id" placeholder="活动 ID，例如 lecture_001" required>
          <input name="phone" placeholder="手机号（优先匹配）">
          <input name="email" placeholder="邮箱（手机号不便时可填）">
          <button type="submit">手动补签</button>
        </form>
        <p class="sub">适用于现场来宾网络不稳定、二维码失效、工作人员代为登记等场景。系统同样会校验名单并防止重复签到。</p>
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
        <button type="submit">登记并补签</button>
      </form>
      <p class="sub">适用于现场临时来宾、未提前报名但允许入场的情况。系统会先检查该手机号/邮箱是否已在名单中；若没有，则新增为临时访客并立即签到，后续也会出现在导出表格里。</p>
    </div>

    <div class="card" style="margin-top:16px;">
      <h3 style="margin-top:0;">二维码使用方式</h3>
      <p>给每场讲座固定一个链接：</p>
      <p><code>/m/checkin?event_id=lecture_001</code></p>
      <p>把这个链接生成二维码，现场让参会者扫码即可。</p>
    </div>

    <div class="card" style="margin-top:16px;">
      <h3 style="margin-top:0;">活动列表</h3>
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>活动 ID</th><th>名单人数</th><th>成功签到</th><th>未匹配尝试</th><th>操作</th></tr>
          </thead>
          <tbody>{table_rows or '<tr><td colspan="5">暂无活动</td></tr>'}</tbody>
        </table>
      </div>
    </div>
    """
    return page("讲座签到后台", body)


@app.route("/admin/records")
def admin_records():
    event_id = request.args.get("event_id", "").strip()
    msg_text = request.args.get("msg", "").strip()
    msg_type = request.args.get("msg_type", "ok").strip()
    msg_html = make_msg_html(msg_text, msg_type)

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT c.checked_in_at, c.status, c.message,
               r.name, r.phone, r.email, r.organization, r.source,
               c.submitted_phone, c.submitted_email
        FROM checkins c
        LEFT JOIN registrants r ON c.registrant_id = r.id
        WHERE c.event_id = ?
        ORDER BY c.checked_in_at DESC
        """,
        (event_id,),
    )
    rows = cur.fetchall()
    conn.close()

    table_rows = "".join(
        f"<tr>"
        f"<td>{x['checked_in_at']}</td>"
        f"<td>{x['status']}</td>"
        f"<td>{x['name'] or ''}</td>"
        f"<td>{(x['organization'] or '') + ('（临时访客）' if x['source'] == 'walkin' else '')}</td>"
        f"<td>{x['phone'] or x['submitted_phone'] or ''}</td>"
        f"<td>{x['email'] or x['submitted_email'] or ''}</td>"
        f"<td>{x['message'] or ''}</td>"
        f"</tr>"
        for x in rows
    )

    body = f"""
    <div class="topbar">
      <div>
        <div class="brand" style="font-size:24px;">签到记录</div>
        <div class="sub">活动 ID：{event_id}</div>
      </div>
      <div>
        <a href="/admin">返回后台</a> | <a href="/admin/export?event_id={event_id}">导出 Excel</a>
      </div>
    </div>

    {msg_html}

    <div class="card">
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>签到时间</th><th>状态</th><th>姓名</th><th>单位</th><th>手机号</th><th>邮箱</th><th>说明</th></tr>
          </thead>
          <tbody>{table_rows or '<tr><td colspan="7">暂无记录</td></tr>'}</tbody>
        </table>
      </div>
    </div>
    """
    return page("签到记录", body)


@app.route("/admin/export")
def admin_export():
    event_id = request.args.get("event_id", "").strip()

    conn = get_conn()
    query = """
        SELECT c.event_id AS 活动ID,
               c.checked_in_at AS 签到时间,
               c.status AS 状态,
               c.message AS 说明,
               r.name AS 姓名,
               r.organization AS 单位,
               r.source AS 来源,
               r.phone AS 名单手机号,
               r.email AS 名单邮箱,
               c.submitted_phone AS 提交手机号,
               c.submitted_email AS 提交邮箱
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

    return redirect(url_for("admin", msg=f"已删除活动：{event_id}", msg_type="ok"))


@app.route("/admin/manual_checkin", methods=["POST"])
def admin_manual_checkin():
    event_id = request.form.get("event_id", "").strip()
    phone = normalize_phone(request.form.get("phone", ""))
    email = normalize_email(request.form.get("email", ""))

    if not event_id:
        return redirect(url_for("admin", msg="请先填写活动ID。", msg_type="err"))
    if not phone and not email:
        return redirect(url_for("admin", msg="请至少填写手机号或邮箱。", msg_type="err"))

    registrant = find_registrant(event_id, phone, email)
    if registrant is None:
        return redirect(
            url_for(
                "admin",
                msg=f"未在活动 {event_id} 的名单中找到该手机号/邮箱，请检查活动ID或输入信息。",
                msg_type="err",
            )
        )

    ok, insert_msg = insert_checkin(
        event_id=event_id,
        registrant_id=registrant["id"],
        submitted_phone=phone,
        submitted_email=email,
        status="success",
        message="后台手动补签",
        ip="admin-manual",
        user_agent="admin-manual",
    )

    if not ok:
        return redirect(url_for("admin_records", event_id=event_id, msg=insert_msg, msg_type="err"))

    display_name = registrant["name"] or "该参会者"
    return redirect(
        url_for(
            "admin_records",
            event_id=event_id,
            msg=f"手动补签成功：{display_name}",
            msg_type="ok",
        )
    )


@app.route("/admin/walkin_checkin", methods=["POST"])
def admin_walkin_checkin():
    event_id = request.form.get("event_id", "").strip()
    name = request.form.get("name", "").strip()
    phone = normalize_phone(request.form.get("phone", ""))
    email = normalize_email(request.form.get("email", ""))
    organization = request.form.get("organization", "").strip()

    if not event_id:
        return redirect(url_for("admin", msg="请先填写活动ID。", msg_type="err"))
    if not name:
        return redirect(url_for("admin", msg="请填写临时访客姓名。", msg_type="err"))
    if not phone and not email:
        return redirect(url_for("admin", msg="临时访客至少需要手机号或邮箱其一。", msg_type="err"))

    existing = find_registrant(event_id, phone, email)
    if existing is not None:
        ok, insert_msg = insert_checkin(
            event_id=event_id,
            registrant_id=existing["id"],
            submitted_phone=phone,
            submitted_email=email,
            status="success",
            message="后台补签（已存在名单）",
            ip="admin-walkin",
            user_agent="admin-walkin",
        )
        if not ok:
            return redirect(url_for("admin_records", event_id=event_id, msg=insert_msg, msg_type="err"))

        display_name = existing["name"] or name
        return redirect(
            url_for(
                "admin_records",
                event_id=event_id,
                msg=f"已在名单中找到该来宾，补签成功：{display_name}",
                msg_type="ok",
            )
        )

    walkin = create_walkin_registrant(event_id, name, phone, email, organization)
    ok, insert_msg = insert_checkin(
        event_id=event_id,
        registrant_id=walkin["id"],
        submitted_phone=phone,
        submitted_email=email,
        status="success",
        message="临时访客现场登记并签到",
        ip="admin-walkin",
        user_agent="admin-walkin",
    )

    if not ok:
        return redirect(url_for("admin_records", event_id=event_id, msg=insert_msg, msg_type="err"))

    return redirect(
        url_for(
            "admin_records",
            event_id=event_id,
            msg=f"临时访客登记并补签成功：{name}",
            msg_type="ok",
        )
    )


# -----------------------------
# 路由：移动端签到
# -----------------------------
@app.route("/m/checkin", methods=["GET", "POST"])
def mobile_checkin():
    event_id = request.args.get("event_id", "").strip() or request.form.get("event_id", "").strip()
    result = ""

    if request.method == "POST":
        phone = normalize_phone(request.form.get("phone", ""))
        email = normalize_email(request.form.get("email", ""))

        if not event_id:
            result = '<div class="err">缺少活动 ID，请联系工作人员。</div>'
        elif not phone and not email:
            result = '<div class="err">请输入手机号或邮箱。</div>'
        else:
            registrant = find_registrant(event_id, phone, email)
            ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
            ua = request.headers.get("User-Agent", "")

            if registrant is None:
                insert_checkin(
                    event_id=event_id,
                    registrant_id=None,
                    submitted_phone=phone,
                    submitted_email=email,
                    status="failed",
                    message="名单未匹配",
                    ip=ip,
                    user_agent=ua,
                )
                result = '<div class="err">未在名单中匹配到该手机号或邮箱，请联系现场工作人员。</div>'
            else:
                ok, msg = insert_checkin(
                    event_id=event_id,
                    registrant_id=registrant["id"],
                    submitted_phone=phone,
                    submitted_email=email,
                    status="success",
                    message="签到成功",
                    ip=ip,
                    user_agent=ua,
                )
                if ok:
                    display_name = registrant["name"] or "参会者"
                    org = registrant["organization"] or ""
                    result = (
                        '<div style="background:#ecfdf5;border:2px solid #22c55e;border-radius:18px;padding:18px 16px;text-align:center;margin-bottom:12px;">'
                        '<div style="font-size:44px;line-height:1;margin-bottom:8px;">✅</div>'
                        '<div style="font-size:24px;font-weight:800;color:#166534;">签到成功</div>'
                        f'<div style="margin-top:10px;font-size:16px;color:#14532d;">欢迎你，<strong>{display_name}</strong></div>'
                        f'{f"<div style=\"margin-top:6px;font-size:14px;color:#166534;\">单位：{org}</div>" if org else ""}'
                        '<div style="margin-top:10px;font-size:13px;color:#15803d;">系统已记录本次签到，请勿重复提交。</div>'
                        '</div>'
                    )
                else:
                    result = f'<div class="err">{msg}</div>'

    body = f"""
    <div class="mini-phone">
      <div class="wx-header">
        <div class="badge">扫码签到</div>
        <div class="wx-title">讲座签到</div>
        <div class="wx-sub">活动 ID：{event_id or '未提供'}</div>
      </div>

      <div class="wx-card">
        <h3 style="margin:0 0 8px;">填写签到信息</h3>
        <p class="sub" style="margin-top:0;">请输入报名时填写的手机号，若手机号不便输入，也可填写邮箱。</p>
        {result}
        <form method="post" class="grid" style="margin-top:12px;">
          <input type="hidden" name="event_id" value="{event_id}">
          <input name="phone" placeholder="请输入手机号">
          <input name="email" placeholder="或输入邮箱（可选）">
          <button class="wx-btn" type="submit">立即签到</button>
        </form>
      </div>

      <div class="wx-card">
        <h3 style="margin:0 0 8px;">签到说明</h3>
        <ul style="padding-left:18px; margin:8px 0; line-height:1.8;">
          <li>系统优先使用手机号匹配名单。</li>
          <li>若手机号未匹配，再使用邮箱匹配。</li>
          <li>已签到用户不能重复签到。</li>
          <li>若提示未匹配，请联系现场工作人员处理。</li>
        </ul>
      </div>
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

    if not event_id:
        return jsonify({"ok": False, "message": "缺少 event_id"}), 400
    if not phone and not email:
        return jsonify({"ok": False, "message": "请填写手机号或邮箱"}), 400

    registrant = find_registrant(event_id, phone, email)
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
    ua = request.headers.get("User-Agent", "")

    if registrant is None:
        insert_checkin(event_id, None, phone, email, "failed", "名单未匹配", ip, ua)
        return jsonify({"ok": False, "message": "未在名单中找到该用户"}), 404

    ok, msg = insert_checkin(event_id, registrant["id"], phone, email, "success", "签到成功", ip, ua)
    if not ok:
        return jsonify({"ok": False, "message": msg}), 409

    return jsonify(
        {
            "ok": True,
            "message": "签到成功",
            "name": registrant["name"],
            "organization": registrant["organization"],
        }
    )


# 供 gunicorn / wsgi 使用
application = app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False, use_reloader=False)