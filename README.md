# 讲座签到系统

这是一个用于线下讲座、活动、路演场景的轻量级签到系统。系统支持后台上传报名名单，生成活动签到链接，参会者扫码后通过手机号或邮箱完成签到，后台可查看签到记录并导出 Excel。

## 功能概览

- 后台上传 Excel 报名名单
- 按活动 ID 管理不同讲座/活动
- 移动端 H5 签到页，适合二维码扫码访问
- 通过手机号优先匹配报名名单，邮箱作为备用匹配方式
- 防止同一参会者重复签到
- 后台查看签到记录
- 支持导出签到结果 Excel
- 使用 SQLite 本地数据库，部署和维护成本低

## 技术栈

- Python
- Flask
- SQLite
- pandas
- openpyxl
- HTML / CSS / JavaScript

## 项目结构

```text
.
├── app.py                  # Flask 主程序
├── data/
│   └── checkin_mini.db      # SQLite 数据库，首次运行后自动生成
├── requirements.txt         # 依赖文件，可选
└── README.md
```

如果项目中暂时没有 `requirements.txt`，可以自行创建：

```txt
flask
pandas
openpyxl
```

## 本地运行

### 1. 安装依赖

```bash
pip install flask pandas openpyxl
```

### 2. 启动项目

```bash
python app.py
```

当前代码默认监听：

```text
http://127.0.0.1:5001
```

后台入口：

```text
http://127.0.0.1:5001/admin
```

移动端签到页示例：

```text
http://127.0.0.1:5001/m/checkin?event_id=lecture_001
```

## 使用流程

### 1. 准备报名名单 Excel

Excel 至少需要包含以下任一字段：

- 手机号 / 手机 / 电话 / phone / mobile
- 邮箱 / email / mail / 电子邮箱

可选字段：

- 姓名 / name
- 单位 / 机构 / 公司 / organization

示例：

| 姓名 | 手机号 | 邮箱 | 单位 |
|---|---|---|---|
| 张三 | 91234567 | zhangsan@example.com | ABC Company |
| 李四 | 92345678 | lisi@example.com | XYZ Company |

### 2. 后台导入名单

进入后台：

```text
/admin
```

填写活动 ID，例如：

```text
lecture_001
```

上传 Excel 后，系统会将名单写入数据库。每次重新上传同一活动 ID 的名单时，系统会清空该活动原有名单和签到记录，并重新导入。

### 3. 生成签到二维码

每场活动对应一个签到链接：

```text
/m/checkin?event_id=lecture_001
```

将完整链接生成二维码，现场让参会者扫码即可进入签到页。

### 4. 参会者签到

参会者扫码后输入手机号或邮箱：

- 系统优先使用手机号匹配报名名单
- 手机号未匹配时，再使用邮箱匹配
- 匹配成功后写入签到记录
- 已签到用户再次提交会提示重复签到
- 未匹配用户会记录为失败尝试，方便后台核查

### 5. 查看和导出签到记录

后台活动列表中可以进入：

```text
/admin/records?event_id=lecture_001
```

导出 Excel：

```text
/admin/export?event_id=lecture_001
```

## 数据库说明

系统使用 SQLite，本地数据库文件默认保存于：

```text
data/checkin_mini.db
```

主要数据表：

### registrants

用于保存报名名单。

核心字段：

- event_id：活动 ID
- name：姓名
- phone：手机号
- email：邮箱
- organization：单位/机构
- raw_json：原始 Excel 行数据
- created_at：导入时间

### checkins

用于保存签到记录。

核心字段：

- event_id：活动 ID
- registrant_id：报名用户 ID
- submitted_phone：用户提交的手机号
- submitted_email：用户提交的邮箱
- status：签到状态，success 或 failed
- message：签到说明
- ip：提交 IP
- user_agent：浏览器信息
- checked_in_at：签到时间

## 注意事项

1. **活动 ID 要保持一致**  
   上传名单时填写的 `event_id` 必须和二维码链接中的 `event_id` 一致，否则无法匹配名单。

2. **手机号会被标准化处理**  
   系统会自动去除空格、括号、横线等非数字字符。例如 `+852 9123 4567` 会被处理为 `85291234567`。

3. **重复签到不会重复计数**  
   同一个活动中，同一名报名用户只能成功签到一次。

4. **重新上传名单会清空该活动原有数据**  
   当前逻辑下，重新导入同一 `event_id` 的 Excel 会删除该活动原有报名名单和签到记录。

5. **生产环境请修改 secret_key**  
   `app.secret_key` 不应使用默认值，正式部署时应改为随机安全字符串。

## 部署建议

如果部署到 Render、Railway、Fly.io 等平台，需要注意：

- 确认运行命令为：

```bash
python app.py
```

- 当前代码使用 SQLite，本地文件数据库适合轻量级活动场景。
- 如果需要长期保存大量活动数据，建议迁移到 PostgreSQL。
- 如果平台会休眠或重启，需确认数据库文件是否持久化保存。

## 后续可扩展功能

- 管理员登录权限
- 工作人员扫码核验入口
- 手动补签功能
- 临时访客登记
- 活动时间段控制
- 多活动 dashboard
- 邮件通知 / 短信通知
- 数据库迁移到 PostgreSQL
- 更完整的前后端拆分
- 微信小程序原生版本

## 常见问题

### 1. 为什么扫码后显示未匹配？

常见原因：

- 活动 ID 不一致
- 用户输入的手机号与报名表手机号格式不一致
- 报名名单中没有该用户
- Excel 字段名未被系统正确识别

可以检查后台导入的 Excel 字段，确保包含手机号或邮箱字段。

### 2. 为什么重复提交会失败？

系统通过 `event_id + registrant_id` 做唯一约束。同一活动中同一用户只能成功签到一次，避免重复计数。

### 3. 可以只用手机号签到吗？

可以。邮箱不是必填项，只是作为手机号无法匹配时的备用方式。

### 4. 能否导出未签到名单？

当前版本导出的是签到记录。如果需要未签到名单，可以在后续版本中增加“报名名单与签到记录对比导出”功能。
