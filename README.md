# 招新 Gmail → 飞书同步器

此本地程序同步候选人的 `序号、姓名、学号、邮箱、电话、专业、应聘组别、面试链接` 到飞书电子表格。每封 Gmail 邮件只会成功写入一次；本地 SQLite 会保存已处理的 Gmail message ID，避免定时任务重复写入。

## 文件说明

| 文件 | 用途 |
| --- | --- |
| `recruitment_sync.py` | 核心同步程序：读取 Gmail 简历、解析候选人信息、写入飞书表格、发送 Offer 邮件和审核卡片。 |
| `interview_scheduler.py` | 面试调度程序：分配面试官、发送面试时间确认卡片、接收确认结果并回写表格。 |
| `config.example.json` | 配置模板。复制为 `config.json` 后填写 Gmail、飞书和面试相关配置。 |
| `requirements.txt` | Python 依赖列表。 |
| `run.ps1` | 主入口，用于授权 Gmail、检查连接、同步简历和发送通知。 |
| `install-schedule.ps1` | 安装定时同步任务。 |
| `install-interview-listener.ps1` | 安装飞书面试卡片监听服务。 |
| `install-interview-notice-schedule.ps1` | 安装面试通知扫描定时任务。 |
| `run-hidden.vbs` | 隐藏窗口运行简历同步，供定时任务调用。 |
| `run-interview-listener.ps1` | 启动飞书面试卡片监听服务。 |
| `run-interview-listener-hidden.vbs` | 隐藏窗口启动面试监听服务。 |
| `run-interview-notices-hidden.vbs` | 隐藏窗口启动面试通知扫描。 |
| `authorize-gmail-hidden.vbs` | 隐藏窗口启动 Gmail 授权流程。 |
| `.gitignore` | 防止配置密钥、OAuth 文件、候选人资料、数据库和日志被提交。 |
| `README.md` | 项目说明、安装配置和使用方法。 |

### 常用入口

```powershell
# 首次授权 Gmail
powershell -ExecutionPolicy Bypass -File .\run.ps1 authorize-gmail

# 检查 Gmail 和飞书连接
powershell -ExecutionPolicy Bypass -File .\run.ps1 check

# 执行一次简历同步
powershell -ExecutionPolicy Bypass -File .\run.ps1 sync
```

## 一次性配置

1. 将 Google Cloud 下载的 **Desktop app** OAuth JSON 放到 `secrets/gmail-client.json`。
2. 复制 `config.example.json` 为 `config.json`，仅在本机填写飞书 App ID、App Secret、电子表格 token 和通知联系人信息。
3. 在 Gmail 创建用户标签 `招聘/待同步`，并让招新邮件自动或手动打上该标签。
4. 在飞书电子表格的首行创建下列列标题：`序号`、`姓名`、`学号`、`应聘组别`、`邮箱`、`电话`、`专业`、`面试时间`、`面试状态`、`面试官`、`面试链接`、`简历原件`。若实际列名不同，在 `config.json` 的 `feishu.fields` 和 `interview.fields` 中对应修改。`简历原件` 列也会在首次成功归档时自动创建。

若该电子表格存在多个可见工作表，在 `config.json` 的 `feishu.sheet_id` 填入目标工作表 ID；`check` 会在错误信息中列出可选值。只有一个可见工作表时无需填写。
5. 使用 PowerShell 运行：

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\run.ps1 authorize-gmail
   powershell -ExecutionPolicy Bypass -File .\run.ps1 check
   ```

首次授权会打开浏览器；使用招新 Gmail 账号允许读取和发送邮件。`check` 会验证 Gmail 标签、飞书凭证和目标电子表格表头，不会写入数据。

## 运行

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1 sync
powershell -ExecutionPolicy Bypass -File .\run.ps1 status
```

## 填写面试链接后的 Offer 审核与发送

面试官在某一行的 `面试链接` 填入有效的 `https://` 或 `http://` 会议链接后，系统不会立即向候选人发送邮件。机器人会先向唯一的简历审核接收人私信完整 Offer 预览和审核编号。

审核人核对后，在与机器人的单聊中发送 `/批准Offer OFFER-xxxx` 才会触发招新 Gmail 向该行 `邮箱` 发送面试通知；发送 `/拒绝Offer OFFER-xxxx` 则永久不发送。通知正文使用 `config.json` 中配置的组织、联系人和面试信息模板。候选人邮箱为空、格式无效，或链接不是网页链接时，程序不会发送邮件，并会把原因写入本地日志。

同一候选人、同一邮箱、同一链接只发送一次；面试官修改链接后，程序会自动补发一封更新通知。需要立即扫描一次可手动运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1 send-interview-notices
```

此前只授予“只读 Gmail”权限时，必须重新执行一次 `authorize-gmail` 并在浏览器中允许“发送邮件”，旧授权不会自动获得该权限：

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1 authorize-gmail
```

不想看到 PowerShell 窗口时，直接双击 `authorize-gmail-hidden.vbs`；浏览器仍会打开，完成授权后脚本会自行退出。

## 飞书面试时间确认

程序会从指定群的全部人类成员中随机抽取一位面试官，并仅向这一个人私聊发送邀请和候选人的原始 PDF 或 DOCX 简历。机器人只读取群成员作为面试官名单，不会在群内发布面试邀请或简历；飞书的群成员接口本身会过滤机器人。被抽中的面试官在与机器人的单聊中选择时间，程序仅接受落在设定时段内的时间，并写回 `面试时间`、`面试状态`、`面试官` 三列。随后面试官自行预约腾讯会议、填写 `面试链接`；面试由该面试官与一位负责招新的同学共同参与，并在表格最后的 `分数` 列共同记录结果。

在 `config.json` 的 `interview` 中配置允许时段；`recipient_group` 保持为空，机器人被首次拉入群聊时会自动写入该群：

```json
"recipient_group": {
  "chat_id": "",
  "name": "【广研内部】项目讨论群"
},
"time_windows": [
  {"start": "2026-08-06 00:00", "end": "2026-12-31 12:00"}
]
```

在飞书开发者后台为当前自建应用：添加 **机器人**能力；申请并发布 `im:message:send_as_bot`、`im:message.p2p_msg:readonly`、`im:message.group_at_msg:readonly`、`im:chat.members:read`、`im:chat.members:bot_access` 权限；在 **事件与回调** 中选择 **使用长连接接收事件**，订阅 `im.message.receive_v1` 和 **机器人进群** `im.chat.member.bot.added_v1`。在 **回调配置** 中选择 **使用长连接接收回调**，再添加 **卡片回传交互** `card.action.trigger`。在 **可用范围** 中覆盖名单群里的所有面试官。然后把机器人加入名单群并安装监听器：

```powershell
python -m pip install -r requirements.txt
powershell -ExecutionPolicy Bypass -File .\install-interview-listener.ps1
```

安装脚本优先创建隐藏计划任务；如果当前 Windows 账户不允许创建任务，则会自动改为当前用户的隐藏开机启动项。

把机器人首次拉入目标名单群后，它会自动绑定该群，并在群内发送完整流程和面试官操作说明；不需要发送任何命令。后续即使机器人又被拉入其他群，也不会覆盖已有绑定。之后所有邀请和时间确认都只在单聊里进行。确认 `time_windows` 已填写后，将 `interview.enabled` 改为 `true`，再运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1 dispatch-interviews
```

如需把当前已同步的测试简历也补发给名单群成员进行联调，运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1 queue-existing-interviews
```

后续每封新简历会自动向被随机选中的面试官发送预约卡片和原始简历文件。卡片固定显示允许的年份，面试官只能选择卡片提供的月/日和时间，然后点击“确认面试时间”。默认会优先发送附件名中含“简历”、`resume` 或 `cv` 的 PDF、DOC、DOCX；若有多份符合文件则只发最符合的一份，单文件超过 30MB 会跳过并记录本地日志。此功能要求在飞书应用的“权限管理”中添加并发布 `im:resource`（获取与上传图片或文件资源）；如需关闭它，可在 `config.json` 将 `interview.send_resume_with_invitation` 设为 `false`。

如需同步检查每份发给面试官的简历，并审核候选人 Offer，请在与机器人的**单聊**中发送 `/设置简历审核`。机器人会将该账号安全地保存为唯一审核接收人，并在后续每次给面试官发送原始简历时，同步私发一份给该账号；候选人 Offer 也会先私发给该账号审核，绝不会发到群内。发送 `/取消简历审核` 可随时停止。已在该命令之前发送的历史简历不会补发。

卡片日期的起点是该份邀请发起当天，终点由 `time_windows` 决定；例如邀请在 8 月 20 日发出且终点为 9 月 3 日时，只能选择 8 月 20 日至 9 月 3 日。

只有被随机抽中的面试官可以确认时间；确认结果会写回表格。`面试官`列保存面试官的显示姓名。

安装每两分钟自动运行一次的 Windows 任务：

```powershell
powershell -ExecutionPolicy Bypass -File .\install-schedule.ps1 -Minutes 2
```

取消自动任务：

```powershell
Unregister-ScheduledTask -TaskName 'Recruitment Gmail to Feishu Sync' -Confirm:$false
```

## 将原件同步到飞书表格

原始 PDF、DOCX 会继续保留在本机备份目录，同时上传到飞书云盘，并将每个文件的访问链接写入表格的 `简历原件` 列；`.doc` 和没有真实附件的邮件会留空。飞书开发者后台需要为当前自建应用添加并发布 **查看、评论、编辑和管理云空间中所有文件**（`drive:drive`）权限。程序会在应用云盘根目录自动创建“招新简历原件”文件夹，并保存其 token；无需人工创建或填写 token。

```json
"feishu": {
  "resume_archive": {
    "enabled": true,
    "folder_name": "招新简历原件",
    "folder_token": "",
    "file_url_base": "https://你的租户.feishu.cn"
  }
}
```

配置完成后运行一次，历史本地备份也会补写：

```powershell
python recruitment_sync.py sync-original-resumes
```

之后每次正常同步会自动完成原件上传和表格链接写入；原件上传失败仅会进入重试队列，不影响候选人入库或面试邀请。

## 简历格式规则

候选人可使用任意排版的邮件正文、`.docx`、`.doc` 或带有可复制文字的 PDF。程序会在表格和普通段落中识别字段标签及常见的学号、邮箱、手机号，不要求统一模板：

```text
姓名：张三
学号：2026123456
邮箱：zhangsan@example.com
电话：13800138000
专业：计算机科学与技术
```

邮箱只从简历附件或邮件正文中提取；Gmail 发件人邮箱不会被自动当作候选人邮箱写入。

姓名、学号、应聘组别、邮箱、电话、专业等任何无法明确识别的信息都会留空，绝不阻塞入库或向内部面试官发送邀请。`应聘组别`只接受“开发组”“测试组”“运营组”三个值；其他表述会留空。只有需要发出面试通知时，该行必须有有效的候选人邮箱和网页面试链接。

- `.doc` 由本机已安装的 Microsoft Word 读取；请保持 Word 可正常启动。
- PDF 需要含有可复制的文字。扫描件图片没有文字层，仍会提示人工处理；如需识别扫描件，可后续增加 OCR。
- PDF 解析依赖 `pypdf`。首次部署或更换 Python 后，运行 `python -m pip install -r requirements.txt`。

## 可选：大模型第二层识别

规则解析会先运行；仅当姓名、学号、邮箱、电话、专业或应聘组别中存在空值时，才会调用可选的大模型第二层。模型只能补全空值或格式异常值，返回结果仍会经过邮箱和手机号校验；模型不可用、超时或返回格式错误时，候选人仍会按已有规则结果入库并发送内部面试邀请。

程序兼容 OpenAI Chat Completions 风格的接口。为了不把简历中的个人信息发送到外部服务，优先使用本机模型服务。以 Ollama 为例，安装并下载你选择的模型后，把 `config.json` 的 `llm` 改为：

```json
"llm": {
  "enabled": true,
  "base_url": "http://127.0.0.1:11434/v1",
  "model": "qwen3:8b",
  "api_key_env": "",
  "timeout_seconds": 45,
  "max_input_characters": 12000
}
```

若改用云端 OpenAI 兼容服务，将 `base_url` 改为该服务的 `/v1` 地址，填写其模型名，并把密钥存入 Windows 用户环境变量而不是 `config.json`：

```powershell
[Environment]::SetEnvironmentVariable('RECRUITMENT_LLM_API_KEY', '你的密钥', 'User')
```

然后将 `api_key_env` 设为 `RECRUITMENT_LLM_API_KEY`、重新打开登录会话或重启电脑，再把 `enabled` 设为 `true`。配置读取发生在每次同步启动时，无需重新授权 Gmail 或飞书。

## 本地数据与安全

- `config.json`、`secrets/`、`data/`、`logs/` 已被 `.gitignore` 排除，不会被 Git 提交。
- OAuth refresh token、飞书密钥和候选人处理状态仅保存在本机。
- 每份成功入库的简历都会自动备份到 `data/candidate-backups/<Gmail消息ID>/`：其中 `candidate.json` 保存入库字段和飞书行标识，原始 `.pdf`、`.doc`、`.docx` 简历附件会一并保留。若备份临时失败，后续同步会自动重试，不会重复入库。
- 请不要把 `config.json`、OAuth JSON、`data/gmail-token.json` 或日志发送到聊天中。
