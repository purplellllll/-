# 配置模板

本目录只保存可公开提交的配置模板，不保存真实密钥和授权文件。

| 文件 | 说明 |
| --- | --- |
| `config.example.json` | 配置示例。复制到项目根目录并命名为 `config.json`，再填写飞书、Gmail、面试和通知配置。 |

本机配置文件放在项目根目录：

```powershell
Copy-Item .\config\config.example.json .\config.json
```

Gmail OAuth 客户端文件放在 `secrets/gmail-client.json`；`config.json`、`secrets/`、运行数据和日志均不应提交到 Git。
