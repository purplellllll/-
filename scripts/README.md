# Windows 脚本

本目录存放 Windows 启动、授权、计划任务和隐藏窗口包装脚本。脚本会自动回到项目根目录，调用 `src/` 中的 Python 程序和根目录的 `config.json`。

| 文件 | 说明 |
| --- | --- |
| `run.ps1` | 统一命令入口：授权、连接检查、简历同步、状态查询、面试分发和通知扫描。 |
| `run-interview-listener.ps1` | 启动飞书长连接监听器，处理面试卡片回传。 |
| `install-schedule.ps1` | 安装或更新定时简历同步任务。 |
| `install-interview-listener.ps1` | 安装或更新面试卡片监听任务。 |
| `install-interview-notice-schedule.ps1` | 安装或更新面试通知扫描任务。 |
| `authorize-gmail-hidden.vbs` | 隐藏 PowerShell 窗口启动 Gmail 授权。 |
| `run-hidden.vbs` | 隐藏窗口运行简历同步。 |
| `run-interview-listener-hidden.vbs` | 隐藏窗口运行面试监听器。 |
| `run-interview-notices-hidden.vbs` | 隐藏窗口运行面试通知扫描。 |

推荐从项目根目录运行：

```powershell
.\scripts\run.ps1 check
```
