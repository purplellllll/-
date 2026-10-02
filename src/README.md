# 核心代码

本目录只放业务逻辑代码，入口脚本由上级 `scripts/` 目录调用。

| 文件 | 说明 |
| --- | --- |
| `recruitment_sync.py` | Gmail 简历同步主程序：读取邮件、解析简历、写入飞书表格、归档原件，并处理 Offer 审核与通知。 |
| `interview_scheduler.py` | 面试调度程序：分配面试官、发送面试时间确认卡片、监听确认结果并回写表格。 |

直接运行时请从项目根目录执行，例如：

```powershell
python src\recruitment_sync.py self-test
```
