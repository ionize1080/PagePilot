# 半便携部署与更新

运行环境为 64 位 Windows、桌面 Word 和 .NET Framework 4.8。加载项采用 Office 2010 PIA / AnyCPU；Python 与 Electron 随包提供。目标机不需要 PowerShell、Python、Node、.NET SDK 或 VSTO Runtime。

## 首次启用

1. 完整解压便携包到固定的新目录，例如 `D:\Apps\PagePilot\0.4.1`。
2. 用平时运行 Word 的 Windows 账户打开 `PagePilot.Manager.exe`，通常无需管理员权限。
3. 点击“启用此版本”。Word 正在运行时，可请求正常关闭并处理保存提示；无响应时可确认强制关闭后继续。强制关闭会丢失未保存修改，只处理当前账户、当前桌面会话的 Word。
4. 点击“启动 Word 并验证”。只有当前版本同时完成连接和功能区加载才显示“已验证”。注册成功本身不代表已经加载。
5. 使用 Word 的 PagePilot 选项卡。模型设置仍使用 Chatbox 原有界面。

`Enable.cmd`、`Disable.cmd` 直接调用同一部署器，不依赖 PowerShell。旧 PowerShell 辅助脚本也委托管理器，统一使用相同校验、事务和日志。

管理器根据实际 WINWORD.EXE 架构选择 32/64 位注册视图，只写当前用户 PagePilot COM 类和 Word 加载项。它不创建服务、开机启动项、GAC、MSI 或 PATH 项，不清空 Office 禁用列表。

## 更新与存储

推荐按版本分别解压。关闭 Word 后，从新目录启用；不要覆盖使用中的 DLL。启用前校验全部运行文件，随后保存原注册、切换并回读。重复启用已正确注册的版本会保留已有回退点。

| 数据 | 默认位置 |
|---|---|
| 运行配置 | `%LOCALAPPDATA%\PagePilot\runtime.json` |
| 模型配置、API Key、会话 | `%APPDATA%\PagePilot\Chatbox-1.20.3` |
| 设置窗口缓存 | `%APPDATA%\PagePilot\Settings` |
| 偏好、标题规则 | `%LOCALAPPDATA%\PagePilot` |
| 文档修改备份 | `%LOCALAPPDATA%\PagePilot\backups` |
| 详细事件日志 | `%LOCALAPPDATA%\PagePilot\logs` |
| 部署事务 | `%LOCALAPPDATA%\PagePilot\deployment`（固定） |

在“存储与日志”中可修改配置文件及各数据目录，可选择复制旧配置和会话，保留原文件且不覆盖已有目标文件。默认目录与旧版兼容，不要求重新填写模型参数。详见[配置与日志](CONFIGURATION-AND-LOGS.md)。发布包不包含用户数据。

## 回退、移动和停用

关闭 Word 后选择“回退软件版本”，恢复最近一次部署之前的注册快照。保留旧目录；旧 DLL 不存在或其他程序改变注册时会拒绝回退。Word 因加载失败将 LoadBehavior 从 3 改为 2 的已知变化允许回退。写入过程失败会尝试恢复原注册并记录错误；中断事务需先回退。

移动前先停用，移动后重新启用。停用检查当前注册是否属于此版本目录，避免旧目录误停新版。回退只恢复加载项注册，不撤销数据迁移，也不删除日志、配置或文档备份。

## 排查与命令行

点击“诊断详情”查看完整检查并导出诊断包；侧边栏和诊断页均有“打开日志文件夹”。验证超时后 Word 保持打开，可处理首次启动提示、检查 COM 加载项和禁用项目。以其他管理员账户启动不会替平时用户注册加载项。

```text
PagePilot.Manager.exe enable --root "D:\Apps\PagePilot\0.4.1" --verify
PagePilot.Manager.exe enable --config "D:\PagePilotData\runtime.json"
PagePilot.Manager.exe doctor --output "D:\PagePilotData\diagnostic.txt"
PagePilot.Manager.exe export-diagnostics --output "D:\PagePilotData\diagnostic.zip"
PagePilot.Manager.exe open-logs
PagePilot.Manager.exe rollback
```

`--config` 指定已有 JSON 文件；使用 enable 时记为当前用户的运行配置。`--output` 避免完成消息框，适合自动化。命令行不会隐式关闭 Word；确需强制关闭时，必须同时指定 `--force-close-word --accept-unsaved-loss`，表示接受未保存修改丢失。返回码 0 表示命令成功，1 表示失败；未带 `--verify` 的 enable 只确认注册。另支持 `status`、`disable`、`self-test`，后者只操作独立测试子键。

## 目标机验收

本次本机 Word 16 x64 加载已验证。历史日志证明旧版曾连接 Word 14.0，0.4.1 的 Win10 / Office 2010 完整实机验收仍待完成。目标机请检查首次启用、完全退出后再启动、各侧栏、模型配置与日志路径、关闭后更新和回退。实际 Word 2010 验收不能由模拟能力测试代替。
