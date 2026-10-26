# 半便携部署与更新

运行环境为 64 位 Windows、已安装的桌面 Word 和 .NET Framework 4.8。加载项使用 Office 2010 PIA / AnyCPU；Python 与 Electron 运行组件随包提供。目标机不需要 PowerShell、Python、Node、.NET SDK 或 VSTO Runtime。

## 首次启用

1. 完整解压便携包到固定目录，例如 `D:\Apps\PagePilot\0.4.0-rc.1`。
2. 保存并关闭所有 Word 窗口。
3. 运行 `PagePilot.Manager.exe`，选择“启用此版本”。`Enable.cmd` 是调用同一个 exe 的可选快捷方式，不调用 PowerShell。
4. 打开 Word，使用 PagePilot 选项卡。模型设置仍使用原有 Chatbox 界面。

管理器根据实际 WINWORD.EXE 的 PE 架构选择 32/64 位注册视图，只写当前用户的 PagePilot COM 类和 Word 加载项。它不创建系统服务、启动项、GAC、MSI、卸载列表条目或 PATH 项，不清除 Office 的 DisabledItems。

## 同一电脑更新

推荐**按版本分别解压目录**。关闭 Word 后，从新目录的管理器启用新版；不要覆盖正在使用的文件。管理器先校验完整文件清单，再保存原注册状态、切换路径并回读。检测到其他组件占用相同类标识时会拒绝覆盖。

用户数据位置保持不变：

| 数据 | 路径 |
|---|---|
| 模型配置、API Key、会话 | `%APPDATA%\PagePilot\Chatbox-1.20.3` |
| 设置窗口缓存 | `%APPDATA%\PagePilot\Settings` |
| 配色、显示方式、导航参数、标题规则 | `%LOCALAPPDATA%\PagePilot` |
| 文档修改备份 | `%LOCALAPPDATA%\PagePilot\backups` |
| 部署事务记录 | `%LOCALAPPDATA%\PagePilot\deployment` |

启用、停用、更新和回退不会删除这些目录。此版本未改变配置格式，无需导出或重新填写模型参数。发布包按白名单构建，不包含这些用户数据。

## 回退、移动和停用

关闭 Word 后选择“回退注册”，恢复最近一次部署前的注册快照。请保留旧版本目录；旧 DLL 不存在、注册被其他操作改变时会拒绝自动回退。中断的部署会要求先恢复挂起事务。写入过程中出现错误会尝试恢复原状态并保留事务记录。

移动当前版本前先停用，移动后重新启用。停用会验证当前注册路径属于此目录，避免旧目录管理器误停用新版。仅停用加载项不清除个人数据。

完整性检查可用管理器“检查完整文件”。命令行也支持 `enable`、`disable`、`rollback`、`status`、`doctor`、`--root <目录>`、`--output <报告路径>`。`--self-test` 在独立测试子键中验证部署事务，不改实际 PagePilot 注册。

## 发布与目标机验收

0.4.0-rc.1 是候选版。已在本机 Word 16 x64 验证，但 Word 2010 / Win10 21H1 尚未实机验证。目标机应验证启用、Word 冷启动、各侧栏、132 项工具发现、常用对象编辑、关闭后更新和回退；成功后再将候选版升级为正式版。Office 2010 不支持的功能由能力接口明确降级。
