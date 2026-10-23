# PagePilot v0.1.0

首个正式版本：AI Powered Word 增强工具，固定集成 Chatbox 1.20.3。

## 功能

- 正则替换、自动导航各有独立功能区图标和独立 WPF 侧栏，可分别关闭或同时打开。
- 正则查找预览、捕获组替换、选区范围、内容一致性检查与超时保护。
- 本地规则或 AI 提取导航结构；普通导航不修改文字、样式、大纲级别或书签。明确应用层级时可生成 Word 原生导航与多级编号。
- AI 助手使用完整 Chatbox 界面，接入 word-mcp-live 全部 120 个工具与 6 个 PagePilot 扩展。支持核稿、字体、段落、页面、表格、批注、修订、原生题注等操作。
- Electron 设置窗口复用 Chatbox 运行组件，管理手动/自动批准与标题规则，并可直接打开 Chatbox 原有模型设置。
- MCP 仅用 stdio，Word 通信用当前用户专用命名管道，不监听本地 HTTP/TCP 端口。

## 使用与更新

完整解压 `PagePilot-0.1.0-portable-win-x64.zip` 到固定目录。关闭 Word 后运行 `Enable.cmd`，下次启动 Word 自动加载。没有安装向导、MSI、ClickOnce 或系统软件列表条目；仅写入当前用户 COM/Word 加载项注册项。停用用 `Disable.cmd`。

配置与会话在 `%APPDATA%\PagePilot`，规则、日志和备份在 `%LOCALAPPDATA%\PagePilot`，发布包不含 API Key 或用户会话。更新前关闭 Word，保留上述数据目录，以新版完整便携目录重新启用。

## 实测与边界

**实际测试：Windows build 26300、Word 16.0.20430.20180 x64。目标 Win10 21H1 + Word 2010 暂无测试环境，兼容性尚未经真实目标机验证。** 使用 Office 2010 PIA 编译和 .NET Framework 4.8；AI 组件要求 64 位 Windows。

已通过核心 11 项测试、真实 Word/stdin-stdout 工具调用、手动拒绝、Electron 自动批准、WPF 捕获组替换、无侵入导航、原生题注与 20 项格式/表格/批注/修订/PDF 回归检查。真实 DeepSeek API 与 Chatbox 界面均完成工具调用和核稿测试。观察到相关进程没有监听 TCP 端口。

126 项工具及参数完整发布，不代表 120 项上游工具的所有行为和边界均已逐一测试。Word 2010 无法支持的新版批注等接口仍需目标机验证；上游数字签名工具只是元数据/可见签名说明，不是证书加密签名。大型文档、复杂嵌套表格和多窗口生命周期仍需扩展验证。构建成功；NuGet 在线漏洞数据获取在本机不可用，未将其计为通过的检查。详见 [验证记录](TESTING.md)。

源码包包含 PagePilot、上游 Word MCP 源码及固定 Chatbox 1.20.3 完整源码归档；许可证与校验值随包提供。
