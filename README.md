# PagePilot

AI Powered Word 增强工具。面向 Windows 10 21H1、Word 2010 的原生桌面加载项，集成固定版本 Chatbox 1.20.3。

## 当前验证状态

本机真实测试环境：Windows build 26300、Word 16.0.20430.20180、64 位。**Win10 21H1 / Word 2010 的实际运行验证尚未完成，不能把本机测试当成目标环境测试。**详细记录见 `docs/TESTING.md`。代码通过 Office 2010 PIA 编译，避免将 Office 新版特性作为基本功能前提。

## 使用

1. 将完整便携压缩包解压到固定目录，例如 `D:\APP\PagePilot`。
2. 关闭 Word，运行 `Enable.cmd`。它仅注册当前用户的 COM 类和 Word 加载项；没有 MSI、ClickOnce 或安装向导，不创建“已安装的应用”条目。
3. 启动 Word，打开 **PagePilot** 选项卡。可打开正则替换、自动导航或 AI 助手。
4. 在 Chatbox 原有设置中配置模型供应商与 API Key。密钥不包含在发布包中。
5. 要停用，关闭 Word 后运行 `Disable.cmd`。移动目录前先停用，移动后重新启用。

这是准便携程序：Word 自动加载必须保留当前用户注册项；Chatbox 设置和会话保存在 `%APPDATA%\PagePilot`，本地规则、日志和操作备份保存在 `%LOCALAPPDATA%\PagePilot`。不会修改其他加载项或创建系统启动任务。卸载列表中没有条目，不代表系统或管理工具无法检测进程和文件。

## 功能

* **统一图标**：翡翠、海蓝、石墨三套配色，可在设置中预览并即时切换；设置采用通用齿轮图形。

* **原生正则侧栏**：.NET 正则、大小写选项、选区范围、捕获组替换、命中预览与定位、执行前内容一致性校验、500ms 正则超时。`^` / `$` 按 Word 段落匹配。为保护结构，替换不跨段落、单元格或分页符，不插入这些控制符。正文与正文表格是初始范围，页眉/页脚等可通过相应 MCP 工具处理。
* **无侵入导航**：通过内置中文/数字/英文标题规则或 `级别|正则` 自定义规则提取段落位置；不写样式、不设大纲级别、不增加隐藏书签。正文变化后刷新。目录采用可折叠多级树，单击即可将标题定位到视图上部；顶部留白可在设置中调节。展开“AI 目录要求”可指定只识别大标题或识别全部层级。AI 可调用 `pagepilot_set_navigation` 返回结构。明确选择“应用大纲级别”或让 AI 应用层级时才修改文档，可另选 Word 原生多级编号。
* **AI 助手**：真正的 Chatbox 1.20.3，独立窗口与 Word 侧栏可随时切换并保留草稿，保留其原生供应商、会话、附件、模型与 MCP 设置。接入 word-mcp-live 的全部 120 个工具，另有 PagePilot 上下文、导航、层级、西文格式、表格标准化、原生题注、版本能力指导、对象检查与实时对象编辑工具，共 129 个。完整参数见 `docs/tool-inventory.json`。
* **批准**：默认每个修改工具手动批准，显示工具名和参数；在对话输入栏点击盾牌，二次确认后变为红色闪电、开启自动批准；仅对当前文档、本次 Word 会话有效，切换文档或重启后恢复手动。正则替换和手动应用大纲无需 AI 批准。模型没有修改批准模式的工具。拒绝后不执行。
* **恢复**：原生批量修改和实时 MCP 修改前写入包含未保存编辑的 Flat OPC XML 备份，可用 Word 打开。文件工具修改前复制文件备份。上游实时工具尽可能组合撤销；旧版 Word 不保证一次 Ctrl+Z 撤销整个操作。

## 通信

`Chatbox → stdio MCP → 当前用户专用 Windows 命名管道 → Word 主线程`。

PagePilot 不监听 HTTP/TCP 端口。命名管道采用当前用户 ACL、随机名称和每次启动生成的会话令牌。模型供应商 API 是用户配置的出站 HTTPS 请求。整个工具目录均通过 stdio 发布，不使用 SSE 或 streamable HTTP。

## 实现与边界

加载器采用 .NET Framework 4.8 COM Add-in 和 Office 原生 CustomTaskPane。正则/导航分别使用独立的 WPF 侧栏（COM 宿主通过 ElementHost 承载），**不是 VSTO 部署包**，从而不需要额外安装 VSTO Runtime。Word 2010 的 32/64 位加载使用 AnyCPU 和分别注册的用户 COM 视图；AI 运行组件为 x64，因此便携包需要 64 位 Windows。

MCP 的“全覆盖”指全部上游工具及参数可发现、可调用；不是声称 Word 2010 支持新版 Word 的每项能力。实时编辑与关闭文件编辑是两类工具，已打开文档优先使用 `word_live_*`。AI 每次操作前读取 `pagepilot_capabilities`，逐项获知可用性及前置条件；Word 2010 的批注回复/解决会提前拦截。详细覆盖说明见 `docs/MCP-COVERAGE.md`。上游 `add_digital_signature` 只生成签名说明/元数据，**不是基于证书的加密数字签名**。全文核稿是模型分析，不保证发现所有问题。

AI 提供可切换的侧栏与独立窗口；正则与导航各有独立图标和可独立开关的 WPF 侧栏；设置为左右分栏的独立 Electron 窗口，打开设置不会同时打开 Chatbox，可直接打开 Chatbox 原有模型设置。PDF 转换已修复上游会退出整个 Word 进程的问题。大型文档、复杂嵌套表格及多 Word 进程仍需扩展回归验证。

## 构建

需要 .NET SDK、.NET Framework 4.8 开发引用、Office 2010 PIA、Extensibility 引用、Node.js 和 Python 3.12。`build.ps1` 编译加载项；`scripts/Prepare-Portable.ps1` 从锁定来源准备运行组件；`scripts/Package.py` 打包白名单内容。构建依赖不要求安装到目标机。

## 许可

GPL-3.0，详见 `LICENSE`、`THIRD-PARTY-NOTICES.md` 及随包保留的第三方许可证。
