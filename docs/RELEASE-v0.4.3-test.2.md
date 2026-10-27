# PagePilot 0.4.3-test.2 · Chatbox 请求标识

沿用已验证可启用的 0.4.3 部署器和加载项，程序文件版本仍为 0.4.3.0。本包仅调整 Chatbox 启动时的默认 User-Agent，保留机器级替换注册和扩展诊断。

## 请求头变化

0.4.3-test.1 中，同样的 Chatbox 包元数据和随包 Electron 运行时在本地回环测试实际发送：

```text
Mozilla/5.0 (...) PagePilotChatbox/1.20.3 Chrome/152.0.7977.130 Electron/44.7.0 Safari/537.36
```

本包在加载 Chatbox 原有主程序前设置全局默认 User-Agent，测试实际收到：

```text
Mozilla/5.0 (...) Chrome/152.0.7977.130 Safari/537.36 Chatbox/1.20.3
```

仅替换默认客户端标识，移除 Electron 与 PagePilotChatbox 产品标记，保留系统与 Chromium 兼容字段。没有拦截或改写认证头、供应商专用头、请求正文、URL、模型参数或流式响应。主进程 Node/供应商 SDK 若自行设置其他 User-Agent，仍沿用其原行为。因此这不是“所有请求头逐字等于官方 Chatbox”，也不是更换运行时；应用内部仍使用 Electron。

## 验证

使用随包 Electron、相同应用名与版本在隐藏测试窗口中向 127.0.0.1 发请求，不连接模型服务，不使用真实密钥。已验证渲染进程 fetch 和 Electron net.fetch 的默认 UA 不含 Electron；认证头、自定义头与 JSON 正文保持完整。Chromium 忽略渲染脚本直接指定的 User-Agent 时，也使用新的默认 UA。实际供应商、企业代理或网关仍可能添加或改变请求头，目标机最终对外头应以服务端收到的内容为准。

运行日志增加 `network.identity`，记录默认 UA，便于确认启动的是新包，不记录请求密钥。

## 更新

关闭 Word 和 PagePilot 助手，完整解压到新目录。右键 `Replace-Machine.cmd` → 以管理员身份运行，再从新管理器“整台电脑”范围启动 Word 并验证。无需重新填写模型配置。旧目录可保留以供回退。
