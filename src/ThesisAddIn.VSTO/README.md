# Word VSTO 宿主

本目录包含一个最小 Word VSTO 插件项目、跨目标框架的协议契约和独立的 .NET 10 协议宿主：

- <code>ThesisAddIn.VSTO.csproj</code>：Windows + 桌面版 Word 的 .NET Framework 4.8 VSTO 外壳，只引用 Protocol 项目，不引用 Core。
- <code>Protocol/</code>：版本化 JSON Lines 请求/响应、错误码和进程客户端，目标 <code>netstandard2.0</code>。
- <code>ProtocolHost/</code>：目标 <code>net10.0</code> 的控制台宿主，通过进程边界调用 Core 的只读分析服务。

协议字段、错误处理和联调命令见 <code>docs/word-engine-protocol.md</code>。Windows/Word 实机验证记录见 <code>docs/vsto-environment-validation.md</code>。

## 当前环境状态（2026-10-16）

当前主机为 Windows 11 25H2 x64，安装了 Microsoft Word 桌面版 16.0.20430.20146 x64、Visual Studio Community 2026 18.5.11723.231 和 .NET Framework 4.8 targeting pack。尚未检测到 Visual Studio Office/SharePoint development 工作负载、VSTO MSBuild targets 或 VSTO runtime，因此本机无法构建或加载此插件。协议宿主和协议自动化测试由 CI 验证；这不代表 Word 加载项已通过实机验收。

未修改 Office 全局安全配置。Issue #50 在 VSTO 构建、Word 加载、启用/禁用及启动/退出检查完成前保持打开。
