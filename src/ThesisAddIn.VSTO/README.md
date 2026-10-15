# Word VSTO 宿主

本目录现在包含一个最小 Word VSTO 插件项目、跨目标框架的协议契约和独立的 .NET
10 协议宿主：

- ThesisAddIn.VSTO.csproj：Windows + 桌面版 Word 的 .NET Framework 4.8 VSTO
  外壳。它只引用 Protocol/ThesisAddIn.Protocol.csproj，不引用 Core。
- Protocol/：版本化 JSON Lines 请求/响应、错误码和进程客户端，目标 netstandard2.0。
- ProtocolHost/：目标 net10.0 的控制台宿主，通过进程边界调用 Core 的只读分析服务。

协议字段、错误处理和联调命令见 docs/word-engine-protocol.md。Windows/Word 实机
版本与启停验证见 docs/vsto-environment-validation.md。

当前开发环境没有 Windows、桌面版 Word 和 Visual Studio Office/SharePoint 工作负载，
因此 Word COM 加载、插件启用/禁用以及 Word 启停仍标记为阻塞；协议宿主和自动化协议
测试可在跨平台 CI 中验证。未修改 Office 全局安全配置。
