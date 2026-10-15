# Word 插件引擎进程协议

## 边界

src/ThesisAddIn.VSTO/Protocol/ 是插件与引擎共用的契约程序集，目标框架为
netstandard2.0，不引用 ThesisEngine.Core。Word VSTO 插件目标为 .NET Framework
4.8；ThesisEngine.ProtocolHost 目标为 net10.0，在另一个进程中引用 Core。插件不会
加载 net10.0 程序集。

协议使用标准输入/标准输出上的 JSON Lines：每行一个 UTF-8 JSON 请求，每个请求只
产生一行响应。宿主标准输出只能写协议响应，诊断日志应写标准错误。连接关闭后宿主
退出，不保存或修改输入文档。

## 版本 1.0

分析请求：

~~~json
{"protocolVersion":"1.0","requestId":"a1b2","operation":"analyze","inputPath":"C:\\work\\sample.docx"}
~~~

成功响应：

~~~json
{"protocolVersion":"1.0","requestId":"a1b2","ok":true,"result":{"diagnostics":[],"ruleVersion":"demo-0.1.0","coverage":"正文 run","unsupportedScopes":[]}}
~~~

失败响应：

~~~json
{"protocolVersion":"1.0","requestId":"a1b2","ok":false,"error":{"code":"timeout","message":"引擎未在规定时间内返回结果。"}}
~~~

requestId 必须原样回显，插件据此拒绝过期或串线的响应。protocolVersion 必须精确
匹配 1.0；不匹配时宿主不调用引擎，返回 protocol_version_mismatch。当前只实现
analyze，输入必须是存在的 .docx 文件。

错误码和处理方式：

| 错误码 | 触发条件 | 插件行为 |
| --- | --- | --- |
| invalid_request | JSON 字段、操作或输入路径无效 | 中文提示，保留 Word 当前文档 |
| protocol_version_mismatch | 契约版本不一致 | 停止调用，提示升级插件或宿主 |
| protocol_error | JSON 损坏、响应 ID 不匹配 | 停止调用，不采纳结果 |
| timeout | 超过客户端配置的超时 | 终止宿主进程，提示重试 |
| engine_exit | 宿主无法启动或提前退出 | 提示引擎不可用 |
| engine_error | Core 读取或分析失败 | 展示错误，不写入文档 |
| cancelled | 调用方取消 | 终止宿主进程 |

## 构建和手工联调

在安装 .NET 10 SDK 的机器上构建协议宿主：

~~~powershell
dotnet build src\ThesisAddIn.VSTO\ProtocolHost\ThesisEngine.ProtocolHost.csproj
dotnet run --project src\ThesisAddIn.VSTO\ProtocolHost\ThesisEngine.ProtocolHost.csproj -- --protocol-stdio
~~~

VSTO 项目使用 Visual Studio 的 Office/SharePoint development 工作负载、桌面版
Word、.NET Framework 4.8 targeting pack 和 VSTO runtime 构建。设置
THESIS_ENGINE_PROTOCOL_HOST 为宿主 DLL 的完整路径后启动 Word 插件；插件会在
ThisAddIn_Startup 中创建进程客户端。插件只引用 Protocol 项目，不引用 Core。

协议测试位于 tests/ThesisEngine.Tests/EngineProtocolTests.cs，包括 JSON 往返、宿主
分析、非法输入和超时终止。正式规则仍以 Core 当前的 demo 规则版本为准；本协议不把
demo 结果声明为论文合规结论。
