# Word 插件引擎进程协议

当前版本：**1.2**（1.1 新增 plan、apply 与 annotate；1.2 增加合并诊断的多位置数据和位置列表过滤）。

## 边界

src/ThesisAddIn.VSTO/Protocol/ 是插件与引擎共用的契约程序集，目标框架为
netstandard2.0，不引用 ThesisEngine.Core。Word VSTO 插件目标为 .NET Framework
4.8；ThesisEngine.ProtocolHost 目标为 net10.0，在另一个进程中引用 Core。插件不会
加载 net10.0 程序集。

协议使用标准输入/标准输出上的 JSON Lines：每行一个 UTF-8 JSON 请求，每个请求只
产生一行响应。宿主标准输出只能写协议响应，诊断日志应写标准错误。连接关闭后宿主
退出。

宿主本身只读取输入文档：analyze 与 plan 不写入任何文件；apply 和 annotate 只写入调用方指定的
临时输出文件，且**拒绝覆盖已存在的输出文件**。Word 插件再用临时结果刷新唯一的标注工作副本，
用户的原始提交文档不会被覆盖。

## 版本 1.0：analyze

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
匹配宿主版本；不匹配时宿主不调用引擎，返回 protocol_version_mismatch。输入必须是
存在的 .docx 文件。

## 版本 1.1 / 1.2：plan、apply 与 annotate

版本 1.2 在 1.1 的操作上增加 `locations` 请求过滤和 `occurrences` 诊断位置列表；
版本字段严格匹配，插件与协议宿主需要一同更新。

### plan：生成可审阅的修复计划

~~~json
{"protocolVersion":"1.2","requestId":"p1","operation":"plan","inputPath":"C:\\work\\sample.docx"}
~~~

可选过滤字段 `ruleId`、`location` 与 `locations` 用于“只计划当前这一项”。合并诊断会
通过 `locations` 提交该项的所有位置；旧式单位置请求仍受支持。过滤位置为空时计划全部
可自动修复项。

成功响应：

~~~json
{"protocolVersion":"1.2","requestId":"p1","ok":true,"plan":{"patches":[{"ruleId":"demo/east-asian-font","location":"word/document.xml#paragraph[0]/run[0]","expectedBefore":"黑体","desiredAfter":"宋体","reason":"将正文 run 的东亚字体改为规则目标值。"}],"deferred":[{"ruleId":"demo/western-font","severity":"warning","location":"word/document.xml#paragraph[1]/run[0]","expected":"Times New Roman","actual":"Arial","fixLevel":"suggest","message":"…","occurrences":[{"location":"word/document.xml#paragraph[1]/run[0]","actual":"Arial"}]}],"ruleVersion":"demo-0.1.0","coverage":"…","unsupportedScopes":[]}}
~~~

`patches` 按每个受影响的位置展开；`patches` 只包含处理分类为 auto 的项目；`deferred`
是留给用户的合并诊断，既包括引擎判定不能自动修复的，也包括被
`ruleId`/`location(s)` 过滤掉的 auto 项。插件必须先展示计划并
取得用户确认，再调用 apply。

### apply：执行修复

~~~json
{"protocolVersion":"1.2","requestId":"a1","operation":"apply","inputPath":"C:\\work\\sample.docx","outputPath":"C:\\work\\sample-fixed.docx","ruleId":"demo/east-asian-font","locations":["word/document.xml#paragraph[0]/run[0]","word/document.xml#paragraph[1]/run[0]"]}
~~~

引擎仍将完整修复结果写到调用方指定的临时 `.docx`，以便整批校验、原子失败和重新标注；
Word 插件会用该结果更新同一条标注工作副本路径，不再提示用户为每次修复另存一份。原始
提交文件保持不变。引擎宿主在应用前重新分析输入，且 `outputPath` 不得已存在；全部前置
条件在写入之前检查完毕，失败时不会产生部分修复文件。

成功响应：

~~~json
{"protocolVersion":"1.2","requestId":"a1","ok":true,"apply":{"outputPath":"C:\\work\\sample-fixed.docx","applied":[],"failures":[],"unresolved":[]}}
~~~

`failures` 非空时表示该批次未发布任何输出；`unresolved` 是被过滤掉、本次未处理的补丁。

### annotate：分析并生成带 Word 原生批注的新副本

~~~json
{"protocolVersion":"1.2","requestId":"n1","operation":"annotate","inputPath":"C:\\work\\sample.docx","outputPath":"C:\\work\\sample-annotated.docx"}
~~~

`outputPath` 必须是与输入不同且尚不存在的 `.docx` 文件。宿主分析输入文档，并在新副本中
为每条合并诊断添加一个 Word 原生批注，并将其所有位置黄色高亮；响应同时返回 `result`（诊断与规则覆盖信息）和
`annotation`（输出路径与批注数量）。即使没有诊断，也会产生未经修改的新副本。原文不被覆盖。

批注正文只显示约定的四项：要求参考来源、规则、期望值和实际值。插件所有权、稳定位置和原
有高亮状态写入 Word 隐藏文字，不会显示为额外字段。重复分析时，宿主仅移除带有本工具所有权
标记的旧批注，再按最新诊断生成批注；普通用户批注、未带所有权标记的旧批注和原有高亮都会
保留。不能安全识别或还原的插件元数据会使整个操作失败，不发布半成品文件。

诊断位置的 run 序号只计算正文段落中包含文字的 run，Word 批注引用等无文字结构不会改变位置。

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
| patch_conflict | 输出文件已存在，或修复前置条件冲突 | 中文提示，保留原文件并要求更换文件名 |

## 构建和手工联调

在安装 .NET 10 SDK 的机器上构建协议宿主：

~~~powershell
dotnet build src\ThesisAddIn.VSTO\ProtocolHost\ThesisEngine.ProtocolHost.csproj
dotnet run --project src\ThesisAddIn.VSTO\ProtocolHost\ThesisEngine.ProtocolHost.csproj -- --protocol-stdio
~~~

VSTO 项目需要 Visual Studio 的 Office/SharePoint development 工作负载、桌面版
Word、.NET Framework 4.8 targeting pack 和 VSTO runtime。它是经典（非 SDK 风格）
工程，必须用 Visual Studio 自带的 MSBuild 构建；工程内会自动调用 dotnet CLI 构建
netstandard2.0 的契约程序集。完整命令见 src/ThesisAddIn.VSTO/README.md。

插件只引用 Protocol 程序集，不引用 Core。设置 THESIS_ENGINE_PROTOCOL_HOST 为宿主
DLL 的完整路径后启动 Word，插件会在 ThisAddIn_Startup 中创建进程客户端。

Word 会锁定正在编辑的 .docx，引擎无法直接读取，因此插件在分析或修复前从 Word 当前内存内容
生成临时快照，再以快照作为 inputPath。首次分析会生成一份标注工作副本；后续分析和修复通过
临时文件刷新同一工作副本路径，不会每次为用户新建文档。修复时插件关闭并重新打开该路径以
释放 Word 文件锁；作者完成修改后可自行保存或另存为。原始提交文件保持不变。

协议测试位于 tests/ThesisEngine.Tests/EngineProtocolTests.cs 和
WordAnnotationWriterTests.cs，包括 JSON 往返、宿主分析/批注、重复刷新、原文不变、用户批注保留、非法输入和超时终止。正式规则仍以 Core 当前的 demo 规则版本为准；本协议不把
demo 结果声明为论文合规结论。
