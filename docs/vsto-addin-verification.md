# Word 插件实测记录（Issue #51）

> **验收边界：** 下方标注为 2026-10-16 的 Word 操作记录来自此前构建，不能代替当前 PR head 的复测。当前 head 已完成构建和自动化验证；是否通过桌面 Word 交互验收，见最新记录。

## 当前复测结论（2026-10-17；代码提交 `e906101`）

本节记录本次最新实测，覆盖前文中“尚未加载引擎”“新窗口没有侧栏”“原文锚点联动未验证”等旧状态。旧记录保留作历史，不代表当前状态。

- Release VSTO 使用 `D:\vs2022\MSBuild\Current\Bin\MSBuild.exe` 和 VS 2022 VSTO targets 构建、签名成功；无警告、无错误。`ThesisEngine.Tests` Release 测试通过 **67/67**。
- 修复了 Word 从起始页启动时没有活动文档窗口的场景：启动期读取 `Application.ActiveWindow` 可能抛 COM 异常，Office 会把加载项 `LoadBehavior` 从 3 降为 2。现在捕获该状态并等 `WindowActivate` 创建窗格。官方 VSTO 文档也说明，启动加载出错会把该值改为 2：[VSTO 加载项注册项](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins)。
- Word 从起始页重启后打开合成样例 `known_issues_sample.docx`，中文侧栏自动出现，分析列出 **3 条**问题；注册值保持 `LoadBehavior=3`。
- 同一 Word 进程再打开合成副本 `known_issues_batch_fixed.docx`，第二个窗口也自动出现独立侧栏，分析列出 **1 条**待人工确认问题；切回第一个窗口仍保留 **3 条**，两个窗口的结果没有串线。
- 在合成标注文件 `known_issues.annotated.docx` 中点选正文锚点：第一段切换到 `demo/east-asian-font` 建议，第二段切换到 `demo/western-font` 建议。侧栏问题到正文的定位此前已通过实测。Word 对象模型没有纯悬停事件；这里验证的是点击/选中回退，不称为悬停。
- 本次只做分析和选区联动，没有改动测试文档内容。单项/批量修复、预览、另存及重分析的实测结果见下方“Word 实机验收更新”。

代码已提交到本机分支 `codex/issue-51-window-lifecycle`，待推送到 PR #65 的分支并等待新 CI 与复审。PR #65、Issue #51/#63 暂不关闭；Issue #50 仍需完成 Word 内错误场景和启用/禁用的完整验收。

## PR #65 最新本机复核（2026-10-17）

本机编译工作树为 `098928b`；`ThisAddIn.cs`、`EngineGateway.cs` 与 `DiagnosticsPane.cs` 的 blob 与当前 PR head `6194bce` 一致。PR base `main` 为 `9b7e10a`。
VS Community 2022 17.14.37806.7 安装在 `D:\vs2022`；Office/SharePoint 开发工作负载、
VSTO targets、.NET Framework 4.8 targeting pack 与 VSTO Runtime 10.0 均可用。

在 Windows PowerShell 5.1 中运行 `scripts/New-VstoDevCertificate.ps1` 成功创建/复用当前用户
本机开发证书。脚本需 UTF-8 BOM，否则 PowerShell 5.1 会误读中文字符串；该兼容修复已纳入本分支。
VS 2022 MSBuild 的 Debug 与 Release VSTO 构建均成功；`mage.exe -Verify` 确认 Release 输出的
`ThesisAddIn.VSTO.vsto` 和 `ThesisAddIn.VSTO.dll.manifest` 均有有效签名。本机
`dotnet test tests/ThesisEngine.Tests/ThesisEngine.Tests.csproj --configuration Release --no-restore`
通过 **67/67**。

当前 PR #65 head `6194bce` 对应的本机 Release 产物已在桌面 Word 中成功加载，并显示中文「论文格式检查」任务窗格。打开合成样例 `known_issues_sample.docx` 后，任务窗格初始提示未配置 `THESIS_ENGINE_PROTOCOL_HOST`，因此本轮尚未执行分析、定位、单项/批量修复或另存操作。随后已将引擎宿主 DLL 路径写入当前用户环境变量并通知 Windows 更新环境；重启 Word 后，Office 的登录/激活提示阻止继续打开样例。当前 Office 标题为“未经授权产品”，本轮不登录账户、不绕过激活。样例 SHA-256 为 `B8A3DD37B94AAC2B0AB100F20698106369F22311542DBAF2457B271A4EF674CE`，与仓库预期值一致；本轮没有对样例执行修改。

PR 的 CI run 87（run ID `39015449160`）通过，但它仍不构建 `ThesisAddIn.VSTO.csproj`。VSTO Debug/Release 构建与 manifest 签名已通过，插件加载和任务窗格显示已验证；协议调用和完整交互闭环仍未验证。Issue #51 保持打开，PR 保持 Draft。
安全复核发现此前提交的 `ThesisAddIn.VSTO_TemporaryKey.pfx` 确实包含私钥，当前源码树已移除并忽略 PFX；由于旧文件曾出现在公开历史中，必须视为已泄露，不再信任或复用。所有演示样例均为合成 DOCX，不含真实论文或个人信息。

## Word 交互实测环境（此前一次记录）

| 项目 | 实测值 |
| --- | --- |
| 操作系统 | Windows 11 25H2，64 位，内部版本 26200.9457 |
| Microsoft Word | 桌面版 16.0.20430.20146，64 位；**未激活**，标题栏显示“未经授权产品”，打开文档时显示“兼容模式” |
| Visual Studio | Community 2022 17.14.36616.10，本轮为验证新装了 Office/SharePoint 开发工作负载 |
| .NET SDK | 10.0.401（构建与测试）、8.0.425 |
| VSTO runtime | 10.0（本机已存在） |
| 构建方式 | 协议宿主用 `dotnet build`；插件用 Visual Studio 自带 MSBuild（原因见 `src/ThesisAddIn.VSTO/README.md`） |

## 实测步骤与结果（2026-10-16）

样例：`tests/fixtures/synthetic/demo/known_issues_sample.docx`，含 3 条已知问题。

| # | 操作 | 实测结果 |
| --- | --- | --- |
| 1 | 打开样例后点「分析当前文档」 | 侧边栏列出 3 条诊断，显示规则编号、位置、期望值、实际值、严重度与处理方式；面板显示规则版本 `demo-0.1.0`、已检查范围与未覆盖范围 |
| 2 | 点击列表中的问题 | Word 选中对应段落（选中即定位，未修改文档内容） |
| 3 | 选中 `demo/east-asian-font` 后点「修复此项」 | 预览窗口显示“将修改 1 处，保留 2 处”；确认并另存到 `Documents\论文格式检查-已修复.docx`；重新分析当前文档后该条消失 |
| 4 | 选中 `demo/western-font` 后点「修复此项」 | 插件拒绝自动修复并提示“该问题不能自动修复（处理方式：需确认）”，与 #49 确立的安全边界一致 |
| 5 | 点「预览并修复全部」 | 预览显示“将修改 2 处，保留 1 处”；确认后另存 `Documents\论文格式检查-已修复-1016-223733.docx`；面板状态显示“修复完成：已应用 2 处（demo/east-asian-font、demo/font-size）”，并重新分析当前文档 |
| 6 | 再次修复并沿用已存在的文件名 | 引擎拒绝写入并给出中文提示要求更换文件名——验证“不覆盖已有文件”生效 |
| 7 | 检查原始文件 | `known_issues_sample.docx` 与仓库工作区自始至终未被改动 |

修复结果复核：对第 5 步输出的文件重新运行 analyze，仅剩 1 条 `demo/western-font`
（期望 Times New Roman，实际 Arial，处理方式 suggest），与预览中的“保留 1 处”一致。

## 实测中发现并修复的问题

1. **插件工程无法构建**：原工程是 SDK 风格，需要 MSBuild 18，而 VSTO 任务只能用 VS 自带的
   MSBuild 17，两者互斥。已改为经典工程格式，并在工程内自动调用 dotnet CLI 构建契约程序集。
2. **`Document.SaveCopyAs` 在本机被拒绝**：未激活的 Office 返回 `0x800A1704`。插件改为
   优先使用 `SaveCopyAs`（可包含未保存修改），失败时退回文件复制，并在面板状态中写明本次
   分析的是“当前内容”还是“已保存的文件内容”。
3. **plan/apply 的成功响应被误判为协议异常**：客户端只校验 `result` 字段。已改为按操作校验
   对应载荷，并补了回归测试。
4. **任务窗格按钮被裁掉**：「预览并修复全部」在窄窗格中不可见，按钮区已允许换行。
5. **修复后列表与当前文档不一致**：原先重新分析的是输出文件。已改为重新分析当前文档，
   使列表始终对应用户正在编辑的内容。
6. **默认输出文件名重复**：固定名字会让第二次修复撞上“不覆盖已有文件”的安全规则，
   默认文件名已加入时间戳。

## 已知限制

- Word 对象模型不提供鼠标悬停事件：验收要求的“从正文批注锚点激活建议卡片”实现为
  **点击/选中锚点触发**，不宣称支持纯悬停。
- 定位按段落序号映射：引擎只统计正文段落，含表格的文档可能错位，此时插件给出提示而不
  选中错误内容。
- 处理方式为 suggest / manual 的问题不会被自动修复，只列出供人工处理。
- 当前规则范围仅覆盖演示规则（正文东亚字体、西文字体、字号），不代表论文合规结论。
- 未在其他 Word 版本上验证，也未提供安装包/部署流程。
- 测试套件依赖 `tests/fixtures/synthetic/demo/` 下的样例文件：Word 打开这些文件会锁定它们，
  导致相关用例失败。运行测试前请关闭 Word；持续集成环境无此问题。



## Office 激活后的增量复测（2026-10-17）

用户已确认 Office 激活恢复。桌面用户环境中的 THESIS_ENGINE_PROTOCOL_HOST 已指向本机 Release 协议宿主，并已广播环境变更。ProtocolHost Release --no-restore 构建通过（0 warning、0 error）。协议层正式合成样例分析返回 3 项诊断；2 项自动修复另存至新 DOCX，保留 1 项人工处理，recheck 后剩余 1 项，输入文件 SHA-256 未变。另一个格式样例的 5 项补丁写入新输出成功。完整自动测试最终通过 67/67；PR #65 Actions run #99 的 Linux build/test 和 Windows VSTO Release 签名构建均通过。

当前 Word 正打开合成样例，但两次 Computer Use 窗口获取都将请求的 Word 窗口映射为 E: 盘文件资源管理器；未操作错误窗口，并按安全恢复流程停止 UI 自动化。因此尚未验证当前 Word 进程读取引擎路径、引擎调用、诊断定位、批注联动、单项/批量修复及另存/重开。Office 激活已不再是阻塞项；PR #65 保持 Draft，Issue #50/#51 保持打开。
