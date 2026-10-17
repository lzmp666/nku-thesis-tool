# Word 插件实测记录（Issue #51）

> **验收边界：** 下方标注为 2026-10-16 的 Word 操作记录来自此前构建，不能代替当前 PR head 的复测。当前 head 已完成构建和自动化验证；是否通过桌面 Word 交互验收，见最新记录。

## PR #65 最新本机复核（2026-10-17）

复核代码基于 PR head `67f8c925`，并合入 `main`（`9b7e10a`），本地集成提交为 `21d290e`。
VS Community 2022 17.14.37806.7 安装在 `D:\vs2022`；Office/SharePoint 开发工作负载、
VSTO targets、.NET Framework 4.8 targeting pack 与 VSTO Runtime 10.0 均可用。

在 Windows PowerShell 5.1 中运行 `scripts/New-VstoDevCertificate.ps1` 成功创建/复用当前用户
本机开发证书。脚本需 UTF-8 BOM，否则 PowerShell 5.1 会误读中文字符串；该兼容修复已纳入本分支。
VS 2022 MSBuild 的 Debug 与 Release VSTO 构建均成功；`mage.exe -Verify` 确认 Release 输出的
`ThesisAddIn.VSTO.vsto` 和 `ThesisAddIn.VSTO.dll.manifest` 均有有效签名。本机
`dotnet test tests/ThesisEngine.Tests/ThesisEngine.Tests.csproj --configuration Release --no-restore`
通过 **67/67**。

当前 PR head 的构建产物尚未在桌面 Word 中加载；侧边栏、锚点点击联动、定位、单项/批量修复和
另存结果还没有针对这份构建逐项复测。GitHub Actions 只构建协议、引擎宿主和自动化测试，不包含
`ThesisAddIn.VSTO.csproj`。在完成 Word 实测前，#51 保持打开且 PR 保持 Draft。

安全复核发现此前提交的 `ThesisAddIn.VSTO_TemporaryKey.pfx` **确实包含私钥**，原 README 的
“不包含任何真实密钥”表述错误。当前 PR 补丁已从工作树移除该 PFX，并改为由开发者在本机证书库生成
不可导出的调试私钥；由于旧 PFX 曾出现在公开分支历史中，请将旧证书视为已泄露，不再信任或复用。

记录在真实桌面 Word 中运行中文问题侧边栏的实测过程与结果。全部样例均为仓库内的
合成演示文件，不含真实论文或个人信息。

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

