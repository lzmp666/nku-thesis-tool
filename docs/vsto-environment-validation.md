# VSTO 环境验证记录

## 本机安装与功能区入口复验（2026-10-20）

源码基于 main `d609051`，本轮修复分支为 `fix/issue-51-pane-entry`。以下结果仅覆盖本轮列出的操作，
后文带日期的记录保留为历史证据，不能替代当前版本的完整阶段验收。

- 使用 VS 2022 MSBuild 签名生成 Release 插件；`mage -Verify` 对部署清单和程序集清单均返回签名有效。
  新版安装在 `D:\NKUThesisTool\LocalSigned-2026-10-20`，当前用户 Word 注册项指向该目录，`LoadBehavior=3`。
- 重启桌面 Word 16.0.17932.21056（Office LTSC 2024 x64），打开安装目录中的合成样例。
  顶部「论文格式检查」选项卡可见，中文侧栏自动出现。
- 点击侧栏的关闭按钮后，点击「论文格式检查 → 打开检查面板」成功恢复同一文档的侧栏；
  再次分析得到预期的 3 条诊断。加载日志记录了功能区恢复操作及完整的引擎分析往返。
- 当前代码的 Release 自动测试通过 67/67（失败 0、跳过 0）。本轮没有执行单项/批量写入，
  不以该结果替代修复闭环的完整实机复测。
- 本轮只打开合成文件。分析前后输入 SHA-256 均为
  `B8A3DD37B94AAC2B0AB100F20698106369F22311542DBAF2457B271A4EF674CE`。
- 普通文档的分析目前不自动创建原生 Word 批注；批注联动使用预先生成的标注副本。
  纯悬停的当前构建复验、完整正式规则覆盖和 #52 的最终端到端验收仍未完成。

## 最新实机状态（2026-10-17；代码提交 `e906101`）

以下结果取代后文中早先“Office 激活阻断”“Word 窗口无法聚焦”“没有任务窗格”的历史记录：

| 项目 | 最新结果 |
| --- | --- |
| VSTO Release 构建 | 使用 `D:\vs2022` 下 VS 2022 MSBuild 和 VSTO targets 成功构建并签名，0 warning、0 error |
| 引擎自动化测试 | `ThesisEngine.Tests` Release：67/67 通过 |
| Word 启动 | 从 Word 起始页启动后打开合成 DOCX，中文插件侧栏自动出现；当前用户 `LoadBehavior` 保持为 3 |
| 多文档窗口 | 同一 Word 进程第二个文档窗口自动出现独立侧栏；首个文档显示 3 条诊断，第二个显示 1 条，切换窗口后各自结果保留 |
| 正文批注锚点 | 选中第一段高亮后激活 `demo/east-asian-font`，选中第二段后切换到 `demo/western-font`。实测的是点击/选中，不是悬停 |
| 原始测试文档 | 本轮仅分析和选择锚点，没有执行修复或保存；完整单项/批量修复另存记录见 [Word 插件实测记录](vsto-addin-verification.md) |
| 待办 | 新提交推送到 PR #65 后等待 CI 和复审；Issue #50 的 Word 内错误场景及完整启用/禁用流程仍需验收 |
| Office 全局安全配置 | 未修改 |

自动加载问题的原因是 Word 在无活动文档的起始页启动时，`Application.ActiveWindow` 可能抛出 COM 异常，导致 Office 把加载项的 `LoadBehavior` 从 3 降为 2。当前实现捕获无文档状态，等 Word 的窗口激活事件出现后再创建任务窗格；[Microsoft VSTO 文档](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins)说明启动加载失败会将 `LoadBehavior` 降为 2。

## 当前环境核查（2026-10-17）

| 项目 | 实测 / 检查结果 |
| --- | --- |
| 操作系统 | Windows 11 25H2，64 位，内部版本 26200.9711 |
| Microsoft Word | 桌面版 16.0.20430.20146，64 位；用户已确认 Office 激活恢复，Word 正打开合成样例；Computer Use 两次将目标窗口映射为 E: 盘资源管理器，尚未完成实机验证 |
| Visual Studio | Community 2022 17.14.37806.7，安装在 `D:\vs2022` |
| Office/SharePoint development 工作负载 | 已安装；`Microsoft.VisualStudio.Tools.Office.targets` 可用 |
| .NET Framework 4.8 targeting pack | 已检测到 |
| VSTO runtime | 10.0，已检测到 |
| .NET SDK | 10.0.401（协议宿主与测试） |
| 开发证书 | 已创建当前用户本机证书，不可导出；未使用或重新信任旧 PFX |
| VSTO 构建 | Debug 与 Release 构建通过；Release `.vsto` 与 `.dll.manifest` 均经 `mage.exe -Verify` 验证签名有效 |
| 自动化测试 | `ThesisEngine.Tests` Release：67/67 通过 |
| 当前 PR head 的 Word 加载/交互 | 已确认插件加载并显示中文任务窗格；引擎环境变量未被首次 Word 进程识别，重启后被 Office 登录/激活提示阻断；分析、定位、修复、另存尚未验证 |
| Office 全局安全配置 | 未修改 |

环境已从缺少 VSTO 工具链/签名问题推进到桌面 Word 实机交互待验收。Office 激活已恢复；构建、签名、任务窗格加载和协议宿主层合成样例验证均已通过。桌面用户环境已配置 THESIS_ENGINE_PROTOCOL_HOST 并广播变化。当前 Word 正打开合成样例，但计算机使用工具两次将 Word 请求映射为 E: 盘文件资源管理器，因此未安全执行 UI 操作；仍未验证当前 Word 进程读取引擎路径、通过 JSON Lines 调用宿主、插件启用/禁用及正常启动/退出。

## 已由自动化测试覆盖的协议能力

跨平台 CI 覆盖 JSON Lines 往返、协议宿主分析合成 DOCX、协议版本不匹配、引擎退出、超时终止，以及客户端配置的引擎可执行文件不存在时返回中文 <code>engine_exit</code>。这些测试不等同于 Word COM 插件实机测试。

## 完成 Issue #50 所需的后续实测

当前环境已具备构建所需组件。剩余验收是在桌面 Word 中加载当前构建并：

1. 确认加载项可启用/关闭，Word 正常启动和退出。
2. 打开合成 DOCX，通过版本化 JSON Lines 协议调用 .NET 10 宿主并核对结构化诊断。
3. 检查宿主不可启动、超时、提前退出和协议版本不匹配时显示中文提示。
4. 记录 Windows、Word、Visual Studio、.NET Framework、VSTO runtime 的版本与位数，以及当前构建提交和逐项结果。不得关闭或降低 Office 全局安全配置来绕过加载问题。

在完成上述实测前，Issue #50 保持打开；构建、签名和自动化测试的通过不替代 Word COM 插件实机验收。


## Office 激活后的增量复测（2026-10-17）

ProtocolHost Release --no-restore 构建通过（0 warning、0 error）。协议层正式合成样例分析返回 3 项诊断；plan/apply 将 2 项自动修复写入新 DOCX，保留 1 项人工处理；recheck 后仍剩 1 项人工处理，输入文件 SHA-256 未变。另一个格式样例的 5 项补丁写入新输出成功。完整自动测试在样例文件锁释放后通过 67/67；PR #65 Actions run #99 通过 Linux build/test 和 Windows VSTO Release 签名构建。

桌面 Word 的 UI 验收仍未完成：应用列表显示 Word 正打开 known_issues_sample 合成样例，但两次 Computer Use 窗口获取都将请求的 Word 窗口映射成 E: 盘资源管理器。未对错误窗口执行操作，并按安全恢复流程停止 UI 自动化。待能安全聚焦 Word 后，再验证宿主调用、诊断、定位、批注联动、单项/批量修复、另存与重开。

## 最新 Word 实机复测（2026-10-17，PR #65 当前代码树）

以下为本日完成的复测，取代本文件较早的“Word UI 尚未完成”状态。Word 实测和本机验证针对插件代码提交 5b515dc69f5b167d8c61a2f3b83586310ed6177f 的源文件内容；该代码已推送到 PR #65。

| 验收项 | 结果 |
| --- | --- |
| VSTO Release 构建 | 通过；使用本机 VS 2022 MSBuild 17 和当前用户开发证书 |
| 自动测试 | ThesisEngine.Tests 67/67 通过 |
| 从 Word 首页启动插件 | 通过；没有活动文档时延迟到窗口激活再创建侧栏，插件未被 Word 禁用 |
| 同进程多文档窗口 | 通过；每个窗口显示独立任务窗格和结果（示例分别为 3 项、1 项） |
| 正文锚点到建议卡片 | 点击/选中后切换到对应诊断通过；未宣称或验证纯悬停 |
| 原文定位与单项/批量修复闭环 | 合成 DOCX 验收通过；单项副本复查剩余 2 项，批量副本剩余 1 项人工确认；原文 SHA-256 不变 |
| GitHub Actions | run #103（当时 PR head d9551a4）的 build-test 与 vsto-build 均成功；文档后续更新触发的新提交仍需 CI |
| Word 界面错误路径及加载项启用/禁用 | 尚未完成；属于 Issue #50 剩余验收 |

新增窗口行为的根因是启动阶段无文档时访问 Application.ActiveWindow 会抛出 COM 异常，Office 随后会降低 VSTO LoadBehavior。当前代码捕获此状态并等待 WindowActivate，再按 Word 窗口创建和维护侧栏。多窗口实测确认此修复有效。

因此 #51 的主流程和 #63 的点击/选中回退已有实机证据；#50 仍开放至 Word 界面错误路径和启用/禁用测试结束。PR #65 合并前仍需等待最新提交的 Actions 和维护者复审。#52 的端到端说明需在 #51 合并后按 Issue 范围单独整理；现有 PR #67 范围过宽且不可合并。
