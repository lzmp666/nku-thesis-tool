# VSTO 环境验证记录

## 当前环境核查（2026-10-17）

| 项目 | 当前记录 |
| --- | --- |
| Windows | 25H2，build 26200.9711；注册表 ProductName 显示 Windows 10 Home China |
| Word | 桌面版 16.0.20430.20146，64 位（来自 2026-10-16 的环境记录，本轮未重新启动 Word 核验） |
| Visual Studio | Community 2022 安装在 `D:\\vs2022`；本轮 MSBuild 版本为 17.14.60.43110 |
| Office/SharePoint 开发工具 | 已安装；`D:\\vs2022\\MSBuild\\Microsoft\\VisualStudio\\v17.0\\OfficeTools\\Microsoft.VisualStudio.Tools.Office.targets` 存在 |
| .NET Framework 4.8 targeting pack | 已检测到 |
| VSTO Runtime 10.0 | 已检测到 `VSTOInstaller.exe` 与 `VSTOLoader.dll` |
| 本机开发签名证书 | 当前用户证书库中尚无项目开发证书；生成证书时本受限构建会话返回 `0x80070005 ACCESS_DENIED` |
| VSTO 完整签名构建 | 阻塞：不能在当前会话写入证书库，尚未生成可安装的已签名清单 |
| 当前 PR #67 的 Word 实测 | 待完成；不能以历史 #51 测试或 CI 通过替代 |

本机已具备在 D 盘安装的 VS2022、Office/SharePoint VSTO targets、.NET Framework 4.8 targeting pack 和 VSTO Runtime。当前会话尝试创建开发签名证书时被访问控制拒绝；因此尚未证明 PR #67 可产出可安装的已签名 VSTO 清单，也未在 Word 中加载当前分支。

## 自动化测试与 CI 的边界

PR #67 原始 head `8374cb98a6dd850a6efc1355235b99585a8abb51` 本地运行 `dotnet test tests/ThesisEngine.Tests/ThesisEngine.Tests.csproj -c Release --no-restore -m:1 -p:UseSharedCompilation=false`，通过 68/68。CI run 83 通过，但 workflow 只构建 .NET 解决方案，不包含 `ThesisAddIn.VSTO.csproj`。这些结果覆盖引擎、协议宿主和自动化用例，不证明 VSTO 工程可签名、加载，也不替代 Word UI 验收。

## Issue #52 的待验收流程

仅使用合成样例 `tests/fixtures/synthetic/demo/known_issues_sample.docx`，不得使用真实论文或个人信息：

1. 在具备 VSTO targets 的 Windows 上生成本机开发证书并构建当前 PR head 的插件，记录提交 SHA 与构建输出。
2. 在桌面 Word 中打开样例并分析，确认三条演示诊断与原文位置对应。
3. 单独修复 `demo/font-size`；确认预览只包含字号，另存输出可在 Word 中打开且无修复提示，重新分析后仅字号问题消失。
4. 重新打开原始样例执行批量修复；确认只应用 `auto` 项、西文字体建议保留，重新分析输出与预览一致。
5. 核对原始样例 SHA-256 未变化，记录 Windows、Word、VS/MSBuild、VSTO Runtime、插件 head SHA、协议与规则版本。

完成上述实测前，Issue #52 保持打开，PR 不应合并；任何未通过项需记录具体原因，不用静态截图或 CI 结果代替。

## 历史 #51 实测记录

`docs/vsto-addin-verification.md` 保存了 2026-10-16 在早期 #51 代码上的 Word UI 记录。它仅说明当时记录的版本与样例，不代表 PR #67 当前 head 已重复通过；两份记录不应混作同一轮验收。
