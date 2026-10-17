# VSTO 环境验证记录

## 当前环境核查（2026-10-17）

| 项目 | 实测 / 检查结果 |
| --- | --- |
| 操作系统 | Windows 11 25H2，64 位，内部版本 26200.9711 |
| Microsoft Word | 桌面版 16.0.20430.20146，64 位；当前未运行。此前记录显示 Office 未激活，处于精简功能模式 |
| Visual Studio | Community 2022 17.14.37806.7，安装在 `D:\vs2022` |
| Office/SharePoint development 工作负载 | 已安装；`Microsoft.VisualStudio.Tools.Office.targets` 可用 |
| .NET Framework 4.8 targeting pack | 已检测到 |
| VSTO runtime | 10.0，已检测到 |
| .NET SDK | 10.0.401（协议宿主与测试） |
| 开发证书 | 已创建当前用户本机证书，不可导出；未使用或重新信任旧 PFX |
| VSTO 构建 | Debug 与 Release 构建通过；Release `.vsto` 与 `.dll.manifest` 均经 `mage.exe -Verify` 验证签名有效 |
| 自动化测试 | `ThesisEngine.Tests` Release：67/67 通过 |
| 当前 PR head 的 Word 加载/交互 | 尚未复测；待完成当前构建在桌面 Word 中的逐项验收 |
| Office 全局安全配置 | 未修改 |

环境阻塞已从“缺少 VSTO 工具链/签名失败”推进到“当前 PR 构建尚未在 Word 加载和验收”。构建与清单签名通过，不代表加载项在 Word 中已通过验收。Office 未激活状态可能限制部分 Word 功能，实测时如遇此限制需记录具体功能与结果，不改变 Office 全局安全设置。

## 已由自动化测试覆盖的协议能力

跨平台 CI 覆盖 JSON Lines 往返、协议宿主分析合成 DOCX、协议版本不匹配、引擎退出、超时终止，以及客户端配置的引擎可执行文件不存在时返回中文 <code>engine_exit</code>。这些测试不等同于 Word COM 插件实机测试。

## 完成 Issue #50 所需的后续实测

当前环境已具备构建所需组件。剩余验收是在桌面 Word 中加载当前构建并：

1. 确认加载项可启用/关闭，Word 正常启动和退出。
2. 打开合成 DOCX，通过版本化 JSON Lines 协议调用 .NET 10 宿主并核对结构化诊断。
3. 检查宿主不可启动、超时、提前退出和协议版本不匹配时显示中文提示。
4. 记录 Windows、Word、Visual Studio、.NET Framework、VSTO runtime 的版本与位数，以及当前构建提交和逐项结果。不得关闭或降低 Office 全局安全配置来绕过加载问题。

在完成上述实测前，Issue #50 保持打开；构建、签名和自动化测试的通过不替代 Word COM 插件实机验收。
