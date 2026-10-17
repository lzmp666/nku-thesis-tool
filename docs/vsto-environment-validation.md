# VSTO 环境验证记录

## 当前环境核查（2026-10-16）

| 项目 | 实测 / 检查结果 |
| --- | --- |
| 操作系统 | Windows 11 25H2，64 位，内部版本 26200.9711（注册表 ProductName 仍显示 Windows 10 Home China） |
| Microsoft Word | 桌面版 16.0.20430.20146，64 位；本次核查时 Word 未运行 |
| Visual Studio | Visual Studio Community 2026，18.5.11723.231 |
| Office/SharePoint development 工作负载 | 未检测到；Visual Studio Installer 的 Office workload 查询没有匹配安装 |
| .NET Framework 4.8 targeting pack | 已检测到 |
| VSTO MSBuild targets | 未检测到；Visual Studio 18.0/17.0 OfficeTools targets 路径均不存在 |
| VSTO runtime | 未检测到；对应运行时注册表项和卸载项均未找到 |
| Word COM 加载项构建、启用/禁用 | 未测试：缺少 Office workload、VSTO targets 和 VSTO runtime |
| Word 加载项启动/退出 | 未测试：插件无法构建和加载 |
| Office 全局安全配置 | 未修改 |

当前机器具备 Windows、桌面版 Word、Visual Studio 和 .NET Framework 4.8 targeting pack，但缺少 VSTO 开发与运行组件。不能据此宣称最小插件可构建、加载或通过 Word 实机验收。此结果是部分环境阻塞，不是“没有 Windows/Word”的阻塞。

## 已由自动化测试覆盖的协议能力

跨平台 CI 覆盖 JSON Lines 往返、协议宿主分析合成 DOCX、协议版本不匹配、引擎退出、超时终止，以及客户端配置的引擎可执行文件不存在时返回中文 <code>engine_exit</code>。这些测试不等同于 Word COM 插件实机测试。

## 完成 Issue #50 所需的后续实测

在安装 Office/SharePoint development 工作负载与 VSTO runtime 的 Windows 机器上：

1. 确认 .NET Framework 4.8 targeting pack 和 VSTO runtime 可用。
2. 构建 <code>src/ThesisAddIn.VSTO/ThesisAddIn.VSTO.csproj</code>，记录构建结果和生成的清单。
3. 安装或调试最小 VSTO 插件，启动 Word，确认 COM Add-ins 中可启用和禁用该插件。
4. 打开合成 DOCX，通过版本化 JSON Lines 协议调用 .NET 10 宿主，核对结构化诊断。
5. 验证插件启用时 Word 可正常启动和退出；宿主不可启动、超时、提前退出和协议版本不匹配时显示中文提示。
6. 记录 Windows、Word、Visual Studio、.NET Framework、VSTO runtime 的版本与位数，并附实测结果。不得关闭或降低 Office 全局安全配置来绕过加载问题。

在完成上述实测前，Issue #50 保持打开；协议实现和自动化测试可单独合入，不据此自动关闭总体验收任务。

## Issue #52 修复闭环验收记录（2026-10-16）

验收样例限定为 `tests/fixtures/synthetic/demo/known_issues_sample.docx`。自动化协议测试验证单项修复只应用所选规则和位置、批量计划只包含 `auto` 项、另存输出、输入字节不变，以及重新分析输出后已修复问题消失、不可自动修复的西文字体诊断仍保留。自动测试不等同于 Word 实机验收。

| 项目 | 当前记录 |
| --- | --- |
| Windows | Windows 11 25H2，内部版本 26200.9711（环境核查记录） |
| Word | 桌面版 16.0.20430.20146，64 位；本次环境核查时未运行 |
| Visual Studio / VSTO | Visual Studio Community 2026 18.5.11723.231；缺少 Office/SharePoint workload、VSTO MSBuild targets 和 VSTO runtime |
| 插件版本 | 未生成可验收的插件二进制；项目当前没有单独发布版本号 |
| 协议版本 | 1.1 |
| 引擎规则版本 | `demo-0.1.0`（仅合成演示规则） |
| .NET / CI | GitHub Actions 使用 .NET 10；本地用户级环境已配置 .NET SDK 10.0.100（同时保留 8.0.425），可按 `global.json` 构建并测试解决方案；Linux 环境仍不能替代 Word/VSTO 实机验收 |
| Word 插件实测 | 待完成：当前 Windows 环境缺少 VSTO 构建与运行组件，未启动本分支插件 |

### 可重复运行步骤

在 Windows 上安装仓库要求的 .NET 10 SDK、Visual Studio Office/SharePoint development workload、.NET Framework 4.8 targeting pack 和 VSTO runtime 后，按 [Word 插件运行说明](../src/ThesisAddIn.VSTO/README.md) 构建并注册插件。随后：

1. 打开 `known_issues_sample.docx`，分析并核对 3 条诊断及其原文位置。
2. 单独修复 `demo/font-size`，确认预览只含该属性；检查复查结果只消除了字号问题。
3. 从原始样例执行批量修复，确认预览只含两个 `auto` 项，西文字体明确列为保留项。
4. 确认原文件 SHA-256 不变；在 Word 中打开另存输出，记录是否出现修复提示，并重新分析输出。
5. 将 Windows、Word、Visual Studio/VSTO、插件程序集版本、协议版本、规则版本、命令、输出路径和结果补入本节。未通过的项目注明原因，不用静态图片代替 Word 实测。

当前已记录的 Windows/Word 版本仅说明目标机器具备桌面 Word；由于 VSTO 组件缺失，尚未完成插件构建、加载、单项/批量 UI 操作或输出 DOCX 的 Word 打开验证。不得将这些项目标记为通过，也不得把剩余诊断解释成整篇论文合规。

本次本地回归（2026-10-17）：`dotnet build NkuThesisTool.sln --configuration Release` 构建成功，`dotnet test NkuThesisTool.sln --configuration Release` 通过 68/68；结果仅覆盖引擎、协议宿主和自动化测试，不改变上述 Word/VSTO 实机验收状态。
