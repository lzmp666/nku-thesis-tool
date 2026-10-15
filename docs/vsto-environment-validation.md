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
