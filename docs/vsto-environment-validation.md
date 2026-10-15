# VSTO 环境验证记录

## 当前记录

| 项目 | 记录 |
| --- | --- |
| 验证状态 | **阻塞：当前开发环境不是 Windows，未安装桌面版 Word、Visual Studio Office/SharePoint 工作负载或 VSTO runtime** |
| 操作系统 | 未实测 |
| Windows 位数 | 未实测 |
| Word 版本/位数 | 未实测 |
| Visual Studio 版本 | 未实测 |
| .NET Framework 4.8 targeting pack | 未实测 |
| VSTO runtime | 未实测 |
| Office COM 加载项启用/禁用 | 未实测 |
| Word 启动/退出 | 未实测 |

当前仓库没有目标 Windows + 桌面版 Word，不能据此宣称插件已通过 Word 实机验证；
本记录不以截图替代测试，也没有修改 Office 全局安全配置。

## 目标 Windows 机验证步骤

1. 在 Visual Studio Installer 安装 Office/SharePoint development 工作负载、.NET
   Framework 4.8 targeting pack，并确认 VSTO runtime 可用。
2. 使用 src/ThesisAddIn.VSTO/ThesisAddIn.VSTO.csproj 构建 Debug 版本，确认生成的
   VSTO 清单和签名配置符合团队部署要求。
3. 设置 THESIS_ENGINE_PROTOCOL_HOST 指向已构建的
   ThesisEngine.ProtocolHost.dll，按协议说明启动调试。
4. 启动桌面版 Word，确认加载项出现在 COM Add-ins 中且可启用；禁用后再次启动，
   确认 Word 仍能正常退出。
5. 打开合成 DOCX，执行一次分析，核对侧端收到 ruleVersion、诊断位置和
   unsupportedScopes；宿主不可用、超时、退出和协议版本错误都应显示中文提示。
6. 记录上表中的实际版本、位数、启用状态和测试结果，附在后续验证 PR；不要关闭或
   降低 Office 全局安全设置来绕过加载问题。

## 可重复的最小检查

宿主进程本身可在非 Windows 环境构建和测试，因为它不加载 Word COM。VSTO 的 COM
加载、Word 启停和插件启用/禁用必须在目标 Windows 机完成；没有该环境时保持本记录
为阻塞状态，继续推进不会把协议测试伪装成 Word 实测。
