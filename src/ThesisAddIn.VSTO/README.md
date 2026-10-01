# VSTO 环境验证与后续入口
Phase 0 不做重 UI；当前未创建/验证 VSTO 项目，也未将其加入 Core/CLI 解决方案。
需要 Windows、桌面版 Word、支持 Office/SharePoint 开发工具的 Visual Studio、.NET Framework 4.8 targeting pack 及 VSTO runtime。
验证步骤：
1. Visual Studio 安装器确认 Office/SharePoint development 工作负载（版本名称可能不同），Word 可正常创建/另存 .docx。
2. 新建 C# Word VSTO Add-in；目标 .NET Framework 4.8，确认 ThisAddIn_Startup 能载入，ThisAddIn_Shutdown 无错误。
3. 调试启动 Word，检查 COM Add-ins 可见并已启用；分别记录 Office/VS/系统版本与位数、信任配置。
4. 保存最小项目和环境验证记录到本目录，通过单独 Issue/PR 审查。不要禁用全局 Office 安全设置。
5. 后续通过版本化 CLI/进程协议接 Core（net10.0）；VSTO 不直接引用此 Core。先验证诊断/补丁闭环，再做定位和确认 UI。
验收：启动/退出 Word 不崩溃、插件可启停、调用故障有清晰提示；安装部署另建任务。
