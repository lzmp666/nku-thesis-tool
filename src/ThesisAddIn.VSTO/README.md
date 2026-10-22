# Word 插件（VSTO）

在桌面 Word 中以原文黄色标记和单一问题检索侧栏展示格式问题；悬停原文标记时，侧栏选中对应问题并显示建议。
分析、定位、单项修复、批量修复和重新验证操作放在 Word 顶部功能区，详细问题只显示在检索侧栏中。

## 为什么是经典工程

本工程是**经典（非 SDK 风格）** csproj，只能用 Visual Studio 自带的 .NET Framework
版 MSBuild 构建：VSTO 的构建任务依赖该版本。仓库根目录的 `global.json` 要求 .NET
SDK 10，SDK 风格工程需要 MSBuild 18 及以上，两者互斥；经典工程不参与 SDK 解析，因此
可以在 VS MSBuild 下正常构建。契约程序集仍然是 netstandard2.0，由工程内的
`BuildProtocolContracts` 目标调用 dotnet CLI 构建。

## 环境要求

- Windows + 桌面版 Word
- Visual Studio 2022 并勾选 **Office/SharePoint 开发** 工作负载
- .NET Framework 4.8 targeting pack、VSTO runtime（10.0）
- .NET 10 SDK（用于构建协议宿主与契约程序集）

## 构建

```powershell
# 1. 引擎宿主（.NET 10，任意平台可用）
dotnet build src\ThesisAddIn.VSTO\ProtocolHost\ThesisEngine.ProtocolHost.csproj -c Release

# 2. 找到 VS 2022 的 MSBuild（支持自定义安装目录，例如 D:\vs2022）
$vswhere = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
$vsPath = & $vswhere -latest -products * -version "[17.0,18.0)" -property installationPath
if (-not $vsPath) { throw "未找到 Visual Studio 2022。" }
$msbuild = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) { throw "未找到 VS 2022 MSBuild：$msbuild" }

# VSTO 要求签名 ClickOnce 清单；此脚本生成只保存在当前用户证书库中的开发证书。
.\scripts\New-VstoDevCertificate.ps1
& $msbuild src\ThesisAddIn.VSTO\ThesisAddIn.VSTO.csproj /p:Configuration=Debug `
  /p:SignManifests=true "/p:ManifestCertificateThumbprint=$env:THESIS_VSTO_CERT_THUMBPRINT"
```

如果 NuGet 源暂时不可用，但协议项目已还原过依赖，可先用 `--no-restore` 构建契约程序集，
再让 VSTO MSBuild 复用该输出，避免重复访问 NuGet：

```powershell
dotnet build src\ThesisAddIn.VSTO\Protocol\ThesisAddIn.Protocol.csproj -c Release --no-restore
& $msbuild src\ThesisAddIn.VSTO\ThesisAddIn.VSTO.csproj /p:Configuration=Release `
  /p:SkipProtocolContractBuild=true
```

公开仓库不保存签名私钥。上述脚本只导出公开证书，并将它加入当前用户的 Root 与
TrustedPublisher 存储，以便本机演示；私钥保留在当前用户的个人证书库且不可导出。不要将
私钥或 PFX 加入版本控制，也不要把本机开发证书用于发布版本。若之前已下载过旧测试 PFX，
请将它视为已泄露，停止使用并删除对它的信任。

## 安装到 Word（当前用户）

```powershell
# 先按上方步骤创建并信任本机开发证书，再以 SignManifests=true 构建插件。
# 注册加载项
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\ThesisAddIn.VSTO"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name FriendlyName -Value "论文格式检查"
Set-ItemProperty -Path $key -Name Description   -Value "南开研究生学位论文格式检查插件"
Set-ItemProperty -Path $key -Name LoadBehavior  -Value 3 -Type DWord
$configuration = "Release"
$addinDirectory = (Resolve-Path "src\ThesisAddIn.VSTO\bin\$configuration").Path
Set-ItemProperty -Path $key -Name Manifest `
  -Value ("file:///" + ($addinDirectory -replace '\\','/') + "/ThesisAddIn.VSTO.vsto|vstolocal")

# 告诉插件引擎宿主在哪
[Environment]::SetEnvironmentVariable("THESIS_ENGINE_PROTOCOL_HOST",
  (Resolve-Path "src\ThesisAddIn.VSTO\ProtocolHost\bin\Release\net10.0\ThesisEngine.ProtocolHost.dll").Path, "User")
```

设置路径后需要**重新启动 Word**，以便插件重新建立引擎连接。卸载时删除上述注册
表项、证书与环境变量即可。

插件优先读取当前用户的 `THESIS_ENGINE_PROTOCOL_HOST` 设置，防止更新后启动器仍把旧进程环境传给 Word；未配置用户级路径时才使用进程环境。已配置路径无效时会报错，不会静默换用旧引擎。

## 使用

1. 在 Windows 桌面版 Microsoft Word 中打开 `.docx`。分析、修复、问题/位置导航都在顶部「论文格式检查」功能区；
   右侧「问题检索」面板只显示搜索框和合并后的问题列表。关闭侧栏后可从功能区的「问题检索」按钮重新打开。在启动失败时先看
   `%LOCALAPPDATA%\NkuThesisTool\addin.log`。
2. 点「分析当前文档」。插件读取当前 Word 窗口中的内存内容（包含未保存编辑），直接更新本窗口的
   黄色原文标记和检索列表；首次和重复分析都不另开文档、不更名，也不自动保存。未能可靠取得当前内容时会停止，不会改用磁盘上的旧版本。
   原文标记使用可恢复的书签保存原有高亮状态，不创建重复的 Word 回复式批注；已有人工批注会保留。
3. 鼠标停在黄色标记上会选中问题检索中的对应卡片并显示建议；也可从功能区用「上一项」「下一项」切换问题，
   用「上一处」「下一处」逐处定位同一问题的所有位置，或直接点击检索结果定位。这些操作不修改正文。
4. 「修复当前问题」处理当前选中项；「预览并修复全部」处理修复引擎支持的问题。
   两者都会先弹出中文预览，作者确认后才修改当前文档。无法解析的字体、未知格式及不支持的范围会保留并说明原因。
5. 确认后，格式修复和原文标记刷新直接作用于当前打开的文档，不会生成新结果文档、关闭或重开 Word 文档；一次确认对应一个 Word 撤销操作。改动在当前会话中尚未保存，可用 Ctrl+Z 撤销；作者完成所有处理后再按 Ctrl+S 保存当前文档，或使用 Word 的“另存为”。手动修改后可点「验证已解决」重新分析；只有原诊断消失才会清除对应标记。

开发回归可在编译后运行 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\Test-WordInPlaceRepair.ps1 -Configuration Debug`。脚本用完整合成论文在 Word 中检查表格前后的位置映射、分析标记、文档不另开不自动保存；再检查分组单项/批量修复、整体撤销/恢复和零问题文档的持久识别。视觉交互验收仍见 `docs/demo/word-acceptance.md`。

分析、修复计划、写入和复查通过异步进程调用运行，Word 界面可继续响应；Word 正常关闭时，
插件会取消正在等待的引擎请求并尝试终止宿主进程。文档副本、Word 定位和界面更新仍在 Word UI 线程执行。

## 已知限制

- Word 不提供正文悬停事件；插件以 250 毫秒间隔查询鼠标下的正文范围，把黄色标记与
  当前问题卡片联动，不改变 Word 当前选区。问题检索只显示一份问题列表，详细建议不会再重复创建为 Word 批注。
- 引擎和 Word 定位都按文档顺序计入正文及表格单元格里的段落；表格单元格仍不套用正文演示规则。图表题注、图片质量与版式、公式、页眉页脚等是未覆盖范围，界面会列出而不判为合规。
- 未知值和人工处理的问题不会自动修改；可修复项始终先显示预览并等待作者确认。
- 当前规则范围仅覆盖演示规则（正文东亚字体、西文字体、字号），不代表论文合规结论。
- 加载项未打包发布（无 ClickOnce 安装包），当前按上面的步骤在本机注册。



