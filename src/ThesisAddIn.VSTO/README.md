# Word 插件（VSTO）

在桌面 Word 中提供中文问题侧边栏：列出分析发现的问题、点击定位到原文、预览并修复
可自动处理的项目，并把结果另存为新的 .docx。

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

设置环境变量后需要**重新启动 Word**（环境变量在进程启动时读取）。卸载时删除上述注册
表项、证书与环境变量即可。

## 使用

1. Word 右侧出现「论文格式检查」面板；在启动失败时先看
   `%LOCALAPPDATA%\NkuThesisTool\addin.log`。
2. 打开一份 .docx，点「分析当前文档」。问题卡片显示规则、位置、期望值、实际值、严重度与
   处理方式；面板上方显示规则版本与已检查范围，底部显示未覆盖范围。
3. 点击问题卡片，Word 会选中对应段落（选中即为可见定位，不修改文档内容）。
4. 「修复此项」只处理当前选中项；「预览并修复全部」只处理处理分类为 auto 的项目。
   两者都会先弹出中文预览、再由你选择输出文件，结果另存为新 .docx，原文件不变。
5. 修复完成后插件会自动重新分析输出文件并刷新列表。

分析、修复计划、写入和复查通过异步进程调用运行，Word 界面可继续响应；Word 正常关闭时，
插件会取消正在等待的引擎请求并尝试终止宿主进程。文档副本、Word 定位和界面更新仍在 Word UI 线程执行。

## 已知限制

- Word 不提供正文悬停事件；插件以 250 毫秒间隔查询鼠标下的正文范围，将批注锚点映射到
  建议卡片，不改变 Word 当前选区。点击/选中锚点也会触发同一联动。悬停映射依赖 Word 能返回范围且批注内容含有
  唯一规则编号、期望值和实际值；不满足时保留当前选择。
- 定位按段落序号映射：引擎只统计正文段落，Word 的 Paragraphs 还包含表格单元格内的
  段落，含表格的文档可能错位；此时插件给出提示而不是选中错误内容。
- 处理分类为 suggest / manual 的问题不会被自动修复，插件会单独列出。
- 当前规则范围仅覆盖演示规则（正文东亚字体、西文字体、字号），不代表论文合规结论。
- 加载项未打包发布（无 ClickOnce 安装包），当前按上面的步骤在本机注册。


