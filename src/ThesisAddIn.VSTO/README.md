# Word 插件（VSTO）

在桌面 Word 中以黄色原文标记、原生连接批注和左侧问题检索展示格式问题；悬停标记联动对应建议。
分析、定位、单项修复、批量选择、撤销和图表排版操作放在顶部「论文格式检查」功能区。

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
   左侧「问题检索」面板显示搜索框和合并后的问题列表。关闭侧栏后可从功能区的检索按钮重新打开。在启动失败时先看
   `%LOCALAPPDATA%\NkuThesisTool\addin.log`。
2. 点「分析当前文档」。插件读取当前 Word 窗口中的内存内容（包含未保存编辑），直接更新本窗口的
   黄色原文标记和检索列表；首次和重复分析都不另开文档、不更名，也不自动保存。未能可靠取得当前内容时会停止，不会改用磁盘上的旧版本。
   原文标记使用可恢复的书签保存原有高亮状态，并创建对应的原生 Word 连接批注；已有人工批注会保留。
3. 鼠标停在黄色标记上会选中问题检索中的对应卡片并显示建议；也可从功能区用「上一项」「下一项」切换问题，
   用「上一处」「下一处」逐处定位同一问题的所有位置，或直接点击检索结果定位。这些操作不修改正文。
4. 卡片「解决」或「修复当前问题」直接修复选中项；「批量修复」进入列表勾选模式，再点「修复所选」。不再弹修复预览。未知值及不支持的范围会保留并说明原因。
5. 修复直接作用于当前打开的文档，不生成结果副本、不重开文档、不自动保存。卡片「撤销」恢复该项修复，目标已被用户改动时停止撤销；作者完成后自主保存或另存。修复后重新分析，只有诊断消失才清除对应标记。
6. 「图表排版」可设置已与题注唯一配对的行内图片及简单表格：图片居中、超宽等比例缩小；表格居中、字体字号、单元格对齐、三线表和跨页重复表头。表内字号和线宽是作者选择的预设，不冒充校级统一要求。每个对象独立形成 Word 撤销操作，可用 Ctrl+Z 撤销。复杂结构会列出跳过原因，详见 `docs/demo/word-object-layout-2026-10-24.md`。

开发回归可在编译后运行 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\Test-WordInPlaceRepair.ps1 -Configuration Debug`。脚本用完整合成论文在 Word 中检查表格前后的位置映射、分析标记、文档不另开不自动保存；再检查分组单项/批量修复、整体撤销/恢复和零问题文档的持久识别。视觉交互验收仍见 `docs/demo/word-acceptance.md`。

分析、修复计划和复查通过异步进程调用运行。Word COM 写入仍在 UI 线程，在规则组、分段写入和批注清理之间让出消息循环；修复期间检索仍可筛选和滚动，重复修改操作被禁用。相同复查结果复用剩余批注和范围，避免全文重建。完整快照校验和修复后核验保留；复杂对象的单次 COM 调用仍可能短暂阻塞。Word 关闭时会取消引擎请求并尝试终止宿主。

## 已知限制

- Word 不提供正文悬停事件；插件以 250 毫秒间隔查询鼠标下的正文范围，把黄色标记与
  当前问题卡片及原生批注联动，不改变 Word 当前选区。Word 自带「答复」等原生控件无法由加载项任意移除；修复应使用插件卡片的「解决」。
- 表格单元格不套用正文演示规则；表内设置通过显式「图表排版」选择应用。浮动、合并、嵌套等复杂对象，以及流程图或图片内部文字、图片清晰度、公式、页眉页脚不自动修复。
- 未知值和人工处理的问题不会自动修改；题注位置无法唯一确认时不会搬动对象。
- 当前分析只覆盖已实现规则，不代表全论文合规；诊断为 0 不代表所有图表都已合规。
- 开发包通过签名 ClickOnce 清单安装，仍标记 PREVIEW，尚未作为正式 Release 发布。



