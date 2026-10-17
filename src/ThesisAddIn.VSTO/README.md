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

# 2. 插件本体（必须用 Visual Studio 自带的 MSBuild）
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  src\ThesisAddIn.VSTO\ThesisAddIn.VSTO.csproj /p:Configuration=Debug
```

构建产物包括 `bin\Debug\ThesisAddIn.VSTO.dll` 与签名后的 `ThesisAddIn.VSTO.vsto`。
工程目录下的 `ThesisAddIn.VSTO_TemporaryKey.pfx` 是 Visual Studio 生成 VSTO 项目时
同样会创建的临时签名证书，用于清单签名，不包含任何真实密钥。

## 安装到 Word（当前用户）

```powershell
# 信任工程自带的临时证书（仅当前用户）
$pfx = "src\ThesisAddIn.VSTO\ThesisAddIn.VSTO_TemporaryKey.pfx"
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($pfx, "")
foreach ($name in @("Root","TrustedPublisher")) {
  $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($name,"CurrentUser")
  $store.Open("ReadWrite"); $store.Add($cert); $store.Close()
}

# 注册加载项
$key = "HKCU:\Software\Microsoft\Office\Word\Addins\ThesisAddIn.VSTO"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name FriendlyName -Value "论文格式检查"
Set-ItemProperty -Path $key -Name Description   -Value "南开研究生学位论文格式检查插件"
Set-ItemProperty -Path $key -Name LoadBehavior  -Value 3 -Type DWord
Set-ItemProperty -Path $key -Name Manifest `
  -Value ("file:///" + ((Resolve-Path "src\ThesisAddIn.VSTO\bin\Debug").Path -replace '\\','/') + "/ThesisAddIn.VSTO.vsto|vstolocal")

# 告诉插件引擎宿主在哪
[Environment]::SetEnvironmentVariable("THESIS_ENGINE_PROTOCOL_HOST",
  (Resolve-Path "src\ThesisAddIn.VSTO\ProtocolHost\bin\Release\net10.0\ThesisEngine.ProtocolHost.dll").Path, "User")
```

设置环境变量后需要**重新启动 Word**（环境变量在进程启动时读取）。卸载时删除上述注册
表项、证书与环境变量即可。

## 使用

1. Word 右侧出现「论文格式检查」面板；在启动失败时先看
   `%LOCALAPPDATA%\NkuThesisTool\addin.log`。
2. 打开一份 .docx，点「分析当前文档」。列表显示规则、位置、期望值、实际值、严重度与
   处理方式；面板上方显示规则版本与已检查范围，底部显示未覆盖范围。
3. 点击列表中的问题，Word 会选中对应段落（选中即为可见定位，不修改文档内容）。
4. 「修复此项」只处理当前选中项；「预览并修复全部」只处理处理分类为 auto 的项目。
   两者都会先弹出中文预览、再由你选择输出文件，结果另存为新 .docx，原文件不变。
5. 修复完成后插件会重新分析新输出文件并刷新列表；已修复且当前规则可检查的问题会从结果中消失，未自动修复的问题仍会显示。当前打开的原始文档不会被改动；如需继续处理剩余问题，请先打开最新输出文件。

## #52 修复闭环验收步骤

验收只使用合成样例 `tests\fixtures\synthetic\demo\known_issues_sample.docx`，不使用真实论文或个人信息。自动化测试覆盖协议宿主的单项/批量修复、另存文件、输入不变和输出复查；Word 界面步骤仍需在具备 VSTO 的 Windows 桌面环境中执行。

1. 在仓库根目录执行 `dotnet build src\ThesisAddIn.VSTO\ProtocolHost\ThesisEngine.ProtocolHost.csproj -c Release` 和 `dotnet test NkuThesisTool.sln -c Release`。
2. 在桌面 Word 中打开上述合成样例，点「分析当前文档」。应看到东亚字体、西文字体和字号三条诊断，并能定位到对应段落。
3. 选中 `demo/font-size`，点「修复此项」。预览应只列出字号；确认并另存为新文件后，侧边栏重新分析输出，字号问题消失，东亚字体和西文字体问题仍列出。
4. 重新打开原始合成样例并点「预览并修复全部」。预览应只列出标记为 `auto` 的东亚字体和字号，西文字体列在保留项中。确认并另存后，输出复查应仅保留西文字体问题。
5. 比较原始样例的 SHA-256，确认未变化；在 Word 中打开输出文件，确认没有文档修复提示。记录 Windows、Word、Visual Studio/VSTO、插件构建版本、协议版本和引擎规则版本。

如果复查仍有诊断，结果只表示这些问题仍待处理；演示规则覆盖有限，不能据此宣称整篇论文合规。

## 已知限制

- 插件没有悬停事件可用：Word 对象模型不提供鼠标悬停通知，因此“从正文批注激活建议
  卡片”实现为**点击/选中锚点**触发（见 ThisAddIn.OnWindowSelectionChange），不宣称
  支持纯悬停。
- 定位按段落序号映射：引擎只统计正文段落，Word 的 Paragraphs 还包含表格单元格内的
  段落，含表格的文档可能错位；此时插件给出提示而不是选中错误内容。
- 处理分类为 suggest / manual 的问题不会被自动修复，插件会单独列出。
- 当前规则范围仅覆盖演示规则（正文东亚字体、西文字体、字号），不代表论文合规结论。
- 加载项未打包发布（无 ClickOnce 安装包），当前按上面的步骤在本机注册。
