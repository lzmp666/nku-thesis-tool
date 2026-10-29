# Windows Word 演示包

本包用于在 Windows 桌面版 Word 中演示 NKU Thesis Tool。只带合成样例，不带真实论文、证书私钥、PFX、账号凭据或构建机目录。加载项注册与引擎路径仅写入当前 Windows 用户，不修改 Office Trust Center，也不卸载或更改其他加载项。

## 运行前

1. 解压整个 ZIP 到一个普通文件夹。
2. 用 `SHA256SUMS.txt` 核验包内文件；`build-info.json` 记录源码提交和签名证书指纹。
3. 运行 `Preflight.ps1`。它检查桌面 Word、VSTO Runtime、.NET Framework 4.8+ 和 .NET 10 Runtime；缺项时先安装对应官方运行时，再重新预检。
4. 该演示证书是开发用自签名证书。只在确认来源和 `build-info.json` 中指纹后，才运行下面带 `-TrustPublisher` 的命令。它只将包内公开证书导入当前用户的 `Root` 与 `TrustedPublisher` 存储；私钥不在包里。若不信任证书，Word 不会加载插件。

## 安装

默认安装在 `D:\NkuThesisTool`，可传入自己的专用目录：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -TrustPublisher
# 自定义目录示例
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -InstallRoot 'D:\Apps\NkuThesisTool' -TrustPublisher
```

脚本会先预检，再将文件放入目标目录，并写入 `HKCU\Software\Microsoft\Office\Word\Addins\ThesisAddIn.VSTO` 与当前用户的 `THESIS_ENGINE_PROTOCOL_HOST` 环境变量。它保留安装前同名注册值/环境变量供卸载恢复。已存在但不含本工具安装记录的目录不会被覆盖。

安装后关闭并重新启动 Word。在 Word 功能区打开“论文格式检查”，先从 `samples/full-thesis` 打开 `synthetic_complete_thesis.docx` 做完整论文验收；该人工合成稿约 14 页，含中英文摘要、六章正文、图、表和公式，并预设三处文字格式问题。`samples` 根目录的 `known_issues_sample.docx` 与 `clean_sample.docx` 可用于快速验证。所有样例都不构成规范合规结论。完整步骤见包内及安装目录中的 `docs/word-acceptance.md`。

## 更新与卸载

更新时用新 ZIP 运行相同安装命令和相同 `-InstallRoot`；只替换包内同名组件，不清空目录中的其他文件，并保留首次安装前的设置快照。卸载：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall.ps1
# 自定义安装目录须传入相同的路径
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall.ps1 -InstallRoot 'D:\Apps\NkuThesisTool'
```

卸载只在当前注册仍指向本工具时恢复原值，只在引擎环境变量仍是本工具路径时恢复原值，并移除本安装明确加入信任存储的证书。清理安装目录时逐文件校验包清单；被修改的文件及用户额外放入的文件会保留。不会改动其他 Office 加载项或全局安全设置。完成后重新启动 Word。

安装、更新、卸载前必须关闭所有 Word 窗口。更新包的发行证书若发生变化，脚本会停止；须先卸载旧版、核验新发行者后重新安装，以保留正确的回滚与证书记录。

## 从源码制作包

维护者应在 PowerShell 7 中从干净、已测试的 Release 提交开始；包脚本拒绝未提交工作树。需要 Windows、桌面 Word/VSTO 开发工具链、.NET Framework targeting pack 和 .NET 10 SDK。包内安装脚本会写入 UTF-8 BOM，兼容 Windows PowerShell 5.1 的中文提示。

```powershell
dotnet build .\src\ThesisAddIn.VSTO\ProtocolHost\ThesisEngine.ProtocolHost.csproj -c Release
$env:THESIS_VSTO_CERT_THUMBPRINT = '<当前用户开发证书指纹>'
& 'D:\vs2022\MSBuild\Current\Bin\MSBuild.exe' `
  .\src\ThesisAddIn.VSTO\ThesisAddIn.VSTO.csproj /t:Build `
  /p:Configuration=Release /p:Platform=AnyCPU /p:VisualStudioVersion=17.0 `
  /p:VSToolsPath='D:\vs2022\MSBuild\Microsoft\VisualStudio\v17.0' `
  /p:SkipProtocolContractBuild=true /p:SignManifests=true `
  "/p:ManifestCertificateThumbprint=$env:THESIS_VSTO_CERT_THUMBPRINT"
.\scripts\New-WindowsWordDemoPackage.ps1
# 可选：需要结构 CLI 时，用下面命令替换上一条
# .\scripts\New-WindowsWordDemoPackage.ps1 -IncludeStructureCli
```

若在构建后手工重签，先签 `.dll.manifest`，再使用 `mage -Update <file.vsto> -AppManifest <file.dll.manifest> -Install false -CertHash <thumbprint>` 更新并签署部署清单。Mage 的 `-Install` 默认值是 true；VSTO 清单必须保留 false。打包和安装脚本会核对该属性及最终程序集清单的大小、哈希，阻止签名有效但依赖哈希过期的包。

Word 1.5.4 可直接通过“论文格式检查 → 大纲确认与排版”核对角色、明确勾选并在当前文档排版；“撤销整篇排版”恢复本轮格式，保存仍由作者决定。含结构 CLI 的包另安装 `cli` 目录，可运行 `dotnet .\cli\ThesisEngine.Cli.dll structure <file.docx> --report <新报告.json>`。完整操作和 CLI 参数见包内 `docs/unstyled-thesis-workflow.md`。默认不调用模型。

生成的 ZIP 位于 `artifacts`。此开发打包流程始终标记为 PREVIEW；验收表被填写不等于已批准发布。包内 `vsto`、`engine`、`samples` 分开存放，省略 PDB，仅包含部署所需 Release 输出、合成稿、公开 `.cer` 和 SHA-256 清单。若源码包含 `docs/demo/word-acceptance-2026-10-22.md`，包会一并带上候选实测记录；它不替代最终同包和独立环境验收。不会将私钥/PFX 导出或加入 Git。正式发布需另行核对受测提交、二进制校验值、独立环境验收和负责人批准，并使用正式代码签名发布流程。
