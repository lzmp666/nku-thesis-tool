[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OutputDirectory = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts'),
    [string]$CertificateThumbprint = $env:THESIS_VSTO_CERT_THUMBPRINT,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$addinOutput = Join-Path $root 'src\ThesisAddIn.VSTO\bin\Release'
$engineOutput = Join-Path $root 'src\ThesisAddIn.VSTO\ProtocolHost\bin\Release\net10.0'
$sampleRoot = Join-Path $root 'tests\fixtures\synthetic\demo'
$requiredAddin = @('ThesisAddIn.VSTO.dll', 'ThesisAddIn.VSTO.vsto', 'ThesisAddIn.VSTO.dll.manifest')
foreach ($name in $requiredAddin) {
    if (-not (Test-Path -LiteralPath (Join-Path $addinOutput $name) -PathType Leaf)) { throw "缺少 Release 加载项输出：$name。请先完成签名构建。" }
}
foreach ($name in @('ThesisEngine.ProtocolHost.dll', 'ThesisEngine.ProtocolHost.exe', 'ThesisEngine.ProtocolHost.runtimeconfig.json', 'ThesisEngine.ProtocolHost.deps.json')) {
    if (-not (Test-Path -LiteralPath (Join-Path $engineOutput $name) -PathType Leaf)) { throw "缺少 Release 引擎输出：$name。请先构建 ProtocolHost。" }
}
foreach ($name in @('known_issues_sample.docx', 'clean_sample.docx', 'manifest.json', 'README.md')) {
    if (-not (Test-Path -LiteralPath (Join-Path $sampleRoot $name) -PathType Leaf)) { throw "缺少合成样例文件：$name" }
}
if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { throw '请设置 THESIS_VSTO_CERT_THUMBPRINT 或传入 -CertificateThumbprint。' }

$certificate = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq ($CertificateThumbprint -replace '\s', '') } | Select-Object -First 1
if (-not $certificate) { throw "当前用户个人证书库中找不到签名证书 $CertificateThumbprint。不会生成未签名演示包。" }
if (-not $certificate.HasPrivateKey) { throw '当前用户证书没有对应私钥，不能作为签名构建证书。' }
foreach ($manifestName in @('ThesisAddIn.VSTO.vsto', 'ThesisAddIn.VSTO.dll.manifest')) {
    $manifestDocument = [xml](Get-Content -LiteralPath (Join-Path $addinOutput $manifestName) -Raw)
    $embeddedNode = $manifestDocument.SelectSingleNode("//*[local-name()='X509Certificate']")
    if ($null -eq $embeddedNode) { throw "清单 $manifestName 未嵌入签名证书。" }
    $embeddedCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($embeddedNode.InnerText.Trim()))
    if ($embeddedCertificate.Thumbprint -ne $certificate.Thumbprint) { throw "清单 $manifestName 的签名证书与 -CertificateThumbprint 不匹配。" }
}

$revision = (& git -C $root rev-parse HEAD 2>$null | Out-String).Trim()
if ($LASTEXITCODE -ne 0 -or -not $revision) { throw '无法从 Git 读取源码提交号。请在仓库工作树中运行。' }
$dirty = (& git -C $root status --porcelain 2>$null | Out-String).Trim()
if ($dirty) { throw '工作树有未提交修改。请从干净提交构建演示包，避免包与源码 SHA 不一致。' }

$outDir = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$acceptancePath = Join-Path $root 'docs\demo\word-acceptance.md'
$acceptanceIncluded = Test-Path -LiteralPath $acceptancePath -PathType Leaf
# A checklist with no pending cells can still contain failed or unrelated results.
# Distribution approval requires a separately reviewed release record; this developer
# packaging path therefore always produces an explicitly labelled preview.
$acceptanceComplete = $false
$readyForDistribution = $false
$suffix = if ($readyForDistribution) { '' } else { '-PREVIEW' }
$zipName = "NkuThesisTool-WordDemo-$($revision.Substring(0, 12))$suffix.zip"
$zipPath = Join-Path $outDir $zipName
if ((Test-Path -LiteralPath $zipPath) -and -not $Force) { throw "输出包已存在：$zipPath。使用 -Force 才会替换这个精确文件。" }

$stage = Join-Path ([System.IO.Path]::GetTempPath()) ("NkuThesisTool-Demo-" + [guid]::NewGuid().ToString('N'))
try {
    $vstoStage = Join-Path $stage 'vsto'
    $engineStage = Join-Path $stage 'engine'
    $samplesStage = Join-Path $stage 'samples'
    $docsStage = Join-Path $stage 'docs'
    New-Item -ItemType Directory -Path $vstoStage, $engineStage, $samplesStage, $docsStage | Out-Null

    Get-ChildItem -LiteralPath $addinOutput -File | Where-Object { $_.Extension -in @('.dll', '.vsto', '.manifest', '.config') } |
        Copy-Item -Destination $vstoStage
    Get-ChildItem -LiteralPath $engineOutput -File | Where-Object { $_.Extension -in @('.dll', '.exe', '.json') } |
        Copy-Item -Destination $engineStage
    foreach ($name in @('known_issues_sample.docx', 'clean_sample.docx', 'manifest.json', 'README.md')) {
        Copy-Item -LiteralPath (Join-Path $sampleRoot $name) -Destination $samplesStage
    }
    Copy-Item -LiteralPath (Join-Path $root 'scripts\Test-WindowsWordDemoEnvironment.ps1') -Destination (Join-Path $stage 'Preflight.ps1')
    Copy-Item -LiteralPath (Join-Path $root 'scripts\Install-WindowsWordDemo.ps1') -Destination (Join-Path $stage 'Install.ps1')
    Copy-Item -LiteralPath (Join-Path $root 'scripts\Uninstall-WindowsWordDemo.ps1') -Destination (Join-Path $stage 'Uninstall.ps1')
    # Windows PowerShell 5.1 requires a BOM to decode Chinese messages as UTF-8.
    foreach ($scriptName in @('Preflight.ps1', 'Install.ps1', 'Uninstall.ps1')) {
        $scriptPath = Join-Path $stage $scriptName
        $scriptText = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($scriptPath, $scriptText, [System.Text.UTF8Encoding]::new($true))
    }
    Copy-Item -LiteralPath (Join-Path $root 'docs\demo\windows-package.md') -Destination (Join-Path $stage 'README.md')
    if ($acceptanceIncluded) {
        Copy-Item -LiteralPath $acceptancePath -Destination (Join-Path $docsStage 'word-acceptance.md')
    } else {
        Copy-Item -LiteralPath (Join-Path $root 'docs\demo\README.md') -Destination (Join-Path $docsStage 'demo-readme.md')
        @'
# Word 插件验收记录

这是 PREVIEW 包。Issue #52 的 Word 桌面实测尚未并入此源码提交，因此本文件不包含通过结论。
请完成 #52 合成文档验收并合并其记录后，重新从该提交生成正式演示包。
'@ | Set-Content -LiteralPath (Join-Path $docsStage 'word-acceptance.md') -Encoding utf8
    }
    [System.IO.File]::WriteAllBytes((Join-Path $stage 'publisher.cer'), $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))

    $buildInfo = [ordered]@{
        product = 'NKU Thesis Tool Word Demo'
        sourceRevision = $revision
        builtAtUtc = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
        signingCertificateThumbprint = $certificate.Thumbprint
        wordAcceptanceChecklistIncluded = $acceptanceIncluded
        wordAcceptanceComplete = $acceptanceComplete
        readyForDistribution = $readyForDistribution
        note = 'Synthetic demonstration only; no private key or real thesis is included.'
    }
    $buildInfo | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $stage 'build-info.json') -Encoding utf8

    $checksumLines = Get-ChildItem -LiteralPath $stage -Recurse -File | Sort-Object FullName | ForEach-Object {
        $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        $relative = $_.FullName.Substring($stage.Length + 1).Replace('\', '/')
        "$hash  $relative"
    }
    $checksumLines | Set-Content -LiteralPath (Join-Path $stage 'SHA256SUMS.txt') -Encoding utf8

    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zipPath -CompressionLevel Optimal
    Write-Output "演示包已创建：$zipPath"
    Write-Output "源码提交：$revision"
    Write-Output "签名证书公钥已放入包内；未导出私钥或 PFX。"
    if (-not $readyForDistribution) { Write-Warning 'Word 验收未全部记录为通过；这是 PREVIEW 包，不可作为已验收演示包分发。' }
}
finally {
    if (Test-Path -LiteralPath $stage) {
        $stageFull = [System.IO.Path]::GetFullPath($stage)
        $temporaryRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\')
        if (-not $stageFull.StartsWith($temporaryRoot + '\', [StringComparison]::OrdinalIgnoreCase) -or
            [System.IO.Path]::GetFileName($stageFull) -notmatch '^NkuThesisTool-Demo-[0-9a-f]{32}$') {
            throw '临时打包目录不在预期范围内，拒绝递归清理。'
        }
        $stageItem = Get-Item -LiteralPath $stageFull -Force
        if (($stageItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw '临时打包目录是重解析点，拒绝清理。' }
        Remove-Item -LiteralPath $stageFull -Recurse -Force
    }
}
