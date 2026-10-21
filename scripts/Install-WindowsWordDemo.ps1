[CmdletBinding()]
param(
    [string]$InstallRoot = 'D:\NkuThesisTool',
    [switch]$TrustPublisher
)

$ErrorActionPreference = 'Stop'
$packageRoot = $PSScriptRoot
$installPath = [System.IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
$driveRoot = [System.IO.Path]::GetPathRoot($installPath).TrimEnd('\')
if ([string]::IsNullOrWhiteSpace($installPath) -or $installPath -eq $driveRoot -or $installPath.Length -lt 8) {
    throw '安装目录必须是明确的子目录，不能是磁盘根目录。'
}
foreach ($relative in @('vsto\ThesisAddIn.VSTO.vsto', 'engine\ThesisEngine.ProtocolHost.dll', 'publisher.cer', 'Preflight.ps1')) {
    if (-not (Test-Path -LiteralPath (Join-Path $packageRoot $relative) -PathType Leaf)) { throw "演示包不完整，缺少 $relative。" }
}
$checksumPath = Join-Path $packageRoot 'SHA256SUMS.txt'
if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) { throw '缺少 SHA256SUMS.txt，拒绝安装不完整的演示包。' }
$packageRootFull = [System.IO.Path]::GetFullPath($packageRoot).TrimEnd('\')
$checksumCount = 0
foreach ($line in (Get-Content -LiteralPath $checksumPath)) {
    if ($line -notmatch '^([0-9a-f]{64})  (.+)$') { continue }
    $expectedHash = $Matches[1]
    $relative = $Matches[2].Replace('/', '\')
    if ($relative.Split('\') -contains '..' -or [System.IO.Path]::IsPathRooted($relative)) { throw '包校验清单含无效路径，拒绝安装。' }
    $filePath = [System.IO.Path]::GetFullPath((Join-Path $packageRootFull $relative))
    if (-not $filePath.StartsWith($packageRootFull + '\', [StringComparison]::OrdinalIgnoreCase)) { throw '包校验路径越出解压目录，拒绝安装。' }
    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { throw "包文件缺失：$relative" }
    $actualHash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualHash -ne $expectedHash) { throw "包文件校验失败：$relative" }
    $checksumCount++
}
if ($checksumCount -lt 10) { throw 'SHA256SUMS.txt 内容不完整，拒绝安装。' }

$buildInfoPath = Join-Path $packageRoot 'build-info.json'
if (-not (Test-Path -LiteralPath $buildInfoPath -PathType Leaf)) { throw '缺少 build-info.json。' }
$buildInfo = Get-Content -LiteralPath $buildInfoPath -Raw | ConvertFrom-Json
$packageCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Join-Path $packageRoot 'publisher.cer'))
if (-not [string]::Equals([string]$buildInfo.signingCertificateThumbprint, $packageCertificate.Thumbprint, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'build-info.json 的证书指纹与 publisher.cer 不一致。'
}
foreach ($manifestRelative in @('vsto\ThesisAddIn.VSTO.vsto', 'vsto\ThesisAddIn.VSTO.dll.manifest')) {
    $manifestPath = Join-Path $packageRoot $manifestRelative
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "缺少签名清单：$manifestRelative" }
    $manifestDocument = [xml](Get-Content -LiteralPath $manifestPath -Raw)
    $embeddedNode = $manifestDocument.SelectSingleNode("//*[local-name()='X509Certificate']")
    if ($null -eq $embeddedNode) { throw "清单未包含签名证书：$manifestRelative" }
    $embeddedCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($embeddedNode.InnerText.Trim()))
    if (-not [string]::Equals($embeddedCertificate.Thumbprint, $packageCertificate.Thumbprint, [StringComparison]::OrdinalIgnoreCase)) {
        throw "签名清单证书与包证书不一致：$manifestRelative"
    }
}

$preflightHost = Join-Path $PSHOME 'pwsh.exe'
if (-not (Test-Path -LiteralPath $preflightHost)) { $preflightHost = Join-Path $PSHOME 'powershell.exe' }
if (-not (Test-Path -LiteralPath $preflightHost)) { throw '无法找到 PowerShell 来运行预检。' }
& $preflightHost -NoProfile -File (Join-Path $packageRoot 'Preflight.ps1')
if ($LASTEXITCODE -ne 0) { throw '环境预检失败；没有安装加载项。' }

$statePath = Join-Path $installPath '.nku-thesis-tool-install-state.json'
if ((Test-Path -LiteralPath $installPath) -and -not (Test-Path -LiteralPath $statePath)) {
    throw "目标目录已存在且没有本工具的安装记录，为避免覆盖其他文件已停止：$installPath"
}
New-Item -ItemType Directory -Path $installPath -Force | Out-Null
$registryPath = 'HKCU:\Software\Microsoft\Office\Word\Addins\ThesisAddIn.VSTO'
$registryNames = @('FriendlyName', 'Description', 'LoadBehavior', 'Manifest')

if (Test-Path -LiteralPath $statePath) {
    $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
    $previousRegistryExisted = [bool]$state.PreviousRegistryExisted
    $previousRegistryValues = $state.PreviousRegistryValues
    $previousEngineHost = [string]$state.PreviousEngineHost
    $addedToRoot = [bool]$state.AddedPublisherToRoot
    $addedToPublisher = [bool]$state.AddedPublisherToTrustedPublisher
} else {
    $previousRegistryExisted = Test-Path -LiteralPath $registryPath
    $previousRegistryValues = [ordered]@{}
    if ($previousRegistryExisted) {
        $old = Get-ItemProperty -LiteralPath $registryPath
        foreach ($name in $registryNames) {
            $property = $old.PSObject.Properties[$name]
            if ($null -ne $property) { $previousRegistryValues[$name] = $property.Value }
        }
    }
    $previousEngineHost = [Environment]::GetEnvironmentVariable('THESIS_ENGINE_PROTOCOL_HOST', 'User')
    $addedToRoot = $false
    $addedToPublisher = $false
}

foreach ($directory in @('vsto', 'engine', 'samples', 'docs')) {
    $source = Join-Path $packageRoot $directory
    $target = Join-Path $installPath $directory
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Get-ChildItem -LiteralPath $source -Force | Copy-Item -Destination $target -Recurse -Force
}
foreach ($file in @('publisher.cer', 'build-info.json', 'SHA256SUMS.txt', 'README.md')) {
    $source = Join-Path $packageRoot $file
    if (Test-Path -LiteralPath $source -PathType Leaf) { Copy-Item -LiteralPath $source -Destination $installPath -Force }
}
New-Item -ItemType Directory -Path (Join-Path $installPath 'docs') -Force | Out-Null

$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Join-Path $installPath 'publisher.cer'))
function Add-PublisherCertificate([string]$StoreName, [System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert) {
    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($StoreName, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $exists = $false
        foreach ($item in $store.Certificates) { if ($item.Thumbprint -eq $Cert.Thumbprint) { $exists = $true; break } }
        if (-not $exists) { $store.Add($Cert) }
        return (-not $exists)
    } finally { $store.Close() }
}
$manifestPath = Join-Path $installPath 'vsto\ThesisAddIn.VSTO.vsto'
$enginePath = Join-Path $installPath 'engine\ThesisEngine.ProtocolHost.dll'
$manifestUri = ([System.Uri]$manifestPath).AbsoluteUri + '|vstolocal'
$state = [ordered]@{
    schemaVersion = 1
    installRoot = $installPath
    installedManifest = $manifestUri
    installedEngineHost = $enginePath
    previousRegistryExisted = $previousRegistryExisted
    previousRegistryValues = $previousRegistryValues
    previousEngineHost = $previousEngineHost
    addedPublisherToRoot = $addedToRoot
    addedPublisherToTrustedPublisher = $addedToPublisher
    publisherThumbprint = $certificate.Thumbprint
}
$state | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding utf8
if ($TrustPublisher) {
    $addedToRoot = (Add-PublisherCertificate 'Root' $certificate) -or $addedToRoot
    $state.addedPublisherToRoot = $addedToRoot
    $state | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding utf8
    $addedToPublisher = (Add-PublisherCertificate 'TrustedPublisher' $certificate) -or $addedToPublisher
    $state.addedPublisherToTrustedPublisher = $addedToPublisher
    $state | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding utf8
} else {
    Write-Warning '未修改证书信任存储。如该签名证书尚未受信任，Word 会拒绝加载或显示信任提示。仅在确认发行者后使用 -TrustPublisher。'
}

New-Item -Path $registryPath -Force | Out-Null
Set-ItemProperty -LiteralPath $registryPath -Name FriendlyName -Value '论文格式检查'
Set-ItemProperty -LiteralPath $registryPath -Name Description -Value '南开研究生学位论文格式检查演示插件'
New-ItemProperty -LiteralPath $registryPath -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
Set-ItemProperty -LiteralPath $registryPath -Name Manifest -Value $manifestUri
[Environment]::SetEnvironmentVariable('THESIS_ENGINE_PROTOCOL_HOST', $enginePath, 'User')

$state | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding utf8
Write-Output "已安装到 $installPath。注册与引擎路径仅写入当前 Windows 用户。请关闭并重新启动 Word。"
