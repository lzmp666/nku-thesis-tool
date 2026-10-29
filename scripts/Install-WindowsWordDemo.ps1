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
$listedFiles = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($line in (Get-Content -LiteralPath $checksumPath)) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($line -notmatch '^([0-9a-f]{64})  (.+)$') { throw '包校验清单格式无效，拒绝安装。' }
    $expectedHash = $Matches[1]
    $relative = $Matches[2].Replace('/', '\')
    if ($relative.Split('\') -contains '..' -or [System.IO.Path]::IsPathRooted($relative)) { throw '包校验清单含无效路径，拒绝安装。' }
    if (-not $listedFiles.Add($relative)) { throw "包校验清单含重复文件：$relative" }
    $filePath = [System.IO.Path]::GetFullPath((Join-Path $packageRootFull $relative))
    if (-not $filePath.StartsWith($packageRootFull + '\', [StringComparison]::OrdinalIgnoreCase)) { throw '包校验路径越出解压目录，拒绝安装。' }
    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { throw "包文件缺失：$relative" }
    $actualHash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualHash -ne $expectedHash) { throw "包文件校验失败：$relative" }
    $checksumCount++
}
if ($checksumCount -lt 10) { throw 'SHA256SUMS.txt 内容不完整，拒绝安装。' }
foreach ($item in (Get-ChildItem -LiteralPath $packageRootFull -Recurse -Force)) {
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw '包内含重解析点，拒绝安装。' }
    if ($item.PSIsContainer -or $item.FullName -eq $checksumPath) { continue }
    $relative = $item.FullName.Substring($packageRootFull.Length + 1)
    if (-not $listedFiles.Contains($relative)) { throw "包内文件未列入校验清单：$relative" }
}

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

# Mage -Update defaults to install=true, which is not the VSTO local deployment
# mode. Check the final signed files together, before changing user settings.
$deploymentDocument = [xml](Get-Content -LiteralPath (Join-Path $packageRoot 'vsto\ThesisAddIn.VSTO.vsto') -Raw)
$deploymentNode = $deploymentDocument.SelectSingleNode("//*[local-name()='deployment']")
if ($null -eq $deploymentNode -or $deploymentNode.GetAttribute('install') -ne 'false') {
    throw 'VSTO 部署清单必须使用 install="false"；请在 Mage 更新时显式指定 -Install false。'
}
$applicationReferences = $deploymentDocument.SelectNodes("//*[local-name()='dependentAssembly'][@codebase='ThesisAddIn.VSTO.dll.manifest']")
if ($applicationReferences.Count -ne 1) { throw 'VSTO 部署清单缺少唯一的程序集清单引用。' }
$applicationReference = $applicationReferences[0]
$applicationManifestPath = Join-Path $packageRoot 'vsto\ThesisAddIn.VSTO.dll.manifest'
if ([int64]$applicationReference.GetAttribute('size') -ne (Get-Item -LiteralPath $applicationManifestPath).Length) {
    throw 'VSTO 部署清单记录的程序集清单大小已过期；请在最终签名后重新更新部署清单。'
}
$digestMethod = $applicationReference.SelectSingleNode("*[local-name()='hash']/*[local-name()='DigestMethod']")
$digestValue = $applicationReference.SelectSingleNode("*[local-name()='hash']/*[local-name()='DigestValue']")
if ($null -eq $digestMethod -or $null -eq $digestValue) { throw 'VSTO 部署清单缺少程序集清单哈希。' }
$hashAlgorithm = switch ($digestMethod.GetAttribute('Algorithm')) {
    'http://www.w3.org/2000/09/xmldsig#sha1' { [Security.Cryptography.SHA1]::Create(); break }
    'http://www.w3.org/2001/04/xmlenc#sha256' { [Security.Cryptography.SHA256]::Create(); break }
    'http://www.w3.org/2000/09/xmldsig#sha256' { [Security.Cryptography.SHA256]::Create(); break }
    default { throw 'VSTO 部署清单使用了不支持的哈希算法。' }
}
try { $actualDigest = [Convert]::ToBase64String($hashAlgorithm.ComputeHash([IO.File]::ReadAllBytes($applicationManifestPath))) }
finally { $hashAlgorithm.Dispose() }
if ($actualDigest -ne $digestValue.InnerText.Trim()) {
    throw 'VSTO 部署清单记录的程序集清单哈希已过期；请先签程序集清单，再更新并签署部署清单。'
}

$preflightHost = Join-Path $PSHOME 'pwsh.exe'
if (-not (Test-Path -LiteralPath $preflightHost)) { $preflightHost = Join-Path $PSHOME 'powershell.exe' }
if (-not (Test-Path -LiteralPath $preflightHost)) { throw '无法找到 PowerShell 来运行预检。' }
& $preflightHost -NoProfile -File (Join-Path $packageRoot 'Preflight.ps1')
if ($LASTEXITCODE -ne 0) { throw '环境预检失败；没有安装加载项。' }
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { throw '请先关闭所有 Word 窗口，再安装或更新插件；当前配置未更改。' }

$statePath = Join-Path $installPath '.nku-thesis-tool-install-state.json'
if ((Test-Path -LiteralPath $installPath) -and -not (Test-Path -LiteralPath $statePath)) {
    throw "目标目录已存在且没有本工具的安装记录，为避免覆盖其他文件已停止：$installPath"
}
New-Item -ItemType Directory -Path $installPath -Force | Out-Null
$registryPath = 'HKCU:\Software\Microsoft\Office\Word\Addins\ThesisAddIn.VSTO'
$registryNames = @('FriendlyName', 'Description', 'LoadBehavior', 'Manifest')

if (Test-Path -LiteralPath $statePath) {
    $state = [System.IO.File]::ReadAllText($statePath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    if ([System.IO.Path]::GetFullPath([string]$state.installRoot).TrimEnd('\') -ne $installPath) { throw '安装记录中的路径不匹配，拒绝更新。' }
    if (-not [string]::Equals([string]$state.publisherThumbprint, $packageCertificate.Thumbprint, [StringComparison]::OrdinalIgnoreCase)) {
        throw '新包的签名证书与已安装版本不同。请先卸载旧版并核验新发行者，再重新安装；没有覆盖旧回滚记录。'
    }
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

if (Test-Path -LiteralPath $installPath) {
    $existingItems = @((Get-Item -LiteralPath $installPath -Force)) + @(Get-ChildItem -LiteralPath $installPath -Recurse -Force)
    foreach ($item in $existingItems) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw '安装目录含重解析点，拒绝复制文件。' }
    }
}

foreach ($directory in @('vsto', 'engine', 'samples', 'docs', 'cli')) {
    $source = Join-Path $packageRoot $directory
    if ($directory -eq 'cli' -and -not (Test-Path -LiteralPath $source -PathType Container)) { continue }
    $target = Join-Path $installPath $directory
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Get-ChildItem -LiteralPath $source -Force | Copy-Item -Destination $target -Recurse -Force
}
foreach ($file in @('Install.ps1', 'Uninstall.ps1', 'Preflight.ps1', 'publisher.cer', 'build-info.json', 'SHA256SUMS.txt', 'README.md')) {
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
    checksumManifestSha256 = (Get-FileHash -LiteralPath $checksumPath -Algorithm SHA256).Hash.ToLowerInvariant()
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
