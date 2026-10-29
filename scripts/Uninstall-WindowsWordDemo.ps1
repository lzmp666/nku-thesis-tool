[CmdletBinding()]
param([string]$InstallRoot = 'D:\NkuThesisTool')

$ErrorActionPreference = 'Stop'
$installPath = [System.IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
$driveRoot = [System.IO.Path]::GetPathRoot($installPath).TrimEnd('\')
if ($installPath -eq $driveRoot -or $installPath.Length -lt 8) { throw '拒绝删除磁盘根目录或过短路径。' }
$statePath = Join-Path $installPath '.nku-thesis-tool-install-state.json'
if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { throw "没有找到本工具的安装记录，未更改任何设置：$statePath" }
$state = [System.IO.File]::ReadAllText($statePath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
if ([System.IO.Path]::GetFullPath([string]$state.installRoot).TrimEnd('\') -ne $installPath) { throw '安装记录中的路径不匹配，拒绝继续。' }
$checksumPath = Join-Path $installPath 'SHA256SUMS.txt'
if (-not (Test-Path -LiteralPath $checksumPath -PathType Leaf)) { throw '缺少包内 SHA256SUMS.txt，拒绝更改当前用户设置或删除文件。' }
$expectedChecksumManifestHash = [string]$state.checksumManifestSha256
if ($expectedChecksumManifestHash -notmatch '^[0-9a-fA-F]{64}$') { throw '安装记录缺少有效的校验清单指纹，拒绝更改当前用户设置或删除文件。' }
$actualChecksumManifestHash = (Get-FileHash -LiteralPath $checksumPath -Algorithm SHA256).Hash
if (-not [string]::Equals($expectedChecksumManifestHash, $actualChecksumManifestHash, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'SHA256SUMS.txt 已被修改，拒绝更改当前用户设置或自动删除文件。'
}
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { throw '请先关闭所有 Word 窗口，再卸载插件；当前配置未更改。' }
$existingItems = @((Get-Item -LiteralPath $installPath -Force)) + @(Get-ChildItem -LiteralPath $installPath -Recurse -Force)
foreach ($item in $existingItems) {
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw '安装目录含重解析点，拒绝更改设置或删除文件。' }
}

$registryPath = 'HKCU:\Software\Microsoft\Office\Word\Addins\ThesisAddIn.VSTO'
if (Test-Path -LiteralPath $registryPath) {
    $current = Get-ItemProperty -LiteralPath $registryPath
    if ([string]::Equals([string]$current.Manifest, [string]$state.installedManifest, [StringComparison]::OrdinalIgnoreCase)) {
        foreach ($name in @('FriendlyName', 'Description', 'LoadBehavior', 'Manifest')) {
            Remove-ItemProperty -LiteralPath $registryPath -Name $name -ErrorAction SilentlyContinue
        }
        if ([bool]$state.previousRegistryExisted) {
            foreach ($property in $state.previousRegistryValues.PSObject.Properties) {
                if ($property.Name -eq 'LoadBehavior') {
                    New-ItemProperty -LiteralPath $registryPath -Name $property.Name -Value ([int]$property.Value) -PropertyType DWord -Force | Out-Null
                } else {
                    New-ItemProperty -LiteralPath $registryPath -Name $property.Name -Value $property.Value -PropertyType String -Force | Out-Null
                }
            }
        } else {
            $remaining = Get-ItemProperty -LiteralPath $registryPath
            $otherValues = @($remaining.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' })
            if ($otherValues.Count -eq 0) { Remove-Item -LiteralPath $registryPath -Force }
        }
    } else {
        Write-Warning 'Word 加载项注册已被用户更改，保留当前注册项。'
    }
}

$currentEngineHost = [Environment]::GetEnvironmentVariable('THESIS_ENGINE_PROTOCOL_HOST', 'User')
if ([string]::Equals($currentEngineHost, [string]$state.installedEngineHost, [StringComparison]::OrdinalIgnoreCase)) {
    [Environment]::SetEnvironmentVariable('THESIS_ENGINE_PROTOCOL_HOST', [string]$state.previousEngineHost, 'User')
} else {
    Write-Warning '引擎宿主环境变量已被更改，保留当前值。'
}

foreach ($storeName in @(@{ Name = 'Root'; Added = [bool]$state.addedPublisherToRoot }, @{ Name = 'TrustedPublisher'; Added = [bool]$state.addedPublisherToTrustedPublisher })) {
    if (-not $storeName.Added) { continue }
    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($storeName.Name, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    try {
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $matches = @($store.Certificates | Where-Object { $_.Thumbprint -eq [string]$state.publisherThumbprint })
        foreach ($certificate in $matches) { $store.Remove($certificate) }
    } finally { $store.Close() }
}

foreach ($line in (Get-Content -LiteralPath $checksumPath)) {
    if ($line -notmatch '^([0-9a-f]{64})  (.+)$') { continue }
    $expectedHash = $Matches[1]
    $relative = $Matches[2].Replace('/', '\')
    if ($relative.Split('\') -contains '..' -or [System.IO.Path]::IsPathRooted($relative)) { throw '包校验清单含无效路径，拒绝删除文件。' }
    $filePath = [System.IO.Path]::GetFullPath((Join-Path $installPath $relative))
    if (-not $filePath.StartsWith($installPath + '\', [StringComparison]::OrdinalIgnoreCase)) { throw '包文件路径越出安装目录，拒绝删除。' }
    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { continue }
    $actualHash = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualHash -eq $expectedHash) { Remove-Item -LiteralPath $filePath -Force }
    else { Write-Warning "文件已被修改，保留：$relative" }
}
if (Test-Path -LiteralPath $checksumPath -PathType Leaf) {
    $currentChecksumManifestHash = (Get-FileHash -LiteralPath $checksumPath -Algorithm SHA256).Hash
    if ([string]::Equals($expectedChecksumManifestHash, $currentChecksumManifestHash, [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $checksumPath -Force
    } else {
        Write-Warning '卸载期间 SHA256SUMS.txt 发生变化，保留该文件。'
    }
}
Remove-Item -LiteralPath $statePath -Force
foreach ($relativeDirectory in @('vsto', 'engine', 'samples', 'docs', 'cli')) {
    $directoryPath = Join-Path $installPath $relativeDirectory
    if ((Test-Path -LiteralPath $directoryPath -PathType Container) -and -not (Get-ChildItem -LiteralPath $directoryPath -Force | Select-Object -First 1)) {
        Remove-Item -LiteralPath $directoryPath -Force
    }
}
$leftovers = @(Get-ChildItem -LiteralPath $installPath -Force -ErrorAction SilentlyContinue)
if ($leftovers.Count -eq 0) { Remove-Item -LiteralPath $installPath -Force }
else { Write-Warning "卸载完成，安装目录保留用户文件：$installPath" }
Write-Output '已卸载本演示插件并恢复本脚本登记前的当前用户配置。请重新启动 Word。'
