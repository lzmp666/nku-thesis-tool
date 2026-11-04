[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$checks = [System.Collections.Generic.List[object]]::new()

function Add-Check([string]$Name, [bool]$Passed, [string]$Detail) {
    $checks.Add([pscustomobject]@{ 项目 = $Name; 结果 = if ($Passed) { '通过' } else { '缺少/未识别' }; 说明 = $Detail })
}

$wordCandidates = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
)
$wordPath = $null
foreach ($key in $wordCandidates) {
    try {
        $candidate = (Get-Item -LiteralPath $key -ErrorAction Stop).GetValue('')
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { $wordPath = $candidate; break }
    } catch { }
}
Add-Check '桌面版 Word' ([bool]$wordPath) $(if ($wordPath) { $wordPath } else { '未从 Windows App Paths 注册表项识别；请确认已安装桌面版 Word' })
if ($wordPath) {
    $wordBitness = if ($wordPath -match '\\Program Files \(x86\)\\') { '32 位（按安装路径识别）' } else { '64 位（按安装路径识别）' }
    Add-Check 'Word 位数' $true $wordBitness
} else {
    Add-Check 'Word 位数' $false '未识别桌面版 Word'
}

$vstoKeys = @(
    'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R',
    'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4'
)
$vstoVersion = $null
foreach ($key in $vstoKeys) {
    try {
        $item = Get-ItemProperty -LiteralPath $key -ErrorAction Stop
        $vstoVersion = $item.Version
        if (-not $vstoVersion) { $vstoVersion = '已注册' }
        break
    } catch { }
}
Add-Check 'VSTO Runtime' ([bool]$vstoVersion) $(if ($vstoVersion) { [string]$vstoVersion } else { '未识别；请安装 Microsoft Visual Studio Tools for Office Runtime' })

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
$net10 = $null
if ($dotnet) {
    $runtimeText = (& $dotnet.Source --list-runtimes 2>$null | Out-String)
    $net10 = $runtimeText -split "`r?`n" | Where-Object { $_ -match '^Microsoft\.NETCore\.App 10\.' } | Select-Object -First 1
}
Add-Check '.NET 10 Runtime' ([bool]$net10) $(if ($net10) { $net10.Trim() } else { '未识别；演示引擎宿主需要 .NET 10 Runtime' })

$frameworkRelease = 0
foreach ($key in @('HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\NET Framework Setup\NDP\v4\Full')) {
    try {
        $release = [int](Get-ItemProperty -LiteralPath $key -Name Release -ErrorAction Stop).Release
        if ($release -gt $frameworkRelease) { $frameworkRelease = $release }
    } catch { }
}
Add-Check '.NET Framework 4.8+' ($frameworkRelease -ge 528040) $(if ($frameworkRelease -ge 528040) { "Release $frameworkRelease" } else { 'VSTO 插件需要 .NET Framework 4.8 或更高版本' })

$wordBuild = '未读取（可在 Word：文件 → 帐户 → 关于 Word 查看）'
if ($wordPath) {
    try { $wordBuild = (Get-Item -LiteralPath $wordPath).VersionInfo.FileVersion } catch { }
}
Add-Check 'Word 文件版本' ([bool]$wordPath) $wordBuild

"" | Write-Output
$checks | Format-Table -AutoSize | Out-String -Width 220 | Write-Output
if ($checks.Where({ -not $_.结果.Equals('通过') }).Count -gt 0) {
    Write-Error '环境预检未通过。请按上表处理缺少项后再安装加载项。'
    exit 1
}
Write-Output '预检完成。安装脚本只写入当前 Windows 用户的 Word 加载项注册。'
