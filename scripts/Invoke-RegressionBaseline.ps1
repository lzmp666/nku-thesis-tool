# Issue #108 的三条主干回归链路，一条命令跑完并留痕到 artifacts/regression/。
#
#   powershell -ExecutionPolicy Bypass -File scripts\Invoke-RegressionBaseline.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\Invoke-RegressionBaseline.ps1 -SkipBuild
#
# 三条链路：
#   1. apply-figure-format -> analyze-figures：修复后复检应为 0 条表格问题
#   2. full-format -> analyze-figures / analyze-styles：套样式 + 改图表后表格不误报、
#      字体字号不残留「手动」误报
#   3. full-format 二次运行幂等：两个输出的 word/document.xml 压掉空白后必须完全一致。
#      比的是解包后的 document.xml，不是整个 docx 的哈希——zip 条目时间戳每次不同，
#      按字节比永远过不了。
#
# 退出码非 0 表示有链路未达标；所有命令输出都写在 artifacts/regression\*.log 里。

[CmdletBinding()]
param(
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
# CLI outputs UTF-8; align console encoding first or Chinese turns into mojibake
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot

$sdk = Join-Path $env:USERPROFILE ".dotnet\dotnet.exe"
$dotnet = if (Test-Path $sdk) { $sdk } else { "dotnet" }
$cliArgs = @("run", "--project", "src/ThesisEngine.Cli/ThesisEngine.Cli.csproj", "--configuration", "Release")
if ($SkipBuild) { $cliArgs += "--no-build" }

$inputDocx = "tests/fixtures/synthetic/full-thesis/synthetic_object_layout_errors.docx"
$artifacts = Join-Path $repoRoot "artifacts/regression"
New-Item -ItemType Directory -Force $artifacts | Out-Null

$failures = 0

function Invoke-Cli {
    param([string[]]$Arguments, [string]$Log)
    Write-Host "  dotnet $($Arguments -join ' ')"
    & $dotnet @cliArgs @Arguments 2>&1 | Tee-Object -FilePath (Join-Path $artifacts $Log) | Write-Host
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "命令失败（退出码 $LASTEXITCODE）：$($Arguments -join ' ')"
        $script:failures++
        return $false
    }
    return $true
}

function Get-CanonicalDocumentXml {
    param([string]$DocxPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($DocxPath)
    try {
        $entry = $zip.GetEntry("word/document.xml")
        $reader = New-Object IO.StreamReader($entry.Open())
        $text = $reader.ReadToEnd()
        $reader.Close()
        return ($text -replace ">\s+<", "><").Trim()
    }
    finally { $zip.Dispose() }
}

function Assert-TableIssuesZero {
    param([string]$Log, [string]$Step)
    $report = Get-Content (Join-Path $artifacts $Log) -Raw
    if ($report -match "表格.*内边框|内部横线或竖线|未显式设置三线表边框") {
        Write-Warning "$Step 复检仍报表格问题（见 artifacts/regression/$Log）"
        $script:failures++
    }
    else {
        Write-Host "  [PASS] $Step 复检 0 条表格问题"
    }
}

Write-Host "=== 链路 1：apply-figure-format -> analyze-figures ==="
$fig1 = Join-Path $artifacts "regression_fig1.docx"
if (Test-Path $fig1) { Remove-Item $fig1 -Force }
if (Invoke-Cli @("apply-figure-format", $inputDocx, $fig1) -Log "1-apply-figure-format.log") {
    Invoke-Cli @("analyze-figures", $fig1) -Log "1-recheck-figures.log" | Out-Null
    Assert-TableIssuesZero -Log "1-recheck-figures.log" -Step "链路 1"
}

Write-Host ""
Write-Host "=== 链路 2：full-format -> analyze-figures / analyze-styles ==="
$full1 = Join-Path $artifacts "regression_full1.docx"
if (Test-Path $full1) { Remove-Item $full1 -Force }
if (Invoke-Cli @("full-format", $inputDocx, $full1) -Log "2-full-format.log") {
    Invoke-Cli @("analyze-figures", $full1) -Log "2-recheck-figures.log" | Out-Null
    Assert-TableIssuesZero -Log "2-recheck-figures.log" -Step "链路 2（图表）"

    Invoke-Cli @("analyze-styles", $full1) -Log "2-recheck-styles.log" | Out-Null
    $styleReport = Get-Content (Join-Path $artifacts "2-recheck-styles.log") -Raw
    if ($styleReport -match "手动字体|手动字号") {
        Write-Warning "链路 2 复检仍报「手动字体/手动字号」（见 artifacts/regression/2-recheck-styles.log）"
        $script:failures++
    }
    else {
        Write-Host "  [PASS] 链路 2 复检无字体字号误报"
    }
}

Write-Host ""
Write-Host "=== 链路 3：full-format 二次运行幂等（比 document.xml） ==="
$full2 = Join-Path $artifacts "regression_full2.docx"
if (Test-Path $full2) { Remove-Item $full2 -Force }
if (Invoke-Cli @("full-format", $inputDocx, $full2) -Log "3-full-format-second-run.log") {
    $first = Get-CanonicalDocumentXml $full1
    $second = Get-CanonicalDocumentXml $full2
    if ($first -ceq $second) {
        Write-Host "  [PASS] 两次输出的 document.xml 完全一致（$($first.Length) 字符）"
        Set-Content (Join-Path $artifacts "3-documentxml-compare.txt") -Value "IDENTICAL $($first.Length) chars"
    }
    else {
        Write-Warning "两次输出的 document.xml 不一致（$($first.Length) vs $($second.Length) 字符）"
        Set-Content (Join-Path $artifacts "3-documentxml-compare.txt") -Value "DIFFERENT $($first.Length) vs $($second.Length) chars"
        $script:failures++
    }
}

Write-Host ""
if ($failures -gt 0) {
    Write-Host "回归基线：$failures 项未达标，留痕见 artifacts/regression/" -ForegroundColor Red
    exit 1
}
Write-Host "回归基线：三条链路全部达标，留痕见 artifacts/regression/" -ForegroundColor Green
