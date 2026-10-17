[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$subject = 'CN=NKU Thesis Tool Local Development'
$personalStore = 'Cert:\CurrentUser\My'
$now = Get-Date

$certificate = Get-ChildItem -Path $personalStore |
    Where-Object { $_.Subject -eq $subject -and $_.HasPrivateKey -and $_.NotAfter -gt $now } |
    Sort-Object -Property NotAfter -Descending |
    Select-Object -First 1

if (-not $certificate) {
    $certificate = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $subject `
        -CertStoreLocation $personalStore `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -KeyExportPolicy NonExportable `
        -NotAfter $now.AddYears(2)
}

$certificateDirectory = Join-Path $env:LOCALAPPDATA 'NkuThesisTool'
New-Item -ItemType Directory -Path $certificateDirectory -Force | Out-Null
$publicCertificatePath = Join-Path $certificateDirectory 'vsto-local-development.cer'
Export-Certificate -Cert $certificate -FilePath $publicCertificatePath -Force | Out-Null

foreach ($storeName in @('Root', 'TrustedPublisher')) {
    $storePath = "Cert:\CurrentUser\$storeName"
    $alreadyTrusted = Get-ChildItem -Path $storePath |
        Where-Object { $_.Thumbprint -eq $certificate.Thumbprint } |
        Select-Object -First 1
    if (-not $alreadyTrusted) {
        Import-Certificate -FilePath $publicCertificatePath -CertStoreLocation $storePath | Out-Null
    }
}

[Environment]::SetEnvironmentVariable('THESIS_VSTO_CERT_THUMBPRINT', $certificate.Thumbprint, 'User')
$env:THESIS_VSTO_CERT_THUMBPRINT = $certificate.Thumbprint

Write-Output "已创建或复用本机 VSTO 开发证书：$($certificate.Thumbprint)"
Write-Output "私钥仅保存在当前用户的个人证书库，不可导出；公开证书：$publicCertificatePath"
Write-Output '已设置用户环境变量 THESIS_VSTO_CERT_THUMBPRINT，并仅信任当前用户的本地开发证书。'
