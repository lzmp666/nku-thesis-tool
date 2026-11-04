$git = "C:\Users\guo11\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe"
$cred = "protocol=https`nhost=github.com`n`n" | & $git credential fill 2>&1 | Out-String
$pass = (($cred -split "`n" | Where-Object { $_ -like "password=*" }) -replace "password=","").Trim()
$h = @{ Authorization = "Bearer $pass"; Accept = "application/vnd.github+json"; "User-Agent" = "Guoooooi" }
$repo = "lzmp666/nku-thesis-tool"

# 找 build-test 的失败日志
$runs = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/actions/runs?branch=feature/issue-96-chat" -Headers $h
foreach ($run in $runs.workflow_runs) {
    Write-Host "$($run.name): $($run.conclusion) ($($run.id))"
}
