$git = "C:\Users\guo11\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe"
$cred = "protocol=https`nhost=github.com`n`n" | & $git credential fill 2>&1 | Out-String
$pass = (($cred -split "`n" | Where-Object { $_ -like "password=*" }) -replace "password=","").Trim()
$h = @{ Authorization = "Bearer $pass"; Accept = "application/vnd.github+json"; "User-Agent" = "Guoooooi" }

# 先获取 main 的最新 sha
$ref = Invoke-RestMethod -Uri "https://api.github.com/repos/lzmp666/nku-thesis-tool/git/ref/heads/main" -Headers $h
$mainSha = $ref.object.sha
Write-Host "main: $mainSha"

# 创建分支
$body = @{ ref = "refs/heads/feature/issue-96-chat"; sha = $mainSha } | ConvertTo-Json
$branch = Invoke-RestMethod -Uri "https://api.github.com/repos/lzmp666/nku-thesis-tool/git/refs" -Method Post -Headers $h -Body $body -ContentType "application/json; charset=utf-8"
Write-Host "branch: $($branch.ref)"
