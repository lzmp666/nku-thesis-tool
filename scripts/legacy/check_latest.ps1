$git = "C:\Users\guo11\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe"
$cred = "protocol=https`nhost=github.com`n`n" | & $git credential fill 2>&1 | Out-String
$pass = (($cred -split "`n" | Where-Object { $_ -like "password=*" }) -replace "password=","").Trim()
$h = @{ Authorization = "Bearer $pass"; Accept = "application/vnd.github+json"; "User-Agent" = "Guoooooi" }
$repo = "lzmp666/nku-thesis-tool"

$runs = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/actions/runs?branch=feature/issue-96-chat" -Headers $h
$latestRun = $runs.workflow_runs[0]
Write-Host "Latest run: $($latestRun.id) - $($latestRun.conclusion)"

$jobs = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/actions/runs/$($latestRun.id)/jobs" -Headers $h
foreach ($job in $jobs.jobs) {
    Write-Host "=== $($job.name): $($job.conclusion) ==="
    foreach ($step in $job.steps) {
        Write-Host "  $($step.name): $($step.conclusion)"
    }
}
