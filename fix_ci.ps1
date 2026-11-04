$git = "C:\Users\guo11\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe"
$cred = "protocol=https`nhost=github.com`n`n" | & $git credential fill 2>&1 | Out-String
$pass = (($cred -split "`n" | Where-Object { $_ -like "password=*" }) -replace "password=","").Trim()
$h = @{ Authorization = "Bearer $pass"; Accept = "application/vnd.github+json"; "User-Agent" = "Guoooooi" }
$repo = "lzmp666/nku-thesis-tool"

# 获取当前分支的最新 commit
$ref = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/ref/heads/feature/issue-96-chat" -Headers $h
$oldCommitSha = $ref.object.sha
Write-Host "old commit: $oldCommitSha"

$oldCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/commits/$oldCommitSha" -Headers $h
$baseTreeSha = $oldCommit.tree.sha
Write-Host "base tree: $baseTreeSha"

# 要上传的文件（修改或新增的）
$files = @(
    "tests/ThesisEngine.Tests/ChatServiceTests.cs",
    "tests/ThesisEngine.Tests/ThesisEngine.Tests.csproj"
)

$treeEntries = @()
foreach ($file in $files) {
    Write-Host "Uploading $file ..."
    $bytes = [IO.File]::ReadAllBytes($file)
    $content = [Convert]::ToBase64String($bytes)
    $blobBody = @{ content = $content; encoding = "base64" } | ConvertTo-Json
    $maxRetries = 3
    for ($i = 0; $i -lt $maxRetries; $i++) {
        try {
            $blob = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/blobs" -Method Post -Headers $h -Body $blobBody -ContentType "application/json; charset=utf-8"
            Write-Host "  -> $($blob.sha)"
            $treeEntries += @{ path = $file; mode = "100644"; type = "blob"; sha = $blob.sha }
            break
        } catch {
            Write-Host "  retry $($i+1)"
            Start-Sleep -Seconds 2
        }
    }
    Start-Sleep -Milliseconds 500
}

# 删除旧的 ChatCommandTests.cs
$treeEntries += @{ path = "tests/ThesisEngine.Tests/ChatCommandTests.cs"; mode = "100644"; type = "blob"; sha = $null }

# 创建新 tree
Write-Host "Creating tree ..."
$treeBody = @{ base_tree = $baseTreeSha; tree = $treeEntries } | ConvertTo-Json -Depth 5
$newTree = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/trees" -Method Post -Headers $h -Body $treeBody -ContentType "application/json; charset=utf-8"
Write-Host "tree: $($newTree.sha)"

# 创建新 commit
Write-Host "Creating commit ..."
$commitBody = @{
    message = "fix(chat): move tests to Core, avoid referencing CLI exe in tests`n`n- Move ChatCommand tests to ChatServiceTests testing only Core types`n- Revert test project CLI reference to ReferenceOutputAssembly=false`n- Fix CI build-test failure"
    tree = $newTree.sha
    parents = @($oldCommitSha)
} | ConvertTo-Json
$newCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/commits" -Method Post -Headers $h -Body $commitBody -ContentType "application/json; charset=utf-8"
Write-Host "commit: $($newCommit.sha)"

# 更新分支
Write-Host "Updating branch ..."
$patchBody = @{ sha = $newCommit.sha; force = $false } | ConvertTo-Json
$branch = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/refs/heads/feature/issue-96-chat" -Method Patch -Headers $h -Body $patchBody -ContentType "application/json; charset=utf-8"
Write-Host "branch updated: $($branch.ref)"
