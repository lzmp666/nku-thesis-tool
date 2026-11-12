$git = "C:\Users\guo11\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe"
$cred = "protocol=https`nhost=github.com`n`n" | & $git credential fill 2>&1 | Out-String
$pass = (($cred -split "`n" | Where-Object { $_ -like "password=*" }) -replace "password=","").Trim()
$h = @{ Authorization = "Bearer $pass"; Accept = "application/vnd.github+json"; "User-Agent" = "Guoooooi" }
$repo = "lzmp666/nku-thesis-tool"

# 获取 main 最新 sha
Write-Host "Fetching main ref ..."
$ref = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/ref/heads/main" -Headers $h
$mainSha = $ref.object.sha
Write-Host "main: $mainSha"

$commit = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/commits/$mainSha" -Headers $h
$treeSha = $commit.tree.sha
Write-Host "base tree: $treeSha"

# 要上传的文件
$files = @(
    "src/ThesisEngine.Core/IChatService.cs",
    "src/ThesisEngine.Core/LocalRuleChatService.cs",
    "src/ThesisEngine.Cli/ChatCommand.cs",
    "src/ThesisEngine.Cli/Program.cs",
    "tests/ThesisEngine.Tests/ChatCommandTests.cs",
    "tests/ThesisEngine.Tests/ThesisEngine.Tests.csproj"
)

# 逐个上传 blob
$treeEntries = @()
foreach ($file in $files) {
    Write-Host "Uploading $file ..."
    $bytes = [IO.File]::ReadAllBytes($file)
    $content = [Convert]::ToBase64String($bytes)
    $blobBody = @{ content = $content; encoding = "base64" } | ConvertTo-Json
    $maxRetries = 3
    $success = $false
    for ($i = 0; $i -lt $maxRetries; $i++) {
        try {
            $blob = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/blobs" -Method Post -Headers $h -Body $blobBody -ContentType "application/json; charset=utf-8"
            Write-Host "  -> $($blob.sha)"
            $treeEntries += @{ path = $file; mode = "100644"; type = "blob"; sha = $blob.sha }
            $success = $true
            break
        } catch {
            Write-Host "  retry $($i+1): $($_.Exception.Message)"
            Start-Sleep -Seconds 2
        }
    }
    if (-not $success) {
        Write-Host "  FAILED after $maxRetries retries"
        exit 1
    }
    Start-Sleep -Milliseconds 500
}

# 创建 tree
Write-Host "Creating tree ..."
$treeBody = @{ base_tree = $treeSha; tree = $treeEntries } | ConvertTo-Json -Depth 5
$newTree = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/trees" -Method Post -Headers $h -Body $treeBody -ContentType "application/json; charset=utf-8"
Write-Host "tree: $($newTree.sha)"

# 创建 commit
Write-Host "Creating commit ..."
$commitBody = @{
    message = "feat(chat): add AI chat command for format questions (#96)`n`n- Add IChatService interface and LocalRuleChatService`n- Add chat CLI command for format Q&A`n- Add unit tests`n- First phase: local rule matching, future: LLM integration"
    tree = $newTree.sha
    parents = @($mainSha)
} | ConvertTo-Json
$newCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/commits" -Method Post -Headers $h -Body $commitBody -ContentType "application/json; charset=utf-8"
Write-Host "commit: $($newCommit.sha)"

# 创建分支
Write-Host "Creating branch ..."
$branchBody = @{ ref = "refs/heads/feature/issue-96-chat"; sha = $newCommit.sha } | ConvertTo-Json
try {
    $branch = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/refs" -Method Post -Headers $h -Body $branchBody -ContentType "application/json; charset=utf-8"
    Write-Host "branch: $($branch.ref)"
} catch {
    Write-Host "branch already exists, updating ..."
    $patchBody = @{ sha = $newCommit.sha; force = $true } | ConvertTo-Json
    $branch = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/refs/heads/feature/issue-96-chat" -Method Patch -Headers $h -Body $patchBody -ContentType "application/json; charset=utf-8"
    Write-Host "branch updated: $($branch.ref)"
}

# 创建 PR
Write-Host "Creating PR ..."
$prBody = @{
    title = "feat(chat): add AI chat command for format questions (#96)"
    head = "feature/issue-96-chat"
    base = "main"
    body = "## 关联 Issue`nCloses #96`n`n## 改动内容`n- 新增 IChatService 接口和 LocalRuleChatService 本地规则实现`n- 新增 chat 命令，用户问格式问题，AI 回答`n- 新增单元测试`n`n## 验收标准`n- 问一级标题用几号字，能回答黑体三号`n- 问正文用什么字体，能回答宋体小四`n- 问非格式问题，礼貌拒绝`n- Release build 0 错误，测试通过`n`n## 后续`n- 第一阶段先做本地规则匹配，后续替换为外接免费大模型"
} | ConvertTo-Json -Depth 3
try {
    $pr = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/pulls" -Method Post -Headers $h -Body $prBody -ContentType "application/json; charset=utf-8"
    Write-Host "PR #$($pr.number): $($pr.html_url)"
} catch {
    Write-Host "PR may already exist: $($_.Exception.Message)"
}
