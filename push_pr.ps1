$git = "C:\Users\guo11\AppData\Local\GitHubDesktop\app-3.6.6\resources\app\git\cmd\git.exe"
$cred = "protocol=https`nhost=github.com`n`n" | & $git credential fill 2>&1 | Out-String
$pass = (($cred -split "`n" | Where-Object { $_ -like "password=*" }) -replace "password=","").Trim()
$h = @{ Authorization = "Bearer $pass"; Accept = "application/vnd.github+json"; "User-Agent" = "Guoooooi" }

$repo = "lzmp666/nku-thesis-tool"

# 1. 获取 main 的 commit sha 和 tree sha
$ref = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/ref/heads/main" -Headers $h
$mainSha = $ref.object.sha
Write-Host "main commit: $mainSha"

$commit = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/commits/$mainSha" -Headers $h
$treeSha = $commit.tree.sha
Write-Host "base tree: $treeSha"

# 2. 要上传的文件列表
$files = @(
    "src/ThesisEngine.Core/IChatService.cs",
    "src/ThesisEngine.Core/LocalRuleChatService.cs",
    "src/ThesisEngine.Cli/ChatCommand.cs",
    "src/ThesisEngine.Cli/Program.cs",
    "tests/ThesisEngine.Tests/ChatCommandTests.cs",
    "tests/ThesisEngine.Tests/ThesisEngine.Tests.csproj"
)

# 3. 为每个文件创建 blob
$treeEntries = @()
foreach ($file in $files) {
    $bytes = [IO.File]::ReadAllBytes($file)
    $content = [Convert]::ToBase64String($bytes)
    $blobBody = @{ content = $content; encoding = "base64" } | ConvertTo-Json
    $blob = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/blobs" -Method Post -Headers $h -Body $blobBody -ContentType "application/json; charset=utf-8"
    Write-Host "blob: $file -> $($blob.sha)"
    $treeEntries += @{ path = $file; mode = "100644"; type = "blob"; sha = $blob.sha }
}

# 4. 创建新 tree
$treeBody = @{ base_tree = $treeSha; tree = $treeEntries } | ConvertTo-Json -Depth 5
$newTree = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/trees" -Method Post -Headers $h -Body $treeBody -ContentType "application/json; charset=utf-8"
Write-Host "new tree: $($newTree.sha)"

# 5. 创建 commit
$commitBody = @{
    message = "feat(chat): add AI chat command for format questions (#96)`n`n- Add IChatService interface and LocalRuleChatService implementation`n- Add chat CLI command that answers format-related questions`n- Add unit tests for ChatCommand and LocalRuleChatService`n- First phase: local rule matching, future: replace with LLM"
    tree = $newTree.sha
    parents = @($mainSha)
} | ConvertTo-Json
$newCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/commits" -Method Post -Headers $h -Body $commitBody -ContentType "application/json; charset=utf-8"
Write-Host "new commit: $($newCommit.sha)"

# 6. 创建分支
$branchBody = @{ ref = "refs/heads/feature/issue-96-chat"; sha = $newCommit.sha } | ConvertTo-Json
$branch = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/git/refs" -Method Post -Headers $h -Body $branchBody -ContentType "application/json; charset=utf-8"
Write-Host "branch: $($branch.ref)"

# 7. 创建 PR
$prBody = @{
    title = "feat(chat): add AI chat command for format questions (#96)"
    head = "feature/issue-96-chat"
    base = "main"
    body = "## 关联 Issue`nCloses #96`n`n## 改动内容`n- 新增 IChatService 接口和 LocalRuleChatService 本地规则实现`n- 新增 chat 命令，用户问格式问题，AI 回答`n- 新增单元测试`n`n## 验收标准`n- 问一级标题用几号字，能回答黑体三号`n- 问正文用什么字体，能回答宋体小四`n- 问非格式问题，礼貌拒绝`n- Release build 0 错误，测试通过`n`n## 后续`n- 第一阶段先做本地规则匹配，后续替换为外接免费大模型"
} | ConvertTo-Json -Depth 3
$pr = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/pulls" -Method Post -Headers $h -Body $prBody -ContentType "application/json; charset=utf-8"
Write-Host "PR #$($pr.number): $($pr.html_url)"
