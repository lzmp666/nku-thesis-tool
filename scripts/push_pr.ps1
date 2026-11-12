# Upload local files through the GitHub data API and open a PR from a new ref.
# This is a fallback for machines where a normal git push is blocked; the daily
# workflow is plain `git switch` / `git push`. -DryRun performs GETs only.
#
#   .\push_pr.ps1 -Branch feature/xxx -Files @('README.md') `
#       -CommitMessage 'docs: update readme' -Title 'docs: update readme' `
#       -Body 'Closes #123' -DryRun
[CmdletBinding()]
param(
    [string]$Repo = $env:NKU_THESIS_REPO,
    [string]$Branch = $env:NKU_THESIS_BRANCH,
    [string]$Base = "main",
    [string[]]$Files = @(),
    [string]$CommitMessage = "",
    [string]$Title = "",
    [string]$Body = "",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Invoke-GitHubApi.ps1")

$git = Get-NkuGit
$slug = Get-NkuRepoSlug -Repo $Repo -Git $git
$name = Get-NkuBranch -Branch $Branch -Git $git -Base $Base -ForbidBase
$headers = Get-NkuGitHubHeaders -Git $git

$paths = @($Files | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
if ($paths.Count -eq 0) { throw "Pass at least one existing file with -Files." }
# The API needs repo-relative tree paths. Resolve each file and reject anything
# outside the repository instead of sending an absolute Windows path to GitHub.
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path.TrimEnd("\")
$treePaths = @()
foreach ($path in $paths) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing file: $path" }
    $full = (Resolve-Path -LiteralPath $path).Path
    if (-not $full.StartsWith($repoRoot + "\", [StringComparison]::OrdinalIgnoreCase)) {
        throw "File is outside the repository: $path"
    }
    $treePaths += $full.Substring($repoRoot.Length + 1) -replace "\\", "/"
}
if (-not $DryRun) {
    if ([string]::IsNullOrWhiteSpace($CommitMessage)) { throw "Pass -CommitMessage before a real run." }
    if ([string]::IsNullOrWhiteSpace($Title)) { throw "Pass -Title before a real run." }
    if ([string]::IsNullOrWhiteSpace($Body)) { throw "Pass -Body before a real run." }
}

$baseRef = Invoke-NkuGitHubApi -Path "repos/$slug/git/ref/heads/$Base" -Method Get -Headers $headers
$baseSha = $baseRef.object.sha
$commit = Invoke-NkuGitHubApi -Path "repos/$slug/git/commits/$baseSha" -Method Get -Headers $headers
Write-Host "base:      $Base $baseSha (tree $($commit.tree.sha))"

$treeEntries = @()
for ($i = 0; $i -lt $paths.Count; $i++) {
    $path = $paths[$i]
    $bytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $path).Path)
    $blobBody = @{ content = [Convert]::ToBase64String($bytes); encoding = "base64" }
    $blob = Invoke-NkuGitHubApi -Path "repos/$slug/git/blobs" -Method Post -Headers $headers -Body $blobBody -DryRun:$DryRun
    $blobSha = if ($null -eq $blob) { "<pending-blob>" } else { $blob.sha }
    $treeEntries += @{
        path = $treePaths[$i]
        mode = "100644"
        type = "blob"
        sha  = $blobSha
    }
    Write-Host "file:      $($treePaths[$i]) ($($bytes.Length) bytes) -> $blobSha"
}

$treeBody = @{ base_tree = $commit.tree.sha; tree = $treeEntries }
$newTree = Invoke-NkuGitHubApi -Path "repos/$slug/git/trees" -Method Post -Headers $headers -Body $treeBody -DryRun:$DryRun
$treeSha = if ($null -eq $newTree) { "<pending-tree>" } else { $newTree.sha }

$commitBody = @{
    message = $CommitMessage
    tree    = $treeSha
    parents = @($baseSha)
}
$newCommit = Invoke-NkuGitHubApi -Path "repos/$slug/git/commits" -Method Post -Headers $headers -Body $commitBody -DryRun:$DryRun
$commitSha = if ($null -eq $newCommit) { "<pending-commit>" } else { $newCommit.sha }

$existing = $null
try {
    $existing = Invoke-NkuGitHubApi -Path "repos/$slug/git/ref/heads/$name" -Method Get -Headers $headers
} catch {
    # A missing ref is the expected case for a new branch.
}
if ($null -eq $existing -or -not $existing.object.sha) {
    $refBody = @{ ref = "refs/heads/$name"; sha = $commitSha }
    Invoke-NkuGitHubApi -Path "repos/$slug/git/refs" -Method Post -Headers $headers -Body $refBody -DryRun:$DryRun | Out-Null
    Write-Host "ref:       refs/heads/$name"
} else {
    Write-Host "ref:       $name already exists at $($existing.object.sha); not overwritten"
}

$prBody = @{ title = $Title; head = $name; base = $Base; body = $Body }
$pr = Invoke-NkuGitHubApi -Path "repos/$slug/pulls" -Method Post -Headers $headers -Body $prBody -DryRun:$DryRun
if ($DryRun) {
    Write-Host "dry-run complete: no blob/tree/commit/ref/PR written"
    return
}
Write-Host "PR #$($pr.number): $($pr.html_url)"
