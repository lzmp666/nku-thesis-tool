# Create a remote branch from a base ref, idempotently. -DryRun does GETs only.
#
#   .\push_branch.ps1 -Branch feature/issue-96-chat -DryRun
#   .\push_branch.ps1 -Branch feature/issue-96-chat
#   .\push_branch.ps1 -Repo owner/other-repo -Branch docs/xxx -Base main -DryRun
[CmdletBinding()]
param(
    [string]$Repo = $env:NKU_THESIS_REPO,
    [string]$Branch = $env:NKU_THESIS_BRANCH,
    [string]$Base = "main",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "scripts\Invoke-GitHubApi.ps1")

$git = Get-NkuGit
$slug = Get-NkuRepoSlug -Repo $Repo -Git $git
$name = Get-NkuBranch -Branch $Branch -Git $git -Base $Base -ForbidBase
$headers = Get-NkuGitHubHeaders -Git $git

$baseRef = Invoke-NkuGitHubApi -Path "repos/$slug/git/ref/heads/$Base" -Method Get -Headers $headers
$baseSha = $baseRef.object.sha
Write-Host "base:      $Base $baseSha"

$existing = $null
try {
    $existing = Invoke-NkuGitHubApi -Path "repos/$slug/git/ref/heads/$name" -Method Get -Headers $headers
} catch {
    # A missing ref is the expected case for a new branch.
}
if ($null -ne $existing -and $existing.object.sha) {
    Write-Host "already exists: $name $($existing.object.sha)"
    Write-Host "nothing created"
    return
}

$payload = @{ ref = "refs/heads/$name"; sha = $baseSha }
$created = Invoke-NkuGitHubApi -Path "repos/$slug/git/refs" -Method Post -Headers $headers -Body $payload -DryRun:$DryRun
if ($DryRun) {
    Write-Host "dry-run complete: no branch created"
    return
}
Write-Host "created:   $($created.ref) $($created.object.sha)"
