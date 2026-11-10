# Read-only: print the sha of a branch ref. Never writes to the remote.
#
#   .\check_branch.ps1
#   .\check_branch.ps1 -Repo lzmp666/nku-thesis-tool -Branch main
#   $env:NKU_THESIS_BRANCH = 'docs/issue-109-ai-onboarding'; .\check_branch.ps1
[CmdletBinding()]
param(
    [string]$Repo = $env:NKU_THESIS_REPO,
    [string]$Branch = $env:NKU_THESIS_BRANCH
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "scripts\Invoke-GitHubApi.ps1")

$git = Get-NkuGit
$slug = Get-NkuRepoSlug -Repo $Repo -Git $git
$name = Get-NkuBranch -Branch $Branch -Git $git -Base "main"
$headers = Get-NkuGitHubHeaders -Git $git

$ref = Invoke-NkuGitHubApi -Path "repos/$slug/git/ref/heads/$name" -Method Get -Headers $headers
Write-Host "repo:      $slug"
Write-Host "branch:    $name"
Write-Host "sha:       $($ref.object.sha)"
Write-Host "read-only: no remote object was created or changed"
