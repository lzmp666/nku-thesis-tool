# Shared plumbing for the three early GitHub API helper scripts at the repo root.
# Dot-source this file from those scripts; function names are prefixed on purpose.
#
# - git is resolved from PATH, never from a machine-specific absolute path.
# - The token comes from `git credential fill`; nothing is written to disk.
# - Repo / branch come from -Repo / -Branch, then NKU_THESIS_REPO / NKU_THESIS_BRANCH,
#   then the current git remote / branch.
# - Every POST goes through Invoke-NkuGitHubApi -DryRun, so a caller's -DryRun
#   performs read-only GETs only and creates no remote objects.

function Get-NkuGit {
    $cmd = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $cmd) { throw "git was not found on PATH. Install Git or add it to PATH." }
    return $cmd.Source
}

function Get-NkuRepoSlug {
    param([string]$Repo, [string]$Git)
    if (-not [string]::IsNullOrWhiteSpace($Repo)) { return $Repo.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($env:NKU_THESIS_REPO)) { return $env:NKU_THESIS_REPO.Trim() }
    $url = (& $Git remote get-url origin 2>$null | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($url)) {
        throw "Cannot resolve the repository. Pass -Repo or set NKU_THESIS_REPO."
    }
    $url = $url.Trim()
    if ($url -match '^https?://[^/]+/([^/]+/[^/]+?)(?:\.git)?/?$') { return $Matches[1] }
    if ($url -match '^git@[^:]+:([^/]+/[^/]+?)(?:\.git)?$') { return $Matches[1] }
    throw "Cannot parse owner/name from origin URL: $url. Pass -Repo explicitly."
}

function Get-NkuBranch {
    param([string]$Branch, [string]$Git, [string]$Base, [switch]$ForbidBase)
    $name = $Branch
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $env:NKU_THESIS_BRANCH }
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = (& $Git rev-parse --abbrev-ref HEAD 2>$null | Select-Object -First 1)
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        throw "Cannot resolve the branch. Pass -Branch or set NKU_THESIS_BRANCH."
    }
    $name = $name.Trim()
    if ($ForbidBase -and $name -eq $Base) {
        throw "Refusing to use the base branch '$Base' as the target. Pass -Branch explicitly."
    }
    return $name
}

function Get-NkuGitHubHeaders {
    param([string]$Git)
    $cred = "protocol=https`nhost=github.com`n`n" | & $Git credential fill 2>&1 | Out-String
    $line = $cred -split "`n" | Where-Object { $_ -like "password=*" } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($line)) {
        throw "git credential fill returned no password. Check your GitHub credentials."
    }
    $token = ($line -replace "password=", "").Trim()
    if ([string]::IsNullOrWhiteSpace($token)) {
        throw "git credential fill returned an empty token."
    }
    return @{
        Authorization = "Bearer $token"
        Accept        = "application/vnd.github+json"
        "User-Agent"  = "nku-thesis-tool-script"
    }
}

function Invoke-NkuGitHubApi {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateSet("Get", "Post")][string]$Method = "Get",
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [object]$Body,
        [switch]$DryRun
    )
    $uri = "https://api.github.com/" + $Path.TrimStart("/")
    if ($DryRun) {
        Write-Host "[dry-run] would $Method $uri"
        if ($null -ne $Body) {
            $json = $Body | ConvertTo-Json -Depth 8 -Compress
            if ($json.Length -gt 400) {
                $json = $json.Substring(0, 400) + "... (truncated, $($json.Length) chars total)"
            }
            Write-Host "[dry-run] payload: $json"
        }
        return $null
    }
    $params = @{
        Uri         = $uri
        Method      = $Method
        Headers     = $Headers
        TimeoutSec  = 60
    }
    if ($null -ne $Body) {
        $params.Body = $Body | ConvertTo-Json -Depth 8
        $params.ContentType = "application/json; charset=utf-8"
    }
    return Invoke-RestMethod @params
}
