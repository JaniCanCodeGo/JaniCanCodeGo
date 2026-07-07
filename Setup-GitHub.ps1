<#
.SYNOPSIS
    One-shot GitHub setup for this repo on a local Windows machine.

.DESCRIPTION
    Installs Git and the GitHub CLI (if missing), signs you in to GitHub,
    configures your git identity, wires up the Git credential manager so
    future push/pull just work, and (optionally) clones or connects this
    repository.

    Run it from an ordinary PowerShell window - no admin rights are needed
    when installing per-user with winget. If winget is unavailable the
    script prints manual download links instead of failing.

.EXAMPLE
    # Recommended: allow scripts for this one session, then run it
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    .\Setup-GitHub.ps1

.EXAMPLE
    # Provide your name/email up front so it doesn't prompt
    .\Setup-GitHub.ps1 -GitName "Jane Doe" -GitEmail "jane@example.com"

.NOTES
    Repo: JaniCanCodeGo/JaniCanCodeGo
#>

[CmdletBinding()]
param(
    [string]$GitName,
    [string]$GitEmail,
    [string]$RepoSlug   = "JaniCanCodeGo/JaniCanCodeGo",
    [string]$ClonePath
)

$ErrorActionPreference = "Stop"

function Write-Step   ($m) { Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-Ok     ($m) { Write-Host "    [OK]   $m" -ForegroundColor Green }
function Write-Warn2  ($m) { Write-Host "    [WARN] $m" -ForegroundColor Yellow }
function Write-Info2  ($m) { Write-Host "    [INFO] $m" -ForegroundColor Gray }

function Test-Command($name) {
    return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

function Update-PathFromMachine {
    # winget-installed tools land on PATH but the current session doesn't see
    # them until we refresh from the registry.
    $machine = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $user    = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machine;$user"
}

function Install-WithWinget($id, $friendly, $downloadUrl) {
    if (Test-Command "winget") {
        Write-Info2 "Installing $friendly via winget..."
        winget install --id $id --exact --silent --accept-source-agreements --accept-package-agreements
        Update-PathFromMachine
    } else {
        Write-Warn2 "winget not found. Install $friendly manually from:"
        Write-Host  "           $downloadUrl" -ForegroundColor Yellow
        throw "$friendly is required but could not be auto-installed."
    }
}

Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  GitHub setup for $RepoSlug" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta

# ---------------------------------------------------------------------------
# 1. Git
# ---------------------------------------------------------------------------
Write-Step "Checking for Git"
if (Test-Command "git") {
    Write-Ok "Git is installed ($(git --version))"
} else {
    Install-WithWinget "Git.Git" "Git" "https://git-scm.com/download/win"
    if (-not (Test-Command "git")) {
        throw "Git still not on PATH. Close and reopen PowerShell, then re-run."
    }
    Write-Ok "Git installed ($(git --version))"
}

# ---------------------------------------------------------------------------
# 2. GitHub CLI
# ---------------------------------------------------------------------------
Write-Step "Checking for GitHub CLI (gh)"
if (Test-Command "gh") {
    Write-Ok "GitHub CLI is installed ($(gh --version | Select-Object -First 1))"
} else {
    Install-WithWinget "GitHub.cli" "GitHub CLI" "https://cli.github.com/"
    if (-not (Test-Command "gh")) {
        throw "gh still not on PATH. Close and reopen PowerShell, then re-run."
    }
    Write-Ok "GitHub CLI installed"
}

# ---------------------------------------------------------------------------
# 3. Git identity
# ---------------------------------------------------------------------------
Write-Step "Configuring your git identity"
$existingName  = (git config --global user.name)  2>$null
$existingEmail = (git config --global user.email) 2>$null

if (-not $GitName)  { $GitName  = $existingName }
if (-not $GitEmail) { $GitEmail = $existingEmail }

if (-not $GitName)  { $GitName  = Read-Host "    Your name (for commits)" }
if (-not $GitEmail) { $GitEmail = Read-Host "    Your GitHub email (for commits)" }

git config --global user.name  "$GitName"
git config --global user.email "$GitEmail"
Write-Ok "git identity set to: $GitName <$GitEmail>"

# ---------------------------------------------------------------------------
# 4. Authenticate to GitHub (this is the "grant permissions" step)
# ---------------------------------------------------------------------------
Write-Step "Signing in to GitHub"
$alreadyAuthed = $false
try {
    gh auth status 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { $alreadyAuthed = $true }
} catch { $alreadyAuthed = $false }

if ($alreadyAuthed) {
    Write-Ok "Already authenticated with GitHub."
} else {
    Write-Info2 "A browser window will open. Approve access to grant permissions."
    Write-Info2 "Choose: GitHub.com  ->  HTTPS  ->  'Login with a web browser'."
    # -w git => also configure git to use gh as the credential helper
    gh auth login --hostname github.com --git-protocol https --web
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub sign-in failed or was cancelled. Re-run the script to retry."
    }
    Write-Ok "Signed in to GitHub."
}

# Make git use gh's token for HTTPS operations - no PATs to manage.
Write-Step "Wiring up the git credential helper"
gh auth setup-git
Write-Ok "git will now use your GitHub login for push/pull over HTTPS."

# ---------------------------------------------------------------------------
# 5. Connect or clone the repo
# ---------------------------------------------------------------------------
Write-Step "Connecting the repository"
$insideRepo = $false
try {
    git rev-parse --is-inside-work-tree 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { $insideRepo = $true }
} catch { $insideRepo = $false }

if ($insideRepo) {
    Write-Ok "You're already inside a git working tree."
    Write-Info2 "Verifying you can reach GitHub..."
    git ls-remote --heads origin 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Ok "Success - your credentials work against 'origin'."
    } else {
        Write-Warn2 "Could not reach 'origin'. Check the remote URL: git remote -v"
    }
} else {
    if (-not $ClonePath) { $ClonePath = Join-Path (Get-Location) (Split-Path $RepoSlug -Leaf) }
    Write-Info2 "Cloning $RepoSlug into $ClonePath ..."
    gh repo clone $RepoSlug $ClonePath
    if ($LASTEXITCODE -eq 0) {
        Write-Ok "Cloned to $ClonePath"
        Write-Info2 "Next:  cd `"$ClonePath`""
    } else {
        Write-Warn2 "Clone failed. Confirm you have access to $RepoSlug."
    }
}

Write-Host "`n============================================" -ForegroundColor Magenta
Write-Host "  Done. GitHub is connected." -ForegroundColor Green
Write-Host "  Quick test:  git fetch; git status" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Magenta
