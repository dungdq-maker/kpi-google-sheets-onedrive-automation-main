$ErrorActionPreference = "Stop"

$WorkDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $WorkDir

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw "Git is not installed or not available in PATH."
}

if (-not (Test-Path (Join-Path $WorkDir ".git"))) {
    throw "This folder is not a Git repository. Run 'git init' first or clone from GitHub."
}

$branch = (& git branch --show-current).Trim()
if ([string]::IsNullOrWhiteSpace($branch)) {
    $branch = "main"
    & git branch -M $branch
}

$remoteNames = @(& git remote)
$remoteUrl = ""
if ($remoteNames -contains "origin") {
    $remoteUrl = (& git remote get-url origin).Trim()
}
if ([string]::IsNullOrWhiteSpace($remoteUrl)) {
    $remoteUrl = Read-Host "Nhap GitHub repo URL cho remote origin, vi du https://github.com/org/repo.git"
    if ([string]::IsNullOrWhiteSpace($remoteUrl)) {
        throw "Remote URL is required to push to GitHub."
    }
    if ($remoteNames -contains "origin") {
        & git remote set-url origin $remoteUrl
    } else {
        & git remote add origin $remoteUrl
    }
}

Write-Host ""
Write-Host "Files changed:"
& git status --short
Write-Host ""

$confirm = Read-Host "Ban co muon commit va push tat ca thay doi len GitHub khong? (Y/N)"
if ($confirm -notmatch "^(Y|y|Yes|yes)$") {
    Write-Host "Cancelled. No changes were pushed."
    exit 0
}

$message = Read-Host "Nhap commit message"
if ([string]::IsNullOrWhiteSpace($message)) {
    $message = "Update KPI automation"
}

& git add -A
$staged = (& git diff --cached --name-only)
if (-not $staged) {
    Write-Host "No staged changes to commit."
    exit 0
}

& git commit -m $message
if ($LASTEXITCODE -ne 0) {
    throw "git commit failed."
}

& git push -u origin $branch
if ($LASTEXITCODE -ne 0) {
    throw "git push failed. Check GitHub remote URL and authentication."
}

Write-Host ""
Write-Host "Pushed latest local changes to GitHub branch '$branch'."
