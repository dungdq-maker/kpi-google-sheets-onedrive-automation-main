$ErrorActionPreference = "Stop"

$WorkDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogDir  = Join-Path $WorkDir "logs"
$LogFile = Join-Path $LogDir ("kpi_automation_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
Set-Location $WorkDir

$env:PYTHONIOENCODING = "utf-8"

function Write-Log {
    param([string]$Message)
    $Message | Tee-Object -FilePath $LogFile -Append
}

function Test-GitRepository {
    return (Test-Path (Join-Path $WorkDir ".git"))
}

function Update-FromGitHubIfApproved {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-Log "Git not found. Skipping update check."
        return
    }

    if (-not (Test-GitRepository)) {
        Write-Log "This folder is not a Git repository. Skipping update check."
        return
    }

    $remoteUrl = (& git remote get-url origin 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($remoteUrl)) {
        Write-Log "Git remote 'origin' is not configured. Skipping update check."
        return
    }

    $branch = (& git branch --show-current).Trim()
    if ([string]::IsNullOrWhiteSpace($branch)) {
        Write-Log "Cannot detect current Git branch. Skipping update check."
        return
    }

    Write-Log "Checking GitHub updates from $remoteUrl ($branch)..."
    & git fetch origin $branch 2>&1 | Tee-Object -FilePath $LogFile -Append
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Git fetch failed. Continuing with local code."
        return
    }

    $localCommit = (& git rev-parse HEAD 2>$null).Trim()
    $remoteCommit = (& git rev-parse "origin/$branch" 2>$null).Trim()
    if ([string]::IsNullOrWhiteSpace($localCommit) -or [string]::IsNullOrWhiteSpace($remoteCommit)) {
        Write-Log "Cannot compare local/remote commits. Continuing with local code."
        return
    }

    if ($localCommit -eq $remoteCommit) {
        Write-Log "Code is already up to date."
        return
    }

    $answer = Read-Host "Co ban cap nhat moi tren GitHub. Ban co muon update code truoc khi chay automation? (Y/N)"
    if ($answer -notmatch "^(Y|y|Yes|yes)$") {
        Write-Log "User skipped GitHub update."
        return
    }

    $dirty = (& git status --porcelain)
    if ($dirty) {
        Write-Log "Local working tree has uncommitted changes. Skipping pull to avoid overwriting local edits."
        Write-Log "Please commit/stash local changes, then rerun this script."
        return
    }

    Write-Log "Pulling latest code from GitHub..."
    & git pull --ff-only origin $branch 2>&1 | Tee-Object -FilePath $LogFile -Append
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Git pull failed. Continuing with local code."
        return
    }
    Write-Log "GitHub update applied."
}

Write-Log "Started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Update-FromGitHubIfApproved

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 "$WorkDir\automate_kpi.py" --download 2>&1 | Tee-Object -FilePath $LogFile -Append
} else {
    & python "$WorkDir\automate_kpi.py" --download 2>&1 | Tee-Object -FilePath $LogFile -Append
}

$ExitCode = $LASTEXITCODE
Write-Log "Finished at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') with exit code $ExitCode"

exit $ExitCode
