$ErrorActionPreference = "Stop"

$TaskName = "KPI GoogleSheet To OneDrive Automation"
$WorkDir = "C:\Users\admin\OneDrive\Documents\Excel"
$Runner = Join-Path $WorkDir "run_kpi_automation.ps1"

if (-not (Test-Path $Runner)) {
    throw "Runner script not found: $Runner"
}

$TaskCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$Runner`""

# Default: run on day 3 of every month at 09:00 local time.
schtasks.exe /Create `
    /TN $TaskName `
    /TR $TaskCommand `
    /SC MONTHLY `
    /D 3 `
    /ST 09:00 `
    /F | Out-Null

Write-Host "Scheduled task created/updated: $TaskName"
Write-Host "Schedule: day 3 of every month at 09:00"
Write-Host "Runner: $Runner"
