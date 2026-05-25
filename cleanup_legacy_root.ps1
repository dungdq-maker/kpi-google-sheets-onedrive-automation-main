$ErrorActionPreference = "Continue"

$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$Archive = Join-Path $Workspace "data\archive\legacy-root"
New-Item -ItemType Directory -Force -Path $Archive | Out-Null

$Patterns = @(
    "ACCA.xlsx",
    "CMA.xlsx",
    "IT.xlsx",
    ".ACCA.download.xlsx",
    "normalized_output_*.csv",
    "normalized_output_*.xlsx"
)

foreach ($Pattern in $Patterns) {
    Get-ChildItem -Path $Workspace -Filter $Pattern -File -Force | ForEach-Object {
        $Target = Join-Path $Archive $_.Name
        try {
            Move-Item -LiteralPath $_.FullName -Destination $Target -Force
            Write-Host "Moved: $($_.Name) -> data\archive\legacy-root"
        }
        catch {
            Write-Warning "Could not move $($_.Name). Close Excel/OneDrive locks and run cleanup again."
        }
    }
}
