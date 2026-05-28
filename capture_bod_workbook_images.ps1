param(
    [string]$OutputDir = "docs\generated_assets"
)

$ErrorActionPreference = "Stop"

function Get-LatestWorkbook {
    $file = Get-ChildItem -LiteralPath "data\output\final" -Filter "von_hoa_*.xlsx" |
        Where-Object { $_.Name -notlike "~$*" } |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if (-not $file) {
        throw "Cannot find latest output workbook in data\output\final."
    }
    return $file.FullName
}

function Export-ExcelRangePng {
    param(
        [object]$Excel,
        [string]$WorkbookPath,
        [object]$SheetName,
        [string]$RangeAddress,
        [string]$OutputPath
    )

    $resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $resolvedOutput = if ([System.IO.Path]::IsPathRooted($OutputPath)) {
        $OutputPath
    } else {
        Join-Path (Get-Location) $OutputPath
    }
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resolvedOutput) | Out-Null

    $workbook = $null
    $chartObject = $null
    try {
        $workbook = $null
        for ($attempt = 1; $attempt -le 3 -and $workbook -eq $null; $attempt++) {
            try {
                $workbook = $Excel.Workbooks.Open($resolvedWorkbook, 0, $true, 5, "", "", $true)
            }
            catch {
                if ($attempt -eq 3) { throw }
                Start-Sleep -Seconds 2
            }
        }
        if ($workbook -eq $null) {
            throw "Excel could not open workbook: $resolvedWorkbook"
        }
        $sheet = $workbook.Worksheets.Item($SheetName)
        $sheet.Activate() | Out-Null
        $range = $sheet.Range($RangeAddress)

        $range.CopyPicture(1, 2) | Out-Null
        Start-Sleep -Milliseconds 250

        $chartObject = $sheet.ChartObjects().Add($range.Left, $range.Top, $range.Width, $range.Height)
        $chart = $chartObject.Chart
        $chart.Paste() | Out-Null
        Start-Sleep -Milliseconds 250
        $chart.Export($resolvedOutput, "PNG") | Out-Null
    }
    finally {
        if ($chartObject -ne $null) {
            try { $chartObject.Delete() | Out-Null } catch {}
        }
        if ($workbook -ne $null) {
            try { $workbook.Close($false) | Out-Null } catch {}
        }
    }
}

$outputDirPath = if ([System.IO.Path]::IsPathRooted($OutputDir)) {
    $OutputDir
} else {
    Join-Path (Get-Location) $OutputDir
}
New-Item -ItemType Directory -Force -Path $outputDirPath | Out-Null

$latestOutput = Get-LatestWorkbook
$rawDir = Join-Path (Get-Location) "data\input\raw"
$captureStagingDir = Join-Path (Get-Location) "data\output\staging\bod_capture"
New-Item -ItemType Directory -Force -Path $captureStagingDir | Out-Null
$cfaWorkbook = Join-Path $rawDir "CFA.xlsx"
$cfaCleanWorkbook = Join-Path $captureStagingDir "CFA_excel_clean.xlsx"
$cmaWorkbook = Join-Path $rawDir "CMA.xlsx"
$cmaCleanWorkbook = Join-Path $captureStagingDir "CMA_excel_clean.xlsx"
@"
from openpyxl import load_workbook
wb = load_workbook(r'$cfaWorkbook')
wb.save(r'$cfaCleanWorkbook')
wb = load_workbook(r'$cmaWorkbook')
wb.save(r'$cmaCleanWorkbook')
"@ | py -3 -

$items = @(
    # Slide 5 - latest output workbook.
    @{ Workbook = $latestOutput; Sheet = 17; Range = "A1:S37"; Output = "bod_slide5_cost_it.png" },
    @{ Workbook = $latestOutput; Sheet = 4; Range = "A1:M18"; Output = "bod_slide5_von_hoa.png" },
    @{ Workbook = $latestOutput; Sheet = "Data media ACCA+CFA+CMA"; Range = "A1:J18"; Output = "bod_slide5_data_media.png" },

    # Slide 6 - raw input workbooks.
    @{ Workbook = (Join-Path $rawDir "ACCA.xlsx"); Sheet = 5; Range = "A1:Q35"; Output = "sx_source_acca.png" },
    @{ Workbook = $cfaCleanWorkbook; Sheet = 9; Range = "A1:Q35"; Output = "sx_source_cfa.png" },
    @{ Workbook = $cmaCleanWorkbook; Sheet = 12; Range = "A1:I35"; Output = "sx_source_cma.png" },

    # Slide 7 - checkpoint / approval.
    @{ Workbook = $latestOutput; Sheet = "Huong_dan_Approval"; Range = "A1:E18"; Output = "approval_guide.png" },
    @{ Workbook = $latestOutput; Sheet = "Check_IT_CPNS"; Range = "A1:O25"; Output = "check_it_cpns.png" },
    @{ Workbook = $latestOutput; Sheet = "Check_Media_Timesheet"; Range = "A1:L24"; Output = "check_media_timesheet.png" },

    # Slide 9 - SX checkpoint.
    @{ Workbook = $latestOutput; Sheet = "checkpoint data SX"; Range = "A1:J24"; Output = "sx_checkpoint.png" },
    @{ Workbook = $latestOutput; Sheet = "Check_SX_Downstream"; Range = "A1:T24"; Output = "sx_downstream.png" },

    # Slide 10 - result sheets.
    @{ Workbook = $latestOutput; Sheet = "IT_Approval_Result"; Range = "A1:I20"; Output = "it_approval_result.png" },
    @{ Workbook = $latestOutput; Sheet = "Project_Master_Approval_Result"; Range = "A1:H28"; Output = "project_master_approval_result.png" },
    @{ Workbook = $latestOutput; Sheet = "Media_Approval_Result"; Range = "A1:I22"; Output = "media_approval_result.png" },

    # Slide 12 - SX result.
    @{ Workbook = $latestOutput; Sheet = "SX_Approval_Result"; Range = "A1:I22"; Output = "sx_approval_result.png" }
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    foreach ($item in $items) {
        $outPath = Join-Path $outputDirPath $item.Output
        Export-ExcelRangePng -Excel $excel -WorkbookPath $item.Workbook -SheetName $item.Sheet -RangeAddress $item.Range -OutputPath $outPath
        Write-Output "Exported $($item.Output) from $($item.Sheet)"
    }
}
finally {
    try { $excel.Quit() | Out-Null } catch {}
}

Write-Output "Latest output workbook: $latestOutput"
