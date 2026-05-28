param(
    [string]$FinalWorkbook,
    [string]$PayrollWorkbook,
    [string]$OutputDir = "docs\generated_assets",
    [int]$MaxRows = 28,
    [int]$MaxCols = 12,
    [switch]$OnlySxInput
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Resolve-LatestWorkbook {
    param([string]$Folder)
    $latest = Get-ChildItem -Path $Folder -Filter "*.xlsx" -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $latest) {
        throw "Khong tim thay workbook .xlsx trong $Folder"
    }
    return $latest.FullName
}

function Resolve-PayrollWorkbook {
    $folder = Join-Path $env:USERPROFILE "OneDrive\BCQT 2026"
    if (-not (Test-Path -LiteralPath $folder)) {
        return $null
    }
    $candidates = Get-ChildItem -Path $folder -Filter "*.xlsx" -File |
        Where-Object { $_.Name -like "*2026.xlsx" -and $_.Name -notlike "~$*" } |
        Sort-Object LastWriteTime -Descending
    foreach ($candidate in $candidates) {
        if ($candidate.Name -like "*V*n h*a*" -or $candidate.Name -like "*nh*n s*") {
            return $candidate.FullName
        }
    }
    return ($candidates | Select-Object -First 1).FullName
}

function Item-ValueOrDefault {
    param(
        [hashtable]$Item,
        [string]$Key,
        [object]$Default
    )
    if ($Item.ContainsKey($Key)) {
        return $Item[$Key]
    }
    return $Default
}

function Export-SheetPicture {
    param(
        [object]$Excel,
        [string]$WorkbookPath,
        [string]$SheetName,
        [string]$OutputPath,
        [int]$Rows,
        [int]$Cols,
        [int]$StartRow = 1,
        [int]$StartCol = 1
    )

    if (-not (Test-Path -LiteralPath $WorkbookPath)) {
        Write-Warning "Bo qua vi khong tim thay workbook: $WorkbookPath"
        return
    }

    $workbook = $null
    try {
        $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
        $workbook = $Excel.Workbooks.Open($resolvedWorkbookPath, 0, $true, 5, "", "", $true, 1, "", $false, $false, 0, $false, $true, 1)
        $sheet = $null
        foreach ($candidate in $workbook.Worksheets) {
            if ($candidate.Name -eq $SheetName) {
                $sheet = $candidate
                break
            }
        }
        if (-not $sheet -and $SheetName.Contains("*")) {
            foreach ($candidate in $workbook.Worksheets) {
                if ($candidate.Name -like $SheetName) {
                    $sheet = $candidate
                    break
                }
            }
        }

        if (-not $sheet) {
            Write-Warning "Bo qua vi khong co sheet '$SheetName' trong $WorkbookPath"
            return
        }

        $used = $sheet.UsedRange
        $lastRow = [Math]::Min([int]$used.Row + [int]$used.Rows.Count - 1, $StartRow + $Rows - 1)
        $lastCol = [Math]::Min([int]$used.Column + [int]$used.Columns.Count - 1, $StartCol + $Cols - 1)
        if ($lastRow -lt $StartRow) { $lastRow = $StartRow }
        if ($lastCol -lt $StartCol) { $lastCol = $StartCol }

        $range = $sheet.Range($sheet.Cells($StartRow, $StartCol), $sheet.Cells($lastRow, $lastCol))
        $sheet.Activate() | Out-Null
        $range.Select() | Out-Null
        Start-Sleep -Milliseconds 150
        [System.Windows.Forms.Clipboard]::Clear()
        $copied = $false
        for ($attempt = 1; $attempt -le 3 -and -not $copied; $attempt++) {
            try {
                $range.CopyPicture(1, 2) | Out-Null
                $copied = $true
            }
            catch {
                if ($attempt -eq 3) {
                    throw
                }
                Start-Sleep -Milliseconds 700
            }
        }
        Start-Sleep -Milliseconds 500

        $image = $null
        for ($i = 0; $i -lt 10 -and -not $image; $i++) {
            $image = [System.Windows.Forms.Clipboard]::GetImage()
            if (-not $image) {
                Start-Sleep -Milliseconds 250
            }
        }
        if (-not $image) {
            throw "Khong lay duoc anh tu clipboard sau CopyPicture cho sheet '$SheetName'"
        }
        $image.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $image.Dispose()
        Write-Output "Exported $SheetName -> $OutputPath"
    }
    finally {
        if ($workbook) {
            try {
                $workbook.Close($false) | Out-Null
            }
            catch {
                Write-Warning "Khong dong duoc workbook '$WorkbookPath': $($_.Exception.Message)"
            }
        }
    }
}

if (-not $FinalWorkbook) {
    $FinalWorkbook = Resolve-LatestWorkbook -Folder "data\output\final"
}
if (-not $PayrollWorkbook) {
    $PayrollWorkbook = Resolve-PayrollWorkbook
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$resolvedOutputDir = (Resolve-Path -LiteralPath $OutputDir).Path

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $true
try {
    # Minimize the automation window, while keeping Excel visible enough for CopyPicture.
    $excel.WindowState = -4140
}
catch {
    Write-Warning "Khong the minimize Excel automation window: $($_.Exception.Message)"
}

try {
    $finalSheets = @(
        @{ Sheet = "Timesheet IT"; File = "timesheet_it.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Timesheet Media"; File = "timesheet_media.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Data media ACCA+CFA+CMA"; File = "data_media_sheet.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Data SX ACCA+CMA"; File = "data_sx_acca_cma.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Data SX CFA"; File = "data_sx_cfa.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Data SX+Media SC"; File = "data_sx_media_sc.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Timesheet SX"; File = "timesheet_sx.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "SX_Allocation_Build"; File = "sx_allocation_build.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Huong_dan_Approval"; File = "approval_guide.png"; Rows = 22; Cols = 10 },
        @{ Sheet = "Check_IT_CPNS"; File = "check_it_cpns.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Check_Media_Timesheet"; File = "check_media_timesheet.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Check_Payroll"; File = "check_payroll.png"; Rows = 24; Cols = 12 },
        @{ Sheet = "checkpoint data SX"; File = "sx_checkpoint.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Check_SX_Downstream"; File = "sx_downstream.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "IT_Approval_Result"; File = "it_approval_result.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Project_Master_Approval_Result"; File = "project_master_approval_result.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "Media_Approval_Result"; File = "media_approval_result.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "SX_Approval_Result"; File = "sx_approval_result.png"; Rows = 26; Cols = 12 },
        @{ Sheet = "SX_Downstream_Approval_Result"; File = "sx_downstream_approval_result.png"; Rows = 26; Cols = 12 }
    )

    if (-not $OnlySxInput) {
        foreach ($item in $finalSheets) {
            try {
                Export-SheetPicture `
                    -Excel $excel `
                    -WorkbookPath $FinalWorkbook `
                    -SheetName $item.Sheet `
                    -OutputPath (Join-Path $resolvedOutputDir $item.File) `
                    -Rows $item.Rows `
                    -Cols $item.Cols `
                    -StartRow (Item-ValueOrDefault -Item $item -Key "StartRow" -Default 1) `
                    -StartCol (Item-ValueOrDefault -Item $item -Key "StartCol" -Default 1)
            }
            catch {
                Write-Warning "Export loi sheet '$($item.Sheet)': $($_.Exception.Message)"
            }
        }

        $payrollSheets = @(
            @{ Sheet = "*full*time*"; File = "payroll_ft.png"; Rows = 26; Cols = 12 },
            @{ Sheet = "*part*time*"; File = "payroll_pt.png"; Rows = 26; Cols = 12 }
        )

        foreach ($item in $payrollSheets) {
            try {
                Export-SheetPicture `
                    -Excel $excel `
                    -WorkbookPath $PayrollWorkbook `
                    -SheetName $item.Sheet `
                    -OutputPath (Join-Path $resolvedOutputDir $item.File) `
                    -Rows $item.Rows `
                    -Cols $item.Cols `
                    -StartRow (Item-ValueOrDefault -Item $item -Key "StartRow" -Default 1) `
                    -StartCol (Item-ValueOrDefault -Item $item -Key "StartCol" -Default 1)
            }
            catch {
                Write-Warning "Export loi sheet '$($item.Sheet)': $($_.Exception.Message)"
            }
        }
    }

    $rawDir = "data\input\raw"
    $sxSourceSheets = @(
        @{ Workbook = (Join-Path $rawDir "ACCA.xlsx"); Sheet = "0426"; File = "sx_source_acca.png"; Rows = 26; Cols = 14; StartRow = 16; StartCol = 1 },
        @{ Workbook = (Join-Path $rawDir "CFA.xlsx"); Sheet = "Apr 26"; File = "sx_source_cfa.png"; Rows = 26; Cols = 14; StartRow = 16; StartCol = 1 },
        @{ Workbook = (Join-Path $rawDir "CMA.xlsx"); Sheet = "*Sup"; File = "sx_source_cma.png"; Rows = 26; Cols = 14; StartRow = 16; StartCol = 1 }
    )

    foreach ($item in $sxSourceSheets) {
        try {
            Export-SheetPicture `
                -Excel $excel `
                -WorkbookPath $item.Workbook `
                -SheetName $item.Sheet `
                -OutputPath (Join-Path $resolvedOutputDir $item.File) `
                -Rows $item.Rows `
                -Cols $item.Cols `
                -StartRow (Item-ValueOrDefault -Item $item -Key "StartRow" -Default 1) `
                -StartCol (Item-ValueOrDefault -Item $item -Key "StartCol" -Default 1)
        }
        catch {
            Write-Warning "Export loi SX input '$($item.Workbook)' / '$($item.Sheet)': $($_.Exception.Message)"
        }
    }
}
finally {
    try {
        $excel.Quit() | Out-Null
    }
    catch {
        Write-Warning "Excel COM da dong hoac mat ket noi truoc khi Quit: $($_.Exception.Message)"
    }
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    catch {
        Write-Warning "Khong release duoc Excel COM object: $($_.Exception.Message)"
    }
}
