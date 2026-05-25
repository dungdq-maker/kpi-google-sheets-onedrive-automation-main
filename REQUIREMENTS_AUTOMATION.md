# Requirements and Automation Design

## Objective

Automatically consolidate KPI/personnel cost data from department Google Sheets into one normalized staging table.

Current implemented output:

```text
data/output/staging/normalized_output_*.xlsx
data/output/staging/normalized_output_*.csv
```

Future final output:

```text
data/output/final/von_hoa_output_*.xlsx
```

## Scope

The automation currently does:

1. Read source definitions from `config/sources.json`.
2. Download Google Sheets as `.xlsx` files into `data/input/raw`.
3. Parse ACCA/CMA source files.
4. Parse IT matrix source files.
5. Normalize all rows into one staging table.
6. Write `.xlsx` and `.csv` staging output into `data/output/staging`.
7. If `data/input/template/von_hoa_template.xlsx` exists, write final capitalization output with sheet `3. vốn hóa` into `data/output/final`.
8. Write run logs into `logs`.

The automation intentionally skips final output when the real template is missing, because generating a guessed workbook structure can produce incorrect results.

## Environment requirements

- Windows.
- Python 3.10 or newer.
- Python package:
  - `openpyxl`
- OneDrive desktop app signed in if cloud sync is required.
- Google Sheet links must allow export/download.

Install dependency:

```powershell
pip install -r requirements.txt
```

## Source files

| Department | Raw local file | Google Sheet ID |
| --- | --- | --- |
| SX ACCA+CMA | `data/input/raw/CMA.xlsx` | `1jOaBolZ78dbelYkoFL5jSUE0-yJn5425` |
| SX ACCA+CMA | `data/input/raw/ACCA.xlsx` | `16w4-UpSFnjVGpMJlfY9m8LIPTdMZy1dQ` |
| IT | `data/input/raw/IT.xlsx` | `1x9FBjRHISImjCII7GqOcTTn40_BmAnRN` |

Source URLs are configured in:

```text
config/sources.json
```

## Normalized output schema

The staging table has these columns:

- `source_file`
- `source_sheet`
- `source_type`
- `year`
- `month`
- `department`
- `program`
- `position`
- `employee`
- `product_or_project`
- `reference_link`
- `new_or_old`
- `project_name`
- `subject_or_system`
- `product_feature`
- `deliverable`
- `component`
- `complexity`
- `unit`
- `actual_quantity`
- `kpi_standard`
- `total_kpi`

## ACCA/CMA parsing logic

The script:

1. Finds a header row containing employee/product fields.
2. Detects whether the sheet is monthly or employee-month format.
3. For monthly sheets, infers year/month from the sheet name, for example `Apr 26` means April 2026.
4. For employee-month sheets, reads year/month from row values.
5. Skips empty rows.
6. Extracts product, link, subject, product feature, deliverable, component, actual quantity, KPI standard, and total KPI.

## IT parsing logic

The script:

1. Reads the year from the sheet name, for example `2026`.
2. Uses row 13 for project/category/system/month fields.
3. Uses row 14 for employee names across columns.
4. Unpivots each non-zero employee cell into one output row.
5. Extracts position from employee suffix when available, for example `Employee - BA`.

## Run commands

Full workflow:

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

Python with download:

```powershell
python automate_kpi.py --download
```

Python using existing raw files:

```powershell
python automate_kpi.py
```

Custom work directory:

```powershell
python automate_kpi.py --work-dir "C:\Users\admin\OneDrive\Documents\Excel" --download
```

## Scheduled automation

Create/update scheduled task:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Default schedule:

- Day 3 every month.
- 09:00 local time.

## Risks and limitations

- If Google Sheets change headers or layout, parser mapping may need updates.
- If Google Sheet formulas do not export cached values, formula-derived columns may be blank.
- Direct write to the final workbook is not implemented yet.
- When final workbook writing is added, the destination file should be closed before the job runs.

## Required work for final workbook output

To complete the final workbook step:

1. Put the official workbook at `data/input/template/von_hoa_template.xlsx`.
2. Confirm the destination sheet is `3. vốn hóa`.
3. Confirm header names and mapping from staging columns to target columns.
4. Add replace-by-year-month logic to avoid duplicate rows.
5. Save the resulting workbook under `data/output/final`.
