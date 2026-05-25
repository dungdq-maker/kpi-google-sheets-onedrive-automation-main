# Usage Guide

This guide explains how to run the automation from Google Sheet links to staging output.

## 1. Check input config

Open:

```text
config/sources.json
```

Each input source must have:

- `url`: Google Sheet URL.
- `file`: exported file name under `data/input/raw`.
- `department`: department label in output.

Example:

```json
{
  "IT": {
    "url": "https://docs.google.com/spreadsheets/d/1x9FBjRHISImjCII7GqOcTTn40_BmAnRN/edit",
    "file": "IT.xlsx",
    "department": "IT"
  }
}
```

## 2. Install dependency

Run once:

```powershell
cd "C:\Users\admin\OneDrive\Documents\Excel"
pip install -r requirements.txt
```

## 3. Run full workflow

Use this command for normal usage:

```powershell
powershell -ExecutionPolicy Bypass -File ".\run_kpi_automation.ps1"
```

What it does:

1. Reads Google Sheet links from `config/sources.json`.
2. Downloads raw Excel files into `data/input/raw`.
3. Normalizes ACCA/CMA and IT data.
4. Writes staging output into `data/output/staging`.
5. Writes final capitalization output into `data/output/final`.
6. Writes run log into `logs`.

Final output requires the real template file:

```text
data/input/template/von_hoa_template.xlsx
```

If this file is missing, the workflow intentionally skips final output and only creates staging output.

## 4. Run by Python

Download latest Google Sheets and create output:

```powershell
python automate_kpi.py --download
```

Use already downloaded files in `data/input/raw`:

```powershell
python automate_kpi.py
```

Run from another folder:

```powershell
python automate_kpi.py --work-dir "C:\Users\admin\OneDrive\Documents\Excel" --download
```

## 5. Locate files after running

Raw input:

```text
data/input/raw/CMA.xlsx
data/input/raw/ACCA.xlsx
data/input/raw/IT.xlsx
```

Staging output:

```text
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.xlsx
data/output/staging/normalized_output_YYYYMMDD_HHMMSS.csv
```

Run logs:

```text
logs/kpi_automation_YYYYMMDD_HHMMSS.log
```

Final capitalization output:

```text
data/output/final/von_hoa_output_YYYYMMDD_HHMMSS.xlsx
```

## 6. Validate output

For normalized review, open the latest `.xlsx` file in:

```text
data/output/staging/
```

Check these columns first:

- `year`
- `month`
- `department`
- `employee`
- `product_or_project`
- `actual_quantity`
- `total_kpi`

For final capitalization review, open the latest file in:

```text
data/output/final/
```

It should contain a sheet named `3. vốn hóa`.

If no file appears in `data/output/final`, check the log for:

```text
FINAL SKIPPED
```

Then place the real template workbook at `data/input/template/von_hoa_template.xlsx` and run again.

## 7. Schedule automatic run

Create or update the scheduled task:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

Default task:

```text
KPI GoogleSheet To OneDrive Automation
```

Default schedule:

- Day 3 every month.
- 09:00 local time.

## 8. Change schedule

Open:

```text
setup_kpi_scheduled_task.ps1
```

Edit:

```powershell
/D 3
/ST 09:00
```

Then run setup again:

```powershell
powershell -ExecutionPolicy Bypass -File ".\setup_kpi_scheduled_task.ps1"
```

## 9. Clean generated outputs

To remove generated output and logs, delete files under:

```text
data/output/staging/
data/output/final/
logs/
```

Keep `.gitkeep` files if present.

To move old generated files left in the project root:

```powershell
powershell -ExecutionPolicy Bypass -File ".\cleanup_legacy_root.ps1"
```

If a file cannot be moved, close Excel or wait for OneDrive to release the file, then run the cleanup again.

## 10. Troubleshooting

If Google Sheet download fails:

- Check that the URL in `config/sources.json` is correct.
- Check that the sheet allows export/download.
- Check internet connection.

If a source file is locked:

```text
Warning: ACCA.xlsx is locked. Using downloaded temporary copy.
```

This is acceptable. The script will continue with the temporary downloaded file.

If output is missing:

- Check the latest file in `logs`.
- Run the workflow again with `--download`.
