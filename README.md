# KPI Google Sheets to OneDrive Automation

Automation nay tai 2 nguon du lieu tu Google Sheets:

- `IT`
- `media`

Ket qua duoc ghi vao workbook template `von hoa chi phi nhan su` trong `data/output/final`.

## Current Scope

Phiên ban hien tai chi xu ly:

- `Timesheet IT`
- `Data media ACCA+CFA+CMA`

Nhom SX/KPI cu da duoc loai bo khoi pipeline.

## Input

Input duoc cau hinh trong:

```text
config/sources.json
```

Moi source co:

- `url`: Google Sheet URL
- `file`: ten file xlsx luu trong `data/input/raw`
- `type`: `matrix` cho IT, `task_table` cho media

Khi chay voi `--download`, script se doi Google Sheet URL sang link export `.xlsx` va luu file vao:

```text
data/input/raw/IT.xlsx
data/input/raw/MEDIA.xlsx
```

## Processing

### IT

Nguon IT duoc xu ly nhu matrix va copy sang sheet:

```text
Timesheet IT
```

### Media

Nguon media duoc xu ly nhu bang task va copy sang sheet:

```text
Data media ACCA+CFA+CMA
```

Script se lay cac cot co y nghia nhu:

- `BU`
- `Ten du an SX` / `Ten du an`
- `Ten nhiem vu`
- `So gio hoan thanh cong viec` / `Thoi gian leader phan bo`
- `Thang` / `Ngay tao`
- `Nguoi lam task`
- `Ma nhan vien`

Neu `Ma nhan vien` bi thieu, script se thu lookup trong sheet `Mã nhân viên` cua template.

## Output

File output duoc tao tai:

```text
data/output/final/von_hoa_YYYYMMDD_HHMMSS.xlsx
```

Neu muon dung URL moi trong ngay, co the override truc tiep:

```powershell
python automate_kpi.py --download --it-url "..." --media-url "..."
```

## Run

```powershell
python automate_kpi.py --download
python automate_kpi.py
```

## PDF guide

To generate the Vietnamese step-by-step PDF guide with diagrams, run:

```powershell
py -3 generate_huong_dan_pdf.py
```

The file is written to:

```text
docs/huong_dan_kpi_automation.pdf
```

To generate the operator guide for maintaining code and GitHub publishing, run:

```powershell
py -3 generate_operator_guide_pdf.py
```

The file is written to:

```text
docs/huong_dan_van_hanh_code_github.pdf
```

Neu `python` khong co san tren PowerShell, dung:

```powershell
py -3 automate_kpi.py --download
```
