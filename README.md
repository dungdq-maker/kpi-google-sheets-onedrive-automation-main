# KPI Google Sheets to OneDrive Automation

Automation này đồng bộ dữ liệu từ Google Sheets và file approval local vào workbook template
`Vốn hóa chi phí nhân sự 2026.xlsx`.

Pipeline hiện tại bao gồm 3 luồng chính:

- `IT`
- `Media`
- `SX`

Kết quả cuối cùng được ghi vào `data/output/final`.

## Current Scope

Phiên bản hiện tại xử lý:

- `Timesheet IT`
- `Data media ACCA+CFA+CMA`
- `Data SX ACCA+CMA`
- `SX_Allocation_Build`
- `Timesheet SX`
- `4.1 Chi phí nhân sự SX`
- `3.Vốn hóa`

`Data SX consol` đã được loại khỏi flow SX.  
Luồng SX hiện đi theo:

`Data SX ACCA+CMA -> SX_Allocation_Build -> Timesheet SX -> 4.1 Chi phí nhân sự SX -> 3.Vốn hóa`

## Input

Input được cấu hình trong:

```text
config/sources.json
```

Mỗi source có:

- `url`: Google Sheet URL
- `file`: tên file xlsx lưu trong `data/input/raw`
- `type`: kiểu source tương ứng

Khi chạy với `--download`, script sẽ tải source từ Google Sheets và lưu file vào:

```text
data/input/raw/
```

## Processing

### IT

Nguồn IT được xử lý như matrix và đổ vào:

```text
Timesheet IT
```

Pipeline IT vẫn giữ cơ chế approval/carry-forward cho các sheet liên quan đến mapping và downstream.

### Media

Nguồn media được xử lý như task table và đổ vào:

```text
Data media ACCA+CFA+CMA
```

Các cột thường dùng gồm:

- `BU`
- `Tên dự án`
- `Tên nhiệm vụ`
- `Số giờ hoàn thành công việc`
- `Tháng` / `Ngày tạo`
- `Người làm task`
- `Mã nhân viên`

Nếu `Mã nhân viên` bị thiếu, script sẽ lookup trong sheet `Mã nhân viên` của template.

### SX

Nguồn SX được gộp từ 3 file nguồn và xuất ra staging trước khi đổ vào template:

- `ACCA`
- `CMA`
- `CFA`

Luồng SX hiện tại:

1. Build staging `Data SX ACCA+CMA` từ các nguồn SX.
2. Tạo bảng trung gian `SX_Allocation_Build`.
3. Rebuild công thức `Timesheet SX` từ `SX_Allocation_Build`.
4. Downstream sang `4.1 Chi phí nhân sự SX`.
5. Downstream sang `3.Vốn hóa`.

Các checkpoint SX chính:

- `checkpoint data SX`
- `Check_SX_Downstream`
- `Check_Vonhoa_Month_Block`

## Approval Flow

Khi chạy với `--approval-file`, workbook approved sẽ được dùng để carry forward
những sheet người dùng đã duyệt/chỉnh sửa.

Các sheet approval/checkpoint chính:

- `Check_Payroll`
- `checkpoint data SX`
- `Check_SX_Downstream`
- `Check_Vonhoa_Month_Block`
- `Check_IT_CPNS`
- `IT_New_Project_Master`
- `Check_IT_Downstream`
- `Check_Media_Timesheet`

Các sheet kết quả approval được build lại theo lần chạy hiện tại.

## Output

File output được tạo tại:

```text
data/output/final/von_hoa_YYYYMMDD_HHMMSS.xlsx
```

Ví dụ:

```powershell
python automate_kpi.py --download --sx-year 2026 --sx-month 5
python automate_kpi.py --approval-file "data/output/final/von_hoa_20260526_162209.xlsx" --sx-year 2026 --sx-month 5
```

## Run

```powershell
python automate_kpi.py --download
python automate_kpi.py
python automate_kpi.py --sx-year 2026 --sx-month 5
python automate_kpi.py --approval-file "data/output/final/von_hoa_20260526_162209.xlsx" --sx-year 2026 --sx-month 5
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

To generate the BOD / manager presentation script with speaker notes, run:

```powershell
py -3 generate_thuyet_trinh_bod_manager.py
```

The file is written to:

```text
docs/thuyet_trinh_bod_manager.pdf
```

To generate the browser presentation deck for screen sharing, run:

```powershell
py -3 generate_bod_presentation_html.py
```

The file is written to:

```text
docs/thuyet_trinh_bod_manager.html
```

Nếu `python` không có sẵn trên PowerShell, dùng:

```powershell
py -3 automate_kpi.py --download
```
