from __future__ import annotations

import argparse
import csv
import io
import json
import re
import shutil
import sys
import urllib.request
import unicodedata
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

WORK_DIR = Path(__file__).parent

DIRS = {
    "raw": WORK_DIR / "data" / "input" / "raw",
    "template": WORK_DIR / "data" / "input" / "template",
    "staging": WORK_DIR / "data" / "output" / "staging",
    "final": WORK_DIR / "data" / "output" / "final",
    "logs": WORK_DIR / "logs",
}

DEFAULT_SOURCES = {
    "IT": {
        "url": "https://docs.google.com/spreadsheets/d/1xk2Bd95uXoOJAeBqkOSzV1wEo9peOF-e/edit?usp=sharing&ouid=111275165352399420864&rtpof=true&sd=true",
        "file": "IT.xlsx",
        "type": "matrix",
    },
    "MEDIA": {
        "url": "https://docs.google.com/spreadsheets/d/17Eaj28S4nRgR1Iy6yKzkB7wo9XzS68aP/edit?gid=646344289#gid=646344289",
        "file": "MEDIA.xlsx",
        "type": "task_table",
    },
}

TEMPLATE_CANDIDATES = [
    "von_hoa_template.xlsx",
    "Vốn hóa chi phí nhân sự 2026.xlsx",
    "Vốn hóa chi phí nhân sự 2026 (1).xlsx",
    "Vốn hóa chi phí nhân sự 2026 (2).xlsx",
]

SHEET_IT = "Timesheet IT"
SHEET_MEDIA = "Data media ACCA+CFA+CMA"
SHEET_TIMESHEET_MEDIA = "Timesheet Media"
SHEET_IT_COST = "Chi phí nhân sự IT"
SHEET_SALARY_FULLTIME = "Lương nhân viên full time"
SHEET_SALARY_PARTTIME = "Lương nhân viên part time"
SHEET_PROJECT_CATALOG = "1.Danh mục dự án"
SHEET_IT_CHECKING = "Checking Vốn hóa IT"
SHEET_CAPITALIZATION = "3.Vốn hóa"
SHEET_EMPLOYEE = "Mã nhân viên"

PAYROLL_SOURCE_PATH = Path(r"C:\Users\admin\OneDrive\BCQT 2026\Vốn hóa chi phí nhân sự 2026.xlsx")
PAYROLL_SHEETS = [SHEET_SALARY_FULLTIME, SHEET_SALARY_PARTTIME]

IT_HEADER_ROW = 13
IT_TEMPLATE_ROW = 13
MEDIA_HEADER_ROW = 5
MEDIA_TEMPLATE_ROW = 6
MEDIA_AUDIT_FIELDS = [
    "source_sheet",
    "source_row",
    "bu",
    "project_name",
    "task_name",
    "hours",
    "month",
    "employee",
    "mnv",
]
APPROVAL_YES_VALUES = {"yes", "y", "true", "1", "apply", "approved"}


if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def clean(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        v = re.sub(r"\s+", " ", value).strip()
        return v or None
    return value


def norm(value: Any) -> str:
    value = clean(value)
    if value is None:
        return ""
    text = str(value).lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_num(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def setup_dirs() -> None:
    for path in DIRS.values():
        path.mkdir(parents=True, exist_ok=True)


def load_source_config() -> dict[str, dict[str, Any]]:
    cfg_path = WORK_DIR / "config" / "sources.json"
    if not cfg_path.exists():
        return DEFAULT_SOURCES.copy()

    with cfg_path.open("r", encoding="utf-8") as f:
        loaded = json.load(f)

    result: dict[str, dict[str, Any]] = {}
    for key, fallback in DEFAULT_SOURCES.items():
        item = dict(fallback)
        if key in loaded and isinstance(loaded[key], dict):
            item.update({k: v for k, v in loaded[key].items() if v is not None})
        result[key] = item
    return result


def find_template() -> Path | None:
    for name in TEMPLATE_CANDIDATES:
        p = DIRS["template"] / name
        if p.exists():
            return p

    candidates: list[Path] = []
    for folder in [DIRS["template"], WORK_DIR]:
        if folder.exists():
            candidates.extend(folder.glob("*.xlsx"))

    for p in candidates:
        low = p.name.lower()
        if "von_hoa" in low or "vốn hóa" in low or "von hoa" in low:
            return p
    return None


def extract_spreadsheet_id(url: str) -> str | None:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    return None


def extract_gid(url: str) -> str | None:
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    if "gid" in query and query["gid"]:
        return query["gid"][0]
    frag = parsed.fragment
    if frag.startswith("gid="):
        return frag.split("=", 1)[1]
    return None


def build_export_url(url: str, full_workbook: bool = True) -> str:
    sheet_id = extract_spreadsheet_id(url)
    if not sheet_id:
        return url

    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    gid = extract_gid(url)
    if gid and not full_workbook:
        export_url += f"&gid={gid}"
    return export_url


def download_source(url: str, out_path: Path, full_workbook: bool = True) -> None:
    export_url = build_export_url(url, full_workbook=full_workbook)
    tmp = out_path.with_name(f".{out_path.stem}.download.xlsx")
    req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})

    try:
        with urllib.request.urlopen(req) as resp, tmp.open("wb") as f:
            shutil.copyfileobj(resp, f)

        try:
            tmp.replace(out_path)
        except PermissionError:
            shutil.copy2(tmp, out_path)
            tmp.unlink(missing_ok=True)
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        if out_path.exists():
            print(f"    download skipped, using existing local file: {exc}")
            return
        raise


def download_all(sources: dict[str, dict[str, Any]]) -> None:
    for key, meta in sources.items():
        out = DIRS["raw"] / meta["file"]
        print(f"  {key} -> {out.name}")
        try:
            download_source(meta["url"], out, full_workbook=True)
        except Exception as exc:
            print(f"    FAILED: {exc}")


def ensure_row_style(ws, template_row: int, target_row: int, max_col: int | None = None) -> None:
    max_col = max_col or ws.max_column
    if target_row <= template_row:
        return

    ws.row_dimensions[target_row].height = ws.row_dimensions[template_row].height
    for col in range(1, max_col + 1):
        src = ws.cell(template_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.font:
            dst.font = copy(src.font)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.border:
            dst.border = copy(src.border)
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)


def clear_values(ws, start_row: int, end_row: int, start_col: int, end_col: int) -> None:
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(r, c).value = None


def find_header_row(ws, required_terms: list[str], max_scan_row: int = 20) -> tuple[int | None, dict[str, int]]:
    for row_idx in range(1, min(ws.max_row or 0, max_scan_row) + 1):
        headers: dict[str, int] = {}
        for col in range(1, (ws.max_column or 0) + 1):
            h = norm(ws.cell(row_idx, col).value)
            if h:
                headers[h] = col
        if sum(1 for term in required_terms if any(term in h for h in headers)) >= max(2, len(required_terms) // 2):
            return row_idx, headers
    return None, {}


def get_header_col(headers: dict[str, int], *aliases: str) -> int | None:
    for alias in aliases:
        alias_n = norm(alias)
        if alias_n in headers:
            return headers[alias_n]
        for header, col in headers.items():
            if alias_n in header:
                return col
    return None


def cell_value(ws, row: int, headers: dict[str, int], *aliases: str) -> Any:
    col = get_header_col(headers, *aliases)
    if col is None:
        return None
    return clean(ws.cell(row, col).value)


def parse_month_value(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.month
    if is_num(value):
        n = int(float(value))
        return n if 1 <= n <= 12 else None
    text = str(value).strip()
    if not text:
        return None
    compact_prefix = re.match(r"^(0[1-9]|1[0-2])\d{2}(?:\D|$)", text)
    if compact_prefix:
        return int(compact_prefix.group(1))
    m = re.search(r"\b(1[0-2]|0?[1-9])\b", text)
    if m:
        return int(m.group(1))
    for name, num in [
        ("jan", 1), ("feb", 2), ("mar", 3), ("apr", 4), ("may", 5), ("jun", 6),
        ("jul", 7), ("aug", 8), ("sep", 9), ("oct", 10), ("nov", 11), ("dec", 12),
    ]:
        if name in text.lower():
            return num
    return None


def parse_hours(value: Any) -> float | None:
    if value is None:
        return None
    if is_num(value):
        return float(value)
    text = str(value).strip().replace(",", ".")
    if not text:
        return None
    m = re.search(r"(\d+(?:\.\d+)?)", text)
    if m:
        return float(m.group(1))
    return None


def build_employee_lookup(template_wb) -> dict[str, str]:
    if SHEET_EMPLOYEE not in template_wb.sheetnames:
        return {}

    ws = template_wb[SHEET_EMPLOYEE]
    headers = {}
    for c in range(1, ws.max_column + 1):
        h = norm(ws.cell(1, c).value)
        if h:
            headers[h] = c

    name_col = get_header_col(headers, "ten nhan vien", "ho ten", "employee")
    code_col = get_header_col(headers, "ma nhan vien", "mnv", "employee code")
    if name_col is None or code_col is None:
        return {}

    lookup: dict[str, str] = {}
    for r in range(2, ws.max_row + 1):
        name = clean(ws.cell(r, name_col).value)
        code = clean(ws.cell(r, code_col).value)
        if name and code:
            lookup[norm(name)] = str(code)
    return lookup


def lookup_employee_code(lookup: dict[str, str], name: Any) -> str | None:
    if not name:
        return None
    return lookup.get(norm(name))


def find_it_sheet(wb):
    for ws in wb.worksheets:
        if str(ws.title).isdigit():
            return ws
    return wb.worksheets[0]


def read_it_sheet(path: Path) -> tuple[str, list[list[Any]]]:
    wb = load_workbook(path, data_only=True)
    ws = find_it_sheet(wb)
    rows = [
        [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        for r in range(1, ws.max_row + 1)
    ]
    return ws.title, rows


def write_it_sheet(ws, rows: list[list[Any]], source_path: Path) -> None:
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= IT_HEADER_ROW:
            ws.unmerge_cells(str(mc))

    for r in range(IT_HEADER_ROW, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    needed_max_row = IT_HEADER_ROW + max(0, len(rows) - IT_HEADER_ROW)
    if needed_max_row > ws.max_row:
        for r in range(ws.max_row + 1, needed_max_row + 1):
            ensure_row_style(ws, IT_TEMPLATE_ROW, r, ws.max_column)

    for offset, row in enumerate(rows[IT_HEADER_ROW - 1 :]):
        r = IT_HEADER_ROW + offset
        for c, value in enumerate(row, start=1):
            ws.cell(r, c).value = clean(value)

    # Fill down cột "Tên dự án" theo từng block merge gốc của source IT.
    # Nếu chỉ copy nguyên trạng, các dòng tháng 2..12 sẽ bị blank ở cột B
    # và SUMIFS phía dưới không khớp criteria project name.
    src_wb = load_workbook(source_path, data_only=False)
    src_ws = find_it_sheet(src_wb)
    for mc in src_ws.merged_cells.ranges:
        if mc.min_col != 2 or mc.max_col != 2:
            continue
        if mc.min_row < IT_HEADER_ROW + 2:
            continue
        project_name = clean(src_ws.cell(mc.min_row, 2).value)
        if project_name is None:
            continue
        top_style = ws.cell(mc.min_row, 2)
        for r in range(mc.min_row, mc.max_row + 1):
            cell = ws.cell(r, 2)
            cell.value = project_name
            if r != mc.min_row and top_style.has_style:
                cell._style = copy(top_style._style)
                cell.number_format = top_style.number_format
                cell.font = copy(top_style.font)
                cell.fill = copy(top_style.fill)
                cell.border = copy(top_style.border)
                cell.alignment = copy(top_style.alignment)
                cell.protection = copy(top_style.protection)


def sheet_bu_hint(sheet_name: str) -> str | None:
    name = norm(sheet_name)
    if "cfa" in name:
        return "CFA"
    if "cma" in name:
        return "CMA"
    if "acca" in name:
        return "ACCA"
    return None


def extract_media_rows(path: Path, employee_lookup: dict[str, str]) -> list[dict[str, Any]]:
    wb = load_workbook(path, data_only=True)

    rows: list[dict[str, Any]] = []
    required = ["ngay tao", "ten nhiem vu", "nguoi lam task"]

    for ws in wb.worksheets:
        header_row, headers = find_header_row(ws, required_terms=required, max_scan_row=8)
        if header_row is None:
            continue

        bu_from_sheet = sheet_bu_hint(ws.title)
        for r in range(header_row + 1, ws.max_row + 1):
            raw = {
                "bu": cell_value(ws, r, headers, "bu", "phong ban giao nhiem vu", "bo phan") or bu_from_sheet,
                "project_name": cell_value(
                    ws, r, headers,
                    "ten du an sx",
                    "ten du an",
                    "ten du an sx ( diền theo list ở sheet \"ten du an\")",
                    "ten du an sx ( diền theo list ở sheet 'ten du an')",
                ),
                "task_name": cell_value(ws, r, headers, "ten nhiem vu", "nhiem vu", "ten task", "task"),
                "hours": parse_hours(
                    cell_value(
                        ws,
                        r,
                        headers,
                        "so gio hoan thanh cong viec",
                        "thoi gian can de hoan thanh cong viec",
                        "so gio lam viec thuc te",
                        "thoi gian leader phan bo",
                        "thoi gian hoan thanh task",
                        "thoi gian",
                        "total gio",
                    )
                ),
                "month": parse_month_value(cell_value(ws, r, headers, "thang", "ngay tao"))
                or parse_month_value(cell_value(ws, r, headers, "ten nhiem vu", "nhiem vu", "ten task", "task")),
                "employee": cell_value(
                    ws,
                    r,
                    headers,
                    "nguoi lam task (dien ho va ten)",
                    "nguoi lam task",
                    "giao cho",
                    "nguoi tao",
                    "nguoi thuc hien",
                ),
                "mnv": cell_value(ws, r, headers, "ma nhan vien", "mnv"),
            }

            if not any([
                raw["bu"],
                raw["project_name"],
                raw["task_name"],
                raw["hours"],
                raw["month"],
                raw["employee"],
                raw["mnv"],
            ]):
                continue

            mnv = raw["mnv"] or lookup_employee_code(employee_lookup, raw["employee"])
            rows.append(
                {
                    "source_sheet": ws.title,
                    "source_row": r,
                    "bu": raw["bu"],
                    "project_name": raw["project_name"],
                    "task_name": raw["task_name"],
                    "hours": raw["hours"],
                    "month": raw["month"],
                    "employee": raw["employee"],
                    "mnv": mnv,
                }
            )

    return rows


def derive_project_name(task_name: Any, fallback: Any = None) -> Any:
    if fallback:
        return fallback
    if not task_name:
        return None
    text = str(task_name).strip()
    text = re.sub(r"^\d{3,4}[_-]", "", text)
    text = re.sub(r"^(ACCA|CFA|CMA|IT|MEDIA)[_-]", "", text, flags=re.I)
    if "_" in text:
        candidate = text.split("_", 1)[0].strip()
        if candidate and len(candidate) <= 40:
            return candidate
    if " - " in text:
        candidate = text.split(" - ", 1)[0].strip()
        if candidate and len(candidate) <= 40:
            return candidate
    return None


def ensure_media_rows(ws, count: int) -> None:
    needed_max_row = MEDIA_TEMPLATE_ROW + max(0, count - 1)
    if needed_max_row > ws.max_row:
        for r in range(ws.max_row + 1, needed_max_row + 1):
            ensure_row_style(ws, MEDIA_TEMPLATE_ROW, r, ws.max_column)


def media_formula_h(row: int) -> str:
    return (
        f"=_xlfn.XLOOKUP(1, "
        f"('Lương nhân viên full time'!$A:$A='Data media ACCA+CFA+CMA'!E{row})*"
        f"('Lương nhân viên full time'!$E:$E='Data media ACCA+CFA+CMA'!G{row}), "
        f"'Lương nhân viên full time'!$AG:$AG)"
    )


def write_media_sheet(ws, rows: list[dict[str, Any]]) -> None:
    clear_values(ws, MEDIA_TEMPLATE_ROW, ws.max_row, 1, min(ws.max_column, 9))
    ensure_media_rows(ws, len(rows))

    for idx, row in enumerate(rows):
        r = MEDIA_TEMPLATE_ROW + idx
        project_name = derive_project_name(row["task_name"], row["project_name"])
        ws.cell(r, 1).value = row["bu"]
        ws.cell(r, 2).value = project_name
        ws.cell(r, 3).value = row["task_name"]
        ws.cell(r, 4).value = row["hours"]
        ws.cell(r, 5).value = row["month"]
        ws.cell(r, 6).value = row["employee"]
        ws.cell(r, 7).value = row["mnv"]
        ws.cell(r, 8).value = media_formula_h(r)
        ws.cell(r, 9).value = f"=D{r}/H{r}"


def write_media_audit_sheet(wb, sheet_name: str, rows: list[dict[str, Any]]) -> None:
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(title=sheet_name)

    ws.append(MEDIA_AUDIT_FIELDS)
    for row in rows:
        ws.append([row.get(field) for field in MEDIA_AUDIT_FIELDS])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def reset_sheet(wb, sheet_name: str):
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)
        return ws
    return wb.create_sheet(title=sheet_name)


def append_table(ws, start_row: int, title: str, headers: list[str], rows: list[list[Any]]) -> int:
    ws.cell(start_row, 1).value = title
    header_row = start_row + 1
    for c, value in enumerate(headers, start=1):
        ws.cell(header_row, c).value = value
    for r_offset, row in enumerate(rows, start=1):
        for c, value in enumerate(row, start=1):
            ws.cell(header_row + r_offset, c).value = value
    if rows:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{header_row + len(rows)}"
    return header_row + len(rows) + 2


def clear_worksheet(ws) -> None:
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)


def copy_sheet_content(src_ws, dst_ws) -> None:
    clear_worksheet(dst_ws)

    for row in src_ws.iter_rows():
        for src_cell in row:
            dst_cell = dst_ws.cell(src_cell.row, src_cell.column)
            dst_cell.value = src_cell.value
            if src_cell.has_style:
                dst_cell._style = copy(src_cell._style)
            if src_cell.number_format:
                dst_cell.number_format = src_cell.number_format
            if src_cell.font:
                dst_cell.font = copy(src_cell.font)
            if src_cell.fill:
                dst_cell.fill = copy(src_cell.fill)
            if src_cell.border:
                dst_cell.border = copy(src_cell.border)
            if src_cell.alignment:
                dst_cell.alignment = copy(src_cell.alignment)
            if src_cell.protection:
                dst_cell.protection = copy(src_cell.protection)

    for col_key, col_dim in src_ws.column_dimensions.items():
        dst_dim = dst_ws.column_dimensions[col_key]
        dst_dim.width = col_dim.width
        dst_dim.hidden = col_dim.hidden
        dst_dim.outlineLevel = col_dim.outlineLevel

    for row_idx, row_dim in src_ws.row_dimensions.items():
        dst_dim = dst_ws.row_dimensions[row_idx]
        dst_dim.height = row_dim.height
        dst_dim.hidden = row_dim.hidden
        dst_dim.outlineLevel = row_dim.outlineLevel

    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

    dst_ws.freeze_panes = src_ws.freeze_panes
    if src_ws.auto_filter and src_ws.auto_filter.ref:
        dst_ws.auto_filter.ref = src_ws.auto_filter.ref


def find_payroll_header_row(ws) -> int | None:
    for r in range(1, min(ws.max_row or 0, 20) + 1):
        values = {norm(ws.cell(r, c).value) for c in range(1, (ws.max_column or 0) + 1)}
        if "thang" in values and ("ma nhan vien" in values or "ma nhan su" in values):
            return r
    return None


def payroll_duplicate_keys(ws, header_row: int | None) -> int:
    if header_row is None:
        return 0
    headers = {norm(ws.cell(header_row, c).value): c for c in range(1, (ws.max_column or 0) + 1)}
    month_col = headers.get("thang")
    mnv_col = headers.get("ma nhan vien") or headers.get("ma nhan su")
    if not month_col or not mnv_col:
        return 0

    seen: set[tuple[Any, Any]] = set()
    duplicates = 0
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        month = clean(ws.cell(r, month_col).value)
        mnv = clean(ws.cell(r, mnv_col).value)
        if month is None and mnv is None:
            continue
        key = (month, mnv)
        if key in seen:
            duplicates += 1
        else:
            seen.add(key)
    return duplicates


def write_payroll_checkpoint_sheet(wb, rows: list[list[Any]]) -> None:
    ws = reset_sheet(wb, "Check_Payroll")
    append_table(
        ws,
        1,
        "Payroll Sync Check",
        [
            "Sheet",
            "Source path",
            "Source modified",
            "Source exists",
            "Source sheet exists",
            "Output sheet exists",
            "Rows copied",
            "Columns copied",
            "Header row",
            "Duplicate Month+MNV",
            "Status",
        ],
        rows,
    )
    ws.freeze_panes = "A3"
    for col, width in {"A": 32, "B": 90, "C": 22, "K": 60}.items():
        ws.column_dimensions[col].width = width


def sync_payroll_from_onedrive(wb, payroll_path: Path = PAYROLL_SOURCE_PATH) -> list[list[Any]]:
    source_exists = payroll_path.exists()
    modified = datetime.fromtimestamp(payroll_path.stat().st_mtime) if source_exists else None
    rows: list[list[Any]] = []

    if not source_exists:
        for sheet_name in PAYROLL_SHEETS:
            rows.append([sheet_name, str(payroll_path), None, False, False, find_sheet(wb, sheet_name) is not None, 0, 0, None, 0, "failed: payroll source file not found"])
        write_payroll_checkpoint_sheet(wb, rows)
        return rows

    source_wb = load_workbook(payroll_path, data_only=False)
    for sheet_name in PAYROLL_SHEETS:
        source_ws = find_sheet(source_wb, sheet_name)
        target_ws = find_sheet(wb, sheet_name)
        source_sheet_exists = source_ws is not None
        output_sheet_exists = target_ws is not None
        copied_rows = 0
        copied_cols = 0
        header_row = None
        duplicate_count = 0
        status = "OK"

        if source_ws is None:
            status = "failed: source sheet missing"
        elif target_ws is None:
            status = "failed: output sheet missing"
        else:
            copy_sheet_content(source_ws, target_ws)
            copied_rows = source_ws.max_row or 0
            copied_cols = source_ws.max_column or 0
            header_row = find_payroll_header_row(source_ws)
            duplicate_count = payroll_duplicate_keys(source_ws, header_row)
            if header_row is None:
                status = "warning: header row not detected"
            elif duplicate_count:
                status = f"warning: {duplicate_count} duplicate Month+MNV rows"

        rows.append([sheet_name, str(payroll_path), modified, source_exists, source_sheet_exists, output_sheet_exists, copied_rows, copied_cols, header_row, duplicate_count, status])

    write_payroll_checkpoint_sheet(wb, rows)
    return rows


def write_approval_guide_sheet(wb, output_path: Path, approval_file: Path | None = None) -> None:
    ws = reset_sheet(wb, "Huong_dan_Approval")
    rows = [
        ["Step", "Approval source workbook", "Command / Value", "Instruction"],
        [1, str(output_path), "", "Open this workbook and review Check_Payroll / Check_IT_CPNS / IT_New_Project_Master / Check_IT_Downstream / Check_Media_Timesheet."],
        [2, str(output_path), "YES / Y / TRUE / 1 / APPLY / APPROVED", "Enter one of these values in Apply? for rows you want to approve. Leave blank or enter NO to skip."],
        [3, str(output_path), "CLEAN_COST_PROJECT_TEXT", "IT action: cleans extra whitespace in Chi phí nhân sự IT project names."],
        [4, str(output_path), "FILL_COST_MONTH_FORMULA", "IT action: fills missing month formulas in an existing Chi phí nhân sự IT row."],
        [5, str(output_path), "ADD_COST_ROW", "IT action: adds a missing Chi phí nhân sự IT row when an existing project row and employee template row are available."],
        [6, str(output_path), "ADD_PROJECT_MASTER_FIRST", "IT checkpoint action: review IT_New_Project_Master and approve/fill master metadata before adding downstream rows."],
        [7, str(output_path), "ADD_TO_PROJECT_CATALOG / ADD_TO_CAPITALIZATION / ADD_MNV_TO_IT_CHECKING", "Downstream actions: catalog and capitalization can auto-apply; MNV checking remains manual until salary base source is validated."],
        [8, str(output_path), "FIX_MEDIA_WEIGHT_OR_MONTH", "Media action: fills missing month when inferable from task name and refreshes Data media H/I formulas."],
        [9, str(output_path), f"py -3 automate_kpi.py --approval-file \"{output_path}\"", "After marking approvals in this workbook, run this command and point --approval-file to this exact edited workbook."],
        [10, str(approval_file) if approval_file else "(none)", "", "If this run used an approval file, that file is listed here. To change your mind, edit that approval file or rerun without --approval-file."],
        [11, str(output_path), "py -3 automate_kpi.py", "Rerun without --approval-file to ignore all approvals and generate a fresh workbook from template/source data."],
        [12, str(output_path), "IT_Approval_Result / Project_Master_Approval_Result / Downstream_Approval_Result / Media_Approval_Result", "The next output workbook will include these result sheets with applied/skipped/failed statuses."],
    ]
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row, start=1):
            ws.cell(r, c).value = value

    command_start = len(rows) + 3
    command_rows = [
        ["Step", "When to use", "Command", "What it does", "Which file to edit/use"],
        [
            1,
            "Create output from existing local raw input files",
            "py -3 automate_kpi.py",
            "Uses data/input/raw/IT.xlsx and data/input/raw/MEDIA.xlsx, writes Timesheet IT/Data media, creates checkpoint sheets, then opens the output workbook.",
            "No approval file is needed. The newest output is also saved in data/output/final.",
        ],
        [
            2,
            "Download latest source files first, then automate",
            "py -3 automate_kpi.py --download",
            "Downloads IT/MEDIA source XLSX files from config/sources.json, then creates output and checkpoints.",
            "Use when Google Sheets/source files changed and local raw files need refresh.",
        ],
        [
            3,
            "Review checkpoint and choose fixes",
            "(edit workbook, no terminal command)",
            "Open Check_IT_CPNS, IT_New_Project_Master, Check_IT_Downstream, and Check_Media_Timesheet. Put YES in Apply? only for rows you approve.",
            f"Edit the approval source workbook shown in column B, usually {output_path}.",
        ],
        [
            4,
            "Apply approved fixes",
            f'py -3 automate_kpi.py --approval-file "{output_path}"',
            "Reads YES rows from the approval workbook, applies supported fixes, then creates a new output workbook.",
            "The --approval-file path must point to the exact workbook where you entered YES.",
        ],
        [
            5,
            "Download latest sources and apply approved fixes in one run",
            f'py -3 automate_kpi.py --download --approval-file "{output_path}"',
            "Refreshes raw IT/MEDIA files first, then applies approved fixes from the selected approval workbook.",
            "Use carefully: approvals are read from the old workbook, data is refreshed from latest source.",
        ],
        [
            6,
            "Change your mind and ignore all approvals",
            "py -3 automate_kpi.py",
            "Creates a fresh output without reading any YES rows.",
            "Do not pass --approval-file.",
        ],
        [
            7,
            "Change your mind on selected rows",
            "(edit workbook, then rerun command in step 4)",
            "Change Apply? from YES to NO or blank for rows you no longer approve, then rerun with --approval-file.",
            "Edit the same approval workbook before rerunning.",
        ],
        [
            8,
            "Check what was applied",
            "(open output workbook)",
            "Review IT_Approval_Result, Project_Master_Approval_Result, Downstream_Approval_Result, and Media_Approval_Result for applied/skipped/failed statuses.",
            "Use the newest output workbook created after running --approval-file.",
        ],
        [
            9,
            "Run without opening output",
            "py -3 automate_kpi.py --no-open",
            "Creates output from local raw files but does not open the result workbook.",
            "Use for scheduled/background runs.",
        ],
        [
            10,
            "Open output explicitly after run",
            "py -3 automate_kpi.py --open",
            "Creates output from local raw files and opens the result workbook. This is now the default behavior.",
            "No approval file is needed unless you also add --approval-file.",
        ],
    ]
    for r_offset, row in enumerate(command_rows, start=command_start):
        for c, value in enumerate(row, start=1):
            ws.cell(r_offset, c).value = value

    process_start = command_start + len(command_rows) + 3
    process_rows = [
        ["Process", "Checkpoint sheet", "Action to approve", "What user should fill", "Result sheet"],
        [
            "IT mapping / cost formula",
            "Check_IT_CPNS",
            "CLEAN_COST_PROJECT_TEXT / FILL_COST_MONTH_FORMULA / ADD_COST_ROW",
            "Put YES in Apply? only for rows with confirmed mismatch. ADD_COST_ROW needs an existing project row in Chi phi nhan su IT.",
            "IT_Approval_Result",
        ],
        [
            "Payroll sync",
            "Check_Payroll",
            "No approval action",
            "Script reads the OneDrive local file and replaces Lương nhân viên full time / Lương nhân viên part time in the output. Review source path, modified time, row counts, and duplicate Month+MNV warnings.",
            "Check_Payroll",
        ],
        [
            "IT new project master",
            "IT_New_Project_Master",
            "ADD_PROJECT_MASTER_FIRST",
            "Check Suggested project code, Suggested BU, System, then put YES. This adds catalog row and Chi phi nhan su IT rows for employees detected in Timesheet IT.",
            "Project_Master_Approval_Result",
        ],
        [
            "Carry forward approved IT project master",
            "Project_Master_Approval_Result",
            "CARRY_FORWARD_PROJECT_MASTER",
            "If you use a previous output workbook as --approval-file, applied project master rows are copied forward automatically so Chi phi nhan su IT rows are not lost in the next output.",
            "Project_Master_Approval_Result",
        ],
        [
            "IT downstream",
            "Check_IT_Downstream",
            "ADD_TO_PROJECT_CATALOG / ADD_TO_CAPITALIZATION",
            "Fill/confirm catalog metadata columns, then put YES. ADD_TO_CAPITALIZATION copies an existing IT row in 3.Von hoa and keeps translated formulas.",
            "Downstream_Approval_Result",
        ],
        [
            "IT checking MNV",
            "Check_IT_Downstream",
            "ADD_MNV_TO_IT_CHECKING",
            "Review manually for now. The script reports this action but does not auto-add because Checking Von hoa IT needs a validated salary base row.",
            "Downstream_Approval_Result",
        ],
        [
            "Media source formula",
            "Check_Media_Timesheet",
            "FIX_MEDIA_WEIGHT_OR_MONTH",
            "Put YES for Data media rows where month/weight formula can be repaired from source information.",
            "Media_Approval_Result",
        ],
        [
            "Ignore or undo approvals",
            "Any checkpoint",
            "NO / blank",
            "Change Apply? to NO or blank before rerun. To ignore all approvals, run without --approval-file.",
            "No approval result sheet is needed.",
        ],
    ]
    for r_offset, row in enumerate(process_rows, start=process_start):
        for c, value in enumerate(row, start=1):
            ws.cell(r_offset, c).value = value

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["D"].width = 110
    ws.column_dimensions["E"].width = 90


def numeric(value: Any) -> float:
    if is_num(value):
        return float(value)
    return 0.0


def month_number(value: Any) -> int | None:
    if isinstance(value, bool):
        return None
    if is_num(value) and int(value) == value and 1 <= int(value) <= 12:
        return int(value)
    return None


def find_sheet(wb, sheet_name: str):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    target = norm(sheet_name)
    for ws in wb.worksheets:
        if norm(ws.title) == target:
            return ws
    return None


def project_exists_in_column(ws, project: Any, col: int, start_row: int = 1) -> bool:
    target = clean(project)
    if target is None:
        return False
    for r in range(start_row, ws.max_row + 1):
        if clean(ws.cell(r, col).value) == target:
            return True
    return False


def find_project_code_in_cost(cost, project: Any) -> Any:
    if cost is None:
        return None
    target = clean(project)
    for r in range(3, cost.max_row + 1):
        if clean(cost.cell(r, 2).value) == target:
            return cost.cell(r, 1).value
    return None


def next_it_project_code(wb) -> str:
    max_num = 2600
    code_re = re.compile(r"^IT(\d+)$", re.I)
    for ws_name, col in [(SHEET_PROJECT_CATALOG, 2), (SHEET_IT_COST, 1)]:
        ws = find_sheet(wb, ws_name)
        if ws is None:
            continue
        for r in range(1, ws.max_row + 1):
            value = clean(ws.cell(r, col).value)
            if not value:
                continue
            match = code_re.match(str(value))
            if match:
                max_num = max(max_num, int(match.group(1)))
    return f"IT{max_num + 1}"


def collect_it_source_projects(ts) -> dict[str, dict[str, Any]]:
    projects: dict[str, dict[str, Any]] = {}
    for r in range(IT_HEADER_ROW + 2, ts.max_row + 1):
        month = month_number(ts.cell(r, 5).value)
        project = clean(ts.cell(r, 2).value)
        if month is None or not project:
            continue
        row_total = 0.0
        employees = []
        for c in range(6, min(ts.max_column, 26) + 1):
            value = numeric(ts.cell(r, c).value)
            if not value:
                continue
            row_total += value
            employees.append(str(ts.cell(IT_HEADER_ROW + 1, c).value or get_column_letter(c)))
        if not row_total:
            continue
        item = projects.setdefault(
            str(project),
            {
                "project": project,
                "system": clean(ts.cell(r, 4).value),
                "months": set(),
                "employees": set(),
                "rows": [],
                "total": 0.0,
            },
        )
        if not item.get("system") and clean(ts.cell(r, 4).value):
            item["system"] = clean(ts.cell(r, 4).value)
        item["months"].add(month)
        item["employees"].update(employees)
        item["rows"].append(r)
        item["total"] += row_total
    return projects


def write_it_new_project_master_sheet(wb) -> None:
    ts = find_sheet(wb, SHEET_IT)
    cost = find_sheet(wb, SHEET_IT_COST)
    catalog = find_sheet(wb, SHEET_PROJECT_CATALOG)
    if ts is None:
        return

    rows: list[list[Any]] = []
    next_code_num = int(next_it_project_code(wb).replace("IT", ""))
    for item in sorted(collect_it_source_projects(ts).values(), key=lambda x: str(x["project"])):
        exists_in_cost = project_exists_in_column(cost, item["project"], 2, 3) if cost is not None else False
        exists_in_catalog = project_exists_in_column(catalog, item["project"], 3, 2) if catalog is not None else False
        if exists_in_cost and exists_in_catalog:
            continue

        existing_cost_code = find_project_code_in_cost(cost, item["project"])
        if existing_cost_code:
            suggested_code = existing_cost_code
        else:
            suggested_code = f"IT{next_code_num}"
            next_code_num += 1
        months = sorted(item["months"])
        issue_parts = []
        if not exists_in_catalog:
            issue_parts.append("missing in 1.Danh mục dự án")
        if not exists_in_cost:
            issue_parts.append("missing in Chi phí nhân sự IT")
        rows.append(
            [
                "Timesheet IT",
                item["project"],
                item.get("system"),
                suggested_code,
                "SAPP",
                min(months) if months else None,
                max(months) if months else None,
                ", ".join(sorted(item["employees"])),
                ", ".join(str(r) for r in item["rows"]),
                item["total"],
                exists_in_cost,
                exists_in_catalog,
                "; ".join(issue_parts),
                "ADD_PROJECT_MASTER_FIRST",
                None,
                None,
            ]
        )

    ws = reset_sheet(wb, "IT_New_Project_Master")
    append_table(
        ws,
        1,
        "New Project Master Required",
        [
            "Project source",
            "Project name",
            "System",
            "Suggested project code",
            "Suggested BU",
            "Start month",
            "End month",
            "Employees detected",
            "Timesheet rows",
            "Total input %",
            "Exists in Chi phí nhân sự IT",
            "Exists in 1.Danh mục dự án",
            "Issue",
            "Recommended action",
            "Apply?",
            "Approval notes",
        ],
        rows,
    )
    ws.freeze_panes = "A3"
    for col, width in {"A": 18, "B": 60, "C": 24, "D": 22, "H": 80, "I": 24, "M": 42, "N": 28}.items():
        ws.column_dimensions[col].width = width


def collect_it_cost_projects(cost) -> dict[str, dict[str, Any]]:
    projects: dict[str, dict[str, Any]] = {}
    for r in range(3, cost.max_row + 1):
        code = clean(cost.cell(r, 1).value)
        project = clean(cost.cell(r, 2).value)
        if not code or not project:
            continue
        item = projects.setdefault(
            str(code),
            {
                "code": code,
                "project": project,
                "system": clean(cost.cell(r, 3).value),
                "bu": clean(cost.cell(r, 6).value),
                "mnvs": set(),
                "cost_rows": [],
            },
        )
        if not item.get("system") and clean(cost.cell(r, 3).value):
            item["system"] = clean(cost.cell(r, 3).value)
        if not item.get("bu") and clean(cost.cell(r, 6).value):
            item["bu"] = clean(cost.cell(r, 6).value)
        mnv = clean(cost.cell(r, 7).value)
        if mnv:
            item["mnvs"].add(str(mnv))
        item["cost_rows"].append(r)
    return projects


def values_in_column(ws, col: int, start_row: int = 1) -> set[str]:
    if ws is None:
        return set()
    result: set[str] = set()
    for r in range(start_row, ws.max_row + 1):
        value = clean(ws.cell(r, col).value)
        if value is not None:
            result.add(str(value))
    return result


def collect_capitalization_projects(ws) -> dict[str, dict[str, Any]]:
    projects: dict[str, dict[str, Any]] = {}
    if ws is None:
        return projects
    for r in range(3, ws.max_row + 1):
        code = clean(ws.cell(r, 2).value)
        if not code:
            continue
        projects[str(code)] = {
            "row": r,
            "year": ws.cell(r, 1).value,
            "project": ws.cell(r, 3).value,
            "bu": ws.cell(r, 4).value,
            "classification": ws.cell(r, 5).value,
            "start": ws.cell(r, 6).value,
            "end": ws.cell(r, 7).value,
        }
    return projects


def write_it_downstream_checkpoint_sheet(wb) -> None:
    cost = find_sheet(wb, SHEET_IT_COST)
    if cost is None:
        return

    catalog = find_sheet(wb, SHEET_PROJECT_CATALOG)
    checking = find_sheet(wb, SHEET_IT_CHECKING)
    capitalization = find_sheet(wb, SHEET_CAPITALIZATION)

    projects = collect_it_cost_projects(cost)
    catalog_codes = values_in_column(catalog, 2, 2)
    capitalization_codes = values_in_column(capitalization, 2, 3)
    capitalization_projects = collect_capitalization_projects(capitalization)
    checking_mnvs = values_in_column(checking, 1, 4)

    detail_rows: list[list[Any]] = []
    missing_catalog = 0
    missing_capitalization = 0
    projects_with_missing_checking_mnvs = 0

    for item in sorted(projects.values(), key=lambda x: str(x["code"])):
        code = str(item["code"])
        cap_meta = capitalization_projects.get(code, {})
        mnvs = sorted(item["mnvs"])
        missing_mnvs = [mnv for mnv in mnvs if mnv not in checking_mnvs]
        exists_catalog = code in catalog_codes
        exists_capitalization = code in capitalization_codes
        if not exists_catalog:
            missing_catalog += 1
        if not exists_capitalization:
            missing_capitalization += 1
        if missing_mnvs:
            projects_with_missing_checking_mnvs += 1

        issue_parts = []
        actions = []
        if not exists_catalog:
            issue_parts.append("missing in 1.Danh mục dự án")
            actions.append("ADD_TO_PROJECT_CATALOG")
        if not exists_capitalization:
            issue_parts.append("missing in 3.Vốn hóa")
            actions.append("ADD_TO_CAPITALIZATION")
        if missing_mnvs:
            issue_parts.append("missing MNV in Checking Vốn hóa IT")
            actions.append("ADD_MNV_TO_IT_CHECKING")

        if not issue_parts:
            continue

        detail_rows.append(
            [
                code,
                item["project"],
                item.get("system"),
                item.get("bu"),
                cap_meta.get("year") or 2026,
                item.get("bu") or cap_meta.get("bu") or "SAPP",
                cap_meta.get("classification"),
                cap_meta.get("start"),
                cap_meta.get("end"),
                True,
                ", ".join(str(r) for r in item["cost_rows"]),
                ", ".join(mnvs),
                exists_catalog,
                exists_capitalization,
                len(missing_mnvs),
                ", ".join(missing_mnvs),
                "; ".join(issue_parts),
                "; ".join(actions),
                None,
                None,
            ]
        )

    summary_rows = [
        ["Projects in Chi phí nhân sự IT", len(projects)],
        ["Missing in 1.Danh mục dự án", missing_catalog],
        ["Missing in 3.Vốn hóa", missing_capitalization],
        ["Projects with MNV missing in Checking Vốn hóa IT", projects_with_missing_checking_mnvs],
    ]

    ws = reset_sheet(wb, "Check_IT_Downstream")
    next_row = append_table(ws, 1, "Summary", ["Metric", "Count"], summary_rows)
    append_table(
        ws,
        next_row,
        "Details",
        [
            "Project code",
            "Project name",
            "System",
            "BU",
            "Catalog year",
            "Catalog BU",
            "Catalog classification",
            "Catalog start date",
            "Catalog end date",
            "Catalog capitalization flag",
            "Chi phí nhân sự IT rows",
            "MNVs in cost",
            "Exists in 1.Danh mục dự án",
            "Exists in 3.Vốn hóa",
            "Missing Checking MNV count",
            "Missing Checking MNVs",
            "Issue",
            "Recommended action",
            "Apply?",
            "Approval notes",
        ],
        detail_rows,
    )
    ws.freeze_panes = "A3"
    for col, width in {"A": 22, "B": 60, "E": 14, "F": 16, "G": 22, "K": 24, "L": 80, "P": 40, "Q": 42, "R": 52}.items():
        ws.column_dimensions[col].width = width


def write_it_checkpoint_sheet(wb) -> None:
    ts = find_sheet(wb, SHEET_IT)
    cost = find_sheet(wb, SHEET_IT_COST)
    if ts is None or cost is None:
        return

    source: dict[tuple[int, str, int], dict[str, Any]] = {}
    for r in range(IT_HEADER_ROW + 2, ts.max_row + 1):
        month = month_number(ts.cell(r, 5).value)
        if month is None:
            continue
        project = ts.cell(r, 2).value
        if not project:
            continue
        for c in range(6, min(ts.max_column, 26) + 1):
            value = numeric(ts.cell(r, c).value)
            if not value:
                continue
            key = (month, str(project), c)
            item = source.setdefault(
                key,
                {
                    "month": month,
                    "project": project,
                    "employee_col": c,
                    "employee": ts.cell(IT_HEADER_ROW + 1, c).value,
                    "rows": [],
                    "value": 0.0,
                },
            )
            item["value"] += value
            item["rows"].append(r)

    processed: dict[tuple[int, str, int], float] = {}
    clean_candidates: dict[tuple[int, str, int], list[tuple[int, str]]] = {}
    sum_col_re = re.compile(r"'Timesheet IT'!\$([A-Z]+):\$\1")
    for r in range(3, cost.max_row + 1):
        raw_project = cost.cell(r, 2).value
        if not raw_project:
            continue
        for c in range(1, cost.max_column + 1):
            month = month_number(cost.cell(1, c).value)
            if month is None:
                continue
            formula = cost.cell(r, c).value
            if not isinstance(formula, str) or "SUMIFS" not in formula or "Timesheet IT" not in formula:
                continue
            match = sum_col_re.search(formula)
            if not match:
                continue
            sum_col = column_letters_to_number(match.group(1))
            exact_total = 0.0
            for tr in range(1, ts.max_row + 1):
                if ts.cell(tr, 2).value == raw_project and ts.cell(tr, 5).value == month:
                    exact_total += numeric(ts.cell(tr, sum_col).value)
            processed[(month, str(raw_project), sum_col)] = processed.get((month, str(raw_project), sum_col), 0.0) + exact_total
            clean_candidates.setdefault((month, str(clean(raw_project)), sum_col), []).append((r, str(raw_project)))

    summary_rows: list[list[Any]] = []
    detail_rows: list[list[Any]] = []
    all_months = range(1, 13)
    for month in all_months:
        input_total = sum(item["value"] for key, item in source.items() if key[0] == month)
        processed_total = sum(value for key, value in processed.items() if key[0] == month)
        summary_rows.append([month, input_total, processed_total, processed_total - input_total])

    for key, item in sorted(source.items(), key=lambda kv: (kv[0][0], str(kv[0][1]), kv[0][2])):
        month, project, employee_col = key
        output = processed.get(key, 0.0)
        diff = output - item["value"]
        if abs(diff) < 1e-9:
            continue
        cleaned_key = (month, str(clean(project)), employee_col)
        candidates = clean_candidates.get(cleaned_key, [])
        existing_cost_row = find_cost_row_by_project_and_sum_col(cost, project, employee_col)
        issue = "missing row/formula in Chi phí nhân sự IT"
        recommended_action = "ADD_COST_ROW"
        if find_cost_project_sample_row(cost, project) is None:
            issue = "new project master required"
            recommended_action = "ADD_PROJECT_MASTER_FIRST"
        if existing_cost_row:
            issue = "missing month formula in Chi phí nhân sự IT"
            recommended_action = "FILL_COST_MONTH_FORMULA"
        if candidates:
            issue = "criteria text mismatch, likely whitespace"
            recommended_action = "CLEAN_COST_PROJECT_TEXT"
        detail_rows.append(
            [
                month,
                item["project"],
                item["employee"],
                get_column_letter(employee_col),
                ", ".join(str(r) for r in item["rows"]),
                item["value"],
                output,
                diff,
                ", ".join(str(row) for row, _ in candidates),
                candidates[0][1] if candidates else None,
                bool(candidates),
                issue,
                recommended_action,
                None,
                None,
            ]
        )

    employee_month_totals: dict[tuple[int, str], dict[str, Any]] = {}
    for r in range(3, cost.max_row + 1):
        mnv = cost.cell(r, 7).value
        employee = cost.cell(r, 8).value
        if not mnv and not employee:
            continue
        employee_key = str(mnv or employee)
        for c in range(1, cost.max_column + 1):
            month = month_number(cost.cell(1, c).value)
            if month is None:
                continue
            formula = cost.cell(r, c).value
            if not isinstance(formula, str) or "SUMIFS" not in formula or "Timesheet IT" not in formula:
                continue
            total = processed.get((month, str(cost.cell(r, 2).value), column_from_it_sumifs(formula)), 0.0)
            item = employee_month_totals.setdefault(
                (month, employee_key),
                {"month": month, "mnv": mnv, "employee": employee, "value": 0.0, "cost_rows": []},
            )
            item["value"] += total
            if total:
                item["cost_rows"].append(r)

    overload_rows = [
        [
            item["month"],
            item["mnv"],
            item["employee"],
            item["value"],
            item["value"] - 1,
            ", ".join(str(r) for r in sorted(set(item["cost_rows"]))),
            "over 100%",
        ]
        for item in sorted(employee_month_totals.values(), key=lambda x: (x["month"], str(x["mnv"] or x["employee"])))
        if item["value"] > 1 + 1e-9
    ]

    ws = reset_sheet(wb, "Check_IT_CPNS")
    next_row = append_table(
        ws,
        1,
        "Summary",
        ["Month", "Input Timesheet IT", "Processed Chi phí nhân sự IT", "Diff"],
        summary_rows,
    )
    append_table(
        ws,
        next_row,
        "Details",
        [
            "Month",
            "Project",
            "Employee",
            "Employee column",
            "Timesheet rows",
            "Input value",
            "Processed value",
            "Diff",
            "Cost rows matched after clean",
            "Cost project sample",
            "Matched after clean",
            "Issue",
            "Recommended action",
            "Apply?",
            "Approval notes",
        ],
        detail_rows,
    )
    append_table(
        ws,
        ws.max_row + 3,
        "Employee Month Over 100%",
        ["Month", "MNV", "Employee", "Total % work", "Over by", "Cost rows", "Issue"],
        overload_rows,
    )
    ws.freeze_panes = "A3"
    for col, width in {"A": 12, "B": 60, "C": 24, "E": 18, "I": 24, "J": 60, "L": 36}.items():
        ws.column_dimensions[col].width = width


def column_from_it_sumifs(formula: Any) -> int | None:
    if not isinstance(formula, str):
        return None
    match = re.search(r"'Timesheet IT'!\$([A-Z]+):\$\1", formula)
    if not match:
        return None
    return column_letters_to_number(match.group(1))


def column_letters_to_number(letters: str) -> int:
    total = 0
    for ch in letters:
        total = total * 26 + ord(ch.upper()) - ord("A") + 1
    return total


def media_standard_hours(wb, month: Any, mnv: Any) -> float | None:
    salary = find_sheet(wb, SHEET_SALARY_FULLTIME)
    if salary is None or not month or not mnv:
        return None
    for r in range(1, salary.max_row + 1):
        if salary.cell(r, 1).value == month and salary.cell(r, 5).value == mnv:
            value = salary.cell(r, 33).value
            if is_num(value) and value:
                return float(value)
            working_days = salary.cell(r, 14).value
            extra_hours = salary.cell(r, 20).value
            if is_num(working_days):
                return float(working_days) * 8 + numeric(extra_hours)
    return None


def media_weight_value(wb, data_ws, row: int) -> tuple[float | None, str | None]:
    weight = data_ws.cell(row, 9).value
    if is_num(weight):
        return float(weight), None
    hours = data_ws.cell(row, 4).value
    month = month_number(data_ws.cell(row, 5).value)
    mnv = data_ws.cell(row, 7).value
    standard_hours = media_standard_hours(wb, month, mnv)
    if is_num(hours) and standard_hours:
        issue = None
        if not weight:
            issue = "missing weight formula/value in Data media"
        return float(hours) / standard_hours, issue
    if is_num(hours) and month and mnv and not standard_hours:
        return None, "missing salary standard hours for media employee/month"
    if any([hours, month, mnv, data_ws.cell(row, 2).value]):
        return None, "cannot compute media weight"
    return None, None


def write_media_checkpoint_sheet(wb) -> None:
    data_ws = find_sheet(wb, SHEET_MEDIA)
    media_ws = find_sheet(wb, SHEET_TIMESHEET_MEDIA)
    if data_ws is None or media_ws is None:
        return

    source: dict[tuple[Any, str, str], dict[str, Any]] = {}
    source_issues: list[list[Any]] = []
    for r in range(MEDIA_TEMPLATE_ROW, data_ws.max_row + 1):
        project = data_ws.cell(r, 2).value
        month = month_number(data_ws.cell(r, 5).value)
        mnv = data_ws.cell(r, 7).value
        employee = data_ws.cell(r, 6).value
        if not any([project, month, mnv, data_ws.cell(r, 4).value]):
            continue
        value, issue = media_weight_value(wb, data_ws, r)
        if issue:
            source_issues.append(
                [
                    r,
                    project,
                    month,
                    employee,
                    mnv,
                    data_ws.cell(r, 4).value,
                    data_ws.cell(r, 9).value,
                    issue,
                    "FIX_MEDIA_WEIGHT_OR_MONTH",
                    None,
                    None,
                ]
            )
        if value is None or not project or not mnv or not month:
            continue
        key = (month, str(project), str(mnv))
        item = source.setdefault(
            key,
            {"month": month, "project": project, "employee": employee, "mnv": mnv, "rows": [], "value": 0.0},
        )
        item["value"] += value
        item["rows"].append(r)

    target_rows: dict[tuple[int, str, str], list[int]] = {}
    clean_target_rows: dict[tuple[int, str, str], list[int]] = {}
    for r in range(4, media_ws.max_row + 1):
        project = media_ws.cell(r, 4).value
        mnv = media_ws.cell(r, 6).value
        if not project or not mnv:
            continue
        for c in range(9, min(media_ws.max_column, 20) + 1):
            month = month_number(media_ws.cell(1, c).value)
            formula = media_ws.cell(r, c).value
            if month and isinstance(formula, str) and "Data media ACCA+CFA+CMA" in formula:
                target_rows.setdefault((month, str(project), str(mnv)), []).append(r)
                clean_target_rows.setdefault((month, str(clean(project)), str(mnv)), []).append(r)

    summary_rows: list[list[Any]] = []
    detail_rows: list[list[Any]] = []
    for month in range(1, 13):
        input_total = sum(item["value"] for key, item in source.items() if key[0] == month)
        processed_total = sum(item["value"] for key, item in source.items() if key[0] == month and key in target_rows)
        summary_rows.append([month, input_total, processed_total, processed_total - input_total])

    for key, item in sorted(source.items(), key=lambda kv: (str(kv[0][0]), str(kv[0][1]), str(kv[0][2]))):
        if key in target_rows:
            continue
        cleaned_key = (key[0], str(clean(key[1])), key[2])
        clean_rows = sorted(set(clean_target_rows.get(cleaned_key, [])))
        issue = "missing row/formula in Timesheet Media"
        recommended_action = "ADD_TIMESHEET_MEDIA_ROW"
        if clean_rows:
            issue = "criteria text mismatch, likely whitespace"
            recommended_action = "CLEAN_MEDIA_PROJECT_TEXT"
        detail_rows.append(
            [
                item["month"],
                item["project"],
                item["employee"],
                item["mnv"],
                ", ".join(str(r) for r in item["rows"]),
                item["value"],
                0,
                -item["value"],
                None,
                ", ".join(str(r) for r in clean_rows),
                bool(clean_rows),
                issue,
                recommended_action,
                None,
                None,
            ]
        )

    employee_month_totals: dict[tuple[int, str], dict[str, Any]] = {}
    for key, item in source.items():
        if key not in target_rows:
            continue
        employee_key = str(item["mnv"] or item["employee"])
        month = int(item["month"])
        total_item = employee_month_totals.setdefault(
            (month, employee_key),
            {"month": month, "mnv": item["mnv"], "employee": item["employee"], "value": 0.0, "timesheet_rows": []},
        )
        total_item["value"] += item["value"]
        total_item["timesheet_rows"].extend(target_rows.get(key, []))

    overload_rows = [
        [
            item["month"],
            item["mnv"],
            item["employee"],
            item["value"],
            item["value"] - 1,
            ", ".join(str(r) for r in sorted(set(item["timesheet_rows"]))),
            "over 100%",
        ]
        for item in sorted(employee_month_totals.values(), key=lambda x: (x["month"], str(x["mnv"] or x["employee"])))
        if item["value"] > 1 + 1e-9
    ]

    ws = reset_sheet(wb, "Check_Media_Timesheet")
    next_row = append_table(
        ws,
        1,
        "Summary",
        ["Month", "Input Data media", "Processed Timesheet Media", "Diff"],
        summary_rows,
    )
    next_row = append_table(
        ws,
        next_row,
        "Details",
        [
            "Month",
            "Project",
            "Employee",
            "MNV",
            "Data media rows",
            "Input value",
            "Processed value",
            "Diff",
            "Timesheet Media rows matched exact",
            "Timesheet Media rows matched after clean",
            "Matched after clean",
            "Issue",
            "Recommended action",
            "Apply?",
            "Approval notes",
        ],
        detail_rows,
    )
    next_row = append_table(
        ws,
        next_row,
        "Source Issues",
        [
            "Data media row",
            "Project",
            "Month",
            "Employee",
            "MNV",
            "Hours",
            "Weight value/formula",
            "Issue",
            "Recommended action",
            "Apply?",
            "Approval notes",
        ],
        source_issues,
    )
    append_table(
        ws,
        next_row,
        "Employee Month Over 100%",
        ["Month", "MNV", "Employee", "Total % work", "Over by", "Timesheet Media rows", "Issue"],
        overload_rows,
    )
    ws.freeze_panes = "A3"
    for col, width in {"A": 14, "B": 42, "C": 24, "D": 16, "E": 18, "J": 28, "L": 34, "M": 28}.items():
        ws.column_dimensions[col].width = width


def approval_value_is_yes(value: Any) -> bool:
    return norm(value) in APPROVAL_YES_VALUES


def find_section_header(ws, section_name: str) -> int | None:
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == section_name:
            return r + 1
    return None


def read_it_mapping_approvals(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise SystemExit(f"Approval file not found: {path}")

    wb = load_workbook(path, data_only=False)
    ws = find_sheet(wb, "Check_IT_CPNS")
    if ws is None:
        raise SystemExit("Approval file does not contain sheet Check_IT_CPNS.")

    header_row = find_section_header(ws, "Details")
    if header_row is None:
        raise SystemExit("Approval file does not contain Details section in Check_IT_CPNS.")

    headers = {str(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}
    required = ["Month", "Project", "Employee", "Employee column", "Timesheet rows", "Issue", "Recommended action", "Apply?"]
    missing = [name for name in required if name not in headers]
    if missing:
        raise SystemExit(f"Approval file is missing required columns in Details: {missing}")

    approvals: list[dict[str, Any]] = []
    for r in range(header_row + 1, ws.max_row + 1):
        marker = ws.cell(r, 1).value
        if marker in {"Employee Month Over 100%", "Summary", "Details"}:
            break
        if not any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
            continue
        if not approval_value_is_yes(ws.cell(r, headers["Apply?"]).value):
            continue

        approvals.append(
            {
                "approval_row": r,
                "month": month_number(ws.cell(r, headers["Month"]).value),
                "project": ws.cell(r, headers["Project"]).value,
                "employee": ws.cell(r, headers["Employee"]).value,
                "employee_col": ws.cell(r, headers["Employee column"]).value,
                "timesheet_rows": ws.cell(r, headers["Timesheet rows"]).value,
                "cost_rows_after_clean": ws.cell(r, headers.get("Cost rows matched after clean", 0)).value if headers.get("Cost rows matched after clean") else None,
                "issue": ws.cell(r, headers["Issue"]).value,
                "action": ws.cell(r, headers["Recommended action"]).value,
            }
        )
    return approvals


def read_media_approvals(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise SystemExit(f"Approval file not found: {path}")

    wb = load_workbook(path, data_only=False)
    ws = find_sheet(wb, "Check_Media_Timesheet")
    if ws is None:
        return []

    approvals: list[dict[str, Any]] = []
    for section in ["Details", "Source Issues"]:
        header_row = find_section_header(ws, section)
        if header_row is None:
            continue
        headers = {str(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}
        if "Apply?" not in headers or "Recommended action" not in headers:
            continue

        for r in range(header_row + 1, ws.max_row + 1):
            marker = ws.cell(r, 1).value
            if marker in {"Summary", "Details", "Source Issues", "Employee Month Over 100%"}:
                break
            if not any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
                continue
            if not approval_value_is_yes(ws.cell(r, headers["Apply?"]).value):
                continue

            approvals.append(
                {
                    "section": section,
                    "approval_row": r,
                    "data_row": ws.cell(r, headers.get("Data media row", 0)).value if headers.get("Data media row") else None,
                    "project": ws.cell(r, headers.get("Project", 0)).value if headers.get("Project") else None,
                    "month": month_number(ws.cell(r, headers.get("Month", 0)).value) if headers.get("Month") else None,
                    "employee": ws.cell(r, headers.get("Employee", 0)).value if headers.get("Employee") else None,
                    "mnv": ws.cell(r, headers.get("MNV", 0)).value if headers.get("MNV") else None,
                    "action": ws.cell(r, headers["Recommended action"]).value,
                }
            )
    return approvals


def read_downstream_approvals(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise SystemExit(f"Approval file not found: {path}")

    wb = load_workbook(path, data_only=False)
    ws = find_sheet(wb, "Check_IT_Downstream")
    if ws is None:
        return []

    header_row = find_section_header(ws, "Details")
    if header_row is None:
        return []
    headers = {str(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}
    required = ["Project code", "Project name", "Recommended action", "Apply?"]
    missing = [name for name in required if name not in headers]
    if missing:
        raise SystemExit(f"Approval file is missing required columns in Check_IT_Downstream: {missing}")

    approvals: list[dict[str, Any]] = []
    for r in range(header_row + 1, ws.max_row + 1):
        marker = ws.cell(r, 1).value
        if marker in {"Summary", "Details"}:
            break
        if not any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
            continue
        if not approval_value_is_yes(ws.cell(r, headers["Apply?"]).value):
            continue
        approvals.append(
            {
                "approval_row": r,
                "project_code": ws.cell(r, headers["Project code"]).value,
                "project_name": ws.cell(r, headers["Project name"]).value,
                "system": ws.cell(r, headers.get("System", 0)).value if headers.get("System") else None,
                "catalog_year": ws.cell(r, headers.get("Catalog year", 0)).value if headers.get("Catalog year") else 2026,
                "catalog_bu": ws.cell(r, headers.get("Catalog BU", 0)).value if headers.get("Catalog BU") else None,
                "catalog_classification": ws.cell(r, headers.get("Catalog classification", 0)).value if headers.get("Catalog classification") else None,
                "catalog_start": ws.cell(r, headers.get("Catalog start date", 0)).value if headers.get("Catalog start date") else None,
                "catalog_end": ws.cell(r, headers.get("Catalog end date", 0)).value if headers.get("Catalog end date") else None,
                "catalog_capitalization": ws.cell(r, headers.get("Catalog capitalization flag", 0)).value if headers.get("Catalog capitalization flag") else True,
                "action": ws.cell(r, headers["Recommended action"]).value,
            }
        )
    return approvals


def read_project_master_approvals(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise SystemExit(f"Approval file not found: {path}")

    wb = load_workbook(path, data_only=False)
    ws = find_sheet(wb, "IT_New_Project_Master")
    if ws is None:
        return []

    header_row = find_section_header(ws, "New Project Master Required")
    if header_row is None:
        return []
    headers = {str(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}
    required = ["Project name", "Suggested project code", "Suggested BU", "Recommended action", "Apply?"]
    missing = [name for name in required if name not in headers]
    if missing:
        raise SystemExit(f"Approval file is missing required columns in IT_New_Project_Master: {missing}")

    approvals: list[dict[str, Any]] = []
    for r in range(header_row + 1, ws.max_row + 1):
        if not any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1)):
            continue
        if not approval_value_is_yes(ws.cell(r, headers["Apply?"]).value):
            continue
        approvals.append(
            {
                "approval_row": r,
                "project": ws.cell(r, headers["Project name"]).value,
                "system": ws.cell(r, headers.get("System", 0)).value if headers.get("System") else None,
                "project_code": ws.cell(r, headers["Suggested project code"]).value,
                "bu": ws.cell(r, headers["Suggested BU"]).value,
                "start_month": month_number(ws.cell(r, headers.get("Start month", 0)).value) if headers.get("Start month") else None,
                "end_month": month_number(ws.cell(r, headers.get("End month", 0)).value) if headers.get("End month") else None,
                "timesheet_rows": ws.cell(r, headers.get("Timesheet rows", 0)).value if headers.get("Timesheet rows") else None,
                "action": ws.cell(r, headers["Recommended action"]).value,
            }
        )
    return approvals


def read_project_master_result_projects(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        raise SystemExit(f"Approval file not found: {path}")

    wb = load_workbook(path, data_only=False, read_only=True)
    ws = find_sheet(wb, "Project_Master_Approval_Result")
    if ws is None:
        return []

    header_row = find_section_header(ws, "Project Master Approval Result")
    if header_row is None:
        return []
    headers = {str(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}
    required = ["Action", "Project code", "Project name", "Status"]
    if any(name not in headers for name in required):
        return []

    projects: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for r in range(header_row + 1, ws.max_row + 1):
        action = clean(ws.cell(r, headers["Action"]).value)
        code = clean(ws.cell(r, headers["Project code"]).value)
        project = clean(ws.cell(r, headers["Project name"]).value)
        status = clean(ws.cell(r, headers["Status"]).value)
        if action != "ADD_PROJECT_MASTER_FIRST" or not code or not project:
            continue
        if not status or not str(status).startswith("applied:"):
            continue
        key = (str(code), str(project))
        if key in seen:
            continue
        seen.add(key)
        projects.append({"project_code": code, "project": project})
    return projects


def copy_row(ws, source_row: int, target_row: int) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for c in range(1, ws.max_column + 1):
        src = ws.cell(source_row, c)
        dst = ws.cell(target_row, c)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.font:
            dst.font = copy(src.font)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.border:
            dst.border = copy(src.border)
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)
        if isinstance(src.value, str) and src.value.startswith("="):
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
        else:
            dst.value = src.value


def copy_row_between_sheets(source_ws, source_row: int, target_ws, target_row: int) -> None:
    target_ws.row_dimensions[target_row].height = source_ws.row_dimensions[source_row].height
    for c in range(1, source_ws.max_column + 1):
        src = source_ws.cell(source_row, c)
        dst = target_ws.cell(target_row, c)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.font:
            dst.font = copy(src.font)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.border:
            dst.border = copy(src.border)
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)
        if isinstance(src.value, str) and src.value.startswith("="):
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
        else:
            dst.value = src.value


def parse_int_list(value: Any) -> list[int]:
    if value is None:
        return []
    return [int(item) for item in re.findall(r"\d+", str(value))]


def find_cost_row_by_project_and_sum_col(cost, project: Any, sum_col: int) -> int | None:
    for r in range(3, cost.max_row + 1):
        if clean(cost.cell(r, 2).value) != clean(project):
            continue
        for c in range(1, cost.max_column + 1):
            if column_from_it_sumifs(cost.cell(r, c).value) == sum_col:
                return r
    return None


def find_cost_project_sample_row(cost, project: Any) -> int | None:
    for r in range(3, cost.max_row + 1):
        if clean(cost.cell(r, 2).value) == clean(project):
            return r
    return None


def find_cost_employee_template_row(cost, sum_col: int) -> int | None:
    for r in range(3, cost.max_row + 1):
        for c in range(1, cost.max_column + 1):
            if column_from_it_sumifs(cost.cell(r, c).value) == sum_col:
                return r
    return None


def find_cost_percent_col(cost, month: int) -> int | None:
    for c in range(1, cost.max_column + 1):
        if month_number(cost.cell(1, c).value) == month and norm(cost.cell(2, c).value) == "cong viec":
            return c
    return None


def fill_cost_month_formulas(cost, row: int, month: int, sum_col: int) -> str:
    percent_col = find_cost_percent_col(cost, month)
    if percent_col is None:
        return f"failed: no % công việc column for month {month}"

    sum_letter = get_column_letter(sum_col)
    percent_letter = get_column_letter(percent_col)
    salary_letter = get_column_letter(percent_col + 1)
    cost_letter = get_column_letter(percent_col + 2)
    changes = []

    percent_formula = (
        f"=SUMIFS('Timesheet IT'!${sum_letter}:${sum_letter},"
        f"'Timesheet IT'!$B:$B,'Chi phí nhân sự IT'!$B{row},"
        f"'Timesheet IT'!$E:$E,'Chi phí nhân sự IT'!{percent_letter}$1)"
    )
    salary_formula = (
        f"=SUMIFS('Lương nhân viên full time'!$AF:$AF,"
        f"'Lương nhân viên full time'!$E:$E,'Chi phí nhân sự IT'!$G{row},"
        f"'Lương nhân viên full time'!$A:$A,'Chi phí nhân sự IT'!{percent_letter}$1)"
    )
    cost_formula = f"={percent_letter}{row}*{salary_letter}{row}"

    targets = [
        (percent_col, percent_formula, percent_letter),
        (percent_col + 1, salary_formula, salary_letter),
        (percent_col + 2, cost_formula, cost_letter),
    ]
    for col, formula, letter in targets:
        cell = cost.cell(row, col)
        if cell.value in (None, ""):
            cell.value = formula
            changes.append(letter)

    return f"applied: filled {', '.join(changes)} row {row}" if changes else f"skipped: month {month} formulas already exist in row {row}"


def apply_clean_cost_project_text(cost, approval: dict[str, Any]) -> str:
    rows = parse_int_list(approval.get("cost_rows_after_clean"))
    if not rows:
        return "skipped: no cost row listed"
    changed = []
    for r in rows:
        current = cost.cell(r, 2).value
        cleaned = clean(current)
        if current != cleaned:
            cost.cell(r, 2).value = cleaned
            changed.append(r)
    return f"applied: cleaned rows {changed}" if changed else f"skipped: already clean rows {rows}"


def apply_add_cost_row(cost, approval: dict[str, Any]) -> str:
    employee_col = approval.get("employee_col")
    if not employee_col:
        return "failed: missing Employee column"
    sum_col = column_letters_to_number(str(employee_col))
    project = approval.get("project")
    month = approval.get("month")

    existing = find_cost_row_by_project_and_sum_col(cost, project, sum_col)
    if existing:
        if month:
            return fill_cost_month_formulas(cost, existing, month, sum_col)
        return f"skipped: existing cost row {existing}"

    project_row = find_cost_project_sample_row(cost, project)
    if project_row is None:
        return "failed: no existing project row to copy project metadata"

    template_row = find_cost_employee_template_row(cost, sum_col)
    if template_row is None:
        return "failed: no employee template row for Timesheet IT column"

    target_row = cost.max_row + 1
    copy_row(cost, template_row, target_row)
    for c in range(1, 7):
        cost.cell(target_row, c).value = clean(cost.cell(project_row, c).value)
    cost.cell(target_row, 2).value = clean(project)
    fill_status = fill_cost_month_formulas(cost, target_row, month, sum_col) if month else "skipped: no month"
    return f"applied: added cost row {target_row} from employee template row {template_row} and project row {project_row}; {fill_status}"


def apply_fill_cost_month_formula(cost, approval: dict[str, Any]) -> str:
    employee_col = approval.get("employee_col")
    month = approval.get("month")
    project = approval.get("project")
    if not employee_col or not month:
        return "failed: missing Employee column or Month"
    sum_col = column_letters_to_number(str(employee_col))
    existing = find_cost_row_by_project_and_sum_col(cost, project, sum_col)
    if not existing:
        return "failed: no existing cost row"
    return fill_cost_month_formulas(cost, existing, month, sum_col)


def write_approval_result_sheet(wb, results: list[list[Any]]) -> None:
    ws = reset_sheet(wb, "IT_Approval_Result")
    headers = ["Approval row", "Action", "Project", "Employee", "Employee column", "Status"]
    append_table(ws, 1, "IT Approval Result", headers, results)
    for col, width in {"A": 14, "B": 28, "C": 60, "D": 24, "E": 16, "F": 80}.items():
        ws.column_dimensions[col].width = width


def write_media_approval_result_sheet(wb, results: list[list[Any]]) -> None:
    ws = reset_sheet(wb, "Media_Approval_Result")
    headers = ["Approval row", "Section", "Action", "Data media row", "Project", "Employee", "MNV", "Status"]
    append_table(ws, 1, "Media Approval Result", headers, results)
    for col, width in {"A": 14, "B": 16, "C": 28, "D": 16, "E": 42, "F": 24, "G": 16, "H": 80}.items():
        ws.column_dimensions[col].width = width


def write_downstream_approval_result_sheet(wb, results: list[list[Any]]) -> None:
    ws = reset_sheet(wb, "Downstream_Approval_Result")
    headers = ["Approval row", "Action", "Project code", "Project name", "Status"]
    append_table(ws, 1, "Downstream Approval Result", headers, results)
    for col, width in {"A": 14, "B": 34, "C": 18, "D": 60, "E": 90}.items():
        ws.column_dimensions[col].width = width


def write_project_master_approval_result_sheet(wb, results: list[list[Any]]) -> None:
    ws = reset_sheet(wb, "Project_Master_Approval_Result")
    headers = ["Approval row", "Action", "Project code", "Project name", "Status"]
    append_table(ws, 1, "Project Master Approval Result", headers, results)
    for col, width in {"A": 14, "B": 30, "C": 18, "D": 60, "E": 110}.items():
        ws.column_dimensions[col].width = width


def first_empty_row(ws, key_cols: list[int], start_row: int = 2) -> int:
    for r in range(start_row, ws.max_row + 1):
        if all(clean(ws.cell(r, c).value) is None for c in key_cols):
            return r
    return ws.max_row + 1


def bool_value(value: Any, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = norm(value)
    if text in {"yes", "y", "true", "1", "co", "x"}:
        return True
    if text in {"no", "n", "false", "0", "khong"}:
        return False
    return default


def apply_add_to_project_catalog(catalog, approval: dict[str, Any]) -> str:
    code = clean(approval.get("project_code"))
    project = clean(approval.get("project_name"))
    if not code or not project:
        return "failed: missing project code or project name"
    if project_exists_in_column(catalog, code, 2, 2):
        return f"skipped: project code {code} already exists in catalog"

    target_row = first_empty_row(catalog, [2, 3], start_row=2)
    ensure_row_style(catalog, 2, target_row, catalog.max_column)
    catalog.cell(target_row, 1).value = approval.get("catalog_year") or 2026
    catalog.cell(target_row, 2).value = code
    catalog.cell(target_row, 3).value = project
    catalog.cell(target_row, 4).value = approval.get("system")
    catalog.cell(target_row, 5).value = approval.get("catalog_bu") or "SAPP"
    catalog.cell(target_row, 6).value = approval.get("catalog_classification")
    catalog.cell(target_row, 7).value = approval.get("catalog_start")
    catalog.cell(target_row, 8).value = approval.get("catalog_end")
    catalog.cell(target_row, 12).value = bool_value(approval.get("catalog_capitalization"), default=True)
    return f"applied: added catalog row {target_row}"


def find_capitalization_template_row(capitalization) -> int | None:
    for r in range(3, capitalization.max_row + 1):
        code = clean(capitalization.cell(r, 2).value)
        if code and str(code).startswith("IT"):
            return r
    return 3 if capitalization.max_row >= 3 else None


def apply_add_to_capitalization(capitalization, approval: dict[str, Any]) -> str:
    code = clean(approval.get("project_code"))
    project = clean(approval.get("project_name"))
    if not code or not project:
        return "failed: missing project code or project name"
    if project_exists_in_column(capitalization, code, 2, 3):
        return f"skipped: project code {code} already exists in 3.Von hoa"

    template_row = find_capitalization_template_row(capitalization)
    if template_row is None:
        return "failed: no capitalization template row found"

    target_row = first_empty_row(capitalization, [2, 3], start_row=3)
    copy_row(capitalization, template_row, target_row)
    capitalization.cell(target_row, 1).value = approval.get("catalog_year") or 2026
    capitalization.cell(target_row, 2).value = code
    capitalization.cell(target_row, 3).value = project
    capitalization.cell(target_row, 4).value = approval.get("catalog_bu") or "HO"
    capitalization.cell(target_row, 5).value = approval.get("catalog_classification")
    capitalization.cell(target_row, 6).value = approval.get("catalog_start")
    capitalization.cell(target_row, 7).value = approval.get("catalog_end")
    capitalization.cell(target_row, 8).value = None
    return f"applied: added 3.Von hoa row {target_row} from template row {template_row}"


def source_entries_for_project(ts, project: Any) -> dict[int, dict[str, Any]]:
    entries: dict[int, dict[str, Any]] = {}
    target = clean(project)
    for r in range(IT_HEADER_ROW + 2, ts.max_row + 1):
        if clean(ts.cell(r, 2).value) != target:
            continue
        month = month_number(ts.cell(r, 5).value)
        if month is None:
            continue
        for c in range(6, min(ts.max_column, 26) + 1):
            value = numeric(ts.cell(r, c).value)
            if not value:
                continue
            item = entries.setdefault(c, {"months": set(), "rows": [], "total": 0.0})
            item["months"].add(month)
            item["rows"].append(r)
            item["total"] += value
    return entries


def project_master_to_catalog_approval(approval: dict[str, Any]) -> dict[str, Any]:
    return {
        "project_code": approval.get("project_code"),
        "project_name": approval.get("project"),
        "system": approval.get("system"),
        "catalog_year": 2026,
        "catalog_bu": approval.get("bu") or "SAPP",
        "catalog_classification": None,
        "catalog_start": None,
        "catalog_end": None,
        "catalog_capitalization": True,
    }


def apply_project_master_approval(wb, approval: dict[str, Any]) -> str:
    action = clean(approval.get("action"))
    if action != "ADD_PROJECT_MASTER_FIRST":
        return f"skipped: unsupported action {action}"

    ts = find_sheet(wb, SHEET_IT)
    cost = find_sheet(wb, SHEET_IT_COST)
    catalog = find_sheet(wb, SHEET_PROJECT_CATALOG)
    if ts is None or cost is None or catalog is None:
        return "failed: missing Timesheet IT, Chi phí nhân sự IT, or 1.Danh mục dự án"

    project = clean(approval.get("project"))
    code = clean(approval.get("project_code"))
    if not project or not code:
        return "failed: missing project name or project code"

    statuses = [apply_add_to_project_catalog(catalog, project_master_to_catalog_approval(approval))]
    source_entries = source_entries_for_project(ts, project)
    if not source_entries:
        statuses.append("failed: no Timesheet IT source entries found")
        return "; ".join(statuses)

    added_rows = []
    skipped_rows = []
    for employee_col, source in sorted(source_entries.items()):
        existing = find_cost_row_by_project_and_sum_col(cost, project, employee_col)
        if existing:
            for month in sorted(source["months"]):
                fill_cost_month_formulas(cost, existing, month, employee_col)
            skipped_rows.append(str(existing))
            continue

        template_row = find_cost_employee_template_row(cost, employee_col)
        if template_row is None:
            statuses.append(f"failed: no employee template for Timesheet IT column {get_column_letter(employee_col)}")
            continue

        target_row = cost.max_row + 1
        copy_row(cost, template_row, target_row)
        cost.cell(target_row, 1).value = code
        cost.cell(target_row, 2).value = project
        cost.cell(target_row, 3).value = approval.get("system")
        cost.cell(target_row, 4).value = None
        cost.cell(target_row, 5).value = None
        cost.cell(target_row, 6).value = approval.get("bu") or "SAPP"
        for month in sorted(source["months"]):
            fill_cost_month_formulas(cost, target_row, month, employee_col)
        added_rows.append(str(target_row))

    if added_rows:
        statuses.append(f"added cost rows {', '.join(added_rows)}")
    if skipped_rows:
        statuses.append(f"updated existing cost rows {', '.join(skipped_rows)}")
    return "; ".join(statuses)


def apply_it_mapping_approvals(wb, approval_file: Path | None) -> None:
    if approval_file is None:
        return
    cost = find_sheet(wb, SHEET_IT_COST)
    if cost is None:
        raise SystemExit(f"Output workbook does not contain sheet {SHEET_IT_COST}.")

    approvals = read_it_mapping_approvals(approval_file)
    results: list[list[Any]] = []
    for approval in approvals:
        action = clean(approval.get("action"))
        try:
            if action == "CLEAN_COST_PROJECT_TEXT":
                status = apply_clean_cost_project_text(cost, approval)
            elif action == "ADD_COST_ROW":
                status = apply_add_cost_row(cost, approval)
            elif action == "FILL_COST_MONTH_FORMULA":
                status = apply_fill_cost_month_formula(cost, approval)
            elif action == "ADD_PROJECT_MASTER_FIRST":
                status = "skipped: approve this action in IT_New_Project_Master"
            else:
                status = f"skipped: unsupported action {action}"
        except Exception as exc:
            status = f"failed: {exc}"
        results.append(
            [
                approval.get("approval_row"),
                action,
                approval.get("project"),
                approval.get("employee"),
                approval.get("employee_col"),
                status,
            ]
        )

    write_approval_result_sheet(wb, results)


def find_row_by_clean_value(ws, col: int, value: Any, start_row: int = 1) -> int | None:
    target = clean(value)
    if target is None:
        return None
    for r in range(start_row, ws.max_row + 1):
        if clean(ws.cell(r, col).value) == target:
            return r
    return None


def carry_forward_project_master_results(wb, approval_file: Path | None) -> list[list[Any]]:
    if approval_file is None:
        return []

    projects = read_project_master_result_projects(approval_file)
    if not projects:
        return []

    source_wb = load_workbook(approval_file, data_only=False)
    source_cost = find_sheet(source_wb, SHEET_IT_COST)
    source_catalog = find_sheet(source_wb, SHEET_PROJECT_CATALOG)
    target_cost = find_sheet(wb, SHEET_IT_COST)
    target_catalog = find_sheet(wb, SHEET_PROJECT_CATALOG)
    if source_cost is None or source_catalog is None or target_cost is None or target_catalog is None:
        return [[None, "CARRY_FORWARD_PROJECT_MASTER", None, None, "failed: missing source/target catalog or cost sheet"]]

    results: list[list[Any]] = []
    for project in projects:
        code = project["project_code"]
        project_name = project["project"]
        statuses = []

        if find_row_by_clean_value(target_catalog, 2, code, start_row=2):
            statuses.append(f"skipped catalog: project code {code} already exists")
        else:
            source_catalog_row = find_row_by_clean_value(source_catalog, 2, code, start_row=2)
            if source_catalog_row is None:
                statuses.append("failed catalog: approved project not found in approval file catalog")
            else:
                target_catalog_row = first_empty_row(target_catalog, [2, 3], start_row=2)
                copy_row_between_sheets(source_catalog, source_catalog_row, target_catalog, target_catalog_row)
                statuses.append(f"carried catalog row {target_catalog_row}")

        existing_cost_rows = [
            r
            for r in range(3, target_cost.max_row + 1)
            if clean(target_cost.cell(r, 1).value) == clean(code) and clean(target_cost.cell(r, 2).value) == clean(project_name)
        ]
        if existing_cost_rows:
            statuses.append(f"skipped cost: rows already exist {', '.join(str(r) for r in existing_cost_rows)}")
        else:
            source_cost_rows = [
                r
                for r in range(3, source_cost.max_row + 1)
                if clean(source_cost.cell(r, 1).value) == clean(code) and clean(source_cost.cell(r, 2).value) == clean(project_name)
            ]
            if not source_cost_rows:
                statuses.append("failed cost: approved project rows not found in approval file cost sheet")
            else:
                target_rows = []
                for source_row in source_cost_rows:
                    target_row = target_cost.max_row + 1
                    copy_row_between_sheets(source_cost, source_row, target_cost, target_row)
                    target_rows.append(target_row)
                statuses.append(f"carried cost rows {', '.join(str(r) for r in target_rows)}")

        results.append([None, "CARRY_FORWARD_PROJECT_MASTER", code, project_name, "; ".join(statuses)])
    return results


def apply_project_master_approvals(wb, approval_file: Path | None) -> None:
    if approval_file is None:
        return

    approvals = read_project_master_approvals(approval_file)
    results: list[list[Any]] = carry_forward_project_master_results(wb, approval_file)
    for approval in approvals:
        action = clean(approval.get("action"))
        try:
            status = apply_project_master_approval(wb, approval)
        except Exception as exc:
            status = f"failed: {exc}"
        results.append(
            [
                approval.get("approval_row"),
                action,
                approval.get("project_code"),
                approval.get("project"),
                status,
            ]
        )

    if results:
        write_project_master_approval_result_sheet(wb, results)


def apply_fix_media_weight_or_month(data_ws, approval: dict[str, Any]) -> str:
    row_value = approval.get("data_row")
    if row_value is None:
        return "failed: missing Data media row"
    row = int(row_value)
    if row < MEDIA_TEMPLATE_ROW or row > data_ws.max_row:
        return f"failed: Data media row {row} is outside sheet range"

    changes = []
    month = month_number(data_ws.cell(row, 5).value)
    if month is None:
        month = parse_month_value(data_ws.cell(row, 3).value)
        if month is not None:
            data_ws.cell(row, 5).value = month
            changes.append("month")

    data_ws.cell(row, 8).value = media_formula_h(row)
    data_ws.cell(row, 9).value = f"=D{row}/H{row}"
    changes.extend(["H formula", "I formula"])

    if month is None:
        return f"partial: filled formulas row {row}, but month could not be inferred"
    return f"applied: filled {', '.join(changes)} row {row}"


def apply_media_approvals(wb, approval_file: Path | None) -> None:
    if approval_file is None:
        return
    data_ws = find_sheet(wb, SHEET_MEDIA)
    if data_ws is None:
        raise SystemExit(f"Output workbook does not contain sheet {SHEET_MEDIA}.")

    approvals = read_media_approvals(approval_file)
    results: list[list[Any]] = []
    for approval in approvals:
        action = clean(approval.get("action"))
        try:
            if action == "FIX_MEDIA_WEIGHT_OR_MONTH":
                status = apply_fix_media_weight_or_month(data_ws, approval)
            else:
                status = f"skipped: unsupported action {action}"
        except Exception as exc:
            status = f"failed: {exc}"
        results.append(
            [
                approval.get("approval_row"),
                approval.get("section"),
                action,
                approval.get("data_row"),
                approval.get("project"),
                approval.get("employee"),
                approval.get("mnv"),
                status,
            ]
        )

    if results:
        write_media_approval_result_sheet(wb, results)


def apply_downstream_approvals(wb, approval_file: Path | None) -> None:
    if approval_file is None:
        return
    catalog = find_sheet(wb, SHEET_PROJECT_CATALOG)
    if catalog is None:
        raise SystemExit(f"Output workbook does not contain sheet {SHEET_PROJECT_CATALOG}.")
    capitalization = find_sheet(wb, SHEET_CAPITALIZATION)
    if capitalization is None:
        raise SystemExit(f"Output workbook does not contain sheet {SHEET_CAPITALIZATION}.")

    approvals = read_downstream_approvals(approval_file)
    results: list[list[Any]] = []
    for approval in approvals:
        actions = [clean(action) for action in str(approval.get("action") or "").split(";")]
        statuses = []
        for action in actions:
            if not action:
                continue
            try:
                if action == "ADD_TO_PROJECT_CATALOG":
                    statuses.append(apply_add_to_project_catalog(catalog, approval))
                elif action == "ADD_TO_CAPITALIZATION":
                    statuses.append(apply_add_to_capitalization(capitalization, approval))
                elif action == "ADD_MNV_TO_IT_CHECKING":
                    statuses.append("skipped: ADD_MNV_TO_IT_CHECKING not enabled; Checking Von hoa IT salary base needs manual source validation")
                else:
                    statuses.append(f"skipped: unsupported action {action}")
            except Exception as exc:
                statuses.append(f"failed {action}: {exc}")
        if statuses:
            results.append(
                [
                    approval.get("approval_row"),
                    approval.get("action"),
                    approval.get("project_code"),
                    approval.get("project_name"),
                    "; ".join(statuses),
                ]
            )

    if results:
        write_downstream_approval_result_sheet(wb, results)


def run(
    download: bool,
    open_after: bool = True,
    it_url: str | None = None,
    media_url: str | None = None,
    approval_file: Path | None = None,
) -> Path:
    setup_dirs()
    sources = load_source_config()
    if it_url:
        sources["IT"]["url"] = it_url
    if media_url:
        sources["MEDIA"]["url"] = media_url
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    print("=" * 60)
    print(f"Von hoa chi phi nhan su 2026 - {datetime.now():%d/%m/%Y %H:%M}")
    print("=" * 60)

    if download:
        print("[1/4] Downloading source files...")
        download_all(sources)
    else:
        print("[1/4] Using local files in data/input/raw")

    missing = [k for k, meta in sources.items() if not (DIRS["raw"] / meta["file"]).exists()]
    if missing:
        raise SystemExit(
            f"Missing source files: {missing}\n"
            "Run with --download or place the XLSX files in data/input/raw."
        )

    print("[2/4] Reading sources...")
    employee_lookup = {}

    template = find_template()
    if template is None:
        raise SystemExit(
            "Template file not found.\n"
            f"Put the workbook in: {DIRS['template']}\n"
            f"Expected one of: {', '.join(TEMPLATE_CANDIDATES)}"
        )

    template_wb = load_workbook(template)
    employee_lookup = build_employee_lookup(template_wb)

    it_source = DIRS["raw"] / sources["IT"]["file"]
    media_source = DIRS["raw"] / sources["MEDIA"]["file"]

    it_sheet_name, it_rows = read_it_sheet(it_source)
    print(f"  IT -> sheet '{it_sheet_name}', {len(it_rows)} rows")

    media_rows = extract_media_rows(media_source, employee_lookup)
    media_sheets = sorted({row["source_sheet"] for row in media_rows})
    print(f"  MEDIA -> {len(media_rows)} rows from sheets: {', '.join(media_sheets) if media_sheets else '(none)'}")

    print("[3/4] Preparing output workbook...")
    output_path = DIRS["final"] / f"von_hoa_{ts}.xlsx"
    shutil.copy2(template, output_path)

    wb = load_workbook(output_path)
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    written_it = False
    written_media = False
    payroll_rows = sync_payroll_from_onedrive(wb)
    payroll_ok = all(row[-1] == "OK" for row in payroll_rows)
    print(f"  payroll sync <- {PAYROLL_SOURCE_PATH} ({len(payroll_rows)} sheets, ok={payroll_ok})")

    if SHEET_IT in wb.sheetnames:
        write_it_sheet(wb[SHEET_IT], it_rows, it_source)
        written_it = True
        print(f"  wrote {SHEET_IT}")
    else:
        print(f"  missing sheet: {SHEET_IT}")

    if SHEET_MEDIA in wb.sheetnames:
        write_media_sheet(wb[SHEET_MEDIA], media_rows)
        written_media = True
        print(f"  wrote {SHEET_MEDIA}")
    else:
        print(f"  missing sheet: {SHEET_MEDIA}")

    media_by_bu = {
        "CFA": [row for row in media_rows if norm(row.get("bu")) == "cfa"],
        "CMA": [row for row in media_rows if norm(row.get("bu")) == "cma"],
        "ACCA": [row for row in media_rows if norm(row.get("bu")) == "acca"],
    }
    for bu, rows in media_by_bu.items():
        write_media_audit_sheet(wb, f"Media_{bu}", rows)
        print(f"  audit sheet Media_{bu} <- {len(rows)} rows")

    if approval_file is not None:
        apply_it_mapping_approvals(wb, approval_file)
        apply_project_master_approvals(wb, approval_file)
        apply_downstream_approvals(wb, approval_file)
        apply_media_approvals(wb, approval_file)
        print(f"  applied approvals from {approval_file}")

    write_it_checkpoint_sheet(wb)
    write_it_new_project_master_sheet(wb)
    write_it_downstream_checkpoint_sheet(wb)
    write_media_checkpoint_sheet(wb)
    write_approval_guide_sheet(wb, output_path, approval_file)
    print("  wrote checkpoint sheets: Check_Payroll, Check_IT_CPNS, IT_New_Project_Master, Check_IT_Downstream, Check_Media_Timesheet")

    wb.save(output_path)

    print("[4/4] Writing staging files...")
    try:
        csv_path = DIRS["staging"] / f"media_rows_{ts}.csv"
        with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=MEDIA_AUDIT_FIELDS)
            writer.writeheader()
            writer.writerows(media_rows)

        xlsx_path = DIRS["staging"] / f"media_rows_{ts}.xlsx"
        staging_wb = Workbook()
        staging_ws = staging_wb.active
        staging_ws.title = "media_rows"
        staging_ws.append(MEDIA_AUDIT_FIELDS)
        for row in media_rows:
            staging_ws.append([row.get(field) for field in MEDIA_AUDIT_FIELDS])
        staging_ws.freeze_panes = "A2"
        staging_ws.auto_filter.ref = staging_ws.dimensions
        staging_wb.save(xlsx_path)
    except Exception:
        pass

    print("=" * 60)
    print("DONE")
    print(f"Output: {output_path}")
    print(f"IT written: {written_it}")
    print(f"Media written: {written_media}")
    print(f"Payroll sync: {payroll_ok}")
    print("=" * 60)

    if open_after:
        import subprocess

        subprocess.Popen(["explorer", str(output_path)])

    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Von hoa chi phi nhan su 2026 automation")
    parser.add_argument("--download", action="store_true", help="Download the latest source XLSX files first")
    parser.add_argument("--open", action="store_true", dest="open_after", default=True, help="Open the result workbook after finishing")
    parser.add_argument("--no-open", action="store_false", dest="open_after", help="Do not open the result workbook after finishing")
    parser.add_argument("--it-url", dest="it_url", help="Override the IT source URL for this run")
    parser.add_argument("--media-url", dest="media_url", help="Override the media source URL for this run")
    parser.add_argument(
        "--approval-file",
        type=Path,
        help="Workbook containing Check_IT_CPNS approvals. Put YES in Apply? for rows to automate.",
    )
    args = parser.parse_args()
    run(
        download=args.download,
        open_after=args.open_after,
        it_url=args.it_url,
        media_url=args.media_url,
        approval_file=args.approval_file,
    )


if __name__ == "__main__":
    main()
