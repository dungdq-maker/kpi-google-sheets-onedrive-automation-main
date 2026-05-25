from __future__ import annotations

import argparse
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.graphics.shapes import Circle, Drawing, Line, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


WORK_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = WORK_DIR / "docs" / "huong_dan_van_hanh_code_github.pdf"
GITHUB_REPO_URL = "https://github.com/dungdq-maker/kpi-google-sheets-onedrive-automation-main"


def pick_font(candidates: list[Path]) -> Path | None:
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def register_fonts() -> tuple[str, str, str]:
    regular = pick_font(
        [
            Path("C:/Windows/Fonts/arial.ttf"),
            Path("C:/Windows/Fonts/tahoma.ttf"),
            Path("C:/Windows/Fonts/segoeui.ttf"),
        ]
    )
    bold = pick_font(
        [
            Path("C:/Windows/Fonts/arialbd.ttf"),
            Path("C:/Windows/Fonts/tahomabd.ttf"),
            Path("C:/Windows/Fonts/segoeuib.ttf"),
        ]
    )
    mono = pick_font(
        [
            Path("C:/Windows/Fonts/consola.ttf"),
            Path("C:/Windows/Fonts/cour.ttf"),
        ]
    )

    regular_name = "Helvetica"
    bold_name = "Helvetica-Bold"
    mono_name = "Courier"

    if regular:
        regular_name = "GuideRegular"
        pdfmetrics.registerFont(TTFont(regular_name, str(regular)))
    if bold:
        bold_name = "GuideBold"
        pdfmetrics.registerFont(TTFont(bold_name, str(bold)))
    elif regular:
        bold_name = regular_name
    if mono:
        mono_name = "GuideMono"
        pdfmetrics.registerFont(TTFont(mono_name, str(mono)))

    return regular_name, bold_name, mono_name


def p(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(escape(text).replace("\n", "<br/>"), style)


def cell(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(escape(text).replace("\n", "<br/>"), style)


def flow_diagram() -> Drawing:
    d = Drawing(520, 160)
    steps = [
        ("1", 20, 70, 70, 42, "#1D4ED8", "git clone"),
        ("2", 110, 70, 80, 42, "#0F766E", "edit code"),
        ("3", 210, 70, 80, 42, "#EA580C", "py -3 test"),
        ("4", 310, 70, 90, 42, "#7C3AED", "git commit"),
        ("5", 420, 70, 80, 42, "#DC2626", "git push"),
    ]
    for idx, x, y, w, h, fill, label in steps:
        d.add(Rect(x, y, w, h, rx=8, ry=8, fillColor=colors.HexColor(fill), strokeColor=colors.HexColor(fill)))
        d.add(String(x + 10, y + 24, label, fontName="Helvetica-Bold", fontSize=10, fillColor=colors.white))
    for x1, x2 in [(90, 108), (190, 208), (290, 308), (400, 418)]:
        d.add(Line(x1, 91, x2, 91, strokeColor=colors.HexColor("#64748B"), strokeWidth=2))
        d.add(Circle(x2, 91, 3, fillColor=colors.HexColor("#64748B"), strokeColor=colors.HexColor("#64748B")))
    d.add(String(20, 130, "Luồng vận hành code", fontName="Helvetica-Bold", fontSize=12, fillColor=colors.HexColor("#0F172A")))
    d.add(String(20, 112, "Dùng cho người duy trì code trước khi phát hành cho người dùng cuối.", fontName="Helvetica", fontSize=9, fillColor=colors.HexColor("#334155")))
    return d


def build_story(styles: dict[str, ParagraphStyle]) -> list:
    story: list = []

    story.append(Spacer(1, 18 * mm))
    story.append(p("HƯỚNG DẪN VẬN HÀNH CODE VÀ ĐĂNG LÊN GITHUB", styles["cover_title"]))
    story.append(Spacer(1, 4 * mm))
    story.append(p("Tài liệu riêng cho người duy trì code: clone, sửa, test, commit, push và phát hành cho người dùng cuối.", styles["cover_subtitle"]))
    story.append(Spacer(1, 8 * mm))
    story.append(flow_diagram())
    story.append(PageBreak())

    story.append(p("1. Mục tiêu của tài liệu", styles["h1"]))
    story.extend(
        [
            p("Tài liệu này không dành cho người dùng cuối. Nó dành cho người vận hành code, người sửa logic, và người chịu trách nhiệm đẩy bản mới lên GitHub.", styles["body"]),
            p("Quy trình chuẩn: lấy source từ GitHub -> sửa code -> test local -> commit -> push -> nếu cần thì tạo ZIP/Release cho người dùng cuối.", styles["body"]),
        ]
    )

    role_table = Table(
        [
            [cell("Vai trò", styles["table"]), cell("Công việc", styles["table"])],
            [cell("Người vận hành code", styles["table"]), cell("Sửa code, test, commit, push, phát hành bản mới.", styles["table"])],
            [cell("Người dùng cuối", styles["table"]), cell("Chỉ tải ZIP hoặc clone bản phát hành để chạy.", styles["table"])],
        ],
        colWidths=[45 * mm, 135 * mm],
    )
    role_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4ED8")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(role_table)
    story.append(PageBreak())

    story.append(p("2. Lấy source lần đầu", styles["h1"]))
    story.append(p("Lệnh khuyến nghị cho người vận hành code là clone trực tiếp từ GitHub.", styles["body"]))
    story.append(p("git clone https://github.com/dungdq-maker/kpi-google-sheets-onedrive-automation-main.git", styles["code"]))
    story.append(Spacer(1, 2 * mm))
    story.extend(
        [
            p("Sau khi clone:", styles["subhead"]),
            p("cd kpi-google-sheets-onedrive-automation-main", styles["code"]),
            p("code .", styles["code"]),
        ]
    )
    story.append(Spacer(1, 2 * mm))
    story.append(p("Nếu cần mở bằng file explorer thay vì terminal, dùng Start Menu để mở Visual Studio Code rồi Open Folder vào thư mục vừa clone.", styles["note"]))

    story.append(PageBreak())

    story.append(p("3. Cài môi trường và chạy thử local", styles["h1"]))
    story.append(p("Trước khi sửa hay push, hãy cài dependencies và chạy thử code để biết bản hiện tại có lỗi gì.", styles["body"]))
    story.append(p("py -3 -m pip install -r requirements.txt", styles["code"]))
    story.append(Spacer(1, 2 * mm))
    story.append(p("py -3 automate_kpi.py --download --open", styles["code"]))
    story.append(Spacer(1, 2 * mm))
    story.extend(
        [
            p("Các lệnh kiểm tra nhanh:", styles["subhead"]),
            p("py -3 -m py_compile automate_kpi.py", styles["code"]),
            p("py -3 -m py_compile generate_huong_dan_pdf.py", styles["code"]),
            p("git status", styles["code"]),
        ]
    )
    story.append(p("Nếu có lỗi, sửa code rồi chạy lại trước khi commit.", styles["note"]))
    story.append(PageBreak())

    story.append(p("4. Quy trình sửa code", styles["h1"]))
    edit_table = Table(
        [
            [cell("Bước", styles["table"]), cell("Câu lệnh / thao tác", styles["table"])],
            [cell("1", styles["table"]), cell("git pull origin main", styles["table"])],
            [cell("2", styles["table"]), cell("Mở project trong VS Code và sửa file cần thiết.", styles["table"])],
            [cell("3", styles["table"]), cell("Chạy lại `py -3 automate_kpi.py --download --open` hoặc lệnh test tương ứng.", styles["table"])],
            [cell("4", styles["table"]), cell("Kiểm tra output trong `data/output/final` và log trong `logs`.", styles["table"])],
        ],
        colWidths=[22 * mm, 158 * mm],
    )
    edit_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F766E")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(edit_table)
    story.append(Spacer(1, 4 * mm))
    story.append(p("Nguyên tắc: không push code khi chưa chạy được local ít nhất một lần.", styles["note"]))

    story.append(PageBreak())

    story.append(p("5. Commit và push lên GitHub", styles["h1"]))
    story.append(p("Khi bản sửa đã ổn, dùng đúng chuỗi lệnh git sau để đưa code lên GitHub.", styles["body"]))
    story.append(p("git status", styles["code"]))
    story.append(p("git add .", styles["code"]))
    story.append(p('git commit -m "Update automation and guide"', styles["code"]))
    story.append(p("git push origin main", styles["code"]))
    story.extend(
        [
            p("Lệnh thay thế thường dùng:", styles["subhead"]),
            p("git branch", styles["code"]),
            p("git remote -v", styles["code"]),
            p("git log --oneline -5", styles["code"]),
        ]
    )
    story.append(p("Nếu repo dùng nhánh khác `main`, thay `main` bằng nhánh thật của dự án.", styles["note"]))

    story.append(PageBreak())

    story.append(p("6. Phát hành cho người dùng cuối", styles["h1"]))
    release_table = Table(
        [
            [cell("Cách phát hành", styles["table"]), cell("Khi nào dùng", styles["table"])],
            [cell("ZIP từ GitHub", styles["table"]), cell("Khi người dùng cuối chỉ cần tải source về máy.", styles["table"])],
            [cell("GitHub Release", styles["table"]), cell("Khi muốn phát bản ổn định, có version tag rõ ràng.", styles["table"])],
            [cell("README / PDF", styles["table"]), cell("Khi cần hướng dẫn cách cài và chạy.", styles["table"])],
        ],
        colWidths=[55 * mm, 125 * mm],
    )
    release_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4ED8")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(release_table)
    story.append(Spacer(1, 4 * mm))
    story.append(p(f"Repo chính: {GITHUB_REPO_URL}", styles["link"]))
    story.append(p("Nếu muốn tạo ZIP thủ công từ git, có thể dùng `git archive --format zip --output release.zip HEAD`.", styles["note"]))

    story.append(PageBreak())

    story.append(p("7. Kiểm tra trước khi bàn giao", styles["h1"]))
    checklist = Table(
        [
            [cell("Checklist", styles["table"]), cell("Đã xong?", styles["table"])],
            [cell("Chạy `py -3 -m py_compile automate_kpi.py`", styles["table"]), cell("☐", styles["table"])],
            [cell("Chạy `py -3 automate_kpi.py --download --open`", styles["table"]), cell("☐", styles["table"])],
            [cell("Commit và push lên GitHub", styles["table"]), cell("☐", styles["table"])],
            [cell("Cập nhật PDF hướng dẫn nếu có thay đổi lớn", styles["table"]), cell("☐", styles["table"])],
        ],
        colWidths=[155 * mm, 25 * mm],
    )
    checklist.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7C3AED")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FAF5FF")),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(checklist)
    story.append(Spacer(1, 4 * mm))
    story.append(p("Tài liệu này để người vận hành code làm việc nhanh hơn và ít sai bước hơn khi phát hành bản mới.", styles["note"]))

    story.append(PageBreak())

    story.append(p("8. Lỗi thường gặp", styles["h1"]))
    errors = Table(
        [
            [cell("Vấn đề", styles["table"]), cell("Cách xử lý", styles["table"])],
            [cell("Không chạy được `py`", styles["table"]), cell("Cài Python Launcher hoặc dùng đúng đường dẫn `py -3` trên Windows.", styles["table"])],
            [cell("Push bị reject", styles["table"]), cell("Chạy `git pull --rebase origin main` rồi push lại.", styles["table"])],
            [cell("Local khác GitHub", styles["table"]), cell("Kiểm tra `git status` và `git log` trước khi commit mới.", styles["table"])],
            [cell("Quên update PDF", styles["table"]), cell("Chạy lại script tạo PDF operator guide sau khi thay đổi quy trình.", styles["table"])],
        ],
        colWidths=[55 * mm, 125 * mm],
    )
    errors.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#B91C1C")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFF7ED")),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(errors)

    return story


def build_pdf(output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    regular_font, bold_font, mono_font = register_fonts()
    base = getSampleStyleSheet()
    styles = {
        "cover_title": ParagraphStyle(
            "cover_title",
            parent=base["Title"],
            fontName=bold_font,
            fontSize=22,
            leading=26,
            textColor=colors.HexColor("#0F172A"),
            alignment=TA_CENTER,
        ),
        "cover_subtitle": ParagraphStyle(
            "cover_subtitle",
            parent=base["Heading2"],
            fontName=regular_font,
            fontSize=12,
            leading=16,
            textColor=colors.HexColor("#334155"),
            alignment=TA_CENTER,
        ),
        "h1": ParagraphStyle(
            "h1",
            parent=base["Heading1"],
            fontName=bold_font,
            fontSize=17,
            leading=21,
            textColor=colors.HexColor("#0F172A"),
        ),
        "subhead": ParagraphStyle(
            "subhead",
            parent=base["Heading3"],
            fontName=bold_font,
            fontSize=11.5,
            leading=14,
            textColor=colors.HexColor("#1D4ED8"),
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["BodyText"],
            fontName=regular_font,
            fontSize=10.5,
            leading=15,
            textColor=colors.HexColor("#111827"),
        ),
        "note": ParagraphStyle(
            "note",
            parent=base["BodyText"],
            fontName=regular_font,
            fontSize=9.5,
            leading=13,
            textColor=colors.HexColor("#7C2D12"),
            backColor=colors.HexColor("#FFFBEB"),
            borderPadding=6,
            borderColor=colors.HexColor("#FBBF24"),
            borderWidth=0.5,
            borderRadius=4,
        ),
        "link": ParagraphStyle(
            "link",
            parent=base["BodyText"],
            fontName=bold_font,
            fontSize=10.5,
            leading=15,
            textColor=colors.HexColor("#1D4ED8"),
        ),
        "code": ParagraphStyle(
            "code",
            parent=base["Code"],
            fontName=mono_font,
            fontSize=10,
            leading=13,
            textColor=colors.HexColor("#0F172A"),
            backColor=colors.HexColor("#E2E8F0"),
        ),
        "table": ParagraphStyle(
            "table",
            parent=base["BodyText"],
            fontName=regular_font,
            fontSize=9.5,
            leading=13,
        ),
    }

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Huong dan van hanh code va GitHub",
        author="Codex",
    )

    story = build_story(styles)

    def on_page(canvas, _doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#F7F4EF"))
        canvas.rect(0, 0, A4[0], A4[1], stroke=0, fill=1)
        canvas.setFillColor(colors.HexColor("#1D4ED8"))
        canvas.rect(0, A4[1] - 14 * mm, A4[0], 14 * mm, stroke=0, fill=1)
        canvas.setFillColor(colors.white)
        canvas.setFont(bold_font, 9)
        canvas.drawString(15 * mm, A4[1] - 9 * mm, "Huong dan van hanh code va GitHub")
        canvas.setFont(regular_font, 8)
        canvas.setFillColor(colors.HexColor("#4B5563"))
        canvas.drawRightString(A4[0] - 15 * mm, 8 * mm, f"Trang {canvas.getPageNumber()}")
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Sinh PDF huong dan van hanh code va GitHub")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Duong dan PDF can tao")
    args = parser.parse_args()
    out = build_pdf(args.output)
    print(f"Da tao PDF: {out}")


if __name__ == "__main__":
    main()
