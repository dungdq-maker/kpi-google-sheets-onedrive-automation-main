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
DEFAULT_OUTPUT = WORK_DIR / "docs" / "thuyet_trinh_bod_manager.pdf"


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
        regular_name = "BodRegular"
        pdfmetrics.registerFont(TTFont(regular_name, str(regular)))
    if bold:
        bold_name = "BodBold"
        pdfmetrics.registerFont(TTFont(bold_name, str(bold)))
    elif regular:
        bold_name = regular_name
    if mono:
        mono_name = "BodMono"
        pdfmetrics.registerFont(TTFont(mono_name, str(mono)))

    return regular_name, bold_name, mono_name


def p(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(escape(text).replace("\n", "<br/>"), style)


def build_flow(font_bold: str) -> Drawing:
    d = Drawing(520, 120)
    steps = [
        ("Nguồn", 20, 40, 76, 36, "#1E3A8A"),
        ("Staging", 112, 40, 76, 36, "#334155"),
        ("Checkpoint", 204, 40, 76, 36, "#A16207"),
        ("Approval", 296, 40, 76, 36, "#0F766E"),
        ("Output", 388, 40, 92, 36, "#111827"),
    ]
    for label, x, y, w, h, fill in steps:
        d.add(Rect(x, y, w, h, rx=8, ry=8, fillColor=colors.HexColor(fill), strokeColor=colors.HexColor(fill)))
        d.add(String(x + 8, y + 20, label, fontName=font_bold, fontSize=10, fillColor=colors.white))
    for x1, x2 in [(96, 110), (188, 202), (280, 294), (372, 386)]:
        d.add(Line(x1, 58, x2, 58, strokeColor=colors.HexColor("#64748B"), strokeWidth=2))
        d.add(Circle(x2, 58, 3, fillColor=colors.HexColor("#64748B"), strokeColor=colors.HexColor("#64748B")))
    d.add(String(20, 100, "Flow tổng thể", fontName=font_bold, fontSize=12, fillColor=colors.HexColor("#0F172A")))
    return d


def make_box(rows: list[list[Paragraph]], col_widths: list[float], header_fill: str, body_fill: str, font_size: float = 9.0) -> Table:
    table = Table(rows, colWidths=col_widths)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_fill)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor(body_fill)),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    return table


def build_story(styles: dict[str, ParagraphStyle]) -> list:
    story: list = []

    story.append(Spacer(1, 14 * mm))
    story.append(p("TỪ THAO TÁC TAY ĐẾN VẬN HÀNH CHUẨN HÓA", styles["cover_title"]))
    story.append(Spacer(1, 4 * mm))
    story.append(
        p(
            "Script thuyết trình 5-7 phút cho BOD / Manager: flow bài toán, lợi ích khi giải, và tác dụng của tự động hóa.",
            styles["cover_subtitle"],
        )
    )
    story.append(Spacer(1, 7 * mm))
    story.append(build_flow(styles["cover_title"].fontName))
    story.append(Spacer(1, 4 * mm))

    cover_stats = make_box(
        [
            [
                p("Nếu làm tay: 8 giờ / Associate / chu kỳ", styles["table"]),
                p("Nếu làm tay: 4 giờ / Senior / chu kỳ", styles["table"]),
                p("Mục tiêu: chuyển sang kiểm tra + phê duyệt", styles["table"]),
            ]
        ],
        [58 * mm, 54 * mm, 64 * mm],
        "#0F172A",
        "#F8FAFC",
        font_size=9.2,
    )
    story.append(cover_stats)
    story.append(Spacer(1, 5 * mm))
    story.append(
        p(
            "Thông điệp ngắn: dự án này không chỉ để chạy nhanh hơn, mà để ra quyết định nhanh hơn, ít sai hơn, và dễ bàn giao hơn.",
            styles["note"],
        )
    )
    story.append(PageBreak())

    story.append(p("1. Vì sao bài toán ra đời", styles["h1"]))
    story.extend(
        [
            p(
                "Bài toán xuất hiện vì dữ liệu đang nằm rải rác ở nhiều nơi: Google Sheets, Excel, OneDrive và các workbook output khác nhau. Làm tay nghĩa là copy / paste, đối chiếu, sửa lại, rồi chờ người khác kiểm tra. Cách đó chậm, dễ lệch và rất khó scale.",
                styles["body"],
            ),
            p(
                "Nếu một chu kỳ xử lý tay tiêu tốn 8 giờ của Associate hoặc 4 giờ của Senior, thì chi phí không chỉ là thời gian. Nó còn là độ trễ ra quyết định, sai sót do thao tác thủ công, và rủi ro khi bàn giao cho người khác.",
                styles["body"],
            ),
            p(
                "Mục tiêu tối thượng của dự án là đưa dữ liệu về một luồng chuẩn, để tính được chi phí vốn hóa một cách nhất quán, có kiểm soát và có thể audit lại.",
                styles["body"],
            ),
        ]
    )

    problem_table = make_box(
        [
            [p("Flow", styles["table"]), p("Câu hỏi cần nghĩ", styles["table"]), p("Ý nghĩa", styles["table"])],
            [p("Source", styles["table"]), p("Dữ liệu đầu vào có đúng và đủ không?", styles["table"]), p("Khi source sai thì mọi bước sau đều sai.", styles["table"])],
            [p("Staging", styles["table"]), p("Dữ liệu đã được clean và chuẩn hóa chưa?", styles["table"]), p("Đây là lớp trung gian để gom, lọc và loại lỗi trước khi đổ vào output.", styles["table"])],
            [p("Checkpoint / Approval", styles["table"]), p("Dòng nào cần người dùng duyệt?", styles["table"]), p("Cho phép kiểm soát các ngoại lệ thay vì sửa tay trong bóng tối.", styles["table"])],
            [p("Output", styles["table"]), p("Kết quả cuối cùng có đáng tin để dùng không?", styles["table"]), p("Đây là workbook dùng để báo cáo, chia sẻ và bàn giao.", styles["table"])],
        ],
        [26 * mm, 72 * mm, 80 * mm],
        "#1E3A8A",
        "#F8FAFC",
        font_size=8.8,
    )
    story.append(problem_table)
    story.append(Spacer(1, 4 * mm))
    story.append(
        p(
            "Khi nói phần này, hãy nhấn mạnh rằng tự động hóa không chỉ để \"nhanh\", mà để biến một quy trình rời rạc thành một quy trình có thể lặp lại và kiểm soát.",
            styles["note"],
        )
    )
    story.append(PageBreak())

    story.append(p("2. Lợi ích khi tự động hóa", styles["h1"]))
    benefit_table = make_box(
        [
            [p("Lợi ích", styles["table"]), p("Tác dụng thực tế", styles["table"])],
            [p("Tiết kiệm thời gian", styles["table"]), p("Nếu làm tay mất 8h với Associate hoặc 4h với Senior, automation chuyển phần nặng sang kiểm tra và phê duyệt, giúp tiết kiệm gần như một ngày công Associate hoặc nửa ngày Senior cho mỗi chu kỳ.", styles["table"])],
            [p("Giảm sai sót", styles["table"]), p("Không còn phụ thuộc vào thao tác copy/paste, đặt công thức thủ công hay sửa lẻ tẻ nhiều nơi.", styles["table"])],
            [p("Dễ audit và bàn giao", styles["table"]), p("Có checkpoint, result sheet và log để truy ngược từng quyết định.", styles["table"])],
            [p("Dễ scale", styles["table"]), p("Khi khối lượng tăng, hệ thống vẫn chạy theo cùng một logic thay vì tăng người làm tay.", styles["table"])],
        ],
        [40 * mm, 138 * mm],
        "#0F766E",
        "#F8FAFC",
        font_size=8.8,
    )
    story.append(benefit_table)
    story.append(Spacer(1, 5 * mm))

    story.append(p("3. Ví dụ IT và Media: khi nào cần checkpoint và approval", styles["h1"]))
    story.extend(
        [
            p(
                "Với IT, dữ liệu đi từ Timesheet IT sang Chi phí nhân sự IT, rồi tiếp tục qua các bước kiểm tra project master và downstream. Nói ngắn gọn: mục tiêu không chỉ là kéo được dữ liệu vào file, mà là bảo đảm project và nhân sự được map đúng trước khi tính chi phí.",
                styles["body"],
            ),
            p(
                "Các điểm cần nhắc tên khi thuyết trình là `Check_IT_CPNS`, `IT_New_Project_Master`, `Check_IT_Downstream`, rồi các result sheet như `IT_Approval_Result`, `Project_Master_Approval_Result` và `Downstream_Approval_Result`.",
                styles["body"],
            ),
            p(
                "Với Media, luồng thường đi từ Timesheet Media và Data media ACCA+CFA+CMA. Ở đây checkpoint giúp kiểm tra mapping BU, project, task, tháng và người làm task trước khi dữ liệu đi vào output.",
                styles["body"],
            ),
            p(
                "Các sheet cần nhớ là `Check_Media_Timesheet` và `Media_Approval_Result`. Nếu dữ liệu Media sai, vấn đề thường nằm ở mapping nguồn hoặc danh mục, không phải ở bước output cuối.",
                styles["body"],
            ),
        ]
    )
    it_media_table = make_box(
        [
            [p("Mảng", styles["table"]), p("Điểm cần kiểm tra", styles["table"]), p("Kết quả mong muốn", styles["table"])],
            [p("IT", styles["table"]), p("Timesheet IT, Check_IT_CPNS, IT_New_Project_Master, Check_IT_Downstream", styles["table"]), p("Project và chi phí nhân sự IT được map đúng, output có result sheet rõ ràng.", styles["table"])],
            [p("Media", styles["table"]), p("Timesheet Media, Data media ACCA+CFA+CMA, Check_Media_Timesheet", styles["table"]), p("Task / tháng / người làm task khớp, dữ liệu đi vào output sạch hơn.", styles["table"])],
        ],
        [24 * mm, 82 * mm, 72 * mm],
        "#1E3A8A",
        "#F8FAFC",
        font_size=8.8,
    )
    story.append(it_media_table)
    story.append(Spacer(1, 5 * mm))

    story.append(p("4. Ví dụ SX: vì sao phải clean staging trước khi gộp output", styles["h1"]))
    story.extend(
        [
            p(
                "Với SX, vấn đề thường không nằm ở bước cuối mà nằm ở dữ liệu đầu vào: MNV thiếu, mapping lệch, hoặc downstream bị kéo sai. Vì vậy, cách xử lý đúng là clean data ở staging trước, rồi mới gộp vào file output.",
                styles["body"],
            ),
            p(
                "Đây là lý do code có checkpoint riêng cho `checkpoint data SX` và `Check_SX_Downstream`. Một cái để nhìn lỗi ở nguồn dữ liệu SX, một cái để nhìn lỗi lan sang `SX_Allocation_Build`, `Timesheet SX`, `4.1 Chi phí nhân sự SX` và `3.Vốn hóa`.",
                styles["body"],
            ),
            p(
                "Nói ngắn gọn cho BOD: tự động hóa giúp chúng ta không chỉ tạo file nhanh hơn, mà còn chuẩn hóa quyết định, giảm rủi ro dữ liệu, và đảm bảo chi phí vốn hóa được tính trên bộ dữ liệu sạch hơn.",
                styles["body"],
            ),
        ]
    )

    story.append(Spacer(1, 4 * mm))
    story.append(p("5. Khi có dự án mới hoặc nhân viên mới thì phải chỉnh gì?", styles["h1"]))
    changes_table = make_box(
        [
            [p("Tình huống", styles["table"]), p("Sheet thường cần xem / chỉnh", styles["table"]), p("Lý do", styles["table"])],
            [p("Dự án mới", styles["table"]), p("1.Danh mục dự án, IT_New_Project_Master, Check_IT_Downstream, 3.Vốn hóa", styles["table"]), p("Project mới phải được map vào danh mục và đi qua downstream đúng.", styles["table"])],
            [p("Nhân viên mới", styles["table"]), p("Mã nhân viên, Check_Payroll, Timesheet IT / Media / SX, checkpoint tương ứng", styles["table"]), p("Nếu thiếu MNV hoặc map sai thì output sẽ lệch ngay ở bước đầu.", styles["table"])],
            [p("Lương mới", styles["table"]), p("Lương nhân viên full time, Lương nhân viên part time, Check_Payroll", styles["table"]), p("Đây là nguồn để đồng bộ payroll vào workbook output.", styles["table"])],
        ],
        [28 * mm, 96 * mm, 54 * mm],
        "#0F766E",
        "#F8FAFC",
        font_size=8.7,
    )
    story.append(changes_table)
    story.append(Spacer(1, 4 * mm))
    story.append(
        p(
            "Struggle thực tế thường không phải là thiếu code, mà là thay đổi ở nguồn: tên dự án đổi, nhân viên mới chưa map, hoặc sheet nguồn bị đổi tên. Khi đó hệ thống cần kiểm tra lại source, staging và checkpoint chứ không chỉ sửa output cuối.",
            styles["note"],
        )
    )

    closing_box = make_box(
        [[p("Chốt thông điệp: workflow này biến một công việc nặng thao tác tay thành một quy trình có kiểm soát, có thể mở rộng và có thể bàn giao.", styles["table"])]],
        [178 * mm],
        "#111827",
        "#F8FAFC",
        font_size=9.2,
    )
    story.append(Spacer(1, 2 * mm))
    story.append(closing_box)
    return story


def build_pdf(output_path: Path) -> Path:
    regular_font, bold_font, mono_font = register_fonts()
    base = getSampleStyleSheet()
    styles = {
        "cover_title": ParagraphStyle(
            "cover_title",
            parent=base["Title"],
            fontName=bold_font,
            fontSize=21,
            leading=25,
            textColor=colors.HexColor("#0F172A"),
            alignment=TA_CENTER,
        ),
        "cover_subtitle": ParagraphStyle(
            "cover_subtitle",
            parent=base["BodyText"],
            fontName=regular_font,
            fontSize=10.5,
            leading=14,
            textColor=colors.HexColor("#334155"),
            alignment=TA_CENTER,
        ),
        "h1": ParagraphStyle(
            "h1",
            parent=base["Heading1"],
            fontName=bold_font,
            fontSize=15,
            leading=18,
            textColor=colors.HexColor("#0F172A"),
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["BodyText"],
            fontName=regular_font,
            fontSize=10,
            leading=13,
            spaceAfter=5,
            textColor=colors.HexColor("#0F172A"),
        ),
        "note": ParagraphStyle(
            "note",
            parent=base["BodyText"],
            fontName=regular_font,
            fontSize=9,
            leading=12,
            textColor=colors.HexColor("#475569"),
        ),
        "table": ParagraphStyle(
            "table",
            parent=base["BodyText"],
            fontName=regular_font,
            fontSize=8.8,
            leading=11,
            textColor=colors.HexColor("#0F172A"),
        ),
    }
    styles["code"] = ParagraphStyle(
        "code",
        parent=base["Code"],
        fontName=mono_font,
        fontSize=9.2,
        leading=12,
        textColor=colors.HexColor("#0F172A"),
        backColor=colors.HexColor("#F8FAFC"),
        borderPadding=4,
        leftIndent=4,
    )

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )
    story = build_story(styles)
    doc.build(story)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate BOD presentation PDF.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output PDF path")
    args = parser.parse_args()
    out = build_pdf(Path(args.output))
    print(f"Da tao PDF: {out}")


if __name__ == "__main__":
    main()
