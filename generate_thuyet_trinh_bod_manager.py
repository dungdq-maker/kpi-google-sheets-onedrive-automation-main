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
        regular_name = "WorkflowRegular"
        pdfmetrics.registerFont(TTFont(regular_name, str(regular)))
    if bold:
        bold_name = "WorkflowBold"
        pdfmetrics.registerFont(TTFont(bold_name, str(bold)))
    elif regular:
        bold_name = regular_name
    if mono:
        mono_name = "WorkflowMono"
        pdfmetrics.registerFont(TTFont(mono_name, str(mono)))

    return regular_name, bold_name, mono_name


def p(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(escape(text).replace("\n", "<br/>"), style)


def build_flow(font_bold: str) -> Drawing:
    d = Drawing(520, 120)
    steps = [
        ("1", 20, 40, 76, 36, "#1D4ED8", "Nguồn"),
        ("2", 112, 40, 76, 36, "#0F766E", "Xử lý"),
        ("3", 204, 40, 76, 36, "#EA580C", "Kiểm tra"),
        ("4", 296, 40, 76, 36, "#7C3AED", "Phê duyệt"),
        ("5", 388, 40, 92, 36, "#DC2626", "Kết quả"),
    ]
    for _, x, y, w, h, fill, label in steps:
        d.add(Rect(x, y, w, h, rx=8, ry=8, fillColor=colors.HexColor(fill), strokeColor=colors.HexColor(fill)))
        d.add(String(x + 8, y + 20, label, fontName=font_bold, fontSize=10, fillColor=colors.white))
    for x1, x2 in [(96, 110), (188, 202), (280, 294), (372, 386)]:
        d.add(Line(x1, 58, x2, 58, strokeColor=colors.HexColor("#64748B"), strokeWidth=2))
        d.add(Circle(x2, 58, 3, fillColor=colors.HexColor("#64748B"), strokeColor=colors.HexColor("#64748B")))
    d.add(String(20, 100, "Workflow tổng thể", fontName=font_bold, fontSize=12, fillColor=colors.HexColor("#0F172A")))
    return d


def build_story(styles: dict[str, ParagraphStyle]) -> list:
    story: list = []

    story.append(Spacer(1, 16 * mm))
    story.append(p("HƯỚNG DẪN WORKFLOW VÀ CÁCH SUY NGHĨ VẬN HÀNH", styles["cover_title"]))
    story.append(Spacer(1, 4 * mm))
    story.append(p("Tài liệu ngắn gọn, tập trung vào cách nhìn hệ thống: đầu vào -> xử lý -> checkpoint -> approval -> kết quả -> phát hành.", styles["cover_subtitle"]))
    story.append(Spacer(1, 8 * mm))
    story.append(build_flow(styles["cover_title"].fontName))
    story.append(PageBreak())

    story.append(p("1. Cách nhìn workflow", styles["h1"]))
    story.extend(
        [
            p("Hệ thống này không nên được hiểu là một file Python đơn lẻ. Nó là một chuỗi xử lý có ý nghĩa rõ ràng: lấy dữ liệu gốc, chuẩn hóa, kiểm tra, xin phê duyệt, ghi kết quả, và phát hành bản mới.", styles["body"]),
            p("Khi đọc code hoặc chạy automation, trước tiên phải trả lời 3 câu hỏi: dữ liệu nào là source of truth, sheet nào là checkpoint, và trạng thái nào là trạng thái cuối cùng được tin cậy.", styles["body"]),
        ]
    )

    thinking_table = Table(
        [
            [p("Giai đoạn", styles["table"]), p("Câu hỏi cần tư duy", styles["table"]), p("Đầu ra mong đợi", styles["table"])],
            [p("Nguồn", styles["table"]), p("Dữ liệu lấy từ đâu và có được cập nhật chưa?", styles["table"]), p("File raw / Google Sheet / OneDrive hợp lệ.", styles["table"])],
            [p("Chuyển đổi", styles["table"]), p("Code đang đối chiếu / ghép / tính gì?", styles["table"]), p("Staging đã chuẩn hóa.", styles["table"])],
            [p("Kiểm tra", styles["table"]), p("Dòng nào cần người dùng xem lại trước khi duyệt?", styles["table"]), p("Sheet Check_* hoặc checkpoint data SX.", styles["table"])],
            [p("Phê duyệt", styles["table"]), p("Dòng nào được phê duyệt để ghi sang workbook mới?", styles["table"]), p("Result sheet có applied / skipped / failed.", styles["table"])],
            [p("Phát hành", styles["table"]), p("Bản code mới đã lên GitHub chưa?", styles["table"]), p("Repo đồng bộ và user có bản mới để chạy.", styles["table"])],
        ],
        colWidths=[30 * mm, 82 * mm, 68 * mm],
    )
    thinking_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4ED8")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.2),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(thinking_table)
    story.append(Spacer(1, 4 * mm))
    story.append(p("2. Cách suy nghĩ khi vận hành", styles["h1"]))
    story.extend(
        [
            p("1. Xác định source of truth trước. Nếu user hỏi 'dữ liệu nào đúng', phải biết workbook nào là nguồn chính, workbook nào chỉ là kết quả.", styles["body"]),
            p("2. Không sửa trực tiếp kết quả nếu có thể sửa ở nguồn. Càng gần source of truth càng tốt, vì như vậy con đường debug ngắn hơn.", styles["body"]),
            p("3. Mỗi approval phải có checkpoint. Checkpoint là nơi chứng minh rằng hệ thống đã phát hiện vấn đề và cho phép người dùng quyết định.", styles["body"]),
            p("4. Approval không phải là thay đổi ngay lập tức. Quy trình đúng là đánh dấu Apply?, chạy lại với --approval-file, và xem result sheet.", styles["body"]),
            p("5. Nếu code thay đổi, user phải được lấy bản mới. Nếu không, giữa repo và tài liệu sẽ bị lệch version.", styles["body"]),
        ]
    )

    ops_table = Table(
        [
            [p("Tình huống", styles["table"]), p("Cách nghĩ", styles["table"]), p("Hành động", styles["table"])],
            [p("IT / Media", styles["table"]), p("Kiểm tra nguồn -> checkpoint -> approve -> xem kết quả.", styles["table"]), p("Dùng workbook output mới nhất và sheet Huong_dan_Approval.", styles["table"])],
            [p("SX", styles["table"]), p("Kiểm tra merge_SX.py -> checkpoint data SX / Check_SX_Downstream -> approve.", styles["table"]), p("Xử lý dựa vào missing MNV, SX_Allocation_Build và downstream action.", styles["table"])],
            [p("Payroll", styles["table"]), p("Nhập liệu từ OneDrive và luôn đối chiếu Check_Payroll.", styles["table"]), p("Kiểm tra Rows copied và Duplicate Month+MNV.", styles["table"])],
            [p("GitHub", styles["table"]), p("Bản phát hành code phải đi qua publish và commit rõ ràng.", styles["table"]), p("git add, git commit, git push hoặc script publish_to_github.ps1.", styles["table"])],
        ],
        colWidths=[28 * mm, 88 * mm, 64 * mm],
    )
    ops_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F766E")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.1),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(ops_table)
    story.append(PageBreak())

    story.append(p("3. Thứ tự ưu tiên khi gặp vấn đề", styles["h1"]))
    story.extend(
        [
            p("1. Kiểm tra log trước. Log cho biết biến nào thiếu, file nào không tải được, sheet nào không tồn tại.", styles["body"]),
            p("2. Mở workbook checkpoint / result trước khi mở code. Nếu sheet dữ liệu đã sai, sửa cơ sở dữ liệu và workbook trước.", styles["body"]),
            p("3. Chỉ sửa code khi vấn đề là logic. Nếu vấn đề do dữ liệu nguồn, sửa code không giải quyết được gốc rễ.", styles["body"]),
            p("4. Sau khi sửa code, test local bằng py -3, rồi mới publish lên GitHub.", styles["body"]),
            p("5. Khi thay đổi có khả năng ảnh hưởng user, cập nhật PDF / hướng dẫn cùng lúc để không lệch tài liệu.", styles["body"]),
        ]
    )

    sequence_table = Table(
        [
            [p("Trước", styles["table"]), p("Trong lúc", styles["table"]), p("Sau đó", styles["table"])],
            [p("Xác định nguồn", styles["table"]), p("Chọn đúng checkpoint / approval", styles["table"]), p("Xem result sheet và gửi bản mới", styles["table"])],
            [p("Đọc log", styles["table"]), p("Sửa phần liên quan nhất", styles["table"]), p("Push lên GitHub", styles["table"])],
            [p("Test local", styles["table"]), p("Không bỏ qua dữ liệu cảnh báo", styles["table"]), p("Cập nhật tài liệu", styles["table"])],
        ],
        colWidths=[56 * mm, 66 * mm, 60 * mm],
    )
    sequence_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7C3AED")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.1),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(sequence_table)
    story.append(Spacer(1, 4 * mm))
    story.extend(
        [
            p("Kết luận: workflow được thiết kế để có thể debug từng lớp. Khi gặp lỗi, đừng nhảy thẳng vào code mà hãy xác định đang đứng ở lớp nào: source, staging, checkpoint, approval, hay publish.", styles["note"]),
        ]
    )
    return story


def build_bod_story(styles: dict[str, ParagraphStyle]) -> list:
    story: list = []

    story.append(Spacer(1, 16 * mm))
    story.append(p("WORKFLOW DỰ ÁN: CÁCH GIẢI THÍCH CHO BOD VÀ MANAGER", styles["cover_title"]))
    story.append(Spacer(1, 4 * mm))
    story.append(p("Tài liệu này đi theo 4 bước: flow bài toán -> câu hỏi cần tư duy -> cách xử lý vấn đề -> đầu ra.", styles["cover_subtitle"]))
    story.append(Spacer(1, 8 * mm))
    story.append(build_flow(styles["cover_title"].fontName))
    story.append(PageBreak())

    story.append(p("1. Flow bài toán", styles["h1"]))
    story.extend(
        [
            p(
                "Dự án không nên được nhìn như một file Python đơn lẻ. Đây là một chuỗi vận hành: lấy dữ liệu gốc, làm sạch ở staging, kiểm tra bằng checkpoint, duyệt bằng approval, rồi ghi ra output và phát hành bản mới.",
                styles["body"],
            ),
            p(
                "Khi trình bày cho BOD hoặc manager, nên nói theo logic vận hành: bài toán là gì, cần hỏi gì để ra quyết định, xử lý vấn đề ở lớp nào, và đầu ra cuối cùng là gì.",
                styles["body"],
            ),
        ]
    )

    flow_table = Table(
        [
            [p("Mảng", styles["table"]), p("Câu hỏi cần tư duy", styles["table"]), p("Cách xử lý vấn đề", styles["table"]), p("Đầu ra", styles["table"])],
            [
                p("IT / Media", styles["table"]),
                p("Dữ liệu nguồn nào đúng và dòng nào lệch?", styles["table"]),
                p("So sánh source với staging, dùng checkpoint để chỉ ra dòng cần duyệt, rồi chạy lại approval.", styles["table"]),
                p("IT_Approval_Result, Media_Approval_Result và workbook final.", styles["table"]),
            ],
            [
                p("SX", styles["table"]),
                p("Dữ liệu SX đã sạch chưa, MNV nào thiếu, downstream nào bị ảnh hưởng?", styles["table"]),
                p("Clean data ở staging trước khi gộp vào file output; kiểm tra checkpoint data SX và Check_SX_Downstream trước khi duyệt.", styles["table"]),
                p("SX_Approval_Result, SX_Downstream_Approval_Result, SX_Allocation_Build, Timesheet SX, 4.1 Chi phí nhân sự SX, 3.Vốn hóa.", styles["table"]),
            ],
            [
                p("Payroll", styles["table"]),
                p("File lương nào là bản mới nhất, có trùng Month+MNV không?", styles["table"]),
                p("Đồng bộ từ OneDrive, đối chiếu Check_Payroll, chỉ giữ bản nguồn hợp lệ trước khi ghi output.", styles["table"]),
                p("Check_Payroll trạng thái OK và số dòng copied chính xác.", styles["table"]),
            ],
            [
                p("GitHub", styles["table"]),
                p("Người dùng đang chạy bản code nào và bản phát hành đã được đẩy lên chưa?", styles["table"]),
                p("Sửa local -> test -> publish_to_github.ps1 -> user tải bản mới hoặc refresh repo.", styles["table"]),
                p("Repo đồng bộ, user chạy được bản mới.", styles["table"]),
            ],
        ],
        colWidths=[24 * mm, 65 * mm, 61 * mm, 34 * mm],
    )
    flow_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4ED8")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 8.8),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(flow_table)
    story.append(Spacer(1, 4 * mm))
    story.append(p("2. Cách tư duy khi xử lý vấn đề", styles["h1"]))
    story.extend(
        [
            p("1. Nếu vấn đề nằm ở nguồn, sửa nguồn trước. Nếu vấn đề nằm ở dữ liệu trung gian, sửa ở staging trước khi đổ vào output.", styles["body"]),
            p("2. Nếu vấn đề là quyết định nghiệp vụ, dùng checkpoint và approval để người có trách nhiệm xác nhận.", styles["body"]),
            p("3. Không nên sửa trực tiếp output nếu vẫn có thể trace ngược về staging. Làm vậy dễ mất dấu nguyên nhân và khó kiểm soát chất lượng.", styles["body"]),
            p("4. Với SX, nguyên tắc quan trọng là clean data ở staging trước khi gộp vào file output.", styles["body"]),
            p("5. Khi code đổi, phải publish lên GitHub rồi mới thông báo cho người dùng cuối. Nếu không, tài liệu và bản chạy sẽ lệch nhau.", styles["body"]),
            p("6. Khi thuyết trình, luôn đi theo mẫu 4 câu: bài toán là gì -> câu hỏi cần tư duy là gì -> xử lý thế nào -> đầu ra cuối cùng là gì.", styles["body"]),
        ]
    )

    takeaway_table = Table(
        [
            [p("Điểm chốt", styles["table"]), p("Thông điệp cho BOD / manager", styles["table"])],
            [p("Dữ liệu phải sạch trước khi vào output", styles["table"]), p("Đặc biệt với SX, dữ liệu được clean ở staging trước khi gộp vào file output, thay vì sửa trực tiếp ở file cuối.", styles["table"])],
            [p("Checkpoint tạo minh bạch", styles["table"]), p("Checkpoint cho thấy hệ thống đang chỉ ra đúng dòng cần xử lý, không để người duyệt làm việc trong bóng tối.", styles["table"])],
            [p("Approval là quyết định có kiểm soát", styles["table"]), p("Người duyệt chỉ cần chọn đúng dòng để Apply?, sau đó hệ thống ghi sang result sheet và giữ lịch sử.", styles["table"])],
            [p("Publish bảo đảm đồng bộ", styles["table"]), p("GitHub là điểm phát hành; nếu code đổi thì user phải nhận đúng bản mới.", styles["table"])],
        ],
        colWidths=[42 * mm, 136 * mm],
    )
    takeaway_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F766E")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("FONTNAME", (0, 0), (-1, -1), styles["table"].fontName),
                ("FONTSIZE", (0, 0), (-1, -1), 9.0),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(takeaway_table)
    story.append(Spacer(1, 4 * mm))
    story.append(p("Câu chốt khi thuyết trình: dự án được thiết kế để giữ dữ liệu sạch ở từng lớp; riêng SX phải clean ở staging trước khi gộp vào output.", styles["note"]))
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
    story = build_bod_story(styles)
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
