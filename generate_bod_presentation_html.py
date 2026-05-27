from __future__ import annotations

import argparse
from pathlib import Path


WORK_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = WORK_DIR / "docs" / "thuyet_trinh_bod_manager.html"


def section(title: str, eyebrow: str, body: str, notes: str | None = None) -> str:
    note_html = f'<div class="notes">{notes}</div>' if notes else ""
    return f"""
    <section class="slide">
      <div class="eyebrow">{eyebrow}</div>
      <h2>{title}</h2>
      <div class="content">{body}</div>
      {note_html}
    </section>
    """


def card(title: str, text: str, tag: str | None = None) -> str:
    tag_html = f'<div class="tag">{tag}</div>' if tag else ""
    return f"""
    <div class="card">
      {tag_html}
      <h3>{title}</h3>
      <p>{text}</p>
    </div>
    """


def build_html() -> str:
    return f"""<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Thuyet trinh BOD / Manager</title>
  <style>
    :root {{
      --bg: #08162b;
      --panel: rgba(14, 28, 56, 0.9);
      --panel-soft: rgba(20, 39, 72, 0.86);
      --line: rgba(255,255,255,.08);
      --text: #e8eef9;
      --muted: #91a4c7;
      --gold: #d8b15d;
      --teal: #2ccbb8;
      --red: #ff8686;
      --blue: #81a8ff;
      --green: #6ee7b7;
      --shadow: 0 18px 60px rgba(0,0,0,.28);
      --radius: 22px;
    }}
    * {{ box-sizing: border-box; }}
    html, body {{ margin: 0; height: 100%; background: var(--bg); color: var(--text); font-family: "Segoe UI", Arial, sans-serif; text-rendering: optimizeLegibility; -webkit-font-smoothing: antialiased; }}
    body {{
      background:
        linear-gradient(rgba(255,255,255,.03) 1px, transparent 1px),
        linear-gradient(90deg, rgba(255,255,255,.03) 1px, transparent 1px),
        radial-gradient(circle at top, rgba(24,58,110,.55), transparent 40%),
        var(--bg);
      background-size: 72px 72px, 72px 72px, auto, auto;
    }}
    .deck {{
      width: 100%;
      min-height: 100vh;
    }}
    .slide {{
      min-height: 100vh;
      padding: 64px 64px 56px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      gap: 18px;
      border-bottom: 1px solid var(--line);
    }}
    .eyebrow {{
      color: var(--gold);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .32em;
      font-weight: 700;
    }}
    h1, h2, h3, p {{ margin: 0; }}
    h1 {{
      font-family: "Times New Roman", Georgia, serif;
      font-size: clamp(56px, 8vw, 92px);
      line-height: .92;
      letter-spacing: 0;
      text-align: center;
      max-width: 1000px;
      word-spacing: normal;
      word-break: normal;
      overflow-wrap: normal;
      text-wrap: balance;
      font-kerning: normal;
    }}
    h2 {{
      font-family: "Times New Roman", Georgia, serif;
      font-size: clamp(36px, 5vw, 62px);
      line-height: 1;
      letter-spacing: 0;
      text-align: center;
      max-width: 1100px;
      word-spacing: normal;
      word-break: normal;
      overflow-wrap: normal;
      text-wrap: balance;
      font-kerning: normal;
    }}
    .subtitle {{
      max-width: 900px;
      color: var(--muted);
      font-size: 18px;
      line-height: 1.6;
      text-align: center;
    }}
    .hero-meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-top: 18px;
      justify-content: center;
    }}
    .pill {{
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 10px 16px;
      font-size: 13px;
      color: var(--text);
      background: rgba(255,255,255,.03);
    }}
    .grid-2 {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
      width: min(1200px, 100%);
    }}
    .grid-3 {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 18px;
      width: min(1280px, 100%);
    }}
    .card, .box {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 24px;
      backdrop-filter: blur(8px);
    }}
    .card h3 {{
      font-size: 24px;
      margin-bottom: 10px;
      font-family: "Times New Roman", Georgia, serif;
      letter-spacing: 0;
    }}
    .card p, .box p, li {{
      color: #c7d2e8;
      line-height: 1.7;
      font-size: 16px;
    }}
    .tag {{
      display: inline-block;
      color: var(--gold);
      border: 1px solid rgba(216,177,93,.35);
      background: rgba(216,177,93,.08);
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .18em;
      margin-bottom: 14px;
    }}
    .metrics {{
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      width: min(1200px, 100%);
    }}
    .metric {{
      padding: 18px;
      border-radius: 18px;
      background: rgba(255,255,255,.03);
      border: 1px solid var(--line);
      text-align: center;
    }}
    .metric strong {{
      display: block;
      color: var(--gold);
      font-size: 34px;
      line-height: 1;
      margin-bottom: 6px;
      font-family: "Times New Roman", Georgia, serif;
    }}
    .metric span {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.4;
    }}
    .flow {{
      display: flex;
      gap: 14px;
      flex-wrap: wrap;
      align-items: center;
      width: min(1200px, 100%);
      justify-content: center;
    }}
    .node {{
      min-width: 140px;
      padding: 18px 20px;
      border-radius: 18px;
      color: white;
      font-weight: 700;
      font-size: 18px;
      text-align: center;
      box-shadow: var(--shadow);
    }}
    .arrow {{
      color: rgba(255,255,255,.35);
      font-size: 28px;
      font-weight: 300;
    }}
    .node.source {{ background: linear-gradient(135deg, #1e3a8a, #0f172a); }}
    .node.staging {{ background: linear-gradient(135deg, #334155, #0f172a); }}
    .node.checkpoint {{ background: linear-gradient(135deg, #a16207, #4b2e00); }}
    .node.approval {{ background: linear-gradient(135deg, #0f766e, #052e2b); }}
    .node.output {{ background: linear-gradient(135deg, #111827, #020617); }}
    .list {{
      padding-left: 20px;
      margin: 0;
    }}
    .list li {{ margin-bottom: 10px; }}
    .muted {{ color: var(--muted); }}
    .two-col {{
      display: grid;
      grid-template-columns: 1.2fr .8fr;
      gap: 18px;
      width: min(1280px, 100%);
    }}
    .center {{
      text-align: center;
    }}
    .table {{
      width: 100%;
      border-collapse: collapse;
      overflow: hidden;
      border-radius: 18px;
      border: 1px solid var(--line);
      background: rgba(255,255,255,.03);
      box-shadow: var(--shadow);
    }}
    .table th, .table td {{
      padding: 16px 18px;
      text-align: left;
      vertical-align: top;
      border-bottom: 1px solid var(--line);
      font-size: 15px;
      line-height: 1.6;
    }}
    .table th {{
      color: var(--gold);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .18em;
      background: rgba(255,255,255,.02);
    }}
    .table tr:last-child td {{ border-bottom: none; }}
    .accent {{
      color: var(--gold);
      font-weight: 700;
    }}
    .notes {{
      margin-top: 12px;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.6;
      max-width: 1100px;
      text-align: center;
    }}
    .footer {{
      margin-top: 14px;
      color: var(--muted);
      font-size: 13px;
    }}
    @media (max-width: 1100px) {{
      .grid-3, .grid-2, .metrics, .two-col {{
        grid-template-columns: 1fr;
      }}
      .slide {{ padding: 40px 22px; }}
      h1 {{ max-width: none; }}
    }}
  </style>
</head>
<body>
  <main class="deck">
    <section class="slide">
      <div class="eyebrow">BOD / MANAGER DECK</div>
      <h1>Từ thao tác tay đến vận hành chuẩn hóa</h1>
      <p class="subtitle">Một deck ngắn để show trên màn hình: tại sao bài toán này ra đời, workflow chạy như thế nào, và vì sao tự động hóa tạo ra lợi ích rõ ràng cho chi phí, chất lượng và tốc độ ra quyết định.</p>
      <div class="hero-meta">
        <div class="pill">Mục tiêu tối thượng: tính được chi phí vốn hóa</div>
        <div class="pill">Đầu vào: source rải rác</div>
        <div class="pill">Đầu ra: workbook sạch, có checkpoint và approval</div>
      </div>
      <div class="metrics">
        <div class="metric"><strong>8h</strong><span>Associate nếu làm tay cho 1 chu kỳ</span></div>
        <div class="metric"><strong>4h</strong><span>Senior nếu làm tay cho 1 chu kỳ</span></div>
        <div class="metric"><strong>1 flow</strong><span>Chuẩn hóa từ source → staging → output</span></div>
        <div class="metric"><strong>0 guess</strong><span>Giảm copy/paste và kiểm tra cảm tính</span></div>
      </div>
    </section>

    <section class="slide">
      <div class="eyebrow">FLOW BÀI TOÁN</div>
      <h2>Vì sao bài toán này ra đời?</h2>
      <div class="flow">
        <div class="node source">Source</div><div class="arrow">→</div>
        <div class="node staging">Staging</div><div class="arrow">→</div>
        <div class="node checkpoint">Checkpoint</div><div class="arrow">→</div>
        <div class="node approval">Approval</div><div class="arrow">→</div>
        <div class="node output">Output</div>
      </div>
      <div class="two-col">
        <div class="box">
          <div class="tag">Bối cảnh</div>
          <ul class="list">
            <li>Dữ liệu nằm rải rác ở Google Sheets, Excel, OneDrive và các workbook output khác nhau.</li>
            <li>Làm tay nghĩa là copy/paste, đối chiếu, sửa lại, rồi chờ người khác kiểm tra.</li>
            <li>Cách đó chậm, dễ lệch và rất khó scale khi khối lượng tăng.</li>
            <li>Automation ra đời để biến một quy trình rời rạc thành một quy trình có thể lặp lại và kiểm soát.</li>
          </ul>
        </div>
        <div class="box">
          <div class="tag">Câu hỏi cần trả lời</div>
          <ul class="list">
            <li>Dữ liệu đầu vào có đúng và đủ không?</li>
            <li>Dữ liệu nào cần clean ở staging trước?</li>
            <li>Dòng nào cần người dùng duyệt?</li>
            <li>Kết quả cuối cùng có đáng tin để dùng không?</li>
          </ul>
        </div>
      </div>
      <div class="notes">Khi nói slide này, bạn chỉ cần chốt: source sai thì tất cả sai; staging là nơi làm sạch; checkpoint giúp nhìn ra ngoại lệ; approval là nơi người có trách nhiệm ra quyết định; output là thứ dùng để báo cáo và bàn giao.</div>
    </section>

    <section class="slide">
      <div class="eyebrow">LỢI ÍCH ĐỊNH LƯỢNG</div>
      <h2>Vì sao tự động hóa đáng làm?</h2>
      <table class="table">
        <tr><th>Lợi ích</th><th>Tác dụng thực tế</th></tr>
        <tr><td><span class="accent">Tiết kiệm thời gian</span></td><td>Nếu làm tay mất 8h với Associate hoặc 4h với Senior, automation chuyển phần nặng sang kiểm tra và phê duyệt. Mỗi chu kỳ có thể tiết kiệm gần 1 ngày công Associate hoặc nửa ngày Senior.</td></tr>
        <tr><td><span class="accent">Giảm sai sót</span></td><td>Không còn phụ thuộc vào thao tác copy/paste, công thức thủ công hay sửa lẻ tẻ nhiều nơi.</td></tr>
        <tr><td><span class="accent">Dễ audit</span></td><td>Có checkpoint, result sheet và log để truy ngược từng quyết định.</td></tr>
        <tr><td><span class="accent">Dễ scale</span></td><td>Khi khối lượng tăng, hệ thống vẫn chạy theo cùng một logic thay vì tăng người làm tay.</td></tr>
      </table>
      <div class="two-col" style="margin-top:18px;">
        <div class="card">
          <div class="tag">Ví dụ SX</div>
          <h3>Clean staging trước, rồi mới gộp output</h3>
          <p>Với SX, vấn đề thường không nằm ở bước cuối mà nằm ở dữ liệu đầu vào: MNV thiếu, mapping lệch, hoặc downstream bị kéo sai. Vì vậy cách xử lý đúng là clean data ở staging trước, rồi mới gộp vào file output.</p>
        </div>
        <div class="card">
          <div class="tag">Ý nghĩa</div>
          <p>Code có checkpoint riêng cho <span class="accent">checkpoint data SX</span> và <span class="accent">Check_SX_Downstream</span> để nhìn rõ lỗi ở nguồn và lỗi lan sang <span class="accent">SX_Allocation_Build</span>, <span class="accent">Timesheet SX</span>, <span class="accent">4.1 Chi phí nhân sự SX</span> và <span class="accent">3.Vốn hóa</span>.</p>
        </div>
      </div>
      <div class="notes">Thông điệp ngắn cho BOD: tự động hóa không chỉ tạo file nhanh hơn, mà còn chuẩn hóa quyết định, giảm rủi ro dữ liệu, và giúp chi phí vốn hóa được tính trên bộ dữ liệu sạch hơn.</div>
    </section>

    <section class="slide">
      <div class="eyebrow">VÍ DỤ THỰC TẾ</div>
      <h2>IT, Media và những khó khăn thường gặp</h2>
      <div class="grid-3">
        {card("IT", "Dòng nào đã vào Timesheet IT, dòng nào sang Chi phí nhân sự IT, dòng nào cần tạo mới project master? Chốt bằng checkpoint để tránh sửa sai ở output.", "Need to know")}
        {card("Media", "Dữ liệu đầu vào thường phải đối chiếu giữa Timesheet Media và Data media ACCA+CFA+CMA. Cần chắc rằng mapping dự án, tháng và người làm task là đúng.", "Need to know")}
        {card("Struggles", "Khó nhất không phải code, mà là dữ liệu đầu vào không sạch, tên sheet thay đổi, nhân sự mới chưa map xong, hoặc project mới chưa có master đúng.", "Common pain")}
      </div>
      <div class="two-col" style="margin-top:18px;">
        <div class="box">
          <div class="tag">Khi có dự án mới</div>
          <ul class="list">
            <li>Thường phải bổ sung ở 1.Danh mục dự án.</li>
            <li>Có thể phải cập nhật IT_New_Project_Master nếu cần đưa project vào luồng IT.</li>
            <li>Rà lại Check_IT_Downstream để bảo đảm project mới rơi đúng vào chi phí và báo cáo.</li>
          </ul>
        </div>
        <div class="box">
          <div class="tag">Khi có nhân viên mới</div>
          <ul class="list">
            <li>Phải map ở Mã nhân viên.</li>
            <li>Kiểm tra lại Check_Payroll nếu lương / tháng thay đổi.</li>
            <li>Nếu là người tham gia SX, còn phải bảo đảm mapping không làm hỏng Check_SX_Downstream.</li>
          </ul>
        </div>
      </div>
      <div class="notes">Bạn có thể nói thẳng: một dự án mới hoặc nhân viên mới không chỉ sửa một ô. Nó thường kéo theo cập nhật ở nhiều lớp: danh mục, mapping, checkpoint, approval và sheet kết quả.</div>
    </section>

    <section class="slide">
      <div class="eyebrow">TÁC ĐỘNG DỮ LIỆU</div>
      <h2>Thay đổi một đầu vào thì đụng bao nhiêu sheet?</h2>
      <table class="table">
        <tr><th>Tình huống</th><th>Sheet thường cần xem / chỉnh</th><th>Ghi chú</th></tr>
        <tr><td><span class="accent">Dự án mới</span></td><td>1.Danh mục dự án, IT_New_Project_Master, Check_IT_Downstream, 3.Vốn hóa</td><td>Có thể thêm cả sheet result của IT nếu duyệt lại.</td></tr>
        <tr><td><span class="accent">Nhân viên mới</span></td><td>Mã nhân viên, Check_Payroll, Timesheet IT / Media / SX, checkpoint tương ứng</td><td>Nếu map chưa xong thì output vẫn chạy nhưng sẽ thiếu / sai dữ liệu.</td></tr>
        <tr><td><span class="accent">Lương mới</span></td><td>Lương nhân viên full time, Lương nhân viên part time, Check_Payroll</td><td>Đây là lớp nguồn, phải đúng trước khi đẩy vào output.</td></tr>
        <tr><td><span class="accent">SX phát sinh lỗi</span></td><td>checkpoint data SX, Check_SX_Downstream, SX_Allocation_Build, Timesheet SX, 4.1 Chi phí nhân sự SX, 3.Vốn hóa</td><td>Nguyên tắc: clean ở staging trước khi gộp output.</td></tr>
      </table>
      <div class="notes">Nếu cần một câu ngắn để trình bày: một thay đổi đầu vào có thể chạm từ 3 đến 6 sheet, tùy nó nằm ở nguồn, ở mapping hay ở downstream. Tự động hóa giúp mình nhìn thấy các điểm chạm đó rõ ràng hơn.</div>
    </section>

    <section class="slide">
      <div class="eyebrow">KẾT LUẬN</div>
      <h2>Đích đến cuối cùng</h2>
      <div class="grid-2">
        <div class="card">
          <div class="tag">Giá trị</div>
          <p>Giảm thời gian thao tác tay, giảm lỗi khi copy/paste, giảm phụ thuộc vào một cá nhân, và tăng khả năng bàn giao.</p>
        </div>
        <div class="card">
          <div class="tag">Quản trị</div>
          <p>Cho phép dùng checkpoint và approval để kiểm soát ngoại lệ thay vì sửa tay ở output cuối.</p>
        </div>
        <div class="card">
          <div class="tag">BOD / Manager</div>
          <p>Đọc theo 4 bước: flow bài toán → câu hỏi cần tư duy → cách xử lý → đầu ra. Với SX, nhấn mạnh clean staging trước khi gộp output.</p>
        </div>
        <div class="card">
          <div class="tag">Một câu chốt</div>
          <p>Workflow này biến một công việc nặng thao tác tay thành một quy trình có kiểm soát, có thể mở rộng và có thể audit.</p>
        </div>
      </div>
      <div class="footer">Mở file HTML này trên trình duyệt để trình chiếu. Nếu muốn, có thể dùng F11 hoặc chế độ toàn màn hình.</div>
    </section>
  </main>
</body>
</html>
"""


def build_html_file(output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(build_html(), encoding="utf-8")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate BOD presentation HTML deck.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output HTML path")
    args = parser.parse_args()
    out = build_html_file(Path(args.output))
    print(f"Da tao HTML: {out}")


if __name__ == "__main__":
    main()
