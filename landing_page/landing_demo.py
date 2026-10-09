# -*- coding: utf-8 -*-
"""
landing_page/landing_demo.py
=============================
Landing Page MẶC ĐỊNH (dùng chung cho mọi tenant CHƯA có landing page riêng).

CƠ CHẾ "MỖI TENANT 1 LANDING PAGE RIÊNG":
- app.py sẽ tìm file `landing_page/landing_{ma_so_thue}.py` của tenant đang đăng nhập
  (ma_so_thue lấy từ Control Plane — xem control_plane.py, đây là khoá quản lý tenant
  hiện tại). Nếu tìm thấy và file có hàm `render()`, dùng file đó.
- Nếu tenant CHƯA có file riêng (chưa custom landing page), app.py tự động dùng file
  này (landing_demo.py) làm mặc định — nội dung y hệt landing page gốc trước khi có
  cơ chế đa-tenant.

MUỐN TẠO LANDING PAGE RIÊNG CHO 1 TENANT?
1. Copy file này thành `landing_page/landing_{ma_so_thue}.py` (đúng mã số thuế của
   tenant đó, VD: `landing_page/landing_0304577099.py`).
2. Sửa nội dung hàm `render()` bên dưới theo đúng thương hiệu/nội dung mong muốn.
3. Không cần sửa gì ở app.py — lần đăng nhập tiếp theo của tenant đó sẽ tự nhận
   đúng file mới.
"""

import streamlit as st
import streamlit.components.v1 as components


def render():
    """
    Landing Page CV cá nhân — NGUYỄN VĂN TUYẾN
    Ứng tuyển vị trí Trợ lý Phó Chủ tịch HĐQT.

    LƯU Ý VỀ ẢNH:
    - Ảnh chân dung (portrait) và ảnh dashboard HRM được nhúng dạng base64.
    - Anh thay thế 2 biến PORTRAIT_B64 và DASHBOARD_B64 bằng chuỗi base64 thật.
    - Nếu chưa có, em để sẵn placeholder SVG để trang vẫn hiển thị đẹp.
    """

    # ========================================================================
    # PLACEHOLDER ẢNH — ANH THAY BẰNG BASE64 THẬT CỦA ANH
    # ========================================================================
    # Cách lấy base64:
    #   1. Mở terminal, chạy: base64 -w 0 anh_chan_dung.jpg > portrait.txt
    #      (trên Windows dùng: certutil -encode anh_chan_dung.jpg portrait.txt)
    #   2. Copy nội dung file portrait.txt vào giữa 2 dấu ngoặc kép bên dưới.
    #   3. Format chuỗi: "data:image/jpeg;base64,<chuỗi_base64>"
    #
    # Tạm thời dùng placeholder SVG (chân dung mặc định + dashboard mặc định).
    # ========================================================================
    PORTRAIT_B64 = "data:image/jpeg;base64,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" alt="Người sáng lập HRM Master"
    DASHBOARD_B64 = (
        "data:image/svg+xml;base64,"
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI5MDAiIGhlaWdodD0iNTAwIj48cmVjdCB3aWR0aD0iOTAwIiBoZWlnaHQ9IjUwMCIgZmlsbD0iI0Y2RjFFNyIvPjxyZWN0IHg9IjAiIHk9IjAiIHdpZHRoPSIyMDAiIGhlaWdodD0iNTAwIiBmaWxsPSIjMUY0QjQzIi8+PHRleHQgeD0iMTAwIiB5PSI1MCIgdGV4dC1hbmNob3I9Im1pZGRsZSIgZm9udC1zaXplPSIxNCIgZmlsbD0iI0ZGRiIgZm9udC1mYW1pbHk9InNhbnMtc2VyaWYiPkhSTSBXb3Jrc3BhY2U8L3RleHQ+PHRleHQgeD0iNDUwIiB5PSIyNTAiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtc2l6ZT0iMjAiIGZpbGw9IiMxMjMwMkEiIGZvbnQtZmFtaWx5PSJzYW5zLXNlcmlmIj5EYXNoYm9hcmQgSEtNIE1hc3RlcjwvdGV4dD48L3N2Zz4="
    )

    # ========================================================================
    # HTML/CSS LANDING PAGE
    # ========================================================================
    _landing_html = f"""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Be+Vietnam+Pro:wght@300;400;500;600;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
  #hrm-landing{{
    --paper: #F6F1E7;
    --paper-deep: #EDE4D3;
    --ink: #262220;
    --ink-soft: #7a7069;
    --pine: #1F4B43;
    --pine-deep: #12302A;
    --amber: #C98A2C;
    --rust: #A23B2E;
    --line: #DCD3C3;
    --radius: 4px;
    background:var(--paper);
    color:var(--ink);
    font-family:'Be Vietnam Pro', sans-serif;
    font-size:14px;
    font-weight:400;
    line-height:1.55;
    border-radius:10px;
    overflow:hidden;
    border:1px solid var(--line);
  }}
  #hrm-landing *{{ box-sizing:border-box; }}
  #hrm-landing img{{ max-width:100%; display:block; }}
  #hrm-landing .hl-wrap{{ max-width:900px; margin:0 auto; padding:0 26px; }}
  #hrm-landing h1,#hrm-landing h2,#hrm-landing h3{{
    font-family:'Fraunces', serif;
    font-weight:500;
    margin:0;
    color:var(--pine-deep);
  }}

  /* ---------- HERO / CV HEADER ---------- */
  #hrm-landing .cv-hero{{
    padding:40px 0 32px;
    border-bottom:1px solid var(--line);
    background:linear-gradient(180deg, var(--paper) 0%, var(--paper-deep) 100%);
  }}
  #hrm-landing .cv-hero-grid{{
    display:flex; gap:30px; align-items:center;
  }}
  #hrm-landing .cv-portrait{{
    flex:0 0 auto;
    width:150px; height:150px;
    border-radius:50%;
    overflow:hidden;
    border:3px solid #fff;
    box-shadow:0 8px 24px rgba(18,48,42,.18), 0 0 0 1px var(--line);
  }}
  #hrm-landing .cv-portrait img{{
    width:100%; height:100%; object-fit:cover;
  }}
  #hrm-landing .cv-hero-text{{ flex:1 1 auto; }}
  #hrm-landing .cv-eyebrow{{
    display:inline-flex; align-items:center; gap:7px;
    font-family:'Space Mono', monospace;
    font-size:10.5px; letter-spacing:.12em; text-transform:uppercase;
    color:var(--rust); margin-bottom:10px; font-weight:400;
  }}
  #hrm-landing .cv-eyebrow::before{{ content:""; width:16px; height:1px; background:var(--rust); display:inline-block; }}
  #hrm-landing .cv-hero h1{{
    font-size:clamp(22px, 3vw, 30px);
    line-height:1.2;
    font-weight:600;
    letter-spacing:-.01em;
    color:var(--pine-deep);
    text-transform:uppercase;
  }}
  #hrm-landing .cv-hero h1 em{{
    font-style:italic; color:var(--amber); font-weight:500;
    text-transform:none; letter-spacing:0;
  }}
  #hrm-landing .cv-tagline{{
    margin-top:10px; font-size:13.5px; font-weight:400;
    color:var(--pine); line-height:1.55;
    border-left:3px solid var(--amber); padding-left:12px;
  }}
  #hrm-landing .cv-meta{{
    margin-top:14px; display:flex; flex-wrap:wrap; gap:8px 18px;
    font-family:'Space Mono', monospace; font-size:11px;
    color:var(--ink-soft);
  }}
  #hrm-landing .cv-meta span{{ display:inline-flex; align-items:center; gap:5px; }}
  #hrm-landing .cv-cta-row{{ margin-top:18px; display:flex; gap:12px; flex-wrap:wrap; }}
  #hrm-landing .cv-btn{{
    font-family:'Be Vietnam Pro', sans-serif; font-weight:600; font-size:12.5px;
    padding:10px 20px; border-radius:var(--radius); border:1px solid transparent;
    display:inline-flex; align-items:center; gap:7px;
    background:var(--pine); color:var(--paper); box-shadow:2.5px 2.5px 0 var(--amber);
    text-decoration:none; cursor:pointer;
  }}
  #hrm-landing .cv-btn.ghost{{
    background:transparent; color:var(--pine-deep);
    border:1px solid var(--pine-deep); box-shadow:2.5px 2.5px 0 var(--paper-deep);
  }}

  /* ---------- SECTION COMMON ---------- */
  #hrm-landing section.cv-sec{{ padding:38px 0; border-bottom:1px solid var(--line); }}
  #hrm-landing .cv-kicker{{
    font-family:'Space Mono', monospace; font-size:10.5px; color:var(--pine);
    letter-spacing:.08em; text-transform:uppercase; margin-bottom:10px; font-weight:400;
  }}
  #hrm-landing .cv-kicker::before{{ content:"— "; color:var(--rust); }}
  #hrm-landing .cv-sec h2{{ font-size:20px; margin-bottom:14px; font-weight:600; }}

  /* ---------- TRIẾT LÝ & MỤC TIÊU ---------- */
  #hrm-landing .cv-philo{{
    background:var(--pine-deep); color:#EFE8D8;
    border-radius:8px; padding:26px 28px; margin-top:6px;
    border-left:4px solid var(--amber);
  }}
  #hrm-landing .cv-philo .cv-quote{{
    font-family:'Fraunces', serif; font-style:italic;
    font-size:16px; line-height:1.65; color:#fff;
    margin:0 0 14px; font-weight:400;
  }}
  #hrm-landing .cv-philo .cv-quote::before{{ content:"“"; color:var(--amber); font-size:26px; line-height:0; vertical-align:-6px; margin-right:3px; }}
  #hrm-landing .cv-philo .cv-quote::after{{ content:"”"; color:var(--amber); font-size:26px; line-height:0; vertical-align:-6px; margin-left:3px; }}
  #hrm-landing .cv-goals{{
    list-style:none; padding:0; margin:14px 0 0;
    display:grid; grid-template-columns:1fr; gap:10px;
  }}
  #hrm-landing .cv-goals li{{
    display:flex; gap:10px; align-items:flex-start;
    font-size:12.5px; color:#c9d8d3; font-weight:300; line-height:1.55;
  }}
  #hrm-landing .cv-goals li::before{{
    content:"◆"; color:var(--amber); font-size:9px; flex:0 0 auto; margin-top:5px;
  }}
  #hrm-landing .cv-goals li b{{ color:#fff; font-weight:600; }}

  /* ---------- 6 NĂNG LỰC CỐT LÕI ---------- */
  #hrm-landing .cv-comp-grid{{
    display:grid; grid-template-columns:repeat(3, 1fr);
    gap:14px; margin-top:6px;
  }}
  #hrm-landing .cv-comp-card{{
    background:#fff; border:1px solid var(--line);
    border-radius:6px; padding:18px 16px;
    position:relative; transition:transform .15s ease, box-shadow .15s ease;
  }}
  #hrm-landing .cv-comp-card:hover{{
    transform:translateY(-2px);
    box-shadow:0 8px 20px rgba(18,48,42,.10);
  }}
  #hrm-landing .cv-comp-num{{
    font-family:'Space Mono', monospace; font-size:22px;
    color:var(--paper-deep); font-weight:700;
    position:absolute; top:10px; right:14px; line-height:1;
  }}
  #hrm-landing .cv-comp-ic{{ font-size:20px; display:block; margin-bottom:10px; }}
  #hrm-landing .cv-comp-card h4{{
    font-family:'Fraunces', serif; font-size:13.5px;
    margin:0 0 6px; color:var(--pine-deep); font-weight:600;
    line-height:1.35; padding-right:30px;
  }}
  #hrm-landing .cv-comp-card p{{
    margin:0; font-size:11.5px; color:var(--ink-soft);
    font-weight:300; line-height:1.55;
  }}

  /* ---------- TIMELINE KINH NGHIỆM ---------- */
  #hrm-landing .cv-tl{{ position:relative; padding-left:28px; margin-top:8px; }}
  #hrm-landing .cv-tl::before{{
    content:""; position:absolute; left:6px; top:6px; bottom:6px; width:1.5px;
    background:repeating-linear-gradient(to bottom, var(--pine) 0 5px, transparent 5px 9px);
  }}
  #hrm-landing .cv-tl-item{{ position:relative; padding-bottom:26px; }}
  #hrm-landing .cv-tl-item:last-child{{ padding-bottom:0; }}
  #hrm-landing .cv-tl-item::before{{
    content:""; position:absolute; left:-28px; top:4px; width:12px; height:12px;
    border-radius:50%; background:var(--paper); border:2.5px solid var(--rust);
  }}
  #hrm-landing .cv-tl-item.milestone::before{{ border-color:var(--amber); background:var(--amber); }}
  #hrm-landing .cv-tl-date{{
    font-family:'Space Mono', monospace; font-size:10.5px; color:var(--rust);
    background:#fff; display:inline-block; padding:2px 9px; border-radius:2px;
    border:1px solid var(--line); margin-bottom:6px; font-weight:400;
  }}
  #hrm-landing .cv-tl-item h3{{
    font-size:14.5px; margin-bottom:3px; font-weight:600;
    color:var(--pine-deep);
  }}
  #hrm-landing .cv-tl-item .cv-role{{
    font-size:11.5px; color:var(--amber); font-weight:600;
    font-family:'Space Mono', monospace; margin-bottom:6px;
    text-transform:uppercase; letter-spacing:.04em;
  }}
  #hrm-landing .cv-tl-item ul{{
    margin:6px 0 0; padding-left:16px;
    font-size:12px; color:var(--ink-soft); font-weight:300; line-height:1.6;
  }}
  #hrm-landing .cv-tl-item ul li{{ margin-bottom:4px; }}
  #hrm-landing .cv-tl-item ul li b{{ color:var(--ink); font-weight:600; }}

  /* Case study HRM Master */
  #hrm-landing .cv-case{{
    margin-top:16px; background:var(--paper-deep);
    border-radius:8px; padding:20px 22px;
    border-left:4px solid var(--amber);
  }}
  #hrm-landing .cv-case .cv-case-badge{{
    display:inline-block;
    font-family:'Space Mono', monospace; font-size:9.5px;
    letter-spacing:.08em; text-transform:uppercase;
    background:var(--amber); color:#2a1c05;
    padding:3px 9px; border-radius:20px; font-weight:700;
    margin-bottom:10px;
  }}
  #hrm-landing .cv-case h3{{
    font-size:16px; margin-bottom:8px; font-weight:600;
  }}
  #hrm-landing .cv-case p{{
    margin:0 0 10px; font-size:12.5px; color:var(--ink-soft);
    font-weight:300; line-height:1.65;
  }}
  #hrm-landing .cv-case .cv-case-mockup{{
    margin-top:12px; background:#fff; border-radius:6px;
    overflow:hidden; border:1px solid var(--line);
    box-shadow:0 6px 16px rgba(18,48,42,.10);
  }}
  #hrm-landing .cv-case .cv-case-bar{{
    background:var(--paper); padding:7px 11px;
    display:flex; align-items:center; gap:5px;
    border-bottom:1px solid var(--line);
  }}
  #hrm-landing .cv-case .cv-dot{{ width:7px; height:7px; border-radius:50%; background:#d8cfbf; }}
  #hrm-landing .cv-case .cv-case-url{{
    margin-left:8px; font-family:'Space Mono', monospace;
    font-size:10px; color:var(--ink-soft);
  }}
  #hrm-landing .cv-case .cv-case-mockup img{{ width:100%; display:block; }}

  /* ---------- HỌC VẤN & KỸ NĂNG ---------- */
  #hrm-landing .cv-edu-grid{{
    display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-top:6px;
  }}
  #hrm-landing .cv-edu-card{{
    background:#fff; border:1px solid var(--line);
    border-radius:6px; padding:16px 18px;
  }}
  #hrm-landing .cv-edu-card .cv-edu-ic{{
    font-size:22px; display:block; margin-bottom:8px;
  }}
  #hrm-landing .cv-edu-card h4{{
    font-family:'Fraunces', serif; font-size:14px;
    margin:0 0 4px; color:var(--pine-deep); font-weight:600;
  }}
  #hrm-landing .cv-edu-card p{{
    margin:0; font-size:11.5px; color:var(--ink-soft);
    font-weight:300; line-height:1.55;
  }}
  #hrm-landing .cv-edu-card p b{{ color:var(--ink); font-weight:600; }}

  #hrm-landing .cv-skill-tags{{
    margin-top:16px; display:flex; flex-wrap:wrap; gap:7px;
  }}
  #hrm-landing .cv-skill-tags span{{
    font-family:'Space Mono', monospace; font-size:10.5px;
    background:var(--pine-deep); color:#EFE8D8;
    padding:5px 11px; border-radius:20px; font-weight:400;
  }}

  /* ---------- FOOTER ---------- */
  #hrm-landing footer{{
    padding:18px 0 22px; text-align:center;
    font-family:'Space Mono', monospace;
    font-size:10px; color:var(--ink-soft); font-weight:400;
    background:var(--paper-deep);
  }}

  @media (max-width: 720px){{
    #hrm-landing .cv-hero-grid{{ flex-direction:column; text-align:center; }}
    #hrm-landing .cv-portrait{{ margin:0 auto; }}
    #hrm-landing .cv-tagline{{ border-left:none; border-top:3px solid var(--amber); padding-left:0; padding-top:10px; }}
    #hrm-landing .cv-comp-grid{{ grid-template-columns:1fr; }}
    #hrm-landing .cv-edu-grid{{ grid-template-columns:1fr; }}
    #hrm-landing .cv-cta-row{{ justify-content:center; }}
  }}
</style>

<div id="hrm-landing">

  <!-- ================= HERO / CV HEADER ================= -->
  <section class="cv-hero">
    <div class="hl-wrap">
      <div class="cv-hero-grid">
        <div class="cv-portrait">
          <img src="{PORTRAIT_B64}" alt="Chân dung Nguyễn Văn Tuyến">
        </div>
        <div class="cv-hero-text">
          <div class="cv-eyebrow">Curriculum Vitae · 2026</div>
          <h1>Nguyễn Văn Tuyến<br><em>Ứng tuyển Trợ lý Phó Chủ tịch HĐQT</em></h1>
          <p class="cv-tagline">
            Thấu hiểu tư duy Lãnh đạo &nbsp;·&nbsp; Am hiểu Quản trị &amp; Đối ngoại &nbsp;·&nbsp; Tiên phong Ứng dụng AI/Công nghệ
          </p>
          <div class="cv-meta">
            <span>📞 0961 778 150</span>
            <span>✉️ nguyen.support999@gmail.com</span>
            <span>🎂 1977</span>
            <span>🚗 Bằng B1 · 20 năm lái xe an toàn</span>
          </div>
          <div class="cv-cta-row">
            <a class="cv-btn" href="http://demo-hrm9.streamlit.app" target="_blank" rel="noopener">→ Đăng nhập trải nghiệm HRM Master</a>
            <a class="cv-btn ghost" href="#" onclick="alert('Anh có thể upload file CV PDF lên Google Drive và thay link vào đây.'); return false;">⬇ Tải file CV PDF</a>
          </div>
        </div>
      </div>
    </div>
  </section>

  <!-- ================= KHỐI 1: TRIẾT LÝ & MỤC TIÊU ================= -->
  <section class="cv-sec">
    <div class="hl-wrap">
      <div class="cv-kicker">Triết lý làm việc &amp; Mục tiêu nghề nghiệp</div>
      <h2>Lắng nghe — Thấu hiểu — Chuẩn bị — Làm bệ phóng</h2>

      <div class="cv-philo">
        <p class="cv-quote">
          Ở vai trò Trợ lý, vai trò quan trọng nhất là lắng nghe, thấu hiểu, chuẩn bị chu đáo nguồn lực và làm bệ phóng cho các quyết định của Phó Chủ tịch. Tôi không đến để làm thay, mà đến để làm cho mọi việc trở nên nhẹ nhàng, chính xác và kín đáo hơn.
        </p>
        <ul class="cv-goals">
          <li><b>Trở thành "cánh tay phải" tin cậy</b> — thấu hiểu và chia sẻ áp lực công việc với Phó Chủ tịch.</li>
          <li><b>Chuẩn bị chu đáo hồ sơ, đàm phán/đối ngoại chuyên nghiệp</b> với đối tác và cơ quan Quản lý nhà nước.</li>
          <li><b>Điều phối lịch trình, theo dõi đôn đốc công việc đa phòng ban</b> chính xác, bảo mật tuyệt đối hình ảnh và thông tin Lãnh đạo.</li>
        </ul>
      </div>
    </div>
  </section>

  <!-- ================= KHỐI 2: 6 NĂNG LỰC CỐT LÕI ================= -->
  <section class="cv-sec">
    <div class="hl-wrap">
      <div class="cv-kicker">Năng lực cốt lõi</div>
      <h2>6 điểm mạnh phù hợp trực tiếp với vị trí</h2>

      <div class="cv-comp-grid">
        <div class="cv-comp-card">
          <span class="cv-comp-num">01</span>
          <span class="cv-comp-ic">🌐</span>
          <h4>Đối ngoại &amp; Quan hệ Công chúng</h4>
          <p>Kinh nghiệm làm việc trực tiếp với cơ quan nhà nước (Thuế, BHXH, Lao động, PCCC, Hải quan); đàm phán và giao thương thành công với đối tác quốc tế (Hàn Quốc).</p>
        </div>

        <div class="cv-comp-card">
          <span class="cv-comp-num">02</span>
          <span class="cv-comp-ic">📊</span>
          <h4>Soạn thảo &amp; Tổng hợp Báo cáo Cấp cao</h4>
          <p>Thành thạo xây dựng văn bản hành chính, quy chế, báo cáo phân tích, slide trình bày (PowerPoint) chuẩn mực; kỹ năng xử lý số liệu chuyên sâu (Excel/Dashboard).</p>
        </div>

        <div class="cv-comp-card">
          <span class="cv-comp-num">03</span>
          <span class="cv-comp-ic">🗓️</span>
          <h4>Sắp xếp, Điều phối &amp; Quản trị Tiến độ</h4>
          <p>Hơn 15 năm kinh nghiệm điều hành, phân công công việc, quản lý lịch trình tiếp khách/công tác, đôn đốc và theo dõi việc triển khai sau họp.</p>
        </div>

        <div class="cv-comp-card">
          <span class="cv-comp-num">04</span>
          <span class="cv-comp-ic">🔒</span>
          <h4>Bảo mật &amp; Tác phong Chuyên nghiệp</h4>
          <p>Luôn đặt tính bảo mật thông tin nội bộ và hình ảnh Lãnh đạo lên hàng đầu; xử lý tình huống linh hoạt, chịu áp lực cao.</p>
        </div>

        <div class="cv-comp-card">
          <span class="cv-comp-num">05</span>
          <span class="cv-comp-ic">🤖</span>
          <h4>Đột phá Hiệu suất bằng AI &amp; Số hóa</h4>
          <p>Thành thạo ứng dụng AI (ChatGPT, Claude...) và các công cụ công nghệ để tóm tắt tài liệu, dịch thuật đối ngoại, lập báo cáo và nhắc lịch tự động cho HĐQT/Ban Điều hành.</p>
        </div>

        <div class="cv-comp-card">
          <span class="cv-comp-num">06</span>
          <span class="cv-comp-ic">🌱</span>
          <h4>Tinh thần cầu thị &amp; Thích ứng nhanh</h4>
          <p>Sẵn sàng lắng nghe, học hỏi và nắm bắt phong cách làm việc của Ban Điều hành. Hợp tác trực tiếp với Phó Chủ tịch là cơ hội để đóng góp trải nghiệm, đồng thời không ngừng hoàn thiện bản thân.</p>
        </div>
      </div>
    </div>
  </section>

  <!-- ================= KHỐI 3: HÀNH TRÌNH KINH NGHIỆM ================= -->
  <section class="cv-sec">
    <div class="hl-wrap">
      <div class="cv-kicker">Hành trình kinh nghiệm &amp; Thành tựu</div>
      <h2>Hơn 20 năm quản trị — từ Bảo Việt đến HRM Master</h2>

      <div class="cv-tl">

        <div class="cv-tl-item milestone">
          <div class="cv-tl-date">05/2026 — nay</div>
          <h3>Chuyên viên Hành chính Nhân sự · Tham mưu Quản trị</h3>
          <div class="cv-role">Doanh nghiệp hiện tại · gần 80 nhân viên</div>
          <ul>
            <li><b>Tham mưu &amp; Xây dựng thể chế:</b> Chủ trì soạn thảo Nội quy lao động, Chức năng nhiệm vụ các phòng ban, Hồ sơ PCCC, An toàn lao động và hồ sơ làm việc với Cơ quan Quản lý Nhà nước.</li>
            <li><b>Chuẩn hóa quy trình &amp; Số hóa:</b> Khảo sát quy trình vận hành, ứng dụng AI xây dựng hệ thống quản trị HRM số hóa giúp tự động hóa theo dõi trọn đời mỗi nhân viên, nghiệp vụ BHXH, chấm công, C&amp;B, báo cáo, công văn, hợp đồng kinh tế — <b>tối ưu 40% thời gian xử lý thủ tục hành chính.</b></li>
          </ul>
        </div>

        <div class="cv-tl-item">
          <div class="cv-tl-date">2023 — 04/2026</div>
          <h3>Freelancer · Tư vấn du học và Xuất khẩu lao động</h3>
          <div class="cv-role">CTCP Cung ứng Nhân lực Nghệ An</div>
          <ul>
            <li>Tư vấn và xây dựng lộ trình cho hàng trăm khách hàng.</li>
            <li>Tìm kiếm, đàm phán và ký kết hợp đồng với các đối tác nước ngoài để cung ứng nguồn nhân lực Việt Nam.</li>
          </ul>
        </div>

        <div class="cv-tl-item milestone">
          <div class="cv-tl-date">2010 — 2023</div>
          <h3>Founder &amp; CEO</h3>
          <div class="cv-role">CTCP Đầu tư và Thương mại An Phú Lộc</div>
          <ul>
            <li><b>Đối ngoại &amp; Thương mại Quốc tế:</b> Trực tiếp tìm kiếm, đàm phán, ký kết hợp đồng nhập khẩu với đối tác Hàn Quốc; hoàn tất trọn vẹn hồ sơ thủ tục Hải quan và Logistics.</li>
            <li><b>Quan hệ công quyền &amp; Pháp lý:</b> Trực tiếp làm việc với cơ quan Thuế, Lao động, chính quyền địa phương để đảm bảo doanh nghiệp tuân thủ đúng pháp luật.</li>
            <li><b>Xây dựng văn bản &amp; Báo cáo:</b> Xây dựng toàn bộ hệ thống quy chế nội bộ, quy trình vận hành, báo cáo tài chính - kinh doanh định kỳ.</li>
            <li><b>Tổ chức sự kiện &amp; Đối ngoại:</b> Sắp xếp lịch tiếp đón đối tác, tổ chức chương trình tài trợ, sự kiện nội bộ và hoạt động gắn kết cộng đồng.</li>
            <li><i>Doanh nghiệp tạm dừng hoạt động năm 2023 do ảnh hưởng kéo dài của đại dịch Covid-19.</i></li>
          </ul>
        </div>

        <div class="cv-tl-item">
          <div class="cv-tl-date">1999 — 2010</div>
          <h3>Trưởng phòng Nghiệp vụ</h3>
          <div class="cv-role">Tập đoàn Bảo Việt · Bảo Việt Nhân thọ Nghệ An</div>
          <ul>
            <li>Được bổ nhiệm Phụ trách phòng từ năm 24 tuổi nhờ năng lực quản lý và thành tích xuất sắc.</li>
            <li><b>Điều phối &amp; Quản trị đội ngũ:</b> Quản lý và điều phối trực tiếp hệ thống ~300 nhân sự/tư vấn viên; chịu trách nhiệm lập kế hoạch, theo dõi tiến độ, tổ chức các cuộc họp giao ban và hội nghị.</li>
            <li><b>Phối hợp đa phòng ban:</b> Làm cầu nối điều phối công việc giữa Ban Lãnh đạo Tập đoàn với các bộ phận chức năng.</li>
          </ul>
        </div>

      </div>

      <!-- Case study HRM Master -->
      <div class="cv-case">
        <span class="cv-case-badge">Case study tiêu biểu</span>
        <h3>HRM Master — Phần mềm Quản trị Nhân sự</h3>
        <p>
          Phần mềm HR/C&amp;B/BHXH/Hành chính đang vận hành thực tế; tự phân tích nghiệp vụ, thiết kế kiến trúc dữ liệu và chỉ đạo AI xây dựng toàn bộ hệ thống <b>(hơn 15.000 dòng code)</b>.
          Các module: Dashboard, Nhân viên, BHXH, Chấm công, Tính thu nhập, Công văn và Hợp đồng kinh tế, Chat nội bộ, Chatbot pháp luật lao động, Audit.
        </p>
        <div class="cv-case-mockup">
          <div class="cv-case-bar">
            <span class="cv-dot"></span><span class="cv-dot"></span><span class="cv-dot"></span>
            <span class="cv-case-url">demo-hrm9.streamlit.app</span>
          </div>
          <img src="{DASHBOARD_B64}" alt="Dashboard HRM Master">
        </div>
      </div>

    </div>
  </section>

  <!-- ================= KHỐI 4: HỌC VẤN & KỸ NĂNG ================= -->
  <section class="cv-sec">
    <div class="hl-wrap">
      <div class="cv-kicker">Học vấn, Kỹ năng &amp; Bằng lái xe</div>
      <h2>Nền tảng học thuật &amp; công cụ hỗ trợ công việc</h2>

      <div class="cv-edu-grid">
        <div class="cv-edu-card">
          <span class="cv-edu-ic">🎓</span>
          <h4>Thạc sĩ Quản trị Kinh doanh (MBA)</h4>
          <p><b>Đại học Irvine (Hoa Kỳ)</b> kết hợp HSB Hà Nội · 2007</p>
        </div>
        <div class="cv-edu-card">
          <span class="cv-edu-ic">🏛️</span>
          <h4>Cử nhân Quản trị Kinh doanh</h4>
          <p><b>Đại học Kinh tế Quốc dân (NEU)</b> · 1999</p>
        </div>
        <div class="cv-edu-card">
          <span class="cv-edu-ic">🗣️</span>
          <h4>Ngoại ngữ</h4>
          <p>Tiếng Anh đọc hiểu văn bản, đàm phán hợp đồng thương mại qua văn bản/email; ứng dụng AI dịch thuật chuẩn xác tài liệu đối ngoại.</p>
        </div>
        <div class="cv-edu-card">
          <span class="cv-edu-ic">🚗</span>
          <h4>Bằng lái xe</h4>
          <p><b>Hạng B1</b> · Kinh nghiệm <b>20 năm lái xe an toàn</b>, sẵn sàng tự lái phục vụ Lãnh đạo công tác trong/ngoài tỉnh.</p>
        </div>
      </div>

      <div class="cv-skill-tags">
        <span>PowerPoint trình bày báo cáo</span>
        <span>Excel số liệu chuyên sâu</span>
        <span>AI Agents (Claude, ChatGPT)</span>
        <span>Dashboard quản trị</span>
        <span>Soạn thảo văn bản hành chính</span>
        <span>Quản trị tiến độ &amp; điều phối</span>
        <span>Bảo mật thông tin</span>
        <span>Đối ngoại cơ quan nhà nước</span>
      </div>
    </div>
  </section>

  <footer>
    © 2026 Nguyễn Văn Tuyến · CV ứng tuyển vị trí Trợ lý Phó Chủ tịch HĐQT
  </footer>

</div>

<script>
  // Iframe của components.html có chiều cao cố định theo tham số height=... truyền
  // từ Python, không tự co giãn theo nội dung — nên phải tự đo và set lại chiều cao
  // bằng JS. window.frameElement trỏ đúng tới thẻ <iframe> đang chứa nội dung này
  // (vẫn cùng-origin với trang Streamlit nên truy cập được, không cần postMessage).
  function _hrmResizeIframe() {{
    try {{
      var h = document.documentElement.scrollHeight;
      if (window.frameElement && h > 0) {{
        window.frameElement.style.height = (h + 20) + "px";
      }}
    }} catch (e) {{}}
  }}
  window.addEventListener("load", _hrmResizeIframe);
  document.fonts && document.fonts.ready && document.fonts.ready.then(_hrmResizeIframe);
  new ResizeObserver(_hrmResizeIframe).observe(document.body);
  setTimeout(_hrmResizeIframe, 200);
  setTimeout(_hrmResizeIframe, 800);
  setTimeout(_hrmResizeIframe, 2000);
</script>
"""
    components.html(_landing_html, height=50, scrolling=False)