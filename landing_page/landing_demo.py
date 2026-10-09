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
    PORTRAIT_B64 = (
        "data:image/svg+xml;base64,"
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyMDAiIGhlaWdodD0iMjAwIj48cmVjdCB3aWR0aD0iMjAwIiBoZWlnaHQ9IjIwMCIgZmlsbD0iI0U4RTBFMCIvPjxjaXJjbGUgY3g9IjEwMCIgY3k9Ijc1IiByPSIzNSIgZmlsbD0iI0I4QjhCOCIvPjxwYXRoIGQ9Ik0zMCAxOTAgUTMwIDEyMCAxMDAgMTIwIFExNzAgMTIwIDE3MCAxOTBaIiBmaWxsPSIjQjhCOEI4Ii8+PHRleHQgeD0iMTAwIiB5PSIxOTUiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtc2l6ZT0iMTEiIGZpbGw9IiM2NjYiIGZvbnQtZmFtaWx5PSJzYW5zLXNlcmlmIj5DaMOibiBkdW5nPC90ZXh0Pjwvc3ZnPg=="
    )
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
            <a class="cv-btn ghost" href="https://drive.google.com/file/d/1GbDDNIsG37tsVW4cMwHK8nZ3k8xmNKtY/view?usp=drive_link">⬇ Tải file CV PDF</a>
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