# -*- coding: utf-8 -*-
"""
landing_page/landing_demo.py
=============================
Landing Page CV cá nhân — NGUYỄN VĂN TUYẾN
Phiên bản "Portfolio of Impact" — thiết kế cho chương trình AI Thực Chiến VinUni.

Mục tiêu: Chứng minh năng lực thực chiến qua sản phẩm HRM Master đã vận hành,
thay vì chỉ liệt kê kỹ năng. Nêu bật: Tư duy hệ thống, chuyển bài toán quản trị
thành sản phẩm, vai trò AI Agent, và kết quả đo lường được.

LƯU Ý VỀ ẢNH:
- Các sơ đồ (kiến trúc, luồng dữ liệu) và ảnh chân dung được nhúng base64.
- Anh thay thế các biến *_B64 bằng chuỗi base64 thật.
"""

import streamlit as st
import streamlit.components.v1 as components


def render():
    """
    Landing Page CV — Portfolio of Impact.
    """
    # ========================================================================
    # PLACEHOLDER ẢNH — ANH THAY BẰNG BASE64 THẬT
    # ========================================================================
    # Hướng dẫn lấy base64: dùng trang base64-image.de hoặc lệnh certutil trên Windows.
    # ========================================================================
    PORTRAIT_B64 = (
        "data:image/svg+xml;base64,"
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyMDAiIGhlaWdodD0iMjAwIj48cmVjdCB3aWR0aD0iMjAwIiBoZWlnaHQ9IjIwMCIgZmlsbD0iI0U4RTBFMCIvPjxjaXJjbGUgY3g9IjEwMCIgY3k9Ijc1IiByPSIzNSIgZmlsbD0iI0I4QjhCOCIvPjxwYXRoIGQ9Ik0zMCAxOTAgUTMwIDEyMCAxMDAgMTIwIFExNzAgMTIwIDE3MCAxOTBaIiBmaWxsPSIjQjhCOEI4Ii8+PHRleHQgeD0iMTAwIiB5PSIxOTUiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtc2l6ZT0iMTEiIGZpbGw9IiM2NjYiIGZvbnQtZmFtaWx5PSJzYW5zLXNlcmlmIj5DaMOibiBkdW5nPC90ZXh0Pjwvc3ZnPg=="
    )
    # Sơ đồ kiến trúc hệ thống HRM Master
    ARCHITECTURE_B64 = (
        "data:image/svg+xml;base64,"
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI5MDAiIGhlaWdodD0iNTAwIj48cmVjdCB3aWR0aD0iOTAwIiBoZWlnaHQ9IjUwMCIgZmlsbD0iI0Y2RjFFNyIvPjx0ZXh0IHg9IjQ1MCIgeT0iMjUwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LXNpemU9IjIwIiBmaWxsPSIjMTIzMDJBIiBmb250LWZhbWlseT0ic2Fucy1zZXJpZiI+U8ahIMSQLiBLaeG6v04gVFLDc0MgSFJNIE1hc3RlcjwvdGV4dD48L3N2Zz4="
    )
    # Sơ đồ luồng dữ liệu
    DATAFLOW_B64 = (
        "data:image/svg+xml;base64,"
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI5MDAiIGhlaWdodD0iNTAwIj48cmVjdCB3aWR0aD0iOTAwIiBoZWlnaHQ9IjUwMCIgZmlsbD0iI0Y2RjFFNyIvPjx0ZXh0IHg9IjQ1MCIgeT0iMjUwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LXNpemU9IjIwIiBmaWxsPSIjMTIzMDJBIiBmb250LWZhbWlseT0ic2Fucy1zZXJpZiI+U8ahIElM4bZVIEzDkFkgTOG7hkMgVEhVTkcgVOG7qiBIVU5HPC90ZXh0Pjwvc3ZnPg=="
    )
    # Ảnh dashboard HRM Master
    DASHBOARD_B64 = (
        "data:image/svg+xml;base64,"
        "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI5MDAiIGhlaWdodD0iNTAwIj48cmVjdCB3aWR0aD0iOTAwIiBoZWlnaHQ9IjUwMCIgZmlsbD0iI0Y2RjFFNyIvPjxyZWN0IHg9IjAiIHk9IjAiIHdpZHRoPSIyMDAiIGhlaWdodD0iNTAwIiBmaWxsPSIjMUY0QjQzIi8+PHRleHQgeD0iMTAwIiB5PSI1MCIgdGV4dC1hbmNob3I9Im1pZGRsZSIgZm9udC1zaXplPSIxNCIgZmlsbD0iI0ZGRiIgZm9udC1mYW1pbHk9InNhbnMtc2VyaWYiPkhSTSBXb3Jrc3BhY2U8L3RleHQ+PHRleHQgeD0iNDUwIiB5PSIyNTAiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtc2l6ZT0iMjAiIGZpbGw9IiMxMjMwMkEiIGZvbnQtZmFtaWx5PSJzYW5zLXNlcmlmIj5EYXNoYm9hcmQgSEtNIE1hc3RlcjwvdGV4dD48L3N2Zz4="
    )

    # ========================================================================
    # HTML/CSS LANDING PAGE — PORTFOLIO OF IMPACT
    # ========================================================================
    _landing_html = f"""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Be+Vietnam+Pro:wght@300;400;500;600;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
  #hrm-portfolio {{
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
    --highlight: #F0B429; /* Màu vàng gold cho điểm nhấn */
    background: var(--paper);
    color: var(--ink);
    font-family: 'Be Vietnam Pro', sans-serif;
    font-size: 14px;
    font-weight: 400;
    line-height: 1.6;
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid var(--line);
    max-width: 1000px;
    margin: 0 auto;
  }}
  #hrm-portfolio * {{ box-sizing: border-box; }}
  #hrm-portfolio img {{ max-width: 100%; display: block; }}
  #hrm-portfolio .hp-wrap {{ max-width: 900px; margin: 0 auto; padding: 0 26px; }}
  #hrm-portfolio h1, #hrm-portfolio h2, #hrm-portfolio h3, #hrm-portfolio h4 {{
    font-family: 'Fraunces', serif;
    font-weight: 600;
    margin: 0;
    color: var(--pine-deep);
    line-height: 1.25;
  }}
  #hrm-portfolio .hp-kicker {{
    font-family: 'Space Mono', monospace;
    font-size: 11px;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: var(--rust);
    margin-bottom: 10px;
    font-weight: 400;
  }}
  #hrm-portfolio .hp-kicker::before {{ content: "— "; color: var(--amber); }}

  /* ---------- HERO ---------- */
  #hrm-portfolio .hp-hero {{
    padding: 50px 0 40px;
    background: linear-gradient(180deg, #fff 0%, var(--paper) 100%);
    border-bottom: 1px solid var(--line);
    position: relative;
  }}
  #hrm-portfolio .hp-hero::after {{
    content: "";
    position: absolute;
    bottom: -1px; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, var(--amber), var(--rust), var(--pine));
  }}
  #hrm-portfolio .hp-hero-grid {{
    display: flex; gap: 40px; align-items: center;
  }}
  #hrm-portfolio .hp-portrait {{
    flex: 0 0 auto;
    width: 170px; height: 170px;
    border-radius: 50%;
    overflow: hidden;
    border: 4px solid #fff;
    box-shadow: 0 12px 40px rgba(18,48,42,.20), 0 0 0 1px var(--line);
    transition: transform .3s ease;
  }}
  #hrm-portfolio .hp-portrait:hover {{ transform: scale(1.03); }}
  #hrm-portfolio .hp-portrait img {{ width: 100%; height: 100%; object-fit: cover; }}
  #hrm-portfolio .hp-hero-text {{ flex: 1 1 auto; }}
  #hrm-portfolio .hp-hero h1 {{
    font-size: clamp(26px, 3.5vw, 38px);
    font-weight: 700;
    letter-spacing: -.02em;
    margin-bottom: 6px;
  }}
  #hrm-portfolio .hp-hero h1 .sub {{
    display: block;
    font-family: 'Be Vietnam Pro', sans-serif;
    font-size: 16px;
    font-weight: 500;
    color: var(--pine);
    letter-spacing: 0;
    margin-top: 4px;
  }}
  #hrm-portfolio .hp-tagline {{
    margin-top: 12px;
    font-size: 15px;
    font-weight: 400;
    color: var(--ink-soft);
    line-height: 1.5;
    max-width: 55ch;
  }}
  #hrm-portfolio .hp-tagline b {{ color: var(--pine-deep); font-weight: 600; }}
  #hrm-portfolio .hp-meta {{
    margin-top: 16px;
    display: flex; flex-wrap: wrap; gap: 10px 22px;
    font-family: 'Space Mono', monospace;
    font-size: 11.5px;
    color: var(--ink-soft);
  }}
  #hrm-portfolio .hp-meta span {{ display: inline-flex; align-items: center; gap: 6px; }}
  #hrm-portfolio .hp-cta-row {{ margin-top: 22px; display: flex; gap: 14px; flex-wrap: wrap; }}
  #hrm-portfolio .hp-btn {{
    font-family: 'Be Vietnam Pro', sans-serif;
    font-weight: 600;
    font-size: 13px;
    padding: 11px 24px;
    border-radius: var(--radius);
    border: 1px solid transparent;
    display: inline-flex; align-items: center; gap: 8px;
    background: var(--pine);
    color: var(--paper);
    box-shadow: 3px 3px 0 var(--amber);
    text-decoration: none;
    transition: all .2s ease;
  }}
  #hrm-portfolio .hp-btn:hover {{ transform: translate(-1px, -1px); box-shadow: 4px 4px 0 var(--amber); }}
  #hrm-portfolio .hp-btn.ghost {{
    background: transparent; color: var(--pine-deep);
    border: 1px solid var(--pine-deep); box-shadow: 3px 3px 0 var(--paper-deep);
  }}
  #hrm-portfolio .hp-btn.ghost:hover {{ background: var(--paper-deep); }}

  /* ---------- SECTION COMMON ---------- */
  #hrm-portfolio section.hp-sec {{ padding: 40px 0; border-bottom: 1px solid var(--line); }}
  #hrm-portfolio .hp-sec:last-of-type {{ border-bottom: none; }}
  #hrm-portfolio .hp-sec h2 {{
    font-size: clamp(22px, 2.8vw, 28px);
    margin-bottom: 16px;
    font-weight: 700;
  }}
  #hrm-portfolio .hp-sec h2 em {{ font-style: italic; color: var(--amber); }}

  /* ---------- HIGHLIGHT METRICS ---------- */
  #hrm-portfolio .hp-metrics-grid {{
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 14px; margin: 24px 0;
  }}
  #hrm-portfolio .hp-metric-card {{
    background: #fff;
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 20px 18px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,.04);
    transition: transform .2s ease, box-shadow .2s ease;
  }}
  #hrm-portfolio .hp-metric-card:hover {{
    transform: translateY(-3px);
    box-shadow: 0 8px 20px rgba(18,48,42,.10);
  }}
  #hrm-portfolio .hp-metric-num {{
    font-family: 'Fraunces', serif;
    font-size: 32px;
    font-weight: 700;
    color: var(--pine-deep);
    line-height: 1;
    margin-bottom: 6px;
  }}
  #hrm-portfolio .hp-metric-num small {{ font-size: 18px; font-weight: 500; color: var(--amber); }}
  #hrm-portfolio .hp-metric-label {{
    font-size: 12px;
    color: var(--ink-soft);
    font-weight: 400;
    line-height: 1.4;
  }}

  /* ---------- COMPETENCY GRID ---------- */
  #hrm-portfolio .hp-comp-grid {{
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-top: 24px;
  }}
  #hrm-portfolio .hp-comp-card {{
    background: #fff;
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 20px 18px;
    position: relative;
    transition: all .2s ease;
  }}
  #hrm-portfolio .hp-comp-card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(18,48,42,.08);
    border-color: var(--amber);
  }}
  #hrm-portfolio .hp-comp-num {{
    font-family: 'Space Mono', monospace;
    font-size: 26px;
    color: var(--paper-deep);
    font-weight: 700;
    position: absolute; top: 12px; right: 16px; line-height: 1;
  }}
  #hrm-portfolio .hp-comp-ic {{ font-size: 24px; display: block; margin-bottom: 12px; }}
  #hrm-portfolio .hp-comp-card h4 {{
    font-size: 14.5px;
    margin: 0 0 8px;
    padding-right: 36px;
    font-weight: 600;
  }}
  #hrm-portfolio .hp-comp-card p {{
    margin: 0;
    font-size: 12.5px;
    color: var(--ink-soft);
    font-weight: 300;
    line-height: 1.6;
  }}

  /* ---------- SYSTEM THINKING / ARCHITECTURE ---------- */
  #hrm-portfolio .hp-sys {{
    background: var(--pine-deep);
    color: #EFE8D8;
    border-radius: 12px;
    padding: 32px 28px;
    margin-top: 20px;
    position: relative;
    overflow: hidden;
  }}
  #hrm-portfolio .hp-sys::before {{
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg, var(--amber), var(--rust));
  }}
  #hrm-portfolio .hp-sys h3 {{ color: #fff; font-size: 20px; margin-bottom: 14px; }}
  #hrm-portfolio .hp-sys p {{ color: #c9d8d3; font-size: 14px; line-height: 1.7; margin: 0 0 14px; }}
  #hrm-portfolio .hp-sys ul {{ margin: 0; padding-left: 20px; color: #c9d8d3; font-size: 13.5px; line-height: 1.8; }}
  #hrm-portfolio .hp-sys ul li b {{ color: #fff; }}
  #hrm-portfolio .hp-sys img {{
    border-radius: 8px; margin-top: 20px; border: 1px solid rgba(255,255,255,.15);
    box-shadow: 0 8px 24px rgba(0,0,0,.25);
  }}

  /* ---------- AI AGENT ROLE ---------- */
  #hrm-portfolio .hp-ai-role {{
    background: linear-gradient(135deg, #fdf6e3 0%, #fff 100%);
    border-left: 5px solid var(--amber);
    border-radius: 8px;
    padding: 26px 24px;
    margin-top: 20px;
  }}
  #hrm-portfolio .hp-ai-role h4 {{
    font-size: 17px; margin-bottom: 12px; color: var(--pine-deep);
    display: flex; align-items: center; gap: 10px;
  }}
  #hrm-portfolio .hp-ai-role p {{ margin: 0 0 12px; font-size: 14px; line-height: 1.7; }}
  #hrm-portfolio .hp-ai-role ul {{ margin: 0; padding-left: 20px; font-size: 13.5px; line-height: 1.8; }}
  #hrm-portfolio .hp-ai-role ul li b {{ color: var(--pine-deep); }}

  /* ---------- TIMELINE ---------- */
  #hrm-portfolio .hp-timeline {{ position: relative; padding-left: 30px; margin-top: 24px; }}
  #hrm-portfolio .hp-timeline::before {{
    content: ""; position: absolute; left: 7px; top: 8px; bottom: 8px; width: 2px;
    background: repeating-linear-gradient(to bottom, var(--pine) 0 6px, transparent 6px 10px);
  }}
  #hrm-portfolio .hp-tl-item {{ position: relative; padding-bottom: 28px; }}
  #hrm-portfolio .hp-tl-item:last-child {{ padding-bottom: 0; }}
  #hrm-portfolio .hp-tl-item::before {{
    content: ""; position: absolute; left: -30px; top: 4px;
    width: 14px; height: 14px; border-radius: 50%;
    background: var(--paper); border: 3px solid var(--rust);
    z-index: 1;
  }}
  #hrm-portfolio .hp-tl-item.milestone::before {{ border-color: var(--amber); background: var(--amber); }}
  #hrm-portfolio .hp-tl-date {{
    font-family: 'Space Mono', monospace; font-size: 11px; color: var(--rust);
    background: #fff; display: inline-block; padding: 3px 10px; border-radius: 20px;
    border: 1px solid var(--line); margin-bottom: 8px; font-weight: 400;
  }}
  #hrm-portfolio .hp-tl-item h3 {{ font-size: 16px; margin-bottom: 4px; font-weight: 600; }}
  #hrm-portfolio .hp-tl-item .hp-role {{
    font-size: 12px; color: var(--amber); font-weight: 600;
    font-family: 'Space Mono', monospace; margin-bottom: 6px;
    text-transform: uppercase; letter-spacing: .04em;
  }}
  #hrm-portfolio .hp-tl-item p {{ margin: 0; font-size: 13px; color: var(--ink-soft); font-weight: 300; line-height: 1.6; }}

  /* ---------- RESULT / IMPACT ---------- */
  #hrm-portfolio .hp-impact {{
    background: #fff; border: 2px solid var(--amber); border-radius: 12px;
    padding: 30px 28px; margin-top: 20px; text-align: center;
  }}
  #hrm-portfolio .hp-impact h3 {{ font-size: 22px; margin-bottom: 10px; color: var(--pine-deep); }}
  #hrm-portfolio .hp-impact .big-num {{
    font-family: 'Fraunces', serif; font-size: 52px; font-weight: 700;
    color: var(--rust); line-height: 1; margin: 16px 0 8px;
  }}
  #hrm-portfolio .hp-impact p {{ font-size: 14px; color: var(--ink-soft); max-width: 60ch; margin: 0 auto; }}

  /* ---------- FOOTER ---------- */
  #hrm-portfolio footer {{
    padding: 20px 0 24px; text-align: center;
    font-family: 'Space Mono', monospace; font-size: 10.5px;
    color: var(--ink-soft); background: var(--paper-deep);
  }}

  /* ---------- RESPONSIVE ---------- */
  @media (max-width: 720px) {{
    #hrm-portfolio .hp-hero-grid {{ flex-direction: column; text-align: center; }}
    #hrm-portfolio .hp-portrait {{ margin: 0 auto; }}
    #hrm-portfolio .hp-tagline {{ margin-left: auto; margin-right: auto; }}
    #hrm-portfolio .hp-cta-row {{ justify-content: center; }}
    #hrm-portfolio .hp-comp-grid {{ grid-template-columns: 1fr; }}
    #hrm-portfolio .hp-metrics-grid {{ grid-template-columns: 1fr 1fr; }}
  }}
</style>

<div id="hrm-portfolio">

  <!-- ================= HERO ================= -->
  <section class="hp-hero">
    <div class="hp-wrap">
      <div class="hp-hero-grid">
        <div class="hp-portrait">
          <img src="{PORTRAIT_B64}" alt="Chân dung Nguyễn Văn Tuyến">
        </div>
        <div class="hp-hero-text">
          <div class="hp-kicker">Portfolio of Impact · AI Thực Chiến</div>
          <h1>
            Nguyễn Văn Tuyến
            <span class="sub">Từ bài toán quản trị đến sản phẩm công nghệ đang vận hành</span>
          </h1>
          <p class="hp-tagline">
            Tôi không đến với AI bằng lý thuyết. Tôi đến bằng một <b>sản phẩm thật</b> — 
            <b>HRM Master</b> — được xây dựng để giải quyết chính bài toán nhân sự tôi đang gặp hàng ngày, 
            với <b>hơn 15.000 dòng code</b> do AI Agent viết dưới sự chỉ đạo của tôi.
          </p>
          <div class="hp-meta">
            <span>📞 0961 778 150</span>
            <span>✉️ nguyen.support999@gmail.com</span>
            <span>🎂 1977</span>
          </div>
          <div class="hp-cta-row">
            <a class="hp-btn" href="http://demo-hrm9.streamlit.app" target="_blank" rel="noopener">→ Trải nghiệm HRM Master</a>
            <a class="hp-btn ghost" href="#" onclick="alert('Anh có thể upload file CV PDF lên Google Drive và thay link vào đây.'); return false;">⬇ Tải CV đầy đủ (PDF)</a>
          </div>
        </div>
      </div>
    </div>
  </section>

  <!-- ================= METRICS / IMPACT ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-kicker">Kết quả đo lường được</div>
      <h2>Không phải dự án mô phỏng. <em>Đây là kết quả thật.</em></h2>
      <div class="hp-metrics-grid">
        <div class="hp-metric-card">
          <div class="hp-metric-num">15.000<small>+</small></div>
          <div class="hp-metric-label">Dòng code do AI Agent viết dưới sự chỉ đạo của tôi</div>
        </div>
        <div class="hp-metric-card">
          <div class="hp-metric-num">40<small>%</small></div>
          <div class="hp-metric-label">Thời gian xử lý thủ tục hành chính được tối ưu</div>
        </div>
        <div class="hp-metric-card">
          <div class="hp-metric-num">80<small>+</small></div>
          <div class="hp-metric-label">Nhân viên đang sử dụng hệ thống hàng ngày</div>
        </div>
        <div class="hp-metric-card">
          <div class="hp-metric-num">10<small>+</small></div>
          <div class="hp-metric-label">Module nghiệp vụ đã vận hành thực tế</div>
        </div>
      </div>
    </div>
  </section>

  <!-- ================= CORE COMPETENCIES ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-kicker">Năng lực cốt lõi</div>
      <h2>6 điểm mạnh <em>được chứng minh</em> qua sản phẩm</h2>
      <div class="hp-comp-grid">
        <div class="hp-comp-card">
          <span class="hp-comp-num">01</span>
          <span class="hp-comp-ic">🧠</span>
          <h4>Tư duy hệ thống</h4>
          <p>Nhìn bài toán nhân sự như một hệ thống các thực thể có quan hệ chặt chẽ: Nhân viên → Hợp đồng → BHXH → Chấm công → Lương → Thuế. Thiết kế kiến trúc dữ liệu chuẩn hóa, đảm bảo tính toàn vẹn và khả năng mở rộng.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-num">02</span>
          <span class="hp-comp-ic">⚙️</span>
          <h4>Chuyển bài toán quản trị thành sản phẩm</h4>
          <p>Không dừng ở phân tích nghiệp vụ. Tôi trực tiếp chỉ đạo AI Agent xây dựng sản phẩm hoàn chỉnh: từ database schema, business logic, đến giao diện người dùng và triển khai cloud.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-num">03</span>
          <span class="hp-comp-ic">📊</span>
          <h4>Phân tích nghiệp vụ & Quản trị</h4>
          <p>Hơn 20 năm kinh nghiệm quản trị (13 năm điều hành doanh nghiệp, 10 năm tại Tập đoàn Bảo Việt). Hiểu sâu bài toán doanh nghiệp từ góc nhìn người làm thật, không phải góc nhìn lý thuyết.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-num">04</span>
          <span class="hp-comp-ic">🤖</span>
          <h4>Chỉ đạo AI Agent thực chiến</h4>
          <p>Sử dụng thành thạo Claude, ChatGPT như những "kỹ sư ảo". Biết cách phân rã bài toán, viết prompt hiệu quả, kiểm chứng kết quả và tích hợp code do AI tạo ra vào hệ thống thật.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-num">05</span>
          <span class="hp-comp-ic">📈</span>
          <h4>Đo lường & Tối ưu</h4>
          <p>Đặt mục tiêu đo lường được cho mọi tính năng: giảm 40% thời gian xử lý thủ tục, tự động hóa báo cáo BHXH, giảm sai sót trong tính lương. Kết quả được kiểm chứng qua vận hành thực tế.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-num">06</span>
          <span class="hp-comp-ic">🌱</span>
          <h4>Tự học & Thích ứng nhanh</h4>
          <p>Bắt đầu học Vibe Coding từ con số 0 ở tuổi 49. Trong 20 ngày, hoàn thành bản thiết kế quy trình đầu tiên và đưa vào áp dụng thực tế. Tinh thần học hỏi không giới hạn bởi tuổi tác.</p>
        </div>
      </div>
    </div>
  </section>

  <!-- ================= SYSTEM ARCHITECTURE ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-kicker">Tư duy hệ thống</div>
      <h2>Kiến trúc hệ thống <em>HRM Master</em></h2>
      <div class="hp-sys">
        <h3>Thiết kế để mở rộng và vận hành thật</h3>
        <p>HRM Master không phải là một script đơn lẻ. Đây là một hệ thống hoàn chỉnh với kiến trúc phân tầng rõ ràng, được thiết kế để phục vụ nhiều doanh nghiệp (multi-tenant) và có thể mở rộng theo nhu cầu.</p>
        <ul>
          <li><b>Database Layer:</b> PostgreSQL (Supabase) — chuẩn hóa dữ liệu nhân sự, hợp đồng, BHXH, chấm công, lương.</li>
          <li><b>Application Layer:</b> Python + Streamlit — xử lý nghiệp vụ, tính toán lương, xuất báo cáo.</li>
          <li><b>Storage Layer:</b> Supabase Storage — lưu trữ hồ sơ, ảnh nhân viên, tài liệu công văn.</li>
          <li><b>AI Layer:</b> Anthropic Claude API — Chatbot tư vấn pháp luật lao động, trợ lý ảo cho HR.</li>
          <li><b>Multi-tenant:</b> Mỗi doanh nghiệp có database riêng, landing page riêng, cấu hình riêng.</li>
        </ul>
        <img src="{ARCHITECTURE_B64}" alt="Sơ đồ kiến trúc HRM Master">
      </div>
    </div>
  </section>

  <!-- ================= DATA FLOW ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-kicker">Luồng dữ liệu</div>
      <h2>Từ dữ liệu thô đến <em>quyết định quản trị</em></h2>
      <div class="hp-sys" style="background: var(--paper-deep); color: var(--ink);">
        <h3 style="color: var(--pine-deep);">Dữ liệu chảy qua hệ thống như thế nào?</h3>
        <p style="color: var(--ink-soft);">Mỗi thao tác của người dùng đều được xử lý qua một luồng logic chặt chẽ, đảm bảo tính chính xác và khả năng truy vết.</p>
        <ul style="color: var(--ink-soft);">
          <li><b>Bước 1:</b> HR nhập thông tin nhân viên mới → Hệ thống tự sinh Mã NV, tính ngày bắt đầu đóng BHXH theo quy tắc 14 ngày.</li>
          <li><b>Bước 2:</b> Chấm công hàng ngày (thủ công hoặc Face ID) → Dữ liệu được lưu vào bảng cham_cong.</li>
          <li><b>Bước 3:</b> Cuối tháng, Payroll Engine tổng hợp chấm công, tính lương, phụ cấp, khấu trừ BHXH/thuế TNCN.</li>
          <li><b>Bước 4:</b> Hệ thống tự động tạo báo cáo D02-LT (tăng/giảm BHXH), báo cáo trích nộp, sổ quản lý lao động.</li>
          <li><b>Bước 5:</b> Dashboard hiển thị các chỉ số quản trị cho Ban lãnh đạo.</li>
        </ul>
        <img src="{DATAFLOW_B64}" alt="Sơ đồ luồng dữ liệu HRM Master" style="border: 1px solid var(--line);">
      </div>
    </div>
  </section>

  <!-- ================= AI AGENT ROLE ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-kicker">Vai trò thực tế của AI Agent</div>
      <h2>Tôi <em>chỉ đạo</em> AI, không phụ thuộc vào AI</h2>
      <div class="hp-ai-role">
        <h4>🤖 AI Agent đã làm gì trong dự án này?</h4>
        <p>AI Agent (Claude, ChatGPT) là công cụ để tôi hiện thực hóa tư duy của mình. Tôi là người <b>phân tích nghiệp vụ</b>, <b>thiết kế giải pháp</b>, và <b>kiểm chứng kết quả</b>. AI là người viết code.</p>
        <ul>
          <li><b>Phân rã bài toán:</b> Tôi chia nhỏ hệ thống thành 10+ module độc lập, mỗi module có input/output rõ ràng.</li>
          <li><b>Viết prompt hiệu quả:</b> Tôi mô tả chính xác yêu cầu nghiệp vụ bằng ngôn ngữ tự nhiên, kèm ví dụ và ràng buộc.</li>
          <li><b>Kiểm chứng kết quả:</b> Mọi code do AI tạo ra đều được tôi kiểm tra logic, chạy thử và đối chiếu với nghiệp vụ thực tế.</li>
          <li><b>Tích hợp hệ thống:</b> Tôi kết nối các module, xử lý lỗi, tối ưu hiệu năng và triển khai lên cloud.</li>
          <li><b>Ra quyết định kiến trúc:</b> Tôi quyết định dùng database nào, lưu trữ ra sao, bảo mật thế nào.</li>
        </ul>
        <p style="margin-top: 14px; font-style: italic; color: var(--pine);">
          "AI có thể viết code, nhưng không thể hiểu bài toán quản trị bằng người trong cuộc. 
          Đó là lợi thế của tôi."
        </p>
      </div>
    </div>
  </section>

  <!-- ================= MODULES / FEATURES ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-kicker">Chức năng đã vận hành</div>
      <h2>10+ module <em>đang chạy thật</em> mỗi ngày</h2>
      <div class="hp-comp-grid" style="grid-template-columns: repeat(2, 1fr);">
        <div class="hp-comp-card">
          <span class="hp-comp-ic">🗂️</span>
          <h4>Quản lý Hồ sơ nhân viên</h4>
          <p>Số hóa toàn bộ hồ sơ: thông tin cá nhân, hợp đồng, quá trình công tác, BHXH, ngân hàng. Thay thế hoàn toàn Excel.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">📄</span>
          <h4>Hợp đồng & Quyết định nhân sự</h4>
          <p>Tự sinh số HĐLĐ/HĐTV, tạo file Word theo mẫu Nghị định 30/2020/NĐ-CP, quản lý lịch sử bổ nhiệm, điều chuyển.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">📋</span>
          <h4>BHXH & Báo cáo D02-LT</h4>
          <p>Tự tính ngày bắt đầu/kết thúc đóng BHXH theo quy tắc 14 ngày. Xuất báo cáo D02-LT chuẩn iCare, sổ quản lý lao động.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">🕒</span>
          <h4>Chấm công & Tính lương</h4>
          <p>Chấm công thủ công hoặc Face ID. Payroll Engine tự động tính lương, phụ cấp, khấu trừ theo đúng chính sách công ty.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">💰</span>
          <h4>Tính thu nhập & Thuế TNCN</h4>
          <p>Tự động tính thuế TNCN theo biểu lũy tiến, giảm trừ gia cảnh, quyết toán thuế cuối năm.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">📄</span>
          <h4>Quản lý Công văn & HĐ kinh tế</h4>
          <p>Lưu trữ, tra cứu công văn đến/đi, hợp đồng kinh tế. Tự sinh số công văn theo cấu hình từng doanh nghiệp.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">💬</span>
          <h4>Chat nội bộ</h4>
          <p>Kênh trao đổi nội bộ theo phòng ban, chat 1:1, gửi file. Tự đồng bộ thành viên khi nhân viên điều chuyển.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">🤖</span>
          <h4>Chatbot Tư vấn Pháp luật</h4>
          <p>Trợ lý AI trả lời câu hỏi về BHXH, thuế, hợp đồng lao động, có trích dẫn điều luật cụ thể.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">📊</span>
          <h4>Dashboard Quản trị</h4>
          <p>Biểu đồ trực quan về cơ cấu nhân sự, độ tuổi, giới tính, xu hướng tuyển dụng. Hỗ trợ ra quyết định.</p>
        </div>
        <div class="hp-comp-card">
          <span class="hp-comp-ic">🔍</span>
          <h4>Audit & Phân quyền</h4>
          <p>Ghi log mọi thao tác thay đổi dữ liệu. Phân quyền chi tiết theo vai trò: Admin, HR, Kế toán, Văn thư, Viewer.</p>
        </div>
      </div>
    </div>
  </section>

  <!-- ================= TIMELINE ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-kicker">Hành trình</div>
      <h2>Từ Excel đến <em>sản phẩm AI thực chiến</em></h2>
      <div class="hp-timeline">
        <div class="hp-tl-item milestone">
          <div class="hp-tl-date">02/05/2026</div>
          <h3>Ngày đầu làm chuyên viên HR</h3>
          <p>Tiếp nhận công việc quản trị nhân sự, kế thừa thói quen quản lý hồ sơ bằng Excel. Nhận ra quy trình cần được chuẩn hóa lại từ gốc.</p>
        </div>
        <div class="hp-tl-item">
          <div class="hp-tl-date">Tháng 05/2026</div>
          <h3>Bắt đầu học Vibe Coding</h3>
          <p>Chưa từng viết một dòng code. Học từng khái niệm, từng bước, song song với công việc HR hàng ngày — để tự xây công cụ giải quyết đúng vấn đề mình đang gặp.</p>
        </div>
        <div class="hp-tl-item milestone">
          <div class="hp-tl-date">Sau 20 ngày</div>
          <h3>Hoàn thành bản thiết kế quy trình đầu tiên</h3>
          <p>Đưa ngay vào áp dụng thực tế với hồ sơ nhân sự của công ty — không phải bản demo, mà là dữ liệu thật, quy trình thật.</p>
        </div>
        <div class="hp-tl-item milestone">
          <div class="hp-tl-date">Đến nay · 07/2026</div>
          <h3>HRM Master — sẵn sàng cho nhiều doanh nghiệp</h3>
          <p>Từ một công cụ nội bộ, phát triển thành nền tảng quản trị nhân sự đa doanh nghiệp (multi-tenant), với hơn 15.000 dòng code do AI Agent viết dưới sự chỉ đạo của tôi.</p>
        </div>
      </div>
    </div>
  </section>

  <!-- ================= IMPACT STATEMENT ================= -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <div class="hp-impact">
        <h3>Tôi không chỉ học AI. Tôi đã <em>làm ra sản phẩm</em> bằng AI.</h3>
        <div class="big-num">40%</div>
        <p>Thời gian xử lý thủ tục hành chính được tối ưu. Không phải con số ước lượng — đây là kết quả đo lường được sau khi đưa HRM Master vào vận hành thực tế tại doanh nghiệp.</p>
        <p style="margin-top: 16px; font-weight: 600; color: var(--pine-deep);">
          "Giá trị của AI không nằm ở trí tuệ nhân tạo, mà nằm ở trí tuệ, khát vọng và trách nhiệm 
          của những người tạo ra nó."
        </p>
      </div>
    </div>
  </section>

  <footer>
    © 2026 Nguyễn Văn Tuyến · Portfolio of Impact · Được xây dựng bằng AI Thực Chiến
  </footer>

</div>

<script>
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