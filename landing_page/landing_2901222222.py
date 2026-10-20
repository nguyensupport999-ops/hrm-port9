# -*- coding: utf-8 -*-
"""
landing_page/landing_demo.py
=============================
Landing Page CV cá nhân — NGUYỄN VĂN TUYẾN
Phiên bản "Executive Portfolio" — Thiết kế cho vị trí Trưởng phòng HCNS.
"""

import streamlit as st
import streamlit.components.v1 as components

def render():
    """Landing Page CV — Executive Portfolio."""

    # ========================================================================
    # ẢNH BASE64 (Anh giữ nguyên các biến này, hoặc thay bằng ảnh thật)
    # Lưu ý: Tôi đã rút gọn chuỗi Base64 ở đây để dễ nhìn. 
    # Khi chạy thực tế, anh copy lại chuỗi Base64 gốc của anh vào biến PORTRAIT_B64.
    # ========================================================================
    # PORTRAIT_B64 = "data:image/jpeg;base64,..." (Giữ nguyên ảnh chân dung của anh)
    # DATAFLOW_B64 = "data:image/png;base64,..." (Giữ nguyên sơ đồ luồng dữ liệu)

    # TẠM THỜI DÙNG PLACEHOLDER ĐỂ CODE CHẠY ĐƯỢC (Anh nhớ thay bằng ảnh thật)
    PORTRAIT_B64 = "" 
    DATAFLOW_B64 = ""

    # ========================================================================
    # LINKS
    # ========================================================================
    CV_PDF_URL = "https://drive.google.com/file/d/1fT9L9wD0B56IBSaatUA3chHESR_qSlfU/view?usp=sharing"
    CERT_URL = "https://drive.google.com/file/d/1iIGPnp_71rhRtX4h_5oys5XdMByygQha/view?usp=drive_link"
    DEMO_URL = "http://demo-hrm9.streamlit.app"
    CONTACT_EMAIL = "duhocanphuloc@gmail.com"
    PHONE = "0961 778 150"

    _demo_btn = (
        f'<a class="hp-btn primary" href="{DEMO_URL}" target="_blank" rel="noopener noreferrer">'
        f'🚀 Trải nghiệm HRM Master (Demo)</a>'
    )
    _cv_btn = (
        f'<a class="hp-btn ghost" href="{CV_PDF_URL}" target="_blank" rel="noopener noreferrer">'
        f'📄 Tải CV đầy đủ (PDF)</a>'
    )

    # ========================================================================
    # CSS — PHONG CÁCH EXECUTIVE (NAVY & GOLD)
    # ========================================================================
    _css = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600;700&family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  #hrm-portfolio {
    --navy: #0F172A;
    --navy-light: #1E293B;
    --gold: #C5A059;
    --gold-light: #E2C792;
    --bg: #F8FAFC;
    --text: #334155;
    --text-light: #64748B;
    --border: #E2E8F0;
    
    background: var(--bg);
    color: var(--text);
    font-family: 'Inter', sans-serif;
    font-size: 15px;
    line-height: 1.6;
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid var(--border);
    max-width: 1100px;
    margin: 0 auto;
    box-shadow: 0 10px 30px rgba(0,0,0,0.05);
  }
  
  #hrm-portfolio * { box-sizing: border-box; }
  #hrm-portfolio img { max-width: 100%; display: block; }
  #hrm-portfolio .hp-wrap { max-width: 900px; margin: 0 auto; padding: 0 30px; }
  
  #hrm-portfolio h1, #hrm-portfolio h2, #hrm-portfolio h3, #hrm-portfolio h4 {
    font-family: 'Playfair Display', serif;
    color: var(--navy);
    font-weight: 600;
    margin: 0;
  }

  /* ---------- HERO SECTION ---------- */
  #hrm-portfolio .hp-hero {
    background: var(--navy);
    color: white;
    padding: 60px 0;
    position: relative;
  }
  #hrm-portfolio .hp-hero::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background: radial-gradient(circle at 80% 20%, rgba(197, 160, 89, 0.15) 0%, transparent 60%);
    pointer-events: none;
  }
  #hrm-portfolio .hp-hero-grid {
    display: flex; gap: 50px; align-items: center; position: relative; z-index: 1;
  }
  #hrm-portfolio .hp-portrait {
    flex: 0 0 auto;
    width: 200px; height: 240px;
    border-radius: 8px;
    overflow: hidden;
    border: 2px solid var(--gold);
    box-shadow: 0 20px 40px rgba(0,0,0,0.4);
  }
  #hrm-portfolio .hp-portrait img { width: 100%; height: 100%; object-fit: cover; }
  #hrm-portfolio .hp-hero-text { flex: 1 1 auto; }
  #hrm-portfolio .hp-kicker {
    font-family: 'Inter', sans-serif;
    font-size: 12px; letter-spacing: 0.15em; text-transform: uppercase;
    color: var(--gold); font-weight: 600; margin-bottom: 12px;
  }
  #hrm-portfolio .hp-hero h1 {
    color: white; font-size: 42px; line-height: 1.2; margin-bottom: 10px;
  }
  #hrm-portfolio .hp-hero h1 span.sub {
    display: block; font-family: 'Inter', sans-serif;
    font-size: 18px; color: var(--gold-light); font-weight: 400; margin-top: 8px;
    letter-spacing: 0.05em;
  }
  #hrm-portfolio .hp-tagline {
    font-size: 16px; color: #CBD5E1; margin-top: 20px; max-width: 60ch;
    line-height: 1.7; font-weight: 300;
  }
  #hrm-portfolio .hp-tagline b { color: white; font-weight: 500; }
  
  #hrm-portfolio .hp-cta-row { margin-top: 30px; display: flex; gap: 15px; flex-wrap: wrap; }
  #hrm-portfolio .hp-btn {
    font-family: 'Inter', sans-serif; font-weight: 500; font-size: 14px;
    padding: 12px 24px; border-radius: 6px; text-decoration: none;
    display: inline-flex; align-items: center; gap: 8px; transition: all 0.3s ease;
  }
  #hrm-portfolio .hp-btn.primary {
    background: var(--gold); color: var(--navy); border: 1px solid var(--gold);
  }
  #hrm-portfolio .hp-btn.primary:hover {
    background: var(--gold-light); transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(197, 160, 89, 0.3);
  }
  #hrm-portfolio .hp-btn.ghost {
    background: transparent; color: white; border: 1px solid rgba(255,255,255,0.3);
  }
  #hrm-portfolio .hp-btn.ghost:hover {
    border-color: var(--gold); color: var(--gold);
  }

  /* ---------- METRICS ---------- */
  #hrm-portfolio .hp-metrics-sec {
    background: white; padding: 40px 0; border-bottom: 1px solid var(--border);
  }
  #hrm-portfolio .hp-metrics-grid {
    display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px;
  }
  #hrm-portfolio .hp-metric-item {
    text-align: center; padding: 20px; border-right: 1px solid var(--border);
  }
  #hrm-portfolio .hp-metric-item:last-child { border-right: none; }
  #hrm-portfolio .hp-metric-num {
    font-family: 'Playfair Display', serif; font-size: 38px; font-weight: 700;
    color: var(--navy); line-height: 1; margin-bottom: 8px;
  }
  #hrm-portfolio .hp-metric-num span { color: var(--gold); font-size: 24px; }
  #hrm-portfolio .hp-metric-label {
    font-size: 13px; color: var(--text-light); font-weight: 500; text-transform: uppercase; letter-spacing: 0.05em;
  }

  /* ---------- SECTIONS ---------- */
  #hrm-portfolio section.hp-sec { padding: 60px 0; border-bottom: 1px solid var(--border); }
  #hrm-portfolio .hp-sec h2 {
    font-size: 32px; margin-bottom: 10px; text-align: center; color: var(--navy);
  }
  #hrm-portfolio .hp-sec h2::after {
    content: ""; display: block; width: 60px; height: 3px; background: var(--gold); margin: 15px auto 0;
  }
  #hrm-portfolio .hp-sec-subtitle {
    text-align: center; color: var(--text-light); font-size: 15px; margin-bottom: 40px;
  }

  /* ---------- COMPETENCY CARDS ---------- */
  #hrm-portfolio .hp-comp-grid {
    display: grid; grid-template-columns: repeat(2, 1fr); gap: 24px;
  }
  #hrm-portfolio .hp-comp-card {
    background: white; padding: 30px; border-radius: 8px;
    border: 1px solid var(--border); box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    transition: border-color 0.3s ease, box-shadow 0.3s ease;
  }
  #hrm-portfolio .hp-comp-card:hover {
    border-color: var(--gold); box-shadow: 0 10px 20px rgba(0,0,0,0.05);
  }
  #hrm-portfolio .hp-comp-icon { font-size: 28px; margin-bottom: 15px; }
  #hrm-portfolio .hp-comp-card h4 { font-size: 18px; margin-bottom: 10px; color: var(--navy); }
  #hrm-portfolio .hp-comp-card p { font-size: 14px; color: var(--text-light); line-height: 1.7; margin: 0; }

  /* ---------- SYSTEM ARCHITECTURE ---------- */
  #hrm-portfolio .hp-sys-box {
    background: var(--navy-light); color: white; padding: 40px; border-radius: 12px;
    margin-top: 30px; border-left: 4px solid var(--gold);
  }
  #hrm-portfolio .hp-sys-box h3 { color: var(--gold-light); font-size: 22px; margin-bottom: 15px; }
  #hrm-portfolio .hp-sys-box p { color: #CBD5E1; font-size: 15px; line-height: 1.8; margin-bottom: 20px; }
  #hrm-portfolio .hp-sys-box ul { color: #CBD5E1; font-size: 14px; line-height: 1.8; padding-left: 20px; margin: 0 0 20px 0; }
  #hrm-portfolio .hp-sys-box li { margin-bottom: 8px; }
  #hrm-portfolio .hp-sys-box li b { color: white; }
  #hrm-portfolio .hp-sys-box img { border-radius: 8px; margin-top: 20px; border: 1px solid rgba(255,255,255,0.1); }

  /* ---------- TIMELINE ---------- */
  #hrm-portfolio .hp-timeline { position: relative; max-width: 800px; margin: 0 auto; }
  #hrm-portfolio .hp-timeline::before {
    content: ""; position: absolute; left: 50%; top: 0; bottom: 0; width: 2px;
    background: var(--border); transform: translateX(-50%);
  }
  #hrm-portfolio .hp-tl-item {
    position: relative; width: 50%; padding: 20px 40px; box-sizing: border-box;
  }
  #hrm-portfolio .hp-tl-item:nth-child(odd) { left: 0; text-align: right; }
  #hrm-portfolio .hp-tl-item:nth-child(even) { left: 50%; text-align: left; }
  #hrm-portfolio .hp-tl-item::before {
    content: ""; position: absolute; top: 28px; width: 16px; height: 16px;
    border-radius: 50%; background: white; border: 3px solid var(--gold);
    z-index: 1;
  }
  #hrm-portfolio .hp-tl-item:nth-child(odd)::before { right: -8px; }
  #hrm-portfolio .hp-tl-item:nth-child(even)::before { left: -8px; }
  
  #hrm-portfolio .hp-tl-date {
    font-family: 'Inter', sans-serif; font-size: 12px; font-weight: 600; color: var(--gold);
    text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 5px;
  }
  #hrm-portfolio .hp-tl-item h3 { font-size: 18px; margin-bottom: 5px; color: var(--navy); }
  #hrm-portfolio .hp-tl-item .hp-role { font-size: 14px; color: var(--text-light); margin-bottom: 10px; font-style: italic; }
  #hrm-portfolio .hp-tl-item p { font-size: 14px; color: var(--text); line-height: 1.6; margin: 0; }

  /* ---------- FOOTER ---------- */
  #hrm-portfolio footer {
    background: var(--navy); color: #94A3B8; padding: 30px 0; text-align: center;
    font-family: 'Inter', sans-serif; font-size: 13px;
  }
  #hrm-portfolio footer a { color: var(--gold); text-decoration: none; }
  #hrm-portfolio footer a:hover { text-decoration: underline; }

  /* ---------- RESPONSIVE ---------- */
  @media (max-width: 768px) {
    #hrm-portfolio .hp-hero-grid { flex-direction: column; text-align: center; gap: 30px; }
    #hrm-portfolio .hp-hero h1 { font-size: 32px; }
    #hrm-portfolio .hp-cta-row { justify-content: center; }
    #hrm-portfolio .hp-metrics-grid { grid-template-columns: repeat(2, 1fr); }
    #hrm-portfolio .hp-metric-item { border-right: none; border-bottom: 1px solid var(--border); padding: 15px; }
    #hrm-portfolio .hp-metric-item:nth-child(3), #hrm-portfolio .hp-metric-item:nth-child(4) { border-bottom: none; }
    #hrm-portfolio .hp-comp-grid { grid-template-columns: 1fr; }
    
    #hrm-portfolio .hp-timeline::before { left: 20px; }
    #hrm-portfolio .hp-tl-item { width: 100%; left: 0 !important; text-align: left !important; padding-left: 50px; padding-right: 0; }
    #hrm-portfolio .hp-tl-item::before { left: 12px !important; right: auto !important; }
  }
</style>
"""

    # ========================================================================
    # HTML BODY
    # ========================================================================
    _html_body = f"""
<div id="hrm-portfolio">

  <!-- HERO -->
  <section class="hp-hero">
    <div class="hp-wrap">
      <div class="hp-hero-grid">
        <div class="hp-portrait">
          <!-- Anh thay src bằng biến PORTRAIT_B64 của anh -->
          <img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/..." alt="Nguyễn Văn Tuyến">
        </div>
        <div class="hp-hero-text">
          <div class="hp-kicker">Executive Profile · HRM & Digital Transformation</div>
          <h1>
            Nguyễn Văn Tuyến
            <span class="sub">Trưởng phòng Hành chính Nhân sự</span>
          </h1>
          <p class="hp-tagline">
            Tôi không chỉ quản trị nhân sự bằng kinh nghiệm. Tôi <b>kiến tạo hệ thống nhân sự tự vận hành bằng công nghệ</b>. 
            Từ một người làm quản trị, tôi đã trực tiếp ứng dụng AI để xây dựng phần mềm HRM Master, giải quyết triệt để bài toán 
            vận hành cho doanh nghiệp. Tôi tìm kiếm cơ hội dẫn dắt chuyển đổi số tại vị trí Trưởng phòng HCNS.
          </p>
          <div class="hp-cta-row">
            {_demo_btn}
            {_cv_btn}
          </div>
        </div>
      </div>
    </div>
  </section>

  <!-- METRICS -->
  <div class="hp-metrics-sec">
    <div class="hp-wrap">
      <div class="hp-metrics-grid">
        <div class="hp-metric-item">
          <div class="hp-metric-num">20<span>+</span></div>
          <div class="hp-metric-label">Năm kinh nghiệm<br>quản trị</div>
        </div>
        <div class="hp-metric-item">
          <div class="hp-metric-num">15.000<span>+</span></div>
          <div class="hp-metric-label">Dòng code tự tay<br>chỉ đạo AI viết</div>
        </div>
        <div class="hp-metric-item">
          <div class="hp-metric-num">80<span>%</span></div>
          <div class="hp-metric-label">Thời gian xử lý<br>nghiệp vụ được cắt giảm</div>
        </div>
        <div class="hp-metric-item">
          <div class="hp-metric-num">300<span>+</span></div>
          <div class="hp-metric-label">Nhân sự đã quản lý<br>trực tiếp (Bảo Việt)</div>
        </div>
      </div>
    </div>
  </div>

  <!-- CORE COMPETENCIES -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <h2>Năng lực cốt lõi</h2>
      <div class="hp-sec-subtitle">Sự kết hợp giữa Quản trị thực chiến và Tư duy Công nghệ</div>
      <div class="hp-comp-grid">
        <div class="hp-comp-card">
          <div class="hp-comp-icon">🧠</div>
          <h4>Tư duy Hệ thống (System Thinking)</h4>
          <p>Nhìn bài toán nhân sự như một hệ thống các thực thể liên kết: Nhân viên → Hợp đồng → BHXH → Chấm công → Lương → Thuế. Thiết kế kiến trúc dữ liệu chuẩn hóa, đảm bảo tính toàn vẹn và khả năng mở rộng.</p>
        </div>
        <div class="hp-comp-card">
          <div class="hp-comp-icon">⚙️</div>
          <h4>Chuyển bài toán Quản trị thành Sản phẩm</h4>
          <p>Không dừng ở phân tích nghiệp vụ. Tôi trực tiếp chỉ đạo AI Agent (Claude, ChatGPT) xây dựng sản phẩm hoàn chỉnh: từ Database Schema, Business Logic, đến UI/UX và triển khai Cloud.</p>
        </div>
        <div class="hp-comp-card">
          <div class="hp-comp-icon">📊</div>
          <h4>Quản trị Nhân sự & C&B Toàn diện</h4>
          <p>20 năm kinh nghiệm thực chiến (13 năm Founder/CEO, 10 năm tại Bảo Việt). Hiểu sâu bài toán vận hành doanh nghiệp, quản lý quỹ lương, tối ưu chi phí và xây dựng chính sách đãi ngộ.</p>
        </div>
        <div class="hp-comp-card">
          <div class="hp-comp-icon">🌱</div>
          <h4>Tư duy Chuyển đổi số & Đổi mới</h4>
          <p>Bắt đầu học Vibe Coding từ con số 0 ở tuổi 49. Trong 20 ngày, hoàn thành bản thiết kế quy trình đầu tiên và đưa vào áp dụng thực tế. Tinh thần học hỏi không giới hạn bởi tuổi tác.</p>
        </div>
      </div>
    </div>
  </section>

  <!-- SYSTEM ARCHITECTURE & AI -->
  <section class="hp-sec" style="background-color: #F1F5F9;">
    <div class="hp-wrap">
      <h2>Dự án Tiêu biểu: HRM Master</h2>
      <div class="hp-sec-subtitle">Hệ thống Quản trị Nhân sự tự xây dựng bằng sức mạnh của AI</div>
      
      <div class="hp-sys-box">
        <h3>🤖 Tôi là Nhà thiết kế, AI là Kỹ sư ảo</h3>
        <p>HRM Master không phải là một script đơn lẻ. Đây là một hệ thống hoàn chỉnh với kiến trúc phân tầng rõ ràng, được thiết kế để phục vụ nhiều doanh nghiệp và có thể mở rộng theo nhu cầu.</p>
        <ul>
          <li><b>Database Layer:</b> PostgreSQL (Supabase) — Chuẩn hóa dữ liệu nhân sự, hợp đồng, BHXH, chấm công, lương.</li>
          <li><b>Application Layer:</b> Python + Streamlit — Xử lý nghiệp vụ, tính toán lương, xuất báo cáo D02-LT.</li>
          <li><b>AI Layer:</b> Anthropic Claude API — Chatbot tư vấn pháp luật lao động, trợ lý ảo cho HR.</li>
          <li><b>Multi-tenant:</b> Mỗi doanh nghiệp có database riêng, landing page riêng, cấu hình riêng.</li>
        </ul>
        
        <!-- SƠ ĐỒ DATA FLOW -->
        <div style="text-align: center; margin-top: 20px;">
            <!-- Anh thay src bằng biến DATAFLOW_B64 -->
            <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgA..." alt="Sơ đồ kiến trúc HRM Master" style="max-width: 100%; border-radius: 8px;">
        </div>
      </div>
    </div>
  </section>

  <!-- CAREER JOURNEY -->
  <section class="hp-sec">
    <div class="hp-wrap">
      <h2>Hành trình Nghề nghiệp</h2>
      <div class="hp-sec-subtitle">Từ Quản trị Truyền thống đến Chuyển đổi số</div>
      <div class="hp-timeline">
        <div class="hp-tl-item">
          <div class="hp-tl-date">05/2026 - Nay</div>
          <h3>CV HCNS kiêm Product Owner HRM</h3>
          <div class="hp-role">Doanh nghiệp hiện tại</div>
          <p>Chủ động khảo sát quy trình, xác định điểm nghẽn. Tự phân tích nghiệp vụ, thiết kế và phát triển hệ thống HRM Master. Kết quả: App giúp tiết kiệm ~80% thời gian xử lý nghiệp vụ HCNS.</p>
        </div>
        <div class="hp-tl-item">
          <div class="hp-tl-date">2023 - 04/2026</div>
          <h3>Freelancer Tư vấn Du học & XKLĐ</h3>
          <div class="hp-role">CTCP Cung ứng Nhân lực Nghệ An</div>
          <p>Trực tiếp tư vấn, xây dựng lộ trình cho hàng trăm khách hàng. Đàm phán, ký kết hợp đồng với đối tác nước ngoài.</p>
        </div>
        <div class="hp-tl-item">
          <div class="hp-tl-date">2010 - 2023</div>
          <h3>Founder & CEO</h3>
          <div class="hp-role">CTCP Đầu tư & Thương mại An Phú Lộc</div>
          <p>Thành lập và điều hành doanh nghiệp, xây dựng đội ngũ. Trực tiếp đàm phán nhập khẩu 2 container ô tô từ Hàn Quốc. Tự xây dựng hệ thống chính sách nhân sự, đánh giá KPI.</p>
        </div>
        <div class="hp-tl-item">
          <div class="hp-tl-date">1999 - 2010</div>
          <h3>Trưởng phòng Nghiệp vụ</h3>
          <div class="hp-role">Bảo Việt Nhân thọ Nghệ An</div>
          <p>Quản lý trực tiếp đội ngũ 300 tư vấn viên. Được bổ nhiệm Trưởng phòng sau 3 năm công tác (năm 24 tuổi). Đặt nền móng tư duy quản trị hệ thống, chuẩn hóa quy trình.</p>
        </div>
      </div>
    </div>
  </section>

  <!-- FOOTER -->
  <footer>
    <div class="hp-wrap">
      <p style="margin-bottom: 10px;">
        <strong>Nguyễn Văn Tuyến</strong> · {PHONE} · <a href="mailto:{CONTACT_EMAIL}">{CONTACT_EMAIL}</a>
      </p>
      <p style="margin: 0; font-size: 12px; color: #64748B;">
        © 2026 Nguyễn Văn Tuyến. Portfolio được thiết kế cho vị trí Trưởng phòng HCNS.
      </p>
    </div>
  </footer>

</div>
"""

    _landing_html = _css + _html_body

    # Hiển thị
    try:
        _ua = (st.context.headers.get("User-Agent") or "").lower()
    except Exception:
        _ua = ""
    _la_mobile = any(k in _ua for k in ("mobi", "iphone", "android"))
    
    # Tăng chiều cao để chứa nội dung mới
    components.html(_landing_html, height=5500 if _la_mobile else 4200, scrolling=True)


if __name__ == "__main__":
    render()