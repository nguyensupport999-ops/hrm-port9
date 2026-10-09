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
    """Vẽ Landing Page giới thiệu, hiển thị ở vùng nội dung chính (bên phải sidebar
    đăng nhập) khi CHƯA đăng nhập. Toàn bộ ảnh (dashboard demo + ảnh người sáng lập)
    được nhúng sẵn dạng base64 để không phụ thuộc file ngoài."""
    """Landing page giới thiệu HRM Master + câu chuyện hành trình, hiển thị trong
    vùng nội dung chính (bên phải sidebar) khi CHƯA đăng nhập. Toàn bộ ảnh (dashboard
    demo + ảnh người sáng lập) được nhúng sẵn dạng base64 để không phụ thuộc file ngoài."""
    # LƯU Ý: cố tình KHÔNG dùng st.markdown(..., unsafe_allow_html=True) ở đây.
    # Khối HTML/CSS này rất dài và có nhiều dòng thụt lề (indent) sau dòng trống —
    # trình phân tích Markdown của Streamlit hiểu nhầm các đoạn đó là "code block"
    # (chuẩn Markdown: 4+ dấu cách sau dòng trống = code) nên hiển thị ra y nguyên
    # dạng text/code thay vì render thành trang landing page.
    # st.components.v1.html() không đi qua bước parse Markdown nên tránh được lỗi
    # này hoàn toàn, dù nội dung có bao nhiêu dòng thụt lề đi nữa.
    _landing_html = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Be+Vietnam+Pro:wght@300;400;500;600;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
  #hrm-landing{
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
  }
  #hrm-landing *{ box-sizing:border-box; }
  #hrm-landing img{ max-width:100%; display:block; }
  #hrm-landing .hl-wrap{ max-width:900px; margin:0 auto; padding:0 26px; }
  #hrm-landing h1,#hrm-landing h2,#hrm-landing h3{
    font-family:'Fraunces', serif;
    font-weight:500;
    margin:0;
    color:var(--pine-deep);
  }

  /* ---------- HERO ---------- */
  #hrm-landing .hl-hero{ padding:40px 0 32px; border-bottom:1px solid var(--line); }
  #hrm-landing .hl-eyebrow{
    display:inline-flex; align-items:center; gap:7px;
    font-family:'Space Mono', monospace;
    font-size:10.5px; letter-spacing:.12em; text-transform:uppercase;
    color:var(--rust); margin-bottom:12px; font-weight:400;
  }
  #hrm-landing .hl-eyebrow::before{ content:""; width:16px; height:1px; background:var(--rust); display:inline-block; }
  #hrm-landing .hl-hero h1{
    font-size:clamp(22px, 3vw, 30px);
    line-height:1.22;
    max-width:60ch;
    font-weight:500;
  }
  #hrm-landing .hl-hero h1 em{ font-style:italic; color:var(--amber); }
  #hrm-landing .hl-hero p.hl-lead{
    margin-top:12px; max-width:100ch; font-size:13.5px; font-weight:300;
    color:var(--ink-soft);
  }
  #hrm-landing .hl-cta-row{ margin-top:18px; display:flex; gap:12px; flex-wrap:wrap; }
  #hrm-landing .hl-btn{
    font-family:'Be Vietnam Pro', sans-serif; font-weight:600; font-size:12.5px;
    padding:10px 20px; border-radius:var(--radius); border:1px solid transparent;
    display:inline-flex; align-items:center; gap:7px;
    background:var(--pine); color:var(--paper); box-shadow:2.5px 2.5px 0 var(--amber);
    text-decoration:none;
  }

  #hrm-landing .hl-mockup{ margin-top:30px; position:relative; }
  #hrm-landing .hl-sticky{
    position:absolute; top:-16px; right:5%;
    background:#FBEFC4; color:#4a3a10;
    font-family:'Space Mono', monospace; font-size:10.5px;
    padding:9px 11px; max-width:190px;
    box-shadow:0 5px 12px rgba(0,0,0,.12);
    transform:rotate(3deg); z-index:3; border-radius:2px; font-weight:400;
  }
  #hrm-landing .hl-frame{
    background:#fff; border-radius:8px; overflow:hidden;
    box-shadow:0 16px 36px rgba(18,48,42,.16); border:1px solid var(--line);
  }
  #hrm-landing .hl-frame-bar{
    background:var(--paper-deep); padding:8px 12px; display:flex; align-items:center; gap:6px;
    border-bottom:1px solid var(--line);
  }
  #hrm-landing .hl-dot{ width:7px; height:7px; border-radius:50%; background:#d8cfbf; }
  #hrm-landing .hl-frame-url{
    margin-left:8px; font-family:'Space Mono', monospace; font-size:10px; color:var(--ink-soft);
    background:#fff; border:1px solid var(--line); padding:3px 9px; border-radius:20px; font-weight:400;
  }
  #hrm-landing .hl-frame img{ width:100%; display:block; }

  /* ---------- STORY ---------- */
  #hrm-landing .hl-story{ padding:38px 0; border-bottom:1px solid var(--line); }
  #hrm-landing .hl-story-grid{ display:flex; gap:26px; align-items:flex-start; max-width:680px; }
  #hrm-landing .hl-avatar{
    flex:0 0 auto; width:64px; height:64px; border-radius:50%; overflow:hidden;
    border:2px solid var(--paper-deep); box-shadow:0 3px 8px rgba(0,0,0,.12);
  }
  #hrm-landing .hl-avatar img{ width:100%; height:100%; object-fit:cover; }
  #hrm-landing .hl-kicker{
    font-family:'Space Mono', monospace; font-size:10.5px; color:var(--pine);
    letter-spacing:.08em; text-transform:uppercase; margin-bottom:8px; font-weight:400;
  }
  #hrm-landing .hl-story h2{ font-size:19px; margin-bottom:12px; font-weight:500; }
  #hrm-landing .hl-story blockquote{
    margin:0; padding-left:16px; border-left:2.5px solid var(--amber);
    font-size:13.5px; line-height:1.7; color:var(--ink); font-weight:300;
  }
  #hrm-landing .hl-story blockquote p{ margin:0 0 10px; }
  #hrm-landing .hl-sign{
    margin-top:12px; font-family:'Space Mono', monospace; font-size:11px;
    color:var(--ink-soft); font-weight:400;
  }
  #hrm-landing .hl-sign b{ color:var(--pine-deep); font-family:'Be Vietnam Pro',sans-serif; }

  /* ---------- TIMELINE ---------- */
  #hrm-landing .hl-timeline-sec{ padding:38px 0; border-bottom:1px solid var(--line); background:var(--paper-deep); }
  #hrm-landing .hl-timeline-sec h2{ font-size:19px; margin-bottom:26px; max-width:24ch; font-weight:500; }
  #hrm-landing .hl-tl{ position:relative; padding-left:26px; }
  #hrm-landing .hl-tl::before{
    content:""; position:absolute; left:5px; top:5px; bottom:5px; width:1.2px;
    background:repeating-linear-gradient(to bottom, var(--pine) 0 5px, transparent 5px 9px);
  }
  #hrm-landing .hl-tl-item{ position:relative; padding-bottom:22px; }
  #hrm-landing .hl-tl-item:last-child{ padding-bottom:0; }
  #hrm-landing .hl-tl-item::before{
    content:""; position:absolute; left:-26px; top:2px; width:10px; height:10px;
    border-radius:50%; background:var(--paper-deep); border:2px solid var(--rust);
  }
  #hrm-landing .hl-tl-date{
    font-family:'Space Mono', monospace; font-size:10.5px; color:var(--rust);
    background:#fff; display:inline-block; padding:2px 8px; border-radius:2px;
    border:1px solid var(--line); margin-bottom:5px; font-weight:400;
  }
  #hrm-landing .hl-tl-item h3{ font-size:14.5px; margin-bottom:4px; font-weight:500; }
  #hrm-landing .hl-tl-item p{ margin:0; color:var(--ink-soft); max-width:60ch; font-size:12.5px; font-weight:300; }

  /* ---------- FEATURES ---------- */
  #hrm-landing .hl-features{ padding:38px 0; border-bottom:1px solid var(--line); }
  #hrm-landing .hl-features h2{ font-size:19px; margin-bottom:8px; max-width:24ch; font-weight:500; }
  #hrm-landing .hl-features > .hl-wrap > p.hl-sub{ font-size:12.5px; color:var(--ink-soft); margin:0 0 20px; font-weight:300; }

  #hrm-landing .hl-highlight{
    display:flex; align-items:center; gap:14px;
    background:var(--pine-deep); color:#EFE8D8;
    border-radius:6px; padding:16px 18px; margin-bottom:16px;
  }
  #hrm-landing .hl-highlight .hl-badge{
    font-family:'Space Mono', monospace; font-size:9.5px; letter-spacing:.08em;
    text-transform:uppercase; background:var(--amber); color:#2a1c05;
    padding:3px 8px; border-radius:20px; flex:0 0 auto; font-weight:700;
  }
  #hrm-landing .hl-highlight .hl-ic{ font-size:20px; flex:0 0 auto; }
  #hrm-landing .hl-highlight-text h4{
    font-family:'Fraunces', serif; font-size:14.5px; margin:0 0 3px; color:#fff; font-weight:500;
  }
  #hrm-landing .hl-highlight-text p{ margin:0; font-size:12px; color:#c9d8d3; font-weight:300; }

  #hrm-landing .hl-feat-grid{ display:grid; grid-template-columns:repeat(3, 1fr); gap:1px; background:var(--line); border:1px solid var(--line); }
  #hrm-landing .hl-feat{ background:var(--paper); padding:16px 15px; }
  #hrm-landing .hl-feat .hl-fic{ font-size:16px; margin-bottom:8px; display:block; }
  #hrm-landing .hl-feat h4{ font-family:'Fraunces', serif; font-size:13px; margin:0 0 4px; color:var(--pine-deep); font-weight:500; }
  #hrm-landing .hl-feat p{ margin:0; font-size:11px; color:var(--ink-soft); font-weight:300; }

  /* ---------- INSPIRE ---------- */
  #hrm-landing .hl-inspire{ padding:42px 0; background:var(--pine-deep); color:#EFE8D8; text-align:center; }
  #hrm-landing .hl-inspire .hl-wrap{ max-width:620px; }
  #hrm-landing .hl-inspire h2{ color:#fff; font-size:clamp(18px,2.4vw,24px); line-height:1.35; font-weight:500; }
  #hrm-landing .hl-inspire h2 em{ color:var(--amber); font-style:italic; }
  #hrm-landing .hl-inspire p{ margin-top:10px; color:#c9d8d3; font-size:12.5px; font-weight:300; }
  #hrm-landing .hl-inspire .hl-cta-row{ justify-content:center; margin-top:18px; }
  #hrm-landing .hl-inspire .hl-btn{ box-shadow:2.5px 2.5px 0 var(--rust); }

  #hrm-landing footer{
    padding:16px 0 22px; text-align:center; font-family:'Space Mono', monospace;
    font-size:10px; color:var(--ink-soft); font-weight:400;
  }

  @media (max-width: 720px){
    #hrm-landing .hl-feat-grid{ grid-template-columns:1fr; }
    #hrm-landing .hl-sticky{ display:none; }
    #hrm-landing .hl-story-grid{ flex-direction:column; }
  }
</style>

<div id="hrm-landing">

  <section class="hl-hero">
    <div class="hl-wrap">
      <div class="hl-eyebrow">Về câu chuyện của Tôi đến với Viber Codding · bắt đầu 02/05/2026</div>
      <h1>Từ những file Excel rối ren,<br>đến phần mềm nhân sự <em>của chính mình</em></h1>
      <p class="hl-lead">HRM Master ra đời từ tay một chuyên viên HR chưa từng viết một dòng code — chỉ vì không muốn tiếp tục quản lý hồ sơ nhân sự bằng thói quen cũ. Đây là hành trình, và cũng là lời mời bạn dùng thử.</p>
      <div class="hl-cta-row">
        <span class="hl-btn">← Dùng thử miễn phí</span>
      </div>

      <div class="hl-mockup">
        <div class="hl-sticky">"Ngày đó tôi còn chưa biết Python là gì." — trích nhật ký, tháng 5/2026</div>
        <div class="hl-frame">
          <div class="hl-frame-bar">
            <span class="hl-dot"></span><span class="hl-dot"></span><span class="hl-dot"></span>
            <span class="hl-frame-url">demo-hrm.streamlit.app</span>
          </div>
#
        </div>
      </div>
    </div>
  </section>

  <section class="hl-story">
    <div class="hl-wrap">
      <div class="hl-kicker">Vì sao HRM Master ra đời</div>
      <div class="hl-story-grid">
        <div class="hl-avatar">
 ##
        </div>
        <div>
          <h2>Một thói quen cũ, và câu hỏi<br>"Tại sao cứ phải làm bằng Excel?"</h2>
          <blockquote>
            <p>Tôi nhận việc chuyên viên HR ngày 02/05/2026. Việc đầu tiên tôi học được không phải là quy trình, mà là một thói quen: quản lý hồ sơ nhân sự bằng Excel — thói quen được truyền lại từ các đồng nghiệp, và từ chính người làm HR trước tôi.</p>
            <p>Tôi không trách thói quen đó. Nhưng tôi thấy công việc quản trị nhân sự của công ty mình xứng đáng được chuẩn hóa tốt hơn. Vậy là, ở tuổi mà nhiều người nghĩ "học code không còn dành cho mình", tôi bắt đầu học Vibe Coding — không phải để trở thành lập trình viên, mà để tự tay giải quyết một vấn đề tôi hiểu rõ hơn ai hết.</p>
          </blockquote>
          <div class="hl-sign"><b>Mr.Tuyến</b> - người sáng lập HRM Master</div>
        </div>
      </div>
    </div>
  </section>

  <section class="hl-timeline-sec" id="hanh-trinh">
    <div class="hl-wrap">
      <div class="hl-kicker">Hành trình</div>
      <h2>20 ngày đầu tiên, và những gì đến sau đó</h2>
      <div class="hl-tl">
        <div class="hl-tl-item">
          <div class="hl-tl-date">02 / 05 / 2026</div>
          <h3>Ngày đầu làm chuyên viên HR</h3>
          <p>Tiếp nhận công việc quản trị nhân sự, kế thừa thói quen quản lý hồ sơ bằng Excel từ đồng nghiệp và HR tiền nhiệm. Nhận ra quy trình cần được chuẩn hóa lại từ gốc.</p>
        </div>
        <div class="hl-tl-item">
          <div class="hl-tl-date">Tháng 05 / 2026</div>
          <h3>Bắt đầu học Vibe Coding</h3>
          <p>Chưa từng viết một dòng code trước đó. Học từng khái niệm cơ bản, từng bước, song song với công việc HR hằng ngày — để tự xây công cụ giải quyết đúng vấn đề mình đang gặp.</p>
        </div>
        <div class="hl-tl-item">
          <div class="hl-tl-date">Sau 20 ngày</div>
          <h3>Hoàn thành bản thiết kế quy trình đầu tiên</h3>
          <p>Đưa ngay vào áp dụng thực tế với hồ sơ nhân sự của công ty — không phải bản demo, mà là dữ liệu thật, quy trình thật.</p>
        </div>
        <div class="hl-tl-item">
          <div class="hl-tl-date">Đến nay · 07 / 2026</div>
          <h3>HRM Master - sẵn sàng cho nhiều doanh nghiệp</h3>
          <p>Từ một công cụ nội bộ, phát triển thành nền tảng quản trị nhân sự đa doanh nghiệp — và vẫn đang tiếp tục hoàn thiện mỗi ngày.</p>
        </div>
      </div>
    </div>
  </section>

  <section class="hl-features">
    <div class="hl-wrap">
      <div class="hl-kicker">HRM Master làm được gì</div>
      <h2>Đủ dùng cho việc quản trị nhân sự, 'đo ni đóng giày' cho từng yêu cầu quản lý của doanh nghiệp!</h2>
      <p class="hl-sub">Từ những nghiệp vụ nền tảng, đến điểm khác biệt không nhiều phần mềm nhân sự trên thị trường hiện có.</p>

      <div class="hl-highlight">
        <span class="hl-badge">Đột phá</span>
        <span class="hl-ic">🤖</span>
        <div class="hl-highlight-text">
          <h4>AI Tư vấn Hành chính Nhân sự</h4>
          <p>Trợ lý AI được huấn luyện và sở hữu Big Data về luật liên quan,sẽ phân tích ngữ cảnh và thông tin đầu vào để truy vấn các Điều luật liên quan và đưa ra ý kiến tư vấn tức thì cho bạn - Một công cụ hữu ích cho cả nhà quản lý và Hr.</p>
        </div>
      </div>

      <div class="hl-feat-grid">
        <div class="hl-feat">
          <span class="hl-fic">🗂️</span>
          <h4>Hồ sơ nhân viên</h4>
          <p>Quản lý tập trung thông tin, hợp đồng lao động, quyết định nhân sự - số hóa để thay thế hoàn toàn các file Excel rời rạc.</p>
        </div>
        <div class="hl-feat">
          <span class="hl-fic">🕒</span>
          <h4>Chấm công</h4>
          <p>Không cần trang bị thêm máy chấm công - HRM Master ứng dụng Face ID vào nghiệp vụ chấm công. Theo dõi công, phép, tăng ca theo thời gian thực, đồng bộ trực tiếp vào bảng lương.</p>
        </div>
        <div class="hl-feat">
          <span class="hl-fic">📋</span>
          <h4>BHXH</h4>
          <p>Quản lý mã số, mức đóng, theo dõi Tăng/Giảm và xuất file báo cáo chuẩn form của BHXH để import trực tiếp vào phần mềm BHXH 'tokhaibaohiem.vn'</p>
        </div>
        <div class="hl-feat">
          <span class="hl-fic">💰</span>
          <h4>Lương 3P</h4>
          <p>Tư vấn Xây dựng chính sách lương theo Vị trí - Năng lực - Kết quả, có gợi ý mẫu theo cơ cấu tổ chức riêng. Hoặc tùy chỉnh theo đúng chính sách tiền lương mà DN đang áp dụng</p>
        </div>
        <div class="hl-feat">
          <span class="hl-fic">📄</span>
          <h4>Công văn &amp; HĐ kinh tế</h4>
          <p>Soạn thảo, lưu trữ, tra cứu công văn đến/đi và hợp đồng kinh tế theo đúng quy trình, nghiệp vụ Văn thư lưu trữ.</p>
        </div>
        <div class="hl-feat">
          <span class="hl-fic">💬</span>
          <h4>Chat nội bộ</h4>
          <p>Trao đổi công việc trực tiếp trong hệ thống, gắn liền với đúng hồ sơ nhân sự liên quan. Dễ dàng giao việc và quản lý tiến độ công việc đã giao</p>
        </div>
      </div>
    </div>
  </section>

  <section class="hl-inspire">
    <div class="hl-wrap">
      <h2>Không ai bắt đầu học điều mới ở <em>"đúng độ tuổi"</em> cả.</h2>
      <p>Nếu một chuyên viên HR chưa từng viết code có thể tự xây phần mềm quản trị nhân sự cho công ty mình, thì hành trình học hỏi của bạn — dù bắt đầu ở đâu, tuổi nào — cũng đáng để bắt đầu ngay hôm nay.</p>
      <div class="hl-cta-row">
        <span class="hl-btn">← Dùng thử HRM Master ngay</span>
      </div>
    </div>
  </section>

  <footer>
    © 2026 HRM Master · một sản phẩm được xây từ một câu hỏi nhân sự rất đời thường
  </footer>

</div>

<script>
  // Iframe của components.html có chiều cao cố định theo tham số height=... truyền
  // từ Python, không tự co giãn theo nội dung — nên phải tự đo và set lại chiều cao
  // bằng JS. window.frameElement trỏ đúng tới thẻ <iframe> đang chứa nội dung này
  // (vẫn cùng-origin với trang Streamlit nên truy cập được, không cần postMessage).
  function _hrmResizeIframe() {
    try {
      var h = document.documentElement.scrollHeight;
      if (window.frameElement && h > 0) {
        window.frameElement.style.height = (h + 20) + "px";
      }
    } catch (e) {}
  }
  window.addEventListener("load", _hrmResizeIframe);
  document.fonts && document.fonts.ready && document.fonts.ready.then(_hrmResizeIframe);
  new ResizeObserver(_hrmResizeIframe).observe(document.body);
  setTimeout(_hrmResizeIframe, 200);
  setTimeout(_hrmResizeIframe, 800);
  setTimeout(_hrmResizeIframe, 2000);
</script>
"""
    components.html(_landing_html, height=50, scrolling=False)
