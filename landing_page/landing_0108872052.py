# -*- coding: utf-8 -*-
"""
landing_page/landing_demo.py
=============================
Landing Page MẶC ĐỊNH (dùng cho MỌI tenant chưa có landing page riêng).

QUY ƯỚC:
- Nội dung của file này là bản trích xuất NGUYÊN VẸN (không viết lại) từ hàm
  render_landing_page() / show_landing_page() gốc trong app.py, TRƯỚC khi có cơ
  chế đa-tenant. Được app.py tự động nạp qua _load_tenant_module_or_demo('landing_page',
  'landing', ma_so_thue) khi tenant đang đăng nhập CHƯA có file riêng
  'landing_page/landing_{ma_so_thue}.py'.
- Muốn có Landing Page RIÊNG cho 1 khách hàng: copy file này thành
  'landing_page/landing_{ma_so_thue}.py' (đặt cùng cấp với app.py, dùng đúng Mã số
  thuế của khách hàng đó) rồi tuỳ biến nội dung bên trong hàm render().
- HỢP ĐỒNG BẮT BUỘC: file phải có hàm render() không nhận tham số — app.py chỉ gọi
  module.render(), không gọi gì khác. Bên trong có thể tự do dùng st, components,
  COMPANY_CONFIG, st.session_state... như bản gốc.
"""

import streamlit as st
import streamlit.components.v1 as components
import os

# Import config - ưu tiên config.py (local), fallback to config_template (cloud).
# Dùng from-import để nhận CÙNG 1 dict COMPANY_CONFIG mà app.py đang dùng (dict là
# mutable object dùng chung giữa các module đã import nó — app.py cập nhật branding
# theo tenant vào dict này, landing_demo.py đọc lại sẽ thấy đúng giá trị mới nhất).
try:
    from config import COMPANY_CONFIG
except ImportError:
    from config_template import COMPANY_CONFIG


def render():
    """Hiển thị Landing Page với chuyển ngữ Việt/Anh"""
    
    # Import languages
    try:
        from languages import LANGUAGES
    except ImportError:
        # Fallback nếu chưa có file languages.py
        LANGUAGES = {'vi': {}, 'en': {}}
    
    lang = st.session_state.get('language', 'vi')
    text = LANGUAGES.get(lang, LANGUAGES.get('vi', {}))
    
    # Ẩn UI chrome của Streamlit
    st.markdown("""
        <style>
            [data-testid="stDecoration"],
            [data-testid="stHeader"],
            header[data-testid],
            footer[data-testid],
            .stAppDeployButton,
            .stToolbar,
            .stStatusWidget,
            .stApp > header,
            .stApp > div[data-testid="stToolbar"] {
                display: none !important;
                height: 0 !important;
            }
            html, body, .stApp, .stApp > div {
                margin: 0 !important;
                padding: 0 !important;
            }
            .main > div {
                padding: 0 !important;
                margin: 0 !important;
            }
            .block-container {
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                padding-left: 0 !important;
                padding-right: 0 !important;
                max-width: 100% !important;
            }
            section[data-testid="stMain"] > div {
                padding-top: 0 !important;
            }
            iframe {
                border: none !important;
                display: block !important;
                margin: 0 !important;
                padding: 0 !important;
                width: 100% !important;
            }
            body {
                overflow-x: hidden;
            }
            [data-testid="stDataFrame"] {
                max-height: 700px !important;
            }
            [data-testid="stDataFrame"] > div {
                max-height: 700px !important;
            }
            [data-testid="stDataEditor"] {
                max-height: 700px !important;
            }
            [data-testid="stDataEditor"] > div {
                max-height: 700px !important;
            }
            [data-testid="stDataFrame"] td:last-child,
            [data-testid="stDataEditor"] td:last-child {
                background-color: #E8F5E9 !important;
                font-weight: bold !important;
                cursor: pointer !important;
            }
            
            /* ===== Tăng chiều cao bảng chấm công ===== */
            [data-testid="stDataFrame"] {
                max-height: 800px !important;
            }
            [data-testid="stDataFrame"] > div {
                max-height: 800px !important;
            }
            [data-testid="stDataEditor"] {
                max-height: 800px !important;
            }
            [data-testid="stDataEditor"] > div {
                max-height: 800px !important;
            }
        </style>
    """, unsafe_allow_html=True)
    
    # Đọc file logo động
    import base64
    import requests
    logo_base64 = ""
    logo_src = COMPANY_CONFIG.get("logo_url")
    if logo_src:
        if logo_src.startswith("http://") or logo_src.startswith("https://"):
            try:
                response = requests.get(logo_src, timeout=3)
                if response.status_code == 200:
                    logo_base64 = base64.b64encode(response.content).decode()
            except Exception:
                pass
        elif os.path.exists(logo_src):
            try:
                with open(logo_src, "rb") as f:
                    logo_base64 = base64.b64encode(f.read()).decode()
            except Exception:
                pass
    
    if not logo_base64:
        # Fallback về logo_cty.png mặc định
        logo_path = os.path.join(os.path.dirname(__file__), "logo_cty.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                logo_base64 = base64.b64encode(f.read()).decode()
    
    # Đọc ảnh slider
    def load_img_b64(filename):
        path = os.path.join(os.path.dirname(__file__), "static", filename)
        if os.path.exists(path):
            with open(path, "rb") as f:
                ext = filename.rsplit(".", 1)[-1].lower()
                mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
                return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
        return ""
    
    slide1_src = load_img_b64("anh1.jpeg")
    slide2_src = load_img_b64("anh2.jpeg")
    slide3_src = load_img_b64("anh3.jpeg")
    chu_tich_img = load_img_b64("chu_tich.png")
    chu_ky_img = load_img_b64("123456.png")
    
    # Active class cho language buttons
    vi_active = 'active' if lang == 'vi' else ''
    en_active = 'active' if lang == 'en' else ''
    
    # JavaScript riêng biệt (không có {} để tránh xung đột)
    landing_js = """
    <script>
        // Hàm cuộn mượt đến section bằng window.top
        function scrollToSection(sectionId) {
            var topWin = window.top || window.parent || window;
            var targetElement = document.getElementById(sectionId);
            if (targetElement) {
                var targetRect = targetElement.getBoundingClientRect();
                var offsetTop = targetRect.top + (topWin.scrollY || topWin.pageYOffset);
                topWin.scrollTo({
                    top: offsetTop - 80,
                    behavior: 'smooth'
                });
            }
        }
        
        function handleNavClick(e) {
            e.preventDefault();
            var section = this.getAttribute('data-section');
            if (section) {
                scrollToSection(section);
            }
        }
        
        function initNavigation() {
            var navLinks = document.querySelectorAll('.nav-link');
            for (var i = 0; i < navLinks.length; i++) {
                navLinks[i].removeEventListener('click', handleNavClick);
                navLinks[i].addEventListener('click', handleNavClick);
            }
        }
        
        // Slider
        var currentSlide = 0;
        var slides = document.querySelectorAll('.slide');
        var dots = document.querySelectorAll('.slider-dot');
        var totalSlides = slides.length;
        var autoSlideInterval;
        var progressInterval;
        var progressValue = 0;
        var SLIDE_DURATION = 5000;
        var progressBar = document.getElementById('sliderProgress');
        
        function showSlide(index) {
            for (var i = 0; i < slides.length; i++) {
                slides[i].classList.remove('active');
                if (dots[i]) dots[i].classList.remove('active');
            }
            slides[index].classList.add('active');
            if (dots[index]) dots[index].classList.add('active');
            currentSlide = index;
            resetProgress();
        }
        
        function nextSlide() {
            showSlide((currentSlide + 1) % totalSlides);
        }
        
        function prevSlide() {
            showSlide((currentSlide - 1 + totalSlides) % totalSlides);
        }
        
        function resetProgress() {
            progressValue = 0;
            if (progressBar) progressBar.style.width = '0%';
        }
        
        function startProgress() {
            if (progressInterval) clearInterval(progressInterval);
            progressValue = 0;
            progressInterval = setInterval(function() {
                progressValue += 100 / (SLIDE_DURATION / 100);
                if (progressBar) progressBar.style.width = Math.min(progressValue, 100) + '%';
                if (progressValue >= 100) resetProgress();
            }, 100);
        }
        
        function startAutoSlide() {
            if (autoSlideInterval) clearInterval(autoSlideInterval);
            autoSlideInterval = setInterval(nextSlide, SLIDE_DURATION);
            startProgress();
        }
        
        if (totalSlides > 0) {
            for (var i = 0; i < dots.length; i++) {
                dots[i].addEventListener('click', (function(idx) {
                    return function() {
                        showSlide(idx);
                        if (autoSlideInterval) clearInterval(autoSlideInterval);
                        if (progressInterval) clearInterval(progressInterval);
                        startAutoSlide();
                    };
                })(i));
            }
            
            var prevBtn = document.getElementById('prevBtn');
            var nextBtn = document.getElementById('nextBtn');
            if (prevBtn) {
                prevBtn.addEventListener('click', function() {
                    prevSlide();
                    if (autoSlideInterval) clearInterval(autoSlideInterval);
                    if (progressInterval) clearInterval(progressInterval);
                    startAutoSlide();
                });
            }
            if (nextBtn) {
                nextBtn.addEventListener('click', function() {
                    nextSlide();
                    if (autoSlideInterval) clearInterval(autoSlideInterval);
                    if (progressInterval) clearInterval(progressInterval);
                    startAutoSlide();
                });
            }
            
            var touchStartX = 0;
            var heroSlider = document.querySelector('.hero-slider');
            if (heroSlider) {
                heroSlider.addEventListener('touchstart', function(e) {
                    touchStartX = e.touches[0].clientX;
                });
                heroSlider.addEventListener('touchend', function(e) {
                    var diff = touchStartX - e.changedTouches[0].clientX;
                    if (Math.abs(diff) > 50) {
                        diff > 0 ? nextSlide() : prevSlide();
                        startAutoSlide();
                    }
                });
            }
            startAutoSlide();
        }
        
        // Scroll reveal
        var revealObserver = new IntersectionObserver(function(entries) {
            entries.forEach(function(entry) {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                    revealObserver.unobserve(entry.target);
                }
            });
        }, { threshold: 0.15 });
        document.querySelectorAll('.reveal').forEach(function(el) {
            revealObserver.observe(el);
        });
        
        // Navbar scroll effect
        window.addEventListener('scroll', function() {
            var navbar = document.getElementById('navbar');
            if (navbar) {
                if (window.scrollY > 50) {
                    navbar.style.background = 'rgba(15, 59, 92, 0.98)';
                    navbar.style.boxShadow = '0 4px 20px rgba(0,0,0,0.2)';
                } else {
                    navbar.style.background = 'rgba(15, 59, 92, 0.95)';
                    navbar.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';
                }
            }
        });
        
        // Modal
        var modal = document.getElementById('thuNgoModal');
        var thuNgoBtn = document.getElementById('thuNgoBtn');
        var closeModalBtn = document.getElementById('closeModalBtn');
        
        if (thuNgoBtn && modal) {
            thuNgoBtn.addEventListener('click', function(e) {
                e.preventDefault();
                modal.classList.add('active');
                document.body.style.overflow = 'hidden';
            });
            
            var closeModal = function() {
                modal.classList.remove('active');
                document.body.style.overflow = '';
            };
            
            if (closeModalBtn) closeModalBtn.addEventListener('click', closeModal);
            modal.addEventListener('click', function(e) {
                if (e.target === modal) closeModal();
            });
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape' && modal.classList.contains('active')) closeModal();
            });
        }
        
        // Career link
        var careerLink = document.getElementById('careerLink');
        if (careerLink) {
            careerLink.addEventListener('click', function(e) {
                e.preventDefault();
                alert('Vui lòng liên hệ HR qua email: hr@honlaport.com.vn');
            });
        }
        
        // Language switcher
        function switchLanguage(lang) {
            var topWin = window.top || window.parent || window;
            var url = new URL(topWin.location.href);
            url.searchParams.set('lang', lang);
            topWin.history.replaceState(null, '', url.toString());
            topWin.location.reload();
        }
        
        // Khởi tạo navigation
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', function() {
                initNavigation();
            });
        } else {
            initNavigation();
        }
    </script>
    """
    
    landing_html = f"""
    <!DOCTYPE html>
    <html lang="{lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
        <title></title>   ### Sửa ở đây
        <link href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;14..32,400;14..32,500;14..32,600;14..32,700;14..32,800&display=swap" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            body {{
                font-family: 'Inter', sans-serif;
                background-color: #ffffff;
                color: #1e293b;
                line-height: 1.5;
                overflow-x: hidden;
                width: 100%;
                padding-top: 100px;
            }}
            ::-webkit-scrollbar {{
                width: 8px;
            }}
            ::-webkit-scrollbar-track {{
                background: #f1f1f1;
            }}
            ::-webkit-scrollbar-thumb {{
                background: #0f3b5c;
                border-radius: 4px;
            }}
            
            /* ===== NAVIGATION ===== */
            .navbar {{
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                z-index: 1000;
                padding: 0.8rem 30px;                                       
                background: rgba(15, 59, 92, 0.95);
                backdrop-filter: blur(10px);
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            .nav-container {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                max-width: 1400px;
                margin: 0 auto;
            }}
            
            .logo-circle {{
                width: 86px;
                height: 86px;
                border-radius: 50%;
                overflow: hidden;
                box-shadow: 0 8px 25px rgba(0,0,0,0.25), 0 0 0 3px rgba(255,255,255,0.3);
                background: white;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: transform 0.3s, box-shadow 0.3s;
                cursor: pointer;
            }}
            .logo-circle:hover {{
                transform: scale(1.02);
                box-shadow: 0 12px 30px rgba(0,0,0,0.3);
            }}
            .logo-circle img {{
                width: 100%;
                height: 100%;
                object-fit: cover;
            }}
            
            .nav-links {{
                display: flex;
                gap: 0.5rem;
                align-items: center;
                background: rgba(0,0,0,0.35);
                padding: 5px 15px;
                border-radius: 50px;
                backdrop-filter: blur(5px);
            }}
            .nav-links a {{
                text-decoration: none;
                color: white;
                font-weight: 500;
                font-size: 0.9rem;
                padding: 8px 16px;
                border-radius: 40px;
                transition: all 0.3s;
                cursor: pointer;
            }}
            .nav-links a:hover {{
                background: #f59e0b;
                color: #0f3b5c;
            }}
            
            /* ===== DIVIDER VÀ LANGUAGE SWITCH ===== */
            .nav-links .nav-divider {{
                color: rgba(255,255,255,0.4);
                margin: 0 5px;
                font-size: 14px;
            }}
            .lang-switch {{
                display: inline-flex;
                align-items: center;
                gap: 5px;
                margin-left: 5px;
            }}
            .lang-link {{
                text-decoration: none !important;
                color: white !important;
                font-weight: 500;
                font-size: 0.85rem;
                padding: 8px 8px !important;
                border-radius: 40px;
                transition: all 0.3s;
                background: transparent !important;
                cursor: pointer;
            }}
            .lang-link:hover {{
                background: #f59e0b !important;
                color: #0f3b5c !important;
            }}
            .lang-link.active {{
                background: #f59e0b !important;
                color: #0f3b5c !important;
            }}
            .lang-sep {{
                color: rgba(255,255,255,0.5);
                font-size: 12px;
            }}
            
            /* ===== MOBILE LANGUAGE ===== */
            .mobile-lang {{
                position: fixed;
                top: 15px;
                right: 15px;
                z-index: 10001;
                background: rgba(0,0,0,0.6);
                backdrop-filter: blur(8px);
                border-radius: 30px;
                padding: 6px 12px;
                display: none;
                gap: 8px;
                border: 1px solid rgba(255,255,255,0.2);
            }}
            .mobile-lang a {{
                color: white;
                text-decoration: none;
                font-size: 12px;
                font-weight: 600;
                padding: 4px 8px;
                border-radius: 20px;
                transition: all 0.2s;
                cursor: pointer;
            }}
            .mobile-lang a:hover {{
                background: rgba(255,255,255,0.2);
            }}
            .mobile-lang a.active {{
                background: #f59e0b;
                color: #0f3b5c;
            }}
            
            .dropdown {{
                position: relative;
            }}
            .dropdown-content {{
                display: none;
                position: absolute;
                background: white;
                min-width: 200px;
                box-shadow: 0 8px 16px rgba(0,0,0,0.1);
                border-radius: 8px;
                padding: 0.5rem 0;
                top: 100%;
                left: 0;
                z-index: 1;
            }}
            .dropdown:hover .dropdown-content {{
                display: block;
            }}
            .dropdown-content a {{
                color: #333 !important;
                padding: 8px 16px;
                display: block;
                font-size: 0.85rem;
                background: transparent;
                cursor: pointer;
            }}
            .dropdown-content a:hover {{
                background: #f8fafc;
                color: #f59e0b !important;
            }}
            
            /* ===== HERO SLIDER ===== */
            .hero-slider {{
                height: 550px;
                width: 100%;
                position: relative;
                overflow: hidden;
                margin-top: 0;
                background-color: #0a2a3a;
            }}
            .slides-container {{
                position: relative;
                width: 100%;
                height: 100%;
            }}
            .slide {{
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                opacity: 0;
                transition: opacity 0.8s ease-in-out;
                background-color: #0a2a3a;
            }}
            .slide.active {{
                opacity: 1;
                z-index: 2;
            }}
            .slide-layout {{
                display: flex;
                width: 100%;
                height: 100%;
                align-items: center;
                justify-content: center;
            }}
            .slide-image {{
                flex: 1;
                height: 100%;
                background-size: cover;
                background-position: center center;
                background-repeat: no-repeat;
            }}
            .slide-content {{
                flex: 1;
                padding: 50px 40px;
                color: white;
                z-index: 3;
                text-align: center;
            }}
            .slide-content h1 {{
                font-size: 2.8rem;
                font-weight: 800;
                margin-bottom: 1.2rem;
                letter-spacing: -0.5px;
                text-shadow: 0 2px 5px rgba(0,0,0,0.4);
                line-height: 1.3;
            }}
            .slide-content p {{
                font-size: 1.15rem;
                margin-bottom: 0.8rem;
                line-height: 1.5;
                text-shadow: 0 1px 3px rgba(0,0,0,0.3);
            }}
            .slide-content .highlight {{
                font-size: clamp(0.85rem, 1.1vw, 1.2rem);
                font-weight: 700;
                color: #f59e0b;
                margin-top: 1.2rem;
                display: inline-block;
                background: rgba(0,0,0,0.25);
                padding: 6px 18px;
                border-radius: 40px;
                backdrop-filter: blur(4px);
                white-space: nowrap;
            }}
            .slider-progress {{
                position: absolute;
                bottom: 0;
                left: 0;
                height: 4px;
                background: #f59e0b;
                width: 0%;
                z-index: 10;
                transition: width 0.1s linear;
            }}
            .slider-nav {{
                position: absolute;
                bottom: 20px;
                left: 50%;
                transform: translateX(-50%);
                display: flex;
                gap: 15px;
                z-index: 10;
            }}
            .slider-dot {{
                width: 12px;
                height: 12px;
                border-radius: 50%;
                background: rgba(255,255,255,0.5);
                cursor: pointer;
                transition: all 0.3s;
            }}
            .slider-dot.active {{
                background: #f59e0b;
                width: 30px;
                border-radius: 10px;
            }}
            .slider-arrow {{
                position: absolute;
                top: 50%;
                transform: translateY(-50%);
                z-index: 10;
                background: rgba(255,255,255,0.15);
                border: 2px solid rgba(255,255,255,0.4);
                color: white;
                width: 45px;
                height: 45px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                cursor: pointer;
                font-size: 1.2rem;
                transition: all 0.3s;
                backdrop-filter: blur(4px);
            }}
            .slider-arrow:hover {{
                background: #f59e0b;
                border-color: #f59e0b;
                color: #1e293b;
            }}
            .slider-arrow.prev {{ left: 20px; }}
            .slider-arrow.next {{ right: 20px; }}
            
            /* ===== SCROLL REVEAL ===== */
            .reveal {{
                opacity: 0;
                transform: translateY(30px);
                transition: opacity 0.7s ease, transform 0.7s ease;
            }}
            .reveal.visible {{
                opacity: 1;
                transform: translateY(0);
            }}
            
            /* ===== STATS SECTION ===== */
            .stats-section {{
                padding: 60px 30px;                                 
                background: #0f3b5c;
                color: white;
            }}
            .stats-grid {{
                display: flex;
                justify-content: space-between;
                max-width: 1200px;
                margin: 0 auto;
                gap: 0;
                flex-wrap: nowrap;
            }}
            .stat-card {{
                text-align: center;
                flex: 1;
                min-width: 0;
                padding: 28px 12px;
                border-right: 1px solid rgba(255,255,255,0.2);
                transition: transform 0.3s;
            }}
            .stat-card:hover {{ transform: translateY(-5px); }}
            .stat-card:last-child {{ border-right: none; }}
            .stat-number {{
                font-size: clamp(1.4rem, 2.2vw, 2.4rem);
                font-weight: 800;
                color: #f59e0b;
                margin-bottom: 8px;
                white-space: nowrap;
                line-height: 1.2;
            }}
            .stat-label {{
                font-size: clamp(0.7rem, 1vw, 0.85rem);
                text-transform: uppercase;
                letter-spacing: 1px;
                font-weight: 500;
                white-space: nowrap;
            }}
            
            /* ===== ABOUT & SERVICES ===== */
            .about-section {{
                padding: 80px 30px;                                                       
                background: #f8fafc;
            }}
            .about-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 60px;
                max-width: 1280px;
                margin: 0 auto;
            }}
            .about-tag {{
                color: #f59e0b;
                font-weight: 700;
                letter-spacing: 2px;
                margin-bottom: 1rem;
                font-size: 0.8rem;
            }}
            .about-title {{
                font-size: 2.5rem;
                font-weight: 700;
                color: #0f3b5c;
                margin-bottom: 1.5rem;
            }}
            .about-text {{
                color: #475569;
                line-height: 1.7;
                margin-bottom: 1.5rem;
            }}
            .about-highlight {{
                background: white;
                padding: 20px;
                border-radius: 16px;
                border-left: 4px solid #f59e0b;
            }}
            .about-img {{
                width: 100%;
                border-radius: 24px;
                box-shadow: 0 20px 30px -15px rgba(0,0,0,0.15);
            }}
            .services-section {{
                padding: 80px 30px;                                         
                background: white;
            }}
            .section-header {{
                text-align: center;
                margin-bottom: 50px;
            }}
            .section-header h2 {{
                font-size: 2.2rem;
                color: #0f3b5c;
            }}
            .services-grid {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 30px;
                max-width: 1280px;
                margin: 0 auto;
            }}
            .service-card {{
                background: white;
                padding: 30px 20px;
                border-radius: 20px;
                text-align: center;
                box-shadow: 0 4px 15px rgba(0,0,0,0.05);
                border: 1px solid #e2e8f0;
                transition: all 0.3s;
            }}
            .service-card:hover {{
                transform: translateY(-8px);
                border-color: #f59e0b;
            }}
            .service-icon {{
                font-size: 3rem;
                color: #f59e0b;
                margin-bottom: 20px;
            }}
            .infra-section {{
                padding: 80px 30px;                                         
                background: #f8fafc;
            }}
            .infra-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 50px;
                max-width: 1280px;
                margin: 0 auto;
            }}
            .infra-feature {{
                display: flex;
                gap: 15px;
                margin-bottom: 25px;
            }}
            .infra-feature i {{
                font-size: 1.8rem;
                color: #f59e0b;
            }}
            .careers-section {{
                padding: 80px 30px;                                             
                background: linear-gradient(135deg, #0f3b5c 0%, #1e4a76 100%);
                color: white;
                text-align: center;
            }}
            .btn-white {{
                background: white;
                color: #0f3b5c;
                padding: 12px 35px;
                border-radius: 40px;
                font-weight: 700;
                display: inline-block;
                margin-top: 20px;
                text-decoration: none;
                cursor: pointer;
            }}
            
            /* ===== FOOTER ===== */
            .footer {{
                background: #0f172a;
                color: #cbd5e1;
                padding: 50px 30px 30px;                                        
                width: 100%;
                clear: both;
            }}
            .footer-grid {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 40px;
                max-width: 1280px;
                margin: 0 auto;
                padding-bottom: 40px;
                border-bottom: 1px solid #334155;
            }}
            .footer-col h4 {{
                color: white;
                margin-bottom: 20px;
                font-size: 1.1rem;
            }}
            .footer-col p, .footer-col a {{
                color: #94a3b8;
                text-decoration: none;
                line-height: 1.8;
                font-size: 0.9rem;
                display: block;
                cursor: pointer;
            }}
            .footer-col a:hover {{
                color: #f59e0b;
            }}
            .copyright {{
                text-align: center;
                padding-top: 30px;
                font-size: 0.8rem;
                color: #64748b;
            }}
            
            /* ===== MODAL ===== */
            .modal {{
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100vw;
                height: 100vh;
                z-index: 99999;
                background: rgba(0,0,0,0.85);
                align-items: flex-start;
                justify-content: center;
                overflow-y: auto;
                padding: 80px 20px 20px 20px;
            }}
            .modal.active {{
                display: flex;
            }}
            .modal-content {{
                max-width: 900px;
                width: 100%;
                background: white;
                border-radius: 8px;
                box-shadow: 0 25px 60px rgba(0,0,0,0.4);
                overflow: hidden;
                animation: modalFadeIn 0.3s ease;
            }}
            @keyframes modalFadeIn {{
                from {{ opacity: 0; transform: scale(0.95); }}
                to {{ opacity: 1; transform: scale(1); }}
            }}
            .modal-header {{
                display: flex;
                justify-content: flex-end;
                padding: 10px 15px;
                background: #f0f2f5;
                border-bottom: 1px solid #ddd;
            }}
            .modal-close {{
                background: none;
                border: none;
                font-size: 24px;
                cursor: pointer;
                color: #666;
                transition: color 0.2s;
            }}
            .modal-close:hover {{
                color: #f59e0b;
            }}
            .modal-body {{
                padding: 30px 40px;
                max-height: 80vh;
                overflow-y: auto;
            }}
            .a4-chairman {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 20px;
                margin-bottom: 30px;
                padding: 20px;
                background: #f8f9fa;
                border-radius: 16px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            }}
            .a4-chairman-info {{
                flex: 2;
            }}
            .a4-chairman-avatar {{
                flex: 0 0 auto;
                width: 150px;
                height: 150px;
            }}
            .a4-chairman-avatar img {{
                width: 100%;
                height: 100%;
                border-radius: 50%;
                object-fit: cover;
                border: 3px solid #f59e0b;
                box-shadow: 0 4px 15px rgba(0,0,0,0.15);
            }}
            .a4-chairman-logo {{
                flex: 0 0 auto;
                width: 80px;
                height: 80px;
            }}
            .a4-chairman-logo img {{
                width: 100%;
                height: 100%;
                border-radius: 50%;
                object-fit: cover;
                border: 2px solid #ddd;
                background: white;
                padding: 5px;
            }}
            .a4-chairman-info h2 {{
                font-size: 1.3rem;
                color: #0f3b5c;
                margin-bottom: 5px;
            }}
            .a4-chairman-info .title {{
                color: #f59e0b;
                font-weight: 600;
                margin-bottom: 8px;
            }}
            .a4-chairman-info .company {{
                font-size: 0.8rem;
                color: #666;
                line-height: 1.4;
                font-weight: 500;
            }}
            .a4-body {{
                line-height: 1.7;
                color: #333;
            }}
            .modal-body .a4-body p {{
                text-align: justify;
                text-justify: inter-ideograph;
            }}
            .a4-date {{
                text-align: right !important;
                font-style: italic;
                margin-bottom: 20px;
                color: #666;
            }}
            .vision-box, .mission-box {{
                background: #f0f7ff;
                padding: 20px;
                border-radius: 12px;
                margin: 20px 0;
                border-left: 4px solid #f59e0b;
            }}
            .vision-box h3, .mission-box h3 {{
                color: #0f3b5c;
                margin-bottom: 12px;
                font-size: 1.1rem;
            }}
            .a4-signature-left {{
                margin-top: 40px;
                text-align: left;
            }}
            .sig-block-left {{
                display: inline-block;
                text-align: center;
            }}
            .sig-block-left .sig-image img {{
                max-width: 386px;
                height: auto;
            }}
            .a4-footer {{
                margin-top: 30px;
                padding-top: 15px;
                border-top: 1px solid #ddd;
                text-align: center;
                font-size: 0.75rem;
                color: #999;
                display: flex;
                justify-content: space-between;
                flex-wrap: wrap;
            }}
            
            /* ===== RESPONSIVE ===== */
            @media (max-width: 768px) {{
                .slide-layout {{
                    flex-direction: column;
                }}
                .slide-image {{
                    width: 100%;
                    height: 35%;
                    flex: none;
                }}
                .slide-content {{
                    padding: 25px 20px;
                }}
                .slide-content h1 {{
                    font-size: 1.6rem;
                }}
                .slide-content p {{
                    font-size: 0.9rem;
                }}
                .hero-slider {{
                    height: auto;
                    min-height: 480px;
                }}
                .stats-grid, .about-grid, .services-grid, .infra-grid, .footer-grid {{
                    grid-template-columns: 1fr;
                }}
                .stat-card {{
                    border-right: none;
                    border-bottom: 1px solid rgba(255,255,255,0.2);
                }}
                .nav-links {{
                    display: none;
                }}
                .logo-circle {{
                    width: 60px;
                    height: 60px;
                }}
                .a4-chairman {{
                    flex-wrap: wrap;
                    justify-content: center;
                    text-align: center;
                }}
                .modal-body {{
                    padding: 20px;
                }}
                .mobile-lang {{
                    display: flex !important;
                }}
                .nav-links .lang-switch {{
                    display: none;
                }}
            }}
        </style>
    </head>
    <body>
    
    <!-- Navigation -->
    <nav class="navbar" id="navbar">
        <div class="nav-container">
            <div class="logo-circle">
                <img src="data:image/png;base64,{logo_base64}" alt="Cảng Hòn La">
            </div>
            <div class="nav-links">
                <a class="nav-link" data-section="home">{text.get('nav_home', 'Trang chủ')}</a>
                <div class="dropdown">
                    <a class="nav-link" data-section="about">{text.get('nav_about', 'Giới thiệu')} <i class="fas fa-chevron-down"></i></a>
                    <div class="dropdown-content">
                        <a class="nav-link" data-section="about">{text.get('about_us', 'Về chúng tôi')}</a>
                        <a href="#" id="thuNgoBtn">{text.get('chairman_letter', 'Thư ngỏ của Chủ tịch HĐQT')}</a>
                    </div>
                </div>
                <a class="nav-link" data-section="services">{text.get('nav_services', 'Dịch vụ')}</a>
                <a class="nav-link" data-section="infrastructure">{text.get('nav_infrastructure', 'Vị trí & Hạ tầng')}</a>
                <a class="nav-link" data-section="careers">{text.get('nav_careers', 'Tuyển dụng')}</a>
                <a class="nav-link" data-section="contact">{text.get('nav_contact', 'Liên hệ')}</a>
                <span class="nav-divider">|</span>
                <div class="lang-switch">
                    <a href="#" class="lang-link {vi_active}" onclick="switchLanguage('vi'); return false;">🇻🇳 VI</a>
                    <span class="lang-sep">/</span>
                    <a href="#" class="lang-link {en_active}" onclick="switchLanguage('en'); return false;">🇬🇧 EN</a>
                </div>
            </div>
        </div>
    </nav>

    <!-- Mobile Language Switcher -->
    <div class="mobile-lang">
        <a href="#" class="{vi_active}" onclick="switchLanguage('vi'); return false;">🇻🇳 VI</a>
        <span style="color:white; opacity:0.5;">|</span>
        <a href="#" class="{en_active}" onclick="switchLanguage('en'); return false;">🇬🇧 EN</a>
    </div>

    <!-- Modal Thư ngỏ -->
    <div id="thuNgoModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <button class="modal-close" id="closeModalBtn">✕</button>
            </div>
            <div class="modal-body">
                <div class="a4-chairman">
                    <div class="a4-chairman-avatar">
                        <img src="{chu_tich_img}" alt="Chủ tịch HĐQT">
                    </div>
                    <div class="a4-chairman-info">
                        <h2>Ông Phùng Gia Phát</h2>
                        <p class="title">{text.get('modal_chairman_title', 'Chủ tịch Hội đồng Quản trị')}</p>
                        <p class="company">Công ty Cổ phần Cảng Hòn La</p>
                        <p class="company">Khu kinh tế Hòn La, Xã Quảng Đông, Huyện Quảng Trạch, Tỉnh Quảng Bình</p>
                    </div>                  
                    <div class="a4-chairman-logo">
                        <img src="data:image/png;base64,{logo_base64}" alt="Logo Cảng Hòn La">
                    </div>
                </div>
                <div class="a4-body">
                    <p class="a4-date">Quảng Bình, ngày 21 tháng 3 năm 2025</p>
                    <p class="a4-greeting" style="font-weight: bold; font-size: 1rem;">{text.get('modal_greeting', 'Kính gửi: Quý đối tác, nhà đầu tư và toàn thể cán bộ nhân viên,')}</p>
                    <p>{text.get('modal_content_1', 'Với niềm tự hào sâu sắc, Tôi xin thay mặt Hội đồng Quản trị Công ty Cổ phần Cảng Hòn La gửi lời chào trân trọng nhất đến Quý đối tác, nhà đầu tư và toàn thể cán bộ nhân viên — những người đã và đang đồng hành cùng chúng tôi trên hành trình kiến tạo một cảng biển tầm cỡ quốc tế giữa lòng đất nước Việt Nam.')}</p>
                    <p>{text.get('modal_content_2', 'Ngày 21 tháng 3 năm 2025 là một mốc son lịch sử — ngày chính thức khởi công Dự án Cảng tổng hợp quốc tế Hòn La, dự án được Chính phủ công nhận là Dự án trọng điểm Quốc gia. Đây không chỉ là thành quả của nhiều năm nỗ lực không ngừng, mà còn là khởi đầu của một chương mới trong lịch sử phát triển kinh tế hàng hải miền Trung Việt Nam.')}</p>
                    <div class="vision-box">
                        <h3>{text.get('modal_vision_title', '🎯 Tầm nhìn — Vision 2035')}</h3>
                        <p>{text.get('modal_vision_text', 'Trở thành cảng biển quốc tế hiện đại hàng đầu Đông Nam Á trên tuyến hành lang kinh tế Đông–Tây (EWEC) — nơi kết nối Việt Nam với thế giới, thúc đẩy thương mại, logistic và du lịch tàu biển, đóng góp thiết thực vào chiến lược phát triển kinh tế biển bền vững của Việt Nam đến năm 2035 và tầm nhìn 2045.')}</p>
                    </div>
                    <p>{text.get('modal_content_3', 'Với vị trí địa chiến lược độc đáo, hệ thống hạ tầng quy mô 39,22 ha, năng lực tiếp nhận tàu trọng tải lên đến 70.000 DWT và tàu du lịch quốc tế 225.000 GT, Cảng tổng hợp quốc tế Hòn La sẽ là cửa ngõ hàng hải chiến lược, cầu nối giữa các nền kinh tế trong khu vực và toàn cầu.')}</p>
                    <div class="mission-box">
                        <h3>{text.get('modal_mission_title', '💡 Sứ mệnh - Nhắn gửi đến mỗi thành viên')}</h3>
                        <p>{text.get('modal_mission_text', 'Mỗi cán bộ nhân viên của Công ty cổ phần Cảng Hòn La là một đại sứ của sự chuyên nghiệp và tận tâm. Sứ mệnh của chúng ta là xây dựng một môi trường làm việc đẳng cấp, nơi năng lực được trọng dụng, sáng tạo được khuyến khích và mỗi cá nhân đều tự hào khi đặt bàn tay mình vào công trình lịch sử này. Hãy làm việc với trái tim của người kiến tạo — bởi di sản chúng ta để lại không chỉ là những cầu bến vững chắc, mà còn là những thế hệ nhân lực xuất sắc của đất nước.')}</p>
                    </div>
                    <p>{text.get('modal_content_4', 'Chúng tôi hiểu rằng con đường phía trước còn không ít thách thức. Song Tôi tin tưởng sâu sắc rằng với trí tuệ tập thể, khí phách dân tộc và khát vọng vươn ra biển lớn, Cảng tổng hợp quốc tế Hòn La sẽ hoàn thành xuất sắc sứ mệnh lịch sử được giao phó.')}</p>
                    <p>{text.get('modal_thanks', 'Xin trân trọng cảm ơn sự tin tưởng, đồng hành và cống hiến của tất cả Quý vị.')}<br>{text.get('modal_wishes', 'Chúc Quý đối tác thịnh vượng, toàn thể cán bộ nhân viên sức khỏe và thành công!')}</p>
                </div>
                <div class="a4-signature-left">
                    <div class="sig-block-left">
                        <div class="sig-image">
                            <img src="{chu_ky_img}" alt="Chữ ký Chủ tịch">
                        </div>
                    </div>
                </div>
                <div class="a4-footer">
                    <span>🌐 honlaport.com.vn</span>
                    <span>✉ info@honlaport.com.vn</span>
                    <span>📞 0232.xxxx.xxx</span>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Hero Slider -->
    <section id="home" class="hero-slider">
        <div class="slides-container">
            <div class="slide active">
                <div class="slide-layout">
                    <div class="slide-image" style="background-image: url('{slide1_src}');"></div>
                    <div class="slide-content">
                        <h1>{text.get('hero_title_1', 'CẢNG TỔNG HỢP QUỐC TẾ HÒN LA')}</h1>
                        <p>{text.get('hero_desc_1_1', 'Chính thức khởi công ngày 21 tháng 3 năm 2025')}</p>
                        <p>{text.get('hero_desc_1_2', 'Đưa vào khai thác từ Tháng 5 năm 2026')}</p>
                        <div class="highlight">{text.get('hero_tag_1', '🚢 Cửa ngõ hàng hải chiến lược của Miền Trung')}</div>
                    </div>
                </div>
            </div>
            <div class="slide">
                <div class="slide-layout">
                    <div class="slide-content">
                        <h1>{text.get('hero_title_2', 'KẾT NỐI TOÀN CẦU')}</h1>
                        <p>{text.get('hero_desc_2_1', 'Vị trí chiến lược trên tuyến hành lang kinh tế Đông - Tây (EWEC)')}</p>
                        <p>{text.get('hero_desc_2_2', 'Kết nối trực tiếp với các cảng biển lớn trong khu vực và quốc tế')}</p>
                        <div class="highlight">{text.get('hero_tag_2', '🌏 Hành lang thương mại huyết mạch')}</div>
                    </div>
                    <div class="slide-image" style="background-image: url('{slide2_src}');"></div>
                </div>
            </div>
            <div class="slide">
                <div class="slide-layout">
                    <div class="slide-image" style="background-image: url('{slide3_src}');"></div>
                    <div class="slide-content">
                        <h1>{text.get('hero_title_3', 'HẠ TẦNG ĐẲNG CẤP QUỐC TẾ')}</h1>
                        <p>{text.get('hero_desc_3_1', '04 Cầu Tàu | Tổng chiều dài 970m | Tiếp nhận tàu 70.000 DWT')}</p>
                        <p>{text.get('hero_desc_3_2', 'Tàu du lịch quốc tế 225.000 GT')}</p>
                        <div class="highlight">{text.get('hero_tag_3', '⚓ Hiện đại - Đồng bộ - Chuyên nghiệp')}</div>
                    </div>
                </div>
            </div>
        </div>
        <div class="slider-arrow prev" id="prevBtn">&#8592;</div>
        <div class="slider-arrow next" id="nextBtn">&#8594;</div>
        <div class="slider-nav">
            <div class="slider-dot active" data-slide="0"></div>
            <div class="slider-dot" data-slide="1"></div>
            <div class="slider-dot" data-slide="2"></div>
        </div>
        <div class="slider-progress" id="sliderProgress"></div>
    </section>
    
    <!-- Statistics -->
    <section id="stats" class="stats-section">
        <div class="stats-grid">
            <div class="stat-card reveal"><div class="stat-number">39,22 ha</div><div class="stat-label">{text.get('stat_total_area', 'Tổng diện tích')}</div></div>
            <div class="stat-card reveal"><div class="stat-number">70.000 DWT</div><div class="stat-label">{text.get('stat_max_capacity', 'Trọng tải tàu tối đa')}</div></div>
            <div class="stat-card reveal"><div class="stat-number">970 m</div><div class="stat-label">{text.get('stat_berth_length', 'Chiều dài cầu cảng')}</div></div>
            <div class="stat-card reveal"><div class="stat-number">225.000 GT</div><div class="stat-label">{text.get('stat_cruise_ship', 'Tàu du lịch quốc tế')}</div></div>
        </div>
    </section>
    
    <!-- About -->
    <section id="about" class="about-section">
        <div class="about-grid">
            <div>
                <div class="about-tag">{text.get('about_tag', 'CHÀO MỪNG ĐẾN VỚI CẢNG QUỐC TẾ HÒN LA')}</div>
                <h2 class="about-title">{text.get('about_title', 'Cửa ngõ hàng hải chiến lược của Miền Trung')}</h2>
                <p class="about-text">{text.get('about_text', 'Cảng tổng hợp Quốc tế Hòn La được đầu tư bài bản với hệ thống cơ sở hạ tầng đồng bộ, hiện đại, đáp ứng nhu cầu bốc xếp hàng hóa, trung chuyển container và đón tàu du lịch quốc tế.')}</p>
                <div class="about-highlight"><i class="fas fa-trophy" style="color:#f59e0b"></i> <strong>{text.get('about_highlight', 'Dự án trọng điểm Quốc gia')}</strong></div>
            </div>
            <div><img src="https://images.unsplash.com/photo-1562329264-a2c2d4112b8d?q=80&w=2070" class="about-img"></div>
        </div>
    </section>
    
    <!-- Services -->
    <section id="services" class="services-section">
        <div class="section-header"><h2>{text.get('services_title', 'Dịch vụ của chúng tôi')}</h2></div>
        <div class="services-grid">
            <div class="service-card"><div class="service-icon"><i class="fas fa-ship"></i></div><h3>{text.get('service_bulk', 'Hàng rời & Hàng khô')}</h3></div>
            <div class="service-card"><div class="service-icon"><i class="fas fa-boxes"></i></div><h3>{text.get('service_container', 'Hàng container')}</h3></div>
            <div class="service-card"><div class="service-icon"><i class="fas fa-umbrella-beach"></i></div><h3>{text.get('service_cruise', 'Du lịch tàu biển')}</h3></div>
            <div class="service-card"><div class="service-icon"><i class="fas fa-warehouse"></i></div><h3>{text.get('service_logistics', 'Logistics & Kho bãi')}</h3></div>
        </div>
    </section>
    
    <!-- Infrastructure -->
    <section id="infrastructure" class="infra-section">
        <div class="infra-grid">
            <div>
                <div class="about-tag">{text.get('infra_tag', 'HẠ TẦNG & VỊ TRÍ')}</div>
                <h2 class="about-title">{text.get('infra_title', 'Vị thế vàng trên bản đồ logistics')}</h2>
                <div class="infra-feature"><i class="fas fa-map-marker-alt"></i><div><strong>Quảng Trạch, Quảng Bình</strong><br>{text.get('infra_location', 'Khu kinh tế Hòn La')}</div></div>
                <div class="infra-feature"><i class="fas fa-road"></i><div><strong>{text.get('infra_connection', 'Kết nối hành lang Đông - Tây (EWEC)')}</strong></div></div>
                <div class="infra-feature"><i class="fas fa-anchor"></i><div><strong>{text.get('infra_berths', '04 bến cấp tàu')}</strong><br>Tổng chiều dài 970m</div></div>
            </div>
            <div><img src="https://images.unsplash.com/photo-1578575437130-527eed3abbec?q=80&w=2070" class="about-img"></div>
        </div>
    </section>
    
    <!-- Careers -->
    <section id="careers" class="careers-section">
        <h2>{text.get('careers_title', 'GIA NHẬP ĐỘI NGŨ NHÂN SỰ CỦA CHÚNG TÔI')}</h2>
        <p>{text.get('careers_subtitle', 'Chúng tôi luôn tìm kiếm những nhân tài')}</p>
        <a href="#" class="btn-white" id="careerLink">{text.get('careers_button', '📢 Xem cơ hội việc làm tại đây')}</a>
    </section>
    
    <!-- Footer -->
    <footer id="contact" class="footer">
        <div class="footer-grid">
            <div class="footer-col"><h4 style="font-size:0.95rem; white-space:nowrap;">{text.get('footer_company', 'CÔNG TY CỔ PHẦN CẢNG HÒN LA')}</h4><p>Khu kinh tế Hòn La, Xã Phú Trạch, Tỉnh Quảng Trị</p><p>📞 0232.xxxx.xxx</p><p>📧 info@honlaport.com.vn</p></div>
            <div class="footer-col"><h4>{text.get('footer_quick_links', 'Liên kết nhanh')}</h4>
                <a class="nav-link" data-section="home">{text.get('nav_home', 'Trang chủ')}</a>
                <a class="nav-link" data-section="about">{text.get('nav_about', 'Về chúng tôi')}</a>
                <a class="nav-link" data-section="services">{text.get('nav_services', 'Dịch vụ')}</a>
                <a class="nav-link" data-section="infrastructure">{text.get('nav_infrastructure', 'Hạ tầng')}</a>
                <a class="nav-link" data-section="careers">{text.get('nav_careers', 'Tuyển dụng')}</a>
            </div>
            <div class="footer-col"><h4>{text.get('footer_support', 'Hỗ trợ')}</h4><a href="#">{text.get('footer_faq', 'Câu hỏi thường gặp')}</a><a href="#">{text.get('footer_privacy', 'Chính sách bảo mật')}</a><a href="#">{text.get('footer_terms', 'Điều khoản sử dụng')}</a></div>
            <div class="footer-col"><h4>{text.get('footer_working_hours', 'Giờ làm việc')}</h4><p>🚢 {text.get('footer_working_hours_port', 'Bến cảng: 24/7')}</p><p>🏢 {text.get('footer_working_hours_office', 'Văn phòng: 7:30 - 17:00')}</p><p>📅 {text.get('footer_working_days', 'Thứ 2 - Thứ 7')}</p></div>
        </div>
        <div class="copyright">
            <p>© 2026 - Công ty Cổ phần Cảng Hòn La. All rights reserved.</p>
            <p style="margin-top: 10px;">{text.get('footer_copyright', 'PHÁT TRIỂN BỀN VỮNG - KẾT NỐI TOÀN CẦU')}</p>
        </div>
    </footer>
    
    {landing_js}
    </body>
    </html>
    """
    
    # Render landing page
    components.html(landing_html, height=3150, scrolling=False)
    
    # Nút HRM dùng components.html (giữ nguyên phần còn lại)
    hrm_html = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    background: linear-gradient(135deg, #0f3b5c 0%, #1a4a6e 100%);
    display: flex; justify-content: center; align-items: center;
    min-height: 100px;
    border-top: 3px solid #f59e0b;
    border-bottom: 3px solid #f59e0b;
    padding: 20px;
}
.hrm-button {
    background: linear-gradient(135deg, #f59e0b 0%, #e67e22 100%);
    color: #0f3b5c; font-weight: 800; font-size: 1.2rem;
    border: none; border-radius: 60px; padding: 18px 60px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.3); letter-spacing: 1px;
    cursor: pointer; transition: all 0.3s ease; min-width: 420px;
    font-family: sans-serif;
}
.hrm-button:hover {
    background: linear-gradient(135deg, #e67e22 0%, #d35400 100%);
    transform: translateY(-3px); box-shadow: 0 12px 30px rgba(0,0,0,0.4);
}
@media (max-width: 768px) {
    .hrm-button { font-size: 0.9rem; padding: 14px 30px; min-width: 260px; }
}
</style>
</head>
<body>
    <button class="hrm-button" id="hrmBtn">
        🔐 HRM - QUẢN LÝ NHÂN SỰ / Chỉ dành cho Nhân viên
    </button>
    <script>
    document.getElementById('hrmBtn').addEventListener('click', function() {
        var topWin = window.top || window.parent || window;
        var url = new URL(topWin.location.href);
        url.searchParams.set('goto', 'hrm');
        topWin.location.href = url.toString();
    });
    </script>
</body>
</html>"""
 
    st.markdown("""
        <style>
            .hrm-button-container {
                background: linear-gradient(135deg, #0f3b5c 0%, #1a4a6e 100%);
                border-top: 3px solid #f59e0b;
                border-bottom: 3px solid #f59e0b;
                padding: 20px;
                text-align: center;
            }
            .stButton > button {
                background: linear-gradient(135deg, #f59e0b 0%, #e67e22 100%);
                color: #0f3b5c !important;
                font-weight: 800;
                font-size: 1.2rem;
                border: none;
                border-radius: 60px;
                padding: 18px 60px;
                box-shadow: 0 8px 25px rgba(0,0,0,0.3);
                min-width: 420px;
                transition: all 0.3s ease;
                width: auto !important;
            }
            .stButton > button:hover {
                background: linear-gradient(135deg, #e67e22 0%, #d35400 100%);
                transform: translateY(-3px);
                box-shadow: 0 12px 30px rgba(0,0,0,0.4);
            }
            @media (max-width: 768px) {
                .stButton > button {
                    font-size: 0.9rem;
                    padding: 14px 30px;
                    min-width: 260px;
                }
            }
        </style>
        <div class="hrm-button-container">
    """, unsafe_allow_html=True)

