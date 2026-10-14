# -*- coding: utf-8 -*-
"""
landing_page/landing_0108872052.py
===================================
Landing Page NỘI BỘ (dành cho CBCNV) — Công ty Cổ phần Cảng Hòn La.

Chuyển thể từ landing page đối ngoại (khách hàng, đối tác) → Bản tin nội bộ
cho CBCNV. Mục tiêu: Truyền tải Sứ mệnh, Tầm nhìn, Văn hóa doanh nghiệp để
tạo động lực và gắn kết nội bộ.

HỢP ĐỒNG BẮT BUỘC: file phải có hàm render() không nhận tham số — app.py chỉ
gọi module.render().
"""

import streamlit as st
import streamlit.components.v1 as components
import os
import base64
import requests

# Import config - ưu tiên config.py (local), fallback to config_template (cloud).
try:
    from config import COMPANY_CONFIG
except ImportError:
    from config_template import COMPANY_CONFIG


def render():
    """Hiển thị Bản tin nội bộ với chuyển ngữ Việt/Anh."""

    try:
        from languages import LANGUAGES
    except ImportError:
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
            [data-testid="stVerticalBlock"] {
                gap: 0 !important;
            }
            [data-testid="stElementContainer"],
            .element-container {
                margin: 0 !important;
            }
            [data-testid="stMain"],
            [data-testid="stAppViewContainer"] > .main {
                padding: 0 !important;
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
        </style>
    """, unsafe_allow_html=True)

    # Đọc file logo động
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
        logo_path = os.path.join(os.path.dirname(__file__), "logo_cty.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                logo_base64 = base64.b64encode(f.read()).decode()

    def load_img_b64(filename):
        thu_muc_goc = os.path.dirname(os.path.dirname(__file__))
        path = os.path.join(thu_muc_goc, "static", filename)
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

    vi_active = 'active' if lang == 'vi' else ''
    en_active = 'active' if lang == 'en' else ''

    landing_js = """
    <script>
        function scrollToSection(sectionId) {
            var topWin = window.top || window.parent || window;
            var targetElement = document.getElementById(sectionId);
            if (targetElement) {
                var targetRect = targetElement.getBoundingClientRect();
                var offsetTop = targetRect.top + (topWin.scrollY || topWin.pageYOffset);
                topWin.scrollTo({ top: offsetTop - 80, behavior: 'smooth' });
            }
        }
        function handleNavClick(e) {
            e.preventDefault();
            var section = this.getAttribute('data-section');
            if (section) scrollToSection(section);
        }
        function initNavigation() {
            var navLinks = document.querySelectorAll('.nav-link');
            for (var i = 0; i < navLinks.length; i++) {
                navLinks[i].removeEventListener('click', handleNavClick);
                navLinks[i].addEventListener('click', handleNavClick);
            }
        }
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
        function nextSlide() { showSlide((currentSlide + 1) % totalSlides); }
        function prevSlide() { showSlide((currentSlide - 1 + totalSlides) % totalSlides); }
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
            if (prevBtn) prevBtn.addEventListener('click', function() {
                prevSlide();
                if (autoSlideInterval) clearInterval(autoSlideInterval);
                if (progressInterval) clearInterval(progressInterval);
                startAutoSlide();
            });
            if (nextBtn) nextBtn.addEventListener('click', function() {
                nextSlide();
                if (autoSlideInterval) clearInterval(autoSlideInterval);
                if (progressInterval) clearInterval(progressInterval);
                startAutoSlide();
            });
            startAutoSlide();
        }
        var revealObserver = new IntersectionObserver(function(entries) {
            entries.forEach(function(entry) {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                    revealObserver.unobserve(entry.target);
                }
            });
        }, { threshold: 0.15 });
        document.querySelectorAll('.reveal').forEach(function(el) { revealObserver.observe(el); });
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
            modal.addEventListener('click', function(e) { if (e.target === modal) closeModal(); });
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape' && modal.classList.contains('active')) closeModal();
            });
        }
        function switchLanguage(lang) {
            var topWin = window.top || window.parent || window;
            var url = new URL(topWin.location.href);
            url.searchParams.set('lang', lang);
            topWin.history.replaceState(null, '', url.toString());
            topWin.location.reload();
        }
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', function() { initNavigation(); });
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
        <title>Bản tin nội bộ — Hòn La</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;14..32,400;14..32,500;14..32,600;14..32,700;14..32,800&family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&display=swap" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <style>
            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
            body {{
                font-family: 'Inter', sans-serif;
                background-color: #ffffff;
                color: #1e293b;
                line-height: 1.5;
                overflow-x: hidden;
                width: 100%;
                padding-top: 100px;
            }}
            ::-webkit-scrollbar {{ width: 8px; }}
            ::-webkit-scrollbar-track {{ background: #f1f1f1; }}
            ::-webkit-scrollbar-thumb {{ background: #0f3b5c; border-radius: 4px; }}

            /* ===== NAVIGATION ===== */
            .navbar {{
                position: fixed; top: 0; left: 0; width: 100%;
                z-index: 1000; padding: 0.8rem 30px;
                background: rgba(15, 59, 92, 0.95);
                backdrop-filter: blur(10px);
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            .nav-container {{
                display: flex; justify-content: space-between; align-items: center;
                max-width: 1400px; margin: 0 auto;
            }}
            .logo-circle {{
                width: 86px; height: 86px; border-radius: 50%; overflow: hidden;
                box-shadow: 0 8px 25px rgba(0,0,0,0.25), 0 0 0 3px rgba(255,255,255,0.3);
                background: white; display: flex; align-items: center; justify-content: center;
                transition: transform 0.3s, box-shadow 0.3s; cursor: pointer;
            }}
            .logo-circle:hover {{ transform: scale(1.02); }}
            .logo-circle img {{ width: 100%; height: 100%; object-fit: cover; }}
            .nav-links {{
                display: flex; gap: 0.5rem; align-items: center;
                background: rgba(0,0,0,0.35);
                padding: 5px 15px; border-radius: 50px;
                backdrop-filter: blur(5px);
            }}
            .nav-links a {{
                text-decoration: none; color: white; font-weight: 500;
                font-size: 0.9rem; padding: 8px 16px; border-radius: 40px;
                transition: all 0.3s; cursor: pointer;
            }}
            .nav-links a:hover {{ background: #f59e0b; color: #0f3b5c; }}
            .nav-links .nav-divider {{ color: rgba(255,255,255,0.4); margin: 0 5px; font-size: 14px; }}
            .lang-switch {{ display: inline-flex; align-items: center; gap: 5px; margin-left: 5px; }}
            .lang-link {{
                text-decoration: none !important; color: white !important;
                font-weight: 500; font-size: 0.85rem; padding: 8px 8px !important;
                border-radius: 40px; transition: all 0.3s; background: transparent !important; cursor: pointer;
            }}
            .lang-link:hover {{ background: #f59e0b !important; color: #0f3b5c !important; }}
            .lang-link.active {{ background: #f59e0b !important; color: #0f3b5c !important; }}
            .lang-sep {{ color: rgba(255,255,255,0.5); font-size: 12px; }}

            .mobile-lang {{
                position: fixed; top: 15px; right: 15px; z-index: 10001;
                background: rgba(0,0,0,0.6); backdrop-filter: blur(8px);
                border-radius: 30px; padding: 6px 12px; display: none; gap: 8px;
                border: 1px solid rgba(255,255,255,0.2);
            }}
            .mobile-lang a {{
                color: white; text-decoration: none; font-size: 12px; font-weight: 600;
                padding: 4px 8px; border-radius: 20px; transition: all 0.2s; cursor: pointer;
            }}
            .mobile-lang a.active {{ background: #f59e0b; color: #0f3b5c; }}

            /* ===== HERO SLIDER ===== */
            .hero-slider {{
                height: 550px; width: 100%; position: relative;
                overflow: hidden; margin-top: 0; background-color: #0a2a3a;
            }}
            .slides-container {{ position: relative; width: 100%; height: 100%; }}
            .slide {{
                position: absolute; top: 0; left: 0; width: 100%; height: 100%;
                opacity: 0; transition: opacity 0.8s ease-in-out; background-color: #0a2a3a;
            }}
            .slide.active {{ opacity: 1; z-index: 2; }}
            .slide-layout {{
                display: flex; width: 100%; height: 100%;
                align-items: center; justify-content: center;
            }}
            .slide-image {{
                flex: 1; height: 100%; background-size: cover;
                background-position: center center; background-repeat: no-repeat;
            }}
            .slide-content {{
                flex: 1; padding: 50px 40px; color: white;
                z-index: 3; text-align: center;
            }}
            .slide-content h1 {{
                font-family: 'Fraunces', serif;
                font-size: 2.8rem; font-weight: 700; margin-bottom: 1.2rem;
                letter-spacing: -0.5px; text-shadow: 0 2px 5px rgba(0,0,0,0.4);
                line-height: 1.3;
            }}
            .slide-content p {{
                font-size: 1.15rem; margin-bottom: 0.8rem;
                line-height: 1.5; text-shadow: 0 1px 3px rgba(0,0,0,0.3);
            }}
            .slide-content .highlight {{
                font-size: clamp(0.85rem, 1.1vw, 1.2rem); font-weight: 700;
                color: #f59e0b; margin-top: 1.2rem; display: inline-block;
                background: rgba(0,0,0,0.25); padding: 6px 18px; border-radius: 40px;
                backdrop-filter: blur(4px); white-space: nowrap;
            }}
            .slider-progress {{
                position: absolute; bottom: 0; left: 0; height: 4px;
                background: #f59e0b; width: 0%; z-index: 10; transition: width 0.1s linear;
            }}
            .slider-nav {{
                position: absolute; bottom: 20px; left: 50%;
                transform: translateX(-50%); display: flex; gap: 15px; z-index: 10;
            }}
            .slider-dot {{
                width: 12px; height: 12px; border-radius: 50%;
                background: rgba(255,255,255,0.5); cursor: pointer; transition: all 0.3s;
            }}
            .slider-dot.active {{ background: #f59e0b; width: 30px; border-radius: 10px; }}
            .slider-arrow {{
                position: absolute; top: 50%; transform: translateY(-50%); z-index: 10;
                background: rgba(255,255,255,0.15); border: 2px solid rgba(255,255,255,0.4);
                color: white; width: 45px; height: 45px; border-radius: 50%;
                display: flex; align-items: center; justify-content: center;
                cursor: pointer; font-size: 1.2rem; transition: all 0.3s; backdrop-filter: blur(4px);
            }}
            .slider-arrow:hover {{ background: #f59e0b; border-color: #f59e0b; color: #1e293b; }}
            .slider-arrow.prev {{ left: 20px; }}
            .slider-arrow.next {{ right: 20px; }}

            /* ===== REVEAL ===== */
            .reveal {{ opacity: 0; transform: translateY(30px); transition: opacity 0.7s ease, transform 0.7s ease; }}
            .reveal.visible {{ opacity: 1; transform: translateY(0); }}

            /* ===== STATS ===== */
            .stats-section {{ padding: 60px 30px; background: #0f3b5c; color: white; }}
            .stats-grid {{
                display: flex; justify-content: space-between;
                max-width: 1200px; margin: 0 auto; gap: 0; flex-wrap: nowrap;
            }}
            .stat-card {{
                text-align: center; flex: 1; min-width: 0; padding: 28px 12px;
                border-right: 1px solid rgba(255,255,255,0.2); transition: transform 0.3s;
            }}
            .stat-card:hover {{ transform: translateY(-5px); }}
            .stat-card:last-child {{ border-right: none; }}
            .stat-number {{
                font-family: 'Fraunces', serif;
                font-size: clamp(1.4rem, 2.2vw, 2.4rem); font-weight: 700;
                color: #f59e0b; margin-bottom: 8px; white-space: nowrap; line-height: 1.2;
            }}
            .stat-label {{
                font-size: clamp(0.7rem, 1vw, 0.85rem); text-transform: uppercase;
                letter-spacing: 1px; font-weight: 500; white-space: nowrap;
            }}

            /* ===== SECTION COMMON ===== */
            .about-section {{ padding: 80px 30px; background: #f8fafc; }}
            .about-grid {{
                display: grid; grid-template-columns: 1fr 1fr;
                gap: 60px; max-width: 1280px; margin: 0 auto;
            }}
            .about-tag {{
                color: #f59e0b; font-weight: 700; letter-spacing: 2px;
                margin-bottom: 1rem; font-size: 0.8rem; text-transform: uppercase;
            }}
            .about-title {{
                font-family: 'Fraunces', serif;
                font-size: 2.5rem; font-weight: 700; color: #0f3b5c;
                margin-bottom: 1.5rem; line-height: 1.2;
            }}
            .about-title em {{ font-style: italic; color: #f59e0b; }}
            .about-text {{
                color: #475569; line-height: 1.7; margin-bottom: 1.5rem; font-size: 1rem;
            }}
            .about-highlight {{
                background: white; padding: 20px; border-radius: 16px;
                border-left: 4px solid #f59e0b;
            }}
            .about-img {{
                width: 100%; border-radius: 24px;
                box-shadow: 0 20px 30px -15px rgba(0,0,0,0.15);
            }}

            .services-section {{ padding: 80px 30px; background: white; }}
            .section-header {{ text-align: center; margin-bottom: 50px; }}
            .section-header h2 {{
                font-family: 'Fraunces', serif;
                font-size: 2.2rem; color: #0f3b5c;
            }}
            .section-header h2 em {{ font-style: italic; color: #f59e0b; }}
            .section-header p {{
                color: #64748b; font-size: 1rem; margin-top: 10px; max-width: 60ch;
                margin-left: auto; margin-right: auto;
            }}
            .services-grid {{
                display: grid; grid-template-columns: repeat(4, 1fr);
                gap: 30px; max-width: 1280px; margin: 0 auto;
            }}
            .service-card {{
                background: white; padding: 30px 20px; border-radius: 20px;
                text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.05);
                border: 1px solid #e2e8f0; transition: all 0.3s;
            }}
            .service-card:hover {{ transform: translateY(-8px); border-color: #f59e0b; }}
            .service-icon {{
                font-size: 3rem; color: #f59e0b; margin-bottom: 20px;
            }}
            .service-card h3 {{
                color: #0f3b5c; font-size: 1.1rem; margin-bottom: 8px; font-weight: 600;
            }}
            .service-card p {{ color: #64748b; font-size: 0.9rem; line-height: 1.5; }}

            /* ===== VALUES (Giá trị cốt lõi) ===== */
            .values-section {{
                padding: 80px 30px;
                background: linear-gradient(135deg, #0f3b5c 0%, #1e4a76 100%);
                color: white;
            }}
            .values-section .section-header h2 {{ color: white; }}
            .values-section .section-header p {{ color: rgba(255,255,255,0.75); }}
            .values-grid {{
                display: grid; grid-template-columns: repeat(4, 1fr);
                gap: 24px; max-width: 1280px; margin: 0 auto;
            }}
            .value-card {{
                background: rgba(255,255,255,0.08); padding: 30px 22px;
                border-radius: 20px; text-align: center;
                border: 1px solid rgba(255,255,255,0.15);
                transition: all 0.3s; backdrop-filter: blur(10px);
            }}
            .value-card:hover {{
                background: rgba(255,255,255,0.15);
                transform: translateY(-6px);
                border-color: #f59e0b;
            }}
            .value-icon {{
                font-size: 2.8rem; margin-bottom: 16px;
            }}
            .value-card h3 {{
                color: #f59e0b; font-size: 1.15rem; margin-bottom: 10px; font-weight: 700;
            }}
            .value-card p {{
                color: rgba(255,255,255,0.85); font-size: 0.9rem; line-height: 1.6;
            }}

            /* ===== MISSION / VISION ===== */
            .mission-section {{
                padding: 80px 30px; background: #f8fafc;
            }}
            .mission-grid {{
                display: grid; grid-template-columns: 1fr 1fr;
                gap: 40px; max-width: 1280px; margin: 0 auto;
            }}
            .mission-box {{
                background: white; padding: 40px 32px; border-radius: 20px;
                box-shadow: 0 8px 25px rgba(0,0,0,0.06);
                border-top: 4px solid #f59e0b;
                transition: all 0.3s;
            }}
            .mission-box:hover {{
                transform: translateY(-4px);
                box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            }}
            .mission-box .icon {{
                font-size: 2.8rem; margin-bottom: 16px;
            }}
            .mission-box h3 {{
                font-family: 'Fraunces', serif;
                color: #0f3b5c; font-size: 1.5rem; margin-bottom: 14px; font-weight: 700;
            }}
            .mission-box p {{
                color: #475569; line-height: 1.7; font-size: 0.98rem;
            }}
            .mission-box ul {{
                color: #475569; line-height: 1.8; font-size: 0.95rem;
                padding-left: 20px; margin-top: 12px;
            }}
            .mission-box ul li {{ margin-bottom: 6px; }}
            .mission-box ul li b {{ color: #0f3b5c; }}

            /* ===== FACES (Gương mặt Hòn La) ===== */
            .faces-section {{ padding: 80px 30px; background: white; }}
            .faces-grid {{
                display: grid; grid-template-columns: repeat(3, 1fr);
                gap: 30px; max-width: 1100px; margin: 0 auto;
            }}
            .face-card {{
                text-align: center; padding: 30px 20px; border-radius: 20px;
                background: #f8fafc; border: 1px solid #e2e8f0;
                transition: all 0.3s;
            }}
            .face-card:hover {{
                transform: translateY(-6px);
                border-color: #f59e0b;
                box-shadow: 0 12px 30px rgba(0,0,0,0.08);
            }}
            .face-avatar {{
                width: 120px; height: 120px; border-radius: 50%;
                margin: 0 auto 16px; overflow: hidden;
                border: 4px solid #f59e0b;
                box-shadow: 0 8px 20px rgba(0,0,0,0.1);
                display: flex; align-items: center; justify-content: center;
                background: linear-gradient(135deg, #0f3b5c, #1e4a76);
                color: white; font-family: 'Fraunces', serif;
                font-size: 2.5rem; font-weight: 700;
            }}
            .face-card h4 {{
                color: #0f3b5c; font-size: 1.1rem; margin-bottom: 6px; font-weight: 700;
            }}
            .face-card .role {{
                color: #f59e0b; font-size: 0.85rem; font-weight: 600;
                margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px;
            }}
            .face-card .quote {{
                color: #64748b; font-size: 0.9rem; font-style: italic; line-height: 1.6;
                position: relative; padding: 0 10px;
            }}
            .face-card .quote::before {{ content: '"'; color: #f59e0b; font-size: 1.5rem; font-weight: 700; }}

            /* ===== CAREERS ===== */
            .careers-section {{
                padding: 80px 30px;
                background: linear-gradient(135deg, #0f3b5c 0%, #1e4a76 100%);
                color: white; text-align: center;
            }}
            .careers-section h2 {{
                font-family: 'Fraunces', serif;
                font-size: 2.2rem; margin-bottom: 16px;
            }}
            .careers-section p {{ font-size: 1.05rem; opacity: 0.9; margin-bottom: 10px; }}
            .btn-white {{
                background: #f59e0b; color: #0f3b5c; padding: 14px 40px;
                border-radius: 40px; font-weight: 700;
                display: inline-block; margin-top: 24px;
                text-decoration: none; cursor: pointer;
                transition: all 0.3s;
                box-shadow: 0 8px 20px rgba(245, 158, 11, 0.3);
            }}
            .btn-white:hover {{
                transform: translateY(-2px);
                box-shadow: 0 12px 28px rgba(245, 158, 11, 0.4);
            }}

            /* ===== FOOTER ===== */
            .footer {{
                background: #0f172a; color: #cbd5e1;
                padding: 50px 30px 30px; width: 100%; clear: both;
            }}
            .footer-grid {{
                display: grid; grid-template-columns: repeat(4, 1fr);
                gap: 40px; max-width: 1280px; margin: 0 auto;
                padding-bottom: 40px; border-bottom: 1px solid #334155;
            }}
            .footer-col h4 {{ color: white; margin-bottom: 20px; font-size: 1.1rem; }}
            .footer-col p, .footer-col a {{
                color: #94a3b8; text-decoration: none; line-height: 1.8;
                font-size: 0.9rem; display: block; cursor: pointer;
            }}
            .footer-col a:hover {{ color: #f59e0b; }}
            .copyright {{
                text-align: center; padding-top: 30px;
                font-size: 0.8rem; color: #64748b;
            }}

            /* ===== MODAL ===== */
            .modal {{
                display: none; position: fixed; top: 0; left: 0;
                width: 100vw; height: 100vh; z-index: 99999;
                background: rgba(0,0,0,0.85);
                align-items: flex-start; justify-content: center;
                overflow-y: auto; padding: 80px 20px 20px 20px;
            }}
            .modal.active {{ display: flex; }}
            .modal-content {{
                max-width: 900px; width: 100%; background: white;
                border-radius: 8px; box-shadow: 0 25px 60px rgba(0,0,0,0.4);
                overflow: hidden; animation: modalFadeIn 0.3s ease;
            }}
            @keyframes modalFadeIn {{
                from {{ opacity: 0; transform: scale(0.95); }}
                to {{ opacity: 1; transform: scale(1); }}
            }}
            .modal-header {{
                display: flex; justify-content: flex-end;
                padding: 10px 15px; background: #f0f2f5;
                border-bottom: 1px solid #ddd;
            }}
            .modal-close {{
                background: none; border: none; font-size: 24px;
                cursor: pointer; color: #666; transition: color 0.2s;
            }}
            .modal-close:hover {{ color: #f59e0b; }}
            .modal-body {{
                padding: 30px 40px; max-height: 80vh; overflow-y: auto;
            }}
            .a4-chairman {{
                display: flex; align-items: center; justify-content: space-between;
                gap: 20px; margin-bottom: 30px; padding: 20px;
                background: #f8f9fa; border-radius: 16px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            }}
            .a4-chairman-info {{ flex: 2; }}
            .a4-chairman-avatar {{
                flex: 0 0 auto; width: 150px; height: 150px;
            }}
            .a4-chairman-avatar img {{
                width: 100%; height: 100%; border-radius: 50%;
                object-fit: cover; object-position: center 15%;
                border: 3px solid #f59e0b;
                box-shadow: 0 4px 15px rgba(0,0,0,0.15);
            }}
            .a4-chairman-logo {{
                flex: 0 0 auto; width: 80px; height: 80px;
            }}
            .a4-chairman-logo img {{
                width: 100%; height: 100%; border-radius: 50%;
                object-fit: cover; border: 2px solid #ddd;
                background: white; padding: 5px;
            }}
            .a4-chairman-info h2 {{
                font-size: 1.3rem; color: #0f3b5c; margin-bottom: 5px;
            }}
            .a4-chairman-info .title {{
                color: #f59e0b; font-weight: 600; margin-bottom: 8px;
            }}
            .a4-chairman-info .company {{
                font-size: 0.8rem; color: #666; line-height: 1.4; font-weight: 500;
            }}
            .a4-body {{ line-height: 1.7; color: #333; }}
            .a4-body p {{ text-align: justify; text-justify: inter-ideograph; margin-bottom: 12px; }}
            .a4-date {{
                text-align: right !important; font-style: italic;
                margin-bottom: 20px; color: #666;
            }}
            .vision-box, .mission-box-modal {{
                background: #f0f7ff; padding: 20px; border-radius: 12px;
                margin: 20px 0; border-left: 4px solid #f59e0b;
            }}
            .vision-box h3, .mission-box-modal h3 {{
                color: #0f3b5c; margin-bottom: 12px; font-size: 1.1rem;
            }}
            .a4-signature-left {{ margin-top: 40px; text-align: left; }}
            .sig-block-left {{ display: inline-block; text-align: center; }}
            .sig-block-left .sig-image img {{ max-width: 386px; height: auto; }}
            .a4-footer {{
                margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd;
                text-align: center; font-size: 0.75rem; color: #999;
                display: flex; justify-content: space-between; flex-wrap: wrap;
            }}

            /* ===== RESPONSIVE ===== */
            @media (max-width: 768px) {{
                .slide-layout {{ flex-direction: column; }}
                .slide-image {{ width: 100%; height: 35%; flex: none; }}
                .slide-content {{ padding: 25px 20px; }}
                .slide-content h1 {{ font-size: 1.6rem; }}
                .slide-content p {{ font-size: 0.9rem; }}
                .hero-slider {{ height: auto; min-height: 480px; }}
                .stats-grid, .about-grid, .services-grid, .values-grid, .mission-grid, .faces-grid, .footer-grid {{
                    grid-template-columns: 1fr;
                }}
                .stat-card {{ border-right: none; border-bottom: 1px solid rgba(255,255,255,0.2); }}
                .nav-links {{ display: none; }}
                .logo-circle {{ width: 60px; height: 60px; }}
                .a4-chairman {{ flex-wrap: wrap; justify-content: center; text-align: center; }}
                .modal-body {{ padding: 20px; }}
                .mobile-lang {{ display: flex !important; }}
                .nav-links .lang-switch {{ display: none; }}
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
                <a class="nav-link" data-section="home">Trang chủ</a>
                <div class="dropdown" style="position:relative;">
                    <a class="nav-link" data-section="about">Về chúng ta <i class="fas fa-chevron-down"></i></a>
                    <div class="dropdown-content" style="display:none; position:absolute; background:white; min-width:220px; box-shadow:0 8px 16px rgba(0,0,0,0.15); border-radius:8px; padding:8px 0; top:100%; left:0; z-index:1;">
                        <a class="nav-link" data-section="values" style="color:#333 !important; padding:10px 18px; display:block; font-size:0.9rem; cursor:pointer;">Giá trị cốt lõi</a>
                        <a class="nav-link" data-section="mission" style="color:#333 !important; padding:10px 18px; display:block; font-size:0.9rem; cursor:pointer;">Sứ mệnh & Tầm nhìn</a>
                        <a class="nav-link" data-section="faces" style="color:#333 !important; padding:10px 18px; display:block; font-size:0.9rem; cursor:pointer;">Gương mặt Hòn La</a>
                        <a href="#" id="thuNgoBtn" style="color:#333 !important; padding:10px 18px; display:block; font-size:0.9rem; cursor:pointer;">Lá thư gửi đại gia đình</a>
                    </div>
                </div>
                <a class="nav-link" data-section="services">Hoạt động</a>
                <a class="nav-link" data-section="careers">Phát triển cùng chúng tôi</a>
                <a class="nav-link" data-section="contact">Liên hệ</a>
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

    <!-- Modal Lá thư -->
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
                        <p class="title">Chủ tịch Hội đồng Quản trị</p>
                        <p class="company">Công ty Cổ phần Cảng Hòn La</p>
                        <p class="company">Khu kinh tế Hòn La, Xã Quảng Đông, Huyện Quảng Trạch, Tỉnh Quảng Bình</p>
                    </div>
                    <div class="a4-chairman-logo">
                        <img src="data:image/png;base64,{logo_base64}" alt="Logo Cảng Hòn La">
                    </div>
                </div>
                <div class="a4-body">
                    <p class="a4-date">Quảng Bình, ngày 21 tháng 3 năm 2025</p>
                    <p style="font-weight: bold; font-size: 1rem;">Thân gửi đại gia đình Hòn La,</p>
                    <p>Hôm nay, khi ngồi viết những dòng này, tôi nhìn lại hành trình chúng ta đã đi cùng nhau. Có những ngày nắng cháy da trên công trường, có những đêm thức trắng bên bản vẽ, có những giây phút căng thẳng trước mỗi cột mốc quan trọng. Nhưng cũng có những buổi chiều cả đội ngồi lại, cười vang vì một câu chuyện nhỏ, và nhìn về phía biển — nơi tương lai của chúng ta đang dần thành hình.</p>
                    <p>Tôi muốn nói với mỗi người rằng: <b>các bạn không chỉ đang làm một công việc. Các bạn đang viết nên lịch sử.</b></p>
                    <div class="vision-box">
                        <h3>🎯 Tầm nhìn chung của chúng ta</h3>
                        <p>Đến năm 2035, Hòn La sẽ là cảng biển quốc tế hiện đại hàng đầu Đông Nam Á — nơi mà mỗi cán bộ nhân viên chúng ta hôm nay có thể tự hào nói với con cháu: <i>"Ông/bà của con đã góp tay xây dựng nên nơi này."</i></p>
                    </div>
                    <p>Chúng ta không đi một mình. Chúng ta có nhau. Có Ban Lãnh đạo luôn lắng nghe, có đồng nghiệp luôn sẵn sàng hỗ trợ, có những con người Hòn La kiên cường và đầy khát vọng.</p>
                    <div class="mission-box-modal">
                        <h3>💡 Sứ mệnh của mỗi người Hòn La</h3>
                        <p>Mỗi người trong chúng ta — dù là kỹ sư trên cầu tàu, nhân viên văn phòng, công nhân vận hành hay đội ngũ hậu cần — đều đang giữ một mảnh ghép quan trọng của Hòn La. Chúng ta làm việc với trái tim của người kiến tạo, bởi di sản chúng ta để lại không chỉ là những cầu bến vững chắc, mà còn là hình ảnh một tập thể đoàn kết, chuyên nghiệp và nhân văn.</p>
                    </div>
                    <p>Phía trước còn nhiều thử thách. Nhưng tôi tin vào chúng ta. Tin vào trí tuệ tập thể, tin vào khí phách dân tộc, và tin vào khát vọng vươn ra biển lớn của mỗi người Hòn La.</p>
                    <p>Hãy cùng nhau viết tiếp câu chuyện này nhé. Một câu chuyện đáng tự hào — không chỉ cho chúng ta hôm nay, mà cho cả những thế hệ mai sau.</p>
                    <p style="margin-top: 20px;">Thân ái và trân trọng,</p>
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

    <!-- Hero Slider — Khoảnh khắc Hòn La -->
    <section id="home" class="hero-slider">
        <div class="slides-container">
            <div class="slide active">
                <div class="slide-layout">
                    <div class="slide-image" style="background-image: url('{slide1_src}');"></div>
                    <div class="slide-content">
                        <h1>NGÔI NHÀ CHUNG HÒN LA</h1>
                        <p>Nơi mỗi người là một thành viên của đại gia đình</p>
                        <p>Nơi công việc không chỉ là công việc — mà là sự nghiệp và niềm tự hào</p>
                        <div class="highlight">🤝 Cùng nhau, chúng ta lớn hơn</div>
                    </div>
                </div>
            </div>
            <div class="slide">
                <div class="slide-layout">
                    <div class="slide-content">
                        <h1>CHÚNG TA ĐANG VIẾT LỊCH SỬ</h1>
                        <p>Mỗi cây cầu, mỗi bến tàu — đều có dấu tay của chúng ta</p>
                        <p>Từ công trường đến văn phòng — chúng ta là một đội</p>
                        <div class="highlight">🌟 Kiến tạo di sản, không chỉ công trình</div>
                    </div>
                    <div class="slide-image" style="background-image: url('{slide2_src}');"></div>
                </div>
            </div>
            <div class="slide">
                <div class="slide-layout">
                    <div class="slide-image" style="background-image: url('{slide3_src}');"></div>
                    <div class="slide-content">
                        <h1>PHÁT TRIỂN CÙNG NHAU</h1>
                        <p>Hòn La đầu tư vào con người — đào tạo, cơ hội, và một tương lai vững chắc</p>
                        <p>Bạn đồng hành cùng Hòn La — Hòn La đồng hành cùng bạn</p>
                        <div class="highlight">🚀 Tương lai của bạn, tương lai của chúng ta</div>
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

    <!-- Stats — Về con người, không phải về cơ sở hạ tầng -->
    <section id="stats" class="stats-section">
        <div class="stats-grid">
            <div class="stat-card reveal">
                <div class="stat-number">80+</div>
                <div class="stat-label">CBCNV đồng hành</div>
            </div>
            <div class="stat-card reveal">
                <div class="stat-number">15+</div>
                <div class="stat-label">Năm gắn bó (người lâu nhất)</div>
            </div>
            <div class="stat-card reveal">
                <div class="stat-number">1</div>
                <div class="stat-label">Tầm nhìn chung</div>
            </div>
            <div class="stat-card reveal">
                <div class="stat-number">2035</div>
                <div class="stat-label">Cột mốc chúng ta hướng tới</div>
            </div>
        </div>
    </section>

    <!-- About -->
    <section id="about" class="about-section">
        <div class="about-grid">
            <div>
                <div class="about-tag">Chào mừng đến với Hòn La</div>
                <h2 class="about-title">Nơi công việc trở thành <em>sự nghiệp</em></h2>
                <p class="about-text">Hòn La không chỉ là một cảng biển. Đây là nơi mỗi người chúng ta dành phần lớn thời gian của cuộc đời — nơi chúng ta xây dựng sự nghiệp, kết nối bạn bè, và tạo ra những giá trị có ý nghĩa.</p>
                <p class="about-text">Chúng ta có thể đến từ những phòng ban khác nhau, làm những công việc khác nhau. Nhưng tất cả chúng ta đều có một điểm chung: <b>chúng ta là người Hòn La</b>.</p>
                <div class="about-highlight"><i class="fas fa-heart" style="color:#f59e0b"></i> <strong>Chúng ta không chỉ làm việc cùng nhau — chúng ta trưởng thành cùng nhau</strong></div>
            </div>
            <div><img src="https://images.unsplash.com/photo-1521737604893-d14cc237f11d?q=80&w=2084" class="about-img" alt="Đội ngũ Hòn La"></div>
        </div>
    </section>

    <!-- Values (Giá trị cốt lõi) -->
    <section id="values" class="values-section">
        <div class="section-header">
            <h2>Giá trị cốt lõi của chúng ta</h2>
            <p>4 giá trị định hình con người Hòn La — dù bạn ở vị trí nào, làm công việc gì</p>
        </div>
        <div class="values-grid">
            <div class="value-card reveal">
                <div class="value-icon">💛</div>
                <h3>Tận tâm</h3>
                <p>Chúng ta làm việc bằng cả trái tim, không chỉ bằng đôi tay. Mỗi công việc dù lớn hay nhỏ đều được làm với sự trân trọng.</p>
            </div>
            <div class="value-card reveal">
                <div class="value-icon">🎯</div>
                <h3>Chuyên nghiệp</h3>
                <p>Chúng ta giữ chuẩn mực cao trong từng chi tiết — từ cách giao tiếp với đồng nghiệp đến chất lượng công việc bàn giao.</p>
            </div>
            <div class="value-card reveal">
                <div class="value-icon">💡</div>
                <h3>Đổi mới</h3>
                <p>Chúng ta không ngừng học hỏi, cải tiến, và tìm cách làm tốt hơn. Một ý tưởng nhỏ có thể tạo ra thay đổi lớn.</p>
            </div>
            <div class="value-card reveal">
                <div class="value-icon">🤝</div>
                <h3>Gắn kết</h3>
                <p>Chúng ta là một đội. Thành công của một người là thành công của cả tập thể, và khó khăn của một người là khó khăn chung.</p>
            </div>
        </div>
    </section>

    <!-- Mission & Vision -->
    <section id="mission" class="mission-section">
        <div class="section-header">
            <h2>Sứ mệnh & Tầm nhìn</h2>
            <p>Chúng ta biết mình đang đi đâu, và tại sao điều đó quan trọng</p>
        </div>
        <div class="mission-grid">
            <div class="mission-box">
                <div class="icon">🎯</div>
                <h3>Tầm nhìn 2035</h3>
                <p>Trở thành cảng biển quốc tế hiện đại hàng đầu Đông Nam Á trên tuyến hành lang kinh tế Đông–Tây, là nơi kết nối Việt Nam với thế giới.</p>
                <ul>
                    <li><b>Với khách hàng:</b> Đối tác tin cậy, dịch vụ đẳng cấp quốc tế</li>
                    <li><b>Với CBCNV:</b> Nơi làm việc lý tưởng, thu nhập ổn định, cơ hội phát triển</li>
                    <li><b>Với cộng đồng:</b> Động lực phát triển kinh tế miền Trung</li>
                </ul>
            </div>
            <div class="mission-box">
                <div class="icon">💡</div>
                <h3>Sứ mệnh của chúng ta</h3>
                <p>Xây dựng một môi trường làm việc đẳng cấp — nơi năng lực được trọng dụng, sáng tạo được khuyến khích, và mỗi cá nhân đều tự hào khi đặt bàn tay mình vào công trình lịch sử này.</p>
                <ul>
                    <li><b>Phát triển con người:</b> Đào tạo, kèm cặp, tạo cơ hội thăng tiến</li>
                    <li><b>Xây dựng văn hóa:</b> Tôn trọng, minh bạch, hợp tác, sáng tạo</li>
                    <li><b>Lan tỏa giá trị:</b> Mỗi người Hòn La là một đại sứ của sự chuyên nghiệp</li>
                </ul>
            </div>
        </div>
    </section>

    <!-- Faces — Gương mặt Hòn La -->
    <section id="faces" class="faces-section">
        <div class="section-header">
            <h2>Gương mặt Hòn La</h2>
            <p>Những con người đang viết nên câu chuyện của chúng ta mỗi ngày</p>
        </div>
        <div class="faces-grid">
            <div class="face-card reveal">
                <div class="face-avatar">P</div>
                <h4>Anh/chị A</h4>
                <div class="role">Phòng Kỹ thuật</div>
                <p class="quote">"Mỗi ngày đến Hòn La là một ngày tôi thấy mình có ích. Đồng nghiệp như anh em trong nhà."</p>
            </div>
            <div class="face-card reveal">
                <div class="face-avatar">B</div>
                <h4>Anh/chị B</h4>
                <div class="role">Phòng Vận hành</div>
                <p class="quote">"Công việc vất vả nhưng vui. Nhìn cảng lớn lên từng ngày, tôi thấy tự hào."</p>
            </div>
            <div class="face-card reveal">
                <div class="face-avatar">C</div>
                <h4>Anh/chị C</h4>
                <div class="role">Phòng Kinh doanh</div>
                <p class="quote">"Ở đây, tôi được học hỏi và thử thách bản thân mỗi ngày. Hòn La đầu tư vào con người."</p>
            </div>
        </div>
        <div style="text-align: center; margin-top: 32px; color: #64748b; font-size: 0.9rem;">
            <i class="fas fa-quote-left" style="color: #f59e0b;"></i>
            Mỗi người Hòn La là một câu chuyện đáng kể.
        </div>
    </section>

    <!-- Careers -->
    <section id="careers" class="careers-section">
        <h2>Phát triển cùng chúng tôi</h2>
        <p>Bạn đang tìm kiếm một nơi để phát triển sự nghiệp?</p>
        <p>Hòn La đang tìm kiếm những người đồng hành — không chỉ giỏi chuyên môn, mà còn giàu khát vọng.</p>
        <a href="#" class="btn-white" onclick="alert('Liên hệ Phòng Hành chính Nhân sự để biết thêm chi tiết: hr@honlaport.com.vn'); return false;">
            📢 Xem cơ hội nghề nghiệp
        </a>
    </section>

    <!-- Footer -->
    <footer id="contact" class="footer">
        <div class="footer-grid">
            <div class="footer-col">
                <h4 style="font-size:0.95rem; white-space:nowrap;">CÔNG TY CỔ PHẦN CẢNG HÒN LA</h4>
                <p>Khu kinh tế Hòn La, Xã Phú Trạch, Tỉnh Quảng Trị</p>
                <p>📞 0232.xxxx.xxx</p>
                <p>📧 info@honlaport.com.vn</p>
            </div>
            <div class="footer-col">
                <h4>Liên kết nhanh</h4>
                <a class="nav-link" data-section="home">Trang chủ</a>
                <a class="nav-link" data-section="about">Về chúng ta</a>
                <a class="nav-link" data-section="values">Giá trị cốt lõi</a>
                <a class="nav-link" data-section="mission">Sứ mệnh & Tầm nhìn</a>
                <a class="nav-link" data-section="faces">Gương mặt Hòn La</a>
            </div>
            <div class="footer-col">
                <h4>Dành cho CBCNV</h4>
                <a href="#">Sổ tay nhân viên</a>
                <a href="#">Quy trình nội bộ</a>
                <a href="#">Phúc lợi & chính sách</a>
                <a href="#">Gửi phản hồi, góp ý</a>
            </div>
            <div class="footer-col">
                <h4>Kết nối với chúng tôi</h4>
                <p>📅 Thứ 2 - Thứ 7</p>
                <p>🏢 Văn phòng: 7:30 - 17:00</p>
                <p>🚢 Bến cảng: 24/7</p>
                <p style="margin-top: 12px;">
                    <a href="#" style="color:#f59e0b;">📘 Facebook nội bộ</a>
                </p>
            </div>
        </div>
        <div class="copyright">
            <p>© 2026 - Công ty Cổ phần Cảng Hòn La. Bản tin nội bộ dành cho CBCNV.</p>
            <p style="margin-top: 10px;">CHÚNG TA LÀ MỘT ĐỘI · CHÚNG TA LÀ HÒN LA</p>
        </div>
    </footer>

    {landing_js}
    </body>
    </html>
    """

    components.html(landing_html, height=3600, scrolling=False)