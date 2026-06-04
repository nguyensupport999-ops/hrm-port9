'''
Tóm lại bài học rút ra từ ca này: khi HTML nằm trong components.html (iframe), việc giao tiếp với trang Streamlit cha luôn phải dùng window.top thay vì window.parent — đặc biệt trên Streamlit Cloud nơi có thể có nhiều tầng iframe lồng nhau. Và không bao giờ dùng replaceState rồi lại location.href cùng lúc vì chúng triệt tiêu nhau.
Nếu sau này cần thêm tính năng hay gặp bug mới, cứ ping lại nhé!
'''
import streamlit as st
import psycopg2
import psycopg2.extras
from datetime import datetime, date, timedelta
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import tempfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import requests
from openpyxl.styles import Font, Alignment
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import qrcode
from io import BytesIO
import os
import pathlib
import streamlit.components.v1 as components
import urllib.parse

# Xử lý đổi ngôn ngữ từ request
def handle_language_change():
    """Xử lý thay đổi ngôn ngữ từ query params"""
    query_params = st.query_params
    if 'lang' in query_params:
        new_lang = query_params['lang']
        if new_lang in ['vi', 'en']:
            st.session_state.language = new_lang
            st.query_params.clear()
            st.rerun()
    
    # Cũng kiểm tra POST request (cho fetch từ client)
    try:
        # Lấy dữ liệu từ request (nếu có)
        import sys
        if hasattr(st, 'context') and hasattr(st.context, 'headers'):
            content_length = int(st.context.headers.get('content-length', 0))
            if content_length > 0:
                body = sys.stdin.read(content_length) if content_length else ''
                if 'set_language=' in body:
                    new_lang = body.replace('set_language=', '').strip()
                    if new_lang in ['vi', 'en']:
                        st.session_state.language = new_lang
                        st.rerun()
    except:
        pass

# Gọi hàm xử lý ngôn ngữ trước khi hiển thị landing page
handle_language_change()

def show_landing_page():
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
            [data-testid="stSidebar"],
            [data-testid="collapsedControl"],
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
        </style>
    """, unsafe_allow_html=True)
    
    # Đọc file logo
    import base64
    logo_path = os.path.join(os.path.dirname(__file__), "logo_cty.png")
    logo_base64 = ""
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
    chu_tich_img = load_img_b64("333.png")
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
        <title>Cảng Quốc tế Hòn La</title>
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

    # Nút HRM thuần Python
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🔐 HRM - QUẢN LÝ NHÂN SỰ / Chỉ dành cho Nhân viên", width='stretch'):
            st.session_state.show_hrm = True
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

st.set_page_config(page_title="HRM-Port", page_icon="🏗️", layout="wide")

# ========== XỬ LÝ ĐA NGÔN NGỮ ==========
def init_language():
    """Khởi tạo và xử lý chuyển đổi ngôn ngữ"""
    if 'language' not in st.session_state:
        st.session_state.language = 'vi'
    
    # Kiểm tra query params để đổi ngôn ngữ
    query_params = st.query_params
    if 'lang' in query_params:
        new_lang = query_params['lang']
        if new_lang in ['vi', 'en']:
            st.session_state.language = new_lang
            # Xóa param để tránh lặp
            st.query_params.clear()
            st.rerun()

# Gọi hàm khởi tạo
init_language()

# ========== KHỞI TẠO SESSION STATE ==========
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None

# show_hrm=False → hiện landing | show_hrm=True → vào HRM (sidebar login)
if 'show_hrm' not in st.session_state:
    st.session_state.show_hrm = False

# ========== KIỂM TRA URL PARAMS (Từ nút Nhân viên trên Landing Page) ==========
query_params = st.query_params
if query_params.get('goto') == 'hrm':
    st.session_state.show_hrm = True   # Chỉ thoát landing, KHÔNG tự đăng nhập
    st.query_params.clear()
    st.rerun()


# ========== HIỂN THỊ LANDING PAGE NẾU CHƯA VÀO HRM ==========
logo_path = "logo_cty.png"
if os.path.exists(logo_path):
    with st.sidebar:
        st.image(logo_path, width='stretch')
        st.divider()

# Trong phần main hoặc ở cuối file, đảm bảo:
if not st.session_state.logged_in and not st.session_state.get('show_hrm', False):
    # Xử lý đổi ngôn ngữ
    handle_language_change()
    
    # Ẩn sidebar
    st.markdown("""
        <style>
            [data-testid="stSidebar"] { display: none !important; }
            [data-testid="collapsedControl"] { display: none !important; }
            header { display: none !important; }
            footer[data-testid], #stDecoration { display: none !important; }
        </style>
    """, unsafe_allow_html=True)
    show_landing_page()
    st.stop()

# Nếu show_hrm=True hoặc logged_in=True → chạy tiếp HRM, sidebar tự hiện

# ========== PHẦN CODE HRM BẮT ĐẦU TỪ ĐÂY ==========

st.markdown("""
    <style>
        /* ===== ẨN MANAGE APP - dùng mọi selector có thể ===== */
        [data-testid="stToolbar"],
        [data-testid="manage-app-button"],
        [data-testid="stAppDeployButton"],
        .stDeployButton,
        #MainMenu,
        div[class*="toolbar"],
        div[class*="StatusWidget"],
        div[class*="viewerBadge"],
        div[class*="manage-app"],
        button[kind="managedApp"],
        [data-testid="stBottom"] > div:last-child { 
            display: none !important; 
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
        }
        footer[data-testid] { display: none !important; }

        /* ===== PADDING TOP / BOTTOM = 5px ===== */
        .stApp > div[data-testid="stAppViewContainer"] > section[data-testid="stMain"] > div {
            padding-top: 5px !important;
            padding-bottom: 5px !important;
        }
        .block-container {
            padding-top: 5px !important;
            padding-bottom: 5px !important;
        }

        /* ===== LOGO SIDEBAR: hình tròn đổ bóng, sát top, căn giữa ===== */
        [data-testid="stSidebar"] > div:first-child {
            padding-top: 0 !important;
        }
        [data-testid="stSidebar"] [data-testid="stImage"] {
            display: flex !important;
            justify-content: center !important;
            margin-top: 0 !important;
            padding-top: 0 !important;
        }
        [data-testid="stSidebar"] [data-testid="stImage"] img {
            width: 150px !important;
            height: 150px !important;
            border-radius: 50% !important;
            object-fit: cover !important;
            box-shadow: 0 4px 20px rgba(0,0,0,0.25) !important;
            display: block !important;
            margin: 4px auto 0 auto !important;
        }
    </style>
    <script>
        // MutationObserver: ẩn Manage App ngay khi DOM thay đổi
        function hideManageApp() {
            const selectors = [
                '[data-testid="manage-app-button"]',
                '[data-testid="stToolbar"]',
                '[data-testid="stAppDeployButton"]',
                '.stDeployButton',
                'button[kind="managedApp"]'
            ];
            selectors.forEach(sel => {
                document.querySelectorAll(sel).forEach(el => {
                    el.style.cssText = 'display:none!important;visibility:hidden!important';
                });
            });
        }
        hideManageApp();
        // Quan sát DOM liên tục — bắt được dù Streamlit render trễ
        const observer = new MutationObserver(hideManageApp);
        observer.observe(document.body, { childList: true, subtree: true });
    </script>
""", unsafe_allow_html=True)

def to_float_or_none(val):
    """Chuyển đổi giá trị sang float hoặc None, tránh lỗi numeric"""
    if val is None or str(val).strip() == '':
        return None
    try:
        return float(val)
    except:
        return None

def format_date(d):
    if d is None or pd.isna(d): return ''
    try: return d.strftime('%d/%m/%Y') if hasattr(d,'strftime') else str(d)[:10]
    except: return str(d)

def parse_date(s):
    """Chuyển đổi nhiều định dạng ngày tháng về date object"""
    if not s or str(s).strip() == '':
        return None
    
    # Nếu đã là date object
    if hasattr(s, 'strftime'):
        return s
    
    s = str(s).strip()
    
    # Các định dạng cần thử
    formats = [
        '%d/%m/%Y',      # 18/04/2026
        '%d-%m-%Y',      # 18-04-2026
        '%Y-%m-%d',      # 2026-04-18
        '%Y/%m/%d',      # 2026/04/18
        '%d.%m.%Y',      # 18.04.2026
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    
    # Nếu không có định dạng nào phù hợp
    print(f"⚠️ Không thể parse ngày: {s}")
    return None

def get_xung_ho(gioi_tinh, ho_ten=""):
    """
    Lấy cách xưng hô phù hợp dựa trên giới tính
    - Nếu giới tính là Nam -> trả về "Anh"
    - Nếu giới tính là Nữ -> trả về "Chị"
    - Nếu giới tính là None hoặc rỗng -> trả về "Anh/Chị"
    """
    if gioi_tinh == "Nam":
        return "Anh"
    elif gioi_tinh == "Nữ":
        return "Chị"
    else:
        return "Anh/Chị"

def get_loi_chuc_sinh_nhat(ho_ten, gioi_tinh, tuoi=None):
    """
    Tạo lời chúc sinh nhật có xưng hô phù hợp
    """
    xung_ho = get_xung_ho(gioi_tinh, ho_ten)
    
    loi_chuc = f"""
🎉🎂 CHÚC MỪNG SINH NHẬT {xung_ho.upper()} {ho_ten.upper()} 🎂🎉

Thân gửi {xung_ho}: {ho_ten},

Nhân dịp sinh nhật của {xung_ho}, thay mặt Ban Lãnh đạo Công ty CP Cảng Hòn La, 
xin gửi đến {xung_ho} những lời chúc tốt đẹp nhất.

Chúc {xung_ho} luôn mạnh khỏe, hạnh phúc và thành công trong công việc 
cũng như trong cuộc sống.

"""
    
    if tuoi:
        loi_chuc += f"Chúc mừng {xung_ho} tròn {tuoi} tuổi! 🎂\n\n"
    
    loi_chuc += f"""
Cảm ơn {xung_ho} đã luôn đồng hành và đóng góp cho sự phát triển của Công ty.

Trân trọng!

🏗️ CÔNG TY CP CẢNG HÒN LA
    """
    
    return loi_chuc

def auto_check_birthday():
    if 'sinh_nhat_hom_nay_list' not in st.session_state:
        st.session_state.sinh_nhat_hom_nay_list = []

    today_str = date.today().strftime('%Y-%m-%d')
    
    # Key mới: gắn với cả ngày lẫn trạng thái login
    # → mỗi lần login lại đều query DB, không bị cache nhầm
    check_key = f"{today_str}_{st.session_state.get('username', 'guest')}"
    
    if st.session_state.get('last_birthday_check') == check_key:
        return  # Đã check cho user này hôm nay rồi
    
    try:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, dien_thoai, email_lien_he
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_sinh IS NOT NULL
            AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
            AND EXTRACT(DAY FROM ngay_sinh) = EXTRACT(DAY FROM CURRENT_DATE)
        """)
        birthday_today = c.fetchall()
        db.close()
        
        st.session_state.sinh_nhat_hom_nay_list = [
            {
                'ho_ten': nv['ho_ten'],
                'ma_nv': nv['ma_nv'],
                'xung_ho': get_xung_ho(nv.get('gioi_tinh'), nv['ho_ten'])
            }
            for nv in birthday_today
        ]
        
        for nv in birthday_today:
            xung_ho = get_xung_ho(nv.get('gioi_tinh'), nv['ho_ten'])
            st.toast(f"🎂 Sinh nhật {xung_ho} {nv['ho_ten']} hôm nay!", icon="🎂")
        
        # Đánh dấu đã check — dùng check_key gắn với username
        st.session_state.last_birthday_check = check_key
        
    except Exception as e:
        st.warning(f"⚠️ Không thể kiểm tra sinh nhật: {e}")

# Import config - ưu tiên config.py (local), fallback to config_template (cloud)
try:
    from config import COMPANY_CONFIG, BHXH_CONFIG, EMAIL_CONFIG, TELEGRAM_CONFIG, USERS
    print("Using local config.py")
except ImportError:
    from config_template import COMPANY_CONFIG, BHXH_CONFIG, EMAIL_CONFIG, TELEGRAM_CONFIG, USERS
    print("Using config_template.py")
    

def da_chuyen_doi_chinh_thuc(nv_id):
    """Kiểm tra xem nhân viên đã có quyết định chuyển từ thử việc sang chính thức chưa"""
    try:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT * FROM quyet_dinh_nhan_su 
            WHERE nhan_vien_id = %s AND loai_quyet_dinh = 'CHINH_THUC'
            ORDER BY ngay_quyet_dinh DESC LIMIT 1
        """, (nv_id,))
        result = c.fetchone()
        db.close()
        
        # Debug: in ra log để kiểm tra
        print(f"Checking nv_id={nv_id}, found={result is not None}")
        if result:
            print(f"Quyet dinh: {result}")
        
        return result is not None, result
    except Exception as e:
        print(f"Error in da_chuyen_doi_chinh_thuc: {e}")
        return False, None

def lay_thong_tin_truoc_chuyen_doi(nv_id):
    """Lấy thông tin nhân viên trước khi chuyển đổi (từ lich_su_cong_tac)"""
    try:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        # Lấy lịch sử công tác cũ nhất (thời gian thử việc)
        c.execute("""
            SELECT * FROM lich_su_cong_tac 
            WHERE nhan_vien_id = %s 
            ORDER BY tu_ngay ASC LIMIT 1
        """, (nv_id,))
        result = c.fetchone()
        db.close()
        return result
    except:
        return None
# ========== DATABASE CONNECTION (SUPABASE) ==========
def get_connection():
    # Đọc từ st.secrets (không có DB_)
    if 'connections' in st.secrets and 'supabase' in st.secrets.connections:
        return psycopg2.connect(
            host=st.secrets.connections.supabase.host,  # không có DB_
            port=st.secrets.connections.supabase.port,
            user=st.secrets.connections.supabase.user,
            password=st.secrets.connections.supabase.password,
            database=st.secrets.connections.supabase.database
        )
    
    # Fallback: đọc từ .env (có DB_)
    from dotenv import load_dotenv
    load_dotenv()
    return psycopg2.connect(
        host=os.getenv('DB_HOST'),      # có DB_
        port=os.getenv('DB_PORT'),
        user=os.getenv('DB_USER'),
        password=os.getenv('DB_PASSWORD'),
        database=os.getenv('DB_NAME')
    )

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None
if 'selected_nv_id' not in st.session_state:
    st.session_state.selected_nv_id = None
if 'edit_uv_id' not in st.session_state:
    st.session_state.edit_uv_id = None
if 'bhxh_family_nv_id' not in st.session_state:
    st.session_state.bhxh_family_nv_id = None
if 'bhxh_family_nv_name' not in st.session_state:
    st.session_state.bhxh_family_nv_name = None
if 'bhxh_family_members' not in st.session_state:
    st.session_state.bhxh_family_members = []
if 'show_chuyen_nv_form' not in st.session_state:
    st.session_state.show_chuyen_nv_form = False
if 'chuyen_uv_id' not in st.session_state:
    st.session_state.chuyen_uv_id = None
if 'chuyen_uv_data' not in st.session_state:
    st.session_state.chuyen_uv_data = {}

def force_center(p):
    pPr = p._p.get_or_add_pPr()
    for jc in pPr.findall(qn('w:jc')):
        pPr.remove(jc)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)

st.markdown("""
<style>
    [data-testid="stDataFrame"] > div {
        overflow-x: auto !important;
    }
    [data-testid="stDataFrame"] table {
        min-width: 2000px !important;
        width: max-content !important;
    }
</style>
""", unsafe_allow_html=True)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def to_int_or_none(val):
    """Chuyển đổi giá trị sang int hoặc None"""
    if val is None or str(val).strip() == '':
        return None
    try:
        return int(float(val))
    except:
        return None
        
def tao_noi_dung_zalo(nv):
    ZC = COMPANY_CONFIG
    return f"""Gửi anh/chị: {nv.get('ho_ten','')},

Thông tin đã cập nhật:
- Họ tên: {nv.get('ho_ten','')}
- Ngày sinh: {format_date(nv.get('ngay_sinh'))}
- CCCD: {nv.get('so_cccd','')}
- Ngày cấp: {format_date(nv.get('ngay_cap_cccd'))}
- Thường trú: {nv.get('thuong_tru','')}
- Số BHXH: {nv.get('ma_so_bhxh','')}
- TK NH: {nv.get('so_tai_khoan_nh','')}
- CN NH: {nv.get('chi_nhanh_nh','')}

{ZC.get('loi_nhan_zalo','Vui lòng kiểm tra và phản hồi nếu có sai sót. Xin Cảm ơn!')}"""

def remove_table_border(tbl):
    for row in tbl.rows:
        for cell in row.cells:
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            b = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders')
            if b is not None: tcPr.remove(b)

# ========== CÁC HÀM TẠO HỢP ĐỒNG (GIỮ NGUYÊN) ==========
def tao_hop_dong(nv):
    # ... giữ nguyên code cũ (quá dài, tôi giữ lại)
    CC = COMPANY_CONFIG; doc = Document()
    s = doc.styles['Normal']; s.font.name='Times New Roman'; s.font.size=Pt(13)
    s.paragraph_format.space_after=Pt(0); s.paragraph_format.space_before=Pt(0)
    sec = doc.sections[0]; sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(3.5); sec.right_margin=Cm(2)
    def add_p(text='', bold=False, size=Pt(13)):
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        if bold and p.runs:
            p.runs[0].bold = True
        if p.runs:
            p.runs[0].font.size = size
        return p
    def al(label,value):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(5))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r=p.add_run(f'{label}'); r.font.size=Pt(13)
        r=p.add_run('\t: '); r.font.size=Pt(13)
        r=p.add_run(f'{value}'); r.font.size=Pt(13)
    ht=doc.add_table(rows=4,cols=2); ht.alignment=WD_TABLE_ALIGNMENT.CENTER; ht.autofit=False; remove_table_border(ht)
    for row in ht.rows: row.cells[0].width=Cm(6); row.cells[1].width=Cm(11)
    c=ht.rows[0].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CÔNG TY CỔ PHẦN'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[0].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CẢNG HÒN LA'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Độc lập - Tự do - Hạnh phúc'); r.bold=True; r.italic=True; r.font.size=Pt(13)
    c=ht.rows[2].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*12); r.font.size=Pt(9)
    c=ht.rows[2].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*20); r.font.size=Pt(9)
    c=ht.rows[3].cells[1]; p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(20)
    ngay_ky = nv.get("ngay_ky_hd")
    ngay_vao = nv.get("ngay_vao_lam")
    nk = ngay_ky if ngay_ky else ngay_vao
    ns = 'Quảng Trị, ngày ... tháng ... năm ......'
    if nk and hasattr(nk, 'day'):
        ns = f'Quảng Trị, ngày {nk.day} tháng {nk.month:02d} năm {nk.year}'
    run = p.add_run(ns)
    run.font.size = Pt(13)
    run.italic = True
    c=ht.rows[3].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f'Số: {nv.get("so_hdld","...")}'); r.italic=True; r.font.size=Pt(12)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run('HỢP ĐỒNG LAO ĐỘNG')
    r.bold = True
    r.font.size = Pt(18)
    force_center(p)
    p2 = doc.add_paragraph('- Căn cứ thông tư 10/2020/TT-LĐTBXH ngày 12/11/2020 hướng dẫn thi hành một số điều của Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019 về nội dung của hợp đồng lao động;')
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p2 = doc.add_paragraph('- Căn cứ nhu cầu sử dụng lao động trong đơn vị.')
    doc.add_paragraph('Chúng tôi gồm:')
    p=doc.add_paragraph(); r=p.add_run(f'BÊN A: {CC["ten_cong_ty"]} (Người sử dụng LĐ)'); r.bold=True
    al('Đại diện',f"Ông {CC['dai_dien']}"); al('Chức vụ',CC['chuc_vu']); al('Mã số thuế',CC['ma_so_thue'])
    al('Điện thoại',CC['dien_thoai_cty']); al('Địa chỉ',CC['dia_chi']); doc.add_paragraph()
    p=doc.add_paragraph(); r=p.add_run('BÊN B: (Người lao động)'); r.bold=True
    sk=nv.get('so_tai_khoan_nh','')
    if nv.get('chi_nhanh_nh'): sk+=f' - {nv.get("chi_nhanh_nh")}'
    gt = nv.get('gioi_tinh','')
    xung_ho = 'Ông' if gt == 'Nam' else ('Bà' if gt == 'Nữ' else 'Ông/Bà')
    al(xung_ho, nv.get('ho_ten',''))
    al('Ngày sinh',format_date(nv.get('ngay_sinh')))
    al('Số CMND/CCCD',nv.get('so_cccd','')); al('Ngày cấp',format_date(nv.get('ngay_cap_cccd')))
    al('Nơi cấp',nv.get('noi_cap_cccd','')); al('Số TKNH',sk)
    al('Điện thoại',nv.get('dien_thoai','')); al('Thường trú',nv.get('thuong_tru',''))
    doc.add_paragraph('Thoả thuận ký kết Hợp đồng lao động với những điều khoản dưới đây:')
    p=doc.add_paragraph(); r=p.add_run('Điều 1. Thời hạn và công việc hợp đồng:'); r.bold=True
    ngay_hieu_luc = nv.get("ngay_ky_hd") or nv.get("ngay_vao_lam")
    ns2 = '.../.../..........'
    if ngay_hieu_luc and hasattr(ngay_hieu_luc, 'day'):
        ns2 = f'{ngay_hieu_luc.day} tháng {ngay_hieu_luc.month:02d} năm {ngay_hieu_luc.year}'
    elif ngay_hieu_luc:
        ns2 = str(ngay_hieu_luc)
    add_p(f'-    Bên B làm việc theo chế độ hợp đồng lao động không xác định thời hạn;')
    add_p(f'-    Thời gian: Từ ngày {ns2};')
    add_p('-    Địa điểm làm việc: Tại Cảng tổng hợp quốc tế Hòn La và các địa điểm khác theo sự sắp xếp của Công ty;')
    add_p(f'-    Vị trí: {nv.get("chuc_danh_nghe","")};')
    add_p('-    Công việc phải làm: Thực hiện công việc theo đúng chuyên môn dưới sự quản lý, điều hành của cấp trên;')
    add_p('-    Mức lương và phụ cấp: Theo thỏa thuận;')
    add_p('-    Hình thức trả lương: Tiền mặt hoặc chuyển khoản, theo lần chi trả;')
    add_p('-    Kỳ hạn trả lương: Theo quy định Công ty;')
    add_p('-    Chế độ nâng lương: Theo thỏa thuận.')
    p=doc.add_paragraph(); r=p.add_run('Điều 2. Chế độ làm việc:'); r.bold=True
    add_p('-    Thời gian làm việc: Theo tính chất công việc, do nhu cầu kinh doanh của Công ty nên thời gian làm việc của bên B là linh hoạt nhưng phải đảm bảo hoàn thành công việc được giao;')
    add_p('-    Thời gian nghỉ ngơi của người lao động: Theo thỏa thuận và phù hợp với quy định của pháp luật;')
    add_p('-    Ngoài giờ làm việc: Người lao động phải tự chịu trách nhiệm về các hoạt động cá nhân của mình.')
    p=doc.add_paragraph(); r=p.add_run('Điều 3. Nghĩa vụ, quyền lợi NLĐ:'); r.bold=True
    p=doc.add_paragraph(); r=p.add_run('1. Nghĩa vụ:'); r.bold=True
    add_p('-    Hoàn thành những công việc được giao và sẵn sàng chấp nhận mọi sự điều động khi có yêu cầu;')
    add_p('-    Chấp hành nghiêm túc nội quy, kỷ luật lao động, an toàn lao động và các quy định của Công ty và pháp luật của Nhà nước;')
    add_p('-    Người lao động có trách nhiệm tuân thủ đầy đủ quy định về an toàn lao động, quy trình vận hành thiết bị và hướng dẫn của Công ty. Trường hợp NLĐ cố ý vi phạm hoặc vi phạm nghiêm trọng quy định an toàn lao động gây thiệt hại thì phải chịu trách nhiệm theo quy định pháp luật và nội quy Công ty;')
    add_p('-    Bồi thường vi phạm vật chất : Phải bồi thường vật chất do cá nhân vi phạm quy định của Công ty về bảo quản trang thiết bị được giao.')
    p=doc.add_paragraph(); r=p.add_run('2. Quyền Lợi:'); r.bold=True
    add_p('-    Phương tiện đi lại: Tự túc;')
    add_p('-    Được Công ty đóng Bảo hiểm xã hội, bảo hiểm y tế, BHTN: theo chế độ hiện hành của Nhà nước và Quy định của Công ty;')
    add_p('-    Được Công ty cấp đầy đủ bảo hộ lao động theo đúng vị trí làm việc;')
    add_p('-    Được phân công công việc theo yêu cầu của Công ty phù hợp với khả năng và trình độ chuyên môn mà người lao động đáp ứng;')
    add_p('-    Các quyền lợi khác thực hiện theo quy định của Pháp luật Lao động như tạm dừng, chấm dứt hợp đồng.')
    p=doc.add_paragraph(); r=p.add_run('Điều 4. Nghĩa vụ, quyền hạn NSDLĐ:'); r.bold=True
    add_p('-    Bảo đảm việc làm và thực hiện đầy đủ những điều đã cam kết trong hợp đồng;')
    add_p('-    Thanh toán đầy đủ, đúng hạn các chế độ và quyền lợi cho người lao động theo hợp đồng;')
    add_p('-    Điều hành người lao động hoàn thành công việc theo hợp đồng;')
    add_p('-    Tạm hoãn, chấm dứt hợp đồng, kỷ luật người lao động theo quy định của pháp luật, và nội quy lao động của Công ty.')
    p=doc.add_paragraph(); r=p.add_run('Điều 5. Điều khoản chung:'); r.bold=True
    add_p('-    Những nội dung về quan hệ lao động không ghi trong hợp đồng này thì được áp dụng theo pháp luật lao động;')
    add_p('-    Những thoả thuận khác (nếu có): không;')
    add_p('-    Hợp đồng này có hiệu lực từ ngày ký và được làm thành 02 bản, Bên A giữ 01 bản, Bên B giữ 01 có giá trị pháp lý như nhau, để làm căn cứ thực hiện.')
    add_p('Bản HĐ này lập tại văn phòng Công ty CP Cảng Hòn La.'); doc.add_paragraph()
    ts=doc.add_table(rows=3,cols=2); ts.alignment=WD_TABLE_ALIGNMENT.CENTER; remove_table_border(ts)
    c=ts.rows[0].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[0].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI SỬ DỤNG LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[1].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[1].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[2].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(nv.get('ho_ten','').upper()); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[2].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(CC['dai_dien'].upper()); r.bold=True; r.font.size=Pt(13)
    tf=tempfile.NamedTemporaryFile(delete=False,suffix='.docx'); doc.save(tf.name); return tf.name

def tao_hop_dong_thu_viec(nv):
    # ... giữ nguyên code cũ (tương tự)
    CC = COMPANY_CONFIG; doc = Document()
    s = doc.styles['Normal']; s.font.name='Times New Roman'; s.font.size=Pt(13)
    s.paragraph_format.space_after=Pt(0); s.paragraph_format.space_before=Pt(0)
    s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sec = doc.sections[0]; sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(3.5); sec.right_margin=Cm(2)
    def al(label,value):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(5))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r=p.add_run(f'{label}'); r.font.size=Pt(13)
        r=p.add_run('\t: '); r.font.size=Pt(13)
        r=p.add_run(f'{value}'); r.font.size=Pt(13)
    ht=doc.add_table(rows=4,cols=2); ht.alignment=WD_TABLE_ALIGNMENT.CENTER; ht.autofit=False; remove_table_border(ht)
    for row in ht.rows: row.cells[0].width=Cm(6); row.cells[1].width=Cm(11)
    c=ht.rows[0].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CÔNG TY CỔ PHẦN'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[0].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CẢNG HÒN LA'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Độc lập - Tự do - Hạnh phúc'); r.bold=True; r.italic=True; r.font.size=Pt(13)
    c=ht.rows[2].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*12); r.font.size=Pt(9)
    c=ht.rows[2].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*20); r.font.size=Pt(9)
    c=ht.rows[3].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f'Số: {nv.get("so_hdld",".../.../HĐTV-CHL")}'); r.italic=True; r.font.size=Pt(12)
    c=ht.rows[3].cells[1]; p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(20)
    nk=nv.get("ngay_vao_lam") or nv.get("ngay_ky_hd")
    ns='Quảng Trị, ngày ... tháng ... năm ......'
    if nk and hasattr(nk,'day'): ns=f'Quảng Trị, ngày {nk.day} tháng {nk.month:02d} năm {nk.year}'
    run = p.add_run(ns)
    run.font.size = Pt(13)
    run.italic = True
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run('HỢP ĐỒNG THỬ VIỆC')
    r.bold = True
    r.font.size = Pt(18)
    force_center(p)
    p2 = doc.add_paragraph('- Căn cứ thông tư 10/2020/TT-LĐTBXH ngày 12/11/2020 hướng dẫn thi hành một số điều của Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019 về nội dung của hợp đồng lao động;')
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph('- Căn cứ nhu cầu sử dụng lao động trong đơn vị.')
    doc.add_paragraph('Chúng tôi gồm:')
    p=doc.add_paragraph(); r=p.add_run(f'BÊN A: {CC["ten_cong_ty"]} (Người sử dụng LĐ)'); r.bold=True
    al('Đại diện',f"Ông {CC['dai_dien']}"); al('Chức vụ',CC['chuc_vu']); al('Mã số thuế',CC['ma_so_thue'])
    al('Điện thoại',CC['dien_thoai_cty']); al('Địa chỉ',CC['dia_chi'])
    p=doc.add_paragraph(); r=p.add_run('BÊN B: (Người lao động)'); r.bold=True
    sk=nv.get('so_tai_khoan_nh','')
    if nv.get('chi_nhanh_nh'): sk+=f' - {nv.get("chi_nhanh_nh")}'
    gt = nv.get('gioi_tinh','')
    xung_ho = 'Ông' if gt == 'Nam' else ('Bà' if gt == 'Nữ' else 'Ông/Bà')
    al(xung_ho, nv.get('ho_ten',''))
    al('Ngày sinh',format_date(nv.get('ngay_sinh')))
    al('Số CMND/CCCD',nv.get('so_cccd','')); al('Ngày cấp',format_date(nv.get('ngay_cap_cccd')))
    al('Nơi cấp',nv.get('noi_cap_cccd','')); al('Số TKNH',sk)
    al('Điện thoại',nv.get('dien_thoai','')); al('Thường trú',nv.get('thuong_tru',''))
    doc.add_paragraph('Thoả thuận ký kết Hợp đồng Thử việc với những điều khoản dưới đây:')
    p=doc.add_paragraph(); r=p.add_run('Điều 1. Thời hạn và công việc hợp đồng:'); r.bold=True
    ns2='.../.../..........'
    if nk and hasattr(nk,'day'): ns2=f'{nk.day} tháng {nk.month} năm {nk.year}'
    elif nk: ns2=str(nk)
    p=doc.add_paragraph(f'-    Bên B làm việc theo chế độ hợp đồng thử việc, có thời hạn 01 tháng;')
    if nk and hasattr(nk,'day'):
        nkt=nk+timedelta(days=30)
        p=doc.add_paragraph(f'     + Bắt đầu: {nk.day:02d}/{nk.month:02d}/{nk.year}')
        p=doc.add_paragraph(f'     + Kết thúc: {nkt.day:02d}/{nkt.month:02d}/{nkt.year}')
    else: doc.add_paragraph('     + Bắt đầu: .../.../......'); doc.add_paragraph('     + Kết thúc: .../.../......')
    p=doc.add_paragraph('-    Địa điểm làm việc: Tại Cảng tổng hợp quốc tế Hòn La và các địa điểm khác theo sự sắp xếp của Công ty;')
    p=doc.add_paragraph(f'-    Vị trí: {nv.get("chuc_danh_nghe","")};')
    p=doc.add_paragraph('-    Công việc phải làm: Thực hiện công việc theo đúng chuyên môn dưới sự quản lý, điều hành của cấp trên;')
    p=doc.add_paragraph('-    Mức lương và phụ cấp: Theo thỏa thuận;')
    p=doc.add_paragraph('-    Hình thức trả lương: Tiền mặt hoặc chuyển khoản, theo lần chi trả;')
    p=doc.add_paragraph('-    Kỳ hạn trả lương: Theo quy định Công ty;')
    p=doc.add_paragraph(); r=p.add_run('Điều 2. Chế độ làm việc:'); r.bold=True
    p=doc.add_paragraph('-    Thời gian làm việc: Theo tính chất công việc, do nhu cầu kinh doanh của Công ty nên thời gian làm việc của bên B là linh hoạt nhưng phải đảm bảo hoàn thành công việc được giao;')
    p=doc.add_paragraph('-    Thời gian nghỉ ngơi của người lao động: Theo thỏa thuận và phù hợp với quy định của pháp luật;')
    p=doc.add_paragraph('-    Ngoài giờ làm việc: Người lao động phải tự chịu trách nhiệm về các hoạt động cá nhân của mình.')
    p=doc.add_paragraph(); r=p.add_run('Điều 3. Nghĩa vụ, quyền lợi NLĐ:'); r.bold=True
    p=doc.add_paragraph(); r=p.add_run('1. Nghĩa vụ:'); r.bold=True
    p=doc.add_paragraph('-    Hoàn thành những công việc được giao và sẵn sàng chấp nhận mọi sự điều động khi có yêu cầu;')
    p=doc.add_paragraph('-    Chấp hành nghiêm túc nội quy, kỷ luật lao động, an toàn lao động và các quy định của Công ty và pháp luật của Nhà nước;')
    p=doc.add_paragraph('-    Người lao động có trách nhiệm tuân thủ đầy đủ quy định về an toàn lao động, quy trình vận hành thiết bị và hướng dẫn của Công ty. Trường hợp NLĐ cố ý vi phạm hoặc vi phạm nghiêm trọng quy định an toàn lao động gây thiệt hại thì phải chịu trách nhiệm theo quy định pháp luật và nội quy Công ty;')
    p=doc.add_paragraph('-    Bồi thường vi phạm vật chất : Phải bồi thường vật chất do cá nhân vi phạm quy định của Công ty về bảo quản trang thiết bị được giao.')
    p=doc.add_paragraph(); r=p.add_run('2. Quyền Lợi:'); r.bold=True
    p=doc.add_paragraph('-    Phương tiện đi lại: Tự túc;')
    p=doc.add_paragraph('-    Được Công ty cấp đầy đủ bảo hộ lao động theo đúng vị trí làm việc;')
    p=doc.add_paragraph('-    Được phân công công việc theo yêu cầu của Công ty phù hợp với khả năng và trình độ chuyên môn mà người lao động đáp ứng;')
    p=doc.add_paragraph('-    Các quyền lợi khác thực hiện theo quy định của Pháp luật Lao động như tạm dừng, chấm dứt hợp đồng.')
    p=doc.add_paragraph(); r=p.add_run('Điều 4. Nghĩa vụ, quyền hạn NSDLĐ:'); r.bold=True
    p=doc.add_paragraph('-    Bảo đảm việc làm và thực hiện đầy đủ những điều đã cam kết trong hợp đồng;')
    p=doc.add_paragraph('-    Thanh toán đầy đủ, đúng hạn các chế độ và quyền lợi cho người lao động theo hợp đồng;')
    p=doc.add_paragraph('-    Điều hành người lao động hoàn thành công việc theo hợp đồng;')
    p=doc.add_paragraph('-    Tạm hoãn, chấm dứt hợp đồng theo quy định của pháp luật, và nội quy lao động của Công ty;')
    p=doc.add_paragraph(); r=p.add_run('Điều 5. Điều khoản chung:'); r.bold=True
    p=doc.add_paragraph('-    Những nội dung về quan hệ lao động không ghi trong hợp đồng này thì được áp dụng theo pháp luật lao động;')
    p=doc.add_paragraph('-    Những thoả thuận khác (nếu có): không;')
    p=doc.add_paragraph('-    Hợp đồng này có hiệu lực từ ngày ký và được làm thành 02 bản, Bên A giữ 01 bản, Bên B giữ 01 có giá trị pháp lý như nhau, để làm căn cứ thực hiện.'); doc.add_paragraph()
    p=doc.add_paragraph('Bản HĐ này lập tại văn phòng Công ty CP Cảng Hòn La.'); doc.add_paragraph()
    ts=doc.add_table(rows=3,cols=2); ts.alignment=WD_TABLE_ALIGNMENT.CENTER; remove_table_border(ts)
    c=ts.rows[0].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[0].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI SỬ DỤNG LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[1].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[1].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[2].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(nv.get('ho_ten','').upper()); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[2].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(CC['dai_dien'].upper()); r.bold=True; r.font.size=Pt(13)
    tf = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(tf.name)
    return tf.name

def gui_email(loai, ds, file=None):
    # Không import trong hàm nữa, dùng EMAIL_CONFIG đã có sẵn
    # from config import EMAIL_CONFIG as EC  <-- XÓA DÒNG NÀY
    
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG['email']
        msg['To'] = EMAIL_CONFIG['nguoi_nhan']
        tn = f"{datetime.now().month:02d}/{datetime.now().year}"
        msg['Subject'] = f"[HRM-Port] Báo cáo {loai} lao động tháng {tn}"
        nd = f"<h3>BÁO CÁO {loai.upper()} LĐ</h3><p>Tháng: <b>{tn}</b></p><p>SL: <b>{len(ds)}</b></p><hr><ul>"
        for nv in ds[:10]:
            nd += f"<li>{nv.get('ho_ten','')} - {nv.get('chuc_danh_nghe','')}</li>"
        if len(ds) > 10:
            nd += f"<li>... và {len(ds)-10} người khác</li>"
        nd += "</ul><p><b>File Excel đính kèm.</b></p>"
        msg.attach(MIMEText(nd, 'html', 'utf-8'))
        if file and os.path.exists(file):
            with open(file, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file)}"')
                msg.attach(part)
        srv = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
        srv.starttls()
        srv.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
        srv.send_message(msg)
        srv.quit()
        return True
    except Exception as e:
        st.error(f"Lỗi email: {e}")
        return False

def gui_telegram(msg):
    # from config import TELEGRAM_CONFIG as TC  <-- XÓA DÒNG NÀY
    
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_CONFIG['bot_token']}/sendMessage"
        r = requests.post(url, data={"chat_id": TELEGRAM_CONFIG['chat_id'], "text": msg, "parse_mode": "HTML"}, timeout=10)
        return r.status_code == 200
    except:
        return False

def tao_bao_cao_tang_giam(tang_list, giam_list, tu_ngay, den_ngay):
    """Tạo báo cáo Word tăng/giảm nhân sự"""
    from docx import Document
    from docx.shared import Pt, Cm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    
    doc = Document()
    
    # Tiêu đề
    title = doc.add_heading('BÁO CÁO TĂNG/GIẢM NHÂN SỰ', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    period = doc.add_paragraph(f'Thời gian: Từ {tu_ngay.strftime("%d/%m/%Y")} đến {den_ngay.strftime("%d/%m/%Y")}')
    period.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Danh sách tăng
    doc.add_heading('I. LAO ĐỘNG TĂNG MỚI', level=1)
    if tang_list:
        table = doc.add_table(rows=1 + len(tang_list), cols=5)
        table.style = 'Table Grid'
        # Header
        headers = ['STT', 'Họ tên', 'Chức danh', 'Loại HĐ', 'Ngày vào làm']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, nv in enumerate(tang_list, 1):
            row = table.rows[idx]
            row.cells[0].text = str(idx)
            row.cells[1].text = nv.get('ho_ten', '')
            row.cells[2].text = nv.get('chuc_danh_nghe', '')
            row.cells[3].text = nv.get('loai_hop_dong', '')
            row.cells[4].text = format_date(nv.get('ngay_vao_lam'))
    else:
        doc.add_paragraph('Không có lao động tăng trong kỳ.')
    
    doc.add_paragraph()
    
    # Danh sách giảm
    doc.add_heading('II. LAO ĐỘNG GIẢM (NGHỈ VIỆC)', level=1)
    if giam_list:
        table = doc.add_table(rows=1 + len(giam_list), cols=5)
        table.style = 'Table Grid'
        headers = ['STT', 'Họ tên', 'Chức danh', 'Loại HĐ', 'Ngày nghỉ việc']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, nv in enumerate(giam_list, 1):
            row = table.rows[idx]
            row.cells[0].text = str(idx)
            row.cells[1].text = nv.get('ho_ten', '')
            row.cells[2].text = nv.get('chuc_danh_nghe', '')
            row.cells[3].text = nv.get('loai_hop_dong', '')
            row.cells[4].text = format_date(nv.get('ngay_ket_thuc'))
    else:
        doc.add_paragraph('Không có lao động giảm trong kỳ.')
    
    # Footer
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(f'Ngày {date.today().day} tháng {date.today().month} năm {date.today().year}')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run('NGƯỜI LẬP BÁO CÁO')
    
    tf = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(tf.name)
    return tf.name

# ========== HÀM TẠO BÁO CÁO THỐNG KÊ NHÂN SỰ (MẪU THEO YÊU CẦU) ==========
def tao_bao_cao_thong_ke_nhan_su(tu_ngay, den_ngay, ds_nhan_vien):
    """
    Tạo báo cáo thống kê nhân sự theo mẫu yêu cầu
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Thong_ke_nhan_su"
    
    # Định nghĩa border
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Header thông tin đơn vị
    ten_cong_ty = COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA")
    
    # Dòng 1: Tên công ty
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    ws['A1'] = ten_cong_ty
    ws['A1'].font = Font(bold=True, size=14, name='Times New Roman')
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # Dòng 2: Tiêu đề báo cáo
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)
    ws['A2'] = "BÁO CÁO THỐNG KÊ NHÂN SỰ"
    ws['A2'].font = Font(bold=True, size=16, name='Times New Roman')
    ws['A2'].alignment = Alignment(horizontal='center')
    
    # Dòng 3: Thời gian
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=12)
    ws['A3'] = f"(Từ ngày {tu_ngay.strftime('%d/%m/%Y')} đến ngày {den_ngay.strftime('%d/%m/%Y')})"
    ws['A3'].font = Font(size=12, name='Times New Roman')
    ws['A3'].alignment = Alignment(horizontal='center')
    
    # Dòng 4: Ngày lập báo cáo
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=12)
    ws['A4'] = f"Ngày lập báo cáo: {date.today().strftime('%d/%m/%Y')}"
    ws['A4'].font = Font(size=11, name='Times New Roman', italic=True)
    ws['A4'].alignment = Alignment(horizontal='right')
    
    # Dòng 5 trống
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=12)
    
    # Header của bảng
    header_row = 6
    headers = [
        "STT", "Mã NV", "Họ và tên", "Ngày sinh", "Giới tính", "Chức danh",
        "Phòng ban", "Loại HĐ", "Ngày vào làm", "Ngày ký HĐ", "Ngày kết thúc", "Ghi chú"
    ]
    
    # Tạo header
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(bold=True, size=11, name='Times New Roman')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
    
    # Đổ dữ liệu
    data_row = header_row + 1
    for idx, nv in enumerate(ds_nhan_vien, 1):
        row = data_row + idx - 1
        
        ws.cell(row=row, column=1, value=idx)
        ws.cell(row=row, column=2, value=nv.get('ma_nv', ''))
        ws.cell(row=row, column=3, value=nv.get('ho_ten', ''))
        ws.cell(row=row, column=4, value=format_date(nv.get('ngay_sinh')))
        ws.cell(row=row, column=5, value='Nam' if nv.get('gioi_tinh') == 'Nam' else 'Nữ' if nv.get('gioi_tinh') == 'Nữ' else '')
        ws.cell(row=row, column=6, value=nv.get('chuc_danh_nghe', ''))
        ws.cell(row=row, column=7, value=nv.get('phong_ban_lam_viec', ''))
        ws.cell(row=row, column=8, value=nv.get('loai_hop_dong', ''))
        ws.cell(row=row, column=9, value=format_date(nv.get('ngay_vao_lam')))
        ws.cell(row=row, column=10, value=format_date(nv.get('ngay_ky_hd')))
        ws.cell(row=row, column=11, value=format_date(nv.get('ngay_ket_thuc')))
        ws.cell(row=row, column=12, value='')
        
        # Định dạng border cho từng ô
        for col_idx in range(1, 13):
            cell = ws.cell(row=row, column=col_idx)
            cell.border = thin_border
            cell.font = Font(size=10, name='Times New Roman')
            if col_idx in [1, 2, 4, 5, 8, 9, 10, 11]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
    
    # Tổng số nhân viên
    total_row = data_row + len(ds_nhan_vien)
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    ws.cell(row=total_row, column=1, value=f"Tổng số nhân viên: {len(ds_nhan_vien)}")
    ws.cell(row=total_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
    ws.cell(row=total_row, column=1).border = thin_border
    
    # Thống kê theo giới tính
    male_count = len([nv for nv in ds_nhan_vien if nv.get('gioi_tinh') == 'Nam'])
    female_count = len([nv for nv in ds_nhan_vien if nv.get('gioi_tinh') == 'Nữ'])
    other_count = len(ds_nhan_vien) - male_count - female_count
    
    stats_row = total_row + 2
    ws.cell(row=stats_row, column=1, value="Thống kê theo giới tính:")
    ws.cell(row=stats_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
    ws.merge_cells(start_row=stats_row, start_column=2, end_row=stats_row, end_column=4)
    ws.cell(row=stats_row, column=2, value=f"Nam: {male_count} người")
    ws.merge_cells(start_row=stats_row, start_column=5, end_row=stats_row, end_column=7)
    ws.cell(row=stats_row, column=5, value=f"Nữ: {female_count} người")
    ws.merge_cells(start_row=stats_row, start_column=8, end_row=stats_row, end_column=10)
    ws.cell(row=stats_row, column=8, value=f"Khác: {other_count} người")
    
    # Thống kê theo loại hợp đồng
    stats_row += 1
    ws.cell(row=stats_row, column=1, value="Thống kê theo loại hợp đồng:")
    ws.cell(row=stats_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
    
    # Lấy các loại hợp đồng khác nhau
    loai_hd_list = list(set([nv.get('loai_hop_dong', '') for nv in ds_nhan_vien]))
    col_start = 2
    for loai_hd in loai_hd_list:
        if loai_hd:
            count = len([nv for nv in ds_nhan_vien if nv.get('loai_hop_dong') == loai_hd])
            ws.merge_cells(start_row=stats_row, start_column=col_start, end_row=stats_row, end_column=col_start+2)
            ws.cell(row=stats_row, column=col_start, value=f"{loai_hd}: {count} người")
            col_start += 3
    
    # Footer ký tên
    sign_row = stats_row + 4
    ws.merge_cells(start_row=sign_row, start_column=9, end_row=sign_row, end_column=12)
    ws.cell(row=sign_row, column=9, value="NGƯỜI LẬP BÁO CÁO")
    ws.cell(row=sign_row, column=9).font = Font(bold=True, size=11, name='Times New Roman')
    ws.cell(row=sign_row, column=9).alignment = Alignment(horizontal='center')
    
    sign_row += 1
    ws.merge_cells(start_row=sign_row, start_column=9, end_row=sign_row, end_column=12)
    ws.cell(row=sign_row, column=9, value="(Ký, ghi rõ họ tên)")
    ws.cell(row=sign_row, column=9).font = Font(size=10, name='Times New Roman', italic=True)
    ws.cell(row=sign_row, column=9).alignment = Alignment(horizontal='center')
    
    sign_row += 2
    ws.merge_cells(start_row=sign_row, start_column=9, end_row=sign_row, end_column=12)
    ws.cell(row=sign_row, column=9, value=COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC').upper())
    ws.cell(row=sign_row, column=9).font = Font(bold=True, size=11, name='Times New Roman')
    ws.cell(row=sign_row, column=9).alignment = Alignment(horizontal='center')
    
    # Điều chỉnh độ rộng cột
    for col_idx in range(1, 13):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18
    
    # Lưu file
    filename = f"Bao_cao_thong_ke_nhan_su_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
    wb.save(filename)
    return filename

# ========== SIDEBAR + LOGIN ==========
st.sidebar.title("🏗️ HRM-Port")
st.sidebar.caption("Quản lý nhân sự cảng biển")

# Hàm kiểm tra đăng nhập từ secrets
def check_login(username, password):
    # Ưu tiên kiểm tra từ st.secrets trước (Streamlit Cloud)
    try:
        if 'users' in st.secrets and username in st.secrets.users:
            if st.secrets.users[username]['password'] == password:
                return True, st.secrets.users[username]['role']
    except:
        pass
    
    # Fallback: kiểm tra từ USERS trong config (local)
    try:
        if username in USERS:
            return USERS[username]['password'] == password, USERS[username]['role']
    except:
        pass
    
    return False, None

if 'logged_in' not in st.session_state: 
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None

if not st.session_state.logged_in:
    st.sidebar.subheader("🔐 Đăng nhập")
    u = st.sidebar.text_input("Tài khoản")
    p = st.sidebar.text_input("Mật khẩu", type="password")
    c1, c2 = st.sidebar.columns(2)
    with c1:
        if st.button("Đăng nhập", width='stretch'):
            success, role = check_login(u, p)
            if success:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = u
                st.rerun()
            else:
                st.sidebar.error("❌ Sai tài khoản hoặc mật khẩu!")
    with c2:
        if st.button("👁️ Xem thử", width='stretch'):
            st.session_state.logged_in = True
            st.session_state.role = "viewer"
            st.session_state.username = "guest"
            st.rerun()
    
    # Nút Back cho Guest
    if st.session_state.get('show_hrm', False) and not st.session_state.get('logged_in', False):
        st.markdown("<br><br><br>", unsafe_allow_html=True)  # Thêm khoảng trống
        if st.button("🔙 Quay lại Landing Page", width='stretch'):
            st.session_state.show_hrm = False
            st.session_state.pop('last_birthday_check', None)
            st.session_state.pop('sinh_nhat_hom_nay_list', None)
            st.rerun()
    
    st.stop()  
    
# Menu theo role
if st.session_state.role == "admin":
    menu_options = ["📊 Dashboard","👤 Ứng viên","✅ Nhân viên","📁 Upload hồ sơ","⚙️ Danh mục","📋 BHXH","📋 Báo cáo 01/PLI","👆 Chấm công (Face ID)","💰 Tính thu nhập"]
else:  # viewer
    menu_options = ["📊 Dashboard","👤 Ứng viên","✅ Nhân viên","📋 BHXH","📋 Báo cáo 01/PLI","👆 Chấm công (Face ID)","💰 Tính thu nhập"]
menu = st.sidebar.radio("📋 Menu", menu_options)
st.sidebar.divider()
st.sidebar.caption(f"👤 {st.session_state.username} ({st.session_state.role})")
# MỚI:
if st.sidebar.button("🚪 Đăng xuất", width='stretch'):
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None
    st.session_state.show_hrm = False
    st.session_state.pop('last_birthday_check', None)
    st.session_state.pop('sinh_nhat_hom_nay_list', None)
    st.cache_data.clear()
    st.rerun()

# ========== DASHBOARD ==========
if menu == "📊 Dashboard":
    st.title("📊 Dashboard")
    db = get_connection()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute("SELECT COUNT(*) t FROM ung_vien")
    tuv = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC')")
    tnv = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM ung_vien WHERE trang_thai='CHO_DUYET'")
    cd = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM ung_vien WHERE trang_thai='TU_CHOI'")
    tc = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM ung_vien WHERE trang_thai='DA_NHAN_VIEC'")
    dn = c.fetchone()['t']
    cl1, cl2, cl3, cl4, cl5 = st.columns(5)
    cl1.metric("Tổng UV", tuv)
    cl2.metric("Nhân viên", tnv)
    cl3.metric("Chờ duyệt", cd)
    cl4.metric("Đã nhận", dn)
    cl5.metric("Từ chối", tc)
    st.divider()
        
    # Gọi kiểm tra sinh nhật (đã xóa 2 dòng debug)
    auto_check_birthday()

    # 👇 Hiển thị banner cố định nếu có sinh nhật hôm nay
    sinh_nhat_list = st.session_state.get('sinh_nhat_hom_nay_list', [])
    if sinh_nhat_list:
        for sn in sinh_nhat_list:
            st.success(
                f"🎂 **Chúc mừng sinh nhật {sn['xung_ho']} {sn['ho_ten']} ({sn['ma_nv']})** — Hôm nay là sinh nhật của {sn['xung_ho']}! 🎉",
                icon="🎂"
            )
    
    # ========== PHẦN SINH NHẬT HOÀN CHỈNH ==========
    st.subheader("🎂 SINH NHẬT")

    # Tạo tabs cho sinh nhật
    tab_trong_thang, tab_hom_nay, tab_lich_su = st.tabs(["📅 Sinh nhật trong tháng", "🎉 Hôm nay", "📜 Lịch sử đã gửi"])

    with tab_trong_thang:
        # Lấy danh sách sinh nhật trong tháng
        c.execute("""
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, dien_thoai, email_lien_he, chuc_danh_nghe
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_sinh IS NOT NULL
            AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
            ORDER BY EXTRACT(DAY FROM ngay_sinh) ASC
        """)
        sinh_nhat_trong_thang = c.fetchall()
        
        if sinh_nhat_trong_thang:
            st.success(f"📅 Tháng {datetime.now().month} có **{len(sinh_nhat_trong_thang)}** nhân viên có sinh nhật:")
            
            # Hiển thị dạng grid
            cols = st.columns(3)
            for idx, sn in enumerate(sinh_nhat_trong_thang):
                with cols[idx % 3]:
                    ngay_sinh = sn.get('ngay_sinh')
                    if ngay_sinh:
                        # Tính tuổi
                        today = date.today()
                        tuoi = today.year - ngay_sinh.year
                        if today.month < ngay_sinh.month or (today.month == ngay_sinh.month and today.day < ngay_sinh.day):
                            tuoi -= 1
                        
                        # Tính ngày sinh nhật trong năm nay
                        sinh_nhat_nam_nay = date(today.year, ngay_sinh.month, ngay_sinh.day)
                        is_today = sinh_nhat_nam_nay == today
                        da_qua = sinh_nhat_nam_nay < today
                        
                        xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                        
                        if is_today:
                            st.markdown(f"""
                            <div style='background: linear-gradient(135deg, #ffd700 0%, #ffed4e 100%); 
                                        padding: 15px; border-radius: 15px; margin: 10px 0; 
                                        border: 2px solid #ff9800; box-shadow: 0 4px 8px rgba(0,0,0,0.1);'>
                                <div style='font-size: 30px; text-align: center;'>🎉🎂</div>
                                <h4 style='text-align: center; color: #d32f2f; margin: 5px 0;'>
                                    <b>HÔM NAY LÀ SINH NHẬT!</b>
                                </h4>
                                <h3 style='text-align: center; color: #333; margin: 5px 0;'>
                                    {xung_ho} <b>{sn['ho_ten']}</b>
                                </h3>
                                <p style='text-align: center; color: #666;'>
                                    📅 {format_date(ngay_sinh)} (🎂 {tuoi} tuổi)<br>
                                    💼 {sn.get('chuc_danh_nghe', 'Chưa cập nhật')}<br>
                                    📞 {sn.get('dien_thoai', 'Chưa cập nhật')}
                                </p>
                            </div>
                            """, unsafe_allow_html=True)
                        elif da_qua:
                            st.markdown(f"""
                            <div style='background-color: #e0e0e0; padding: 12px; border-radius: 10px; margin: 8px 0;'>
                                <b>✅ {sn['ho_ten']}</b><br>
                                🎂 Sinh ngày: {format_date(ngay_sinh)} (đã qua)<br>
                                💼 {sn.get('chuc_danh_nghe', '')}
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            ngay_con_lai = (sinh_nhat_nam_nay - today).days
                            st.markdown(f"""
                            <div style='background-color: #e3f2fd; padding: 12px; border-radius: 10px; margin: 8px 0;'>
                                <b>🎂 {sn['ho_ten']}</b><br>
                                📅 Sinh ngày: {format_date(ngay_sinh)}<br>
                                ⏰ Còn {ngay_con_lai} ngày nữa<br>
                                💼 {sn.get('chuc_danh_nghe', '')}
                            </div>
                            """, unsafe_allow_html=True)
        else:
            st.info("📭 Tháng này không có ai sinh nhật.")
            
            # Hiển thị sinh nhật tháng sau
            c.execute("""
                SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, chuc_danh_nghe
                FROM nhan_vien 
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                AND ngay_sinh IS NOT NULL
                AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE + INTERVAL '1 month')
                ORDER BY EXTRACT(DAY FROM ngay_sinh) ASC
                LIMIT 10
            """)
            sinh_nhat_thang_sau = c.fetchall()
            if sinh_nhat_thang_sau:
                st.caption("📅 Sinh nhật tháng sau:")
                for sn in sinh_nhat_thang_sau:
                    xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                    st.caption(f"🎂 {xung_ho} **{sn['ho_ten']}** - {format_date(sn['ngay_sinh'])}")

    with tab_hom_nay:
        # Lấy danh sách sinh nhật hôm nay
        c.execute("""
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, dien_thoai, email_lien_he, 
                   chuc_danh_nghe, phong_ban_lam_viec
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_sinh IS NOT NULL
            AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
            AND EXTRACT(DAY FROM ngay_sinh) = EXTRACT(DAY FROM CURRENT_DATE)
        """)
        sinh_nhat_hom_nay = c.fetchall()
        
        if sinh_nhat_hom_nay:
            st.balloons()
            for sn in sinh_nhat_hom_nay:
                ngay_sinh = sn.get('ngay_sinh')
                today = date.today()
                tuoi = today.year - ngay_sinh.year
                
                xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                
                st.markdown(f"""
                <div style='background: linear-gradient(135deg, #ff6b6b 0%, #ff8e8e 100%);
                            padding: 25px; border-radius: 20px; margin: 15px 0;
                            text-align: center; color: white;'>
                    <div style='font-size: 50px;'>🎉🎂🎉</div>
                    <h1 style='color: white; margin: 10px 0;'>CHÚC MỪNG SINH NHẬT!</h1>
                    <h2 style='color: #fff3e0; margin: 10px 0;'>{xung_ho} {sn['ho_ten']}</h2>
                    <p style='font-size: 18px;'>
                        🎂 {tuoi} tuổi - Một tuổi mới thật nhiều niềm vui! 🎂
                    </p>
                    <p style='margin-top: 15px;'>
                        📅 {format_date(ngay_sinh)} | 💼 {sn.get('chuc_danh_nghe', '')} | 
                        🏢 {sn.get('phong_ban_lam_viec', '')}
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                # Hiển thị thông tin liên hệ
                # Hiển thị thông tin liên hệ (tất cả đều thấy)
                col_phone, col_email_info = st.columns(2)
                with col_phone:
                    if sn.get('dien_thoai'):
                        st.markdown(f"📞 **SĐT:** {sn['dien_thoai']}")
                    else:
                        st.warning("⚠️ Chưa cập nhật số điện thoại")
                with col_email_info:
                    if sn.get('email_lien_he'):
                        st.markdown(f"📧 **Email:** {sn['email_lien_he']}")
                    else:
                        st.warning("⚠️ Chưa cập nhật email")

                # Nút gửi — CHỈ ADMIN mới thấy
                if st.session_state.get('role') == 'admin':
                    col_btn_zalo, col_btn_email = st.columns(2)
                    
                    with col_btn_zalo:
                        if sn.get('dien_thoai'):
                            sdt = sn['dien_thoai'].replace('+84', '0').replace(' ', '').strip()
                            if st.button(f"📱 Gửi Zalo cho {sn['ho_ten']}", 
                                         key=f"zalo_sn_{sn['id']}", 
                                         width='stretch', 
                                         type="primary"):
                                tuoi_nv = date.today().year - sn['ngay_sinh'].year
                                loi_chuc_nv = get_loi_chuc_sinh_nhat(sn['ho_ten'], sn.get('gioi_tinh'), tuoi_nv)
                                st.code(loi_chuc_nv)
                                st.markdown(f"[👉 NHẤN ĐỂ GỬI QUA ZALO CHO {sn['ho_ten']}](https://zalo.me/{sdt})")
                        else:
                            st.button("📱 Gửi Zalo", disabled=True, 
                                      key=f"zalo_sn_disabled_{sn['id']}", 
                                      width='stretch',
                                      help="Chưa có số điện thoại")

                    with col_btn_email:
                        if sn.get('email_lien_he'):
                            if st.button(f"📧 Gửi Email cho {sn['ho_ten']}", 
                                         key=f"email_sn_{sn['id']}", 
                                         width='stretch'):
                                st.info(f"📧 Email: {sn['email_lien_he']}")
                                st.toast(f"Chức năng gửi email đang phát triển!", icon="📧")
                        else:
                            st.button("📧 Gửi Email", disabled=True, 
                                      key=f"email_sn_disabled_{sn['id']}", 
                                      width='stretch',
                                      help="Chưa có email")
                else:
                    # Tài khoản thường — hiện thông báo thay vì nút
                    st.caption("💡 Liên hệ admin để gửi lời chúc sinh nhật.")
        else:
            st.info("🎉 Hôm nay không có ai sinh nhật.")
            
            # Gợi ý sinh nhật sắp tới
            c.execute("""
                SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh
                FROM nhan_vien 
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                AND ngay_sinh IS NOT NULL
                AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
                AND EXTRACT(DAY FROM ngay_sinh) > EXTRACT(DAY FROM CURRENT_DATE)
                ORDER BY EXTRACT(DAY FROM ngay_sinh) ASC
                LIMIT 3
            """)
            sinh_nhat_sap_toi = c.fetchall()
            if sinh_nhat_sap_toi:
                st.subheader("📅 Sinh nhật sắp tới trong tháng:")
                for sn in sinh_nhat_sap_toi:
                    xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                    st.info(f"🎂 {xung_ho} **{sn['ho_ten']}** - {format_date(sn['ngay_sinh'])}")

    with tab_lich_su:
        if st.session_state.role == "admin":
            st.subheader("📜 Lịch sử đã gửi lời chúc sinh nhật")
            
            # Kiểm tra bảng lịch sử tồn tại chưa
            try:
                c.execute("""
                    SELECT ls.*, nv.ho_ten, nv.ma_nv
                    FROM lich_su_gui_loi_chuc ls
                    JOIN nhan_vien nv ON ls.nhan_vien_id = nv.id
                    WHERE ls.loai_chuc = 'SINH_NHAT'
                    ORDER BY ls.ngay_gui DESC
                    LIMIT 50
                """)
                lich_su = c.fetchall()
                
                if lich_su:
                    ls_data = []
                    for ls in lich_su:
                        ls_data.append({
                            "Ngày gửi": format_date(ls['ngay_gui']),
                            "Mã NV": ls['ma_nv'],
                            "Họ tên": ls['ho_ten'],
                            "Kênh gửi": ls['kenh_gui'],
                            "Trạng thái": "✅ Đã gửi" if ls['trang_thai'] == 'DA_GUI' else ls['trang_thai']
                        })
                    df_ls = pd.DataFrame(ls_data)
                    st.dataframe(df_ls, width='stretch', hide_index=True)
                else:
                    st.info("📭 Chưa có lịch sử gửi lời chúc nào.")
            except Exception as e:
                st.info("📭 Chưa có dữ liệu lịch sử. Bảng lịch sử có thể chưa được tạo.")
        else:
            st.info("🔒 Chỉ Admin mới xem được lịch sử gửi lời chúc.")

    st.divider()

    # Nút gửi lời chúc sinh nhật (chỉ admin)
    if st.session_state.role == "admin" and sinh_nhat_trong_thang:
        with st.expander("💌 GỬI LỜI CHÚC SINH NHẬT", expanded=False):
            st.subheader("Gửi lời chúc sinh nhật đến nhân viên")
            
            # Chọn nhân viên để gửi
            sn_options = {}
            for sn in sinh_nhat_trong_thang:
                ngay_sinh = sn.get('ngay_sinh')
                xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                label = f"{xung_ho} {sn['ho_ten']} - {format_date(ngay_sinh)}"
                sn_options[label] = sn
            
            # SAU KHI SỬA (ĐÚNG) — dùng key động theo từng nhân viên
            selected_label = st.selectbox("Chọn nhân viên:", list(sn_options.keys()), key="chon_sn_gui")
            selected_sn = sn_options[selected_label]

            # Tính tuổi
            if selected_sn.get('ngay_sinh'):
                today = date.today()
                tuoi = today.year - selected_sn['ngay_sinh'].year
                if today.month < selected_sn['ngay_sinh'].month or (
                    today.month == selected_sn['ngay_sinh'].month and 
                    today.day < selected_sn['ngay_sinh'].day
                ):
                    tuoi -= 1
            else:
                tuoi = None

            default_chuc = get_loi_chuc_sinh_nhat(
                selected_sn['ho_ten'], 
                selected_sn.get('gioi_tinh'), 
                tuoi
            )

            # ✅ KEY ĐỘNG theo nv_id — buộc Streamlit re-render lại text_area mỗi khi chọn người mới
            loi_chuc = st.text_area(
                "📝 Lời chúc sinh nhật:", 
                value=default_chuc, 
                height=250, 
                key=f"loi_chuc_sn_{selected_sn['id']}"  # ← thay "loi_chuc_sn_gui" bằng key động này
            )
            
            col_zalo, col_email, col_cancel = st.columns(3)
            
            with col_zalo:
                if st.button("📱 GỬI QUA ZALO", width='stretch', type="primary"):
                    sdt = selected_sn.get('dien_thoai', '')
                    if sdt:
                        sdt = sdt.replace('+84', '0').replace(' ', '').strip()
                        st.code(loi_chuc)
                        st.markdown(f"[👉 NHẤN VÀO ĐÂY ĐỂ GỬI QUA ZALO CHO {selected_sn['ho_ten']}](https://zalo.me/{sdt})")
                        st.success(f"✅ Đã sao chép nội dung! Vui lòng nhấn link Zalo để gửi.")
                        
                        # Lưu lịch sử
                        try:
                            db_log = get_connection()
                            cur_log = db_log.cursor()
                            cur_log.execute("""
                                INSERT INTO lich_su_gui_loi_chuc (nhan_vien_id, loai_chuc, noi_dung, kenh_gui, trang_thai)
                                VALUES (%s, %s, %s, %s, %s)
                            """, (selected_sn['id'], 'SINH_NHAT', loi_chuc[:500], 'ZALO', 'DA_GUI'))
                            db_log.commit()
                            db_log.close()
                            st.toast("Đã lưu lịch sử gửi!", icon="✅")
                        except:
                            pass
                    else:
                        st.error("❌ Nhân viên chưa có số điện thoại! Vui lòng cập nhật SĐT trước.")
            
            with col_email:
                if st.button("📧 GỬI QUA EMAIL", width='stretch'):
                    email = selected_sn.get('email_lien_he', '')
                    if email:
                        try:
                            xung_ho = get_xung_ho(selected_sn.get('gioi_tinh'), selected_sn['ho_ten'])
                            
                            msg = MIMEMultipart()
                            msg['From'] = EMAIL_CONFIG['email']
                            msg['To'] = email
                            msg['Subject'] = f"🎂 Chúc mừng sinh nhật {xung_ho} {selected_sn['ho_ten']} - Công ty CP Cảng Hòn La"
                            
                            html_content = f"""
                            <html>
                            <head>
                                <meta charset="UTF-8">
                            </head>
                            <body style='font-family: "Times New Roman", Arial, sans-serif;'>
                                <div style='background: linear-gradient(135deg, #ffd700 0%, #ff9800 100%); 
                                            padding: 20px; text-align: center; border-radius: 10px;'>
                                    <h1 style='color: white;'>🎂 CHÚC MỪNG SINH NHẬT 🎂</h1>
                                </div>
                                <div style='padding: 20px; line-height: 1.6;'>
                                    {loi_chuc.replace(chr(10), '<br>')}
                                </div>
                                <hr>
                                <p style='color: #999; font-size: 11px; text-align: center;'>
                                    Email được gửi tự động từ hệ thống HRM-Port Công ty CP Cảng Hòn La<br>
                                    Địa chỉ: {COMPANY_CONFIG.get('dia_chi', '')} | Điện thoại: {COMPANY_CONFIG.get('dien_thoai_cty', '')}
                                </p>
                            </body>
                            </html>
                            """
                            msg.attach(MIMEText(html_content, 'html', 'utf-8'))
                            
                            server = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
                            server.starttls()
                            server.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
                            server.send_message(msg)
                            server.quit()
                            
                            st.success(f"✅ Đã gửi lời chúc sinh nhật qua email cho {xung_ho} {selected_sn['ho_ten']}!")
                            
                            # Lưu lịch sử
                            db_log = get_connection()
                            cur_log = db_log.cursor()
                            cur_log.execute("""
                                INSERT INTO lich_su_gui_loi_chuc (nhan_vien_id, loai_chuc, noi_dung, kenh_gui, trang_thai)
                                VALUES (%s, %s, %s, %s, %s)
                            """, (selected_sn['id'], 'SINH_NHAT', loi_chuc[:500], 'EMAIL', 'DA_GUI'))
                            db_log.commit()
                            db_log.close()
                        except Exception as e:
                            st.error(f"❌ Lỗi gửi email: {e}")
                    else:
                        st.error("❌ Nhân viên chưa có email! Vui lòng cập nhật email trước.")
            
            with col_cancel:
                st.write("")  # Placeholder

    # ========== KẾT THÚC PHẦN SINH NHẬT ==========
    
    # ── Phân bố chức danh ──
    import plotly.express as px
    import plotly.graph_objects as go

    db2 = get_connection()
    c2 = db2.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c2.execute("""
        SELECT chuc_danh_nghe, COUNT(*) t 
        FROM nhan_vien 
        WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
        GROUP BY chuc_danh_nghe 
        ORDER BY t DESC
    """)
    data = c2.fetchall()
    db2.close()

    if data:
        st.divider()
        st.subheader("📈 Phân bố nhân sự theo chức danh")

        df_phan_bo = pd.DataFrame(data)
        df_phan_bo.columns = ['Chức danh', 'Số lượng']
        df_phan_bo['Chức danh'] = df_phan_bo['Chức danh'].fillna('Chưa phân loại')

        col_chart, col_table = st.columns([3, 2])

        with col_chart:
            fig = go.Figure(data=[go.Pie(
                labels=df_phan_bo['Chức danh'],
                values=df_phan_bo['Số lượng'],
                hole=0.55,
                textinfo='label+percent',
                textposition='outside',
                textfont=dict(size=12),
                marker=dict(
                    colors=px.colors.qualitative.Safe,
                    line=dict(color='white', width=2)
                ),
                hovertemplate='<b>%{label}</b><br>Số lượng: %{value}<br>Tỷ lệ: %{percent}<extra></extra>'
            )])
            fig.update_layout(
                title=dict(
                    text=f"<b>Tổng: {df_phan_bo['Số lượng'].sum()} nhân viên</b>",
                    x=0.5, y=0.5,
                    xanchor='center', yanchor='middle',
                    font=dict(size=16, color='#0f3b5c')
                ),
                showlegend=False,
                margin=dict(t=40, b=40, l=20, r=20),
                height=400,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
            )
            st.plotly_chart(fig, width='stretch')

        with col_table:
            st.markdown("**📋 Chi tiết**")
            colors = px.colors.qualitative.Safe
            for i, row in df_phan_bo.iterrows():
                color = colors[i % len(colors)]
                pct = row['Số lượng'] / df_phan_bo['Số lượng'].sum() * 100
                st.markdown(f"""
                <div style='display:flex; align-items:center; padding:8px 10px; 
                            margin:4px 0; border-radius:8px; background:#f8f9fa;
                            border-left: 4px solid {color};'>
                    <div style='flex:1; font-size:13px; color:#333;'>
                        {row['Chức danh']}
                    </div>
                    <div style='font-weight:bold; color:{color}; font-size:15px; margin-left:8px;'>
                        {row['Số lượng']}
                    </div>
                    <div style='color:#999; font-size:12px; margin-left:6px;'>
                        ({pct:.0f}%)
                    </div>
                </div>
                """, unsafe_allow_html=True)

    st.divider()

    # ── Thông báo ──
    st.subheader("📌 Thông báo")
    db3 = get_connection()
    c3 = db3.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c3.execute("SELECT ho_ten FROM nhan_vien WHERE DATE(ngay_vao_lam)=CURRENT_DATE")
    hn = c3.fetchall()
    c3.execute("SELECT ho_ten FROM nhan_vien WHERE DATE(ngay_vao_lam)=CURRENT_DATE - INTERVAL '1 day'")
    hq = c3.fetchall()
    if hn:
        st.success(f"🟢 Hôm nay có thêm: **{', '.join([x['ho_ten'] for x in hn])}**")
    if hq:
        st.info(f"🔵 Hôm qua có thêm: **{', '.join([x['ho_ten'] for x in hq])}**")
    if st.session_state.role == "admin":
        c3.execute("""
            SELECT STT, ma_nv, ho_ten, ngay_vao_lam, 
                   (ngay_vao_lam + INTERVAL '30 days')::DATE as ngay_ket_thuc_tv,
                   ((ngay_vao_lam + INTERVAL '30 days')::DATE - CURRENT_DATE) as ngay_con_lai
            FROM nhan_vien 
            WHERE trang_thai = 'THU_VIEC' 
            AND (ngay_vao_lam + INTERVAL '30 days')::DATE <= CURRENT_DATE + INTERVAL '5 days'
            ORDER BY ngay_con_lai ASC
        """)
        tv_sap_het = c3.fetchall()
        for x in tv_sap_het:
            ngay_con_lai = x['ngay_con_lai']
            if isinstance(ngay_con_lai, timedelta):
                ngay_con_lai = ngay_con_lai.days
            if ngay_con_lai < 0:
                st.error(f"🔴 **{x.get('ma_nv','')} {x['ho_ten']}** - ĐÃ QUÁ THỜI HẠN THỬ VIỆC {abs(ngay_con_lai)} NGÀY!")
            elif ngay_con_lai == 0:
                st.error(f"⚠️ **{x.get('ma_nv','')} {x['ho_ten']}** - HÔM NAY LÀ NGÀY CUỐI HỢP ĐỒNG THỬ VIỆC!")
            else:
                st.warning(f"⚠️ **{x.get('ma_nv','')} {x['ho_ten']}** còn **{ngay_con_lai}** ngày sẽ kết thúc hợp đồng thử việc!")
    db3.close()
    
    
    if st.session_state.role == "admin":
        if st.button("💾 BACKUP DỮ LIỆU NGAY", width='stretch'):
            try:
                from backup_nv import backup_nhan_vien
                backup_nhan_vien()
                st.success("✅ Đã backup! Kiểm tra thư mục D:\\HRM_Port\\backup")
            except ImportError:
                st.error("❌ Không tìm thấy module backup_nv. Backup chỉ hoạt động trên local.")
    
    
    
# ========== ỨNG VIÊN ==========
elif menu == "👤 Ứng viên":
    st.title("👤 Ứng viên")
    su = st.text_input("🔍 Tìm kiếm", key="suv")
    
    # Kiểm tra nếu đang chuyển từ ứng viên sang nhân viên (chỉ admin)
    if st.session_state.role == "admin" and 'show_chuyen_nv_form' in st.session_state and st.session_state.show_chuyen_nv_form:
        st.subheader("📝 CHUYỂN ỨNG VIÊN THÀNH NHÂN VIÊN")
        uv_data = st.session_state.get('chuyen_uv_data', {})
        
        # Lấy danh sách chức danh từ database
        db_chuc = get_connection()
        c_chuc = db_chuc.cursor()
        c_chuc.execute("SELECT DISTINCT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
        dschucdanh = [row[0] for row in c_chuc.fetchall()]
        db_chuc.close()
        
        with st.form("chuyen_uv_to_nv_form"):
            st.markdown(f"**Ứng viên:** {uv_data.get('ho_ten', '')}")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                ho_ten_nv = st.text_input("Họ và tên *", value=uv_data.get('ho_ten', ''))
                ngay_sinh_nv = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(uv_data.get('ngay_sinh')), placeholder="dd/mm/yyyy", max_chars=10)
                gioi_tinh_nv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"], index=["", "Nam", "Nữ", "Khác"].index(uv_data.get('gioi_tinh', '')) if uv_data.get('gioi_tinh') in ["Nam", "Nữ", "Khác"] else 0)
                quoc_tich_nv = st.text_input("Quốc tịch", value="Việt Nam")
                dan_toc_nv = st.text_input("Dân tộc", value="Kinh")
            with col2:
                so_cccd_nv = st.text_input("CCCD")
                ngay_cap_cccd_nv = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                noi_cap_cccd_nv = st.text_input("Nơi cấp CCCD")
                nguyen_quan_nv = st.text_input("Nguyên quán")
                thuong_tru_nv = st.text_area("Thường trú", value=uv_data.get('ghi_chu', ''), height=68)
            with col3:
                dien_thoai_nv = st.text_input("SĐT", value=uv_data.get('dien_thoai', ''))
                email_nv = st.text_input("Email")
                chuc_danh_nv = st.selectbox("Chức danh", [""] + dschucdanh, index=([""] + dschucdanh).index(uv_data.get('vi_tri', '')) if uv_data.get('vi_tri', '') in dschucdanh else 0)
                phong_ban_nv = st.text_input("Phòng ban")
                noi_lam_viec_nv = st.text_input("Nơi làm việc", value="Cảng THQT Hòn La")
            
            st.divider()
            st.caption("💼 Hợp đồng & BHXH")
            col4, col5, col6 = st.columns(3)
            with col4:
                loai_hd_chuyen = st.selectbox("Loại HĐ *", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"])
                ngay_vao_lam_chuyen = st.date_input("Ngày vào làm", value=uv_data.get('ngay_vao_lam', date.today()))
                ngay_ket_thuc_chuyen = st.text_input("Ngày kết thúc (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                ma_bhxh_chuyen = st.text_input("Mã BHXH")
                bat_dau_bh_chuyen = st.text_input("Bắt đầu BH (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
            with col5:
                luong_bh_chuyen = st.text_input("Lương BH")
                he_so_luong_chuyen = st.text_input("Hệ số lương")
                pc_chuc_vu_chuyen = st.text_input("PC chức vụ")
                pc_tnvk_chuyen = st.text_input("PC TNVK (%)")
                pc_tnn_chuyen = st.text_input("PC TNN (%)")
            with col6:
                muc_huong_bhyt_chuyen = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"])
                ty_le_dong_chuyen = st.text_input("Tỷ lệ đóng (%)")
                muc_tien_dong_chuyen = st.text_input("Mức tiền đóng")
                phuong_thuc_dong_chuyen = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"])
                nhom_bhxh_chuyen = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"])
            
            st.divider()
            st.caption("🏦 Ngân hàng & KCB")
            col7, col8, col9 = st.columns(3)
            with col7:
                stk_chuyen = st.text_input("STK")
                chi_nhanh_nh_chuyen = st.text_input("Chi nhánh NH")
                tinh_kcb_chuyen = st.text_input("Tỉnh KCB")
                noi_kcb_chuyen = st.text_input("Nơi KCB")
            with col8:
                tinh_nhan_hs_chuyen = st.text_input("Tỉnh/TP nhận HS")
                phuong_nhan_hs_chuyen = st.text_input("Phường/Xã nhận HS")
                dia_chi_nhan_hs_chuyen = st.text_area("Địa chỉ nhận HS", height=100)
            with col9:
                dk_nhan_so_chuyen = st.selectbox("ĐK nhận sổ", ["Có", "Không"])
                ho_so_chuyen = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"])
            
            col_confirm1, col_confirm2 = st.columns(2)
            with col_confirm1:
                if st.form_submit_button("✅ XÁC NHẬN CHUYỂN", width='stretch', type="primary"):
                    if ho_ten_nv:
                        # Kiểm tra định dạng ngày
                        ngay_loi = []
                        if ngay_sinh_nv and not parse_date(ngay_sinh_nv): 
                            ngay_loi.append("Ngày sinh")
                        if ngay_cap_cccd_nv and not parse_date(ngay_cap_cccd_nv): 
                            ngay_loi.append("Ngày cấp CCCD")
                        if ngay_ket_thuc_chuyen and not parse_date(ngay_ket_thuc_chuyen): 
                            ngay_loi.append("Ngày kết thúc")
                        if bat_dau_bh_chuyen and not parse_date(bat_dau_bh_chuyen): 
                            ngay_loi.append("Bắt đầu BH")
                        if ngay_loi:
                            st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                        else:
                            try:
                                db = get_connection()
                                c = db.cursor()
                                
                                # Tạo STT và mã nhân viên mới
                                c.execute("SELECT COALESCE(MAX(STT),0)+1 FROM nhan_vien")
                                stt_moi = c.fetchone()[0]
                                ma_nv = f"NV{stt_moi:03d}"
                                
                                nhl = ngay_vao_lam_chuyen
                                tbd_val = parse_date(bat_dau_bh_chuyen) if bat_dau_bh_chuyen and bat_dau_bh_chuyen.strip() else None
                                
                                # Tạo số hợp đồng theo loại
                                if loai_hd_chuyen == "Thử việc":
                                    trang_thai_nv = 'THU_VIEC'
                                    trang_thai_bhxh = 'CHUA_DONG'
                                    c.execute("""
                                        SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) 
                                        FROM nhan_vien 
                                        WHERE so_hdld LIKE '%/HĐTV-CHL' 
                                        AND trang_thai IN ('THU_VIEC', 'DANG_LAM')
                                    """)
                                    tv_cnt = c.fetchone()[0] + 1
                                    so_hd = f"{tv_cnt:02d}/{nhl.year}/HĐTV-CHL"
                                else:
                                    trang_thai_nv = 'DANG_LAM'
                                    trang_thai_bhxh = 'DANG_DONG'
                                    if not tbd_val:
                                        tbd_val = nhl
                                    c.execute("""
                                        SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) 
                                        FROM nhan_vien 
                                        WHERE so_hdld LIKE '%/HĐLĐ-CHL' 
                                        AND trang_thai = 'DANG_LAM'
                                        AND loai_hop_dong != 'Thử việc'
                                    """)
                                    so_hd_cnt = c.fetchone()[0] or 0
                                    so_hd = f"{so_hd_cnt + 1:02d}/{nhl.year}/HĐLĐ-CHL"
                                
                                # Thêm nhân viên mới
                                c.execute("""
                                    INSERT INTO nhan_vien (STT, ma_nv, so_hdld, ho_ten, chuc_danh_nghe, 
                                        ngay_sinh, gioi_tinh, so_cccd, ngay_cap_cccd, noi_cap_cccd,
                                        nguyen_quan, thuong_tru, dien_thoai, email, email_lien_he, ho_so,
                                        luong_bao_hiem, ma_so_bhxh, ngay_vao_lam, noi_lam_viec,
                                        so_tai_khoan_nh, chi_nhanh_nh, ngay_ky_hd, loai_hop_dong,
                                        nhom_bhxh, thang_bat_dau_bh, thang_ket_thuc_bh, trang_thai, trang_thai_bhxh,
                                        phong_ban_lam_viec, ngay_ket_thuc, quoc_tich, dan_toc, 
                                        he_so_luong, phu_cap_chuc_vu, phu_cap_tnvk, phu_cap_tnn,
                                        muc_huong_bhyt, ty_le_dong, muc_tien_dong, phuong_thuc_dong,
                                        tinh_nhan_hs, phuong_nhan_hs, dia_chi_nhan_hs, 
                                        tinh_kcb, noi_dang_ky_kcb, dang_ky_nhan_so)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                    RETURNING id
                                """, (
                                    stt_moi, ma_nv, so_hd, ho_ten_nv, chuc_danh_nv,
                                    parse_date(ngay_sinh_nv), gioi_tinh_nv, so_cccd_nv, parse_date(ngay_cap_cccd_nv), noi_cap_cccd_nv,
                                    nguyen_quan_nv, thuong_tru_nv, dien_thoai_nv, email_nv, email_nv, ho_so_chuyen,
                                    luong_bh_chuyen, ma_bhxh_chuyen, ngay_vao_lam_chuyen, noi_lam_viec_nv,
                                    stk_chuyen, chi_nhanh_nh_chuyen, ngay_vao_lam_chuyen, loai_hd_chuyen,
                                    nhom_bhxh_chuyen, tbd_val, parse_date(ngay_ket_thuc_chuyen), trang_thai_nv, trang_thai_bhxh,
                                    phong_ban_nv, parse_date(ngay_ket_thuc_chuyen), quoc_tich_nv, dan_toc_nv,
                                    to_float_or_none(he_so_luong_chuyen), to_float_or_none(pc_chuc_vu_chuyen),
                                    to_float_or_none(pc_tnvk_chuyen), to_float_or_none(pc_tnn_chuyen),
                                    muc_huong_bhyt_chuyen, to_float_or_none(ty_le_dong_chuyen), to_float_or_none(muc_tien_dong_chuyen),
                                    phuong_thuc_dong_chuyen, tinh_nhan_hs_chuyen, phuong_nhan_hs_chuyen, dia_chi_nhan_hs_chuyen,
                                    tinh_kcb_chuyen, noi_kcb_chuyen, dk_nhan_so_chuyen
                                ))
                                nhan_vien_id_moi = c.fetchone()[0]
                                # Cập nhật trạng thái ứng viên
                                c.execute("UPDATE ung_vien SET trang_thai='DA_NHAN_VIEC', ma_nv=%s WHERE id=%s", 
                                         (ma_nv, st.session_state['chuyen_uv_id']))
                                
                                # Thêm lịch sử công tác
                                c.execute("""
                                    INSERT INTO lich_su_cong_tac (nhan_vien_id, tu_ngay, chuc_danh, phong_ban, noi_lam_viec, loai_hop_dong, he_so_luong, so_hop_dong)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                                """, (nhan_vien_id_moi, ngay_vao_lam_chuyen, chuc_danh_nv, phong_ban_nv, noi_lam_viec_nv, loai_hd_chuyen, 
                                      to_float_or_none(he_so_luong_chuyen), so_hd))
                                
                                db.commit()
                                db.close()
                                
                                st.success(f"✅ Đã chuyển {ho_ten_nv} thành nhân viên! Mã NV: {ma_nv}")
                                
                                # Xóa session state
                                del st.session_state['show_chuyen_nv_form']
                                del st.session_state['chuyen_uv_id']
                                del st.session_state['chuyen_uv_data']
                                st.rerun()
                                
                            except Exception as e:
                                db.rollback()
                                db.close()
                                st.error(f"❌ Lỗi khi chuyển: {e}")
                    else:
                        st.error("Họ tên không được để trống!")
            
            with col_confirm2:
                if st.form_submit_button("❌ HỦY", width='stretch'):
                    del st.session_state['show_chuyen_nv_form']
                    del st.session_state['chuyen_uv_id']
                    del st.session_state['chuyen_uv_data']
                    st.rerun()
        
        st.divider()
        st.stop()  # Dừng lại để không hiển thị danh sách ứng viên phía dưới
    
    # Đóng db_f
    db_f = get_connection()
    c_f = db_f.cursor()
    c_f.execute("SELECT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
    ds_vi_tri = [row[0] for row in c_f.fetchall()]
    c_f.execute("SELECT DISTINCT vi_tri_du_tuyen FROM ung_vien WHERE vi_tri_du_tuyen IS NOT NULL AND vi_tri_du_tuyen != '' ORDER BY vi_tri_du_tuyen")
    for row in c_f.fetchall():
        if row[0] not in ds_vi_tri:
            ds_vi_tri.append(row[0])
    db_f.close()
    
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        filter_vi_tri = st.selectbox("🔍 Lọc Vị trí dự tuyển:", ["Tất cả"] + ds_vi_tri)
    
    # Chỉ admin mới thấy nút thêm ứng viên
    if st.session_state.role == "admin":
        with st.expander("➕ THÊM ỨNG VIÊN MỚI", expanded=False):
            with st.form("add_uv_form"):
                db_f = get_connection()
                c_f = db_f.cursor()
                c_f.execute("SELECT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
                ds_vt_uv = [row[0] for row in c_f.fetchall()]
                db_f.close()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    ho_ten_uv = st.text_input("Họ và tên *")
                    vi_tri_uv = st.selectbox("Vị trí dự tuyển", [""] + ds_vt_uv)
                    dien_thoai_uv = st.text_input("SĐT")
                with col2:
                    ngay_sinh_uv = st.text_input("Ngày sinh (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    gioi_tinh_uv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"])
                with col3:
                    ngay_vao_lam_uv = st.text_input("Ngày vào làm (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    ghi_chu_uv = st.text_area("Ghi chú")
                
                if st.form_submit_button("💾 LƯU", width='stretch'):
                    if ho_ten_uv:
                        ngay_loi = []
                        if ngay_sinh_uv and not parse_date(ngay_sinh_uv): 
                            ngay_loi.append("Ngày sinh")
                        if ngay_vao_lam_uv and not parse_date(ngay_vao_lam_uv): 
                            ngay_loi.append("Ngày vào làm")
                        if ngay_loi:
                            st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                        else:
                            try:
                                db = get_connection()
                                c = db.cursor()
                                c.execute("""INSERT INTO ung_vien (ho_ten, vi_tri_du_tuyen, dien_thoai, 
                                    ngay_sinh, gioi_tinh, ngay_vao_lam, luong_bao_hiem, trang_thai)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, 'CHO_DUYET')""",
                                    (ho_ten_uv, vi_tri_uv, dien_thoai_uv, parse_date(ngay_sinh_uv),
                                     gioi_tinh_uv, parse_date(ngay_vao_lam_uv), ghi_chu_uv))
                                new_id = c.lastrowid
                                ma_uv = f"UV{new_id:04d}"
                                c.execute("UPDATE ung_vien SET ma_uv = %s WHERE id = %s", (ma_uv, new_id))
                                db.commit()
                                db.close()
                                st.success(f"✅ Đã thêm ứng viên: {ho_ten_uv} (Mã: {ma_uv})")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi thêm ứng viên: {e}")
                    else:
                        st.error("Họ tên không được để trống!")
    
    st.divider()
    
    t1, t2, t3, t4 = st.tabs(["📋 Tất cả", "⏳ Chờ duyệt", "✅ Đã nhận", "❌ Từ chối"])
    tm = {"📋 Tất cả": "", "⏳ Chờ duyệt": "CHO_DUYET", "✅ Đã nhận": "DA_NHAN_VIEC", "❌ Từ chối": "TU_CHOI"}
    
    for tn, tab in zip(tm.keys(), [t1, t2, t3, t4]):
        with tab:
            tt = tm[tn]
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            sql = "SELECT id, ma_uv, ho_ten, vi_tri_du_tuyen, dien_thoai, ngay_vao_lam, luong_bao_hiem, ngay_sinh, trang_thai FROM ung_vien WHERE 1=1"
            params = []
            if tt:
                sql += " AND trang_thai = %s"
                params.append(tt)
            if su:
                sql += " AND (ho_ten LIKE %s OR dien_thoai LIKE %s OR ma_uv LIKE %s)"
                params.extend([f'%{su}%', f'%{su}%', f'%{su}%'])
            if filter_vi_tri != "Tất cả":
                sql += " AND vi_tri_du_tuyen = %s"
                params.append(filter_vi_tri)
            sql += " ORDER BY id ASC"
            c.execute(sql, tuple(params))
            ds = c.fetchall()
            db.close()
            
            if ds:
                df = pd.DataFrame(ds)
                for col in df.columns:
                    if 'ngay' in col.lower():
                        df[col] = df[col].apply(format_date)
                
                display_cols = ['ma_uv', 'ho_ten', 'vi_tri_du_tuyen', 'dien_thoai', 'ngay_vao_lam', 'luong_bao_hiem', 'ngay_sinh', 'trang_thai']
                available_cols = [c for c in display_cols if c in df.columns]
                df_show = df[available_cols]
                
                col_map = {
                    'ma_uv': 'Mã UV',
                    'ho_ten': 'Họ tên',
                    'vi_tri_du_tuyen': 'Vị trí dự tuyển',
                    'dien_thoai': 'SĐT',
                    'ngay_vao_lam': 'Ngày vào làm',
                    'luong_bao_hiem': 'Ghi chú',
                    'ngay_sinh': 'Ngày sinh',
                    'trang_thai': 'Trạng thái',
                }
                df_show.rename(columns=col_map, inplace=True)
                
                st.caption(f"📌 {len(ds)} kết quả.")
                
                if st.session_state.role == "admin":
                    # Admin: hiển thị bảng có checkbox và nút chức năng
                    if 'selected' not in df.columns:
                        df.insert(0, 'selected', False)
                    df_show_with_checkbox = df[['selected'] + [c for c in df.columns if c in display_cols]]
                    df_show_with_checkbox.rename(columns={'selected': 'Chọn'}, inplace=True)
                    
                    edited_df = st.data_editor(
                        df_show_with_checkbox,
                        column_config={"Chọn": st.column_config.CheckboxColumn("Chọn", default=False)},
                        disabled=[col for col in df_show_with_checkbox.columns if col != 'Chọn'],
                        hide_index=True,
                        height=400,
                        key=f"uv_editor_{tn}"
                    )
                    
                    if edited_df is not None:
                        selected_rows = edited_df[edited_df['Chọn'] == True]
                        if len(selected_rows) > 1:
                            st.error("⚠️ Chỉ được chọn 1 ứng viên!")
                        elif len(selected_rows) == 1:
                            selected_idx = selected_rows.index[0]
                            selected_nv = df.iloc[selected_idx]
                            col_btn1, col_btn2 = st.columns(2)
                            with col_btn1:
                                if st.button(f"✏️ SỬA", key=f"edit_sel_{tn}"):
                                    st.session_state['edit_uv_id'] = int(selected_nv['id'])
                                    st.rerun()
                            if tn == "⏳ Chờ duyệt":
                                with col_btn2:
                                    if st.button(f"👥 CHUYỂN SANG NHÂN VIÊN", type="primary", key=f"chuyen_uv_{tn}"):
                                        try:
                                            db = get_connection()
                                            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                            uv_id = int(selected_nv['id'])
                                            c.execute("SELECT * FROM ung_vien WHERE id = %s", (uv_id,))
                                            uv = c.fetchone()
                                            db.close()
                                            
                                            if uv:
                                                # Lưu thông tin ứng viên vào session_state
                                                st.session_state['chuyen_uv_id'] = uv_id
                                                st.session_state['chuyen_uv_data'] = {
                                                    'ho_ten': uv['ho_ten'],
                                                    'vi_tri': uv['vi_tri_du_tuyen'],
                                                    'dien_thoai': uv['dien_thoai'],
                                                    'ngay_sinh': uv['ngay_sinh'],
                                                    'gioi_tinh': uv['gioi_tinh'],
                                                    'ngay_vao_lam': uv['ngay_vao_lam'] or date.today(),
                                                    'ghi_chu': uv['luong_bao_hiem']
                                                }
                                                st.session_state['show_chuyen_nv_form'] = True
                                                st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi: {e}")
                else:
                    # Viewer: chỉ hiển thị bảng (không có checkbox và nút chuyển)
                    # Nhưng viewer vẫn có thể xem thông tin và sửa? Yêu cầu: "bổ sung 1 tk mới có phân quyền giống như user + cho phép thêm mới, sửa thông tin ứng viên"
                    # => Viewer cũng được phép thêm mới và sửa ứng viên
                    
                    # Viewer: hiển thị bảng có nút sửa (không có checkbox)
                    st.dataframe(df_show, width='stretch', hide_index=True, height=400)
                    
                    # Viewer: chọn ứng viên để sửa
                    st.caption("💡 Chọn ứng viên để sửa thông tin:")
                    uv_options = {f"{row['ho_ten']} - {row['vi_tri_du_tuyen']}": row['id'] for row in ds}
                    selected_uv_name = st.selectbox("Chọn ứng viên:", list(uv_options.keys()), key=f"select_uv_{tn}")
                    selected_uv_id = uv_options[selected_uv_name]
                    
                    if st.button(f"✏️ SỬA ỨNG VIÊN", key=f"edit_uv_viewer_{tn}"):
                        st.session_state['edit_uv_id'] = selected_uv_id
                        st.rerun()
                    
                    # Viewer: thêm ứng viên mới (dùng key riêng để tránh trùng với form của admin)
                    with st.expander("➕ THÊM ỨNG VIÊN MỚI", expanded=False):
                        with st.form(f"add_uv_form_viewer_{tn}"):  # Key động theo tab
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                ho_ten_uv = st.text_input("Họ và tên *", key=f"viewer_ho_ten_{tn}")
                                vi_tri_uv = st.selectbox("Vị trí dự tuyển", [""] + ds_vi_tri, key=f"viewer_vi_tri_{tn}")
                                dien_thoai_uv = st.text_input("SĐT", key=f"viewer_dien_thoai_{tn}")
                            with col2:
                                ngay_sinh_uv = st.text_input("Ngày sinh (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10, key=f"viewer_ngay_sinh_{tn}")
                                gioi_tinh_uv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"], key=f"viewer_gioi_tinh_{tn}")
                            with col3:
                                ngay_vao_lam_uv = st.text_input("Ngày vào làm (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10, key=f"viewer_ngay_vao_lam_{tn}")
                                ghi_chu_uv = st.text_area("Ghi chú", key=f"viewer_ghi_chu_{tn}")
                            
                            if st.form_submit_button("💾 LƯU", width='stretch'):
                                if ho_ten_uv:
                                    ngay_loi = []
                                    if ngay_sinh_uv and not parse_date(ngay_sinh_uv): 
                                        ngay_loi.append("Ngày sinh")
                                    if ngay_vao_lam_uv and not parse_date(ngay_vao_lam_uv): 
                                        ngay_loi.append("Ngày vào làm")
                                    if ngay_loi:
                                        st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                                    else:
                                        try:
                                            db = get_connection()
                                            c = db.cursor()
                                            c.execute("""INSERT INTO ung_vien (ho_ten, vi_tri_du_tuyen, dien_thoai, 
                                                ngay_sinh, gioi_tinh, ngay_vao_lam, luong_bao_hiem, trang_thai)
                                                VALUES (%s, %s, %s, %s, %s, %s, %s, 'CHO_DUYET')""",
                                                (ho_ten_uv, vi_tri_uv, dien_thoai_uv, parse_date(ngay_sinh_uv),
                                                 gioi_tinh_uv, parse_date(ngay_vao_lam_uv), ghi_chu_uv))
                                            new_id = c.lastrowid
                                            ma_uv = f"UV{new_id:04d}"
                                            c.execute("UPDATE ung_vien SET ma_uv = %s WHERE id = %s", (ma_uv, new_id))
                                            db.commit()
                                            db.close()
                                            st.success(f"✅ Đã thêm ứng viên: {ho_ten_uv} (Mã: {ma_uv})")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi khi thêm ứng viên: {e}")
                                else:
                                    st.error("Họ tên không được để trống!")
            else:
                st.info("Không có dữ liệu")
    
    # Form sửa ứng viên (admin và viewer đều có thể sửa)
    if 'edit_uv_id' in st.session_state:
        st.divider()
        st.subheader(f"✏️ Sửa ứng viên")
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT * FROM ung_vien WHERE id = %s", (st.session_state['edit_uv_id'],))
        uv_data = c.fetchone()
        db.close()
        if uv_data:
            with st.form("edit_uv_direct"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    ho_ten_e = st.text_input("Họ và tên *", value=uv_data['ho_ten'] or '')
                    vi_tri_e = st.selectbox("Vị trí dự tuyển", [""] + ds_vi_tri,
                        index=([""] + ds_vi_tri).index(uv_data['vi_tri_du_tuyen']) if uv_data['vi_tri_du_tuyen'] in ds_vi_tri else 0)
                    dien_thoai_e = st.text_input("SĐT", value=uv_data['dien_thoai'] or '')
                with col2:
                    ngay_sinh_e = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(uv_data['ngay_sinh']))
                    gioi_tinh_e = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"],
                        index=["", "Nam", "Nữ", "Khác"].index(uv_data['gioi_tinh']) if uv_data['gioi_tinh'] in ["Nam", "Nữ", "Khác"] else 0)
                with col3:
                    ngay_vao_lam_e = st.text_input("Ngày vào làm (dd/mm/yyyy)", value=format_date(uv_data['ngay_vao_lam']))
                    ghi_chu_e = st.text_area("Ghi chú", value=uv_data['luong_bao_hiem'] or '')
                    # Viewer không thể thay đổi trạng thái ứng viên (chỉ admin mới có quyền)
                    if st.session_state.role == "admin":
                        trang_thai_e = st.selectbox("Trạng thái", ["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"],
                            index=["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"].index(uv_data['trang_thai']) if uv_data['trang_thai'] in ["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"] else 0)
                    else:
                        # Viewer: chỉ hiển thị trạng thái, không cho sửa
                        st.text_input("Trạng thái", value=uv_data['trang_thai'], disabled=True)
                        trang_thai_e = uv_data['trang_thai']
                
                col_save, col_del, col_cancel = st.columns(3)
                with col_save:
                    if st.form_submit_button("💾 CẬP NHẬT"):
                        ngay_loi = []
                        if ngay_sinh_e and not parse_date(ngay_sinh_e): 
                            ngay_loi.append("Ngày sinh")
                        if ngay_vao_lam_e and not parse_date(ngay_vao_lam_e): 
                            ngay_loi.append("Ngày vào làm")
                        if ngay_loi:
                            st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                        else:
                            db = get_connection()
                            c = db.cursor()
                            c.execute("""UPDATE ung_vien SET ho_ten=%s, vi_tri_du_tuyen=%s, dien_thoai=%s,
                                ngay_sinh=%s, gioi_tinh=%s, ngay_vao_lam=%s, luong_bao_hiem=%s, trang_thai=%s
                                WHERE id=%s""",
                                (ho_ten_e, vi_tri_e, dien_thoai_e, parse_date(ngay_sinh_e), gioi_tinh_e,
                                 parse_date(ngay_vao_lam_e), ghi_chu_e, trang_thai_e, uv_data['id']))
                            db.commit()
                            db.close()
                            st.success("✅ Đã cập nhật!")
                            del st.session_state['edit_uv_id']
                            st.rerun()
                # Chỉ admin mới có quyền xóa ứng viên
                if st.session_state.role == "admin":
                    with col_del:
                        if st.form_submit_button("🗑️ XÓA"):
                            db = get_connection()
                            c = db.cursor()
                            c.execute("DELETE FROM ung_vien WHERE id = %s", (uv_data['id'],))
                            db.commit()
                            db.close()
                            st.success("🗑️ Đã xóa!")
                            del st.session_state['edit_uv_id']
                            st.rerun()
                with col_cancel:
                    if st.form_submit_button("❌ HỦY"):
                        del st.session_state['edit_uv_id']
                        st.rerun()
    
    # Quản lý danh mục vị trí dự tuyển (chỉ admin)
    if st.session_state.role == "admin":
        st.divider()
        with st.expander("⚙️ Quản lý danh mục Vị trí dự tuyển", expanded=False):
            with st.form("add_vi_tri_uv"):
                ten_vt_moi = st.text_input("Tên vị trí dự tuyển mới *")
                if st.form_submit_button("➕ Thêm"):
                    if ten_vt_moi:
                        db = get_connection()
                        c = db.cursor()
                        c.execute("SELECT COUNT(*) FROM vi_tri_cong_tac WHERE ten_vi_tri = %s", (ten_vt_moi,))
                        if c.fetchone()[0] == 0:
                            c.execute("INSERT INTO vi_tri_cong_tac (ten_vi_tri) VALUES (%s)", (ten_vt_moi,))
                            db.commit()
                            st.success(f"✅ Đã thêm: {ten_vt_moi}")
                            st.rerun()
                        else:
                            st.warning("Vị trí này đã tồn tại!")
                        db.close()
                    else:
                        st.error("Tên không được để trống!")
            
            db = get_connection()
            c = db.cursor()
            c.execute("SELECT id, ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
            ds_vt = c.fetchall()
            db.close()
            if ds_vt:
                st.caption("📋 Danh sách vị trí dự tuyển:")
                for row in ds_vt:
                    st.write(f"- {row[1]}")

# ========== NHÂN VIÊN ==========
elif menu == "✅ Nhân viên":
    st.title("✅ Quản lý nhân viên")
    
    tab_dang_lam, tab_da_nghi, tab_qtct = st.tabs(["📌 ĐANG LÀM VIỆC", "📋 ĐÃ NGHỈ VIỆC", "📜 LỊCH SỬ CÔNG TÁC"])
    
    with tab_dang_lam:
        st.caption("👥 Danh sách nhân viên đang làm việc (bao gồm thử việc)")
        sn = st.text_input("🔍 Tìm kiếm", key="snv_dang_lam")
        
        if st.session_state.role == "admin":
            with st.expander("➕ THÊM NHÂN VIÊN MỚI", expanded=False):
                with st.form("add_nv"):
                    db = get_connection()
                    c = db.cursor()
                    c.execute("SELECT DISTINCT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
                    dcv = [row[0] for row in c.fetchall()]
                    db.close()
                    st.caption("📝 Thông tin cá nhân")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        htn = st.text_input("Họ và tên *")
                        nsn = st.text_input("Ngày sinh (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                        gtn = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"])
                        qtn = st.text_input("Quốc tịch", value="Việt Nam")
                        dtn = st.text_input("Dân tộc", value="Kinh")
                    with c2:
                        scc = st.text_input("CCCD")
                        ncc = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                        ncc2 = st.text_input("Nơi cấp CCCD")
                        nqn = st.text_input("Nguyên quán")
                        ttn = st.text_input("Thường trú")
                    with c3:
                        dtn2 = st.text_input("SĐT")
                        emn = st.text_input("Email")
                        cdn = st.selectbox("Chức danh", [""] + dcv)
                        pbn = st.text_input("Phòng ban")
                        nlv = st.text_input("Nơi làm việc", value="Cảng THQT Hòn La")
                    st.divider()
                    st.caption("💼 Hợp đồng & BHXH")
                    c4, c5, c6 = st.columns(3)
                    with c4:
                        lhd = st.selectbox("Loại HĐ *", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"])
                        nvl = st.text_input("Ngày vào làm (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                        nkt = st.text_input("Ngày kết thúc", placeholder="dd/mm/yyyy", max_chars=10)
                        mbh = st.text_input("Mã BHXH")
                        tbd = st.text_input("Bắt đầu BH (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    with c5:
                        lbh = st.text_input("Lương BH")
                        hsl = st.text_input("Hệ số lương")
                        pcv = st.text_input("PC chức vụ")
                        ptv = st.text_input("PC TNVK (%)")
                        ptn = st.text_input("PC TNN (%)")
                    with c6:
                        mhb = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"])
                        tld = st.text_input("Tỷ lệ đóng (%)")
                        mtd = st.text_input("Mức tiền đóng")
                        ptd = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"])
                        nbh = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"])
                    st.divider()
                    st.caption("🏦 Ngân hàng & KCB")
                    c7, c8, c9 = st.columns(3)
                    with c7:
                        stk = st.text_input("STK")
                        cnh = st.text_input("Chi nhánh NH")
                        tkb = st.text_input("Tỉnh KCB")
                        nkb = st.text_input("Nơi KCB")
                    with c8:
                        ths = st.text_input("Tỉnh/TP nhận HS")
                        phs = st.text_input("Phường/Xã nhận HS")
                        dhs = st.text_area("Địa chỉ nhận HS", height=100)
                    with c9:
                        dks = st.selectbox("ĐK nhận sổ", ["Có", "Không"])
                        hso = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"])
                    
                    if st.form_submit_button("💾 LƯU"):
                        if htn:
                            ngay_loi = []
                            if nsn and not parse_date(nsn):
                                ngay_loi.append("Ngày sinh")
                            if ncc and not parse_date(ncc):
                                ngay_loi.append("Ngày cấp CCCD")
                            if nvl and not parse_date(nvl):
                                ngay_loi.append("Ngày vào làm")
                            if nkt and not parse_date(nkt):
                                ngay_loi.append("Ngày kết thúc")
                            if tbd and not parse_date(tbd):
                                ngay_loi.append("Bắt đầu BH")
                            if ngay_loi:
                                st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                            else:
                                try:
                                    db = get_connection()
                                    c = db.cursor()
                                    c.execute("SELECT COALESCE(MAX(STT),0)+1 FROM nhan_vien")
                                    stt_moi = c.fetchone()[0]
                                    ma_nv = f"NV{stt_moi:03d}"
                                    nhl = parse_date(nvl) or date.today()
                                    if lhd == "Thử việc":
                                        ttnv, ttbh, tbd_val = 'THU_VIEC', 'CHUA_DONG', None
                                        c.execute("""
                                            SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) 
                                            FROM nhan_vien 
                                            WHERE so_hdld LIKE '%/HĐTV-CHL'
                                            AND trang_thai IN ('THU_VIEC', 'DANG_LAM')
                                        """)
                                        tv_cnt = c.fetchone()[0] + 1
                                        so_hd = f"{tv_cnt:02d}/{nhl.year}/HĐTV-CHL"
                                    else:
                                        ttnv, ttbh = 'DANG_LAM', 'DANG_DONG'
                                        tbd_val = parse_date(tbd) or parse_date(nvl)
                                        c.execute("""
                                            SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) 
                                            FROM nhan_vien 
                                            WHERE so_hdld LIKE '%/HĐLĐ-CHL'
                                            AND trang_thai = 'DANG_LAM'
                                            AND loai_hop_dong != 'Thử việc'
                                        """)
                                        so_hd_cnt = c.fetchone()[0] or 0
                                        so_hd = f"{so_hd_cnt + 1:02d}/{nhl.year}/HĐLĐ-CHL"
                                    c.execute("""INSERT INTO nhan_vien (STT, ma_nv, so_hdld, ho_ten, chuc_danh_nghe, ngay_sinh, gioi_tinh,
                                    so_cccd, ngay_cap_cccd, noi_cap_cccd, nguyen_quan, thuong_tru,
                                    dien_thoai, email, email_lien_he, ho_so, luong_bao_hiem, ma_so_bhxh, ngay_vao_lam,
                                    noi_lam_viec, so_tai_khoan_nh, chi_nhanh_nh, ngay_ky_hd, loai_hop_dong,
                                    nhom_bhxh, thang_bat_dau_bh, thang_ket_thuc_bh, trang_thai, trang_thai_bhxh,
                                    phong_ban_lam_viec, ngay_ket_thuc, quoc_tich, dan_toc, he_so_luong, phu_cap_chuc_vu,
                                    phu_cap_tnvk, phu_cap_tnn, muc_huong_bhyt, ty_le_dong, muc_tien_dong, phuong_thuc_dong,
                                    tinh_nhan_hs, phuong_nhan_hs, dia_chi_nhan_hs, tinh_kcb, noi_dang_ky_kcb, dang_ky_nhan_so)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                    %s, %s, %s, %s, %s, %s)""",
                                      (stt_moi, ma_nv, so_hd, htn, cdn, parse_date(nsn), gtn, scc, parse_date(ncc), ncc2, nqn, ttn,
                                       dtn2, emn, emn, '', lbh, mbh, parse_date(nvl), nlv, stk, cnh, parse_date(nvl), lhd,
                                       nbh, tbd_val, None, ttnv, ttbh, pbn, parse_date(nkt), qtn, dtn, 
                                       to_float_or_none(hsl),   # he_so_luong
                                       to_float_or_none(pcv),   # phu_cap_chuc_vu
                                       to_float_or_none(ptv),   # phu_cap_tnvk
                                       to_float_or_none(ptn),   # phu_cap_tnn
                                       mhb, 
                                       to_float_or_none(tld),   # ty_le_dong
                                       to_float_or_none(mtd),   # muc_tien_dong
                                       ptd, ths, phs, dhs, tkb, nkb, dks))
                                    db.commit()
                                    db.close()
                                    st.success(f"✅ Đã lưu nhân viên mới thành công! {htn} - {ma_nv}")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Lỗi: {e}")
                        else:
                            st.error("Họ tên không được để trống!")
                st.divider()
        
        db_f = get_connection()
        c_f = db_f.cursor()
        c_f.execute("SELECT DISTINCT chuc_danh_nghe FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != '' ORDER BY chuc_danh_nghe")
        ds_chuc_danh = [row[0] for row in c_f.fetchall()]
        c_f.execute("SELECT DISTINCT loai_hop_dong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND loai_hop_dong IS NOT NULL AND loai_hop_dong != '' ORDER BY loai_hop_dong")
        ds_loai_hd = [row[0] for row in c_f.fetchall()]
        db_f.close()
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filter_chuc_danh = st.selectbox("🔍 Lọc Chức danh:", ["Tất cả"] + ds_chuc_danh, key="filter_cd_danglam")
        with col_f2:
            filter_loai_hd = st.selectbox("🔍 Lọc Loại HĐ:", ["Tất cả"] + ds_loai_hd, key="filter_lhd_danglam")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        sql = "SELECT * FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC')"
        params = []
        if sn:
            sql += " AND (ho_ten LIKE %s OR dien_thoai LIKE %s OR so_cccd LIKE %s OR ma_nv LIKE %s)"
            params.extend([f'%{sn}%'] * 4)
        if filter_chuc_danh != "Tất cả":
            sql += " AND chuc_danh_nghe = %s"
            params.append(filter_chuc_danh)
        if filter_loai_hd != "Tất cả":
            sql += " AND loai_hop_dong = %s"
            params.append(filter_loai_hd)
        sql += " ORDER BY id DESC"
        c.execute(sql, tuple(params))
        ds = c.fetchall()
        db.close()
        
        if ds:
            df = pd.DataFrame(ds)
            for col in df.columns:
                if 'ngay' in col.lower():
                    df[col] = df[col].apply(format_date)
            
            if 'selected' not in df.columns:
                df.insert(0, 'selected', False)
            
            display_cols = ['selected', 'ma_nv', 'ho_ten', 'ngay_sinh', 'gioi_tinh', 'so_hdld', 'so_cccd', 'dien_thoai',
                            'thuong_tru', 'chuc_danh_nghe', 'loai_hop_dong', 'ngay_vao_lam', 'ma_so_bhxh', 'thang_bat_dau_bh']
            available_cols = [c for c in display_cols if c in df.columns]
            df_show = df[available_cols]
            
            col_map = {
                'selected': 'Chọn',
                'ma_nv': 'Mã NV',
                'ho_ten': 'Họ và tên',
                'ngay_sinh': 'Ngày sinh',
                'gioi_tinh': 'Giới tính',
                'so_hdld': 'Số HĐLĐ',
                'so_cccd': 'CCCD',
                'dien_thoai': 'SĐT',
                'thuong_tru': 'Thường trú',
                'chuc_danh_nghe': 'Chức danh',
                'loai_hop_dong': 'Loại HĐ',
                'ngay_vao_lam': 'Ngày vào làm',
                'ma_so_bhxh': 'Mã số BHXH',
                'thang_bat_dau_bh': 'Bắt đầu BH',
            }
            df_show.rename(columns=col_map, inplace=True)
            
            st.caption(f"📌 {len(ds)} kết quả. Tick chọn 1 nhân viên để thao tác.")
            
            # Nếu là viewer, hiển thị bảng không có checkbox chọn
            if st.session_state.role == "admin":
                edited_df = st.data_editor(
                    df_show,
                    column_config={
                        "Chọn": st.column_config.CheckboxColumn("Chọn để sửa", default=False)
                    },
                    disabled=[col for col in df_show.columns if col != 'Chọn'],
                    hide_index=True,
                    height=400,
                    key="nv_editor_danglam"
                )
            else:
                # Viewer: hiển thị bảng đơn thuần, không có checkbox
                st.dataframe(df_show.drop(columns=['Chọn'], errors='ignore'), width='stretch', hide_index=True, height=400)
                edited_df = None
            
            selected_nv = None
            if edited_df is not None and st.session_state.role == "admin" and 'Chọn' in edited_df.columns:
                selected_rows = edited_df[edited_df['Chọn'] == True]
                if len(selected_rows) > 0:
                    if len(selected_rows) > 1:
                        st.error("⚠️ Chỉ được chọn 1 nhân viên!")
                    else:
                        selected_idx = selected_rows.index[0]
                        selected_nv = df.iloc[selected_idx]
                        nv_id_key = selected_nv['id']
                        
                        # Hiển thị các nút chức năng (chỉ admin mới thấy và mới click được)
                        col_btn1, col_btn2, col_btn3, col_btn4, col_btn5 = st.columns(5)
                        
                        with col_btn1:
                            if st.button(f"✏️ SỬA '{selected_nv['ho_ten']}'", key=f"edit_nv_btn_{nv_id_key}", width='stretch'):
                                st.session_state['selected_nv_id'] = int(selected_nv['id'])
                                st.rerun()
                        
                        with col_btn2:
                            trang_thai_nv = selected_nv.get('trang_thai', '')
                            if trang_thai_nv == 'DANG_LAM':
                                if st.button(f"🖨️ IN HĐLĐ - {selected_nv['ho_ten']}", key=f"print_hdld_btn_{nv_id_key}", width='stretch'):
                                    db = get_connection()
                                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                    c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                    nv_full = c.fetchone()
                                    db.close()
                                    if nv_full:
                                        fp = tao_hop_dong(nv_full)
                                        with open(fp, "rb") as f:
                                            st.download_button(
                                                label="📥 TẢI HĐLĐ",
                                                data=f,
                                                file_name=f"HDLD_{nv_full['ho_ten']}_{datetime.now().strftime('%Y%m%d')}.docx",
                                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                                key=f"download_hdld_{nv_id_key}"
                                            )
                                    else:
                                        st.error("Không tìm thấy thông tin nhân viên!")
                            elif trang_thai_nv == 'THU_VIEC':
                                if st.button(f"🖨️ IN HĐTV - {selected_nv['ho_ten']}", key=f"print_hdtv_btn_{nv_id_key}", width='stretch'):
                                    db = get_connection()
                                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                    c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                    nv_full = c.fetchone()
                                    db.close()
                                    if nv_full:
                                        fp = tao_hop_dong_thu_viec(nv_full)
                                        with open(fp, "rb") as f:
                                            st.download_button(
                                                label="📥 TẢI HĐTV",
                                                data=f,
                                                file_name=f"HDTV_{nv_full['ho_ten']}_{datetime.now().strftime('%Y%m%d')}.docx",
                                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                                key=f"download_hdtv_{nv_id_key}"
                                            )
                                    else:
                                        st.error("Không tìm thấy thông tin nhân viên!")
                            else:
                                st.button(f"📄 {selected_nv['ho_ten']} (Không thể in HĐ)", disabled=True, width='stretch', key=f"disabled_btn_{nv_id_key}")
                        
                        with col_btn3:
                            if st.button(f"📱 GỬI ZALO - {selected_nv['ho_ten']}", key=f"zalo_btn_{nv_id_key}", width='stretch'):
                                db = get_connection()
                                c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                nv_full = c.fetchone()
                                db.close()
                                if nv_full:
                                    ph = nv_full.get('dien_thoai', '')
                                    if ph:
                                        ph = ph.replace('+84', '0').replace(' ', '').strip()
                                        st.code(tao_noi_dung_zalo(nv_full))
                                        st.markdown(f"[👉 MỞ ZALO](https://zalo.me/{ph})")
                                    else:
                                        st.error("Chưa có SĐT!")
                                else:
                                    st.error("Không tìm thấy thông tin nhân viên!")
                        
                        with col_btn4:
                            ma_bhxh = selected_nv.get('ma_so_bhxh', '')
                            chua_co_bhxh = not bool(ma_bhxh and str(ma_bhxh).strip())
                            if chua_co_bhxh:
                                if st.button(f"🏠 NHẬP THÔNG TIN HỘ GIA ĐÌNH - {selected_nv['ho_ten']}", key=f"bhxh_family_btn_{nv_id_key}", width='stretch', type="primary"):
                                    st.session_state['bhxh_family_nv_id'] = int(selected_nv['id'])
                                    st.session_state['bhxh_family_nv_name'] = selected_nv['ho_ten']
                                    st.rerun()
                            else:
                                st.button(f"✅ ĐÃ CÓ BHXH - {selected_nv['ho_ten']}", disabled=True, width='stretch', key=f"has_bhxh_btn_{nv_id_key}")
                        
                        with col_btn5:
                            trang_thai_nv = selected_nv.get('trang_thai', '')
                            if trang_thai_nv == 'THU_VIEC':
                                if f'convert_open_{nv_id_key}' not in st.session_state:
                                    st.session_state[f'convert_open_{nv_id_key}'] = False
                                
                                if not st.session_state[f'convert_open_{nv_id_key}']:
                                    if st.button(f"🔄 CHUYỂN HĐLĐ KHÔNG XĐTH - {selected_nv['ho_ten']}", 
                                                key=f"convert_hdld_btn_{nv_id_key}", 
                                                width='stretch', type="primary"):
                                        st.session_state[f'convert_open_{nv_id_key}'] = True
                                        st.rerun()
                                else:
                                    st.markdown("---")
                                    st.markdown("### 📝 CHUYỂN ĐỔI HỢP ĐỒNG LAO ĐỘNG")
                                    st.caption("Vui lòng nhập đầy đủ thông tin cho quyết định chuyển đổi")
                                    
                                    db_temp = get_connection()
                                    c_temp = db_temp.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                    c_temp.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                    nv_data = c_temp.fetchone()
                                    db_temp.close()
                                    
                                    if nv_data:
                                        ngay_quyet_dinh = st.date_input(
                                            "📅 Ngày quyết định:", 
                                            value=date.today(),
                                            key=f"ngay_qd_{nv_id_key}"
                                        )
                                        
                                        ngay_hieu_luc = st.date_input(
                                            "📅 Ngày hiệu lực (bắt đầu HĐLĐ):", 
                                            value=ngay_quyet_dinh,
                                            key=f"ngay_hl_{nv_id_key}"
                                        )
                                        
                                        # Trong phần chuyển đổi từ THU_VIEC sang DANG_LAM
                                        current_year = datetime.now().year

                                        db_temp2 = get_connection()
                                        c_temp2 = db_temp2.cursor()
                                        c_temp2.execute("""
                                            SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) as max_stt
                                            FROM nhan_vien 
                                            WHERE so_hdld LIKE %s 
                                            AND trang_thai = 'DANG_LAM'
                                            AND loai_hop_dong != 'Thử việc'
                                        """, (f'%/{current_year}/HĐLĐ-CHL',))
                                        result = c_temp2.fetchone()
                                        max_stt = result[0] if result else 0
                                        db_temp2.close()

                                        next_stt = max_stt + 1
                                        stt_str = str(next_stt).zfill(2)
                                        so_hd_moi = f"{stt_str}/{current_year}/HĐLĐ-CHL"
                                        
                                        st.info(f"📄 **Số HĐLĐ mới:** {so_hd_moi} (tự động sinh)")
                                        
                                        ngay_bat_dau_bh = st.date_input(
                                            "📅 Bắt đầu đóng BHXH:", 
                                            value=ngay_hieu_luc,
                                            key=f"ngay_bhxh_{nv_id_key}"
                                        )
                                        st.caption("⚠️ Đây là ngày bắt đầu tính đóng BHXH")
                                        
                                        ly_do_chuyen = st.text_area(
                                            "📝 Lý do/ Nội dung quyết định:", 
                                            value="Hoàn thành thời gian thử việc, chuyển sang hợp đồng lao động không xác định thời hạn",
                                            key=f"ly_do_{nv_id_key}",
                                            height=80
                                        )
                                        
                                        col_confirm1, col_confirm2, col_confirm3 = st.columns([1, 2, 1])
                                        with col_confirm2:
                                            if st.button("✅ XÁC NHẬN CHUYỂN ĐỔI", key=f"confirm_convert_{nv_id_key}", width='stretch', type="primary"):
                                                try:
                                                    db = get_connection()
                                                    c = db.cursor()
                                                    
                                                    # Lưu số HĐ cũ (HĐTV) vào biến để sau này rollback
                                                    so_hd_cu = selected_nv.get('so_hdld', '')
                                                    ngay_vao_lam_cu = selected_nv.get('ngay_vao_lam')
                                                    # Dùng hàm parse_date có sẵn để xử lý nhiều định dạng
                                                    if ngay_vao_lam_cu:
                                                        if hasattr(ngay_vao_lam_cu, 'strftime'):
                                                            # Đã là date object
                                                            pass
                                                        else:
                                                            ngay_vao_lam_cu = parse_date(ngay_vao_lam_cu)
                                                            if not ngay_vao_lam_cu:
                                                                ngay_vao_lam_cu = date.today()
                                                    else:
                                                        ngay_vao_lam_cu = date.today()
                                                    
                                                    # Cập nhật nhân viên sang HĐLĐ không xác định thời hạn
                                                    c.execute("""
                                                        UPDATE nhan_vien SET 
                                                            trang_thai = 'DANG_LAM',
                                                            loai_hop_dong = 'Không xác định thời hạn',
                                                            so_hdld = %s,
                                                            ngay_ky_hd = %s,
                                                            ngay_chinh_thuc = %s,
                                                            thang_bat_dau_bh = %s,
                                                            trang_thai_bhxh = 'DANG_DONG',
                                                            ngay_ket_thuc = NULL
                                                        WHERE id = %s
                                                    """, (so_hd_moi, ngay_quyet_dinh, ngay_hieu_luc, ngay_bat_dau_bh, int(selected_nv['id'])))
                                                    
                                                    # Lưu quyết định chuyển đổi vào bảng quyet_dinh_nhan_su
                                                    c.execute("""
                                                        INSERT INTO quyet_dinh_nhan_su (
                                                            nhan_vien_id, loai_quyet_dinh, ngay_quyet_dinh, ngay_hieu_luc,
                                                            noi_dung, so_quyet_dinh, loai_hop_dong_cu, loai_hop_dong_moi,
                                                            he_so_luong_cu, he_so_luong_moi
                                                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                                    """, (
                                                        int(selected_nv['id']),
                                                        'CHINH_THUC',
                                                        ngay_quyet_dinh,
                                                        ngay_hieu_luc,
                                                        ly_do_chuyen,
                                                        f"QD{ngay_quyet_dinh.strftime('%Y%m%d')}_{selected_nv['ma_nv']}",
                                                        nv_data.get('loai_hop_dong', 'Thử việc'),
                                                        'Không xác định thời hạn',
                                                        nv_data.get('he_so_luong', 0),
                                                        nv_data.get('he_so_luong', 0)
                                                    ))
                                                    
                                                    # Cập nhật lịch sử công tác: kết thúc giai đoạn thử việc
                                                    c.execute("""
                                                        UPDATE lich_su_cong_tac 
                                                        SET den_ngay = %s,
                                                            so_hop_dong = %s
                                                        WHERE nhan_vien_id = %s 
                                                        AND loai_hop_dong = 'Thử việc'
                                                        AND den_ngay IS NULL
                                                    """, (ngay_hieu_luc - timedelta(days=1), so_hd_cu, int(selected_nv['id'])))
                                                    
                                                    # Thêm lịch sử công tác mới cho giai đoạn chính thức
                                                    c.execute("""
                                                        INSERT INTO lich_su_cong_tac (
                                                            nhan_vien_id, tu_ngay, chuc_danh, phong_ban, 
                                                            noi_lam_viec, loai_hop_dong, he_so_luong, so_hop_dong
                                                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                                                    """, (
                                                        int(selected_nv['id']),
                                                        ngay_hieu_luc,
                                                        nv_data.get('chuc_danh_nghe', ''),
                                                        nv_data.get('phong_ban_lam_viec', ''),
                                                        nv_data.get('noi_lam_viec', 'Cảng THQT Hòn La'),
                                                        'Không xác định thời hạn',
                                                        nv_data.get('he_so_luong', 0),
                                                        so_hd_moi
                                                    ))
                                                    
                                                    db.commit()
                                                    db.close()
                                                    
                                                    st.success(f"✅ Đã chuyển {nv_data['ho_ten']} sang HĐLĐ không xác định thời hạn!")
                                                    st.info(f"📄 Số HĐLĐ mới: {so_hd_moi}")
                                                    st.info(f"📅 Ngày hiệu lực: {ngay_hieu_luc.strftime('%d/%m/%Y')}")
                                                    st.info(f"💰 Bắt đầu đóng BHXH từ: {ngay_bat_dau_bh.strftime('%d/%m/%Y')}")
                                                    
                                                    st.session_state[f'convert_open_{nv_id_key}'] = False
                                                    st.rerun()
                                                except Exception as e:
                                                    db.rollback()
                                                    db.close()
                                                    st.error(f"❌ Lỗi: {str(e)}")

                                        if st.button("❌ HỦY", key=f"cancel_convert_{nv_id_key}", width='stretch'):
                                            st.session_state[f'convert_open_{nv_id_key}'] = False
                                            st.rerun()
                            else:
                                st.button(f"✅ ĐÃ LÀ HĐLĐ", disabled=True, width='stretch', key=f"already_hdld_btn_{nv_id_key}")
                        
                        st.divider()
            
            # Form sửa nhân viên (chỉ admin)
            if 'selected_nv_id' in st.session_state and st.session_state.selected_nv_id is not None and st.session_state.role == "admin":
                try:
                    nid = int(st.session_state['selected_nv_id'])
                    db = get_connection()
                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c.execute("SELECT * FROM nhan_vien WHERE id=%s", (nid,))
                    nd = c.fetchone()
                    db.close()
                    
                    if nd:
                        st.subheader(f"✏️ Cập nhật: {nd.get('ho_ten', '')} ({nd.get('ma_nv', '')})")
                        with st.form("edit_nv"):
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                hnv = st.text_input("Họ tên *", value=nd.get('ho_ten', ''))
                                nsnv = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(nd.get('ngay_sinh')), placeholder="dd/mm/yyyy", max_chars=10)
                                gtnv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"], index=["", "Nam", "Nữ", "Khác"].index(nd.get('gioi_tinh', '')) if nd.get('gioi_tinh') in ["Nam", "Nữ", "Khác"] else 0)
                                qtnv = st.text_input("Quốc tịch", value=nd.get('quoc_tich', 'Việt Nam'))
                                dtnv = st.text_input("Dân tộc", value=nd.get('dan_toc', 'Kinh'))
                            with col2:
                                sccv = st.text_input("CCCD", value=nd.get('so_cccd', ''))
                                nccv = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", value=format_date(nd.get('ngay_cap_cccd')), placeholder="dd/mm/yyyy", max_chars=10)
                                ncv = st.text_input("Nơi cấp CCCD", value=nd.get('noi_cap_cccd', ''))
                                nqnv = st.text_input("Nguyên quán", value=nd.get('nguyen_quan', ''))
                                ttnv = st.text_input("Thường trú", value=nd.get('thuong_tru', ''))
                            with col3:
                                dtnv2 = st.text_input("SĐT", value=nd.get('dien_thoai', ''))
                                emnv = st.text_input("Email", value=nd.get('email_lien_he', ''))
                                cdnv = st.text_input("Chức danh", value=nd.get('chuc_danh_nghe', ''))
                                pbnv = st.text_input("Phòng ban", value=nd.get('phong_ban_lam_viec', ''))
                                nlv2 = st.text_input("Nơi làm việc", value=nd.get('noi_lam_viec', 'Cảng THQT Hòn La'))
                            
                            st.divider()
                            st.caption("💼 Hợp đồng & BHXH")
                            col4, col5, col6 = st.columns(3)
                            with col4:
                                lhdv = st.selectbox("Loại HĐ", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"], index=["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"].index(nd.get('loai_hop_dong', 'Thử việc')) if nd.get('loai_hop_dong') in ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"] else 0)
                                nvlv = st.text_input("Ngày vào làm (dd/mm/yyyy)", value=format_date(nd.get('ngay_vao_lam')), placeholder="dd/mm/yyyy", max_chars=10)
                                nktv = st.text_input("Ngày kết thúc (dd/mm/yyyy)", value=format_date(nd.get('ngay_ket_thuc')), placeholder="dd/mm/yyyy", max_chars=10)
                                mbhv = st.text_input("Mã BHXH", value=nd.get('ma_so_bhxh', ''))
                                tbdv = st.text_input("Bắt đầu BH (dd/mm/yyyy)", value=format_date(nd.get('thang_bat_dau_bh')), placeholder="dd/mm/yyyy", max_chars=10)
                            with col5:
                                lbhv = st.text_input("Lương BH", value=nd.get('luong_bao_hiem', ''))
                                hslv = st.text_input("Hệ số lương", value=str(nd.get('he_so_luong', '')))
                                pcvv = st.text_input("PC chức vụ", value=str(nd.get('phu_cap_chuc_vu', '')))
                                ptvv = st.text_input("PC TNVK (%)", value=str(nd.get('phu_cap_tnvk', '')))
                                ptnv = st.text_input("PC TNN (%)", value=str(nd.get('phu_cap_tnn', '')))
                            with col6:
                                mhbv = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"], index=["80%", "95%", "100%"].index(nd.get('muc_huong_bhyt', '80%')) if nd.get('muc_huong_bhyt') in ["80%", "95%", "100%"] else 0)
                                tldv = st.text_input("Tỷ lệ đóng (%)", value=str(nd.get('ty_le_dong', '')))
                                mtdv = st.text_input("Mức tiền đóng", value=str(nd.get('muc_tien_dong', '')))
                                ptdv = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"], index=["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"].index(nd.get('phuong_thuc_dong', 'Hàng tháng')) if nd.get('phuong_thuc_dong') in ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"] else 0)
                                nbhv = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"], index=["", "Văn phòng", "Lao động trực tiếp"].index(nd.get('nhom_bhxh', '')) if nd.get('nhom_bhxh') in ["Văn phòng", "Lao động trực tiếp"] else 0)
                            
                            st.divider()
                            st.caption("🏦 Ngân hàng & KCB")
                            col7, col8, col9 = st.columns(3)
                            with col7:
                                stkv = st.text_input("STK", value=nd.get('so_tai_khoan_nh', ''))
                                cnhv = st.text_input("Chi nhánh NH", value=nd.get('chi_nhanh_nh', ''))
                                tkbv = st.text_input("Tỉnh KCB", value=nd.get('tinh_kcb', ''))
                                nkbv = st.text_input("Nơi KCB", value=nd.get('noi_dang_ky_kcb', ''))
                            with col8:
                                thsv = st.text_input("Tỉnh/TP nhận HS", value=nd.get('tinh_nhan_hs', ''))
                                phsv = st.text_input("Phường/Xã nhận HS", value=nd.get('phuong_nhan_hs', ''))
                                dhsv = st.text_area("Địa chỉ nhận HS", value=nd.get('dia_chi_nhan_hs', ''), height=100)
                            with col9:
                                dksv = st.selectbox("ĐK nhận sổ", ["Có", "Không"], index=["Có", "Không"].index(nd.get('dang_ky_nhan_so', 'Có')) if nd.get('dang_ky_nhan_so') in ["Có", "Không"] else 0)
                                hsov = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"], index=["", "Đã có HS", "Chưa có"].index(nd.get('ho_so', '')) if nd.get('ho_so') in ["Đã có HS", "Chưa có"] else 0)
                            
                            col_save, col_quit, col_delete = st.columns(3)
                            with col_save:
                                if st.form_submit_button("💾 CẬP NHẬT", width='stretch'):
                                    if hnv:
                                        ngay_loi = []
                                        if nsnv and not parse_date(nsnv):
                                            ngay_loi.append("Ngày sinh")
                                        if nccv and not parse_date(nccv):
                                            ngay_loi.append("Ngày cấp CCCD")
                                        if nvlv and not parse_date(nvlv):
                                            ngay_loi.append("Ngày vào làm")
                                        if nktv and not parse_date(nktv):
                                            ngay_loi.append("Ngày kết thúc")
                                        if tbdv and not parse_date(tbdv):
                                            ngay_loi.append("Bắt đầu BH")
                                        if ngay_loi:
                                            st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                                        else:
                                            try:
                                                db_upd = get_connection()
                                                c_upd = db_upd.cursor()
                                                nhl = parse_date(nvlv) or date.today()
                                                ngay_bat_dau_bh = parse_date(tbdv) if tbdv and tbdv.strip() else None
                                                if lhdv == "Thử việc":
                                                    tt_nv, tt_bh, tbd_val = 'THU_VIEC', 'CHUA_DONG', None
                                                else:
                                                    tt_nv, tt_bh = 'DANG_LAM', 'DANG_DONG'
                                                    if ngay_bat_dau_bh:
                                                        tbd_val = ngay_bat_dau_bh
                                                    else:
                                                        tbd_val = parse_date(nvlv)
                                                c_upd.execute("""UPDATE nhan_vien SET ho_ten=%s,chuc_danh_nghe=%s,ngay_sinh=%s,gioi_tinh=%s,
                                                    so_cccd=%s,ngay_cap_cccd=%s,noi_cap_cccd=%s,nguyen_quan=%s,thuong_tru=%s,dien_thoai=%s,
                                                    email=%s,email_lien_he=%s,ho_so=%s,luong_bao_hiem=%s,ma_so_bhxh=%s,ngay_vao_lam=%s,noi_lam_viec=%s,
                                                    so_tai_khoan_nh=%s,chi_nhanh_nh=%s,ngay_ky_hd=%s,loai_hop_dong=%s,nhom_bhxh=%s,
                                                    thang_bat_dau_bh=%s,trang_thai=%s,trang_thai_bhxh=%s,phong_ban_lam_viec=%s,
                                                    ngay_ket_thuc=%s,quoc_tich=%s,dan_toc=%s,he_so_luong=%s,phu_cap_chuc_vu=%s,
                                                    phu_cap_tnvk=%s,phu_cap_tnn=%s,muc_huong_bhyt=%s,ty_le_dong=%s,muc_tien_dong=%s,
                                                    phuong_thuc_dong=%s,tinh_nhan_hs=%s,phuong_nhan_hs=%s,dia_chi_nhan_hs=%s,
                                                    tinh_kcb=%s,noi_dang_ky_kcb=%s,dang_ky_nhan_so=%s WHERE id=%s""",
                                                      (hnv, cdnv, parse_date(nsnv), gtnv, sccv, parse_date(nccv), ncv, nqnv, ttnv, dtnv2,
                                                       emnv, emnv, hsov, lbhv, mbhv, parse_date(nvlv), nlv2, stkv, cnhv, parse_date(nvlv), lhdv,
                                                       nbhv, tbd_val, tt_nv, tt_bh, pbnv, parse_date(nktv), qtnv, dtnv,
                                                       to_float_or_none(hslv), to_float_or_none(pcvv), to_float_or_none(ptvv), to_float_or_none(ptnv),
                                                       mhbv, to_float_or_none(tldv), to_float_or_none(mtdv), ptdv, thsv, phsv, dhsv,
                                                       tkbv, nkbv, dksv, nid))
                                                db_upd.commit()
                                                db_upd.close()
                                                st.success(f"✅ Đã cập nhật: {hnv}")
                                                del st.session_state['selected_nv_id']
                                                st.rerun()
                                            except Exception as e:
                                                st.error(f"❌ Lỗi: {e}")
                                    else:
                                        st.error("Họ tên không được để trống!")
                            
                            with col_quit:
                                if f'nghi_viec_open_{nid}' not in st.session_state:
                                    st.session_state[f'nghi_viec_open_{nid}'] = False
                                
                                if not st.session_state[f'nghi_viec_open_{nid}']:
                                    if st.form_submit_button("🚫 NGHỈ VIỆC", width='stretch', type="secondary"):
                                        st.session_state[f'nghi_viec_open_{nid}'] = True
                                        st.rerun()
                                else:
                                    st.markdown("---")
                                    st.markdown("### 📝 XÁC NHẬN NGHỈ VIỆC")
                                    
                                    default_ngay_nghi = date.today()
                                    ngay_ket_thuc_hien_tai = nd.get('ngay_ket_thuc')
                                    if ngay_ket_thuc_hien_tai and ngay_ket_thuc_hien_tai != '':
                                        try:
                                            if hasattr(ngay_ket_thuc_hien_tai, 'strftime'):
                                                default_ngay_nghi = ngay_ket_thuc_hien_tai
                                        except:
                                            pass
                                    
                                    ngay_nghi = st.date_input(
                                        "📅 Ngày quyết định nghỉ việc (Ngày kết thúc HĐLĐ):", 
                                        value=default_ngay_nghi,
                                        key=f"ngay_nghi_{nid}"
                                    )
                                    
                                    ly_do_nghi = st.text_area(
                                        "📝 Lý do nghỉ việc:", 
                                        value=nd.get('ly_do_nghi', '') if 'ly_do_nghi' in nd else '',
                                        placeholder="VD: Xin nghỉ theo nguyện vọng cá nhân, Chuyển công tác, Hết hạn hợp đồng...",
                                        key=f"ly_do_nghi_{nid}",
                                        height=100
                                    )
                                    
                                    col_xac_nhan, col_huy = st.columns(2)
                                    
                                    with col_xac_nhan:
                                        if st.form_submit_button("✅ XÁC NHẬN NGHỈ VIỆC", width='stretch', type="primary"):
                                            try:
                                                db_nghi = get_connection()
                                                c_nghi = db_nghi.cursor()
                                                
                                                c_nghi.execute("""
                                                    UPDATE nhan_vien 
                                                    SET trang_thai = 'NGHI_VIEC', 
                                                        ngay_ket_thuc = %s,
                                                        ly_do_nghi = %s
                                                    WHERE id = %s
                                                """, (ngay_nghi, ly_do_nghi if ly_do_nghi else None, nid))
                                                
                                                db_nghi.commit()
                                                db_nghi.close()
                                                
                                                st.session_state[f'nghi_viec_open_{nid}'] = False
                                                if 'selected_nv_id' in st.session_state:
                                                    del st.session_state['selected_nv_id']
                                                
                                                st.success(f"✅ Đã cập nhật nghỉ việc cho {nd.get('ho_ten', '')}!")
                                                st.rerun()
                                            except Exception as e:
                                                st.error(f"❌ Lỗi khi cập nhật nghỉ việc: {e}")
                                    
                                    with col_huy:
                                        if st.form_submit_button("❌ HỦY", width='stretch'):
                                            st.session_state[f'nghi_viec_open_{nid}'] = False
                                            st.rerun()
                            
                            with col_delete:
                                # Kiểm tra xem nhân viên đã được chuyển đổi chưa
                                da_chuyen_doi, quyet_dinh = da_chuyen_doi_chinh_thuc(nid)
                                
                                if da_chuyen_doi:
                                    if st.form_submit_button("🔄 XÓA QUYẾT ĐỊNH CHUYỂN ĐỔI", width='stretch', type="secondary"):
                                        try:
                                            db_xoa = get_connection()
                                            c_xoa = db_xoa.cursor()
                                            
                                            # Lấy thông tin HĐTV cũ từ lịch sử
                                            c_xoa.execute("""
                                                SELECT so_hop_dong, tu_ngay FROM lich_su_cong_tac 
                                                WHERE nhan_vien_id = %s AND loai_hop_dong = 'Thử việc'
                                                ORDER BY tu_ngay ASC LIMIT 1
                                            """, (nid,))
                                            lich_su_cu = c_xoa.fetchone()
                                            
                                            if lich_su_cu and lich_su_cu[0]:
                                                so_hd_cu = lich_su_cu[0]
                                                ngay_bat_dau_tv = lich_su_cu[1]
                                                
                                                # Chỉ đổi ký hiệu từ HĐLĐ-CHL thành HĐTV-CHL, giữ nguyên số
                                                if '/HĐLĐ-CHL' in so_hd_cu:
                                                    so_hd_cu = so_hd_cu.replace('/HĐLĐ-CHL', '/HĐTV-CHL')
                                                
                                                # Cập nhật lại nhân viên về trạng thái Thử việc
                                                c_xoa.execute("""
                                                    UPDATE nhan_vien SET 
                                                        trang_thai = 'THU_VIEC',
                                                        loai_hop_dong = 'Thử việc',
                                                        so_hdld = %s,
                                                        trang_thai_bhxh = 'CHUA_DONG',
                                                        ngay_chinh_thuc = NULL,
                                                        thang_bat_dau_bh = NULL,
                                                        ngay_ky_hd = %s
                                                    WHERE id = %s
                                                """, (so_hd_cu, ngay_bat_dau_tv, nid))
                                                
                                                # Xóa quyết định chuyển đổi
                                                if quyet_dinh:
                                                    c_xoa.execute("DELETE FROM quyet_dinh_nhan_su WHERE id = %s", (quyet_dinh['id'],))
                                                
                                                # Xóa lịch sử HĐLĐ mới
                                                c_xoa.execute("""
                                                    DELETE FROM lich_su_cong_tac 
                                                    WHERE nhan_vien_id = %s AND loai_hop_dong = 'Không xác định thời hạn'
                                                """, (nid,))
                                                
                                                # Cập nhật lịch sử thử việc
                                                c_xoa.execute("""
                                                    UPDATE lich_su_cong_tac 
                                                    SET den_ngay = NULL,
                                                        so_hop_dong = %s
                                                    WHERE nhan_vien_id = %s AND loai_hop_dong = 'Thử việc'
                                                """, (so_hd_cu, nid))
                                                
                                                db_xoa.commit()
                                                db_xoa.close()
                                                
                                                st.success("✅ Đã xóa quyết định chuyển đổi! Nhân viên trở lại trạng thái Thử việc.")
                                                del st.session_state['selected_nv_id']
                                                st.rerun()
                                            else:
                                                st.error("❌ Không tìm thấy thông tin lịch sử để khôi phục!")
                                                
                                        except Exception as e:
                                            db_xoa.rollback()
                                            db_xoa.close()
                                            st.error(f"❌ Lỗi khi xóa quyết định chuyển đổi: {e}")
                                else:
                                    if st.form_submit_button("🗑️ XÓA NHÂN VIÊN", width='stretch', type="secondary"):
                                        try:
                                            db_xoa = get_connection()
                                            c_xoa = db_xoa.cursor()
                                            c_xoa.execute("DELETE FROM ho_so_nhan_vien WHERE nhan_vien_id=%s", (nid,))
                                            c_xoa.execute("DELETE FROM quyet_dinh_nhan_su WHERE nhan_vien_id=%s", (nid,))
                                            c_xoa.execute("DELETE FROM lich_su_cong_tac WHERE nhan_vien_id=%s", (nid,))
                                            c_xoa.execute("DELETE FROM nhan_vien WHERE id=%s", (nid,))
                                            db_xoa.commit()
                                            db_xoa.close()
                                            st.success("🗑️ Đã xóa nhân viên!")
                                            del st.session_state['selected_nv_id']
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi khi xóa: {e}")
                        
                        # Nút hủy sửa (đặt ngoài form)
                        if st.button("❌ HỦY SỬA", width='stretch'):
                            del st.session_state['selected_nv_id']
                            st.rerun()
                    
                    else:
                        st.error("Không tìm thấy thông tin nhân viên!")
                        del st.session_state['selected_nv_id']
                        st.rerun()
                
                except Exception as e:
                    st.error(f"Lỗi khi tải thông tin nhân viên: {e}")
                    if 'selected_nv_id' in st.session_state:
                        del st.session_state['selected_nv_id']
                    st.rerun()

            # Form nhập thông tin hộ gia đình (chỉ admin) - Đặt NGOÀI form sửa nhân viên
            if 'bhxh_family_nv_id' in st.session_state and st.session_state.bhxh_family_nv_id is not None and st.session_state.role == "admin":
                nv_id = st.session_state['bhxh_family_nv_id']
                nv_name = st.session_state['bhxh_family_nv_name']
                st.divider()
                st.subheader(f"🏠 NHẬP THÔNG TIN HỘ GIA ĐÌNH CHO: {nv_name}")
                st.caption("Vui lòng nhập đầy đủ thông tin chủ hộ và các thành viên trong hộ gia đình")
                
                db = get_connection()
                c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                c.execute("SELECT * FROM nhan_vien WHERE id = %s", (nv_id,))
                nv_data = c.fetchone()
                db.close()
                
                if 'bhxh_family_members' not in st.session_state:
                    db_temp = get_connection()
                    c_temp = db_temp.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c_temp.execute("SELECT * FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_id,))
                    existing_members = c_temp.fetchall()
                    db_temp.close()
                    st.session_state['bhxh_family_members'] = []
                    for tv in existing_members:
                        st.session_state['bhxh_family_members'].append({
                            'ho_ten': tv['ho_ten'], 'ngay_sinh': tv['ngay_sinh'], 'gioi_tinh': tv['gioi_tinh'],
                            'quoc_tich': tv['quoc_tich'], 'dan_toc': tv['dan_toc'], 'quan_he': tv['quan_he_voi_chu_ho'],
                            'tinh': tv['tinh_thanh_pho'], 'phuong_xa': tv['phuong_xa']
                        })
                
                db_temp = get_connection()
                c_temp = db_temp.cursor()
                c_temp.execute("SELECT ma_tinh, ten_tinh FROM danh_muc_tinh ORDER BY ten_tinh")
                ds_tinh = c_temp.fetchall()
                db_temp.close()
                tinh_options = {ten: ma for ma, ten in ds_tinh}
                
                if st.session_state.bhxh_family_members:
                    st.markdown("**Danh sách thành viên đã thêm:**")
                    tv_data = []
                    for i, tv in enumerate(st.session_state.bhxh_family_members):
                        tv_data.append({"STT": i+1, "Họ và tên": tv['ho_ten'], "Ngày sinh": format_date(tv['ngay_sinh']),
                                        "Giới tính": tv['gioi_tinh'], "Quốc tịch": tv['quoc_tich'], "Dân tộc": tv['dan_toc'],
                                        "Quan hệ chủ hộ": tv['quan_he'], "Tỉnh/TP": tv['tinh'], "Phường/Xã": tv['phuong_xa']})
                    df_tv = pd.DataFrame(tv_data)
                    st.dataframe(df_tv, width='stretch', hide_index=True)
                    col_del1, col_del2, col_del3 = st.columns([1,1,1])
                    with col_del2:
                        tv_to_delete = st.number_input("Nhập STT thành viên cần xóa:", min_value=1, max_value=len(st.session_state.bhxh_family_members), step=1, key="tv_delete_family")
                        if st.button("🗑️ Xóa thành viên", key="btn_del_tv_family"):
                            st.session_state.bhxh_family_members.pop(tv_to_delete - 1)
                            st.rerun()
                
                with st.form(key=f"bhxh_family_form_{nv_id}"):
                    st.markdown("**I. THÔNG TIN CHỦ HỘ:**")
                    col1, col2 = st.columns(2)
                    with col1:
                        ho_ten_chu_ho = st.text_input("Họ và tên chủ hộ", value=nv_data.get('ho_ten_chu_ho', '') if nv_data else '')
                        so_cccd_chu_ho = st.text_input("Số CCCD chủ hộ", value=nv_data.get('so_cccd_chu_ho', '') if nv_data else '')
                        tinh_chu_ho_index = 0
                        tinh_chu_ho_current = nv_data.get('tinh_thanh_pho_chu_ho', '') if nv_data else ''
                        if tinh_chu_ho_current in tinh_options:
                            tinh_chu_ho_index = list(tinh_options.keys()).index(tinh_chu_ho_current) + 1
                        tinh_chu_ho = st.selectbox("Tỉnh/Thành phố (chủ hộ)", options=[""] + list(tinh_options.keys()), index=tinh_chu_ho_index)
                    with col2:
                        phuong_xa_options = []
                        phuong_xa_current = nv_data.get('phuong_xa_chu_ho', '') if nv_data else ''
                        if tinh_chu_ho and tinh_chu_ho != "":
                            ma_tinh = tinh_options.get(tinh_chu_ho)
                            db_temp2 = get_connection()
                            c_temp2 = db_temp2.cursor()
                            c_temp2.execute("SELECT ten_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh,))
                            phuong_xa_options = [row[0] for row in c_temp2.fetchall()]
                            db_temp2.close()
                        phuong_xa_index = 0
                        if phuong_xa_current in phuong_xa_options:
                            phuong_xa_index = phuong_xa_options.index(phuong_xa_current) + 1
                        phuong_xa_chu_ho = st.selectbox("Phường/Xã (chủ hộ)", options=[""] + phuong_xa_options, index=phuong_xa_index)
                    
                    st.markdown("**Thông tin thường trú:**")
                    col_tt1, col_tt2 = st.columns(2)
                    with col_tt1:
                        tinh_thuong_tru_index = 0
                        tinh_thuong_tru_current = nv_data.get('tinh_thanh_pho_thuong_tru', '') if nv_data else ''
                        if tinh_thuong_tru_current in tinh_options:
                            tinh_thuong_tru_index = list(tinh_options.keys()).index(tinh_thuong_tru_current) + 1
                        tinh_thuong_tru = st.selectbox("Tỉnh/Thành phố thường trú", options=[""] + list(tinh_options.keys()), index=tinh_thuong_tru_index)
                        ma_tinh_thuong_tru = tinh_options.get(tinh_thuong_tru, "") if tinh_thuong_tru else ""
                    with col_tt2:
                        phuong_xa_tt_options = []
                        phuong_xa_tt_current = nv_data.get('phuong_xa_thuong_tru', '') if nv_data else ''
                        if tinh_thuong_tru and tinh_thuong_tru != "":
                            ma_tinh_tt = tinh_options.get(tinh_thuong_tru)
                            db_temp3 = get_connection()
                            c_temp3 = db_temp3.cursor()
                            c_temp3.execute("SELECT ten_xa, ma_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh_tt,))
                            phuong_xa_tt_options = c_temp3.fetchall()
                            db_temp3.close()
                        phuong_xa_tt_index = 0
                        ma_phuong_xa_thuong_tru = ""
                        for i, px in enumerate(phuong_xa_tt_options):
                            if px[0] == phuong_xa_tt_current:
                                phuong_xa_tt_index = i + 1
                                ma_phuong_xa_thuong_tru = px[1]
                                break
                        phuong_xa_display_list = [""] + [px[0] for px in phuong_xa_tt_options]
                        phuong_xa_thuong_tru = st.selectbox("Phường/Xã thường trú", options=phuong_xa_display_list, index=phuong_xa_tt_index)
                        for px in phuong_xa_tt_options:
                            if px[0] == phuong_xa_thuong_tru:
                                ma_phuong_xa_thuong_tru = px[1]
                                break
                    
                    st.markdown("**II. THÊM THÀNH VIÊN MỚI:**")
                    st.caption("Điền thông tin vào các cột bên dưới, sau đó bấm '➕ Thêm thành viên'")
                    col_tv1, col_tv2, col_tv3, col_tv4, col_tv5, col_tv6, col_tv7, col_tv8 = st.columns([2,1.3,1,1,1,1.5,1.8,1.8])
                    with col_tv1:
                        ho_ten_tv = st.text_input("Họ và tên", key="tv_ho_ten_family", placeholder="Nguyễn Văn A")
                    with col_tv2:
                        ngay_sinh_tv = st.text_input("Ngày sinh", key="tv_ngay_sinh_family", placeholder="dd/mm/yyyy")
                    with col_tv3:
                        gioi_tinh_tv = st.selectbox("Giới tính", ["Nam", "Nữ"], key="tv_gioi_tinh_family")
                    with col_tv4:
                        quoc_tich_tv = st.text_input("Quốc tịch", value="Việt Nam", key="tv_quoc_tich_family")
                    with col_tv5:
                        dan_toc_tv = st.text_input("Dân tộc", value="Kinh", key="tv_dan_toc_family")
                    with col_tv6:
                        quan_he_tv = st.selectbox("Quan hệ chủ hộ", ["", "Vợ", "Chồng", "Con", "Bố", "Mẹ", "Anh", "Chị", "Em", "Ông", "Bà", "Khác"], key="tv_quan_he_family")
                    with col_tv7:
                        tinh_tv = st.selectbox("Tỉnh/Thành phố", options=[""] + list(tinh_options.keys()), key="tv_tinh_family")
                    with col_tv8:
                        phuong_xa_tv_options = []
                        if tinh_tv and tinh_tv != "":
                            ma_tinh_tv = tinh_options.get(tinh_tv)
                            db_temp4 = get_connection()
                            c_temp4 = db_temp4.cursor()
                            c_temp4.execute("SELECT ten_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh_tv,))
                            phuong_xa_tv_options = [row[0] for row in c_temp4.fetchall()]
                            db_temp4.close()
                        phuong_xa_tv = st.selectbox("Phường/Xã", options=[""] + phuong_xa_tv_options, key="tv_phuong_xa_family")
                    
                    col_btn_add1, col_btn_add2, col_btn_add3 = st.columns([1,1,1])
                    with col_btn_add2:
                        if st.form_submit_button("➕ Thêm thành viên vào danh sách", width='stretch'):
                            if ho_ten_tv:
                                st.session_state.bhxh_family_members.append({
                                    'ho_ten': ho_ten_tv, 'ngay_sinh': parse_date(ngay_sinh_tv), 'gioi_tinh': gioi_tinh_tv,
                                    'quoc_tich': quoc_tich_tv, 'dan_toc': dan_toc_tv, 'quan_he': quan_he_tv,
                                    'tinh': tinh_tv, 'phuong_xa': phuong_xa_tv
                                })
                                st.rerun()
                            else:
                                st.error("Vui lòng nhập họ tên thành viên")
                    
                    st.markdown("---")
                    col_save1, col_save2, col_save3 = st.columns([1,2,1])
                    with col_save2:
                        if st.form_submit_button("💾 LƯU THÔNG TIN CHỦ HỘ", width='stretch', type="primary"):
                            try:
                                db_luu = get_connection()
                                c_luu = db_luu.cursor()
                                c_luu.execute("""UPDATE nhan_vien SET ho_ten_chu_ho=%s, so_cccd_chu_ho=%s, tinh_thanh_pho_chu_ho=%s, phuong_xa_chu_ho=%s,
                                    tinh_thanh_pho_thuong_tru=%s, ma_tinh_thuong_tru=%s, phuong_xa_thuong_tru=%s, ma_phuong_xa_thuong_tru=%s WHERE id=%s""",
                                    (ho_ten_chu_ho, so_cccd_chu_ho, tinh_chu_ho, phuong_xa_chu_ho, tinh_thuong_tru, ma_tinh_thuong_tru, phuong_xa_thuong_tru, ma_phuong_xa_thuong_tru, nv_id))
                                c_luu.execute("DELETE FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_id,))
                                for tv in st.session_state.bhxh_family_members:
                                    c_luu.execute("""INSERT INTO phu_luc_gia_dinh (nhan_vien_id, ho_ten, ngay_sinh, gioi_tinh, quoc_tich, dan_toc, quan_he_voi_chu_ho, tinh_thanh_pho, phuong_xa) 
                                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                                        (nv_id, tv['ho_ten'], tv['ngay_sinh'], tv['gioi_tinh'], tv['quoc_tich'], tv['dan_toc'], tv['quan_he'], tv['tinh'], tv['phuong_xa']))
                                db_luu.commit()
                                db_luu.close()
                                del st.session_state['bhxh_family_nv_id']
                                del st.session_state['bhxh_family_nv_name']
                                del st.session_state['bhxh_family_members']
                                st.success(f"✅ Đã lưu thông tin hộ gia đình cho nhân viên {nv_name}")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi lưu: {e}")
                
                col_cancel1, col_cancel2, col_cancel3 = st.columns([1,2,1])
                with col_cancel2:
                    if st.button("❌ HỦY BỎ", width='stretch'):
                        del st.session_state['bhxh_family_nv_id']
                        del st.session_state['bhxh_family_nv_name']
                        if 'bhxh_family_members' in st.session_state:
                            del st.session_state['bhxh_family_members']
                        st.rerun()
    
    with tab_da_nghi:
        st.caption("📋 Danh sách nhân viên đã nghỉ việc (có thông tin ngày nghỉ)")
        
        col_filter1, col_filter2, col_filter3 = st.columns([2, 2, 1])
        with col_filter1:
            search_nghi = st.text_input("🔍 Tìm kiếm (Tên, Mã NV, SĐT, CCCD)", key="search_da_nghi")
        with col_filter2:
            db_temp = get_connection()
            c_temp = db_temp.cursor()
            c_temp.execute("SELECT DISTINCT EXTRACT(YEAR FROM ngay_ket_thuc) as nam FROM nhan_vien WHERE trang_thai='NGHI_VIEC' AND ngay_ket_thuc IS NOT NULL ORDER BY nam DESC")
            years = [int(row[0]) for row in c_temp.fetchall() if row[0] is not None]
            db_temp.close()
            filter_nam = st.selectbox("📅 Lọc theo năm nghỉ:", ["Tất cả"] + [str(y) for y in years] if years else ["Tất cả"])
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        sql = """
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, so_cccd, dien_thoai, 
                   chuc_danh_nghe, loai_hop_dong, so_hdld, ngay_vao_lam, ngay_ket_thuc,
                   ma_so_bhxh, thang_bat_dau_bh, ly_do_nghi
            FROM nhan_vien 
            WHERE trang_thai = 'NGHI_VIEC'
        """
        params = []
        
        if search_nghi:
            sql += " AND (ho_ten LIKE %s OR ma_nv LIKE %s OR dien_thoai LIKE %s OR so_cccd LIKE %s)"
            params.extend([f'%{search_nghi}%'] * 4)
        
        if filter_nam != "Tất cả" and filter_nam.isdigit():
            sql += " AND EXTRACT(YEAR FROM ngay_ket_thuc) = %s"
            params.append(int(filter_nam))
        
        sql += " ORDER BY ngay_ket_thuc DESC, id DESC"
        c.execute(sql, tuple(params))
        ds_nghi = c.fetchall()
        db.close()
        
        if ds_nghi:
            df_nghi = pd.DataFrame(ds_nghi)
            for col in df_nghi.columns:
                if 'ngay' in col.lower():
                    df_nghi[col] = df_nghi[col].apply(format_date)
            
            display_cols_nghi = ['ma_nv', 'ho_ten', 'ngay_sinh', 'gioi_tinh', 'chuc_danh_nghe', 
                                 'so_hdld', 'ngay_vao_lam', 'ngay_ket_thuc', 'dien_thoai', 'ma_so_bhxh']
            available_cols_nghi = [c for c in display_cols_nghi if c in df_nghi.columns]
            df_show_nghi = df_nghi[available_cols_nghi]
            
            col_map_nghi = {
                'ma_nv': 'Mã NV',
                'ho_ten': 'Họ và tên',
                'ngay_sinh': 'Ngày sinh',
                'gioi_tinh': 'Giới tính',
                'chuc_danh_nghe': 'Chức danh',
                'so_hdld': 'Số HĐLĐ',
                'ngay_vao_lam': 'Ngày vào làm',
                'ngay_ket_thuc': '📅 Ngày nghỉ việc',
                'dien_thoai': 'SĐT',
                'ma_so_bhxh': 'Mã BHXH',
            }
            df_show_nghi.rename(columns=col_map_nghi, inplace=True)
            
            st.caption(f"📌 Tổng số: **{len(ds_nghi)}** nhân viên đã nghỉ việc")
            st.dataframe(df_show_nghi, width='stretch', hide_index=True, height=400)
            
            st.divider()
            st.subheader("🔍 Xem chi tiết / Khôi phục nhân viên")
            
            nv_options = {f"{nv['ma_nv']} - {nv['ho_ten']} (Nghỉ: {format_date(nv.get('ngay_ket_thuc'))})": nv['id'] for nv in ds_nghi}
            selected_nghi_name = st.selectbox("Chọn nhân viên đã nghỉ:", list(nv_options.keys()))
            selected_nghi_id = nv_options[selected_nghi_name]
            
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("SELECT * FROM nhan_vien WHERE id = %s", (selected_nghi_id,))
            nv_nghi_detail = c.fetchone()
            db.close()
            
            if nv_nghi_detail:
                col_detail1, col_detail2 = st.columns(2)
                with col_detail1:
                    st.markdown("**📋 Thông tin cá nhân**")
                    st.write(f"- **Mã NV:** {nv_nghi_detail.get('ma_nv', '')}")
                    st.write(f"- **Họ tên:** {nv_nghi_detail.get('ho_ten', '')}")
                    st.write(f"- **Ngày sinh:** {format_date(nv_nghi_detail.get('ngay_sinh'))}")
                    st.write(f"- **Giới tính:** {nv_nghi_detail.get('gioi_tinh', '')}")
                    st.write(f"- **CCCD:** {nv_nghi_detail.get('so_cccd', '')}")
                    st.write(f"- **SĐT:** {nv_nghi_detail.get('dien_thoai', '')}")
                    st.write(f"- **Chức danh:** {nv_nghi_detail.get('chuc_danh_nghe', '')}")
                
                with col_detail2:
                    st.markdown("**📅 Thông tin hợp đồng & nghỉ việc**")
                    st.write(f"- **Số HĐLĐ:** {nv_nghi_detail.get('so_hdld', '')}")
                    st.write(f"- **Loại HĐ:** {nv_nghi_detail.get('loai_hop_dong', '')}")
                    st.write(f"- **Ngày vào làm:** {format_date(nv_nghi_detail.get('ngay_vao_lam'))}")
                    st.write(f"- **📅 Ngày nghỉ việc:** **{format_date(nv_nghi_detail.get('ngay_ket_thuc'))}**")
                    st.write(f"- **Mã BHXH:** {nv_nghi_detail.get('ma_so_bhxh', '')}")
                    st.write(f"- **Lý do nghỉ:** {nv_nghi_detail.get('ly_do_nghi', 'Chưa có thông tin')}")
                
                if st.session_state.role == "admin":
                    st.divider()
                    col_restore1, col_restore2, col_restore3 = st.columns([1, 2, 1])
                    with col_restore2:
                        if st.button(f"🔄 KHÔI PHỤC NHÂN VIÊN - {nv_nghi_detail['ho_ten']}", width='stretch', type="primary"):
                            try:
                                db = get_connection()
                                c = db.cursor()
                                loai_hd = nv_nghi_detail.get('loai_hop_dong', '')
                                if loai_hd == 'Thử việc':
                                    trang_thai_moi = 'THU_VIEC'
                                else:
                                    trang_thai_moi = 'DANG_LAM'
                                c.execute("""
                                    UPDATE nhan_vien 
                                    SET trang_thai = %s, 
                                        ngay_ket_thuc = NULL
                                    WHERE id = %s
                                """, (trang_thai_moi, selected_nghi_id))
                                db.commit()
                                db.close()
                                st.success(f"✅ Đã khôi phục nhân viên {nv_nghi_detail['ho_ten']}!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi khôi phục: {e}")
        else:
            st.info("📭 Không có nhân viên nào đã nghỉ việc")
    
    with tab_qtct:
        st.caption("📜 Lịch sử công tác và quyết định nhân sự")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien ORDER BY id DESC")
        all_nv = c.fetchall()
        db.close()
        
        if all_nv:
            nv_options = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in all_nv}
            selected_nv_history = st.selectbox("🔍 Chọn nhân viên:", list(nv_options.keys()), key="history_nv")
            nv_id_history = nv_options[selected_nv_history]
            
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("SELECT * FROM nhan_vien WHERE id = %s", (nv_id_history,))
            nv_current = c.fetchone()
            
            st.markdown(f"""
            ### 📌 Thông tin hiện tại của {nv_current['ho_ten']} ({nv_current['ma_nv']})
            | Trường | Giá trị |
            |--------|---------|
            | Trạng thái | {'🟢 Đang làm' if nv_current['trang_thai'] == 'DANG_LAM' else '🔵 Thử việc' if nv_current['trang_thai'] == 'THU_VIEC' else '🔴 Đã nghỉ'} |
            | Loại hợp đồng | {nv_current['loai_hop_dong']} |
            | Ngày vào làm | {format_date(nv_current['ngay_vao_lam'])} |
            | Ngày chính thức | {format_date(nv_current.get('ngay_chinh_thuc')) or 'Chưa có'} |
            | Chức danh | {nv_current['chuc_danh_nghe']} |
            | Phòng ban | {nv_current['phong_ban_lam_viec']} |
            """)
            
            c.execute("""
                SELECT * FROM quyet_dinh_nhan_su 
                WHERE nhan_vien_id = %s 
                ORDER BY ngay_quyet_dinh DESC
            """, (nv_id_history,))
            quyet_dinh_list = c.fetchall()
            
            if quyet_dinh_list:
                st.markdown("### 📋 Các quyết định nhân sự")
                qd_data = []
                for i, qd in enumerate(quyet_dinh_list, 1):
                    loai_qd_map = {
                        'THU_VIEC': '📝 Quyết định thử việc',
                        'CHINH_THUC': '✅ Quyết định chính thức',
                        'DIEU_CHUYEN': '🔄 Quyết định điều chuyển',
                        'BO_NHIEM': '⭐ Quyết định bổ nhiệm',
                        'TANG_LUONG': '💰 Quyết định tăng lương',
                        'NGHI_VIEC': '🚫 Quyết định nghỉ việc'
                    }
                    qd_data.append({
                        "STT": i,
                        "Loại quyết định": loai_qd_map.get(qd['loai_quyet_dinh'], qd['loai_quyet_dinh']),
                        "Số quyết định": qd['so_quyet_dinh'] or '...',
                        "Ngày quyết định": format_date(qd['ngay_quyet_dinh']),
                        "Ngày hiệu lực": format_date(qd['ngay_hieu_luc']),
                        "Nội dung": (qd['noi_dung'][:50] + "...") if qd['noi_dung'] and len(qd['noi_dung']) > 50 else qd['noi_dung']
                    })
                df_qd = pd.DataFrame(qd_data)
                st.dataframe(df_qd, width='stretch', hide_index=True)
                
                with st.expander("🔍 Xem chi tiết quyết định"):
                    qd_options = {f"{format_date(qd['ngay_quyet_dinh'])} - {qd['loai_quyet_dinh']}": qd for qd in quyet_dinh_list}
                    selected_qd_name = st.selectbox("Chọn quyết định:", list(qd_options.keys()), key="qd_detail")
                    selected_qd = qd_options[selected_qd_name]
                    st.markdown(f"""
                    **📄 Chi tiết quyết định:**
                    - **Số quyết định:** {selected_qd.get('so_quyet_dinh', '...')}
                    - **Ngày quyết định:** {format_date(selected_qd.get('ngay_quyet_dinh'))}
                    - **Ngày hiệu lực:** {format_date(selected_qd.get('ngay_hieu_luc'))}
                    - **Loại quyết định:** {selected_qd.get('loai_quyet_dinh')}
                    - **Nội dung:** {selected_qd.get('noi_dung', '...')}
                    - **Người ký:** {selected_qd.get('nguoi_ky', COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC'))}
                    
                    **📊 Thay đổi:**
                    - Chức danh: {selected_qd.get('chuc_danh_cu', '...')} → {selected_qd.get('chuc_danh_moi', '...')}
                    - Phòng ban: {selected_qd.get('phong_ban_cu', '...')} → {selected_qd.get('phong_ban_moi', '...')}
                    - Loại HĐ: {selected_qd.get('loai_hop_dong_cu', '...')} → {selected_qd.get('loai_hop_dong_moi', '...')}
                    """)
            else:
                st.info("📭 Nhân viên này chưa có quyết định nào.")
            
            c.execute("""
                SELECT * FROM lich_su_cong_tac 
                WHERE nhan_vien_id = %s 
                ORDER BY tu_ngay ASC
            """, (nv_id_history,))
            lich_su_list = c.fetchall()
            
            if lich_su_list:
                st.markdown("### 📅 Lịch sử công tác")
                ls_data = []
                for i, ls in enumerate(lich_su_list, 1):
                    ls_data.append({
                        "STT": i,
                        "Từ ngày": format_date(ls['tu_ngay']),
                        "Đến ngày": format_date(ls['den_ngay']) if ls['den_ngay'] else "Đang làm",
                        "Chức danh": ls['chuc_danh'] or '',
                        "Phòng ban": ls['phong_ban'] or '',
                        "Loại HĐ": ls['loai_hop_dong'] or '',
                        "Hệ số lương": ls['he_so_luong'] or ''
                    })
                df_ls = pd.DataFrame(ls_data)
                st.dataframe(df_ls, width='stretch', hide_index=True, height=400)
            else:
                st.info("📭 Chưa có lịch sử công tác. Đang tạo từ dữ liệu hiện tại...")
                loai_hd_dung = nv_current['loai_hop_dong']
                if nv_current['trang_thai'] == 'THU_VIEC':
                    loai_hd_dung = 'Thử việc'
                c.execute("""
                    INSERT INTO lich_su_cong_tac (nhan_vien_id, tu_ngay, chuc_danh, phong_ban, noi_lam_viec, loai_hop_dong, he_so_luong)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (nv_id_history, nv_current['ngay_vao_lam'], nv_current['chuc_danh_nghe'], 
                      nv_current['phong_ban_lam_viec'], nv_current['noi_lam_viec'], 
                      loai_hd_dung, nv_current['he_so_luong']))
                db.commit()
                st.rerun()
            db.close()
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")
    
    st.divider()
    
    # ========== BÁO CÁO THỐNG KÊ NHÂN SỰ (CHO TẤT CẢ MỌI NGƯỜI) ==========
    st.subheader("📊 BÁO CÁO THỐNG KÊ NHÂN SỰ")
    st.caption("Xuất báo cáo tổng hợp danh sách nhân viên đang làm việc")
    
    col_from, col_to, col_btn = st.columns([2, 2, 1])
    with col_from:
        tu_ngay_bc = st.date_input("Từ ngày:", value=date.today().replace(day=1), key="bc_thongke_tu")
    with col_to:
        den_ngay_bc = st.date_input("Đến ngày:", value=date.today(), key="bc_thongke_den")
    with col_btn:
        xuat_bc_thongke = st.button("📄 XUẤT BÁO CÁO THỐNG KÊ", width='stretch')
    
    if xuat_bc_thongke:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        # Lấy danh sách nhân viên đang làm việc trong kỳ
        c.execute("""
            SELECT ma_nv, ho_ten, ngay_sinh, gioi_tinh, chuc_danh_nghe, phong_ban_lam_viec,
                   loai_hop_dong, ngay_vao_lam, ngay_ky_hd, ngay_ket_thuc
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            ORDER BY ho_ten ASC
        """)
        ds_thongke = c.fetchall()
        db.close()
        
        if ds_thongke:
            file_path = tao_bao_cao_thong_ke_nhan_su(tu_ngay_bc, den_ngay_bc, ds_thongke)
            with open(file_path, "rb") as f:
                st.download_button(
                    label="📥 TẢI FILE BÁO CÁO THỐNG KÊ (Excel)",
                    data=f,
                    file_name=f"Bao_cao_thong_ke_nhan_su_{tu_ngay_bc.strftime('%d%m%Y')}_{den_ngay_bc.strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            st.success(f"✅ Đã xuất báo cáo thống kê với {len(ds_thongke)} nhân viên")
        else:
            st.info("📭 Không có nhân viên đang làm việc trong kỳ báo cáo!")
    
    st.divider()
    st.subheader("📊 Báo cáo tăng/giảm nhân sự trong kỳ")
    
    col_from, col_to, col_btn = st.columns([2, 2, 1])
    with col_from:
        tu_ngay_bc = st.date_input("Từ ngày:", value=date.today().replace(day=1), key="bc_tu")
    with col_to:
        den_ngay_bc = st.date_input("Đến ngày:", value=date.today(), key="bc_den")
    with col_btn:
        xuat_bc = st.button("📄 XUẤT BÁO CÁO TĂNG/GIẢM", width='stretch')
    
    if xuat_bc:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT ho_ten, chuc_danh_nghe, phong_ban_lam_viec, loai_hop_dong, ngay_vao_lam,
                   ngay_sinh, so_hdld, ngay_ky_hd
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_vao_lam BETWEEN %s AND %s
            ORDER BY ngay_vao_lam ASC
        """, (tu_ngay_bc, den_ngay_bc))
        tang_list = c.fetchall()
        c.execute("""
            SELECT ho_ten, chuc_danh_nghe, phong_ban_lam_viec, loai_hop_dong, ngay_vao_lam, ngay_ket_thuc,
                   ngay_sinh, so_hdld, ngay_ky_hd
            FROM nhan_vien 
            WHERE trang_thai = 'NGHI_VIEC'
            AND ngay_ket_thuc BETWEEN %s AND %s
            ORDER BY ngay_ket_thuc ASC
        """, (tu_ngay_bc, den_ngay_bc))
        giam_list = c.fetchall()
        db.close()
        if tang_list or giam_list:
            file_path = tao_bao_cao_tang_giam(tang_list, giam_list, tu_ngay_bc, den_ngay_bc)
            with open(file_path, "rb") as f:
                st.download_button(
                    label="📥 TẢI FILE BÁO CÁO TĂNG/GIẢM (Word)",
                    data=f,
                    file_name=f"Bao_cao_tang_giam_{tu_ngay_bc.strftime('%d%m%Y')}_{den_ngay_bc.strftime('%d%m%Y')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.info("Không có biến động nhân sự trong kỳ.")

# ========== CHẤM CÔNG (FACE ID) ==========
elif menu == "👆 Chấm công (Face ID)":
    st.title("👆 Chấm công bằng Face ID")
    st.caption("Quản lý chấm công nhân viên bằng công nghệ nhận diện khuôn mặt")
    
    st.info("""
    ### 🚧 Tính năng đang hoàn thiện
    
    Nội dung đang được phát triển. Các tính năng sắp ra mắt:
    - ✅ Đăng ký khuôn mặt cho nhân viên
    - ✅ Chấm công bằng camera
    - ✅ Lịch sử chấm công theo thời gian thực
    - ✅ Báo cáo đi muộn, về sớm
    - ✅ Tổng hợp công theo tháng
    - ✅ Tích hợp với tính lương tự động
    
    ⏳ **Dự kiến hoàn thành: Quý 4/2026**
    """)
    
    # Hiển thị thống kê demo
    with st.expander("📊 Thống kê chấm công hôm nay (Demo)"):
        col_demo1, col_demo2, col_demo3 = st.columns(3)
        col_demo1.metric("Đã chấm công", "0/0", "0%")
        col_demo2.metric("Đi làm đúng giờ", "0", "---")
        col_demo3.metric("Đi muộn", "0", "---")
        
        st.markdown("**Danh sách nhân viên chưa chấm công:**")
        st.caption("Chưa có dữ liệu - Tính năng đang phát triển")

# ========== TÍNH THU NHẬP ==========
elif menu == "💰 Tính thu nhập":
    st.title("💰 Tính thu nhập (Lương & Phụ cấp)")
    st.caption("Tính toán lương, thưởng và các khoản phụ cấp cho nhân viên")
    
    st.info("""
    ### 🚧 Tính năng đang hoàn thiện
    
    Nội dung đang được phát triển. Các tính năng sắp ra mắt:
    - ✅ Tính lương cơ bản theo hệ số
    - ✅ Tính các khoản phụ cấp (chức vụ, thâm niên, trách nhiệm...)
    - ✅ Tính thuế TNCN
    - ✅ Tính các khoản khấu trừ (BHXH, BHYT, BHTN, đoàn phí...)
    - ✅ Tổng hợp bảng lương tháng
    - ✅ Xuất bảng lương Excel/PDF
    - ✅ Gửi bảng lương qua email/Zalo cho nhân viên
    
    ⏳ **Dự kiến hoàn thành: Quý 4/2026**
    """)
    
    # Thêm form tính thử nghiệm demo
    with st.expander("🧪 Thử nghiệm tính lương (Demo)"):
        db_demo = get_connection()
        c_demo = db_demo.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c_demo.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY ho_ten LIMIT 10")
        nv_list = c_demo.fetchall()
        db_demo.close()
        
        if nv_list:
            nv_options = {f"{nv['ma_nv']} - {nv['ho_ten']}": nv['id'] for nv in nv_list}
            selected_nv = st.selectbox("Chọn nhân viên để tính thử:", list(nv_options.keys()))
            
            col_luong1, col_luong2 = st.columns(2)
            with col_luong1:
                luong_co_ban = st.number_input("Lương cơ bản (VNĐ)", min_value=0, value=5000000, step=500000)
                phu_cap_chuc_vu = st.number_input("Phụ cấp chức vụ (VNĐ)", min_value=0, value=0, step=100000)
            with col_luong2:
                phu_cap_tnvk = st.number_input("Phụ cấp thâm niên VK (%)", min_value=0.0, value=0.0, step=0.5)
                phu_cap_tnn = st.number_input("Phụ cấp thâm niên nghề (%)", min_value=0.0, value=0.0, step=0.5)
            
            tong_luong = luong_co_ban + phu_cap_chuc_vu
            tong_luong += luong_co_ban * phu_cap_tnvk / 100
            tong_luong += luong_co_ban * phu_cap_tnn / 100
            
            st.markdown("---")
            st.subheader("📊 Kết quả tính thử:")
            
            col_kq1, col_kq2, col_kq3 = st.columns(3)
            col_kq1.metric("Lương cơ bản", f"{luong_co_ban:,.0f} VNĐ")
            col_kq2.metric("Phụ cấp", f"{(tong_luong - luong_co_ban):,.0f} VNĐ")
            col_kq3.metric("Tổng thu nhập", f"{tong_luong:,.0f} VNĐ")
            
            st.caption("⚠️ Đây chỉ là tính toán tham khảo. Tính năng chính thức sẽ tích hợp với dữ liệu nhân viên và chấm công.")
        else:
            st.warning("Chưa có nhân viên nào trong hệ thống để thử nghiệm!")

# ========== UPLOAD ==========
elif menu=="📁 Upload hồ sơ" and st.session_state.role=="admin":
    st.title("📁 Quản lý hồ sơ nhân viên")
    tab_upload, tab_list = st.tabs(["📤 UPLOAD HỒ SƠ", "📋 DANH SÁCH HỒ SƠ"])
    
    with tab_upload:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY id DESC")
        nvl = c.fetchall()
        db.close()
        
        if nvl:
            nd = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in nvl}
            cn = st.selectbox("📌 Chọn nhân viên:", list(nd.keys()))
            
            col1, col2 = st.columns(2)
            with col1:
                lh = st.selectbox("📂 Loại hồ sơ:", ["BANG_CAP", "CHUNG_CHI", "CCCD", "HOP_DONG", "SO_YEU_LY_LICH", "KHAC"])
            with col2:
                st.markdown("💡 **Hướng dẫn:**")
                st.caption("- BANG_CAP: Bằng cấp, chứng chỉ")
                st.caption("- CHUNG_CHI: Chứng chỉ nghề")
                st.caption("- CCCD: Căn cước công dân")
                st.caption("- HOP_DONG: Hợp đồng lao động")
                st.caption("- SO_YEU_LY_LICH: Sơ yếu lý lịch")
            
            fl = st.file_uploader("📎 Chọn file:", type=['pdf', 'jpg', 'png', 'jpeg', 'doc', 'docx'])
            
            if fl:
                st.info(f"📄 Tên file: {fl.name} | 📏 Kích thước: {fl.size/1024:.1f} KB")
            
            if fl and st.button("📤 UPLOAD", type="primary", width='stretch'):
                nid = nd[cn]
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                fn = f"{nid}_{timestamp}_{lh}_{fl.name}"
                fp = os.path.join(UPLOAD_FOLDER, fn)
                
                with open(fp, "wb") as f:
                    f.write(fl.getbuffer())
                
                db = get_connection()
                c = db.cursor()
                c.execute("""
                    INSERT INTO ho_so_nhan_vien (nhan_vien_id, loai_ho_so, ten_file, duong_dan_file, ngay_upload) 
                    VALUES (%s, %s, %s, %s, CURRENT_DATE)
                """, (nid, lh, fl.name, fp))
                db.commit()
                db.close()
                
                st.success(f"✅ Đã upload thành công!\n📁 Lưu tại: {fp}")
                st.rerun()
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")
    
    with tab_list:
        st.subheader("📋 Danh sách hồ sơ đã upload")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien ORDER BY id DESC")
        nvl = c.fetchall()
        db.close()
        
        if nvl:
            nd = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in nvl}
            selected_nv = st.selectbox("🔍 Chọn nhân viên để xem hồ sơ:", list(nd.keys()), key="view_hoso")
            nv_id = nd[selected_nv]
            
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("""
                SELECT id, loai_ho_so, ten_file, duong_dan_file, ngay_upload 
                FROM ho_so_nhan_vien 
                WHERE nhan_vien_id = %s 
                ORDER BY ngay_upload DESC, id DESC
            """, (nv_id,))
            hs_list = c.fetchall()
            db.close()
            
            if hs_list:
                st.caption(f"📌 Tổng số: **{len(hs_list)}** hồ sơ")
                hs_data = []
                for i, hs in enumerate(hs_list, 1):
                    hs_data.append({
                        "STT": i,
                        "Loại hồ sơ": hs['loai_ho_so'],
                        "Tên file gốc": hs['ten_file'],
                        "Ngày upload": format_date(hs['ngay_upload']),
                        "ID": hs['id'],
                        "Đường dẫn": hs['duong_dan_file']
                    })
                df_hs = pd.DataFrame(hs_data)
                st.dataframe(df_hs[['STT', 'Loại hồ sơ', 'Tên file gốc', 'Ngày upload']], width='stretch', hide_index=True)
                
                st.divider()
                st.subheader("📄 Xem chi tiết & Tải xuống")
                hs_options = {f"{hs['loai_ho_so']} - {hs['ten_file']} (Ngày: {format_date(hs['ngay_upload'])})": hs for hs in hs_list}
                selected_hs_name = st.selectbox("Chọn hồ sơ:", list(hs_options.keys()))
                selected_hs = hs_options[selected_hs_name]
                
                col_info, col_download = st.columns([2, 1])
                with col_info:
                    st.markdown(f"""
                    **📋 Thông tin hồ sơ:**
                    - **Loại:** {selected_hs['loai_ho_so']}
                    - **Tên file gốc:** {selected_hs['ten_file']}
                    - **Đường dẫn:** `{selected_hs['duong_dan_file']}`
                    - **Ngày upload:** {format_date(selected_hs['ngay_upload'])}
                    """)
                
                with col_download:
                    if os.path.exists(selected_hs['duong_dan_file']):
                        with open(selected_hs['duong_dan_file'], "rb") as f:
                            st.download_button(
                                label="📥 TẢI HỒ SƠ",
                                data=f,
                                file_name=selected_hs['ten_file'],
                                mime="application/octet-stream",
                                width='stretch'
                            )
                    else:
                        st.error("❌ File không tồn tại trên máy chủ!")
                
                st.divider()
                col_del1, col_del2, col_del3 = st.columns([1, 2, 1])
                with col_del2:
                    if st.button("🗑️ XÓA HỒ SƠ NÀY", width='stretch', type="secondary"):
                        try:
                            if os.path.exists(selected_hs['duong_dan_file']):
                                os.remove(selected_hs['duong_dan_file'])
                            db = get_connection()
                            c = db.cursor()
                            c.execute("DELETE FROM ho_so_nhan_vien WHERE id = %s", (selected_hs['id'],))
                            db.commit()
                            db.close()
                            st.success(f"✅ Đã xóa hồ sơ: {selected_hs['ten_file']}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Lỗi khi xóa: {e}")
            else:
                st.info(f"📭 Nhân viên này chưa có hồ sơ nào được upload.")
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")

# ========== DANH MỤC CHỨC DANH ==========
elif menu == "⚙️ Danh mục" and st.session_state.role == "admin":
    st.title("⚙️ Quản lý danh mục Chức danh")
    if st.session_state.role == "admin":
        with st.expander("➕ Thêm chức danh mới", expanded=False):
            with st.form("add_chuc_danh"):
                ten_moi = st.text_input("Tên chức danh *"); mo_ta = st.text_area("Mô tả")
                if st.form_submit_button("💾 LƯU"):
                    if ten_moi:
                        db = get_connection(); c = db.cursor()
                        c.execute("SELECT COALESCE(MIN(t1.id + 1), 1) FROM vi_tri_cong_tac t1 LEFT JOIN vi_tri_cong_tac t2 ON t1.id + 1 = t2.id WHERE t2.id IS NULL AND t1.id >= 1")
                        id_trong = c.fetchone()[0]
                        c.execute("SELECT COALESCE(MAX(id),0) FROM vi_tri_cong_tac")
                        id_max = c.fetchone()[0]
                        id_moi = id_trong if id_trong <= id_max + 1 else id_max + 1
                        c.execute("INSERT INTO vi_tri_cong_tac (id, ten_vi_tri, ghi_chu) VALUES (%s, %s, %s)", (id_moi, ten_moi, mo_ta))
                        db.commit(); db.close(); st.success(f"✅ Đã thêm: {ten_moi}"); st.rerun()
                    else: st.error("Tên chức danh không được để trống!")
        st.subheader("📋 Danh sách chức danh")
        db = get_connection(); c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ten_vi_tri, ghi_chu FROM vi_tri_cong_tac ORDER BY id")
        ds = c.fetchall(); db.close()
        if ds:
            df = pd.DataFrame(ds); df.columns = ['ID', 'Tên chức danh', 'Ghi chú']; st.dataframe(df, width='stretch', hide_index=True)
            st.divider(); cdx = st.number_input("Nhập ID cần xóa:", min_value=1, step=1)
            if st.button("🗑️ XÓA", key="del_cd"):
                db = get_connection(); c = db.cursor()
                c.execute("DELETE FROM vi_tri_cong_tac WHERE id=%s", (cdx,)); db.commit(); db.close(); st.success("🗑️ Đã xóa!"); st.rerun()
        else: st.info("Chưa có chức danh nào")

# ========== BHXH ==========
elif menu == "📋 BHXH":
    st.title("📋 Quản lý BHXH")
    
    t1, t2, t3 = st.tabs(["📊 Tổng quan", "📝 Báo cáo tăng/giảm D02-LT", "💰 Dự toán đóng BHXH"])
    
    with t1:
        st.subheader("📊 Tổng quan tình hình đóng BHXH")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        # Thống kê chung
        c.execute("SELECT COUNT(*) as tong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')")
        tong_ld = c.fetchone()['tong']
        
        c.execute("SELECT COUNT(*) as dang_dong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND trang_thai_bhxh = 'DANG_DONG'")
        dang_dong = c.fetchone()['dang_dong']
        
        c.execute("SELECT COUNT(*) as chua_dong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND trang_thai_bhxh = 'CHUA_DONG'")
        chua_dong = c.fetchone()['chua_dong']
        
        c.execute("SELECT COUNT(*) as da_nghi FROM nhan_vien WHERE trang_thai = 'NGHI_VIEC'")
        da_nghi = c.fetchone()['da_nghi']
        
        db.close()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("👥 Tổng lao động", tong_ld)
        col2.metric("✅ Đang đóng BHXH", dang_dong, delta=f"{dang_dong/tong_ld*100:.0f}%" if tong_ld > 0 else None)
        col3.metric("⏳ Chưa đóng BHXH", chua_dong, delta=f"-{chua_dong}" if chua_dong > 0 else None, delta_color="inverse")
        col4.metric("📋 Đã nghỉ việc", da_nghi)
        
        st.divider()
        
        # Danh sách lao động chưa đóng BHXH
        st.subheader("⚠️ Lao động chưa đóng BHXH")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT ma_nv, ho_ten, chuc_danh_nghe, ngay_vao_lam, loai_hop_dong, thang_bat_dau_bh
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND trang_thai_bhxh = 'CHUA_DONG'
            ORDER BY ngay_vao_lam ASC
        """)
        chua_dong_list = c.fetchall()
        db.close()
        
        if chua_dong_list:
            df_chua_dong = pd.DataFrame(chua_dong_list)
            for col in df_chua_dong.columns:
                if 'ngay' in col.lower():
                    df_chua_dong[col] = df_chua_dong[col].apply(format_date)
            st.dataframe(df_chua_dong, width='stretch', hide_index=True)
            
            if st.session_state.role == "admin":
                st.warning("💡 Hướng dẫn: Vào menu '✅ Nhân viên' -> chọn nhân viên -> sửa thông tin -> cập nhật 'Bắt đầu BH' và chuyển trạng thái BHXH thành 'ĐANG ĐÓNG'")
        else:
            st.success("✅ Tất cả lao động đã được đăng ký đóng BHXH!")
    
    with t2:
        st.subheader("📝 Báo cáo tăng/giảm lao động tham gia BHXH (Mẫu D02-LT)")
        st.caption("Theo Thông tư 56/2017/TT-BYT và Quyết định 595/QĐ-BHXH")
        
        col_from, col_to = st.columns(2)
        with col_from:
            tu_ngay = st.date_input("📅 Từ ngày:", value=date(date.today().year, 1, 1), key="d02_tu")
        with col_to:
            den_ngay = st.date_input("📅 Đến ngày:", value=date.today(), key="d02_den")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        # Lao động tăng trong kỳ
        c.execute("""
            SELECT 
                ma_nv, ho_ten, ma_so_bhxh, ngay_sinh, gioi_tinh, so_cccd,
                chuc_danh_nghe, phong_ban_lam_viec, luong_bao_hiem, he_so_luong,
                COALESCE(thang_bat_dau_bh, ngay_vao_lam) as ngay_bat_dau,
                loai_hop_dong, so_hdld, ngay_vao_lam, thuong_tru
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND COALESCE(thang_bat_dau_bh, ngay_vao_lam) BETWEEN %s AND %s
            ORDER BY COALESCE(thang_bat_dau_bh, ngay_vao_lam) ASC
        """, (tu_ngay, den_ngay))
        tang_list = c.fetchall()
        
        # Lao động giảm trong kỳ
        c.execute("""
            SELECT 
                ma_nv, ho_ten, ma_so_bhxh, ngay_sinh, gioi_tinh, so_cccd,
                chuc_danh_nghe, phong_ban_lam_viec, luong_bao_hiem, he_so_luong,
                thang_ket_thuc_bh as ngay_ket_thuc,
                loai_hop_dong, so_hdld, ngay_vao_lam, thuong_tru, ly_do_nghi
            FROM nhan_vien 
            WHERE trang_thai = 'NGHI_VIEC'
            AND thang_ket_thuc_bh BETWEEN %s AND %s
            ORDER BY thang_ket_thuc_bh ASC
        """, (tu_ngay, den_ngay))
        giam_list = c.fetchall()
        db.close()
        
        col_tang, col_giam = st.columns(2)
        with col_tang:
            st.markdown(f"### 🟢 LAO ĐỘNG TĂNG ({len(tang_list)})")
            if tang_list:
                df_tang = pd.DataFrame(tang_list)
                for col in df_tang.columns:
                    if 'ngay' in col.lower():
                        df_tang[col] = df_tang[col].apply(format_date)
                st.dataframe(df_tang, width='stretch', hide_index=True, height=300)
            else:
                st.info("📭 Không có lao động tăng trong kỳ")
        
        with col_giam:
            st.markdown(f"### 🔴 LAO ĐỘNG GIẢM ({len(giam_list)})")
            if giam_list:
                df_giam = pd.DataFrame(giam_list)
                for col in df_giam.columns:
                    if 'ngay' in col.lower():
                        df_giam[col] = df_giam[col].apply(format_date)
                st.dataframe(df_giam, width='stretch', hide_index=True, height=300)
            else:
                st.info("📭 Không có lao động giảm trong kỳ")
        
        st.divider()
        
        # Chỉ admin mới được xuất Excel
        if st.session_state.role == "admin":
            if tang_list or giam_list:
                if st.button("📥 XUẤT EXCEL D02-LT (Mẫu báo cáo BHXH)", type="primary", width='stretch'):
                    from openpyxl import Workbook
                    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
                    from openpyxl.utils import get_column_letter
                    
                    thin_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "D02-LT"
                    
                    # Header thông tin đơn vị
                    ten_cong_ty = COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA")
                    ma_don_vi_bhxh = COMPANY_CONFIG.get("ma_don_vi_BHXH", "................")
                    
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
                    ws['A1'] = ten_cong_ty
                    ws['A1'].font = Font(bold=True, size=13, name='Times New Roman')
                    ws['A1'].alignment = Alignment(horizontal='center')
                    
                    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
                    ws['A2'] = f"Mã đơn vị BHXH: {ma_don_vi_bhxh}"
                    ws['A2'].font = Font(size=11, name='Times New Roman')
                    ws['A2'].alignment = Alignment(horizontal='center')
                    
                    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=8)
                    ws['A3'] = f"BÁO CÁO TĂNG/GIẢM LAO ĐỘNG THAM GIA BHXH (Mẫu D02-LT)"
                    ws['A3'].font = Font(bold=True, size=12, name='Times New Roman')
                    ws['A3'].alignment = Alignment(horizontal='center')
                    
                    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=8)
                    ws['A4'] = f"(Từ ngày {tu_ngay.strftime('%d/%m/%Y')} đến ngày {den_ngay.strftime('%d/%m/%Y')})"
                    ws['A4'].font = Font(size=11, name='Times New Roman')
                    ws['A4'].alignment = Alignment(horizontal='center')
                    
                    # Tạo sheet riêng cho từng loại
                    current_row = 6
                    
                    # ===== DANH SÁCH TĂNG =====
                    ws.cell(row=current_row, column=1, value="I. DANH SÁCH LAO ĐỘNG TĂNG MỚI THAM GIA BHXH")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=12, name='Times New Roman')
                    current_row += 1
                    
                    if tang_list:
                        headers = ["STT", "Mã NV", "Họ và tên", "Mã số BHXH", "Ngày sinh", "Giới tính", "Số CCCD", "Ngày bắt đầu BH"]
                        for col_idx, header in enumerate(headers, 1):
                            cell = ws.cell(row=current_row, column=col_idx, value=header)
                            cell.font = Font(bold=True, size=10, name='Times New Roman')
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = thin_border
                            cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
                        
                        current_row += 1
                        for idx, nv in enumerate(tang_list, 1):
                            ws.cell(row=current_row, column=1, value=idx)
                            ws.cell(row=current_row, column=2, value=nv.get('ma_nv', ''))
                            ws.cell(row=current_row, column=3, value=nv.get('ho_ten', ''))
                            ws.cell(row=current_row, column=4, value=nv.get('ma_so_bhxh', ''))
                            ws.cell(row=current_row, column=5, value=format_date(nv.get('ngay_sinh')))
                            ws.cell(row=current_row, column=6, value='Nam' if nv.get('gioi_tinh') == 'Nam' else 'Nữ' if nv.get('gioi_tinh') == 'Nữ' else '')
                            ws.cell(row=current_row, column=7, value=nv.get('so_cccd', ''))
                            ws.cell(row=current_row, column=8, value=format_date(nv.get('ngay_bat_dau')))
                            
                            for col_idx in range(1, 9):
                                cell = ws.cell(row=current_row, column=col_idx)
                                cell.border = thin_border
                                if col_idx == 3:
                                    cell.alignment = Alignment(horizontal='left', vertical='center')
                                else:
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                            current_row += 1
                        
                        current_row += 1
                    else:
                        ws.cell(row=current_row, column=1, value="Không có lao động tăng trong kỳ")
                        ws.cell(row=current_row, column=1).font = Font(italic=True, size=10, name='Times New Roman')
                        current_row += 2
                    
                    # ===== DANH SÁCH GIẢM =====
                    ws.cell(row=current_row, column=1, value="II. DANH SÁCH LAO ĐỘNG GIẢM (NGHỈ VIỆC)")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=12, name='Times New Roman')
                    current_row += 1
                    
                    if giam_list:
                        headers = ["STT", "Mã NV", "Họ và tên", "Mã số BHXH", "Ngày sinh", "Giới tính", "Số CCCD", "Ngày kết thúc BH", "Lý do nghỉ"]
                        for col_idx, header in enumerate(headers, 1):
                            cell = ws.cell(row=current_row, column=col_idx, value=header)
                            cell.font = Font(bold=True, size=10, name='Times New Roman')
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = thin_border
                            cell.fill = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
                        
                        current_row += 1
                        for idx, nv in enumerate(giam_list, 1):
                            ws.cell(row=current_row, column=1, value=idx)
                            ws.cell(row=current_row, column=2, value=nv.get('ma_nv', ''))
                            ws.cell(row=current_row, column=3, value=nv.get('ho_ten', ''))
                            ws.cell(row=current_row, column=4, value=nv.get('ma_so_bhxh', ''))
                            ws.cell(row=current_row, column=5, value=format_date(nv.get('ngay_sinh')))
                            ws.cell(row=current_row, column=6, value='Nam' if nv.get('gioi_tinh') == 'Nam' else 'Nữ' if nv.get('gioi_tinh') == 'Nữ' else '')
                            ws.cell(row=current_row, column=7, value=nv.get('so_cccd', ''))
                            ws.cell(row=current_row, column=8, value=format_date(nv.get('ngay_ket_thuc')))
                            ws.cell(row=current_row, column=9, value=nv.get('ly_do_nghi', ''))
                            
                            for col_idx in range(1, 10):
                                cell = ws.cell(row=current_row, column=col_idx)
                                cell.border = thin_border
                                if col_idx in [3, 9]:
                                    cell.alignment = Alignment(horizontal='left', vertical='center')
                                else:
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                            current_row += 1
                        
                        current_row += 1
                    else:
                        ws.cell(row=current_row, column=1, value="Không có lao động giảm trong kỳ")
                        ws.cell(row=current_row, column=1).font = Font(italic=True, size=10, name='Times New Roman')
                        current_row += 2
                    
                    # Footer
                    ws.cell(row=current_row, column=1, value=f"Tổng số lao động tăng: {len(tang_list)}")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
                    current_row += 1
                    ws.cell(row=current_row, column=1, value=f"Tổng số lao động giảm: {len(giam_list)}")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
                    current_row += 2
                    
                    # Ký tên
                    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=8)
                    ws.cell(row=current_row, column=6, value="NGƯỜI LẬP BÁO CÁO")
                    ws.cell(row=current_row, column=6).font = Font(bold=True, size=11, name='Times New Roman')
                    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal='center')
                    current_row += 1
                    
                    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=8)
                    ws.cell(row=current_row, column=6, value="(Ký, ghi rõ họ tên)")
                    ws.cell(row=current_row, column=6).font = Font(size=10, name='Times New Roman', italic=True)
                    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal='center')
                    current_row += 2
                    
                    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=8)
                    ws.cell(row=current_row, column=6, value=COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC').upper())
                    ws.cell(row=current_row, column=6).font = Font(bold=True, size=11, name='Times New Roman')
                    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal='center')
                    
                    # Điều chỉnh độ rộng cột
                    for col_idx in range(1, 10):
                        ws.column_dimensions[get_column_letter(col_idx)].width = 20
                    
                    filename = f"D02-LT_BHXH_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
                    wb.save(filename)
                    
                    with open(filename, "rb") as f:
                        st.download_button(
                            label="📥 TẢI FILE EXCEL D02-LT",
                            data=f,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch'
                        )
                    st.success(f"✅ Đã xuất báo cáo D02-LT với {len(tang_list)} lao động tăng và {len(giam_list)} lao động giảm")
            else:
                st.info("📭 Không có biến động lao động trong kỳ để xuất báo cáo")
        else:
            st.info("🔒 Chỉ Admin mới có quyền xuất file Excel báo cáo BHXH. Bạn đang ở chế độ xem (Viewer).")
    
    with t3:
        st.subheader("💰 DỰ TOÁN ĐÓNG BHXH")
        st.caption("Tính toán các khoản phải nộp Bảo hiểm xã hội, Bảo hiểm y tế, Bảo hiểm thất nghiệp")
        
        st.info("""
        ### 🚧 Tính năng đang hoàn thiện
        
        Nội dung đang được phát triển. Các tính năng sắp ra mắt:
        - ✅ Tính toán mức đóng BHXH theo lương cơ sở
        - ✅ Tính toán các khoản phụ cấp tính đóng BHXH
        - ✅ Bảng kê chi tiết từng nhân viên
        - ✅ Xuất báo cáo kê khai BHXH theo mẫu quy định
        - ✅ Tổng hợp số tiền phải nộp theo tháng/quý/năm
        
        ⏳ **Dự kiến hoàn thành: Quý 3/2026**
        """)
        
        # Thêm một số thông tin tham khảo
        with st.expander("📌 Thông tin tham khảo về tỷ lệ đóng BHXH hiện hành"):
            st.markdown("""
            **Tỷ lệ trích BHXH, BHYT, BHTN theo quy định hiện hành:**
            
            | Loại | Doanh nghiệp | Người lao động | Tổng |
            |------|-------------|----------------|------|
            | BHXH | 17.5% | 8% | 25.5% |
            | BHYT | 3% | 1.5% | 4.5% |
            | BHTN | 1% | 1% | 2% |
            | BHTNLĐ-BNN | 0.5% | 0% | 0.5% |
            | **Tổng cộng** | **22%** | **10.5%** | **32.5%** |
            
            *Lưu ý: Tỷ lệ có thể thay đổi theo quy định mới của Nhà nước.*
            """)
        
        # Thêm form nhập thử nghiệm (có thể comment lại sau)
        with st.expander("🧪 Thử nghiệm tính toán (Demo)"):
            col_demo1, col_demo2 = st.columns(2)
            with col_demo1:
                luong_demo = st.number_input("Lương tháng (VNĐ)", min_value=0, value=5000000, step=500000)
            with col_demo2:
                chon_ty_le = st.selectbox("Áp dụng tỷ lệ", ["Theo quy định", "Tùy chỉnh"])
            
            if chon_ty_le == "Theo quy định":
                ty_le_nld = 10.5
                ty_le_nsdl = 22.0
            else:
                col_ty1, col_ty2 = st.columns(2)
                with col_ty1:
                    ty_le_nld = st.number_input("Tỷ lệ NLĐ (%)", min_value=0.0, max_value=50.0, value=10.5, step=0.5)
                with col_ty2:
                    ty_le_nsdl = st.number_input("Tỷ lệ NSDLĐ (%)", min_value=0.0, max_value=50.0, value=22.0, step=0.5)
            
            tien_nld = luong_demo * ty_le_nld / 100
            tien_nsdl = luong_demo * ty_le_nsdl / 100
            tong_tien = tien_nld + tien_nsdl
            
            st.markdown("---")
            col_kq1, col_kq2, col_kq3 = st.columns(3)
            col_kq1.metric("NLĐ đóng", f"{tien_nld:,.0f} VNĐ", f"({ty_le_nld}%)")
            col_kq2.metric("NSDLĐ đóng", f"{tien_nsdl:,.0f} VNĐ", f"({ty_le_nsdl}%)")
            col_kq3.metric("Tổng tiền", f"{tong_tien:,.0f} VNĐ", "cả 2 bên")
        
# ========== BÁO CÁO TÌNH HÌNH SỬ DỤNG LAO ĐỘNG MẪU 01/PLI (EXCEL) ==========
elif menu == "📋 Báo cáo 01/PLI":
    st.title("📋 Báo cáo tình hình sử dụng lao động")
    st.caption("Theo mẫu 01/PLI Phụ lục I - Nghị định 145/2020/NĐ-CP (sửa đổi bởi Nghị định 35/2022/NĐ-CP)")
    
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    col1, col2 = st.columns(2)
    with col1:
        tu_ngay = st.date_input("📅 Từ ngày:", value=date(date.today().year, 1, 1), key="pli_tu")
    with col2:
        den_ngay = st.date_input("📅 Đến ngày:", value=date.today(), key="pli_den")
    
    db = get_connection()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute("""
        SELECT 
            nv.STT, nv.ma_nv, nv.ho_ten, nv.ma_so_bhxh, nv.ngay_sinh, nv.gioi_tinh,
            nv.so_cccd, nv.chuc_danh_nghe, nv.luong_bao_hiem, nv.he_so_luong,
            nv.phu_cap_chuc_vu, nv.phu_cap_tnvk, nv.phu_cap_tnn, nv.loai_hop_dong,
            nv.ngay_vao_lam, nv.ngay_ky_hd, nv.ngay_ket_thuc, nv.thang_bat_dau_bh,
            nv.thang_ket_thuc_bh, nv.so_hdld, nv.phong_ban_lam_viec, nv.noi_lam_viec
        FROM nhan_vien nv
        WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC', 'NGHI_VIEC')
        AND nv.ngay_vao_lam <= %s
        AND (nv.ngay_ket_thuc IS NULL OR nv.ngay_ket_thuc >= %s)
        ORDER BY nv.STT ASC
    """, (den_ngay, tu_ngay))
    ds_lao_dong = c.fetchall()
    db.close()
    
    st.info(f"📊 Tổng số lao động đang làm việc: **{len(ds_lao_dong)}** người")
    
    # Hiển thị bảng dữ liệu trước khi xuất (cho cả admin và viewer)
    if ds_lao_dong:
        st.subheader("📋 Danh sách lao động")
        df_preview = pd.DataFrame(ds_lao_dong)
        for col in df_preview.columns:
            if 'ngay' in col.lower():
                df_preview[col] = df_preview[col].apply(format_date)
        
        preview_cols = ['ma_nv', 'ho_ten', 'chuc_danh_nghe', 'loai_hop_dong', 'ngay_vao_lam', 'ma_so_bhxh']
        available_preview = [c for c in preview_cols if c in df_preview.columns]
        df_display = df_preview[available_preview]
        col_map_preview = {
            'ma_nv': 'Mã NV',
            'ho_ten': 'Họ tên',
            'chuc_danh_nghe': 'Chức danh',
            'loai_hop_dong': 'Loại HĐ',
            'ngay_vao_lam': 'Ngày vào làm',
            'ma_so_bhxh': 'Mã BHXH'
        }
        df_display.rename(columns=col_map_preview, inplace=True)
        st.dataframe(df_display, width='stretch', hide_index=True, height=400)
        
        st.divider()
        
        # Chỉ admin mới được xuất Excel
        if st.session_state.role == "admin":
            if st.button("📥 XUẤT EXCEL MẪU 01/PLI", type="primary", width='stretch'):
                wb = Workbook()
                ws = wb.active
                ws.title = "BC_Tinh_hinh_su_dung_LD"
                
                ten_cong_ty = COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA")
                dia_chi = COMPANY_CONFIG.get("dia_chi", "")
                ma_so_thue = COMPANY_CONFIG.get("ma_so_thue", "")
                dien_thoai_cty = COMPANY_CONFIG.get("dien_thoai_cty", "")
                
                # Header
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
                ws['A1'] = ten_cong_ty
                ws['A1'].font = Font(bold=True, size=13, name='Times New Roman')
                ws['A1'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=1, start_column=20, end_row=1, end_column=26)
                ws['T1'] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
                ws['T1'].font = Font(bold=True, size=13, name='Times New Roman')
                ws['T1'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
                ws['A2'] = f"Số: 01/BC-01PLI-{datetime.now().year}/CHL"
                ws['A2'].font = Font(size=12, name='Times New Roman')
                ws['A2'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=2, start_column=20, end_row=2, end_column=26)
                ws['T2'] = "Độc lập - Tự do - Hạnh phúc"
                ws['T2'].font = Font(italic=True, size=12, name='Times New Roman')
                ws['T2'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=3, start_column=20, end_row=3, end_column=26)
                ws['T3'] = f"Quảng Trị, ngày {date.today().day} tháng {date.today().month} năm {date.today().year}"
                ws['T3'].font = Font(italic=True, size=12, name='Times New Roman')
                ws['T3'].alignment = Alignment(horizontal='right')
                
                ws.merge_cells('A5:AA5')
                ws['A5'] = "BÁO CÁO TÌNH HÌNH SỬ DỤNG LAO ĐỘNG"
                ws['A5'].font = Font(bold=True, size=13, name='Times New Roman')
                ws['A5'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells('A6:AA6')
                ws['A6'] = f"(Từ ngày {tu_ngay.strftime('%d/%m/%Y')} đến ngày {den_ngay.strftime('%d/%m/%Y')})"
                ws['A6'].font = Font(size=12, name='Times New Roman')
                ws['A6'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells('A8:AA8')
                ws['A8'] = "Kính gửi: SỞ NỘI VỤ TỈNH QUẢNG TRỊ"
                ws['A8'].font = Font(bold=True, size=11, name='Times New Roman')
                ws['A8'].alignment = Alignment(horizontal='left')
                
                ws['A10'] = "1. Thông tin chung về doanh nghiệp:"
                ws['A10'].font = Font(bold=True, size=11, name='Times New Roman')
                
                row_info = 11
                for label in [f"- Tên doanh nghiệp: {ten_cong_ty}", f"- Địa chỉ: {dia_chi}", 
                             f"- Mã số thuế: {ma_so_thue}", f"- Điện thoại: {dien_thoai_cty}"]:
                    ws[f'A{row_info}'] = label
                    ws[f'A{row_info}'].font = Font(size=11, name='Times New Roman')
                    row_info += 1
                
                ws[f'A{row_info + 1}'] = "2. Thông tin tình hình sử dụng lao động của đơn vị:"
                ws[f'A{row_info + 1}'].font = Font(bold=True, size=11, name='Times New Roman')
                
                header_row = 18
                col_widths = [5, 25, 18, 15, 8, 18, 25, 12, 18, 18, 12, 15, 
                             12, 12, 12, 12, 15, 12, 12, 18, 18, 18, 18, 18, 18, 18, 20]
                for i, w in enumerate(col_widths, 1):
                    ws.column_dimensions[get_column_letter(i)].width = w
                
                stt_row = header_row + 3
                for col in range(1, 28):
                    ws.cell(row=stt_row, column=col, value=f"({col})")
                    ws.cell(row=stt_row, column=col).font = Font(size=9, name='Times New Roman')
                    ws.cell(row=stt_row, column=col).alignment = Alignment(horizontal='center')
                    ws.cell(row=stt_row, column=col).border = thin_border
                
                # Merge cells header
                ws.merge_cells(start_row=header_row, start_column=1, end_row=header_row+2, end_column=1)
                ws.cell(row=header_row, column=1, value="STT")
                
                ws.merge_cells(start_row=header_row, start_column=2, end_row=header_row+2, end_column=2)
                ws.cell(row=header_row, column=2, value="Họ và tên")
                
                ws.merge_cells(start_row=header_row, start_column=3, end_row=header_row+2, end_column=3)
                ws.cell(row=header_row, column=3, value="Mã số BHXH")
                
                ws.merge_cells(start_row=header_row, start_column=4, end_row=header_row+2, end_column=4)
                ws.cell(row=header_row, column=4, value="Ngày sinh")
                
                ws.merge_cells(start_row=header_row, start_column=5, end_row=header_row+2, end_column=5)
                ws.cell(row=header_row, column=5, value="Giới tính")
                
                ws.merge_cells(start_row=header_row, start_column=6, end_row=header_row+2, end_column=6)
                ws.cell(row=header_row, column=6, value="Số CCCD/Hộ chiếu")
                
                ws.merge_cells(start_row=header_row, start_column=7, end_row=header_row+2, end_column=7)
                ws.cell(row=header_row, column=7, value="Chức danh nghề, vị trí, công việc")
                
                ws.merge_cells(start_row=header_row, start_column=8, end_row=header_row, end_column=11)
                ws.cell(row=header_row, column=8, value="Vị trí việc làm (2)")
                
                ws.merge_cells(start_row=header_row, start_column=12, end_row=header_row, end_column=17)
                ws.cell(row=header_row, column=12, value="Tiền lương")
                
                ws.merge_cells(start_row=header_row, start_column=20, end_row=header_row, end_column=24)
                ws.cell(row=header_row, column=20, value="Loại và hiệu lực hợp đồng")
                
                ws.merge_cells(start_row=header_row, start_column=18, end_row=header_row+1, end_column=19)
                ws.cell(row=header_row, column=18, value="Ngành nghề nặng nhọc, độc hại")
                
                ws.merge_cells(start_row=header_row+1, start_column=13, end_row=header_row+1, end_column=17)
                ws.cell(row=header_row+1, column=13, value="Phụ cấp")
                
                ws.merge_cells(start_row=header_row+1, start_column=21, end_row=header_row+1, end_column=22)
                ws.cell(row=header_row+1, column=21, value="Hiệu lực HĐLĐ xác định thời hạn")
                
                ws.merge_cells(start_row=header_row+1, start_column=23, end_row=header_row+1, end_column=24)
                cell = ws.cell(row=header_row+1, column=23, value="Hiệu lực HĐLĐ khác (dưới 1 tháng, thử việc)")
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                
                ws.merge_cells(start_row=header_row+1, start_column=8, end_row=header_row+2, end_column=8)
                ws.cell(row=header_row+1, column=8, value="Nhà quản lý")
                
                ws.merge_cells(start_row=header_row+1, start_column=9, end_row=header_row+2, end_column=9)
                ws.cell(row=header_row+1, column=9, value="Chuyên môn kỹ thuật bậc cao")
                
                ws.merge_cells(start_row=header_row+1, start_column=10, end_row=header_row+2, end_column=10)
                ws.cell(row=header_row+1, column=10, value="Chuyên môn kỹ thuật bậc trung")
                
                ws.merge_cells(start_row=header_row+1, start_column=11, end_row=header_row+2, end_column=11)
                ws.cell(row=header_row+1, column=11, value="Khác")
                
                ws.merge_cells(start_row=header_row+1, start_column=12, end_row=header_row+2, end_column=12)
                ws.cell(row=header_row+1, column=12, value="Mức lương/Hệ số lương")
                
                ws.merge_cells(start_row=header_row+1, start_column=20, end_row=header_row+2, end_column=20)
                ws.cell(row=header_row+1, column=20, value="Ngày bắt đầu HĐLĐ không xác định thời hạn")
                
                ws.merge_cells(start_row=header_row, start_column=25, end_row=header_row+2, end_column=25)
                ws.cell(row=header_row, column=25, value="Thời điểm bắt đầu đóng BHXH")
                
                ws.merge_cells(start_row=header_row, start_column=26, end_row=header_row+2, end_column=26)
                ws.cell(row=header_row, column=26, value="Thời điểm kết thúc đóng BHXH")
                
                ws.merge_cells(start_row=header_row, start_column=27, end_row=header_row+2, end_column=27)
                ws.cell(row=header_row, column=27, value="Ghi chú")
                
                for row in range(header_row, header_row + 3):
                    for col in range(1, 28):
                        cell = ws.cell(row=row, column=col)
                        cell.border = thin_border
                
                ws.cell(row=header_row+2, column=13, value="Phụ cấp chức vụ")
                ws.cell(row=header_row+2, column=14, value="Phụ cấp thâm niên VK(%)")
                ws.cell(row=header_row+2, column=15, value="Phụ cấp thâm niên nghề (%)")
                ws.cell(row=header_row+2, column=16, value="Phụ cấp thâm niên nghề (%)")
                ws.cell(row=header_row+2, column=17, value="Các khoản bổ sung")
                ws.cell(row=header_row+2, column=18, value="Ngày bắt đầu")
                ws.cell(row=header_row+2, column=19, value="Ngày kết thúc")
                ws.cell(row=header_row+2, column=21, value="Ngày bắt đầu")
                ws.cell(row=header_row+2, column=22, value="Ngày kết thúc")
                ws.cell(row=header_row+2, column=23, value="Ngày bắt đầu")
                ws.cell(row=header_row+2, column=24, value="Ngày kết thúc")
                
                for row in range(header_row, header_row + 3):
                    for col in range(1, 28):
                        cell = ws.cell(row=row, column=col)
                        if cell.value:
                            cell.font = Font(bold=True, size=10, name='Times New Roman')
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            cell.border = thin_border
                
                for col in range(1, 28):
                    cell = ws.cell(row=stt_row, column=col)
                    cell.font = Font(size=9, name='Times New Roman')
                    cell.alignment = Alignment(horizontal='center')
                    cell.border = thin_border
                
                data_row = stt_row + 1
                for idx, nv in enumerate(ds_lao_dong, 1):
                    row = data_row + idx - 1
                    
                    ws.cell(row=row, column=1, value=idx)
                    ws.cell(row=row, column=2, value=nv.get('ho_ten', ''))
                    ws.cell(row=row, column=3, value=nv.get('ma_so_bhxh', ''))
                    ws.cell(row=row, column=4, value=format_date(nv.get('ngay_sinh')))
                    gt = nv.get('gioi_tinh', '')
                    ws.cell(row=row, column=5, value='Nam' if gt == 'Nam' else 'Nữ' if gt == 'Nữ' else '')
                    ws.cell(row=row, column=6, value=nv.get('so_cccd', ''))
                    ws.cell(row=row, column=7, value=nv.get('chuc_danh_nghe', ''))
                    
                    cd = (nv.get('chuc_danh_nghe') or '').lower()
                    is_quan_ly = any(x in cd for x in ['giám đốc', 'trưởng phòng', 'phó', 'quản lý'])
                    ws.cell(row=row, column=8, value='x' if is_quan_ly else '')
                    is_chuyen_mon_cao = (any(x in cd for x in ['kỹ thuật', 'kĩ thuật']) and any(x in cd for x in ['cao', 'chính', 'kỹ sư']))
                    ws.cell(row=row, column=9, value='x' if is_chuyen_mon_cao else '')
                    is_khac = any(x in cd for x in ['phổ thông', 'lao động', 'tạp vụ', 'bảo vệ', 'tạp vụ', 'lái xe'])
                    ws.cell(row=row, column=11, value='x' if is_khac else '')
                    is_trung = (not is_quan_ly) and (not is_chuyen_mon_cao) and (not is_khac)
                    ws.cell(row=row, column=10, value='x' if is_trung else '')
                    
                    luong = nv.get('luong_bao_hiem', '')
                    heso = nv.get('he_so_luong', '')
                    ws.cell(row=row, column=12, value=f"Hệ số: {heso}" if heso and str(heso).strip() else str(luong) if luong else '')
                    ws.cell(row=row, column=13, value=str(nv.get('phu_cap_chuc_vu', '')) if nv.get('phu_cap_chuc_vu') else '')
                    ws.cell(row=row, column=14, value=f"{nv.get('phu_cap_tnvk', '')}%" if nv.get('phu_cap_tnvk') else '')
                    ws.cell(row=row, column=15, value=f"{nv.get('phu_cap_tnn', '')}%" if nv.get('phu_cap_tnn') else '')
                    ws.cell(row=row, column=16, value='')
                    ws.cell(row=row, column=17, value='')
                    ws.cell(row=row, column=18, value='')
                    ws.cell(row=row, column=19, value='')
                    
                    loai_hd = nv.get('loai_hop_dong', '')
                    ngay_bd = nv.get('ngay_ky_hd') or nv.get('ngay_vao_lam')
                    ngay_kt = nv.get('ngay_ket_thuc')
                    ws.cell(row=row, column=20, value=format_date(ngay_bd) if loai_hd == 'Không xác định thời hạn' else '')
                    
                    if loai_hd == 'Xác định thời hạn':
                        ws.cell(row=row, column=21, value=format_date(ngay_bd))
                        ws.cell(row=row, column=22, value=format_date(ngay_kt) if ngay_kt else '')
                    else:
                        ws.cell(row=row, column=21, value='')
                        ws.cell(row=row, column=22, value='')
                    
                    if loai_hd == 'Thử việc':
                        ws.cell(row=row, column=23, value=format_date(ngay_bd))
                        ws.cell(row=row, column=24, value=format_date(ngay_kt) if ngay_kt else '')
                    else:
                        ws.cell(row=row, column=23, value='')
                        ws.cell(row=row, column=24, value='')
                    
                    ws.cell(row=row, column=25, value=format_date(nv.get('thang_bat_dau_bh')))
                    ws.cell(row=row, column=26, value=format_date(nv.get('thang_ket_thuc_bh')))
                    ws.cell(row=row, column=27, value=nv.get('so_hdld', ''))
                    
                    for col in range(1, 28):
                        cell = ws.cell(row=row, column=col)
                        cell.border = thin_border
                        cell.font = Font(size=10, name='Times New Roman')
                        if col in [1, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                
                total_row = data_row + len(ds_lao_dong)
                ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
                ws.cell(row=total_row, column=1, value=f"Tổng cộng: {len(ds_lao_dong)} người")
                ws.cell(row=total_row, column=1).font = Font(bold=True, size=10, name='Times New Roman')
                ws.cell(row=total_row, column=1).border = thin_border
                
                sign_row = total_row + 3
                ws.merge_cells(start_row=sign_row, start_column=23, end_row=sign_row, end_column=27)
                ws.cell(row=sign_row, column=23, value="ĐẠI DIỆN DOANH NGHIỆP")
                ws.cell(row=sign_row, column=23).font = Font(bold=True, size=11, name='Times New Roman')
                ws.cell(row=sign_row, column=23).alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=sign_row+1, start_column=23, end_row=sign_row+1, end_column=27)
                ws.cell(row=sign_row+1, column=23, value="(Ký, đóng dấu, ghi rõ họ tên)")
                ws.cell(row=sign_row+1, column=23).font = Font(size=10, name='Times New Roman')
                ws.cell(row=sign_row+1, column=23).alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=sign_row+2, start_column=23, end_row=sign_row+2, end_column=27)
                ws.cell(row=sign_row+2, column=23, value=COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC').upper())
                ws.cell(row=sign_row+2, column=23).font = Font(bold=True, size=11, name='Times New Roman')
                ws.cell(row=sign_row+2, column=23).alignment = Alignment(horizontal='center')
                
                filename = f"Bao_cao_01_PLI_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
                wb.save(filename)
                
                with open(filename, "rb") as f:
                    st.download_button(
                        label="📥 TẢI FILE EXCEL",
                        data=f,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width='stretch'
                    )
                st.success(f"✅ Đã xuất báo cáo với {len(ds_lao_dong)} lao động")
        else:
            st.info("🔒 Chỉ Admin mới có quyền xuất file Excel báo cáo 01/PLI. Bạn đang ở chế độ xem (Viewer).")
            st.caption("💡 Với quyền Viewer, bạn có thể xem danh sách lao động ở trên nhưng không thể tải file Excel.")
    else:
        st.warning("⚠️ Không có lao động nào đang làm việc trong kỳ báo cáo!")
            
st.sidebar.divider()
st.sidebar.caption("© 2026 HRM-Port | Cảng biển quốc tế Hòn La")


#===== Hàm xử lý chính =====
def main():
    """Hàm điều khiển chính - phân luồng Landing Page / HRM App"""
    
    # Khởi tạo session state
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.role = None
        st.session_state.username = None
    
    # Kiểm tra URL parameter để xử lý login từ landing page
    query_params = st.query_params
    if query_params.get('login') == 'true':
        # Hiển thị form đăng nhập
        st.session_state['show_login_form'] = True
        st.query_params.clear()
    
    # Nếu chưa đăng nhập, hiển thị Landing Page
    if not st.session_state.logged_in:
        # Ẩn sidebar
        st.markdown("""
            <style>
                [data-testid="stSidebar"] { display: none !important; }
                [data-testid="collapsedControl"] { display: none !important; }
                header { display: none !important; }
            </style>
        """, unsafe_allow_html=True)
        
        show_landing_page()
        
        # Hiển thị form đăng nhập nếu được yêu cầu
        if st.session_state.get('show_login_form', False):
            with st.container():
                st.markdown("""
                <div style="position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);
                            background:white;padding:30px;border-radius:15px;z-index:10000;
                            box-shadow:0 0 30px rgba(0,0,0,0.3);width:350px;">
                    <h3 style="color:#0f3b5c;">🔐 Đăng nhập HRM-Port</h3>
                </div>
                """, unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns([1,2,1])
                with col2:
                    username = st.text_input("Tài khoản", key="landing_user")
                    password = st.text_input("Mật khẩu", type="password", key="landing_pass")
                    if st.button("Đăng nhập", key="landing_login"):
                        success, role = check_login(username, password)
                        if success:
                            st.session_state.logged_in = True
                            st.session_state.role = role
                            st.session_state.username = username
                            st.session_state['show_login_form'] = False
                            st.rerun()
                        else:
                            st.error("Sai tài khoản hoặc mật khẩu!")
                    if st.button("Hủy", key="landing_cancel"):
                        st.session_state['show_login_form'] = False
                        st.rerun()
        st.stop()

# Chạy ứng dụng
if __name__ == "__main__":
    main()