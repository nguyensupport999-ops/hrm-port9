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
          <img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gIoSUNDX1BST0ZJTEUAAQEAAAIYanhsIARAAABtbnRyUkdCIFhZWiAH4wAMAAEAAAAAAABhY3NwQVBQTAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLWp4bCACufkBQHM6b/D/A/Tw9worAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAtkZXNjAAABCAAAAERjcHJ0AAABTAAAACR3dHB0AAABcAAAABRjaGFkAAABhAAAACxjaWNwAAABsAAAAAxyWFlaAAABvAAAABRnWFlaAAAB0AAAABRiWFlaAAAB5AAAABRyVFJDAAAB+AAAACBnVFJDAAAB+AAAACBiVFJDAAAB+AAAACBtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACYAAAAcAFIARwBCAF8ARAA2ADUAXwBTAFIARwBfAFIAZQBsAF8AUwBSAEcAAG1sdWMAAAAAAAAAAQAAAAxlblVTAAAABgAAABwAQwBDADAAAFhZWiAAAAAAAAD21gABAAAAANMtc2YzMgAAAAAAAQxAAAAF3f//8yoAAAeSAAD9kP//+6P///2jAAAD2wAAwIFjaWNwAAAAAAENAAFYWVogAAAAAAAAb58AADj1AAADkFhZWiAAAAAAAABilgAAt4cAABjbWFlaIAAAAAAAACSiAAAPhQAAttZwYXJhAAAAAAADAAAAAmZmAADypwAADVkAABPQAAAKW//bAEMAAgEBAQEBAgEBAQICAgICBAMCAgICBQQEAwQGBQYGBgUGBgYHCQgGBwkHBgYICwgJCgoKCgoGCAsMCwoMCQoKCv/bAEMBAgICAgICBQMDBQoHBgcKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCv/AABEIAuQCvQMBEQACEQEDEQH/xAAeAAACAgMBAQEBAAAAAAAAAAADBAAFAgYHAQgJCv/EAG8QAAEBBQEHCQ4RCAgEBAQGAwQAAQMFBhQTAgcRIyQzNAgVMUNEU1RhcxIWISUyNWNkgYKDkqPwCSJBQkVRUlVicnSToaKys8MXJjZxhMLS4mV1kZSk0+PyN7G08xgnRsRHhcHhVnaV0dTkCmbV/8QAGwEAAwEBAQEBAAAAAAAAAAAAAAMEAgEFBgf/xAA7EQEAAQMCBQEFBgYBBAIDAAAAAQIEESExAwUSFEETBhUiMlEkMzRCRPAjNUNFYXElFoGRsQfBUtHx/9oADAMBAAIRAxEAPwD60Xx77gat88KWoGTE4tYzjQBa3zwoAyAIgCIAiAcQEQBQ9juoBtAEQBEBEAwgIhOIgIgIgCIAisUF0sIgIgIgCKdOGqFAinCITogIhQIrAXSwiAwo2ca1kPFKnLqwCIUBoCKNOGgBqwsNTliJagurAwo2caHNXinIRAMJagCjZxoAVG3iVjmIRRuhqxQ9o2ca4GNFde2pAwVgSjbxIGqUbeJA1SjbxIGpdCdKPi+lDmJDo28SkdSjbxIAVFde2uKAEAwgCKhQNReeBAKmSv6iAq9ZGe6YmJcQmsjPdMQMQIo2hqLzwKxzEC0TA0tQKh0WiaYxAQwJMTPKM7iUbooex3UA/kSArzQgtxdxLUIGgImJ3tWEgFTeqYgHXGHmOgcgNTQEQDCWoEQWIgCJgGCTpajc1WN4klkRLBhMLEQBEAWsZxoA6AIgGEAQPZ7iAiCxEARBaICJhgiWETDEQEQEQWiWETAGmGCJZaJZaIMRMCICJZYioUBqdOGqAwo2ca1lQEtBEtOIpwXTAHR8X0pegRARMAaYESwGlhEBEwBoCJYDTDEQYGgIgsNMCIAiAFQt4lOWLRncSscxAdG3iTE2IDMCUaomqDA0sIGGhYsEAWsZxradmsKCxgSAVo2caArzIKbhwKdOKEgLAJUKA6LzwLOYLYhhmbNCklrYIIzYaArAVSwiADQ9uIUPKNnGp073LvPChrRkhkSj4vpVWADReeBZUGaK69tTp2kVnH9CAiYBFQWIo1hhLAiYnEQYYQYYQWIlgRMLEQBEARAWCAiAYQBEBEFiIAiCxEBEGIgIgIgIgImANMMMIAaWWiWWiAiDEQETDETAiCw0ARLRhoCIAdG3iTdD8QiU6GgIlhEANMAaYBEsB0fF9KAHRt4k3QIlBjQs40sMaPi+lGhmoaYEQA0stlReeBMMwlF54EM5lkqGTCjAaYA0FpRt4l3QMKNnGuFq00LCaqF71LNRRmCKgtEwwoYabsoCB7HdQDaYnY0QXmxRgSi88CdmG8wxWjWNaGGpyNFiEarHGNGzjUYZUTPe9qoDCjbxLOIAdHxfSkguYjwCdZx/QgLEPY7qrASwEUahz9AMKwCJadFQWIgGF56wRMTmEGDA7HdTp3djc0kuIgswgCIBhARAOIBhARAEQBEFiIAiCxEBEBEBEBEBEBEwxEBEstEARADTDETDESwIqAGllhoMRLeeiDESzBEwIlgNWB5RXXtqcMEBEANARAESwXVASjbxJYDQHtGzjS8hKNnGgJRs40BjRXXtoD1ADQA0ARAEQGdC3iTBiHqsToowWMCVB6tNCQpCS1CIAiC3lEwxqnLZUbONbyZhnkSwWsQqLAq5AdYzjRqFfWA8awADaNTgUI0PZQFiFliq1CwCgvqMXQEbW4eigKxSBlrNxJgB51+L6VzV3QpRmhJB5bLKxWA4gOfqMCJgEQWIgsRAESzBEGGExOKHsd1anYxmnA4py0QYYQWIlmCJhYiAYQDCAIgGEBEFmEFiICICICICICICICICIAiAiAiDETDEQETAiCw0ssNWJxFGoRLCICICJgRLCIAasCUfF9KnCIAaoCKcBqgIpwiWEQEQEQEQA6NvEmArReeBUA3RXPtqMDoDEJWCDSjDKhZxoAKYnJ5EqFBM2CoUB0beJANUXngWcwC+s3EsYCazN4v7EjQF6KiVgGQEQClGzjXAzWAyow/ct/sSwYCDCSwsq0wJqrgCVnaH0owEyJdCIAaAqjeqYp253KIdRMMc3UYEQBEAwmJ0QBEB6Hsd1AHSzBgkwsZMMGrfPCkYkwZcAiAYTCxEARBYiAcQBEAwgCIAiCxEFiICICICICICICICIAiAGgCIMEQYGqAYU4LqhOiWA0wIlhEBEBEBEBEsIgImANLAiAiYESw8oW8SYHlF54EBigIgIgIgPKK69tAepYLpgRARAEQBFQEU4DSwImARLBOj4vpVeo1e5ZxLrmrCi88CzmDXlGzjWTGRoS2AKPi+lSAAzY7qpjdyNgdZS/b+hd6ode0RnAmrQTLVOEy1UAQM01AOVh3EuBFgPK269pd0Dyt88K7iADWN4lKFeYgB5ErA5mvPUIgGEARAEQWImARBYiWoRMTmFYDAez3FGYYSwYQYioLMJYFrGcaxo5qOtsCJicwlqDCAIgCILEQWYQEQEQEQEQEQBEBEBEBEBEGImGCJgDQWGqEZheeoDQA0BEwIqAiAinCJYRMLRARARAEQYiWETAioD2iufbQHtF54FOGNG3iWsQA1oImBFOCtF54FQMPaK69tLACYBEsCKcCJgRUBFOEQnRMMRUFmEt6ABuxspfQV3jxDpdBbCiufbSfidxLwI3gSqcHNCw8IUYJ6FpqaYTNozdn+xPMFDCQD1CZxpYYqcE1YA1GEQGVEEYgF8h88Cr0c0ctUjqJZgiYYIlgRAETCzCCxEBEGGEswRMLErG8SziHMmFl0RMMMJZaIMMILMJiMwhQYQBKzj+hAEQBEFiIBhBaICICIAiAiAiAIgBoMRMCJhiJZaILDQEQESwGmB5XN4kDMPazj+hLCVnH9CAlY3iTHNREt0NMLRBaICIMEVBginCJYMJgESwiAGrAGlhEsBoAaYEo28SAGgF0wGaNnGkYkM0oBpgRARLCJhYgez3F38oOZElGJRB+23+1MTg0XngS1ANHxfSu4kETQsHQauAHQvPZQoWFZx/QhOWMDrlQ1sWovPAtZg5loXnsqULEM1Vgodsd1KjcFUkF6NvErAGowYQGNF54EByFLUImFiKgCKcxEsGEARMLESwIgCIBhAETCxEswwgCIMMIAiYDCWBEwsRBZhAMIAiAYSwImFiILRARAEQEQBEANBiJagNMLRMLRBbKuZxpYBQA16CMugxlVs4cN8nwrnZOdQRvbp47h+uuBGRcKDdA3EONzkYcW+XOxd6ngRgZuWhHDv11w3W+eFAwlb54UDDFDRhBaKMwRAETAIlhEGCIBhLMRMCICIAaAiAiAGgIgCUbeJLBdWBEsIlhjQs40AOi88CWGSYWiCxaxnGhQNr0xLdyHXN4kOC5EmJwTUBVaytS1A1C3iTE4qoUJR8X0qdODReeBMWYY0Z3Et5MEXQHRt4lzUK02tUgGQDIXUtQ7G46HHC0tQiYWIgCIMESwIgBoBhBZhMLEQYiWYYTE4iWYYQYMEgDIAoex3UAdAEQWImFmEARWA0EoZ2AyyBEsCJhYiYWiAiAIgIgxEwBqgsNLCVnH9CnANb54VYMh1oWy3H8ITErjd/7V46mbU6ivypxndySe40iHjEO7RV9jlmeJEPj2/H6PPLLhr8OS7iH3PIYxdYmpwWZfR0r6ZvQGiZD7iZaDO/EuFj4TuuXLZx9Fp1Sky9EV4Q1y/wDnHXxLtXpsy1oP0QfVPNyzntjOH+s3mKUbuZy24X0XjVRcy4liM3zLp+DweINRo71S7Ree9F8mx/6Yq6sTXG7x82X8B9cIHU+l70/oyt7mK2Ad8SUX7X3CIdnETDUVQ32ZvRaLxogjDJXvZzCbYe2U7dO1zA6oaRFvRu7z0uN6cXpYw/c7fYRR2jA6hRfR6tSnGGOA/wAmU0uGdnLdqXDXXDtd7P0RvUs3zGuDGxpsA/rEl3Z+O7XMS7mJd0CjQcaDcGQU4d+4f6OQOTi0OmkNIowImAwmAwp1CKgCKcM6FvEsZDGjZxrYCVARThEGIgtEsImARLMBNTCwUsImAGhZxoTgqhQyovPAs5hQLRXXtoxAHogvNiSnZ0bONBWolGzjSzSSYA0sImBEBFQATazCuQYDWN4kamIkFs6FvEllg0bONazBjxVGCIDh689QiYWIlgwgIgwRAEQWiAKHsd1blmN2a6yIlmGEwsRAMJZgiDDCAIgGEARMLEQWYQBEAyHsd1LA6nMRMAiYWIgtEwCILRBgaAIgwOs4/oQnKVjONWDVWRmdZYlhvTqPDuLDc+d+7TEuXGr6/ogN7G83CNeJnkOYn7nc+GidWvxLi7u/tKvsWfUh+eGqt9FS1Rl9oY4yTZ6iF76Et0CHQWJvHTPDXdxn/qq7CGZy+BL5E+RiLRd8ytINNOf2r5+/xjy0XD91AXF4xSaz0NggaKxpkZD6EFgKHcR5Z/lMmPhlggdMGw5xizeg01lv7SHJgyJMcXq6ssJw+/Z13JExDZMEozGJSQiC0RvYCM6uuatrEvFRaDiOYvhtv/nTzFIBkXnjvcl/pPiX2fcVFojdxSxaehCy8rNtlEqDmMuXKTpxBSOwPwES4JIt9HnRtgxGv8d5JDuz6b1I3ont868VGuacxxw/c7ePELSz+p1F2uOxVh+oGpO9EnvUapSXHBL65bDz6iyftHIdvMZ8DqP4lzBkTl9FBmhmBODQjh34L/RyB/P6nVqRvQVDRkPY7qnajcdTnGEwCJgMKcIgB0beJMd0eUV17aWEorr20OMEwIlhEBEBKPi+lMd0DQ4GmB7Rs40BhReeBZzAYNCYE1ZMeoAiCzCE4aoUBqcF6NvEgA0LONLDOjZxpgCozuJUBEsBUV17aW18LBDKITiUfF9KWoe0YPGgBIUOHoCJZgiYYJR8X0rOGcCLLQiC0QETAYQWiDDCWYImJxEAUPY7qwYOuGCIBhAEQDAez3EeCxEAwmFiIBhLBhLMRMLEQHtx1TFux2cqZrrCILDQYxrfPCgxjWN4lXqAKwMMJ/GjTrAEDHEEbymvOfOWqZ9EOvTXqBH8EgTIie/f6QRU2Tv/APk3fe8wiIZmp8AX+/RcL50aY/gt6aNw+WYS32Pgo9k7fPPdvrt5dvXxff3a9fBE1Pmqcb48+X64xWzlNJEVcv8A1SMZ5/FXWVbMUWlGECUghlt5RTltFKi4hZdWGFROd/cD2bxVeC1BFy4QWX0ogtd+0I3dTXe9zuqWH7h92Ahc0djLUIybBjTMAjSB29sLimImFXzIXDfoXMw6tWhmMZRwTH/J11nTyZgrZnhBvToEhvk0CcOoyhfRM1opedew/aM6tJVLN8Wqy8AnzFQjYNNMNMwaDYP/AJpeesbJLUyRiL2MtRrBbuMwtROrM/4N1eV0ZgeO2hDja5ckWLFCVcIC/wBJdd1l6UJF5QtqSvhZrnG2+uLx34nMdR8FVsvsTUQeiiape8UaFCL6c2NmSFHNxVy2zd1fwCbvqLbmNt5n410ozIql+sWpy1Rl7PVMybz53pI405uCyiEPJ0kQm49Zd3HrOZUZ0TmNG/h7HdQ75Nrz1IiAcTAIgIqAIowGgCJgDVADM2e4ppd8MqFnGsZGWSY4IlgOjbxLuZAdG3iXAIrAinAVIalgKiokBK3zwoCwB2O6h2RKLzwLeUylVKgNLAaWGQSAMlgNWAvRt4lOC6WEQETAIlgOj4vpQHC6NvElvQETAIgIhOYQEQYiWBEwtEARBhhAMILRAED2e4szs3GxhZdED2e4gwRAMIAiAYQWImAwllmEswRMLEQETAIgCIBdBbCtufaQbmQlYmcY1UerkvHakGW35U/xpxEIs/HtQIOPE8Y9793zfMK/CKZiH5Y6rv0Vq/7qgplfCy6e4luFOMUPBxolinXw+b9fd/G+Lcqluapl8bzdffm6bou/Ki8at32/piVSB0fLv1G1KyEnsoTJK2wcqxktMt8oIHrQFYrOzMUtP5/4vxchta1ljEBtnOMCaGbYdnHQMQabGJmZu23/AGZAxAevTTWbnQ1gyJLcXjAb4wQFw/w+1EnbtdwzM4UvTkMzBlDh+4ZudcUbtnlqZYw0twzov/6wGtEFTELwyL0hbktpth92tbFFiy9d7Cra4+XfwIGy3g4ZZQeCMtHirXHqkMs3jpQ6qcgBhiCRfTSH7m3xD/bHS7s0sIvMdW2jLa410cv8gfOM2WjLBiRL6jITMlFGGPwn+3vhyM089YtZKbRMU9iTEXlTMz92jrc7VBJuaXCDYRLdhbPn+Q24/wBS7XXVxec1c19bU1zkDOslXUQhMVhRFk+btb3sBNxt3wbr1q5oH7V+h1+iW3qfRBpcfXEKZrTOcJH6bwl/sPu2LhLxlXE5fTYRuDosUZhsI31GqR02gIgIgGEKBEFogIgwNUBFOA1QDCnBijbxLmqcOjbxLpiUfF9KAyovPAhnIyGQ6NvEuagoaEExYUBZagJlqADReeBOzAK0fF9KcA1OAAupaqAPRt4lGETDEQWiWA0AurAYUYLqwOHrz3oImBFQWIpwIlhEBEwGEFogCICJagwmJxEAUPY7qxO5jNbMESwYQBEAwgsRBgiCzCYWYSwIgCJgEQBEANAL1nH9CAQjMzSxLIb+NTQcOAC40h+QrEz899XZ6MILDRDr32pvY4Zc5rX+Ifguf37pXRH1Jqr+j8wb6U3zHF7aZIwbbPn795j85avFcictnGcSpjZguTv2gdDUQ1jLvPChpZa8mBsS3MRKujMaMabW7CYNIgYwOjDcGm5RboZUtW1DZ4IKt6GGwQFoHGTJZt2QYEe24R2NDOMq0NhhpuCisEErQIOi0z+8IBkSDBmGODQ45/eF1nMrwuEUmVlm23IWa65lr2RZRhzHqrKnXCzg5WsxmAJtu43hDkttjBYmEKMQi3cPrDH72oZ+p6iEwll5HiXO3qoLcuECCQjQv9JJxgtqRgZdZWmMx79LxLgh0aMCsKM5ZzJ0Rke6jAYjHJrAq5y52hi9BPhdRW+qFPgoUGmM1++ocUC/xdpZ+45v16BhZXhb/d8rUv32ITfnvMTVQRWEk4WEb78C7uPX3CwoicP6MNRRqvpN1ZN44K+1LVOBFmNdOpohFTohHwPgXXrV5WMOxOYdfCUbRpWARThFQDiWFhRt4lNqGNF54FrUMlwF0wxEsJR8X0rWAIslnEsImBEBKM7iQC6oCwU4YmqqNgRWQiATSzA1OYQNCQGaYDFY3iSwDXM40AqgImAmlgNWAII1RgVWBxNee9BlReeBdyGVHxfStls6FvEgPUwtEswRBYiAiAYo28SA9o2caGsSzQyIgwRAESzBEAwgCIAoex3UwmN2aW0YTCzCAKHsd1AZoAiAlZx/QgIgIhOp5nnGDSXLB00TQaO4BAHtYgRmrJ3cKvsT+vR+PvolPorMzX6Yy+kK9SzWmWx8UO9zbx72S7VqeqrOj4pi8Wqi3MyRjHYjHpsvO3arOF9cqLh0gJrKJxmHBGcXNZdiGo1lb6lOq8w6y0Lz2F0FaxnGuOlyy1zcEqtqM0s9UIJsoj5RTssgwjG9BjFp00HBI1wFDmYbRCBKsR+IyncdnXY+iWVYXB9Zy8Jh1u440bNZyvIQJCCxNCsXyNGZ3DMlpjTKwxi4ozJNsFYboa89REnudkz22LuBo2AOEF0jkyiXDzIkuM4GgNjEl8ukfCUdst9aftVBzpll7izKOsdqoIxJxbfVWcSFHGJbLg4bTGKwRu1k1oVFWhZ+oWitVzBZ0MjJtGbT/drrkxD9G/QUb+B0rX3RoVKxuWv8S/cbW+C3v49z61QzrCmjd+0EFjQcaDcGhHYh/o6lHhaBIE7HlG6ImA4Hs9xY/KDC4BFQA0sIgF8tTFBhLLDVBawUYEQDtIEs9UAJaAVC3iTA8ovPAnZgGlCCtF54E3MgFbBdMBY3qmJZgCYWHRt4lGoe0gSAzo28SNAHRt4kxOTMTpajcmtHPaNnGkYkBLidEtQ42hQiDEQWioLRLMRLAiAIgDUbPbZ/YgDICIAaDBEswRAEQEQDCAIgGEFiJgESwYQBEwCILEQA0BjW+eFACMjQYQdZGzbABwO9ekP+x3HVqxM/Lz0RPVsTlqgTn0gSZXwmVACMnh7W9cHm+Pu89PzCuiMEVVZ0fnLN7YQIW+rAs8/tbfevhrs4LakZLZczQd9MkZwghOMw/wA3a94rSNlbLMGlpoevLAbd83MOCEOzM5Doww7dtDb8Hf7X6dA3KTPBjBGYTG/s6BEtZSzmdMX5tRiWukKlZ7bEDCxEFEb0GtUs1VTu0vhJXMjLMjasuZPhwWDBtytsQXdGcycluVQYuXkdQs+8tcHzaS3cO9kUJYhiG29ttFMu+8Yc7NcQi8WVFi8jglh+KoeqFPayuC9TRPYnXgJ+/c9gXeodpLETU/liMyRkQc8uOsRVEH9pLapPvFxjdbHHIEKCKsKe0mW0/kKaW2jLgmPcbwtdbvZtlg+pZhLcraxn3i51w72bYYPqSxTWVgYPh1r1IHZKwzUfTnB4xViAf6rtd9QnspV5eoxjTC6wyCkWCuzCXpaLfg1MgsIhD5jYKtTEYTzGHzlN16kGEFPsEE0J/wBFiu1jZNlosZgoZewDYK0iG23hb/M+amyfXE/SDHLF646I5CHd33PeO/8A8gbVAyXGHHPHLMHiwT/HP2FC2ds8+HZ/uqPTLXVL7y1NPo3Wprv3NBg18WHc6Rr/AHSy0KG/j5j4ajxo3FcTu+zYNM8GmeDOJog0dhx8Jf41wfDibUZ78S7drjfhahKR1YA7HdQDSAGmANUAwowIGh3wIhxlQs40BigIlgwmBFQDCjCJgRADSwiYCuQ+eBd1BVcDyhbxJge0beJK1BaiuvbTTAEwxFOWXTACasxsYCsgmhOioUORrz3oMqFnGu5kt5Rs41tnEpRs41vVkOjZxrAZoAiAIgIlmCIMRMLRLCJhYiWoRAMB7PcQBEARAEQDCAIgCJhYiWBEwtEANMCKgPnD0RC/mZJsmQm9bBjsdHccR/l8wvQeTVL81tUbff50pEZMmsmOxgsJcZvGXeef/wAPublb2Svk0W+rIZYjGzIHF34Vva60gWfTDln3VKk3DW5xvkiRacOeQyVnDkJxoEBqcndfA71cdwFLcZjM+zjlRoEKcvrS3yfFi+frVLMQPC7lt6JSP2FhW7mFY3e7V4sRWf2rXIwPGJmZWRmH45/uhY65URbRCMvYRgzZZYLEcaVvu2DAl5GLmaGx++/ZkmObxOx/utfhanWL0dYbBCMP61v16vqd7tk2JeijBmRwaVX6hmuqqdSot23S1qapjMMch7/734xdjiYHuuW4/kQkMTI8nON3gcfNLXqxge626yFqUZmnKMOWMlSINc7RTw2z8+9UHvNVFhMu/wAhehyXcIC5574R2tIPtt8+bWPUURYfV3K8BqJYRCfznOlQmEuM6O41syl72e7uPT8xce5uPGXn7bPQp4eHS/8Aw+wZmhAj26GumFWZqSIOZlhkCbbrOT+iCoepak7dsDcf3ZZ9RR2MIZqUJMo8IQVh+0rnvVz3dAQepZF4FiP1LvqjsW2A3mRILkVDb/J1rh8Qm7sl5BLy4lZWmg/Jx16NHEefd2RuM6nxsYD0FUZeRMZcbvtale7uDX9XA7fJ3mTq+JRzD4X1QupjbBRKsQKwsM+433+BbRzS+RJ8kcqElv8AI8SrJvsJoiVDrP0mOo6d/k/Bs0tETOrTw6wLYOVrrcb2sYMFizWNMsbDb0M1Q+7vQ1/RBb7Wp9vrgyG+vqjuIVHSLJ1BozjYSW89w+3i7uuouSPWXXVKOYaiZftZekvnSzfoktxfEkk7EP8AEkDkZwQm46sd98O5+uovJ8atxCXJ2dPJIEQBEARLCIBxALqwIvPBgNakCLQRAESwIgImB5Qt4lnAY0bONGYAZoS0AKLzwJgDovPAmDDFLAaoBU1TtxsCluhoML0fF9KYAKNnGu4lzUIzZ7iUQTS1Dl6FAaWoRBYiYW8orr20B7Rt4kBEBEswRADQYIgtEGCILRBgiAIgCICIAiAIgCIAiYBEssRMLRADrOP6EAoaaHR7w4p7UgjeeYViZ+WOqAv+mX2tUTFppEzIR7ugZnXbp3cO/T+KqLyNHlxOr5I1cU3tnGDuRIObbPqd3b9id9WuZxvB28vmCiCDCrV6rOdVgEFBTWvzjP2cdeefOYGliXza3IwbC39tqx1zK2baHcpOvLmReDhBwYGwrs+vN6tHq9rMu5SDqLYyXYGFwPErzYoph6UUS69LWoohDbCsgdh+teZFFMbQ9SOGuzLwDZN0OBEP/k8NXMQo6MNNmW8jfNjJlbK97Yg9xvFM7dpc8znxBs8GqdoXV7TUFx26L14jMpxB++f7QRinbrv0RzV2nlsukwb0PK6gweE1g7hzwcfFrcREaQX6LqN77UK3s5asMMqD26himmnaDo4UQ6zBr0gcGDogwrDs6v8AVdiybRBL2MFYbWRvHv8Ac9RtKj9QuLNtGsoS46EZBggtwoAOszfbb/asqAtZQlX1Sbh5rIFwJYyoH51wvNi1mSOmRucgL20wjB0KCJiBbBwcLDstXo+o83s1NO0khmBVrPN2qHjPkHVNXmWk3RwrAMQcPpC9CEVUPz01Rt7iiM6IWZcL0nkzD5pLD1pZFIOW3HNtHXjrn1ZahLcIDmSMUYYNgyn+d5hVw7Ojp0o3lYQyXX0+RdmRYx1iM467xMxDz8tpmK91Lhd7lzMcICx0EhGXAuMZa+4u/wCK4U5T7O9CL1cU+3oY9A5bn05vOpFnFjdm7676i4bd/DcdR8VRbq6Zw/ZII0IwKtCyhwRo5G/Lz1poPZ7irlyDwOx3VNO7oyyBqLzwIDOjZxrmQyoW8SaBUwCKMIgImBEARARLBijbxLeoRdCIBdADVgReeC6YC6YATUsAoAaDC9G3iTHcgGbHdWJ3bKGpQL0XngQHLUt6D2jZxpidmlmIgIgIgIgIgIgxjReeBASi88CAyo28SAIgGEANAEQEQBEBEARARAMIDIJdncsZbLDQAjeqYgOI6ua+yHen1P8AFbnd8XbSD9iG267+JzC9mN0VU4h+VLdeBJcfTLt0VPJx7/F4u79Pzdx9hJwkcBv6z20qY3EtSEC/f7Uc/VLmjjUYg8YMjOs2DdHQVJ2mG8yhedLi0Xcta3l2LxuqqYxKnL6WvO6kGMTi0G5gsDt9/wAAy8+L7R6EcOZfe14HUNwiUIQ4DjQNu/cWW5liZXU0Yd1li8vBg/YKwXlPcimFv+TqCe21eer6YZhyuF7STquwtAoKHwFDuIG52Msra/8AZ9rVZ2NTxkGDLUm5UxEmoNBgg0CIwtAw110/Rs40ag2GEukPNZmcf9iY4yo2caWAKLzwLOYUPaK69tVHfC9TCTiE4aYnXQQWHoMWk47Qw26avSec5NfqvXhGBPzWA5jR16OZeXMPzi1Vt7XmTTsIWfxo6dEoKofCd9GXC5dmKr2leg8vZyuDlly3MrkJtzYWL95Tv+x3fVq0jw3JsxlSiY+lszE23lXd31CkzgRl0KRZ5pIQ+i4eJsXFK/792tlK2RJgm6RYQFByjbFy5PtQXHY7v3C7FkR1S/bH0KXVeiaom9q29zHDumsBhzun7K7/AJV50xiXqU1Zh9hIbNA7HdUs7g0sg4gIgBpgEQBFQESwiAIpwGlhEA5WN4lvUA1vnhXMwGKyBEwE1QBEsBpYRAJpYDTAiAVovPAqAVU5hQzY7qWYWQHL0t6AiAiAyovPAjMuZhlRt4kaMIgBphiJYRADQEQEQBEBEARARAEQEQBEBEARAETAiCxEFk0BEB8BeiXX5hJqiRrGNtgYQPrcO47Jac3d838BepDwp2fFl94uLwiUX0HjBm3u3rh/tYry79PduLj4vrlYw+ctiLvgw8Q5fP7V+/cfYSzj0uSoFFpjpJbBx1v7Xn1KgjmETOIU9Mv0L1CPofJ99F4DcEg58jSFHssppy/Sq9jqUb316SDUcGCwv2j6Qn9llZExDa+djIstYvPelgiaE0LLjQF5j02uUfF9Km0eroQM2O6vPejO46GTlG3iTCzFG3iWdAsQw11yIWwQS46sAglog6hMHRt4lYBQwkwCaytWMqCpkFQFTRXPtrig8GEupTNHxfSm6oVqEEuAfQjF6bz1PfDC6Sv1uE0vhXVayCGVdPzGBLUcTQmqzfndqipGawx8G0HPL0OW1vPvbXEvmSZKSro6LHZpem8CF1GIPz3tghhZuOcOKV9+54vUKlleypSS7FzQyzHDm2cO8Q/+33qVeF0rKEwhsXLrKy21rxTlx8dTk6Pon0PHVLG6nO/kBHym5ExuPf7079elzqfE4fvJBTYKYG4NgmYfjunzjk7t3aXCwaeCUMqjS6Doex3UBmgIgBoAtYzjTAlYzjTAOlg4lmJkSzqWGtANAJoAiWBEwBpgRADSwiAGmANLBdUBEsFjQktvYpR8X0rmrpWi88CwY5vReeBChilmBoMRMCITolqEQA0ARARADQEQEQEQEQEQBEBEARAEQEQBEwCJZaGJhZNAa3fcnQO99JcWmjbwB+l/ynaFbw0t4/MPVOy2WEJBJa56q6KGv3jogHfSLT07993/AMVelDypclv7ClRYyFyEIY4c4KgVx8nuPd/D9fzXZEQ44oLCYTCISbV7TaUPZXnu++9as1nWruHobup5Mn85xF5lCyKo3OoY0jVXTGZfuvqQ7w4d7KS+aPgNhUY4fk97WJnJ0Rh0ONhB4hixClqsbCCrcLcwkZew1yN5aFsrwnqNRo8GhJG71sZIrzVw4XVNS1cbLcILD0GJhCwovPAmPPFDDQFuErJ3CyTE6JYRUFvOjwFq5o5o8wf0esGmq0M3IUApRNCWFCuWwt9N/WhOODsd1dndOhpuHotXoJyczm1sFf4WqZO+d7+0ltjIb9lCuRobVD8/dUNe1qmnGZh/tCv4aW9h+e98SWy4RMj8xmfX0VMeXydyvpPMEbe4hZlFbvq+yt+ye4XoIAL4wkXwNmQPZfOLJx9tTswavQXxNeLl82Y/AP1yLNmasN2l2bum2s+IsTbPH71zC6H7rehlX5Tb+GpMgUyGnW7+DkEwh/4F5huPqKGdzY2fQ6hVnA9nuKgLRThEBEBEBEARAEQDAez3FzyDCwA0wwurC0UYDQEQETAyovPAmBkpwXTA9o2cadiQFR8X0pICoW8SA9SwUM2O6uzuAlwK9Ac3S3oMKNnGl5kwsgwRBbGi88CMyGdGzjXQ8XAGgxEANARARARARADQEQBEARARAEQBEARADQWxNXZ3cjYFbYfOurmvtFweDuL18HaO/fxbKzyCNyD3HUWPxrtexYouI/L6b7+5cJi8bmWtrn2MhUCcb076t8U++xcr0cPKmWrwmb2iQmJz5F/ZS0rnL/OOh33rLj3HpPSfFUzRCEyhz3y5WF5E+fOHhVDT4sW2ecwM4uPi3Hp1is611fpj6FteANfw2BQaCg27Wk2r9+zOPVHJ1MP1kZBNZAhwa8nED2JHm7VPEdsmrTSbhSOI9qzaDGjV4T3JauabgXkvZIm5GEvPUKoPY7q55ULMMNYPXgez3ExAsKPi+lbxJOBqLzwKlzBwILD0GJjJqj4vpRoVgWjZxoOZULeJUFsggsK2WNrIEbuBdA2stFuBaCpNCUKhX0V17axmFA4RtCu7liVlYtFhVjONMJzJCNVvqrjLTJ1g2vQeFaTTGXx5qmb2RlzB35lDmLVXxKeuNH5eaq2Ty5bmN+XBuVXrcujR8ze46nJ5ci8YEjGs5mJcP7N64cb0Q56i7XvPHOhz4UYGcJGArbARj+xcx7hR5+ruzW2xguDzM4ZlH+mrHPDs8uUcxwhxGA24l+4zHn1CnKfrP6AJOLSr1U9yyadk8KmADEbZjnbXdp8TFqKrDdL9AQdEcJcrZ2WgRqQDgRqYD6nCICIAaoAinAiAYrOP6FzUGaxnGuYgBLQKUVz7a5oBKNnGuh7ReeBZxADw/wBIJwGSwIgBqwBpYRLAiWCaYA0AMzZ7iAXQA0sOXpb0EQA0BhRs41zLuHiw4iDA0B7Rs40AJMCJZiILSjbxI0bwlG3iRowxovPAjQMUBEGPaNnGgPEARAEQEQBEANMLLpZjD1eXTE782NX7f8EmnVGR2EMBwQqVIC8Ya42z3Fxcd8+XuPImdX53zfMcxxcSr25zktuP8Pz5tMTFy5uMKi7kPaXOf7K8uHbtY3O0dwvRFTHMcXoy8TWv7Ww+x9RS8Rq0ftr6Fpe9Lg8nAzvGm2GBuT73Z74sUbm3UaPquaI36mTqriCzc7mk/Iuh3d8Xj8S+w+stKdHPTTV4T01TWM9pn9q8p7BMwxcUCBZbsIl1ZBaY4/WuuL4Jta1BBpWFGwdnuJ07FQdWSDCYDFG3iVAZUXngXcSGS2WNRG+bFvR3UZYbV5mz3FiNwrTVEplVmrEbnsAjVE6Lu1W+EvkQxN8kKAzYoqBaDkN/aS6yWH9EtQkqh+ZGrKvK1Mzvy6LEeovU5Zf5h4t/w9XxlfSkSLy7F32ALMY1wvoqHy1y0UJlFsboxysclbmFweZoQFBw247c79A2breim+LyjF3MILzNvj1r/ad+mfoKTSw9UFNVyIZbQs6XW1DjevX/AO1eTOsG0v1UCyywN39KWrEPY7qQ6bSwYDVUbDysFKA0wIgCJYEVgReeDCYDCDC6Cw6xvEgMKxnGmAJMBhLAanD2tufaW9AzTQIlgRALoAaWC6sDGi88C7mQHReeBS5DltG3iWHoJR8X0od0YUbONDjxLAdG3iU5j2jZxqgt4mANADSzA0BEARMLDSzETC0SzEQYGgBoBhARAEQEQA0wsAzY7qWYFW5hjf8AmmPPfkdquS2wi/jOVZnzYy16/wAm2u46he48+d3yZrTCRJcciFm20TNqSjn+1gcx1HfXS6zorb3MjVUWCLLMcWL+zt3G2enXBu+ttSTerLup8OmSMYg2FAZA430h96S4UfEV2cav2n1KIht729HCQ67Bk7qoVDcbNzmg1YrVWbTo6YaYzuLwb57lLUjDaNeY9pUVjTGLydHsMDTQtlqy3oaDTHVwrC14Hs9xd8JzQSbJU7HNCVJQywSsA9nuLEhaKsIZudZ8hjXdppmP8Fma4PiVROILUQOCur1Oo0wW3Z3Ejw0rzNjuqJQpDViN1AKUWGqCxEFhm5by62WRmmVgjII/axaTvifVY3mQ4tCYqIzb3GIf708VHCeTew/PbVB3q6Mw5hm0uLLwd2voJfPzHl8xiSGUWIwzgI69t52VTR5a+NDxFhtCBk7LkXLwVngl0iYfqT6BjMglzqhnm0vn8IHxG+qIyl+tMF0Jx2DEqJbC6DNUjphWA4pwbrbr2koPVSBFOESwImARLBgPZ7i35CLoK1vnhVAYpYDTAIlg4lhEApRs41ViQzSQIlhEAuqAGgCJYRTgumBzWi88CVl6OGNG3iXXA0ANLAaAiYA0BEBhRs40BhReeBASi88CWEovPAgApgEQESzA0GBoAaAIgCICICICJicoZsd1LUFciw5bmEx578ltU/CZw5/5qjJoWOOjFk1xvXuPqL6HofPzd5fIBcJEErRCzKKxqLdw4xjx77hLXLORIQ2UZPguEJ+4mGKH498/2r1jm4/fUlMYhp+gfofN7PXqaX0ZMM11sCLIc7fXnUfVUPFW2W79Wr0GRQXs7gaxHVDcBTvG8GCu3Ol1vQtGqtMCN6G3rzeI9Kza5HOicQGvneI+itFHWM41O9bMlkMnoKYhuJbSF1TVqflIXYew39S0RG6yD6OhY9egmMAgqgsRUFrAIIL1UsHAqL1V0CLZaILMIAaFCvU5hQzY7qnMUhqxG6gFTJ0QCaDBglQRDI03IsDE8pxy+3LARlfWg58dO4bzLx8CapO9qGU1wIEDYBHQey8JcenXs8N8zeQ+SZul2kLfUgVg4fP7Vw4XtvLc+mOXZFFhL4vdmMWyGpRaXCpDg4QkYxL7SnHZebXdk77o9Aqi5kX1VUA7O4OdP/7s1Q5zSqiNX7ZiaClxuthaBGqd1brARUAwlgwlgwgIlgRMAdZx/Qs4AwRqAyNNwLQLoAiAyoWcaYGSWESwImBEANLAiYEQC6WBqLzwLWIDFaAaYA0wOdqNQiAXQoYUbONACQA0ANARAEQA0sIpzEQGJqpjcsFcMRADQYiAiAiAiAGgIgIgBG9UxMTq8vZXeW7J75+auqug0Ik2dZyuTI6Oe+cREl6/aPjLXefFXsPm3w5rTrQI5mQszPWjrED4y09x7m493dJ6xtF7mEFxeenJZZls+BAr3D/OWWL/AHcXcKKdlD9KPQ6JONhMKciUWIBxtvvxN3tfiLz6tl1L77gZoQQWELliMm8n/OsWLdSkjZuWvzf+pSHsZa/W1hi856apM/8Abrz3seVEZss/Wo3reGC89o0Hka0G2QUys8AvQynXQez+zJ07lrcLqmqmdizoez3Ezwg8HwdnuIlw1/zzP8Ct8g2tAqasTuBQqLCu+QLkfAUaggYakmE6wHjU5hOtufaQ38SqSzQ0xOijMDDDXGGdY2tRoFIZp3cTWI3axPobTA34Sxysm/fHd/GTWFxoGV8FhYVLrx8yve4T5G8h8ZX95QoxAmCY6xqfrr34eC4XN0uiYmrark6qaJ0oNEi5uJzTi3ziY899q+ggy4yE6qqWg6y3twDmOOhadWM3mF5+NHpU7v2c7R+leeuWAaAtA9nuIBhAMIBgNc8g4uhEARAZUXngWcQDVG3iXdQUM2O6ugsqANXM40sGAk6QCcbwJTBlW+eFYzAZLQRAY1vnhWcwGdYzjRiAOtBEAugIgMqLzwLOIAKuDn689QXTDA0sBoCIAaAiDA6Pi+lARADQESwIlmF0sIqCwaLzwIzJgKYWIlmCIAaDEQA0ANMCJZZU1MKnYo2jD01a4ZF4/NXVu1hjyOGCQSHsfVD2IP3EP2l5UWFnd836S7u+92tfRQ+PfK83xcSLy5BA88HC3FqcDT4soh96e7u/h8zcKN6Terz97cuXGa8RiC2L+Kv7KAg8s85vm7v4Du4s7hec9GH6d6lSDBiQeEwYPQXDXth9j+dIepGz6aNNooJQhYiw7ZS1WVMboX/UVKUpa+YHRt39SKlEfst/WvG4mz6C03Vqw9ASjbxKdboImFLyDfIFtO2ANWhch7HdTE58PZ7i1Gyc4nJxKw7iTsSBq7IqHBuZaYyUDNXGxV0BoDDTf1Lz1DCjD9y3+xWAFRgvRt4lrEM5kGi88C8/MGZDovPAn5ZImbPcSgVNQYrTNnuJZiljPqrfD2IvHzhqjZMMYbWBA7nyjz+Cva4T5m9h8d305aLDhBpm8Ypexw3zd3D5pnmEFiCOYw0LP2luvQl5jTYsYJMUxhQcTb3Dy3cdkVid9DehKGBlX9IQLGTrBwDGXRbX4+c/2KLwpjd+8YQXq/4jfuyJa5Yh7HdUv5nTISyDyAYQGVb54UGDVnH9CxmRqbrbr2lot6sBZhGrYQ01AKLoDyJADrG8S1mADXGcaoxIGUoHome97VwPVgCJgY0XngXMhisAwmANADVgRThEsIlhoaYC6oUIlqA1OWiAiCy6DEQYGgMaLzwIDFARARLAaYA0BEBEANLMRBgaYA0sImFl0sAmphavONost3jSFrhp7x+eF+ISLzLe2mOMtYO4fOYgSU4f9jcvMS4/fX0T4989wiRBIvCIWWYFiQgHkVfuN6d2dgNceT8oo3pupXqoPGJyvrgh1pFEDTCuH+104zvmH31/TqCVL9AtTBBQtxHWFhaqZ7kO7tNrOX+TYxLUhRuNhUQ4PhlhTMw0+NbK856Mq+jM8AlvQwwoWcajVYSi88CnPHo2casLbAEFmMKYWtQoew1vQVfYOdWVoHBcwxCjCxoW8S3iEBqjZxrSdnR8X0od0RMcEo28SNA8oW8SdmDcM6FnGjBTKjO4lw3BKhNNXC8MaD4TUzsWOqCZtyzCpOwd6gg7llb0T13sB1GzICbRVzQLeo4MudhA6msmwQ31cQocqVYavMegrUAmYHkfQQGhXzZYrIM/b2uvZ4b5y8fCV/qWi2GPww9vzD/si9nhvmbx8xX34QWIY+Di/JL04eLLllWJCCwmtzznPrad9JehIRoOWdVLB42adYOXBNlbtzff/AS/Cmnd+6wXaWY3OlLVgEpHRkGGEswRMLLmJZvhEOGFOWYCN9RqsLWlZx/Qs4ARpq62wC4zkMIuhK2iXQmvoS2DGRJuoEUQRMCJYETAIgIgE0ANUAGuZxrXxDEGlKErOP6FnAaWsqA0wBoMRLMDSzETE4aAiCy6DEQA0ANBgiAGgIgIgIgBoCIAaWYGgBoAaYWVNS1DXJ/MCgslRaNGtt7CHPKgdO5a8++fnjPsYN/PGDGZDrtCHZR7h/tXMPMS4+Pdeu5Neu+bcoEEi4chms1kcW0VprBxvQ9xmbj99cPfTN6iWg4zCAZZaCQDCoVQ4R9sejWeVXfhLv6q83iaPSs315eUowstNp7d+MvK4n3j3bLSzdHNjRhumn/KEtUraxodvgAVZisDD16M6G0Lz1a51maFluDHodUEzxoOCt09LXzOFeyMhG9EI5DmRw40rGW3BGhZO1DWfLZoKHWhK6cpV1QsNCrWMXOiGu8TIlxlEsxEsscIJBZmjbxKxnECh7HdQI3eIaWWRBW7DV1K1mNXxIIEuHzLUo1fRgoemnjuH/yna1YzmFKbfflg02ir7BRjMFDb4UFraLAh3MGQr7zYLGnAVdiH6hwpy2TXoKNVBpq0TpLXY2Fh0JedxHo2iiXmPQMIDTb4YbTIM/aqY3eNOz441UUHo2uIwF/2l9Nw3zF5D5fvuiCUb7XjH22NcL0rJ4lT5nm4svXZ8WIHn8/yavTO1ehvl/8An5CSxMy+iDp1YKPw7G79+pX6yw/5OvPXr8JAPIByjbxLmoEXQWNCXDAqLzwLmQCsmCIB2sZxrepbxYMQIL1WpicwlhjReeBNzIEovPAjMg1R8X0rGQi0EVAESwiAGmATLVOA0ANUAnu1ThYIAaA1SjZxqhQEowGrAIowGmANLAaAiAGgIgIgA0XngQYyo28SDA0ANARLAanMDVBaJgRARADQC6AXSwr43W0b+iAHt6Z7T1G+b5zCdw094/NXVdSXBrzZ50sQa7fzJWgPK9tTaWTz4a9eHzbyTpag84iQMsSNkPwnEHxD8gfJ3WMeXb675v4NnzFyolDtl7MzXg2EgtBsOk+u0QqM5lPp/T+Bs0jiPSs31NJUFDjYjg3DYWGkDjqTifePorP8G2KiufbUKwUKCGnK1I3GCyQFBQtmnfq/oTd25jfrvuhC3R5sqgkP37jrhtTt06zdw775QwpmYfOc53zmtCyw6w4RvawsmVYXqgIRBg8tjhH7OjDnUXL1WokG6ATSH7nf1Zhnqw3GS9VucwNwYZAyLDh4+bS8KYrdekrVDhl6GxEw31ZdcCviBhWATMHa6q63e0XwRoUa2FMTuIEbn1hQLu1Y8uGoLobj9a0ZGw9aytTEuquNjQQaxiT40a7Gp1NC9VXpZnDnt9C+9G4MG/ZBM/mhyOyXamwnmcOSTNOcZNbrKGdEX526CCf3LhdZy49GYzGoNBn8zzPNmIcEaPDsY8s/jrrM5g3IF8CMziG4o707HAL/ANkMYMT3l3drSeJy3IO+BBw8jMm2IQg5xtBGUvHvfqDCvLbQ41GYzYGm1FvnbemxazopbjBZoMXHVwLNGBlFhUHEhXZyYDrMK816LBLUEo2FWhP2mqh575b1RcmCMjT8Kh3Pajj/ABF6vCl4N7D451SMtlhi4GtzH7/rF7b5SXzFF4TSW3R/1f517CN1n0NaEUl+S529y+Id4hQ3saKuHu/fyCaGP8mUJ6zD2e4qFCwDNU6c3XB8SAZQHtYzjXMAsumIgCUbeJAe6ytXMDASwYz0P9aYWFW+eFcwWgRuDosVgM69MUgFrG8SAw16agH6zj+hZwBFoGEAurAijAaWETAV3d5+2ueAMugRAa2lqC6sAaWWiAGmGBqcIlgNARADQEQA0ANBgaWETDC6AiWYGmFogCICIAaAGlgiakxs1BIwJVJXwhqo72kZLOjsZawcF+/bj3A+beu7h5nPh3d11FyvYfNtRkJgn5KIHBoNkL+OxB7UOKnc1w7d3fMfA9I7efGukKHcbxkGMNmd+YbmDsa/f/HzNx8S5crype1D6dkqtg0r6Dn1NxHt2a6UKtfyvWhB1v8A1Gb7xe48tzq/VfrOg4j+ijxLh/ujk98V8I8y+Rr5t9syZzcMFqKFeD1vYm2a4ZMs5UfWPl848eJS7MtamOWpymXQ4GR/+mq3hsXcTJGD3mjDGVsaOsOwYx28SMoopbzJt6cwPoBRx/iFjMKYh0y9/BzIOZhwEYhYmtTFo7PBZqMND2Evv1fS6ZK074NliG4lt0rGNMatTglbVoRuhKtjRYbj7q4PCrNNwLTjVY2aGabgwLSdrkaMz4TGZ9ZyoajGpXabYG4MwRbLvhnClMvXtjMafm4MQ/HUuSMarIzUwwaMhuAjAcxo45G1fNr0ImXndMHzdTY0wJwbGj8u2gcfJXYfeO12uRaQgmptvfSuG/YcE1+e/wBh+wXGPu/eLOcqemFaZevjUGt6KPRFw47Xxv3n7i87L0MFQoKYxuBv/TO0eHWOs0Zgsa3+33RjFFpMO4bkHV7t+hZVGciUZoTGUWgsQW4Tql4LRdOlRyt4t/D421Q0HL15rQ9vcPPCr6Thvmbt8kT1LuC2LrbFzhzG9L0d3lvqL0JqQYTGr7oxQgNvlFk47z06xUqph+yIemZDmFCcsEKBEsDBIBpANUXngQBqNvErE4lHxfStYgaiKUGEArQmcaMmJReeBN1LE1mC9pJyC1EEEugDTUATIg0NbAqxlYLzwYVgErG8S1mALW3PtKTQA1zeJdD1ARADQA0sCJga+qFCKcB0beJGhYaAXVBiKcIlmIgtEANAY0LONAYoAaAGozA1QA1OoRUFomBEsBoCIMDQWAb1LFOYUNVBZQ7Z7iYnfHerbMmX8oUWDydxCtZ7JxlNrakdXd8xzHuXK9h820qTayWbzdHJeXPzofa27RnbyytnjvFuezXX1VzyoduvZ3MHgs6P73kGx+sOSRB/265du7s30/uHVo7c3Prbm6XlPbh9PwWChhQRw1gCTW9OzeGwRta4CCA0jtpYPwsDDGayv64HEee9rad8x3+mxqZ41RhnW9gRwmyS+YrbNpod6RpZrgyNAkOH/a5LvOJa7pb0HJdZYBUVv8oSsvSw3yVb0UrmhZb7a7w5Iu4bD+RWV/eIb+7O0zMoMEzbylEzIgB1nOqnpUkavXmMy6hUKiYa+271lM0H6Er1F/aL6CxwzEI9R3tHQJXmg31ej+tejjLz4lugSZJTZg9CViTwoI3sEIcayaFhXnvSnVUUbaz/AO6jaFCgoaANkQPRCWtU67gsa3DhVyaJbgELA6LY+UL0e+eb0tbjVGaZWqDwqnDU40FW5bn3+56lec9FQVhoXRaDYeSXXC9buFKMysQ8s01bMNmhhhtfsQ6VpGcOZ/YgYcs1R8Gq5MfKjhvEvdnwtfJyuDUeG3sGvMQvY4b5q7fJt+kouEW0XLt/9Reg8t9y+gSSyYaVFZmMAt3DB3eUby83zvvXIlul+oAa89adD2O6g3ws8iS20QDNWEh3JtMTMa3zwruJMOBGrOxZpdBdADTAcTAinBIzY0BcgFF0wHTUsxkmADXo1LDzX01WBYBGms6CE5hLBdLA4RvqNXQ9VYRTgRLAiAo0tQiYESwGmAugIlgNLWImJw0FiIAaAGgBoAaAiAGlmBpgEQWGlmBqdQXVhaJYDSwCallRsRMVhrhWrRvZc8stDxuCwK3OpyRCH+1uhl6vCfOXu75avHN1ojEVksNthCn8YpGkb1zDvmLjxruofeImcRLaPpW8VJYbQwTIK3TyHgnLDXfVvLv4b2zU3Ee1Zw+ugsigoFbwd192kvTVpoZpuydmO1nf20lQ1ydY1RhUTV1Q5abK5sbNUSgSNwYKCh4OJQcRTaFwpnDDakPXXoV9KChB5acO4ccIIKV3Wm7WBGatLU4y0ZRxm/NKrl/wdkRdr0eqXndvT9W3wbVD3sJ1C/NebIKf8nJdvVH1zW3FtTabHeeqBG3XM4bD5Qk9fD6+h6PaXfadw1S+KFBaOuCXnVvQtcNXlUwJreis0HXTdJW07uJnhO6BK2z3F7VWzz27dCi0Bek89TnbHdXhwsapNNEEG/NUXEUWjUdempr1Chk7hBt2P2dS+oT2rjt8nVmXqb20a/OWenAL/eBhnhTzxLhb4fLuJxEl5xbKzcImT0eXUxyYW/gwdzPZz9wRpEOgll9+vQ91V4yj96Wbb5o9HIlK9hBoHGL7d46+jLkJmTGwk+MS0M12W793cenXo18vrpedbXFnc7w2yQPRZtRtfZDrJYvxtcf1yM8GUM0zVVid1VPGottpdmlm+9JUzB1kFnyHv3HCBybVed1yui1iV9r2Eb0cNujrd7R5RNrVHq6sVSWZaF5dMYCo6KwBa1dcaNf2DCZLL5elw3l3uz4AvqhhwcM4NgWIYRk/ZVbw3zd2+cb8F6O+TPIsKluWZVrn8WPeWG9tdrHFuJojQ215VN0/Vf0L7UuxvUv3nLqFHAW7+IDu8D8cbM8x7tLouIhbc8mxGr6baFhNftXeHxCryzHC2GfrXo1vNstjqkUiJYYhIBpAEQDwOx3Vqdi4GrOP6FotEAuqAeCNwJQGWAEb1TFwEDQqxdBNAZYTffBLUJQs40O5F1lYnalC6CtFiJYJqwDaEpw9rbr2kN6PUMIlmCILVaoUIlmIgIgtEsMaFnGlu5BTG0QEQBEstEAugF0BEBEsMTtjuqeDI2KqwsRADQYiCy6DC6DHtFc+2k/EWWo28ScWXQoUc6wXnolc6CaPbj6RvOLT+G8a8fEYkmSbB5ZgQcltiD85/GSXtuRi3lm5ee47L9lehlI+g73oes00OIKF2rrP3jt3/qJHEejZ7vrE0I3EJOXpoYFR27KD+7o8hpcaDrNwLKhph1FBLeNnYMRo46896D5e1VerFg8gtfmxk79gHU3Erw9rgcCKXzHK+rJv43/Zx5zb34fO25f7vf5xa4U58IrniVVziGr+iNan+b72l6qFRiaL6kYipp7/AAv8BK9nicCih8xb80uruMvkbU/yfJRN9OCRqfZGImSCOYxVx6E65UxJcOc+nfO7i7u/gK7g9CTmc3ToY8Ckm+RqlzjdSYDNEmyedNA7qWIQRE8odDXbx3cMd3d3cJfRR1roubnt9H6k31L3OqN1BhYMb5+ohNl7h9n9cMYZCXnu/h3Nyl845NFPxLfZr2mm6/gRs6eGZNEaCrAjq8F+Pk5A+bpl8W+51kqHkeWoH+W7yvGug4YhqN3VpJbWBdFenLyYb2d1javdeY1w81uHuL5/iPbtGiTqbWBrz3oTs1GMsNCsKIBDcuG35YzOlY4gsFBIfvz7TW8fa7O42y7VnLy+LMti1PupYvYy0E/nS+hEJdOmM/PviW6J2NfScDiUPzHm9ld4fkhqnbzBd7S/FNQBmOcvYgTYcnaKv1suVcq6Zwpr8t9Gfr78VluWo3PM0HytKY9JLMImAmpu4c7s/Tu7i7WbbjZO5dyx93+h0XutT9DNTnMk1X6ZVcEMiz+yGAfj4z0m2KTg8SiuhnmVpdWtxlzabtSlfOhl8Nk66khsRlsHgDbSzXzvNeYVV1aPu7Pg1Wtvml9IXjr4WqhlcRw2+fAq9jjSH8Pzi+f5j6uHs8KeBG76mvczOHOoTg2u/wAxenw9Hi3U5b2EFBOHqjUo2YGHRZEmvPINCNCDQGr3wgvzXPNNW+HulvX59X5Q2lmv4MZmKj5pepDyZd71GOp8DiArg2NA5jGjkb0vF4svqrGl9AX1L5czwIJxBJWBqP8ALUXE5nHDfR21nN1B29NfbuZtMojALB+43Os8C79Q/mns/Fs6YzhoO3tX1HDflF4eTnnCJagwgCICIBoHY7q3BU7DLrCKwDUQXmxRZAORKvQLDQvPZUgLVta1P8AVJALQq3YXJCUXngRkMKJgTUGGV0tXG9UxVRsJ2epIY5D54FnUMVkF0tQIqCw1OYyrmcaHcAqxxEsCJYDTAIlgNLCJgDTAIlhEsBoCIBNBiILRARLAaYA0ANRmIrCw0ANBiIMKnbHdQBYLJEbnc2hggHygja3KE6/nS8+2SJW17OPh0QsNIHYVi86rLTiem5a2fdvjK+1LIf5TmmQXER1xETi2gEaNo7zm3nw7v3NwueodNnqudT5e9LEnVhhhw79/riLhynRArN3zHf3Xrk3iPPs932xRBayON4VCgjGggmhZE3/FI4m7tm0qN0YbMK8yt7Vo4dfsjRkaF1lgrV5tb17V88FahaIXzDOeiZoFbs4l5+YW+n1akDdR6Le/NwSYDrS/ce2qImFnp9M6Nb1RN7qZ7+Ug/k/vgnW+8P8AbE/144jx6uUTbRh89Sb6E0aZGKz8ow1j8nV3vEmLeX0fqafQxhrx0/Qq+b+Uhhz6E44ADW3J6la9ToFHL+5nOX1DfSvr3974Mlv5XjYMun5PY07IZjPJpVd1Cq19nZiAtRZe/nO9leJfyVfRYPbuIyTrP2Ea76h2kU6vX40TTRiXkawa9v6FQqZXcrbPcUZbrt7w2jTEcOl1jaL/AOy9x5uWjRw3okLyOI9W0aUdpndXjPVIhwWsCftVuWIjLn00Xr4LNEz89BpxDh+4HpNJxdmiNGZpiZy9AvKyxWbNuvQolLd0wCZqR738zGVsZkNwd+zKjMvI6Inwb/8ABdeYMMrecWDuP2dazJHRSzD1HMtQYytCzG8KDKnohuwd62Ch7LUtV0r+C3oZKDC6MC/aFvJGINBXkYIHloWIWugjuoX3OtBdCOBt0emO7CNDouiEtpw1YFNGgq0I8LtdTp3wBqmpMMg18xywPb/3EytBbaS+xtTfLBssXomnbe/xI6g4m762yjEN0u5XCgh7g07PqbiTmt7NpH2RV3xJJCgRriagsQ/cY5eV929iqO71dBl48I8FwcDuge2cL6bknE9R+X+09p2kLFeu8ARAEQWYUYML0E4NczjS1ODARrOAJhIyGBEsF8iT9QiaDCjBGhM41dmQsMu88KlDLLUAwgIGsyAjaLD0V0KhWgxRt4lFqDVF54EZBWj4vpVIDSzAdNU5jyjZxq/Ep9QlhQiWEQBEsIqAYU6cmlqETE4aYYiWWGlqEQYXQWIgIgIgBoCJYRTmBULeJVZgtgtANADSzCpq3Bstsm6INlKR3EEBZubph2ZQ3fH4lD6n2b5PZ3cPnuZJ0j0rsfhTcfgBP8ivm+Hziv8AO/RbzlFrH3DjE/zTSX34TM8aNa4cVDp1UdjXocPifG+UvLX7M7LqeJXCbNB1ZiLeNWI/KXeMceRs19PZT/DflU/jH1mbBKKwrf8ApV6h6hmjp1kNBb0+58Y6d+OisWbTzZXCN9gv2gj+BeY9prx8EleCswnA/Jx9sfLy+JOXs2mFJGzaLQgMfwhec9KWiGMMMMy1ebK6dS3R3ECqzzwd3G+BWHbCPUHaStA5ajUdZlseX1Hf9aLom1bJA73QQTK6g/vP8C1RQTd3WDU1YQtNz/yayUdbVo40YbWxl+vAevO7ZJXWqiXSpTXo8v2Q1OgGm5F+zr0vLxmmx1fP17PbtdmtnArynsChBYWEKhOb1lDM3Dn17CMUO96FW5FiH/ylDuDfOrO4WhR6HeU/DTNEGJe0c74esNumZhPqUo745pvWGw+Urz/jU/ZVoFJMd04w4f8AZ0fG7i0WwcFND6Fer90y0BNN5danBButE9pbTlTAgmW7EBTRoKjCcMahydFLu3IsOyhxxG/zejDjF8+EmMgVuxpC1XLFtGZfQUqQU4INxBAtwY5IsXqTsrJ8DjVFWs4QpeZ4e/ZHZrCZGpUcm4doXzvE+7ezaY7uVHeriFJAqHgBCd7O8TR5HthaaN9D2O6vtZ2flsbjrTKILEy1LLEVCdEGGEAyGbkXRXC2Nc3iXQSom+97EAQII1AWoXVNXAFg/o9dAehIBpUARABru01PiQWrGca5gPAvkCpAiYDCjBdLAdHxfSrsQNQzQvVYpQVrmcaWpwVQ4IgBqgCJYEQEQEQEQEQA1QA1OYXSzEQBEFogsRALoMDQWIlgNMAaAiAXMUfgweVwq2Njg8aq4h1ms5xwGnUOFfIc3+8fovJfwjjeqECCMyJebcPsuWQ+e5xliseayxrMH4od/vLxVcP7x4V5H2R2fUghRk2aoQZNANA+hER1up/ifv3Vx5JfccufkV5l9hm0QVvkI+kckvaeaqjAsFv/AIcf+R2s8RTZNPmmCmmYaLEON0LzK3t2bSDTYJBchCZbv+EL57V9G1uNZYZlp6SulSGysw1VZWzGRQpXYEpCcNggsGo+i1i6TENhDgYRoWx+z+EXudCTuz1brLoR1h+MkItmkzWY0xuWryK3v2suZRnIzNhYWTualgz2kOxOXSZW2O6qZ3efGzezDcj6Cujd5kqSIUXMtUHEehaKZTvQGg3qIZpbbAwgqLIsnXo0ILtYBBUKvS7HAdnuLU7J4AWih6Q1d+FjEiULeJR5bKUdF0f8QsKB+j74VCtYFuQjTcGQLPDSXgBoOY3OuMkTeqYlidlIYsyfKpmgKtNhBu8EJrzZbpe8N6Sv15D6SGc0QXXqCpMaPQkqK0M29ZgYzMYpd4nyOWc/a3PJaDNwRBgS8jk2z1+cOkL7x+OHEJ0QGQRuDosQGSAIgCJhZhUFiKcPaxnGuYCVjONGALW+eFGADlq6A6M7iQDCsAeWqMCIAZvyBVAPLeAsUoMZagMKNnGuZDPIl0BpgKmhVqe4lF54FHl1SqlQIlhEwHFGETAiYETAcU6cuhQiAiWC6AGgxEBEARBaIAaDA0FogBoAaAiAflXr0lcTZ6VkZi3Xoj9S+T5pu++sXJL8/X9x+peZx/vX2XK/wjjk7QWoD6Co+7eXMd27BqLD5XniNODaGoOqLF/2Kwd/9tfR8l4k8SH517SWvaw+lTevbityjtj/ALi+yfFkzcit2nMsO19+XFLR52jUaN6SA4hx86vL4j2rRpFGZp1BT/KV5b6JXxoI3TTVAqCCCVJglE2tWk/ldBhYdNUZiww9CtwMVvqOdoq40exI9Q7tGrxrZUMPQlzaaF5/g9lBvUXoM0unSR1PcVU7oIb2FojVZ+Z5kKJ56i87iPQtGSU9ANoVEgL6B5avS4bzrtsAVEtEGKFnGrE7P9hYmEYZVzeJDg+RadXoZZUbeJGRiC6GmbDcOFq3QlvAzeh0f/bLritM2O6o22tGLMnyAYFWhUWFOecfvd9Y2rxJ+Z9W3QFtYY0LCtcv2Ir0aREo9ROpuD2hzjvJJHNK/wCG9SxtvtWWt3s3JjII/NNZ1wJUlnyuIJ9or7EN1CX2MbPzWdzS0SIgsSsbxId0MIcDQBEARMTsa5nGgMkGDVzONBYKsCICwXngLXpibqGFY3iWdQlGdxLoNB+pW/QuTkIaajcChdU1AeqkF1OESwiAiA1NWKBFGBFQESzBq3zwoLZIAiAIgIgIqAinCJYDQYGlmBoMETE4iCw0ANBiILDQEQGBmx3UNTuYgunLzuJs9KySYDGc9VYzb14/M/nfpFjH2TDnt++GsMmsdi8/icP+IutbqItMNEmuC0QS1xE1rAPofzDA9Uq/CDbiH8FKqF7/ACH5nle1v4aX2HhyzLcf0NHX0mnqPzP9EXyI1Wo2rzVBAmGPzcG5l87W+vtGoGxpTxD2JlrxnTo3sC816LGiuvbTtAaD2O6mlngoIdk7UwrVd6Eh1SxoIKiwbevO1eg0CaTWNaoD5cvmeMtMMyJiGJWskBK1ml2SSQ+g4XpIobfRXPtreYQ4V5oSWpVBqjhQzCNaxqS9BscFCrW4Gr0p0efC7CC9VquSrHcS15T+XhpprE5gMIJLKWzaE0PIlYyFR5F/90BUmqGVUs6JlEkoSysCojO3/rXjhrZpqAMAaxbsNk9bZLqVgoKHr5XrvEe/aRgxJEQYXGsLW+ooOHWpvLSIjLn0wm69mTf8pdOUi4etynXJ4bIrhwFg0cdX8k+7fI+02l2sF7b5M6Hsd1Ba3QWIGEgCUbeJTlYkvWN4lYwyrfPClmMw9jurcsS9ovPAuZhsGsbxLJh4I0NbTm9exPctRmQF8iwIBStL9pWYc+IXTVI6sKLzwLmQaXQVO2O6qvzCNlTlqj0CwDW42ButuvaWQWrmca7gAqgCICJgaulqESwIgIgImA4owIgCICICJgRLCICIBdMMRLLRARBiIAiCy6DEQWiDA0sPQsiZXca1xDrN7HLlht1XL57icP8AiP0O0u/sbnl9IxjJ5HYZt4y8vX1HrTGLRUX1qIIIA3f16HHS8qV+onDD/wDENWhA/wDp0/KPm1z2f+cv2vxFq+ozY0EE3uZQvsf6j8y/RkDY0FpuCwf8HHQGgTTG2Rk2iZlHCB1JxK31FpaNd1jN9TdC816OAaMwI3Ikt6GotEYYgYWQVEGmEbLmCV2Pwfs6sTRlY0JnGtfEnxKummCYWKFVMZfOt8KdQjY1RBZhfPvWmdWubiQ42+95BTfAL0E8Q7JI8EBNCwr0EMRov+hoSsZCNCUZjW41hCXDpVphqQuW8FnWjyJiYTErjnlQ71QvIFGvVV1DzruWz0TTV6TzRdZaLZOQHrQjQm/8kwgGtufaSz1Sds9xQwpkCuP9tTfCURNNNC9VcLVRmx3UOy14xTlMofpqssWamxvP0IIrjki43exyjZql7ieDW25rdoHXjUVvqLq1ypoIb07IZgz5CwfG7aitKf8A6nS9ngfdvzLm/wCMYL1nhm6269pAECNQ6tgo3h6KExisZxqMwJUFiKcw6GarCx0AmaFgTADQmcaoB2kNU5XTDzQkMAoAwSoBpAFo2camyGawADQk0ABBYegxVA0pAXQBEwIgCJYaerAYU6hEBKzj+hMCICIAgez3EA1W+eFTgZLAiAiYESwiYA1QEU6cNLUBpawwmFhpZZdMBhBYaDEQYIhOwesyOi39Rc22fV8lcov/AAdFBQI3t46+P4j9Ht9stZnyaGxq9u4N4BpHfr0Lj7tJyqc3IHofEzsiV/eLit9SSym/4l2u+zU+pKP25xbWv+n1XGjQl+g8R+VWbVI0avN4j27RU9ELLcpfry93tqk02uCWDtyyWuWfR01DWg4UENjbeDuO1s6r5lNENrCBBotnR0xBiDV1RBMXb5PS4Zqib7xrY1zkStpz/EkLzqoetbVeIVUsXka0LYx6OhyLkM29YEEtdBHdQs5Ylc262ETIdflWVzcHR9tXUEXcNgNkY72ld0Ju7VRsFNDWVDWY5AwjTcJy8RXMKrnXZ4BcOxKpjUFN5Dg6weLK0brekhufHVUwfEtpgZxtbxKehHduhwI1vqr3rF4s7rs7Z7iYmCO2O6gKw0Jeer2V2m+eyodlKlN7RW9U7I7Y7qiCpVAVBmx3UQXJQPYb+pUERuyjYU0TQFzrBZhedxHp2mToUCgkryP2+/0hTen8D2O7xdqCU4C1kRXj8Ph/xHv3V19kXJmWGv8A9S+14b8uvGCpeKiALW3PtIdzIlbc+0gZk7XM40svIgRuDosTHD9YzjQWOlrDCYjDrG8SYnFrGcaAzCjXqsQ0tmG1uyuQyqjQqJUhAgjVGpWoS3KZisAGt88K3gMl0PawHjXMh4uhFYA6NvEotQlG3iRqAsj40ahrCpUCJYMKcBqgImBEsCJYRMBxThFQBEwCKcBpYETAiAiC0QWGlrA0GBoLMILDQYiDEQnYXWykcf5Ht8m/Fue32Q9epWPCaviOav1m0jMOM1n5pnwRnBlb/TH6o36F+QEVqhZxD29xKrt8PvlnU+nTvZbh/FU+d/8AkO8xa8N9XzubgMYvsuI/PbRpJpuE1efxHpWZA03o9CmSHsSrWm1tOlnrhLemZC0tqsLX0DCa1MQQtwjQdBWyNFJfGjdHBch2hYrPtNnynJcZ16vwRaN5+gIpMo5NeXefO9rjzm2l32Cm5Hv6uw+VgrG8LQq5ZUSLKkarDR8K5YO1OqwM0IILT17VCW7WwU0sNy7XBXdaXtFHNBu7l570lTAzQjTsuYo3Y3WNIErHGszTBG4a0JQ1qbRzv/4iZF+0JH5Xq/mdUCCCYFWpuqLRYBG0SYgbCFGvVVqU1W1qAqTDFMfMkK2iUjZMzY7qxO7IS4Wqo1kSWFIqAyCQlltl7oyhO6KjbgCfMJsa6Ck5ps+gsle6AZBAq1m6Fm34Z/N7vCkXq0Pz68EVhCJhYiAIgMqLzwIDGs4/oQDoZqAdrfPCgGK3zwoLe1wfEgsfIfPAgFUBahmqpQHW1qykErGcaWUJr0alqGNc3iW8FvV0IgIlmMa2iTC2ZsawNQWVrDPdM/tVeIcyLWm7tUpqVnH9CUbo19DhhARMMRMAiC0TAiWETAiAIowYDNTAYQEQnRMUCJYRBaIAaWA0KA0ARWGCKNOiFAaC3jdnvljiHWTR58bzDKJfK3H8N+wcq+1vnScYiZBTX4W0P14/D4j2ry1mHRfQ5LyOsM/TNfQrtPgrqEDj9jqbe7X2XIuH6UPzf2tuO6zLvF8SNmta4Xrvm2ubirVxQozV5vl6gIa856C6y7zwqrQ7RYBBJxcQtLgwzB0VRYFVnD8ITK1i1f6wRTs5zfDmisDwLzXrzOXyFe6nXnLv3R6CxplhbxGrcON9SqPnXXM4t5y+i2X0oKFBtewj/wCxUdcvH7aN3yZfa9FOvgS1Mp8Fbqe4g/BcbvqbK1dp/qrKrKYnZ0LUwau2Tb5ZvSU2wftz4BGKeOk2ipJx+FFez6gDvwQQ2C4K9a603axhYQW+MDR5Eejrh3tYavfp1SMlyBLL+NRo7Z3OPjXj74lwjEs1VxGsvlCIeiyX0oGY/wCcjUunvwXG3xGJ2bzxE+vq+r1bWm2jal0G9L6KlLExNcBzvKZEBOftzD/N+OvD9b0zqbGm83dbjOqZkqO6DHR+1xxys8rvWmtP7sptdmyXrpXNMNfzxGwbC3H0feUTg2mG9hYa3oZj1U5Chrctrglt55+BxzB0GodicLQKNroLpbhOsbxKX4jEWgr0BDctD2VwbtfMUsbJ0DW07YZI05WCNmyTRKwRrK4I5qX0KJu2tzsw3WRxWsTOHQlvLxqyHzoioAiYEQErG8SAJWN4kAOsbxIDGt88KAfDNQD6AXQDIex3VyXYHrOP6F3QoxWHcSYWiE6IMK1pvm1LMGCjaAfCNQAq4zjV+qbARhqkae1vnhQeCqAIpwbyLDlqWHnSRXap1KoXoImFiIAiAIlmIqC0QWiAiYESwIlgRUAwp05hLUImBEsImANCcNARCgNBgiWEQYIhOGgNHvsAGBsrV5d1Q/QfZ270cDvougmvstXw/pfxH6nVd/ZHX9RgDGgZHAjJtQO4P101v7M7xfMPPHdr761/h0Pxr2g+2XLod8MysCwNVTwpw1I8yjgjQljiPStGqtMYY1I1wsOwJea9OF2EavQTL8I3B0WJalYVzONX4lLmVdHja0FoTT1BfqqXKprjRuDBgxCnre5ay5Bfc1PkGvuWEaCOoDnGYIHVMThTXw4rasJejnSC5FGr5RD9w47WUyfpmN5exm9cFGQ9zqvol5ncQ5tGdREdV888lm2JudHp1T0pp4vmFrBYtq0pZD1lMvUa7OHGYPHIs3i8/wBOtVF3wo0wvJMnPVJxk3Cy9NEQX+/vyVn4lMTwJ8N0la9dGTIzr1OxxB5z9OooeXdXPmW3mXoZZC2AR3/YKZej0IO5hqE1ak697NZlXGpTSa+D1vUtuZU227f7wV5y93eryKCQIcd+43Qk0cHpelcczpuNnauecPcTVFjDmTkDNCrhzTVVQkum4haDXbQrXlQUyJYcGByJbdgGsbxKPVsmsB6F1LUSxG5bdqY2Y3ElhQnaZ3Uj8qcK52WfrW6E922q94EwyN9xWcRyz0b3kcF7YXqGtavizRBjIJRBwKwb+pKu3n3OfLmygeOisAiAiAiAXQBEANLMGrmca7mQZrm8S4D4S3LELAJdcYoLYBVuBBZ9MThpagE0LD0GoDJWKB9CULRY03B0WrSUrWHcSAYVgO1gPGowEaaEgB5DuLuroRbTk6xvElavR0RWphEtQGgGEARBaICILEQEQBEsCJgEU4MJgRUARRgNWBFOEQA0wIlhEsCIUIgsulqGBoQUbCoTlzicPrOs7zs1fA/Q92TRGtfr8MeY2Etxw8Phzc9y137j4io9xenX8ar/AKxi+s82zoE8TvI8kBASsGAO4cOCaMccfNOXWbuF2viUUMWlpeXbRZ30FRw1LmUaMyx+ElrWv1mWbH0pb0DIRqjabBBTa1B8NiCNS21gabRKxxqUbNYF+0LhrnkdNrdheXxHr2rR7iNGhbuWaDboRplYzQV6mHjrAMMMw1xRrZDcIHBujg4vJq6hPdw3CCghBB6BuherRQ8C7vG2GyNKxgVbh7CQQo/drXfNNNgwQTX9DtGjkLCgVsrVoddBAcvztQSsn4eGA/8A9hHpju1JGwggvU/aEh6cgQWNBBqBbEtigszhGtomHrEwpy3EKamNCT3kxK2rWmhK9Nktdm0TVLfu0h1jONTN5kQM1ZcPVzeJZ0TkjTcBuFUmMDtD7ibGxUbKUzY7q75cndgHsN/Uupo3bJK80BySEQ1uffrUxMkxMKk2+lGjTn/QHcLvpud6WrTTdNYm0POvGCChEwsNADQA0ANARAKGbHdS1AVZx/QgC1jONAPBxpdLNBxpawxlbhRr1V1MarfPCgG6xnGgJWM40ACt88K7iQ9rG8Bb/YjR3RgttvQwkFszYJ6iE5WhZxpahjkSAqza1m4EA0FWYUKRaK69tWJsQAhQZyPiWdQOtANLCILESwiAImFogIlpxEGCKgs4o1AaYBEANUARLAinCJgDTAiAiWnRLWCIMDQDEFyKNjtOZ2anWfU9Nrs+8UGqB1X80VmskrGMcWHtsVdzx+I+p9nuT2MeHzxGr4caj0ecGx+ofv8AFfeL5W8jr4nxPuLmqbW1/gPpedggjQq0LMU9t5NevEvzeXGY1pj/APUsnTupEHmQsATV56lcwUzLUh6MSv2msCCrcC2S1qNGmaaacsHTlr8bmgNv92XDmhzNMzQ29nVjczhrFYZiG7/84hzLawwzKJPeU3aWINmFyZaiMN/laWa3ZYrutPFmkajhsFNwBH27hxuda60/aLFt9QE0JwFQE2A+j1H7++Kz1mfdKoNOrdwNydeL1rJtF/I4UbNC0H5Oq7FyVvzrhVq9F52GjTsGa25ogmKHyolpJgZi8/R6IIUbNrcDQEldEtqgk0eq1N4aG7bU2aK2w3Qr05m6jWRKO/FDwM2iXnrD9bu1AHCNrQq0HMK9NnL0zTG/qXS5DN0Ri7GxBSj4vpXNCcBqhK1+NmG89D8LaHA7pMeeMEmJ1gEsGGg9nuLZaIAaAxovPAgMUANALoDGi88CxmTCtGdxLhgaAUrGcaA8QFl0OHMTE6wDjRoSAY59uNcw7k3z1B+97F3Epl+FGvUYAgAmm4Oi1MTh6ahSYCyJUOHK3zwqMA6y1myvQTPaJoTFG6FSdpJagE1CgGs4/oVifVUhLM7KDSyBEA1XM41OnBrG8S7iTBq5nGuFsUGCJhYiWYiCxEwsRBb2sZxrmHcPFg8RADQDIRqAzQYIgtEANMMESyw0wxFQYinLRLLa9NN66SZoMrI3AMf2uuenRxPmUTd3dp+Ha3GdTTJBnWM4gBK9Gj8j0Pe13+odHleV6OSQIKafb0EOsajfVoqHFJ2DojX4SnerOjXN2oUJu1eZ4egvILW1nRS/KmNxZ1MzAS9CEzW41GekuylOTs5qZM9GmGtXCONjRuRIczlu0GlitsMCzMqYhtQQwVa4rf8AkvO9RX2jZDY6GEFhVfWf2qj/ACqRIPQjUzrre321qWDvjPdNNP8A2heh11pu2tGwwW+pLBgWBq31oYtobXBb40EbuBM60UWsLv8ALeCbkQTLD5OrutN20Gwr4wWFS9cldqTMmuCGs2bBR9cu9rCkNgsFN0JQKFXG4J0KJbTqujNDW9E7YoJWhWDF6EvPXKlThhm5b0V571GzBZEq3BaxtGjyNSumeewtphjQ8j6K0QrztD7il/MQUDVs7Fq03TU1B5eB7HdQ4bTCxEFmEBEBEAOsbxICIBdARALpagugFjA0AogGKPi+lGg1L1nH9CAfC6pqAtciTE6a80ewgGueVcxDuZWgSJED5ZxKvRxAzV0LMI2hUid7r01MTgmmm7CApjTTdlLUK+tjXtNQGSsUGA1zyBFgGFOAa3zwoDysw6EgChGm7tBQDSCxEtQiAYTEYaAIgCICICICIBhBgiC0QYIgwNUJ0SwiAiE4aWsRLUIgLSV2ZbQq2di3IL6wTQo1gXmcyezZOemaZRLC8yHoSQvheStp3cUfhqN1NOpuW9BMTzuBRhxqC0TcQr0zkl8SCmBZGs1mWsKaTTKP1VniH2beA5zDD0xtgkPYzCojV8IM5tbgU71ZnKqMvnnN6OFVdKyalFGZ0Zw5XMTUSDmIwzQsf8nXfThP3cmQ4/NZZmEIIj+7Lvo1u+9rZtYbi+KEHW6xxGw+TLfo1Ive1sAFGJ4Cy3WKI/3ZZ7yo/ubY5z/TwFpsDiDj9mWvTrJ7y2h5Br7szxrLWhW7hx2qoneqZPh6oURtvRnZhSYkqK1uZqn5aDD6dRyHdwlM6UPXDRI96INqfwjNZufmDP37/aBycYucx4kU7sU8Kj6uyXq774UzhuDAwMRtCo5bW8y8tcw6S01hrVQ8k5K0EacdWrzHqNkM4ErXJ1CN6liPzCNx4IEuSWhmz3FsuVLGdnuKaCCqtCrTHln6FvEgsymFmEANADQEQEQA0ANLMDQYTQEQGNF54EAvReeBABQA8iQErOP6EA7WM413MjUJcBwPIt3IC0DjZuHYQ7k5WN4kxMHr6q/+zmVgGsEHFQnDrOP6FzDmRasJQKSRBvp9AQFOrFBmsZxrmoHWAcU4J0eW/wD3QDCAYrOP6EJ0QYIgCICILStostQ0MFluWoZGQEQEQYIgsRARBgiC0QYIgsNMMRBaITolqEQoRAGB0zuqn8oaxfvgujmrzuYvSsnE44G2t6KVW9a0ACNolAobRAktVSWnUOsDrU6Hny1qtog8sanPO1cqvuRmMmwZ/GoLAoi/sFRWrtnyUV6Ize8lE1+FM0EmJw/cf0CiONHEfRxYRaMpX9EekG+bMnO1JVvWszDggZM4fDzqiur/ADLr+pjeTLf/AL4B0sRltDQ2b1x2V2p6OBokuObTVOr6zlb0PCsv2uJWjU2EP4SRDqywzSdRa/G8279oIi1y7HO3oZF6EO95FzZWC6bDj2w9Q3PcwvSrtZeNae0dMt51O+oZvXSVK4EcPlMdr9+NbEdHbLtejRaTLzrv2lppPX4NTTBAo1CI5KoIzhxEIi6ckZLmf9y5xOA1Z84iXXvyWyTrJrJrANYcHTOh5/ew0ALU6SsHO5AO0U1tTpfovQ98wl+y95e7Mkh+EyAjuP8AuI4nD9MWV53sOH6pGVLyl5e8THZmJh44IUJgzws9/wCDXmRRQ9qbu7iHyzNFxeKljUruL58sRyH6yHS87euI/U2rt69u1yvoyba9z05fn1G9SLfa1Xs/cze9MIcSq4HyiPkWjp29d9h92peJW9y0sqq50fVOpv8AQnLxcHCBYFArZjjSIvEM49eLy7+nr3e7xa6IfZ0sXhpYlmCuIKG2wsNH6Kx0S8SLmIjDeYLJSvTxDZaMIILAtEq+iP8AbUGIUMWwVhptatp1mrAo40b7amhOoijf+StsGanqYSVCTHlnw9juoDNAEQA0BEBEANARARADQEQC6WoYBUWDooBvIkAnkSAQo2caAHRs40AKjC9r6UAxRt4kBMiQGFY0JAWAU0B4MNcmJ056AvNqAtAo0En6lrQKNBYE0sauD4lGAFQETApK2iYhQOgGEsCKcwwgswgAF6CgCh6E4NQWaQWiAiAQjQbTci2hChZYG6FhzCCTKGREARARAEQEQEQBEBEwxFQYiAIo04aAiYnRLUImGCTQHr3BU7iGWj54moPLV5Fb6i0VygULOV43ga4BWMKYbVRNNTESnjUrdFbIwZliChFwWjNBS1sQ43ff1L96SdDH5sagY/8Adk6uiXo23HpqfKt8z0Le80yZnE53o5sZJcccbeO20G8RU+ppg+rl0eJw128PJOry1LOqYZfnNAh89hPgNb38JhxFC7pvge4Xp0cemYebc8o4tFeX1De59Ee1RUm6q1l+u+7qbJicSayVNYnEHlQmuIDZaWlv1fMXa3ReR1bo7r2Wq7b4o0fQsa9GdvPNgp0E/wDD7flt2w7BYc57vvNvVXeUvmv+lLn6S6te79Fo1Frb3oBU0X9ofAX7B3VQBMMLIGIcvN75i7cdX8RehRc0vOvPZ+4lzzVLejPam0qCwgG8nHozOpw8xivn7iDwMmzs7jq/Tv7i49OoOPeUy9DlHspx6P8AK8E9F9lOMh4ZJvL3zX/B3HOdZ/Xu3nMKbvKYe3/0nx6vE/8AhocY1fWq4/Kg/ni93eVh+sT+DUes8xxOzJqbS0tPSc2lZrw9afYm2mdIa3ff1SOqFv8AElHXup0gEuwgGL2TqIc7lS9JdO851bxYrva65UWnsla2kYcnjF7ydZylqO3sL9V9mdJ1kyKuLJgFS7Gsvt9jWfXPnlEzmKtYaXez1AmpzkGDtFOEIi7lzmIQRE3lE69xiep5vvV588WZXU8uopjD6YvWS3hDcQUIEdwA40dxtblFF+5c0aOvStBAgsiBAWeG+eu21UIfEmoXkEZuJqwoFO2e4uxsGClTnglYCppoVFgYgNHjRqGJU9Y3iXoJhNxLvlOcSkZkPY7qC2aAiAGgIgIgIgIgPMj40BlRs40AJLMDyJaAdG3iXNDADepYuBXIAtYzjQHiAHkSAIgBmUSAWMCQCJmzoKYnNBdU1AXQVFgVjnhY69h8Bah1ZhxpAANNU5BLXpqoPyrgtevZtuI+TKMLANUAwlgwlgUPY7qAOgPTNB7qAgeRhUSHcm0FCILRARAEQBEAwgIgCICICIAiAiAiDBEwCJZaIUBoLRBYiAGmJxQdvCVHEPs3Er70EbBTewP9hea9mWpQUxIewW0ONODQlAobnBjF0qF+blga2QQCyJqoWNMv21jAqKC5/hCXWutnE47D60xNroexa3WQoNBqMzTkiZ0erhutaEwJwF/iB1cVnRtAckwg0JxRNx6q+NRi0Nm3uHla46I5Fho5Cs6ame5tF/Bb24ftDYge2ykVHTLUXMNih8CyJNooeVd3WjM2VOL/AAyj6Gu6a1GoCGFwh+l9C2btSxsNpsFI0Zx94pB4atAoLhMyJSUEXUausyPAu0EqjhvGurpv8ECoQljd4ZisC01VpzgWAINThhWB+6b/AGqvQDKQCVnH9CsCijZ2FWFtBjUaMrHAYX7Qhx4h04Gb7aXw0F2sEIjiYWIgIgBoCIBdAEQEQEQGNb54UBkgPaQJAK0bA2oBRLUF0AmgBG9UxAK1jONASsZxoC0DMC2EAU2jQFXg/o9AYoLLmGGbCDHoZpuymJ16GYYlqFuFgNXoJy7+i5rogoD2i88CjAyYDCYEU4MIBgPZ7iAIgCIBhCcRARAEQEQBEARAEQEQEQEQGVb54UBkgCICIMEQEQBEGBpgRBaILESw0C+9Bq0KiS3ozGHEdCNy1O3ensZ05ea9BYwUxRmw2YMxXJlob0ctwY9Aa9NDWxuCv/ky816U6w4zM4dHsJauYwTDgTDDK0Lkk/oep3S6ggbQzegqpPdIleH6PGwVdNBMXXluoRrTbA05vW+1px1aQygpwVZRG5hX98m6V/AjYKFsrVGxF2uzZolZaxJDQpriFboQC8Gu/exa0tPNh5xpih6FHdmwJUCrFNRQhu7pv8EglEnPKiGwJiIjhz5pqYlZ69VmRBKNoqEbhNWFC7rmcaMSCEaMV6aZaZGo0aZssVjCnM2e4ueT/KLpBmV8sY/SaHn3bY0555hAMIAaAGgMDDUAJARADQA0swSs4/oTC3lbde0lqAKzj+hADrDfb+hALoCIDyiuvbQChmz3EAjReeBAC1lagMss40AnWGBtQC2vLeL+1APBRr1UAbTP1IBajZxpg1YhmmBpYXQcz0aYnJv5nw3aoPmW3qchEwCIBxLBdAEQEQBAja1AMIA5pjQtNQAQupah2N1gHsd1CdmhwRARAEQBEBEAsabrIFWnIA4RlaE4OQGSDBEFiIAiAiDBEGBpgRBaIAiWWor4YdbBdAx7hVwocDnQPDlqQ9WWv3RvRwLzOI9m0MBxmj2FE0vwY1gsFXQRdy2AMzMYFtxgbsEfrXmvSnZz6dJZDMDfhf3dWNzDV4MGZBjKMxZ3g6F4HD6xL9Nd3S+DMMgqXpKtscFjpmwrPeAimVlWVnL+0rg8NjmWuEJzrbg03Zam9Dz+7F1v4/pUvpwV3cm+YC4C1a6IT94s4LK3DlMTELUPAGauS15MmcCW086gGBJrzlEbkWRKBUIEbh6LF5z0FtWcf0K7Ep1HGjVWna+mFquNGLKg4qnniyXpj/8AUh57Z0xOaCQBkANARADQC6AGgBpZhdBhdAeVt17SAwQEQErG8SAy6HDmIAyAXMDQA0AmgE0BgZsd1bjYmd2FCzjXWihmRqhoDLVG0Yy1UBjrzxpOJBnXoJZBWrCRiQ6GmJxEARMAiAiWGFbc+0gMTTMjwIBeCzQHGbeiBzGkIB+sMM6IR1ggJrLw3Hv9zoBo2NBhGOAkA8EgsZBYiAiAIgCICICIAiAiAiDBEFmEANLWImJxEwImGIpyxEGBm6E/CVXkOIXz4KaE2i2Vk+dnNoyG0Nea9iYVtb54V5r0VpBDOit8NPdtpFM6Lj9Ss4gtDpuwP+teZD0voTo61dZJRuVgjWrOVDX+dmNbASwfiVtRxkMJegmxMGAjIyHpoKv3S7KqMzNGa2iCqFz0xN4NBi51N0IBHQO7ltAYc6GGLujurb4JJUaZprVtDDboNBaLTUeHTvQD0JQbq/8AQNb54VvBGQQ1KnU8z5aY4wnZhMTtfrK01JzovXIYWFZUjGsaEu6p2umGK0tXmGKUoGjb7bf7FjQ/RYep3f8A6r0eG8e82MSXpx3yhVPNbCgGEARLCJgDQA0ANADQC6AGlmBGBIMKIBMxAFrbn2kBmgF0A1W+eFANIAmRIBYwJAVhsFTE4WsxfH/YgC0VGgFDNjuoBSiC82Lh+YDNWHSHR4c1Ae0jeHN/sQBEB0hMTmFQBEBFOA0ANARAVfPPBa3WXX0dwdugdAQIKZzJnfmmnZCgNkQDCAYQWYQWiAIgCICIAiAiAIgIgBoAiA9rGcaA8QBEBEGCJgMJZYaDEQW1a+lBa0NVwolw2dILRNXnPQaYaFh6DUl6sw9guz3FJQVd7tgDM4auGLg01eW9daBV2HoqhalE2twpYMoUPaJlamJ1gGHuJeg88QODB1lEacOvQTYWnOuGFwdDq0CCC2UBtEFMUUwYEaZliMaOZ1ebtoVFo6UNNokzdORrfPCnPPa5GjEBXSxlmWJC+G1LbqsMMQyQNViVUrz3qLAALod1UUPOuzzA8+rOG8i7AkrIjD/lCoRtpQBEsCKcwNWFogIgF0AugBoAaWYGgF6Pi+lBhcwNAJoCICIAtYzjQAqzj+hAWoeWaaxMTi0XngQFeasRufGwdb54VzRgZMLVZmz3EAkYHh6ASDtlUYGZsKhwnloajbPBmBroM5HxLgdDTE6IDLTP1IDJAEQA0AugNWgt70P8oR80Ggjv7fR0Bu9F54EBAja02iwoB5AMILEQBEFogCIAiAlY3iQEQBEBEBEBWGxoILTTqdDbSg79Zpkzv5XCkKI2+5yCMU7ffUQbla3sILfegsaPNnbW6wP0fKcyuBu66nDo8trcKAYTDDCAiAGmBkaFWhUSYW5BfCljdqWvlzuNQVee9FrgYajBjdq34TrMQvCvK4nDe1aXbcIKWGk+m9HuzgYanexEMAw9xGracwrWBsu88K9BOsoL8h3qoXoPPWlFwIFAEraJRtGwjcKBuyraL/6KJpZVoJq0nKm7D8JjE6HntY6yZFhz6ZuW1qMmZY/CWGpnVZwY2jDcNS1O0DBmGtWdjWNY3iXNDCBhiGUDDQFqAH0F3hu3bZjZJN5yXE8bRr09ckdh9wvb9N8Z3igg2RG4DQbC3SnoLsPZ7iYnMB7PcQBEARAJoCIAaACagFUBEANLMDQYwMCQClF54EAOi88CAVozuJAIGVgW4UB7pn6kBjlqAcrTOGtQDAaAbohPbQA9Zmcf9iYnKGwX2zlgwpRtCW2yZmx3UwRsxogzMiU7oWswYfqoCvf1HN5PgwKxmcujKckwgGEBEARAZfIkBqU6zqHBrAONA29vpFONmUBfQUMMMPIswgJlpprjCdYOEBfBdS1DsbjocRARBYiCwgq3CgGUARARARAEQEQA0AGibg0DH8IQGbAq3TUA2Hs9xAEQBEGImARMMEQEU6cNWBSztBcIeBZiVDi06wWiXnPQatRt4kvMmKGsbrzR/wD0SfBa0gxnHbrPEUWi9DjNH0DUnD0IW+vNboWYXn+mr7stVmGZbWqr0ze7WUByxQ8v2bq1bKFBK2wbo9uvbR4XgQeW5Erka9NBy7Lk/KQqcFRf+4XIAGsjMOgYjhCyoKGQOi9Vn9ik9NruwQsITEx5zA000O3y9WOKCNRoKjy1TJ2tBdOjcCgVL4OjCDcdFKejGMB1vnhWcKNGNZx/QuaOZQMOsRuFguKjbr/6r0+XfePFvvwj6EvdSRr3eRIlU3b4LbeEzn4i+293Pzbv3ATgzQzSIJgx7hfI8R9fZrCCxqtyKhzC6FmgGEJ0QYGgwugCUfF9KAGgFaLzwIANG3iQC5mz3EswOsbxLujeQKxnGuOvEARARMTkTUtQRQCeWoD2taGgJWt4axAN68eeBDeZehxpDBisZxoAWWoBGjrF0MdZgg9hcBcxALoDdQwwtlMTm0AwgIgIgBmGUQb83eEBQywGE01/GgziH7h+gNoCQGdGzjQDaAIgIgCIARpqCxQ9CQDCC0QEQBEBEARADQCUENNNt60Cn4OgLVAEQEQEQBEGGEwIgxEFhqgF5oCrYJ8nQW4/PIfQfpHMXrWTQzcjUhzR5mMMDMrcCy7JKWZzazQ07CCJXfPRWMwbevNektoNGtxYPlDhcMXIRTGmOAsLbBd6Du8bXBgjA045uYZmW9gWiMtgrYLw7Hr0HnJXB8SMSDOvQeCiwZjdC6BtegzQ3GQ6OuBQGGYdNx66FEZGt2racrGpnyLAhzLQTI1rypHVkIYHBQ+zrzeI9Kz0eGRmrSt18zk3W+eFZPE01Dq0DDWTYhYULONZOwt5DlYydZ1Hgm0OLIuMkdjtPSO++X1Ps3w9HwPPbx9VyuEFQxDwTkfz8Gv0K9fmdLieqlkYKCTTz0hA5OQN5RfI81scvsuSVaOdyoGYFUVv7OvCsXu1NgQ4YQnRARAZULONABMQC6FBOsbxIABvUsQCJvVMQAEtQmRIAaAHWN4kALXlnH/ajEgpWHcSAYQA0AsYH7aAVo2caA8S1AWh/rWglY32m/2o0BgKNYFosxpn60FgZagMKIv20BXIUOgJjz0QEQBEAKtuvaQHPr+98KjlfnXgoJB8Wi+JHHH2lDktjvRwWNQaSwILGs+40hDrbQkA0gCICICID0MwIzwCHSEaMzAW/ocWIOz3EMQdQUDW+eFBmS1bWG7OIceWQYfQnEQEQEQEQDCAIgIgwRBYiAiDDCCy6WsZ4Ge4V3pyV3kNvMvdBQSVmNNP6bxDHDj7y7XocR5tlDg07BURz+tXz+X0cuTTTs9xZUNImajwoZ0c2MujIMZWBrvqM9o9g18FoRlazP7wnvIy3QOZg6NRK2xwaZg2WFacqJg+Jb5K80YN3W7hYUtiCmcMLI9ofrSdY89FZYKsGg5no0BOeesbW5hAyFz7Bh9CuRgNajU6ZZkakCrMnQMPoICiMnasNyJi71udoag10wMOt9RY9RR2jDmzTTF53qPS7Q7Bg0leslsLgNT+FPhfaEEutbNYmWc2wawogbd+/IyeH7YW89Y7uF6fD4bzLy8fUd4e8mbeivdgQOajrePRjplMJHZM24HuPgOnHpF+l8q4fpvxTnl53kYdNCCogh/Cvv4F6LzFVfglfn4vcnhbePo6g5hy+KoyrsOYPnOi88C+CfdjIMMILEQEQA0ANBjGi88CDA6LzwLmXcFTNnuLrhNALpZhdMBI3qmJZhWjZxoARiA9o2caA9ovPAgAVlFprUAzkSAHkSG9HlFde2sZg0I0JbLA1lCQziCWs3EsHPMs4lrUHgjaNG7GzIOZw1lsrr0EjEhtqe897RF+2gPEBlWhebUBRzpOnOwG/oga87c8P35UHzOFTeig0aMt52nYGwOf6OPvSnIb4gGQ9juoBtARAG0P9SAXNjQWhb+gMsiCNQCuRRo35AgHkAwhOTN6dZFhxG6EKDQQYQeQhZhwgs0gsM01gQVcgIEbWoBhAEQDCAiAiDBEFogCICIMRLUOgXj5FZGjefk0DEQ/Rx9+er2+XWL4+/qcP1f2qsOvGQhkXgV0w5oBPTFo+cLJtbO4Hc/AV/MXr2U4KX4TQjTa3f7J747u0XyF8+uhxqM7LP1rh8tGjQayY1MwJZUNamaT2GZYFiH6d6jzptFQHOkZlrI4yHiN/WutP2uFyHfOYZoRqagzLdQ780H3Gal4VdS4/LRB/VOt0N5Nh39Aw2reEnUF/wCIELh30rmHcvP/ABAh8Cas4g/qB/LO322KLphrJH8p5hmhMVfqG9qcgwcZjPRMzCg9RVFo2yDQYODZa1EypE0wxTPQXAYeFAXYYa4dELcKCe0s5aWAYVGuN7K+co0GKFgXeX/xHK5i0bNqKLyhc/zm4v3zUFgHqHoknNa3a9vI77qLlfe8n4Gj8n9p+cdL6uMZr2bXcIxI/JZu4X2EREPh1ga3p3kWY9XvFuMpWYWW9HhBKlacFvxSrzrzTW0OIfr5jmPLn3VjfteXhPcEQWiC0QEQC6DBEGA0XngQAjQkAiaEFw5AAQCeRLGpgJgeD1VssKiL9tABorr20tQDlnEgFteW+03+xAZ1vaLUAVADyJAEQGVF54EAlSBIAWRIMRBgaArNZWqPJOqaytRkatzXoJUQFPM8ZnQL9F4FD+zkRHe0ACC5a3Xuaj4c/cdr7SuGKMOCmXz5016NxEJhFlre/H3Ws+DHT1tOlHxfSgGQ9juoA6AwraJunIDPIjQ36AqzKKsA7AgGjesz+iOt3+0IDOChMDCrdvf6QgHK259pDuZLmm+wmDH8aHBwgqNAOITiICICIAiAYQEQBEBEBEBEARAETAzgQBkajQ8CBz79V8tsU19U7pOxn5OpH51IG3MDeT9f36+rsrJ8tVU/L70RqY9ar1E536zW5PAnGt0sD79FjbS45vwDj0/wbteXzN9FyecU5bvqX5+bfY1Ed7a+DW2746TRg37/AGyoCxD7m/m18df7vt+FOaIkpGTFBClqsZDVRsqAyDKXLswpTQqNOJJFwcMxuWY9UFqIy81B9x4hGXOmCJl5yZ/VNt131HOzlPya3wOAN/vCPUHaSL+Se+FwJq16kEdnJ8O9POjNNAaj1IHZyuw700zcOHUHqQr7OWxQa8yHu079nXPVd7JswcmSxBtCBSnoYhZ5ZgyJc8umg4N0F0LUILAsKFzBIKYz1VyQvw4KuNYWtGzjWyVbHTKMRKr3XWuimkO9uZfovhc6/sSB14I315wdfUcm5Th8V7Tc6fZ0kAhytBH5wWIcAD0cPH3lffdk/Heo1BAqIJ/HGbnHoxx9qtc39hWSkOYMirgt0YnxFxQIzTmA8HGW06rmmWAp3gmshrce/JychedzHhvUsL1wcuCGhGvwjWaOxfHcT7x93ZfggVE6IgwRADQADNjuoMCQETAImB7kXAWqdOqjQvaQar6Pi+lDZoKiQDdGDxoBSi88CADR8X0pahXGwXjXQraNnGuAJARAEQEQA0ATIkAKjb7bf7FjRQqza3D0UeAWrKNJcJ688aThrDcDJnDCMo2nY9+vTecaMDrctQC1Flu6CO11Qe1GdIzM5kacSTJUpDngv8VGH45OMdO/iKQS3eWIMHBYKBBYLmHC2QtEAyHsd1AZoBhAL0YXtfSh3RUmmRoMzWQKA29vuja0OJGoLGjLCijthwhACNMjQZgEEglO/BcdcN8QFzkPgEBVzpM5ssSw/NDBtztzj78gBXsDDDILr0ZAiAH7/SGkIDaAjUBKxuIQDaCxEFogCIAiAiAIgIgCICIAaAIgOs6maVssPnk7MD4kdfRct3fNX7K+gZz0MfsNPsAB7V9GSN5d+fUr6OqMRhDZPy+9F8NMjGpsgrGg0+GDEx1o+9Vr3Ed/RWSgr+R6Nr+La96CNfA559SiTegMOx8JiDwyH8nd9Wvz/mL9Ftp+DDu8zhUZmwvPhdKjWj1YYGkHqSNBrbCpDgqzkLkODdHoLjuFzrM3i/sXNGumBKLzwLuZdwFRs413Uamgw8CAsw4MYgC87JntsRj/AAFoFKwShaiDgUEwpzR4OWPUWMhZBwVNdwtww/aQ2tQw0wt4YZuILHrChrMYuTbs1xK8Fx8ei+JHH2t17t53q9bh0eo8S8u+zfSV669bBL2EruIIE3ojjZQRvzy76t4v0jl3LvTfjl/f95DeY4yiCHgeyQOPWEd+vQjEvLkTQ7CB4Nz2xHK3fULpYu8A7QPifxLtdxhlmds9xcc8DAhNra7BmB0lI0m+le6MjkC19BZj3HllDzHl8eo+vsL/ADZOQL4jiPqrMRKVmEwsNWAuo04FGzjVWVDKiuvbSQzovPAhjAKHESy1SaFgQoD0P9SAcrfPCgPTDVx0BdbDQWQM2O6sTudO6qXGhgkAZAVphntoAKDC+WrmhjKt88KziQUMo8KPACMDw7KmJU5EawXfRBUGMqm+0YXvf9C995aBhmMM7AgNXvnRqdK0GC3vDrB+/IygjeXawfOVnJclyxLEafxoLHxZ/wBcCFx1tIXVNTE5jLuAtQGeV8BYgGstQGdG3hzf7UO6PUOE92oBj+2wwoBCC4ceaEdb26AJGjNZQ64IH9nQFXBYMabYTTGzsfwfeUBdm0ZumqwK4I32EguUdsbyowvQggd3Z9DGiwQUiAiAiAIgCIAiAGgCIAiAGmB5cZawdV8uTXz6TlSCc5F7kCB4MfTWxHK3a+gsnzFbjOqwNwXvX97iBnZfFyBRCOUKJ5i4d96vZRvj70TmTefKGzZdBBYgdw8cj8k4d2Fx9ROlQ+D/AEJafjIJEzWs2h/j3HY1+Ze0m79T5DOLZ+iM6URuWhLxI0elLV6NvEsnkzA1g+YVjQ2mIAWs/nhXRhbwYNZNW+s3EgHg5W9Vq5kwy2SDDdwoy7iVgDe8MZuBcmRiV+Fe6WiOk/zkUSHcFTJYDDWFGICo28SxiGDQcFW28LEKCrhB+i88CdmAe1lyL/3CyFVMsagsswV+bvCOX/xBX9jXmpOvWGRmN/lTngDHP8ssCNqBuOod991d0v0XlfLoflvOr7R36CBMjcacN38i2I5JfUzOIy/PRevcbrt/J8ksn+BQstNIjm8EPX3g7jqF1gQLIgn5zOD2PhLt4gMtNaP2wuGG+EHYN9W04wW5+1yPw1Asc7nW8mDHDX8cgeT7bTpXE5U9Gy57lzw+R41BdNAXg33Lnv03+FTZs9r6FD7tV9+zXmPVEXoFl1OEQEQWiWWGgAUbONDepE0JBpgKCoA5q5DssK269pdcZ1vnhXMO5VxoRvAV1wkbBDcOBLMVVIagwqZWYEAhgN970BMBvvegJ0OHMWNVDJbAeRLmoV5mz3F5BahINwXaoWOnL1HjkY1GtZg+jn34/S/lEBVyVLEaDDPjUbOt4sfufeUBscFraLLc+gLNAEQDCAiAIgAmoAeRBZFX59AVcajVGa4lYLb9IQD7AggwtPsHDjdCAq4LBq2aH8aNOt3G597QF9kYemoBXLDDN4cIB4IIILsCAsUJxEBEBEBEBEARARARAMIAaDA0tQ3u8DK3PTO7jeB8cQvY5c+Yv93aZ2jesgT83B6mT8ovqbJ8/U+foyFr1fggFbj3DiIvYkR4B3aXH1163ktx+/vJbZ0vfzNd0OPZDsJHfvU2dT35BaiyZOc3VQTdJxmJoj3gjXHxHi/Nvabd+mezE44D9Nb30ZrINRNXyj6SGZgeBYUEaNvEjQMKNnGh3RZayt9pn9qzk8QOCIyGyQSCYF3cmIbJBYKsKVqGEGtkLSj4vpWtU4wOx3UTuATEeQrjAlApmCtGzjRqMHAgla4vwoKbhQDesoYau1Tquao3RA/+3WT5nDSr3UqTPfovgv44ZAvzTAiPSd+ST14JuOr5jsLr610vquTcqw/P/ajnb6plUUKDyrkJ1Rrho/ZnVwvtvT9N+bd93xoLIoI/N4RkY/761pslZhNYEFEI435GP3655awLoQX7TY+CcdWjLLw1uYB2/wDERJg27eQ/cXG2JpvSQg1uff2TlDrII3Ih27/Esn7x2gs6CbW2DW+9v4jxdiNZLatGoLBJ2Cf7fT6PvirZ2cyKupWDyLBESH+ZymzdLHE5eqsb8pdSqY1lcEATYL4bmPLn2lhfkV4b1A0tYGmFhqgCJZaJYMILANwBICvrDTOihQz+W4EFogPQwkBmgsCtufaQpzJU3t1DivNotxpah6GGmJypgQSApTA0tQGlqCZmz3EBVF6cpKd2qd3NL6V8+S738adw6NlWb57cNum3LOJKq4mr0eFbRFOrusaNMgoVaEDbr1HiOdwYyZ79U6OI0bASAITACMn7K8TE27qgVYYbyCHVrlfAWIDJANBIAyAGgB1jeJCcoE0wIx/gOQoeGUWIN3hAV4dYYa/NNz+51zZ3GSRkZ555n1lgoOIA64Lri+YaEHkWY4OOgChZZpuPQDyAYQBA0Jw6wL2/oQ3kAOaIKbwhBw6ChaNnGgtnRBBaE1ARARBiIAiAGgIs8M28fQepmgmskkERvR38QJX0nLnyN+WvjRqtNoeD/er6Z5DQI0HQxt/0ah+BLhWuPKP7N3Z+Iq2GpTMEHGpIj7dof0ofLaS8/DXOI7ZPxGnGQTb1Pol80wY3ENjkIGirjk7v+dfFe0nyy/SfZvSp90Xl5myLLWL4SdJfZ0y6rkRq0STMDWDy1GzjRiXV9BAsKHYjK6BgmBcaiFwFBaLZXWcLoOtCYuNrEILCmMxC0QSCdsd1ZjcAHbPcXICuojTdBAWtgtQpX4b/AHdYbwugggg+hgXWdHi9NMXNjVEF0VO9Bzl2EXfqmfnXDqHEIcUuvEQG+H1A7jtl7633Nz6Ze5ym3mt8TzvnPZvoGV5VCDNoggMRig4ePtTkb1ju4+L1a/Q+yh+VdTZI30I3kTMRD8TDx1tMYjbaLW+B4Ot41sRyi5qBTW0II4OHc1YRyi7kPWBZaODvFkH4TOXaAzC6JpEcwbpevvBesXG2e4fI/wAa4aEboQGHliP3F0nxoEabRBDg8HGtvC/7F1kU3tLJ6cZ058mjAVEbNC56chP1vf1Hg1pMqr4oc089OGBx4fHj5QPi3T3vLt56xXWekZRVB3u4gaCbRRub6jhEOJtLLx+oS55fhV7wyu5qvWwSN5cEBre/fr5i95dD6Km/c8jN7iNwXcGIXl8TkD3rLmOVKaEvMe4Sorr21nEKHqSWiWYJR8X0qvUEzFL4JkoF1LVw+NxTV2NyIBXDjmRITlTVQ9AqmJ0o+L6VG7oVovPAuatZgvW+eFdYByJLMIGhIMLJahVmUWFS+WvLTL4U6SzJcsv5omhlg4ceos1n2sxEavnuW73c5X+LkifzHlqKQ+6X3XY0rOFeJq1fXEZM15M515XOt3/sgP2Neq8NdQVsGC/NYLaExOfCMCCyL/qUALWU002s19IsHGjuB82gLwPY7qAbQEQA0Au3AblpoKAXQCMz1msr+jz6AqjNejYK4ggWIfv8SQRvKWoWEsQUOWA9ZAse/wB0P9+TE6zCgvh+2EBYID2tufaQRqhuvRuhU7j5ShwWiNoctPQE1lbwFqAMgwRBiILEQEQEQEQYiAEYaEHpqHRoLA5onY0cGVc+RudX8PlzzLy/fVgMEZI8jjwNu5xl9dZPjamjmGhBW81HAYiDjW3LE+sdqsOTwWtNlaPzUcfl8YiLpzUdkznj+59Iq40hjwUyLnWf4d3xp6+I5JwDYfbefAQH5eeiw3tGSX6IHe4vlhA2AMWkVyAzlHLOb/EXx3tJs+79m9MOmXvTKMNwaGvz+X6BE4ddlgysQYumBVq0nQIL1TUOrqChodhfBIa8tpglD6qzIWFC3iRmAMEEqlB/5d3FJ/pOVNNNMW3AWhcPiH6ln/Tq1CCotCBXHRaLzwIzIwyV6ZXRuN0fYF3Ac6i7uMXwo3zrwQ+w26IRAnRoaNcdW8u/3bj1y9HlPKeut5XO+ddnZusXlZXCg8r89QIRI4A7dbZVHI0l9zefij/tl7cek7E69Ivu+Hy7035Je38XsOjysEFBKiOYMRBx8n5TN3C9CXlrCCBZa4rfY/LCF0BBM17jeW7flhCIiMAetZGza39sI7xGIgIEdpB2303lX6NJA1FuLhFk58HcYxcBjoUVD4bx+oXQWNCrY3RN7EGs7KGBrNe5podoqaPwS54DAKtNjfyiJW3grRdalrgU0hTxNJBpoFQ/3PEBsU9cu7h55TmVuIiIxDzNit9I6VWsoo2ASO/cMycgcV2V/JzHw1dZzOIwjqU8kGwQI1wFrCQ4cP8ARyCYo8dW32Lh2uKXTwjWhBdBhNhuep33e7tQrEN0PoftHn53C75DW41e8gkc2ALB/wAIXl8x5Vl71hz2Jc8mi9dG5Xy3PuOEDY1eDxOG92zvGvLwnqmEARCggYEqwBReeBR5dwyVLhijCN7ApwWoWBepUITgJb0HtYDxphGB0OK40JAVZodG3AagPUsxXG9UxBgFHxfSuatYVhsFa3/+yolOrgd85pl+idOdiCHdKgOu7imxaXtCydWyvXAEAdO4PLUNsh3Nzg5lSlTiHTpZlnnLgr/Bl8Wf6QQvdl5q6kuWKPp2a23Of6R2FbTtjoq3LTUA2gGEBEAMw2iQFDBgjDDNe43+zjoC0rfPCgAoDld9udDTY1CZJkmbbA44jKH+9LEHy6CGGbBWuAgsf74Eb8tkLTIw8i39AQI03QqH9oQFjRcOQDSAIhOGgwwgtEANBgiAiDEQEQHvMGG6EB/d1f7ued3y7CvWX0o3kQUoE9rkEYp0rfdyTvm2yrqTJ1ui2HzREBgB90NZnE73dl5037qt7qWJIlbIpUB+UEYcarvTR96uY2bhC8+8VqVye/dGmBStzqQTPjjZQPtdqrrJJU0aNhMDkgCiBsH+vdsQOPjXdpcDe7eKjw0qzIL+a0AZgz8OPM8GUT6T6jt2uWR1T499GtvZVctSpfAocdKkfhjPBvhrC7/5L5/2kfU+zMfBlpt6VlbB3Fcvx/y/W6dnRYKEYFlrMwjRpt0GjVWr05/TkbNbjhYQmLafaF9Bthv6kOw2iC7KHWzULeJYyFgEFgXoJxTaIHcC4FMds9xR+AIDs9xXTsFhoShdGrGcaBlWmxtoSYmaeUdM86RpxBYK2n4QRtTl2vf5Tb9bwud857NtIcsQWSzecY09luPZVA/DHl289I8u7v3Fyv0Tl/LvTflN/f8AeupxvWT9FAsn1nHo/CXefWo1TLA1lEEBKuDHv+mRHKesdoxqDrOsjjhEYx37NcKZpiF1kf7/ABDE+C9eqGRAmsoiDgd0EUY/YfdrWgC3E4ot0E1n4dwgGTcNaOCFtA1j4S76tGoH02N/J8T4K4QAoJptcbgyfLFndVJeSK2trjtoGemEcquSC8kVtYP2vDnr7yaJmICllYIKCSRXGg4/Q4dyt2teErW40aaaE46I5FOTY6K7xy9hEtZWy0J/HI2BUbTD6hTGr8IKNhV5wR9Q4p8oHJzfwHffKGd1CxCojQhzgql+4IxPZQ+xrSZ7RZbwfg0RG/fdoBjoNbuYd/5N8uaTKlQRu93BJp2Aad+/0hed7uep7wc7jl7oyC7GY7WXgcx5c9+wv1LQ8TF5Xpy+l71gtAmlhKxvEgseuD4lrJgClCvQA6NvEgGsP9IIxIGQCZgaAUo2caWoK0QntoBRAcm1Ql8IwL/ytlc6wfRbZiG9KGPqqqnwQvfSVGb30saymG1523nrJ0aNamq+nAZSiOtUTjo9rcswtyVRtTVh9CfIjsfuhe481ahpicXBGtxA4hAPBdU1AOoBCNRoOChVph1g4Q6TgpZsaa/3jc5CHDaACagKqaKzWWIUefp8nXK93bRwO8ReWjMzxl/fDvo1Fs4Iyfe3yXRK26jzLvptboUFBTUK0CD4ZTkP0A8gGEJ0QYIgIgxEAOs4/oQWIgxEAKqZ7TF33an75OYL9pXe7k3fN+vd3gppmnLo2zW8DdBG2vviL0OHy55d7fuvQSV4LJH6KQCw7YpV6jwmxhBBQQKuOZUP+ELaUtG5pCxAW7yNHHT1IsbjcDvdys/Ojh2jj4/sq1umfmb6JR6Mg4vCTwdJN5+IDx2KQgl09Z0OlsOebw+5j09t8Be7a20UFzL571InoxN+i/FfmAk2+7K8HIBi8ZsiIu4IeZx889JccwqKbfM4cmX6aTuExkEArcgPqT3xA/hLP6qh3dVd1WtmmWJV/wD9bhYcQ7/GLl5ORS5vq0b2X5ddT/fEHbj37AHkSh/6rh7iPu15vMXu2Oz5BvG9ZnP61+PcR+0WTrkF2O6oTPB5Di4gxpioahYKgkyGZReohnZtEEjeBdJbUEaFstVoXYRuDosXACab7Z6AWrm8SkzAQI1XJz+4fP2kAM0miCwmrXuxN3xIuR5ojIWvlDQAe+BAv2HG2L6HlvKnzN9ztRTdO8rXlYKeDIwMRi52eiG1W3L+4+JcL9N5XyP0ofknPuY97D4Mv338L4E7TM4meaI5EISDCet4EGJeOnhfNvOb6vwa+w41lOXjcv5jBq876KRfNvNzPW33QyJugT/SB3DbIl07tPWXfUXa8/nXKfq9XlvMYmH6YXidUhe/1T8IbfCvexwc5zFmumOBh9p7H2O7Xx93Zx07vapqdHjZreerWMLMOMjHXlmBxvIjqIHMOBrEcfsqsTQMaFo8DC3ONR+Fu+reLTD0III03ZyfM+CcLjQgWna9ndlM/gXWWYORBEGt83l2gFQsPOrEDcGkYkdZVMQush/bBDoP8Rc0yyHBWsCgp8c4QMKH4+Mu0TjMNatammN9O9Y4JmIfiR+zPPXr0Y2eS1GaDenY4WT9KB/niX6ochuUFComuIGDucdrmn7Ior0+ndtZraIJxA9oHx0R5T+VTYjc4MwKjsDWZ+IaQP2t6x4uxOZTl4IabREA8lk/k7tdcG4ObX/KO86hDTM0zD08gef3QOhjIEbmqChBEHdrdZt+e/AWc41VNWla92ZHILXRttAeflg4+8u3+5+9Xncx5fEvVsL9rU0yObKzMtB/aF8t7ufQd+oTNnuJb0p2VVC3iXnqRaNvEuaqdEXXA1YGNF54FJmQyTy0U5jHQkAiZs9xL8G+FOZWYVxtrF8++D+T6Sn80UNu/cNzCk4jVpOIc+vRSwbGre+lG8vr9HH3pLXx9Rb586ByyG/bXWD/AHOOt8RPa6OMQ+R2R902MTuDlj5uFv6kqZW4+r6kvesZGTD5pyhw4f6OOQvSeS24OsTE6wQDIXUtQ7G6G9SxAndVmQU2NdA3aNHQ4sN4/UgPDMjCrUBQ1pkzm9JcQ47ZQFDfpbGfyfRbWU6wf0+wlxuonYK8tBY0He+B16Ot2U6JEbN1TE50IJANoTogwRBgaC0QHm4+6u+GPBUPY7q44bQG73uryE0zuytN6XgcI374lwvU5by14l/fuqQO9fJEE6xwEch/wiIr6my4cRD56q9boEJA4JpwA2UKX04a70OajY5uGm/USgOZTRfGmq93kMcgJB5HCIaVZeDxi9DLzML06OX074priVpWBGg7imdPojEScbjN7uLhYnTVS2mSL3QUDMIjt0eQfFj2ZQST8DFpUzhTvs+ZPRO9VTBrzV5qbIpRW7+UoO6LAumEWTtkS2hxd75cXV36xexyi06aco+JVl/OnfsnKZI1G9eY0dbxWKnvIrF3++kPvTr66bb/AI4jq+MKTp8MvbwmBmhty0GMO4q45Ry85tyvOuZw3S/o5eznLOqDvSSFfslk4dgU3S7VuH+8lXdndvHfnva8POZbY69G8+wE1HAWGu8uPTKfsjgZ2GuxGAKbBf8Ay7f9nJdB+DuMYsqXyZON6L8l1884MPQT+mI/f7Wvyz2kp921P2Pkl11Wp8XAGZ0QF8nxHu2hzIlvU9ZA6Z3Un8rf5l3Rs416bz8QEhg+HhDYhts0EjeBaZXwRuFcKeGq2BJsLIv/AN1EFiE7C0I1hFuRucdepw+WvFvL4cwIwEIA2OZBUaOPtnf/AA19fy7lT4++54aFj8ZgZldBLyWj+zMaibuzc/EuOr8dfW8Plb4y+59ENCnqbZ2MMfmzVPsSsP6O2l75+4Xv8v5e+V5jzGIaYZBTZog1E3EOH5GTjj422efDXv7wg8OQ3zdTCIFBH99K++0gdw4xUvy/DxsuiTze3KxN1GdWoh8jTnez52Jz14vuSm4ZFn+YlgfGO4c72l2+93d+6VsxidWd1zqcL+t/S8tfaNna9dHqdy40+oxcOfDdm9x9pfP3NrOXtRVEv1X1Feq3vY6pyDsjUIjkOY/hHRPh+6XRPxOr+Ld+uXzF3aRVD16Zw7TBGVxtabmB8sU7rwI3La7ta28Ld7WstDswBBP/AP8ATR+Vu+rQoYmtogvlBDr5u4XSXpptEEOFwgh6+XSyx2RStCAeEEPXyyqZt/RaENwaQRbLpIeA0KVoQEzdBBT38NEYyW5+EaFBLeajfY/c/Zbt56Re28dVhBG1wDTWkOLcisIIzrvk1QodEkrIjT5rN3Po4/bN2vnb16FKxCCy2tNzA+OiPZlpgZnTs4g44/sxHYXW9ojQlVhTT+dLgI0CwqB/N2gGjQjdOgZ+QeVQ4UrQmsrgqeHn9s6K++Jd7Xdod2asab+USaaI2AxF/wA75Dp8RtXV+7+KgOkhYYJkWkAEboUrqwNotBOx6yqaHM95+CG5dA207/g+1rzuJy16llzDLnsagZkrG0Jy8j3c+m78GhZxqLVXkI0JQKQqNvEmAuqAmRJOgeUV17ay3oXNCS3SmRIMVkZNgoeWm4hwgOANmeNX6r6FbBQelQGKyjNvV52z0d5bdfBmeDXo5Y15MByFxucdYUzimHL9Zg77k6Ob6BptvCafEOOyIY3nLdB73cEjdzWRx0Q7fN2blDrrvR5Beg88zBQqPs67JZ6syLIs+tlmKzIv/sgBB5ZyCAZQAzDaIOtNzA6ApoKaZNGWV2I3OOgHUBol90xvSmV8GnkLEKG7hBUQTgPaHC4FgmJziAIhOIhQGgIgBoAHMhG5FgajsSOpsEk3kr6U7BYZVgNO44QTm17vD4byry8dokfU6yrJBuvk1xDXg8fHU+1ue8XpWXL4jd4dV+vY3fFCZYBZSRt2Twv3Cfu8tXhX1JWrQOn4w79+PugV5U95ce4VqXC9laaTZ2NfxtkPp4QPiYcQRnDHm+fES/8ACpYHGhadwddDR42EbG7A7aKm2II7GmJl5D5pNjVuCEBTjkD5OQscRTYtljZrJJlStOyinG8IoOXRmFV+/HD0am+lGb800SpqfIOxluOfrtH6fbub9IFzfg3nNr9A5ZabRD5uqX5yX4b1UYLnF+wQHtVx3i92qmclS5PNwxXPfrOGbmcl8RQ3Vrqrpqfpd6CLq+5YEK/8A1+qOWAB8ReF3sTyM25Js8qAu7vsu1/D9KvB49tGcL4qfpUa002VZYrc/wBNIb5R28XnKhayNhBQiVaC3Af2pg5G/bXdu/j831K2mc91TEjseRnmQgbfWgd06II7H69fL+03LfeMYfoHstfYhzyiorpxt6/MP6j9Rx9kK6Cu7uMw9lv61BO7UNlCy1egnETCUQFsFRoDYIKYErksLXXsIJuxUKPXDSygcPMnRj87DrfCB+uEQpsW57G43y7+yvr+VcqfIc753g/dHQRkE51ZVgMRHf7o1us3rwx38d4vtuX8vfmnMOYYa3rLBWGVrNcX7/dHT6veb3Z5jmLjvF9Nj6vnQDZXnWjrQpSIfuOERGzdfeWP2FZvOEaogsGjU0m1sEBiI7ioox3GdeGPN7uLvqHFx8Nb7xL0t4CCleVstmkGHRCLZkgf2MD8n6e7Xl6tNHmc2aDY0fNNCQ4fuIdYw+pGxlT8D0i9SzZqfKV/XU93UrQYGNTsBEGHEEPS6dukl9/6xexVcwX0vmK/HB4xznP50NAHAcHEUkAhEP0d1zG2fH+HdJF5E4yspaRqYr7U/wCp0nN9foldtg/BcWVBnHZfN7Rd/A+z1S8a8tcavX5fVmH7Zai3ViXpNVPek55ZbjpFe0h0JGIeQNZvIc8uOrcfD+Dd+uXz13V1LojDuOyCP0NIiVtyzpeY28NyLW/te1iXj9QqGmMbC0cLg41iRyt2gx5NPRNombgHsUFvI3kZtF73jOnPhUOpM+RRscL3vGsVqyYq2ZRuirXAR3sfDrHwvr/vFyGXPp3CCgjHEjQTP1D19Ed6tbva+Y+CvRsnlVFsGskbPNCxFOO6DHIHzT7vFUw3UKhCp4HX4gfHEEb8SvFh7C7NyIIcHhGOIH7F6x2gC6ECOFuh/ZPiOw+4doBQ3AEFs4+MaP2F1/MgNdCrZWgh81wRtRlHS4cnfPX94tJmuzsbBI2F+Y+n09treTm3zz4H8Cr1YbFJErRsKVhwsFQ/caORVZ55/ApG1nGzNZIJrGFj+ERHa3PY7j4ay6cCjevcEIB3gZ0+HWnAtezdZNA0eyfDkIG7A2IBGhDsNA3S9c7PiLPpqe+UMZkgI6gNgjLAcga26O+Zu7drweJy59BZcwU0blUwIJwbQbntvB+vSvd0PS79RntXy99w3r03qv1lYqTw6KiU6c0Hsd1UKCsayLoKMNbNN7RQocgv0xo2ZjPyXBYMv64Hj7UsQfOuizky9/Bb30suJYgsdr7DtZYlTjDUb7hod8G3veBAZcwdTQROpC8TJZYsHfyxOzLegIycdZk+I0dMJChHN5QVZN9pZanC6CNNMjT+tzC9N5q5D4agG0BX9GNGadkDjR+yoC0+Q91ACMNbuPPoBWCtjRof50U9vwfa0AX5EgAoDk06RlkzX9oTK4XsRsLHg+d3X93eftrZBpAMIBhARCdEArzBhjKIIEgh/wBrqz3d6jnfdm3OVbwc7TRpuQOPKKjlvKoeJf8APIh0m9ZeflaCm0Wt9RT6Q/iOd8n6TxF7/u54XvB1plCEFgB/wyHGqnGsZBH7eEEq5K1k2iCCiEcCz78ejH/fVqN7etgjDZpIOw07iHjWI/h/+2sXqql0SNmhhBUbf7FI05/O8bYEEPA9/wAcR+4rpRqA6NGwSVhwq7rwRbfs1wuTgOiXnAzTYJz1RsCgt9HH3l0vP5jMS9SwarqmL4eskE5kE/W/bYyRwN3cYz6yt5dZpb+p+WEtSZGr7l/Y6+9O0CsGReIFRcioGzTu4tOYd8x8SzX6DTTq+PziHD5mvTUddM5sD0Ad498Jd9QvQtafLNT43jEhtg8yRWLmZkFxVP8AlLtKmMS9RpEtimCTJrxtwONcU+LeWnx/Wen9Pc8yvK7XVb1P1V9D29GZjExy1Cb0OrGp2uAXFrB5/prNr3F8x0z9t7db967bV5nHt4mcy3E/R+jEbC16vPwiNhHVGd1vI7Hd4xQqlAZdMNr9fG5+Cvafszz17u7+GvE5jw3rWF7o5I4htGyibtC/KeZ8siX7ZZXpONwXIug3HrxFiqCCoshQ5H0bTBg20aqYX4UF9ROLDNgvqICrS1BfXAx2xYnmEcrO6ZrbReogkDvixikjk+2DjdA8NtKnk7v3Hx19T7L8v96w+D9qb/sIy3qOXxjYGY/gkEbDn4AGJHgw9mV47j4K/UuX8v0fkHMOYK4y+iEYx+2NgQSH/wDyt468S4eekXtWdm8iqox+UQI23C577BxwcaKO3X3fMLXlLgtXXra1+bG49Xv+Dw0V4e8+fecwo+8WdJUyaAtCr4dL7in0fA7KOfO+zuB3f1OoQ4uoLKMajQQ8cCaS4Ap+l8QiOKeco4cbX8Fc9RzsV+296FK1PBJVaQ/fkUuUEZx9zeL+r1a73jvS4pf5kvn1jh4YINu4h/S2H+Axdw875/6dUWkZcql8y6tHU4tuI04lYIK3cQGDUjSO2bvGP19JbUxVGiCXypf9vTmXv5bhUstBx+DXA8f7FwtVUzEYVROU1OOqmnTUUTQDNMrNt67GxeEEbrHuPcXfrH293a8K5tulfE5ftzeJ1SN77VIyZCZzks0dgUWcOmQdg/qfzr527s4mMraanTDMtjb8LfyLEfkrhTQ6TCN17mnl4l9j/trKgsF06mnB74RG28EtaRCcboRqaf6wiX2FnY1mF06nfl4iuTMUxkrcOCM16jVbg0+IvX3lP9NaiIiA51JGA2NxiajT+t8RKfVHhHju4XovKkGVtNACoN021RDuwO85d3Hx1Uw6BJITTja4zMZ6IED72vFvXsU4hbB5abr4bmB8s/gdoDMILXs3LcxnojySlaLG1sbN0DHkE2I/JKqWWvX4JoCDY4BBqCKfEw6n25769bsktWznGA2aJ1AkfaIB0yjNOzPRN/i3A/euFV5ZdbON1kgg4WUjv6bKKbevgWakvctUqoIJlFr5HMnAH4NnDHu9uP4/96HUlVvTuIYAKd+Ravhx9qsqbNoBoIKtgkQCY3czp9kyAMDBK2CPwdHf65Vg5G/IMkxrIFRZaBo8SthyBtpdev8AcLBqzCCrgcHXADtnOuVM1sp43eEgkcy6Bn2Ha68vmPDe7y+9c8jUjxuVugavL93Pd79SRo0z1VErVevqgUpr1W9FAaffCmfnYgr80LHv9oHUauZw0W8tLGRv52jQNgceRj6jalTJsLi+DOnOXLGA07H7nXA0SDQYzEe+z/HV/Y975tdcWd8GMwaTBPyhsBx7jSB6nFqc7bUGWDT77cFczPBLohw5azBTb3xJGxmtTpEsQY2CwZxBTTbd+4Hyghek85cIDXZnnQ0OZwJJDBt3D/Pkby7QG0BhBBh0QW0ICmMmc0yNayyqDb/0htbpANSxK+spj8004h+/f/NoC1QHnR4c1AKmmBwUN+abmHGOQHINT4YHfBvhTLfEot0YghYk+NXalsgwgt600MJBj2p41b7uhP3x+CwGNRtuRQ8juLXu5N37foLqdTWZdNR2Ipnr6n8GvT90vC9+w6pJEsSPBJIrYIDufKCKXGr2psKYqw8rvwozoTV7jyYDkmtjU0RALDgAh7HTkhw3bSfXu15UyYtI5G601+D4EemWFKojZrdBB0cfEj8qrkaqjeDEQPB1v0j5TdojAdEkmV2QSCembj241+oKruK91kUqGNxusjdD7WOI5NWRGIwy0U00ydpoyLYPIsfBqzSmn/STyzCgnPvfGoYG3EQch05qO1rj+JeWY7FG42DK0DfnG4HDhxpDV5709nxdqwJ2bGzX8GrsfH4dWVG8wxxi9r90U8d3C+x5fiIxD53mDnckyXRyxM0botAlR662Pd2btfQ95EPG6XBr4MlhMkx+GZ74bnzlncO16drGiKp8X6pa9NdSYJrMaD19yt/vnwFyq5wq6cw4NOEnReW5bcYQuur+15J3cLkxiHqNlk+BMhF7Z+YZiH+tDy378hIjh/C738YfTuoA9EUvg6mST38FuW67So2MO9cJfIJ63u7vq3g3w/g9SouPaTENU38S/Ve8zf0vE6qq9jCYveijlc1+21iI8RyUlzzfu7hfOXc9Wr2KYwrr4d6IxplbBM+QNbDj7WY7uPcL5GvgPurPnGWgxF2YCyijQNg/4ORtK/M+Z8sih+g2N93YOszfbb/avOektYIFgVjkRhtwQORLeSDlCHxLSdURuiUEZwplpEzEhQXzxj7sdwreHYMXlbt94aS2XrQX8amnT4uNlA+1h8w7tKe4+L674a/WOVcq9N+I8855F3DeAjQggq43T4ho+TY1y79e8u/h3S+g8vlxTYJBMI8EBAGyjHRDJcW5droa+yVpI01kAHp6axHybGmE2mc+ItYymNmytKzQtAqH8Qxw49Lobq49wujJ4KVoJz1Q+VYIAOO/0yMkDZz0+1qBYGb07NfmnZjPeCuFYwEww0I0+a8+/cDvaflH+LuPFXJ1CpvdSRBQjdfDWYiXx6wjlLjqHfjrtU4wnc3jV5Tn2nXLs+RaPYhyecfr6C0vKY3efVS+QdVxemZfAvhxCN0OIfkWzOwjXHUL3LenKKdnxPPskRiZJxfmFgtcOLfMb0PcLE8KZl6nfugaiXVzX2dStfxfnQTLpcixDvXCENGtXjrmN1jds3Nx43iLwLq28Ts9iJy/anU3aoOTL+175xfCkybGnQpkOe2EQbm3pN3i/H91cetXz15Z51hbTU3bQq85u54J5W7xamaBgnSNr873vG/00FDSQyjN17N9jxnvznrFyYzGGkgmAK3N974c9feEu3dncLs6wQLBMEEqDfe8fJyN5e5v62MXZxOgc6LC51oI4gm3kEVkR+c9I7/fV1mkqKhBGhTtQsbo8FtiCBuzvM3zC646OEFrJBHED98MdEew+4dqLy2bNwMyJu6McQP9hddGNwhA6x1+PIx0R7C69Y7UapXVusgOvh2fIxMO/fViTy59Go1BZXDPng3MQcfJ/lPrFXLCq1NIZusvPTG8/F4i9iRG+fA+oiHZdLNCCrdfJr0AfRx9tM5j9xSNQQ6dzvG8tz/B9qDdXG1oiIgGTTaL9FaZ/Ybo4Z8RAWIUbggWXBNp+19tcoCl/KLK0bC6J48PcZ4cgYrM/OIC4CrQtCPtwMVlH8izhSsAjejWhZ/dA/A/iIBqtrW4dH7Y2tG4ZmiMNyI1tQpfdzvvBrs0XnoJNQVdAz7DaV5fMeXPesL9xK+JI80SSbloGI4QvnuJy59PZ3zWzY1Rh1ptPYKB6TQoLGWXwpnfxs3EAgaPT5tLUNpMjIQYb83aExO5U0w2+DM7+aDTrAHNDjkLGx+7dpYkoyDB9O2jn8Hp9pXJ1dQyWINOmRGA4hwgL7nUlwf0sMbivUQDiYnYGGbiadj3+joA+RhB1sbp3D9xpBCAEaHzzh5EdkO6KdAWkFDCgodEEgGUANAKmmhBoDSNUJfBleS73p7DTrB+ePYjoCm1KksmSze+cVsCoK/GjkVOeWJ3PiMQ60toBggzTDKEICof7np1fw+GlvLx0qV9TPG40a45+Mn7XHzn8i9ujlzy72/w6aDePvWSqa4yAfecpxrxNs7KYh5tVQAV8QIKNgQMKAeyViOvSmyw8rqXk8G0R3yjErjrWb3U0hc5J9dmNcntP2bzu1bjXKMKNxuNAhOAoIBbxc/I4Nynr3nxHSJmMahsErQQKSIIRBAmdb90b89u9sUSxAqIPLeDjZPyisZ3VjTQq2uw4iH47wqCWF7uC89EarjWaORbEIrq6acluiRs1gUEIXkrHOa3pJEDdvIyMf8AfVyVrlwbzrBnzV73j0cO7MS/VXEZsXRbx8kc60rOGnafEMdEV50ziHqtfvwTqFNJr+9xXUFOPWRCI7zzCZZWSaqp8lwZv5QzJmvpBAU7iLxoCDwYfOu3MNcu3jz/AFrpfYRmXzjZDILQ3rI/hZ1wIAD/ABFP4S+XNJnkrXoOEwTBn4i9c5R4NfQWiGp8u6sK8abP18NpsFBxxEQsofydxi1dXTmWfD57vy3qOeSZnMFg2ZhNnD3HeLtvGZU1aw1aaZAo9fZJgoFe/cEOxKfsbnq/rr27XgPlOJzPSYjVzUC6maWjKEMCwt9IWrq1wo4ddUTs+jpLjU/3vhZbvuXsI6RAYq4cO3TGjfA937tfDXVpMRmH2dNT9EtRBqntUHf4nY2Vr7sDgx+tDHTlxMEPFpbEaztLPmPhXa8am0mZWTOH0bfCvIflDZWQRo5EWcaRvnJr5TmvLvUfZcjvuzhxUaCMBNfwON4g8AixIH3lfnfM+H0P1Oxu+8XAQSnPWwRuDosSDATTQsoagNLmiaKXdw9hwja1nHqHfhHQrw16C5LM/KLfFBIcU/W9wRnHTvfOWe+SuV97yrlT8n55zzDqFWFpsbzDjrcPtT73Hir7vsnwXVL00I003oae/HevuR+OuuGwoIFWkG7gcEun39Z4tLDJgVF+fJ2UP4fiYMPtb55d7X8S5U0tiysEEEwiazso1vxw/wAp3tE66BXSuabrLF5pOz5GJH+U3apUAGtooJoGkY7wSe8tgZWwUICBm59xlhHf9R9RYzE7GjmmtDkkALBYP4wysiDe1rjMeMiPirz9P3/+ilLGYhBJWlY/Lx3B0X6Ww/7x+qrThTNUOVX7i99C9g0yScNDYP5g0fobhuOod98vZ7vLx+l8r329Sq0KS2G0NudHtjtQa4/nXs2txliqnL5nvganWM3s5NZGboHLYto/J3HV3a7NGIUwZ1M1/W+1qQDnE6ytHCKI5/ZNlggnIoi79f8AEu+yrwrm1ml6ETEv1j1Mer7vR6o6SaEQ4hxMpDXT2IQiIjWZLp36zmN/uPh3C+Zu7WLKNHr0z1O5hUXOs/oT8RECdIU7qwgnQlYj+kCLHxFlVsx6AUEIY3dESdOfB3CCBPYP+sIk9feCuP51OPLR5WC17jThh3LEEdi9f9RevOyBV3oWBTvNMfvjHaA/iOT9mGcYu4d+OuXjtLowRpppuW7oxxHYVK0eCosojhuYHx38DtcypVbK2NxvLs+/xxHYfd+KsbRgtQTvG+emN5EzEaHDh1uIwU5DqhzWmxoC87A8xUUnLE3efeXfxV6kTmMsOzSQFnzmgaPo/JXCQGLWxueI3XYcfmexuXVoo21gbkQWsYX/AMxiO2mPN7uPgI8jLDWSiC7Pucjag/j/AA0BgaEFBKhh26McQPtveIAxtCEF07AGfkEaPDtqsuzoCuCgkbwkGtqLAjc5ObM/y1QoefnTw8fRtzDYu03vm0ONggkbrekZufH8t2NSw0sq3cOjvx1K0Plvq5PwgdVZhkU0EONBUJwGI4P/AAKLslfU4Vf41ILJ2guG9wdhcP8ASIeR92vM4nKnu2XPIlxODXrfyKN512gkOLDhK+b93ZfXxe4c8nSMzPM8zP73lDYOOEDrz9nob6NrkuWMLXBtDiHG0JnhPENg554LGWP4KFn3GfWXVoEGEEG4CoUArGoTbmNukARMTkgo0EZGsiy9x90gLTWUyNde6ew3OOgLQMMMIOiCydAErA8HZ1rTIZURftjrQCVARRhyTVPBtmeNS1K5rMQfEMoWIPnV0uBwYOChuIIFuDFLdiTU6dJF5ONTTYGxuoAcEaOPS419/l3C+i5dw3yN/eu53ur1kEvdaEANb7o3zx16LyW3m7D/APUueHfLn98U01kbreDkOnw69mmMUxCPOpWVpWbz8RiOcHIsR+Vu0ZiAavw4cMACC08iJUbP8xSNFzZWorAEECncDj2I++vviJiWWQIZghj+OzSCxwc0boOBs0G69YP8f3SfExMaKJWZvAfO1SE6qmk0ILIve/71PdUUbjZoQWse6M9Ef3Ha3LDp97uCayQRx0eykKK93V0EL4saosi3gdZh1pM0YQtbwTfY8bKOVu+rXs2SGp7Ksrc9F8SHwM7HgSuPrlEfl131Hi3C8fmK3l+zdL6c7848rPzYJgryNHHQ6+ctUTGo2EF+TkOo13PHdmTURtjnm8W4HuPt8x7hfR8jePzJUwWCwSCyTLAQW6CIpEmj5rbHY/2Ha3ERmZSStZpDor3UIC98I0U+IH+Ji1pxrBoWsvOxWgdmyn4ZKveWofyYBNCj804R7eDw57T73aXbyzuFR3eGel8wQXU+miXzqyNaDCB3kRiBHY7j+dehFzqzhwqWL105mxmLBRm9rEXEWcWuP2svGZzm/tc0voeDczl8rzLh5bKJqPrmcZYi0bmYAZwcBBnpY+t3KKXi3szGqzlvKaafAP5P2we9k/yHQYw7e9j9P6RfNeH1ecvrn0OiNSxBp2j0rNxHPOx1ER/m+Yu3fju1PTKqX2CFG4Kzr5k+0/6d2vKeqob6N7xk0scTTBMeeONYkdmsNr/hXyPH4HW+z5NzntGiC0W4sK/Pb9+l0AG5EoFTUI1HDdBaleovm0Wl4S94XfMnR/HG3WCEwAi1Hw5p6S56t5yLi48r8Rfbcn4D4D2n5xh3+iDjZrgIFlgAPujeRrP05F38O6X6LOYh+PsqIJvTygxFO6cw4fkFloc3pJkIOfIJtiCPj7Wtsjxs2hsIHtEP8s99euYjOQkbwBHAQRmYg+OI+U3fV+LcKHTOYWK6aMAUDAlXR/ZKI9h5vMO1uyS1MY3kQUIge36YR8pu/wCRP8umGhBRqdwIJtA9lyVncYy7TKpxDEKU389poyLPnxGxp+xf7EZiIzIL31o3r1Gj9ZMxUChj9hGuOr/EV3LYmKXlcxfOUbkkydTCJ3mkEi3p3pb/ACnMu7h5m/j82v0K16acQ/NK9XXZSgsajUEhEEjefxQY/Jr4vi4iMv0CwZTte7DnaaaIIHBb5HDuTuFFMK3KtUlqbg74M7vwoFAbdw4sodDx/J/bVFrdw5NL5g1VOoiNJvhuZWkkHC4Aa6hDh/vxPr/rr27a46tkMx9XJJ6vTTmFfZ5mSqhw/hJNIA/HJs3jqx2zm03i0z1aKrB9M6kn0Ww9kYAvX6oc8ho47+kg8z02TPexlcx1HLL5e65XE6S9mK/o/RytghsEABBP3PlFNm7S7XhrTRtaHTgs3OPlHf4xASNGYDaIL2PHsW8pnPxFVZRiZZqaHGTecm9ZF44bk7+IdLYf2H3an8keTMkwTnXkmESrt9PWRBUKG56FBP6w+7uFK6bN0Kh4RjiEBQmm6ySsR0cfGMSOR2tvnfLu8pmuBmsgkEPnjaIfiYd8p/lV84xhG5let6dztF5pObj3A9kOR2V/1aqYdulaC9JKIIHHkYkcdq8WZxD2DBtEF0jgeY9kSN+eb38RaTFzTaLIWH4/hA2chn+5AetyMPLdP3OP+Jd/DQ6KFRBBa+G//Lqjbnn8FysqCIQWveWxw/EEaQRtr74i0mXHXrLjul8IHxI/K/A3xZUhG1s0m0TMQ4+57I/ROQEbrJRZFlH9I/wI0B0LIoL08Pp3+5yCM4+dfDWQyrde6c6B58fSOTWgtgja00c2u3NbLKpll2nZTlGwp06njkrwW+KEOFNQI8QIf6QRmnqXe8thTTzB8pX0dQZG5XnV/GpWOiMQcZ7sjlfI8TlT7Gz51Eq2NVsrwUhlB2Ecfsi8ziPcs9CN72CmwaC1saBsDj10NtCCQBaI3zYgNHMjUamcx/K0rZj2QII2pLUNpgsFCg0Fogv+8mJz4WW7hQDdIEgG0AuqAImAvu1Rhol6+U41f5v1RcIEEghxAdHf7W5Vvu7rS992kvsK9dqd4JK4Vabj35G6N5+IvY4fDfI3t66KEEFBNCzC9t5bw03R/kywpCjhuRD9sDIDT53CrbA33wHyjlLhehDymzBGhUTiOcIGUSxVxptbNIBvvfVPl2Sxwo3pBu0D+We+sdrZJAI7SDTfN76x2tO4EabRBERw5mPH0f5UsuQoq1oRxBpu58dyz31jtVMpJEErY1lvLEf5aPAdaD2O6vOerDnUbNCNmlxWZjdHJ3C9DbZ5TWjTMu17jnywjkvOzXdg3SRwgr3Mka+Rto9vEMsiLNsfPbvqHferzXrNGmiaDYLX35JqqLAcfpNBid2PPWekXo4iI0eS4iEExgQ80TTTnx6PkFRGIP8AbPSYu4X0llERGHj1tzjQVEZB4IxuPcSkBUcpd2jy7U1lq7UsL4YVFBJYgdB7CvfnLt4obLKyd1TGwunYEEr8wOBpObVyJXBQT81pg/rEXRs2+xjxaTKyM3vYKZeui7Tc/H406hpHJ3GPu1R3muHOlzIKRDJJvezdG2hY9+O6g4/hnlp9w7Vnq4nCTsCUFlkwOSp1jcaawe3groT594qO73UdLnDL335lzNh2iHC/eLOHXOZMjMavfNcRsLEP4TGcn8N6dLjRXh+gt4bVC3vb/FcGHc2EywlllGIPtnyjmNsuF4cTCzGG/wBozEHa4WHbA1n9e49x7n16T6aab9zu+NA2SyzXyhxB+6NrtPP3a+E5ryrL9h5JzzLnseMMCZ0V8XzN9/ZOX3y74JkHlo86DadmoeP2zd+kuPrrVvw/U4jHN7ztLN9iXspXBvfXoIRJBzPY506iHJfzP/Tr9a5Xy6OG/Aee3/eRiFvluIgbNPiHXHlbjc/eprzVk3TdPyCH47+PxrtWMBQQ0IKojh2Pp8dUdlu827UKwCVKLXutN0CDjPYkQQqOYSn5cVlgI6Z40AGZiNcCKyID7079ek3n/HuU/HowZgmmaX8cN3QQ9iRHya4xdx59lXIiIjCkAs3LtfMGPIx3hd771WWaOrdmCbRStEI5t8QyMf8AHUjY8rs1lCPmreB6OH8pd/w3CqnXEMFITBQjOvmPcQgesiBHxMx9f7paonGE1+4VGpWmeVzj2HQEiMOHBGPH1sJe7Z2Nfbd3jle74L/p2PUdkkWHGAha9m4jW+HaPvL272tfHcTiPsrGwWN7sHWU0+aWs6zw61/abtYvcK6SslBhNjXPVtEHhz0zwnrPrrlU5jH7/eA08K9uEYbr2aDb6zjPTP2n/e8VvvCEfu5QmanWVgr3cXjmffnj62wYj4/Vqmb9z3e+K752oWPvfMPnUMDIIDD7VreyesV9rceWaoLah7VqzlqcYz+TC+4FGX8mOW1ZD9mMJgXu/SbeN8D1vrVBe2UxpL1rC/y/U293fEle+9YRuC5g8h1Yba7fO/Wcxd+vuOYXhfJRlbhZhGhXJj+NnaPlRhHJ2ammOwd+dzedmc9M7QC9VBMwOO6fEduO84/+8Vu0o26QQLXuNv8AzsRrhRLGwBGgxs2u3P8AhXC0FWzp3G3EDNz5GOI7C69e88T0nhV0tSRmNa9RquDx7h/iYN8muMXcPO+VFlZ6apqqmp3+42wINxI8E9j9I7MTd9Wm2cJ6nsiyuyVoKDBDWVD/AEyIEdkfrPE0U2LqkbMNlWCOJVCZnxsoIUSxRVwQQVd/dx957cu/irSZAsiy03ruRo5H4l3+6gGwgQqIiORzQB/8Y83tDmSmWmm89UcbiMyOP/J7hDqyCDCjTX8bjmIA3R2btdwlzMxspL9O42bltPD3Gtz2n3sN2mbJmNaEb0jCAsHG598Meb5/DcLGyk202CwQ2hNyg/8Awwfx/hrvkMTQsBv51sqKfHU5O3dkWW2WAII0c6B5gjRyPw0Hn61rTRzgd0DWw/K/70ErOtCrdjEED5PyW+KdkYHIsiOzG5yECTmRGg5aq2XGNUxqagpo/OqB7RpFPtq+S5jw31Vhe5cgCDw9BoPYcoXg+n6b3u87w4Hsd1AKRtmWNxVn0OpS1ChkuCmQWV3EFNOHfv8AtffEKFwEEHsZ9+gLxADQWIgsuqDETDHsFkeaJ2a/lWVcQeR7Ibz2RTcPlrz7y+fT14O9FK16C904kiVgKf3wI2wwn17y7X1tlRERq+XqvW51zKL9pVuPKRXHaE35Suzu4RNNZrKAbtGNc7YteU4ZrWayODeDkWO2LoI6bBMt9jyFshj0NZKLbx8cOhwM02hC39+QRY06Gd2UbyLpIzKKfHEdmJToxOrKGnUWRcH3R2VMClmk2iNAC2iH44jlLv8AkRjXIVUbNoqeB7xlhHKXf8i4G9XugggbAFoOPIxxH7iiWNgmk3pI/Zv6MRAc0a3pIedt5HS0f7y7+orkZaRQuemaKLaB8d4n8WLWOIpsRZoMNvpmkBMZkEP0cffnnr/T7WjEQ05ffuviBztGnErQOnoJfIyeH1OeeOered761elZQ+fqCmgKijXOrvA4sN8n6f7xVsLu+J0b40XoffF05+YGdrFkrqNX0QunnOrweHCh+TRZCozGqMK+K/y/2RdOcpzeIdpETmMp2phZFK7/AAgdeIjpA+3YtOUrgwOsvdgVu6I0eZ4jt27UmPjn/t/9gpGZX/8AK6EBYeg/mM8zxHVncIss9TVSrnaWAwr0AAeHrvMdt8w7VveRNX7/AH4R9Ll8z3vWQW9BF20OnxoUMcfsdx1a9Xu9C+l86Xzb2NFe+jxvtRmGfdvOoTpjQRKhg5szgxmU770rnEAR1wQMI/iA+c9I85j06kvNdW6X6Ial+/uHqhb1/Pv7LOIg8h01Q+G4p4Gc5ePMZcfAe3Hp+YXh0ysmHQDQsObPGiAG6CM747hTfzD53fk/CuBX+b3bITA9e5ICtwd0Q9ouMc8jvlwvn+ccn/8AxfpHsv7UdVP2hxbUyyMbf11QcJgrAbeFSzZReIP+2dxO/wAbwa8Lk/LsPc9pb7R9xNojY3RYcgH/AOmcdR412v0fXD8ZNBGG4H8cZTW5GO/afX/UREREYgEI1RBQQCBsPxFRWD8l6xdAhp1FLQED98Ms/DuEf3YfkY6DJH/5giP+GcJOk1/6UMgzNZZIPNw4K+zg8P7C69e8724SL2fiwVSrAv0W2MfMBFGOP2JxjH7z9zwaznXClI1WBBOAsxYY0iH9ku16/wDlEYnc6iCAgbPY8a2I5W7XlmMY3RwWCQiVTc/phHKXf8ifCgc0KikdwDt8YItiOwjOP51PnVOWLZRQRwEF7MRKsI+TXGYVdnHxZZqBNogpWosGnk2xHJuEgDF4ApJArW9eCaxvJ3GLuE+yxMirIQeRSR/XER/wzjq/rru9ThI3oSs4N9+Ilk5G8unHUfXXU5yaMjoJVNzEHHtoiR2zd4y7XYjOrivvkSTBDJIAlWNeyGWEfgI4UxIvnzRqkdQ9Bri9fFo7BcQdFybJj/ta46vxuoXqWt1mEdUOX6h6+LfQ1Hc0P73sZBa4lwiIDQgBwPDNENNebT8B79pI5zExGHqcu1foPGppb+TrXygIHfxgZ05HHzVji7S78XF3C8Sz+Zupo17cytjMfng0+3t8kHH4G8zj94vY4iGxdQggQQUrdnjGj9hG/mXi+XsMjWayQT5R93cfxLaaJVjTaKV35oOfmDEw/sI3r/GWr3WXKVcHRQUI+ePe8exg3ym7d/uqqddEjn0rhmzRO2W5Q4cY4jlF6/kvZ0y90Dr3ND8446oHzxBHYrhfO3q+nY/GzdezX5pv7QONnX3Y1zZ0qEb0dfDcofv+s2+d/wDAufWoHhmCFWm1px9OA4x0RI7JvffJZoMbrTNOxDj2Gh21OXW+LuMGMwga1uXH5Bmf9O4WyzFYZNBtExlO4cD4je4a7uOrXNjGFbW2EDgegbn3x883y78/SrrjKNmmys2hgenkDded55D95YmIl0qbrK2Ca+QNmTj9cYdtTnslx8ZAPhGsjcq0NDYa344fkvBroZBdO4I/B4Plg/JevzayoMQQ2tgj8LgGO8Hd4t+uaZA4RoRoT+uZ1vtco7Fva6SLXZF0cxUuvu1tlZBabQ9rfhJRizC3O3f8T4VSqHz7fvvc85JuvcDAyCIeReevXz3MeXPp7C/csMNXmPcZjnYLjoqN0vBvUQoXSAaDNC3CgGkJwTVyFMlV1MbleBmzQZrICz+xX8OxS3lT6FvWQeCStBdY6DHkDuqjs3Nr6f0nw3etn/RZnROx+ZII373CqcOsNyIjDn6ZTAu3rIQzg+OVDpILCaCeD4YfvF1jwwghuvdQDwjR/vLhDRGCG5bs5/E+DW06rNNOgpriB7fU2NOuuLqiCgmXcHxI/K3aHAoJRBVBvB8d4RALBM3admB8sI8/hKqWVBBDGxqacJvyyIfvrgKh/nRNDjfzyLbz7xVaRDmzpMjGVlfG9/Iyfkl5HMIxCywK3xY3QhOGquIw5DTI31jh8DC4NWEcrdprzIUYUbNClY+h0+ME5QPtlnnOr9Yl4ypbxA/zIvWP5qNAsKeGvTP4FPltwmSYLW60BGtx55DqoqRcXaXbyo9IvoEDYg2hTRfT5eZPxEiZimn/AEnHM6dTSeD74Ror/qbNERiMH7s5oNrb6R/Z426/y0yy0FTM0389o/l9PlBX3a3GkCVAbBeksI0bSSsoHQ6uZpZRSrLED9uHFPvHeWiliMSHkbyKSJQCZt9eX/iXa5GJqn9/vcPZnCrb3UAadwg995Rdsnat1RM8khfktAumbojRRnk1bZ3mqOqlzm+FewwXoH4fD5s8m4GRN3o70uKNvYGyze6rt4mux/wTt59tepj/AIpH/UK3lL75mpU1XEWNDb+Zk+kO9eP1XeYMuOS9cory0xUrpqzD7j56ILGoK4mmCTbiH+jxAcnJnzvfLh+7+xdryXqtUmmaTSQn7a6HRen4NFHdSrfvEn4FY6nu9fBr3l7x/G4KDQHzfEXr2IVItk89P1fi9QvI4cLb5uxvWTl8d4NxmPGu055Y1ExpjiBhNx4+JI+U3efT4mJjKgn+lMbogWYggixh3J3GLUmzYEbNZG5pfmwPhNGP3mLVH9pP/Obnai56aILMQiHOnI65GcaugTuYZBA4RI+C3IcDfPEv/PmFJjV0ULr24ghrLcADE/MYx+875XzsjJBfnTNNadugms8GpG2EGC56Z3y0/EPyLaI8lcKq8jEaM0sTa2d43kWfiBOT9h2u48VSaRDRyaK2NzQ/Cgm0WUNh/efzqunSnMsATuaEbHCAgdz2QY/JOFJEYbZxvr0PKgTcxZBjkfbVrDGeDenhAIOYh4zoMfz+OpGy00hZa4gYPseM6D8L/veKpg8yi593AW0QeG5RybjGffqRsCVoJz0xtw03PvyLYgjyl2mzpGWEjUaMmiaH5oW6CLEfk83cLURiMK2c8URsa1kZoEHGsvCpkbJHzvq/NT0YTJkICklnTaEP9dySB85XZxwvUtM8whHM9LfoPOsz3wr10BjU0QKItOi8OdvYhDyM4Hi7O0uPD9SlRsmxhu16qCVoLjL8Q/yyo34b4fw0rmD1OXt6rTY2bWhYio0fsLv/AGLyFrXo03n2mlxKoTcQPpBG8jXCrtaumnKSrdDTTI1G8iyjaYeP2P1i1pEOKC+iZRdI4JmIPkfLE7e88f0irsZmYylqKXuZW51oIRXZ+IYkgjed/wD4FPGjrocEC1kkfBt8Y+7uFB5WeWohGhXxY3r5QdIIf1v5Xe++XoQj2bGFWmmkVp9PYdcCNrDdqCVmB4KYFG+nhoGCWYP1lh/DCd8u1FMKytEbGzX5pu6NI7C6/lVJIGWxvIYGBiHGj1G09kTCjrTazpHA8x5V89/g9ysqSJpoXOsfrGfmOvPJf5PurtdzAAghvPSG/kbB25BuUuNr8RZA97vTaE3MRAexIH8/crbMiQTpJNNE3c5FiRyd31azho+0LWRr83aHBFHT79ze2eIiQEEbRQQ/s9kGP94/+ouayEoukjjD7IEeSuHnpF3wFt0GBAf1bbeOuNarrQjdjelrRl4F/wBOsBlNMrBTTBH8DO3QNbKNY+To0EHBY0/lY3PgNXx9b7qzUT+j5pee9ICCmoC4D2e4gLTIkAMM0zgFP2wgCITmJWleNTua4gkEZj1b6bnePpu95efla93K2sgWUP8AdBG/PF9Hw3x96qZpD51rCOUBGT0tR/eXjy7XuvAWErTVBDAn4Lc/UvXJHeJTS7CwhM1jO3R90p1RII2hjeXcitlq8Ktgkb+Tk5Qu7lFDfzWNIB4ORbDkeUuEbgWNhdO+i3EZ7x0AgbWxq+nADQWWGuDLYhjPrrk6QNFlNMbZW0W0Q/SOVVeksAxvIgnEDw9miPf9QjSQpZpMopWHBbn4wRbEck46hcmMzqCFFrLK3Rz8YIyj5N/vQFrBQtZII/jm6CMjhy6HRZICooIoVjn80m69zS4B4QRYq5GrI107jZAW/wASdOfBZv7DtaTNegt7sM2/cfO4R8RcP39k6pxicXZ3GLVF5B9Ldb8LTZ3jcPvUwPEVGOiJHYs2vHXLRskStBI2PRAWFPavqjwatRtFkiVmQW+IAbvFq+8R2jEBU3vOnc7wc3tl0Z+Irr1JSHK2GN3xADf6btluraXEC69HmtyfT+TTnFMxvSUDR9HtqgZAX19DI424C974IA58R3UKWynMNVbsJ7CooLLEFC2iWxfr+nXODrmf39HL5jNDekksf1c9ff4ldidZBqaOsssfqe/ertls7O5SdQmfk7gH9YnvvKLlkKt3PJ2ksL8j8IaE2wfkTqV/01mq+8nq/f8AlJ0ufmXh23wzJZCjQNPk54jgjkOo8a4eKvGXM4fHBMuaufUXmP4zeWnuIjwp+O8ekAMykL0jv17h51Cumzrp2Z6olvOon1b+qn1W18+EXoL5sEl4hwe3A/Ih8MpiXTvq7f8AcUfNOqz3V8O9ip+oBoWWjwMLc4+tv8bxeQ0YCoTY3W7RniOwunHUKRsAI00II+Nm5/c/yl//ACJjpaVqKCW8b97xsn7/ABdwmHvZWyI2uMzEPHtv4FOnDklmvc0OK7hFYQR2K4Tr1ukqEbr1NJ88G6OPlnj9Q7+7TYiKYiCRqI0KCEVuff5GORv3N59cCdZJHf4PZAmx8FcLHh15BsAUjv45t8QZRj8n69IvHadjcrM1kr4232PGyflbvqFXOujIMrhsCNfxtrcDiED1bOVu8XcfXSJmI0kwtLGGucHG+x49YR959tMvDqRpWNwxt/NRrNAHelkcp6xcmMxh0CVQsMarjfY8d6YQlZISSK3XutN2gd6YR+59dOLCDN6SPzTc/EIlY0+87Y/+vZreuQcgrTILJJ8bbnz+lsP/AH1PpM4MMyvkVvHDW4iDjZPynrHaon6Fl5WglbG3AR258siJHYlI2wMrZ2jWHJreIE9Hz+DcKuz/AOPYq+NrU7GBRo3XwI+nAxQcPIH2na7jvMW8u1ZEY0Rtlh3WWizGvGOIH7HcdX412o+JhZYrw0zWSCP43h0jEj8n69YVKOCBUUEfxw3T5gxI/wAmTEpmCYIKE/nhrMNBiYf8pf8A8KmnE6J2h0WvcbotvcaP2Ym787Ze6Y3ODA69xweB+BH5O46teVe7G0nr6Ubo4LH44FueHWMP+7uFHEYhojJMrZE4lbDj3A3+8hWTMRDAhp2vZvOrKmUAAboJzcSefwXKkjZvB803RwQdAh9lrePvz3fEBj6usgWP/wAze+9TAw0ILWMHKOEEb92vcJYKBRsKNGv73EEbjyBuuO8vODpjpWVjWhRtxr2zEP8AI4hyfr/FSpzjRwKiMkmaezwcjJ+zcwjSqBqsppC/OnpHmCMcOqNXVrNNFWgb+QPbRHlVLqoLRtmsoTgL9sI5W7/kVOVAsbCoqcH3vhvlbvGKVrI8abR63hYPYQX66InMEn45tH9SO12XYPVrGRv9mFQBq3LSP2r1Fxw6Dufob65I79Tsy4BqspKDDmhxfG4RiSOVuF83zXhvseR3mYcVImcnm8A4HQXz76VhBTUtQvQ9jupideB7HdQnZpgHgUFNmiNuIJAwGkEP1XY/xEtX2N9L3oL1UEveQTDnzyNIIX0PDsvTfM3tXetvN8+VVspFfNITI2yh4RjlvZM5vNMrBQS3jgTdIiVZ5Neg8pdwSaa2NxAI3MD2v8il1dWs0m7ur8QQN5RZVE5pNrWDm1+kDZRlSWqVszsYaEBG8GIp6MjvOoV1lCOp7G+nckAGs2i1DI5LOXH1EuNKpAskRsIO3YczHw+1pyGk7Wt4idJBWCBNNjWE7aMdEFpnwSrTZpjVFt8Q8j53C6SqY4dz7zS4CgjcA78h0HD+SXM4jMuYW5tEZG6EHMD5HDvu/trBq5omGxseVeAD5Ryvr1Oe3OajWwSVSKHaB0tS5VBDcBpEcwdZx3r7wl3i7hWoy8rZFG67DiIPDXplR2Sz/wBRbTLO8cFlrO18cs8w2U2Dp9EEE1+dQ4+nygjbVCsaNWtNOPy/R4bY1CrTNcCNy2Lmnf8A4SKfePi02YyVCrvdm0U7fJ4cU98R3ZqqrWP/AB/7SqW9CH+dUPrtoHevvEdrt47S8gpuRHm5jpaU+TXCLAqyDAdnhzp1k+b9O8TbNys1fRNrZ3mDoerR/hqWzaq3Wt9HI74p4WHEQ8YVz4jpcsthUr76eEM0ALg8ti/5i7ExMOG74ZrAtYAcGjwR0+8d4uWbtRSZ2hc60rg/0cU+8dFkKtxY0FW3rZQCO7fL8quRETVM/v8AeoVU0yvRSRKBoW4Iie+p9+zaZZfNhLW1yNSvK8zwWZq0HEDy4UvVwxq5v6GzqcJZkyZ49qirqBEOHA4+t0G7M79e88dQXUY1WRL6tCCYFBH5u36GPyr/AKv6i8xsTIgoJRe+GJqOxOMZaeOt4h0KN9GHgBNZn+mRA+8vLvqHfiLMTniufoWBuRStDwdviBFYR93cJ3lljoUkUW3xgjyVwutvA6KCSsQabk7+IYn9muOrRenU7lMtCCcAnZ8jLCKbbtrcO0xKkzto43rIC3rOM6DZg3zOP/rrNls7Vuznc2iNcQP3vHdOfCZxZaZzSyiNAlXB1vG/xN3jLtLdNRvIpVAgnD+mRA/Ys3cJGkqGBmRSs44RGCKwj5M46j661nMpieCilbIs/GCPJuP4rtUqA+jBJI/rgn/DXCnTiaFI787dEQIsfBXHVoUsdCkd+xrOuBFj2WzcfzrScKNhGhUAWkP4eNY8s9u8Z9XFqos7NOAI0eVWZiXxrL9pu8ZdqKMbx5bORoLWWBw+B7oiGWEfuLVk5U8M6SSqRv8AGCbH9muFx0iFkQREczD8jI4d3/V+LcKvSWGuGh1s0wiCBcHrCB9rfOn+R+PcuHfNqlxtoeWxrpH2MOH8mvMvW6SE7N59p3HkcHMOMT4nVqmmYsaYlz5mR5mvRuRZjMw8f7CLzONXKC98ONhSsEPKgWYh7emBHbPr/FXbKRUSlcKigmW5/G/Obe8uPi4u4VTjYZWwQSCERzDjyOlsO5P17xeUYpr6RhoZssSPBGY8+Nf4Zw7tLtAhcxw3nWgnOpA8ofkdcYh+GgQwlYOiDot4snKsYHabRMrd0Z6Hdhd7533rVOVEM+jBAn9D13iHkRv40/XJqgNnfnXjYEEobdxUdMO03V3tfx0ALWXnJmkhvByERjAWs7BVk0uDQQbcCID1g9PtPu/xFPDZ2Cm8+80witz+S6472+S6pxTMul2fnTNBFE3SIjYj8l/sVW0asIb07nfYxD8ixH5O4xakVjxvLZpP+UWI/wB2qtfLhq+Jls0xjD58w7U9k3UYvi9eyAu1xfu0qnZo5G8iN+TjiufJIgwzG/0q6G5yRV0Ct6+EfqK+6RJcjt0Jvyl1/wBMt/mLU19+SAr4kkReCbfnh15vMdIh6nL3wtGw38FMo3xw9sxmO/Wvgqt36Db/ACm5YmcOMhuDAsw/S1bboKamJ1nWN4kaFjXOyz9azwzb19I6na94DLEra+Gg2B8Qx3grvqF9nw+H6b4W9vYvYdE8+RVukpd3nR2MOkaPyqHGHQC5d/5H/ctJ1ZG4IEbkP+I7L69MStENCoo3GDQsx/qKxIeCjfSQfL8eQS9feCWcqVh0TYI/Cr8eORbaLtV2sqlYaaEdBDwmn6PZGeJi7v6iZ5eYwlZsbMBPgbco6XWw/K3GMWM4jVSq5WojJWPjjGdcIjYj8k4/nQFllwUD7PGSP8NcfxLuuASDyKCP45o7/Qx/31wCyuFrKEfNRrdAHsh/lL/F3Hirs74IZySFlldht3EPx38H11u9YpbVe5CNx8cNUSxW3xJpyIeiZpBH+mqLLQ+pQf8ApZ/g9mIjk/JOFX5SZKmsopHi/b5IsN/Eu1PO5Dd7zcE6S123kN2VBxHq2LYJ3jdFBSKHKNpp12yxlypoAVaZBYxG6/grn8RWZ1SKutD51pnjeDc4oY/fvF2c5CvlduAOZ45vEEe0/f2azM6gzBQtZNdzd4gj3/LV848o2rQeCG6yRjSesliPk3ZEXjtLf4LK2RQAI7P5Lo+a9IotoVtW1k17viOK32Qjbrk85aKr+0s/nV0zlsmedjmM3fGivvKdapjppin6JZnMi3xctvint4OQ6c+I7WQ8vo/pxEIJ73w5058R0tWTs7vb4mRGgQTgEEFD8fGIjG8OJM/SXWCB7xLflX7xds3atxZ1NC51pQ+TnvvHeLlkKt3PjKyNGn3ugj/04lPJ4gPtLy0s7v6jtVMuwBwWCSTKwErQQCwcQeydED9lzdwvPIPUVcbD5UB3PieWe+ved6lqxG/nRG3EECzD8iyH7C7uOr8ZPmYiMqNVYdWzTH35wPsgRYj/AHdwk7cNP+uEjXTqaX4IWY0OH8ncYtOpzjVhI3RmxvWMLc+RjjrBj2NhhRud3EqV+Ih+R8s6uHdo/S3MYAlfp1O4Bu0PyawgfeXdx/23ade7OUhwTLTdfDsxUvTIiR5T/LSXRJX/ADpmhwaazEVD0yIJ87aKCwTTZpjeHSH8QI+3/IkaRCdZxvp3GyDgf6th33dwsYUl5pN6dP6H2PyMcfkMX9tMjExkstGg6M3nW96BnQY/Kf73ippxMZjynZTt171jCzEPHox+V9ek2RlQk7hUJriBhew8NsfCXeMScYUGDaLXsCB+8+kb1aZx/wA2tJgZWy2N89ZvseO9iUQ3u19YiXQ4IEbNMbHBNz5BGUEeUf8Ai9QidHBDTeeqNVwQGIfkWI/JZu4QI0MTu1uvdEDmIfkY/KevW4nMZCRsHDGwJVBOx4+kVOafc31ffqlhrUFNCNmiPxvDmMT8j3O47zFqvSEbbYJ+asEiE1HMx7jI4dyvr3i8fXK/dq8q5FBD45t8YxMO5P17y4XoI2wSx0kt5qb7Hj2I4/bKXewqpaiwIyNRvC3H0+jj78Td9Rzf27pVWk4jLNS+bRGmjhQPKHGhjkdl3zvlM6vTTQq1xA9wQfI+We+veeOoohtz1s0VmqHi5zMxKEuChj9hONxl39RdHhtBoRoRrgE3P7o7/q/FVBC3bRhVDdwQ/SOzdjTytQQsiqJqjeUOHBHgzHvrHfxLlSTq61ad531k2coPI0cjal6hb2agstcGhH5BEB3RlQSUo1i5mnp3BAJrCyi3h1iRylwuAzBDDWytFzWbnhutsP7DbvVHMatpJWRV8c974aV/Ii9MpS91RBHOO14a9feI7XZ0dEvd5bNIBp3LLt43SJBOnU0uP6y/EtFuZxGQOaayNxrl43Y+O8s0rSIdHmk2tnaMN7Zsfu3az/aTfzmpo/Sg75Q6XI2c8DmmMMmn5REXS7iIjBb0PTv7+olhpuh/tLr7p4rp3hGcZst+Tf5ayqjd8L6v0d3ebv02UJk4YkWMDVbm71w5joNaxmwvlLrh/G+1t7yOh8/anG+20Mz8l80Z9x1vIXzkw9+JfQkFNWWl0EamJ283kJI/KHO9Ec3IAMcQvS5dw3i316+oQqIL/wDj/H6jxV9W+WP7iocHyhAD2MtNTCimHIq3hHkeyJ6coa2i9XHv9H/zE7dOpI3BOklDQaQTbEcncIkNYNCojXAO307pz4W7Wky9ghsE56n8DCydwQO9D8n58ysTESpVMEN1ljYGvjKfKXrkjkrvF2fjrUlkWHc60brjW4+HkWJHefyLmTFrBIKaFKoEECbUVGO+feIziAtDqI05wF+xj94u5BWaDIIHkW0Qfc/ZfXrgAmhoUFBAlVje3IiPy/UO+9WrLfKapZStKxpsEcb/ABAm2II7HcfxJOxp2d5qogtYws/mfCpJ2GuTt17bgZ1vHsfCKqyZq2SaKIJoEDw9b4d5W7xiqsowzUVmnrJCIJg3NrkRyr9SaZy06xKoWskqgA4NzZQvJWNQviG5COd2xo/8C9KyJqasYa3nJcdGwt41pHxHau8p4yr/AP0Kf2xGhXPkrS7U39RrP2FJKyKSY+zhBADnyrxcvdxStci1kj5v9GiufHeID2VggjWH4eDuuTVKhvIeEK3YzJ8ne8k+xamqjTZO0eVgggpocG+99UYQP8R36zv1urZlSyRK9bO4DDtzk2xHeO7T7a3OxWilggRs0zw4rfZCJWxHzlosMlzQ+em+M/6PXCJWNP4RV6YYLXxOnU0RfD742MO+7uFRGca7uGb4pv50xf8Aoeycj+AdqeMY0dVN+CNayxoCVQm4iX5bFc+FW7NNUsryMr1psIvqHH4iHym6DHH8JafyLEKm+habWm7n0jlbvqF1wII1oUrHxw4+wfkZGORvPu3i5OchiEbRQSMRwHP0+tsP5V/tneqOdWxILRQQI+Nhew+KH5S7xdx9RUTOuFBaCdJK+Oe9+j9meXeLuF1O9laiCZr4d7H6P8p9Y7SL1yliHkUEi8bwVFhkY++OXj/PrbTCWG/mtF45X07+mdQ0cjebd4m3sxnDlKBtopWI/pCIug/B3GMu/wCBYvXaHsEyKVohHPfjIx+S9eptFBiVmsC1wmr3vHyflbtbTapKzaKNv45n9Z4dWdHbnubuPrrlWNvqC8rdGaQK32PtTIh2bmMZ9tZ32UiyuFWzLr5G8ocQ8d7EiOze4Vt5Hx4hJT+BJSRWmTUPXN0e1MiNR8DGKSZw0ygnTuaXBxp+IIIemRDk7jGKydI0YGCw0R5p2fPJ0f4+Mu/qLAeYKKSOzzAR/hnHUeNdru9f+v3+/wDYOBVsEgcQjfa7oMflLvq/qKNYzlejCYRHDfY8bJ+VzdwmXiWkGVgstrjcxD8sI5RcdV4RuCvjhuUZNlA+22l3/IqWFFJAVaa/CCPqNFcj9mxeb720Wg2m+IzoaxhaOP0tH7M99e88dS2TVRTTTR6Lc+Rjj9k9erb1JS9vhxvWQNxKgLcRD8Tyzy76tLso8qpCvXhBQU0/byIPDin1R2zd4v6qL3YU7vJWN5141WmtxEHh1ZT/ABEiYzGEy7lZjQuvm5x6wgjflSoct1LrWztfRnaaW4/XedXr4jk3I3MLMzFMOy6kEzXo1/HKDHxDyKbEYhKlFr3YBQNuIcaP2b3bxGzjWr6c76yBYYHtA9HBh9qfOs5dkLMaa/V1RRqC89MkwCNm4jJ3roggj4DxPUNhNCopIhEDoCK8C1/uyn1yYuDcAV7uEdHrgS9MISo3UC9EK9047fjf2Hal8ulAms/J1FzffCIihj/OWi5kGZWyKCRiOYPY6x8LdrYFkg0IIqIGs9j4a9/y1zfiD9CLe76M0ODt4HevvEdrtWMageRQqyaAO1ybYjvHVoscbZuxZwTp3MvyiIo/tLf9QcI3XqaXHZ4i9+8RpDjOCUTZpcG9sPXyHGIW6Dm+9z3660Wtv/8ApOnP+GWVQvqt/q38N4hz6tIv1XmJJvlxQGKxsYd6/dCcy26/W3/7LxauHq9K3vfgfkYwwyDTNCr4IYOJcEY9fnX+H6ltOX1HJUzBzPBXEaD3encPd5d3OYbSCbWtarbFyp9dXhZI5yb3TgLd7/HEEfhr7Dhvhr1vJuRM5Be48lOh/eEt0X7gfSP8xTtEq6tqDTmdb9I/cTtITkK0I3LY3ufSP3Ha2WDWm1r847J3G6Pw3aU0ogstNrjtz44j9xUKJLQTAFbxw1nW8e2I5T1njJjzFKaaEaaQady1P2W76v67vm0BTaolnSZ/HMHXiCW1R2zcYu7d98uR9BDqIRuskEACb7HwR1UEdks7NcUkZWCojSDeAferuNAUggVbG8tzGmRAjsdwiXFbls7Rr+uCLHkXX8tws50ddOyKCQQg4LMDjZP3m1qZrdoME6dzSOadtFqYQqmVfBOncbHC98CLYjk84rpRhVvPvNOQt64RH1d6/wBigWNrlWCsnWdz5qaBkFRYj8ncYu4Sc4jVvdss1RvWQIgLhAyQpc/nQ2ibCATve228Ldq6yR1KqaawOVoR+1GfwK3yjKxn/h3CDeERop8R3juzS7L5lVR2Cf8ADrl5k+4GU2k1NGDf0IPbwiNiuR+8drk75C/vVhdIyDsxbkfYXeYt8v2bLG9BIWTWoaFBIwaFwZ0H47z060nIwM2i14jmDR4bYj9+8s1pMBe5owo2Qb73w16+HWIUqm94FWzUOdh63jlGeTVt7OISUqu95BNe53hFbtESemEcncYxF7sKQAgufaaXARrMefGvtvLS7+otzMU05+iZzjVKTQa2OTOaE3dJVPyVxi7j7tU2ecas1O2yvBOci91CILt7cT8w7s7v668bw9Z5GzQoLKzitP7c/DuF6Tykmi+PBJXgoAJh8NAsLIOIxCJC2ti8f4yzuLhQ+mr75XhzPK8ZjR8rQWbIKfHoARZP9ZWYpuLtKN849Y+5jqVl3UY2NGBSXB/6QiJ74jvMXcKyP5sx+Q+axgUkABM3fEXpniYtGmcghNJpuskICCP64ZYR92gGI3+hEIbmK8kp8R92pGzITGBSuAFQ6REXr6ID786cYtUeTwI1gCCABN3OPakcq/xn2FuJzCVYzuxkE1olT3vGtiOVu1EsYGZFJAAXDyHphHJ+sWvKdiZ0llbYx55Nt+zXGL+2g8HQpJPbt8QIdOR+TcO7S7+un65IetNo5WfmtxGvBNkzk3GMu/rqfSVCBZFI58cNz8QIoxyOxZy7XAwD/RU86h64ZHT9j9eqk3kM2twgQQLH5Nk/KP8A+SzW9I1YGmk0FkbogcxBx6MfvOr+u8UOMSsOzTkVBI2DQBrYjlLvq1rTdMwNyKVQAW7oywjk83cJzodbRSt2c8jydwtnKmacilaH7Q/iBFtUeTuOb+AtpSF6/Ijo/G9v17ehw4jarSzp7h54jtcvHaVsayiy3aIfiR+zPLvF3Cw4spXC1kBfxz3vZk/ym7Sb1qlpE0RrnXgp81HMt6DRx9+JfvLO4XqMLy92FRSRF63P01jUb9jUi9YgwbWhQQ87dFQ6c+U9O78R2sgxfHjQUlXun7K6wfnw6sI7C79Y7XeXxEzn6E8wanqZ4Lzr3rSDjc+/hz0wgjsj95/As8w2V2G7dzcMEgjiB7eQPYkEbyN/GmJQzTQpWgo4Jp9PUdcOwjes8ZLUtLDgvPtJAF8aNtIsHEaeuYfD95G3L92uZ+LDq6y0ELnVoMRD4j+GtOLS+HkU00QO5xnTlcs3atTk8Notb4HweHOlFKstNOEKAwCCM97XpnhLtM/qM/oS0RMor3UACB2+IlGeJk9wscOPimSL5ZdEK9u4Z74Rr7h0s6esp/QhQXoyRGDWN0gl05/EXQegnSWCReON97qPx1zESDcjNojYudweCPfnLNd5jkcvYyPgra3g4z195JZUGJW02uO3OM9fLUJwoJkVv2CGvVlVubC0IjDwcVz5RdJNG6F/8ye/Ud2aNwaDMy3ocH/9suzBbG70IbkWLKh+OgZYesz+DGAr8g8v2bwutTHfCMgxr+97GjfkC1NZFNr4fX2pnklk6Xw8J2gAY4hfReznLet8rz6/7KMvsYLoG5cBue28Lm7PxF9e+OFrtw7nH0hAACwhZbpD8j7tNJeVrQgqHb1wvQpG6IHIdI7X357/ACoCnjZppvSOCMxA+kdmeqpkrG62CU4VfUa36R2ZALxo2iCojc/nojyXrFlokbkUEcAs3RjiOT9YtMgRkJvOs4jZunxDHD9hGuFZZaykqa/NEF/KJK4EjV1Q/AmQV9+zXfVqVp0GdzWQUIhnCIl5NZUlIJWmwTIT8+T9hAL9ZZWf7/GCKMfkrhaTLa9zBAggn81YdIxI/YfdvFGsZzxG6KCEdnIdBj9m92qDlCF0kgkQO4RkY/f9WmPN8sAqKCQWLxvg49iPyj9VzOuGAr3oRuF+b73jWI/ym7xf2FFeLaXU7jWSSJVotHH4QSo+JmZFi0e+LGq2Nn9rjunKriIp0cUF9JvTyh974c6VlmzU8mn2Igm8QV19dajCUrNbaKVZYBb73PX3jvFjlyrmA52RSRAAvl5nlEjTLTON5FJEIrGaRGynyXoG63uf0WH+UvVxRG62jehEdsEqMxz+dzTecgcKvqNcIl5K4d2aumEasYa0K90f0euEaFD8R3aXa75TBSsbgkiYI5v9kGOuiXssG4YHMEc4PDbEfv3qfeO0gyTkTIvHGM0CCPafv0y9cpHveUcEjb+OG+w8NemeI7s7j7xcvWaXEJ9C5574UpSQ3b5iFfRD5MLl7/8A6fyiq4v8qcsfnfQEb02i28cZ058Jd4y7UatWX4YJWhkBB7QO6p/AO7S7TrPZ5dStZJJl982EzVI982ghOvWuRFOK7e1mL0d9aesRP1BSWJJjV663ni+O2XNdn8xvYkQ/go1k8iW1gu7v3d3c3HVXa5iJ3GTYIWskrQCBm4/W+CWpA/ZH7y0u1iyVVTquZoCNq4QCFn4fDXTmn7Ld4xS2TtQE7hZa4CoNHxI/edX9dDqynUKioIID7HwT/es2SmpI0FWmAQTg4wsNa3eSbv8A7ipnERlOMEFr1fF7AREfJXH/AG0XeYiMnQqjjeemach3eRk/3dwpZaWs0trZpogcwBkY/eIAM7Gs56KJmYAGox+8To2LCmnpLQSr7zw7KOUu8Zdohx5NORmuJVwdbx7Hlnt3jLtdYezvkdBKuDrfDnT4jlX65E5AsZCojYRKjeDeUu/5E+zLqYBG5YfNfvPjh96fbW4UreWEkwQI2NuK3MD5ZEOTuP51VVMxGm7BsKtmiNZazT4j5L/YkthzQa2NTS/osw/JsRx+x5tastnKnkb6dxuhBzDgh0GOnMlTDQjJ2cZdTuHHzVM4dqjGNnMy1G8LWm3u9ezQbDXCNHxIj7u4+7U0bOtojbG1oEDwduEfuLoWc7tbAwwJU4PjiOVu1G3u5/fDZWTRCJIZmAB9fozyt3iwXa9WzxJdWjdoJkUFfg4dII+wkXqancuzr1AAmZh/VRgjwGLuPGWM5VK6+gEZNMEft2+Ibn7H/sVtokq3WErBBBQUgKgyAelckdhd3CjVrWC4TTX8cjjMQ40j9xAaoHW309dzcGYoKcfhjy09P4q7M6plte86dxyHwPcFQ6MIH+J1CzeqqWUE/OmaK3f4j9t4icUw4YjTNeppiHZ4jk6IjERAL3xI1WzQeEFtBDpyP92lxnGqrGg18Q3p2/BZmIOMKH4i5ZOVGJ26SRECCe88EFckd/1ansteG3P45nM+RQSAQPB7G1njrtkKhog2ikiDg8ItTFzhi+G0KSMtZp8b+4XNMhkEbrJJEXN3+Iugx/vLtZ8gSC5HA4wdg3MKGuqEC6yRg7wI/fvF0oYLrIfXdic/iLhpgLC0J/2xEvsO10nZl0KIfo7mevvHe+kXcMn6LAaf4X7t27QEZ1FxybF2A/FXXjga/G37oeZBZmlkwG+FBce/BIx7hej0PP7uX6talySWSxesA17AHrj7IyM9/m3fjr7TlXDnhvznnt53kYdWY1l0zCY2wfuGWxA7OEr1N3h7bhnGhBU4OkW6Gi+vkErehue1+p1a1lnAYN8atCfm0HYR/m7RZGGDTcxGwmZfEB8nI3l1/Gqt92VeEbrJBNfAajeYcP8AiIxoFUFkWWm5Q4h+O8Jd7WgDQQIKN29dufHREj8NAKGhmzTG/wCsPIjf7FpgvGza2NvzQQKhxig4cP2PN3C5EYglroVZLGqkg8rgtxA8FPfRDkrh07/feKq6qiumJjyzS3OaQq0IfIMRTvTOyuVMrbFBIIEFBHFDwe2qOyXaWqVM1ZbHOwQ/Ej/vrqUWaZp1kt4GDufSOURnIVczm9GEAm7nh1sRyr9asnKtlfNFEEFCIHh7cI7/AKhVWTlRU3LJWhEE288h6ZT+TuFx10693JAUrQQcI1uPqLYjsL1ePeTlZS1uNzSbNMbcAhaO/I+eVGydrUbjevc0vwgd0RL8SzTssK6dmV18U/Dt8Rsfw1ZT8sIx74puGaT8G4MT4izE9UZN2S+rkUacBe98OFc+StF3l7d+ZmgJlFBwmZ9xLYv18Ys2OJhup7NHWWAQP+jqzx3iI3kN6khrQoIODtFNbZTm/TvFIrFmnoQSh7ZeqNpz++GbkcIgf9HeUu3iuskdWhaNfoTLIXDyDy/w7hPjWuf3+9w80K9s4/pCN/YdJH9QfoRgqP8AJ2f/AEhEXTlVOFYJ+hMXN38gV15S0QGYTPzWi523nkih+DuMoWpjMwkc/vVwVs7aod+b70D2PhCnlnd+K4G8oqOYKeX6OnBfnTPHyiJeS/2LzSTkE6dxv9pKfKxhy+NXrQoKbFzZHPJg5+sj19reNm7T4nwrtXY+jyv9rWSb3BmQa9nERePPx7EiIE/D6t5cXHrETxIiMyOxbjRQWaZp7RzI5Hazjq1BrEPVLhG1sbPmo1vblP8AcIiMQGQQVbNLgE7H8II7FcYxdD0L86ZprTcRURG2I7C7uMZ+GpzBAja2NuJpObYZQ9iRHJWVpcPO9xafiOnBbyCZFKx5uGwf01H+0v8Aq3niPFHMZbHkloVbXYMRB4bbD/uKyfowkk9e645mjjvTCOUWFMloKFr1NIFd8sI+8Reu0vIIzXqaeeo3R35D2JEclcYz/LWMYjEFy8gofPTNLg03dBD0wjks5/IiIZACrZ3ncc33wItvBpW0A7r3r3GohG/7uP5O45hVxGIwWXNogpVcQQLMRAl6+5F1cdR9dSN+TYX6EEG++5Gto/J3Hp7v+BbzE1Y+jpoE2igkYjm8DUY/KXaw4rYJgCaRHDfY8fJ+VVjCStkQR8cN3OPYjkb88u0BQxsw2CSTH432tRjj9kf+47xUZjZzCu1LwQc0XsJYBboNO9LaR2O0eXaRebJ43bPK1FNV8Z/HDcncY0wgfeXdx1CiWTsK1rZ4mnCdugm2I7C6QGsyRgncKP31TfZiY7GH/JnDv0jtVWTNTbdCggG//wAb2zTL1LSEZRGxo84LcHSEfkhcZdvPHeJVP1+qkU0KiCcV2IqMd4P1idlMrNCC1k2h+RbELDpK+kZgZ+TkI/Hv4K9iUR7zMO1Grg3e8DNCbHzjc+BLh5g/YXlmn304hijU1e4C1ljL85jet8FKfeTs7ha5i1y85euC/OmH4doHtiO8drt65TuSkdmvc0w87g5D0wj7xStTLyVwueid3Fb742xH3ifOcaKNmIYbJomhxW7vjf4lotzMU05+iQaaTde53P7YiNj4O4eWdwodfKw1OxuCaz+hiHGRj94rIn+Fln9canbIo+PA2s63w11/mKWy+7aq/GiTPkWs8D3iG+Vu1qyFTIzAFK8Agh26LUsjv8X9hYiNZkM9DkUfD7IRJ6+8Ra/qGfoWehSqOEz2QiVt4iwYN7CAdsEvX3g7jFrukli6EEP8nemeOgD0XTtwFh0elD8/nETONWGdbkR52DP47x3n+muuMya2q/Z3X/Jq0mh+KcG9Rfj0v3V9B6hi9z+Vy+4PKpmPBcD1ZHedQve9nOW4h8/zu+fpiEEEa1xkA1h+F6z66+/fnrOiNxDNvI+9WCBK4Gufm7QP97m0pzE4I5EEFW4dII/3pjKgjcb6R6DT1Fq+7K5eesTvAG16CMmnWRp2YIFDI7C8s7S75hd1cQ01s0m1wTd02I4+8+fVrjpWaTQqLWMFmIg+kEb8SugGN4AgQAtvI6ZEfuO1zOgNBG0UrUNfj4xjqjeRt7XAxlcxgbOfg3MOCLGHD78T/A6/zFPjKdrV7oI02+jFzTafIJSdOajbHzx+8tPsO1W5LfNNjesbWYiH2TnsqypWLI2w03g7gfHeIsqsQo4LlsbrjgcQ4ywjvE2dkymgoZk0RoBrcfXkZRyWctE6qYpjLPkc01s1Ryt98IlYj8lm/sJEaJ1ZM5vPTPBDQsxU2I/JZtV2WWam7SqEFWkTW1nYR+SuOoUX+1ZQ2aTTWnm+98OKfZN82icBQXujfz3AYdh2194jpcnXQK+93l01QfBt5NsR94u3opBljp1fFHbv8b/EVl7iIS0iGs17mnBw+N/iLOxjGaOnU7Ht2h/EbH8NEaQaPfDN/PeL4GZgmjH7x27drljrGXKjF8St5+KEH2PHdBj/ADf+ou2YqdECbuLknJA6jWZKzV+ipFbmNzj7Y5QHO74n6Uv28AG+w7ViRnfCbRshEEwY8eWxvnLvGXa7Y6xlNVuzmn9Fpfgfa70zx3im8usTcAN7qEAs28l6+TDQq1gV7ocE3dEbe/Ud2aWYDGzaKSB9J0kp94K4xarRqbUzhsYFM08N2gg96P2Z24dOxLj69olcxmI0V8v2bdJGRGvzfe+GvXyckClY1ra83eII9p1AsKBQaCG18cNAt6eCPfU8RW96k6ZFgnSWCReNtO0eHUY/KP3ln9VE5cQL9Foubh6L8YUPkXj/ANx3i1M6xGf3+5a3e7icYdviL18Pybh3afbVPlh5BWm860fN7XdOahvw1NmM4dClfpJrwdweG0Y/KXa5uBQsjlYg4LdBNjT7y72/lF0GI2bRQSDwTlYkR3+LcfUU2NcnDBBUUkdnjERyjk7j+dd8tM9Ckh+dt55Fj4Nwt51bKhG0UrRiOb/ZQ0fv8Zd/UUuzqAtCClU81m6CNbRx+xZy7+oqskM4L0lkiLxvhGRj9/1fi3C3OtUQUFKwWsoUXjgW0DUY/KXaS2Eb1k+UEZP2HmP51QwFGzaGNv6H2PGseW939dAXc0MYCcBKrD+sEOsieVu/T3aVYxGsx5V1SwjbWBwODwNjbB+/HexIjv8AqPqKeJyn8qw01gUEAgjGaRlhH7itYZRvpLK4ED29/lhA/wAfqEtX5adqhzNZb3Y8D4PDqwjlLv8AhXoWmzyql3eGaZLGpelKtPy6YIKM53qjGuOr+8UMTPMK/wDEHz8C6CDooI/N0d/GCbHwTjbFptUTOYbK17qMRxufPsoPD+/6v6ijbWMDgvOvegcAhbnjeT9479OqLOfiwxUuYL+lUPrWYgf/AC7ROvdk1LWr0ATZoa4OjeIcY0yIEdju7R/du/KKCqcQtezNGzY3Gz45h64WVOPvLtevZRiMIanoRoVbXHNxDiG2xH3lwkOlb3kFrIJH54mrKItEB3VQPvPNvP4Eq8jEaKo3WgRuGVpnOwex1j8+8RO8ODwVvSSZzdv1usfHeWaxM4mFYcrG0TuLxvDo8Ne065/VI/RJJXSQKMHe98OsfCXeLSrx2kSSMir4573w5787d4tLPeXu8ij9cbmIPDijPEVN792zT+OGvXBfnUAEbtGOI7zGfbWr0UlIJ06mhxW+yERtvPvEf2kfnOG/nROz9m0HxHyWbXIjppw6ZnY2tmmINC5EfvHa5ZCpnfENoY1kTOs4zpz5NYshUzmlushriCMb1nhtizlLvq0uJiYzBgs0YQjQAve+HOvnbtYsj6h43kPSPg47pz4S7XSRo1pr8LeLINAMf+qYgdwe1feI7s7j665JwnsI4w8IsfEWspTTnriXyjP/AKrSaH4b3MYDD6K/KeG/dLp+kXoZt6PnJvKtvjGg9Npmxo3J5u4X2nK6NX53zu8xD6iCocnB2h/if2b1n117zwXhprG2/Q0fE0/2HniIJQ02iBcBbe/x34dwlsBmghGm6xg7nxP8aeFUaEFG42OEDujEj/JrjbFwNcCwGzTr3h0ggp94OmePFpMUkmNaywUg3Bj9bnTnx8WqMZULmVgoJj44czEQfHcs89Y7UjoFEbGjX9cdj4hjiOwu1tOxy2aZpyL2QJseRd3CHHsbN17NyIDIIfkcOH+J1fjICzkiVgg43HzQgMvIIAcxDbXjnmHdpZ838R546z5Ur+uCCNfm7fjX3Yu/QGvhG9JD6LP0zrJ9q9O8RsHuRBSSfl+fIo/41T5ZLQTAEHGJqC3OPSD8o/8A5FPOJnBHgvLBtC0+N7RB4c9+cza5eKaSl7oPo69m7nxNR2W7VqNs0bmmikd+wLEOKl0HUKKcYWNfCN/NaPm8IsnJH3i5pIYythoo/NPB4JiOUu8X+Gu51DCSMijZ5u0Q+CPX3iOrNAK3revlbweHPX3k1VeM0iXrcsmkBhvLeIi8FKXu8Jk7ABb+RbfiLkziMssIKEyNTQBW7vjWUd+8tPsIqnFMy3otQsMaviPzdofxL7Dz/TXYjECW5wU3Xs1+c1tRnafsKmdZzS2t1nCwYggl0pGnNJoD16ml/RZ8+I2Pj4teptDBm+gaxs6xfBtFkGOP8T+dFl/KRVmOI8vosaFNOsfvfDhXPk1yzYqh7OzKMKAQPg8O+2uWQqYRv9FZYCN4OUZ5RIzBSivhGhwWVwGRvMQ+CWxHZnV3jLv6irs2at2Op6CMC1LsINMz0Xhw1R4a0Pf/AHnMJX5ohVLaIL+hMXN3+yDWww0KSD+2CXTn8RDAAX6Lxfo6QQKGOtFgGhUUrUTc+fGnTkjkhRl38xYwRlFKwGDJ7eY3vzbgaz+2l51UsmUWsYFd72lPvHeWadZkVPdCvWn9HPxt0s/mZe3WRStsYh/ErYjkrj+dF67SbNCNrIRA9vcaOR2zd+77xdcLGURhr80ID2RsR+TuMXcefZEtWenbI42/BZ0XEPyMf9/66lsnant8XIzQIHtEPhzr527XJBWN4AoLAIJ2vrkR4dUxjOWAI3kWs8E2gAe2I5R//ImxEZymFmdlFCIBBDWbmrIgOzfH/wDIpbLHQ1V+OZG4QoIBAzt0WsSI/cTsOCBYa0Cu3OPbEfeXaa4VkkJsamkCtzFRWEEdjuMZ9tYqnFM43MFC/PeN4cPXiI+f1FHpEKBIyYbO87v6JuI1xsR+TuPN4raY6aYiUZeibNM79gII0jsVwutsTmmTvfFcB7Q/iNj4K4XJnpjLDnt/r86LqL9v4kfk839hV2blToWxBdZAmM6TwUWGj8lcOvT+NdqSI93tZ6zMb6DKLDiIOM6c+Fzl2pWlPO/XqDwSvp3EHGtiN6qX6r0ywuP/AERCAuEREpT+U7I42iCi8b974IeX4juzVF7spp1VMkMor3UIB9+BxTCOTs/3luISlY1WhA69m4jKCnw/k3aqs2ag5JrY1JJ80hH4gggV0PyaXnXCrDZBG0UkRc7oZQQK68qpbzdqkm39BOXjboMj5u0U17947T+BGENbzkxf+kIiK5H+8VN7u5Sx0KSH+/kRJ058FcYxJbeBf8OiDQfZCI/YS3WQTKK92/7fiLpz4mMQbjVnBsikeYOz0ofjvLS7+oi91l2h7Kx1FBJgjhufcQ7yr956RNmZiYG7OR8iNIje54fDbb+BdckWSevbg05mIHtXyXeu0vZKCrZocVuYIyx/3mMTKpmKdBoNBencacNNZp8Ryj5z06TOkOi4GTRNB7Am6fEU6IiinH0cFN6dzSR0dIiViPydwkxnBho3LZ45eI+TWIjDDE2i17fxw7R6h6+I5K4XFDMJhtEQdt5BDpz+Jd/doaMs9jwuVfeUWpTyYhjo56Mw3fuijVl+H0jSzz6TpCpXCBxx8QdOvHX5byr537Hzj8I/aCSoGFJMrgSsE3EQeHOnI/K3HUfXtF+icusn5PzCpsjTGB6FucfJ+/xd2rsRKQOirQnEDB83V31H1FzIeZEdG67aB8d4K46hdCkCCy0iOYPk/foBUKtCCPNyfR7EflLv+RAQILpJGDu1qMflH+LRoGnmdZIxgzFQB948VyNahxoJkjgA++ERKqO86hR6mLOCZFK2BmnxD7q4Q4xC6SwR+bt8QyOHfvodNQQILCRNVBUOIfiYcPvzyz/dWPKhV3hI3WBTfNTf/wAan+GsBnbu7REwHuvZpoUY6FP0k0fwq2mACOwQM9jOEALCkyabWSsAE3hBT7Jvm1diM5RjGtwyOAEbiNcCCjCO8xdwp/JpJnQkdx/TESyjknH86wz5Mm5FKwATN0EFPvw1QoKRs38yAK3dESevvBqTw6GbhZe7cBYNPjT35u4d2arcZhZDe6i5vCIiAGP3mUXannSrDjGCBUckTOa3gzpz47xBLOV8jCmA7+jbHx069NpZyVgCCjB3B4bk/f4tdvBSl7vI7mLxz3vgj18Pyt31CRekUjXusimnLcwAO9M8R2tzsePe7ra0c03g7195NNdzDbZHB6HQ4PpC8szOq12TR8vzG6N+QMuZywF/5owgLth1k/xLR4vSr+SWCHXud/l8R+/J/gXdKYA09m1s6Rjs8SsR+8xa7w/5S5ffOYvimfnu/C4OPY+TRAwyvjZEbRBN63wR05H5SztFNGSHO9VeaaZLUelYJuPi8OFhA/YXj527cfiL1bOMQxU6NGQgoJJEsSrBMxrc9M8Hd9R9RSWOczLtT03IpIhAXvgSUZ+Go9FbE39BwAd/JevvEVZYWhSsAFwiJFfUxdwusARrrJAGYMfTlGEd+8XDjEa6yQgLg8NtiPDvEp0GNhBBN1j3gcVz+Jdpqd7W/mtCAeEEvfr4u4TC1rRVs0uIHtA+JqN+d3HVoCBG9d5qCxHB+Uf4tx9RZ02kB3uw+nVc3QIONbeI7/7aTe3qmml5JATY3GwK3hNYR94ltlTK2aY3h98CVgLA3p3O9EFmKix5F3cKtko0Ln2nbo5iMRHyf+xJn4af9FvY23npmmtC9kCLEfk83cLJzyNdO5pfghcIdBj8lcJ6d4ab14Nbk7jM0/ZLt5/AtAOBZFBIvG8O53UNH8O8/gXOJOuD7E1JWQtfzV7zwW1/abvF3H1EXs6w5Sr5Wogg35vveM9ffh3CZCUWC9JAouaFnxxqPwl2sKXslZEFF5qZ7Hw2xH79avU1LR5pDbRHxs3cA7r5y7eO3dx+IqXHSgaIM1/WsxEHHemEd51DteWYFBMtNcVuY0wgj7aA1g00001waaBj4hamEfuKthtALDTWQAI72nX+YpLPRqpqN9E00G9BO1Fn38uPHI/h3tmrZhI2SChmmxtxKoW5ybH9muHa7MxEZkbtBviG8+80QeCStoFNbEfIal47tEbhv0r5HepPCzFhERXPg1D/AFFn6F7WhBSPRbQ/iLpWMAG/8OgOhjyJkevvJpWvVLbGNBGhytCKHbyXpnzGLuF3MSoezTkUDgAOHc5RnlPSJDDONtopWl8ELMEWr78Nc0DON5FK8ABw5+1M8fF3CMtobkUkABbfEI29M8HcO6e4WsfFl1n0ApJcBG+zBFX4Jx1H103SZy4aCZR3uyMPsvEfJXCMZqDMLIpWPN98CHQf8aQ6PBMigkYjeHc9GP36xoAJJNork+OcAglt4W7xdwt3s4S0mZJ01+d73w56+RspMSn17cHHbnHtvJrPEdsSoRtF084PjvC3fUIwBoGFRQR/0Le3JsfxLtZVHwjeko/ygoz8O4XcapRtBN/q+G+f3i7sAi3utgoo2DYcsVTL8qvQv73XP/qnB4wbmICPV+E6i4XwHszy2IjD9K9p77TL9WQvb/vHefzr7Z+bvOga1wFmK8isw7y80f7FouaZy4lbp5oPB3tPymbuF10DLWwQfc9R5HmOrXASNNCDgjgHBpGOI/cRGwCNyKCDhbwPbEcrdoBKNmmQUECBm7oHrCO//kWs5ThRuh1kg4O3xDpkRyVxi3a5lxSzSFRStCAQu2jPBroXRres8D3iGuso5dAK30jjceyVW4iHjWMGXbJyobXoI1vOrA8xB8j8L6/66qwyBeTD517z8WONZYP4hEYoXT8ubZ3HfqacxMGMfYSP/J3VP86tAoE381j/AOsRcoVBa6CC/NZwdwccqn3rOIA07soo3RMzEHhwof4l2oMLCE09JNaJU974K6qOUu+rWwxmzIgYeDh0eCZR37xU8NPfBzToUvg8HhuUd+uGMZnawOVpYB/o56Z47z/TXbKYmZJqemZFe8gAPD4ie+I7x3T/AIiRE/HMHPDcivW/L435K4dp3kpjoV7t+abu+NeSuEKPJjQr3x5rdvJdOfExiz/dnP6YkF/QiLm8IiQof4ixnXDg8E/RaPm7fTunI/h3n+ml+W1lJWRQQ807g9j46zetUthkjIreu/ZyFMBY3hBCf/JnvhtrQGmSTkU0PzsPWeCnmU6oqnDKtvWhfnvB2m8ItiPADKm92cpJywFr3M4G/kRG28otXv8AKRR94MYYyaJ3f9niP4iIiIjAMzT+dM7v/wCsrEf7tSutQvpUcbvughMx9RMhT3+69R9dWz/Kkn52830P0pIBCZiIeO6DH7x2ixjQVMJ2yI0ALgEFdOSOUu/TpdPkyQ522YQFg0eCuvrvFuyiMu1BzTkQcICOz7iC23hLt4tRnUlhNBuYB28eXBXPhLvGLCs9GqLn31j9RwSBDfEd2ajiIw61+NnNNjcYNB98Xv8Alr0Ea1om8+8I6Ojw22qPB2lmgM4IbkUQjnBxtH7K/eKNsGNMopXAggW7yXplR2O4xdx94qFELkLB+Tp+bgx8YIo/B5y7Ufl0CCdJIJF5q7WsR+Uu0xwtJORV8cC9jx7bwl3i1jZ17BMAUFPN2+nsR+Vu08ljJJjILBIvNPAIc6EH5R/i1mrE4hkKSGhBNIjnveN5TN3Cp8FvZXNorc3SLAa28JdrCnUM3IpWcdnIyfk7hDTM00LnWg4W0HkvTCP+nuFzywcYb/5edH2XjXk3Lv8AjWd6/wDUf+//AONlgsIUkHmm+yERFcj8ncYy7WtckhaFKzj+kIjbeIutPKNgV7ui9+I15O4R5MaBfP8A0IhMEZk7+P3z4MH4Nzld392tTnKWHVcOssrAYG4+YCawj5M46hRxOam1eab+ap9Fn4gTYj8nnLtac8sI2F+dOsjNzw4VyP8ANqyGFr7NvztoHHyf5uzuEBo19I2ikkgHf41Bg/KWi2mbTG402V5Jj81BNy9/DinUP5K4xl28/cSqozMf4VQ0e93K4ZsEIvi0GPcQWFwccfa7O05t5zC7plx0AIKjkk9n9IulhXuqI0D+a0IC3+JPfqJiQetC51oQFQ8PfZMlusporQtaAsPW+COvnLvGJVkZWwnfr1rG3cEOAD/jSTxZ4bRxscHg8NdOVQWync3p3Qg58CGiuVqIwAZp02hw4gAejH/f+uutrS+HkUa1kbmIRDhRB1ijHTmPOokGdTKIOEQPg8NtiPD4xIsnajMaZRSRAAsGftTEWWBUIa2ivdAA++ERevsPY7hd8hAm0UkvzcOnxF058E4d2l2jPxYcMhZFK8fO2giyD/jWcQ6DW9JH523v8jH/AH10DBB0UEcf0hjvB3HULgP6HBQMHB3r7yi7uCoVbhAge8D5R95dqpzyLuKt4QS6cqV1lML3WyMvhvab7SA+HvQdpN1kkCP30TM/FX9iPydwpOVUavo+d3mIfa7TsiIBw7058fGXf11a+SiDXD+1xnTmo8m4eXHlEBiaaFBQuX0j7u4eIPBjZtaaPAwtz4kj/MWWyAVabNNDQaQRY/s1wjyCrK2NzTQ0GIPiPqcG9f8AUR5BQ402aZpfhBHdcIjYj8lm7j6np1yZAM7MZGY2/wBYsnsLOGw8fsbn0lwuRGBuHffrQo4/C0hwOMK5yf4HuFTZpqiU1G/+Yr8ELtVz5KzShDcOkoU76xm6ABjvCXDv+NVZqmnMMOY89DL3d9/nqjbMgjFUYP8A1nnPEWsR048J923SuYYy8GCyNty5/DgHz+n3x+9tLvvFGr8srn9B4/8ArF+9TrLdyooE2igh/Q9kRU5lskr+xEEoLCo0gfwi5M4jIVdbr3G8Gf1wiP23n8ChhYDOxrY3NJ7At0RGxH5PNq2nMUxndHKXwzGGTs/BYzMNdhj9475hZsoxs1Vqxvi4WzTQ8HHdOfEdqd1L4+RRoAE1mIh8NFc+S/1FRy53mDGduhBYCD/Qtt8+8ePPw1mnef34OeTP0IJK4W0a2lGEd+8TLLWU1W52N/otL4Xaz0zyilaeRvAFK0AB3+1MI+7uE+JzJry7Cor3UI/pCIlGfh3COHrVLd8MEF/5duKLd8atvBuMX9tc/MF0FkUrD7Q/iER+e5h2pVa9lbrIPv7/ABynIZTUb0lIyD9n8IqvDDUIJhopnOYz2NdOfn3i4FfLGRa/zUFtEtvaflLvFqq+w5SbvdhUUbreADvfu0S5Krvdh1s0gG+GIXXdllJLAo3NAHt642xHeWjxarz0zhI51evMbNGqwcQY0DQIK6M/vR1paf4dOvpxE4U0t7O6dzm/N4REnv3tmlf2lP8AnZTsa02dz8OY1xsR+8xazRnojO7aXxDfzoPCZyPiLMTmMn7Gp3C16ml+Dg4LDVK6C3p3fFfhHMxHPJR+DuHnpPqJuZ6c+XGUF6czs4NZ74lGEEdjtLT8NN0iNA18M0xmzt9q+8d5aK5Gto3kU7xjBk9PBHtQP5P0i8ox5ocrP8vx5BHS4jfuYdWlndp/h0Kd6LXqib7DjunNP4O0u/vF09bzPgDChEEZueG1hHKv8YtWOuqWrOWM0ZHI0IC38msI+7uEiyjybUxguRyRg98Ij5K4TWwTTaIJwxrNItTCP3EEobgCkmEQTh/TIjv8W4StJnIYl4QpW/rCJeTcfzptnDlTILAFBHGX4+IEeScJjJeaKLpeEaDYP3EOtiB/j4xZ3MAjbKJrgFm0QR1/mXa0WsJnCYCHBwTfY+COqjlX+Mu1HZKqmE05FBIBAzeDPTCO/wCo+ogkvNO4Am7nhvlX6sLNzuyiNg8Dw9b4blHfrkBod8MPXq+7eakdmfaRFIv4S7d8xcfiKDTda6JNMbCNjb+hzGhj8k4/nTI2S7lI2FRRsCVNoh4zpz4V+8tLt4gLeMtC5+Iwdwcn7DpV2WzNRDoc6xB3CCXTnwecXXHM76MZo4zJLffefcoH+ShPLtUZ1G641Q2vVze8PkiVdO1udQGHkfHz7zz3tbs4xnCSpt0rhQWC3vD4IFmKkCn5NeR/dVv5B/8A0SeFhqOnQqfj4sqSUb6ywhmYsCCtI5RMIDjYQXS8LMP9bnTken7P/wBxd3LGmk2tvoRAJnvlR+JZu0v+0nfnGNC56r6T/tiIvfmrj/to/tJm1bEILXueHBpuPcRCJW3g1O0LBKOar4WE3Ma4vTCOSuMYq9I5Sk/OylcPXqZwGmtz8R/EtEqqemmZVFTK2NTS/CCb1wiL35u0s7hd2jUGZ2y2aogbg3RRj95i0myFRm+HkUbofe+GunK7ZCoxNORa0A4MxDvtpsAGZ+kkFhAPB4c9M8d7/AprJ2p7G8jCcQRmfzxHKXfUKpwxGgqMxxA/e8byvr1Pu2dmrIjXAW8Q10lqCxuWmuIHg0cfKO86tMeZsx2a843dFk5/Eu/qKg9aBB1hoATexPvHUszhomNEaiJnXfbCfb/K5LkupgvKHXsLw0tQcOA7MGdvoh3+MUvLsK+YOqBG0Rjg1oGPcWpg/Kb2mTETGJeYm43APKvh+zOs3cO11xnRfnTluP1vxPLO3H865GsZOQLDRP45t+N+du9rXPBhEI2iCIOoKfJrGHEE/D6tAYhG0UEfxzg+Rj8rd9X9RAJwTIq+aveeHZOP29d9QuyGV7oLp3Wnew470wjz+bWdI3BSCBc9M0jgmt0gi2I+8VUplaaHz03xQDQt0Ru28HcPM2kOsoIa2NzvsVD8iI21P4R5/qKvSC3KdUhObJbvEz5OvRIhIMqxl6P2mdau3dx/BdKur6pI3duoqK904gm8a1Bkb655h3m7v4HuV5HlczoqK90Rg2+NunPiO1lSrggmc6z9mfca9i5P4NXeUZuCGayGwdm0VGUb45ymz8RYnZsaSAggZ3cVrMRD4c9fEd479Z84pNfCkrJIRsbmkCt4TbEfeXaonMRonJwQLXqd3Fdt8R8nap0ziMliGs17nd+z1IhEfxEuIxGHWN8PLZ3jBvbD1yP93+GtWWW6mV9Fv56HhBZiH2UN8Qb/AFEWf8qMq+8ZzsFRTRrIEzQBxQx/m7S7+8WYnMNYNTtR690XAIc6c+TXYdDvh5FGxwfe+GunPk7RFm5UznZtHrQFvEEdf5n+WuRqB43kQUAB/oSsp+yv3lpZpVM5y0eNyIIAL+jbYgf471LgNvgjcwFo+TOsnXdAqo3+i7/eKgXJ95XJwGrf+iY/2eIgZR8TGKiY+OJ/3/8ATngENmskkx+u3eSKJ4+MXJiJqjP7/eSWMrZFBJgNwexv28WjLgd7vIjYgdvENevvHxa7eilJIyOvjfB4IU+TZLadqXTbr8oc7TwzcFq6+ZGzfxObeJd5nwrpbjetD/PeH1258cR3ju0SLxNTszkjp3NQFbwi2/EVE7NPQ+nU7D9vxr8RcmYpjP0G7OCm69XxK03Ma4vX3iYxGNND5ElY3DHn8bNzA4xUSI+b9IptfSd/XFIIFRSxF+14dY/P7Ytf3ZN+QjK+WxsALaH5P4i9SZxGSzxvTs2MG1+Ip3tOR8d5m7tTuvAwq02AQM0DdNsQP2PfPJpZjKChc9M7uK72QIyj5y0+4dqRQZjTGxqaTzffCI2I/J5tVUximISJfENwzQ/C3iyDH7xLpxjMNQJNLKI0eVQfY+HOnPhH6wCUb6dRt+CDj8oox/u7tVsMp2NC56D6JmIcZGP4B3Zqen5W2E75HGh4IF7DjOnPhLvq1RE5jLDOOBdO3EDC3OO6ckD/AB+rQ1BedzQjpoPBC3PiR+8xaI21ZOGwXXqd6Lf4k6DI5L/Y7UM6LAppabNMbP7YiNiP85Z3C6Ht8U0I2dzwtoHsgx+8xaKflTlZqC/Ph+DtA9kGP3jtachhOptbO8Qb4H8NV2WjNW7TL6BhoWrlksIHMQeCnPfmQf4zVG14dCveBVsbh4RrcRDx7Ygj4ju0Vd7pDNO4UkdO53ANM3QRbfiLtWemXDhmW1+/kEf6iwGE01oUEg4XhvHeKdtyy+EGz/xDXlIJ6o+vMXI5TqLj8RelE9UzH0Y2htdabGo0/NYBb2EOdGeFOJdpWkJ20QTCFK0Y2jRaimzVraKJXktW4ZWPyG36YiqxhI3oQAQR/ndvEA4GY38qVFtGvVizk7hJnPTpurKywxvPSfNO0DjFRIcjwlpcKu7jFGJZpnUS93griDuDw16+8Iprx2kjK5psFtzsGYHyflLvqECVpKvSUKPxzeIdYj9/i1JjLQMldJDT43hxEPgpT7ydncKq9jLlLO9aFRTS4rfY8a2+Yd/xpV6KWMq9O5pAbv5NsR2HbEt1YB/nRNDWt3fEfJKmZimnP0c3Q2tmidn4O0PyMn7CiPhp1DLDz0zvQ4MQQQ6cj9hG9f8AdpcziMuMuvcz4d/iViP85ZrX9oH9Q1NGA2aSKLhNiOkWTVWjCNxoKNxt+dtDj/t3CMB7BNCIjjM/ECLHxMZdrSZmzLQgAd/+9u8XcfUWVK6ghtbG61jMRjX3iO7O4ROfAalDIo8hQzLsjZIxn9qdb/K5MZXEENC51qLKB3HB/iYtJdDNrQgsuz5Azqn7DzH8S0WMDRBRuuZ0HEPHyflLiz5t347xc33KL4aKCEM9V+TRj8pnH67nV0oabRQRwzaIgS9ffh3CAEb1kh4JuUW+WED7y89Yg4tNNFBAh4Ibn6asIH7Jd9QsNBaFK4EENz78jXIj8B34iqxrlw10YJI9EFp8wE/4a4UurrKCYYJBIhNbd0ZGP3/Vre6bzhJIogjT45hxEPHth+VzdwuSFFJARsEviOI3vAxRhHzarqiJjEsOIargSMmXiIvBpXBt3EfjMH14H2vTXdu8+J7pU2l3EiaX0YaD0kI3R07+e5h28XmN+Skz4Ar3MP8A6xe/UdptlOuHawv/AIdcvG/sO1ryWxNNopWh5vaxVOR2S1TQeOYaHBJgO2+no+S5sn06zMZxIBkmi6bm7xBHv+WkSYyvb9GN1vvfDnv3aVxFFiTvd4TZ3AabufHEJ05xoR5Iyr07ncALhERdf5l2mcwN5fssAund8ToeyEbe/wDU/wACP7SPzjUWvV8R/RbfGnv3lmsx8NOvh0sb+dM8YGbojX4iTtCcSaG693xX7W7fEcn+7T7PSMnVPJ2yyeD+zkUY/wB2u2btR+djfz2PosxD7IQcfwbtJpz06tLaNhVs7vwtIcAWQfk0iNYVNk6GvR/Qt1uEqqmo1mshAQR+6fw0SGql4Ocgg3f5k/8AbIshUUjeRSs4B6GURt6+8FcO7P7at1yje6Fe6+Xxt058FcO7RQrGEEyKSJgNbuiyDVGicuEbRytH8PBxXPjvPTp0ltc1NcFo7wUfmoxuPj8RYwjvyf4EqfvYV/oW4St1kmA5mfHgj351+kpzN7vIjn53vfDnv+WqLw+kre768j9ghr195NZvC6QYJkcPPO/o17T9hev8Wp/6h36EzBeskf8Ak4sNH795/ArJ+aMIw9CkQ85m6IkK5H5O4xi5r6g/QJIrWtjbDmex47195Na5gOXsQsigmDb4gR887uPcLmxjHLdesiPxEIh3S8jwebu/nFk9ZyL0DYxNXAIdYjj7y8f4u4UnFdsWckBUUby7MQ+1MI7x3/GtTlMp5WC17ncCtZpBFsRyWcVlWYicMHq2tjcQmo3tozxOoUkRiMNsL1uHXzXw32HhxRhHza3e6OUsJXy2aAN40yId5jLv7tYq2aYysFzzzvWG7fEawj7y7Wf5eo+c5KxmvU769m5ioKM8RU7RiHFRK3TqaQMHCbYj7xNSLaSDaKNP42azrfDTzOyfAXnznGix5e8y2aoecbue1MI8A7/jWr0UkJK6dTTDzTdviVYRyWcReikaC9O52cdniVsqtnCsFDrZ3HazMP41+IuTOIyHNpna2Nar4gz1GwY+n5OpH5v/AKd2t2aarZ1aDZFK0fjfCKUMcjl+r+pZrl7PxYcpeytkVedweG+Uu1pwWiOwuOhn/vLvqEAK+G3p2/YFucex8T+dFm7U5FfrjRoerKlmCYOs8glf9NafbeO1VYzmcx/lJVs6FeTZr1e6mCN7efSmfs1p6R35NSXu8K6d2yhYecmPms7VfeURM4qgeFWzIoI/rQNIJ0gbaVWyfoq2d4QFpGi1HeO7RSzs6UENN6bxsI6wf63PXw9Tvj95Z+kXLzSMCkeCBVssTP8A1dYj9+8s1iZ/5ZV+QgEbrKFHze17Hx7NddNtCwSO4OZ7IRLyVx1CnT+R9Ckh/g3fEXX1Ha7iN1LIJrQr3kXNNbpERFD/ABLtNmfjiHGUrYQpWmA5jcfrc6DH7/q0vE5h0zK4VFBYvGwtzj2I/fpszGcOMoJkVfHODj5PyubS9nXsrZEHGDeDjeUu11zwNe8Cojn8bN9j4c9fEcpZ2aTeu0lpWwwQ2tObiIeNbeffqjGYcCCraJ+dwjR/3/FWnDNF0kh4XCMd+HcKVo+1mRAAt4N5V+8QGe7TzQtoHevh+8xbhZDOCNyKL0W0Q7+RdzjAVowVdcNuOjim4Ogs2/yOyrBY20IIfoED9MXrnJl6F4jpbX0Y3GsJp+kRHyVwoto0Vq843T+2CPJXCAwNY3EQTbxx8o5S7Wk5c3LY3rHtDgh0H4K4QC/6UxvsD8n/AAyDA6Jk0xvIohp9lT/Jrjq/qIMLxs3nommiC28ixH+7uE/SIzLjKd43+dD8IL2PyOHcncfzqeNnR5pyKglULMQ8fKPlN2uhmYwIKVnAW3n5YRyfrEJ0CyKVovHDfZDpaPyfr1zy40g2CBBSufr5oFQLDfCPyXf+WrMxnDDZK2tlYftiJW3R27FroZxttbKsvhM7ae+VSgyNwfk3hHQz5BX4a1/Wn9/UfoCEbC/NYCiz5GulOP4RazqDt0bWwWMBYd0ivv8ATWL3V2l6FkV7mMRzb4hEnQfiKfy0YgmRQWPndrug/KJ8xmYUMZJyIGYI5weG2Pj4tJvU1JO93hCjb+N+98OKffwJt4bSwvdN/OkCu9j7Uwjk7N4mVbEQPet69uDTm7memEfN2izeqKXl7vozUAbgzFq+/EXJjMYA17sJsamgCtZui2RVOKZlsrJPTqdnARuY1xdf5iKpmKZmAaCZr1Ozg3SLeNWxHzjxGIiMQF1K3TuaSI5hxGuT19yPMKRWcCN0jeOEIILzSbkTjtiJeCfYtaYUxn/DoDs8bKfdj3tZxrkFI23pfB/29945P+mrMx6jP6FlNODnIgAW/tKMb85zFwponMySGbkV6sALhERevlqyZqV0zmhQW8DFzTeEvfqDJ1n94ZV+BMSUHrJqeJQgnvgT5O4d2a5H82b/AKa1C/4dHm++EyOnI/J3GMUvlt5Bf+HUfO4RZBjp16npKwbIoJMB3a1j471U3mzFLwIKigh/bBArmn37mMYsNvciCgjhnD4098g7s1rOpbI0xnOvAA2bJFUXT8u8s3f1FiI+KZHgODG0UrTBHAuDunPjpl61SxDoulEDBydxTuqcjeXl3jLNBz3cR53CCbEgfat8tLjyao0BzrJe8ACw4SJgiNW3k3HpLj69mvNif+WY/pmIKb+Y55uDP2QY/wB5drHg8CVjWhBReOcHh1iPyt2rZjKMhgopWPrWU9RZOV0HoJhCvdRg432QIdB+DuOrUXlslBcilaMRzb39lDR+Uu8Zd/UVN7Op9IsE6TQSLxztejH7/F/YS/5g78gkFyKVoxHNvpqMfv1qNGGMrM68Rz3vhr2n5S7SSggsilaLnG5O/wAlhvgs5d/UQD0rmdJJoO0fpc7c/P8A/bT73GYPpLyTkOuB3AII9+vi1y9cpY3usOvb85uYHhr18tF6AXrsPPSObwce2W7wulyOC1puqhcBeqPJZ74jlLskdWeUfh2w0JgV64ALB14jRRniO7O4Xh71vX8BBMNZKz/+kIi6cj8ncL0Ua0gmWTR2AfHEeAXJ2Cqlj86ZpHZgz5OUfeLl5DtL5/vvTPr3qhoucFn38unU/fkjuLj8NetZRjRFVs75e8Do4LF4I3Mc7eT+AXk3q2ndaBZdBI+Fwge28ouzvDilN0JwCbiKiI6Rv21roN1rQo2/Nz9PVU5A/JqjdwILIpIi5uCo0Bz2TfFPmM4NY4Teck8EJtvUREBz5V48Spx1NAGZEC4CCqMoItvE6hdNWk0ZDBJfB4PDbbwl2iycqDjmRQSAA7fTvTPn0hpjG+jK0IlXb8aYQR2S7xbhK8gaNBmhwSEA7fT65Ecpd9R9RVxOXDBuAORwAtvPJrPExaVpMuim5FJA+/kRH7CwDHWSVRwsHXgj/YuAoabRys/BwY+IWX92cPEYyEjYRtFrH4YjlLvGXDtcwFhGwqLIW7nGdfX6tdmAzmkI6togtoGdOVwwwblpp53vfo/3dwuHMDekkqnt4QSKH+JdqcgtBv0Vj5vawrnyipvd2KVlDobaummYcFs3Cq2WnQ5jCII/5jE5S6zHpVq9S0roL9Kf71928UUrBZQ6an63GM5p1rcV6ViJnEhhBDyWRWHveb6JNXbcdm79L/YtpmATq4oTIhgxtxDbvmbr9eyhzwqQ8ggcZiA3pXo0Ou7G69rC86KHTknHkuXUde3F30Q4Td3Q3Q6hrdnAsKS163oTG156owb0lzxPLPZROgV170EY6cIA+JuOauiYjjeNF7tkUbiseXUQjL6Ikt5p6SQ9trr28YuBa3xP0piLOCtdOHPFcWWwmWSWrZnNbdbw4FDhug6uAru6ubnjbsodaJfrZgAvKgMzUVvuB3ZzN9a6uXzbjD+ptyxLt5mKZmPECYXMA/QiEfKbv6XbvCr+IxYixz9FIB/Vr3/qVLZN1Gpowc7cAdYOhrc8+8RZuVBRnBzuQh16nTVVWQqSCZ5+69SiFue4usmjf0Th7P6WKuu6pPzt/oQ32KvbveY9fHHXNfNqifvFH6ECH/8ADuL/ACgVd4qWx2ZS1ioLN/Me8n4ij5ir5ewlvFuo29uNm4gj3mfnHa0mkWVndzZRt76txBXvMqhRAMkZ+JvfVuIa95lDkanpK04j+oCrruqW9apYXqnVw2YwHmDo0z/6HXQVXMHOX7Gb1DLWNgMuvcc13bNS3jtKwk5llDBHtxs3ENe8yuxAGlx1cO5cZFrnPlZ669tcgEZoe3QEFh9N0Kol7bcaLIVEo3/w6lf9qXPICM0ED+rnv/UvFVG8uQwnhjOlfwJbF5lIhN4YTbipcgDq42NbW/eqqycndpeqqe3cL1NL+gbZ4s+69L7axafzlVP8ub3FG2Unyg6uOp50Xr/v8X0VNTOZlO8jX6Ey+3hJh92+428zg/5KiPnn/t/9qDX/AMOrj4cb9N82pvDpVn6LPPhxt1zXzarZjUEFnSWH/wBYmfRsJiUAzrJAf6uK/wCpWY+af39AcjTq4rG9rUDhzxXFA8eYPG6KlsmqleazBeWMdepdxr0yr3llYBM/Oq4+GO95rwbr0v8AYuggPioK75j1jx7c3P6k+NlC3njIDYRTelpoI4suLC8ucP2WLzCnswMYBIkEdDdBl3Fj7q6/XzWD/kjhz/ywvvuwT8gvc4RuhUzJcWvHg2ECChnWUD5SUqFHk3Gf+HMv9svCbp9xtxamsvmlNVsA3oQQDB743T/v2XTtjG/2KlR9WMQxcgQd1c9STGzLbjwbCkst26mRrf8AyuuGe7jV1h7mwq2fIjllrJB/N7piTpw+47j2kxIRNZ+a3/zp79xcs/5XTUtYs2f8IjnvriYi7tuPE4f+a5P3sf6c/QBwX/hrNER22yFuea4rRdvfvHKfwQEutspfjcQuM7cQToXXhFqfvEsfgHsoYo2LtuOh+bZS7eu0uMwTo6qF/wD/AJYe3PcqV6/50X5XbpheXbqVZQ5hvsS9uu7aO187ZTouqHC0GD/riCsZGlDFOpge3GyLBMTxKG9WQTkDFRq75j1gZV1c/rs1ZVt/4/8AaN86Q9jD9UaA9J6Led3/AN9/pXH9i9my+Z5Nez6WknTIx/UhX/JePevXp3SC9ZIx/Ur38Ncq3j9+HFC9xTqEPbjZw/iO10G5idXDgiLvQ2WDcbmegqWd0b0vlJ48G6DbuIuua8os7SqDcvrt7e6h/Nt0mNPbbjyZT/mn9/VOA+dXDp6Q6uNgYd1Y8WLWgfvh/pC6detuIcLzLPBpdmqqe3xWYZqAh+1WQFzzPEpPDQU0abEvg7Hg7rBc/wBjFW4tL6X6an9rAurlzxMs0WX8pcq1reTs3LoSzg0FdWPFi1LZNVBzFisPMbmhuJ4kSFhNOnQ+H7Vra69KuZDAhzcPJ6Ih12zFVoo3M9j9pa8gxG/03I/rL/8AdEfKBJuxs1xDm+E/5a5ZO1HMP504O2Lr6NhZUFBOsZH9Zuvu1qN0wE09G90Oz+kn/wBDp5gUrWz0H/h4f/XTpUz8/wC/8M+FsB1vc/Far0b/2Q==" alt="Chân dung Nguyễn Văn Tuyến">
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
            với <b>Khoảng 24.000 dòng code</b> do AI Agent viết dưới sự chỉ đạo của tôi.
          </p>
          <div class="hp-meta">
            <span>📞 0961 778 150</span>
            <span>✉️ nguyen.support999@gmail.com</span>
            <span>🎂 1977</span>
          </div>
          <div class="hp-cta-row">
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
          <div class="hp-metric-num">24.000<small>+</small></div>
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
          <p>Từ một công cụ nội bộ, phát triển thành nền tảng quản trị nhân sự đa doanh nghiệp (multi-tenant), với khoảng 24.000 dòng code do AI Agent viết dưới sự chỉ đạo của tôi.</p>
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