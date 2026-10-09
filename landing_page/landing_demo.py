# -*- coding: utf-8 -*-
"""
landing_page/landing_demo.py
=============================
Landing Page CV Ứng tuyển Vị trí Trợ lý Phó Chủ tịch HĐQT - Nguyễn Văn Tuyến
"""

import streamlit as st
import streamlit.components.v1 as components

def render():
    _landing_html = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600;700&family=Fraunces:opsz,wght@9..144,500;600;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
:root {
  --navy: #0F1B2D;
  --navy-light: #1A2A42;
  --gold: #C9A14E;
  --gold-light: #E2C97E;
  --gold-dim: rgba(201, 161, 78, 0.12);
  --paper: #F8F9FA;
  --text-dark: #1F2937;
  --text-muted: #4B5563;
  --border: #E5E7EB;
  --white: #FFFFFF;
}

#cv-landing * { box-sizing: border-box; margin: 0; padding: 0; }
#cv-landing {
  background: var(--paper);
  color: var(--text-dark);
  font-family: 'Be Vietnam Pro', sans-serif;
  font-size: 14px;
  line-height: 1.6;
  border-radius: 12px;
  overflow: hidden;
  border: 1px solid var(--border);
}

.container { max-width: 900px; margin: 0 auto; padding: 0 24px; }

/* --- HERO SECTION --- */
.hero {
  background: linear-gradient(135deg, var(--navy) 0%, var(--navy-light) 100%);
  color: var(--white);
  padding: 40px 0;
  border-bottom: 3px solid var(--gold);
}
.hero-content {
  display: flex;
  align-items: center;
  gap: 32px;
}
.avatar-wrap {
  flex: 0 0 160px;
  height: 190px;
  border-radius: 10px;
  overflow: hidden;
  border: 3px solid var(--gold);
  box-shadow: 0 8px 20px rgba(0,0,0,0.3);
}
.avatar-wrap img { width: 100%; height: 100%; object-fit: cover; }
.hero-text h1 {
  font-family: 'Fraunces', serif;
  font-size: 24px;
  color: var(--gold-light);
  letter-spacing: 0.5px;
  margin-bottom: 6px;
}
.hero-text .subtitle {
  font-size: 14px;
  font-weight: 600;
  color: #9CA3AF;
  text-transform: uppercase;
  letter-spacing: 1px;
  margin-bottom: 12px;
}
.hero-text .core-msg {
  font-size: 14px;
  font-style: italic;
  color: #E5E7EB;
  background: rgba(255, 255, 255, 0.08);
  padding: 10px 14px;
  border-left: 3px solid var(--gold);
  border-radius: 0 6px 6px 0;
  margin-bottom: 18px;
}
.contact-info {
  display: flex;
  gap: 16px;
  flex-wrap: wrap;
  font-size: 12px;
  color: #D1D5DB;
  margin-bottom: 18px;
}
.cta-row { display: flex; gap: 12px; flex-wrap: wrap; }
.btn-primary {
  background: linear-gradient(110deg, var(--gold) 0%, var(--gold-light) 100%);
  color: var(--navy);
  font-weight: 700;
  padding: 10px 20px;
  border-radius: 6px;
  text-decoration: none;
  font-size: 13px;
  display: inline-block;
  transition: transform 0.2s;
}
.btn-primary:hover { transform: translateY(-2px); }
.btn-secondary {
  background: transparent;
  color: var(--white);
  border: 1px solid var(--gold);
  font-weight: 600;
  padding: 10px 20px;
  border-radius: 6px;
  text-decoration: none;
  font-size: 13px;
  display: inline-block;
}

/* --- SECTION STYLES --- */
.section { padding: 36px 0; border-bottom: 1px solid var(--border); }
.section-title {
  font-family: 'Fraunces', serif;
  font-size: 18px;
  color: var(--navy);
  margin-bottom: 20px;
  display: flex;
  align-items: center;
  gap: 10px;
}
.section-title::before {
  content: "";
  display: inline-block;
  width: 12px;
  height: 12px;
  background: var(--gold);
  border-radius: 2px;
}

/* --- BLOCK 1: PHILOSOPHY --- */
.philosophy-box {
  background: var(--white);
  border: 1px solid var(--border);
  border-left: 4px solid var(--navy);
  padding: 20px;
  border-radius: 6px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.philosophy-box p {
  font-size: 13.5px;
  color: var(--text-dark);
  line-height: 1.7;
}

/* --- BLOCK 2: CORE COMPETENCIES --- */
.grid-2x3 {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
}
.card {
  background: var(--white);
  border: 1px solid var(--border);
  padding: 18px;
  border-radius: 8px;
  transition: box-shadow 0.2s;
}
.card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
.card-icon { font-size: 24px; margin-bottom: 8px; }
.card h4 { font-size: 13.5px; color: var(--navy); margin-bottom: 6px; font-weight: 700; }
.card p { font-size: 12px; color: var(--text-muted); line-height: 1.5; }

/* --- BLOCK 3: TIMELINE --- */
.timeline { position: relative; padding-left: 20px; }
.timeline::before {
  content: "";
  position: absolute;
  left: 4px; top: 8px; bottom: 8px;
  width: 2px;
  background: var(--border);
}
.tl-item { position: relative; margin-bottom: 24px; }
.tl-item:last-child { margin-bottom: 0; }
.tl-item::before {
  content: "";
  position: absolute;
  left: -20px; top: 4px;
  width: 10px; height: 10px;
  border-radius: 50%;
  background: var(--gold);
  border: 2px solid var(--white);
}
.tl-date {
  font-family: 'Space Mono', monospace;
  font-size: 11px;
  font-weight: 700;
  color: var(--gold);
  margin-bottom: 2px;
}
.tl-role { font-size: 14px; font-weight: 700; color: var(--navy); }
.tl-company { font-size: 12px; font-weight: 600; color: var(--text-muted); margin-bottom: 6px; }
.tl-desc { font-size: 12px; color: var(--text-dark); }
.tl-desc li { margin-left: 16px; margin-bottom: 4px; }

/* FEATURED PROJECT */
.project-box {
  background: linear-gradient(135deg, var(--navy-light) 0%, var(--navy) 100%);
  color: var(--white);
  padding: 22px;
  border-radius: 8px;
  margin-top: 24px;
  border: 1px solid var(--gold);
}
.project-box h4 { color: var(--gold-light); font-size: 15px; margin-bottom: 8px; }
.project-box p { font-size: 12.5px; color: #E5E7EB; margin-bottom: 12px; }
.tag-group { display: flex; gap: 8px; flex-wrap: wrap; }
.tag {
  background: rgba(201, 161, 78, 0.2);
  color: var(--gold-light);
  font-size: 11px;
  padding: 3px 8px;
  border-radius: 4px;
  border: 1px solid var(--gold);
}

/* --- BLOCK 4: EDUCATION & SKILLS --- */
.grid-2col { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
.info-group { background: var(--white); border: 1px solid var(--border); padding: 18px; border-radius: 8px; }
.info-group h4 { font-size: 14px; color: var(--navy); margin-bottom: 10px; border-bottom: 1px solid var(--border); padding-bottom: 6px; }
.info-list { list-style: none; }
.info-list li { font-size: 12.5px; margin-bottom: 8px; color: var(--text-dark); display: flex; align-items: flex-start; gap: 8px; }

@media (max-width: 768px) {
  .hero-content { flex-direction: column; text-align: center; }
  .grid-2x3, .grid-2col { grid-template-columns: 1fr; }
  .contact-info, .cta-row { justify-content: center; }
}
</style>

<div id="cv-landing">
  <!-- HEADER / HERO SECTION -->
  <section class="hero">
    <div class="container">
      <div class="hero-content">
        <div class="avatar-wrap">
          <img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAHPBEwDASIAAhEBAxEB/8QAHQAAAgIDAQEBAAAAAAAAAAAAAAYEBQEDCAECAv/EAFkAAAEAMABA0OAgYIBAEBCQEAAgMEBREGElETFzURFSFUU2Fyc4OTs9IHNkF1dJHTFCIyNVZXgZGhI1Jic4KxweFCc3STosLD8CQzNjdE4SVDY2S00f/EABoBAAIDAQEAAAAAAAAAAAAAAAACAQMEBQb/xAAtEQACAQIFAwMDBQEBAAAAAAAAAAECERIDITFBE1FxIjJSYYGxFCIzQpGhwdH/2gAMAwE3253o4REAFmC4B0REBERAREQEREBERAREQEREBERAREQEREBERAREQZaLlsPwK3AadJy2H4FbgCUpCIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgDhA2ERABEQAREQEREBERAEREBERAREQEREBERAREQEREAERAGWi5bD8CtwGnScth+BW4AlKQiIgAiIAIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgDgG4DhEQAREQAREQEREBERABEQAREQEREBERAREQEREBERABlox2w5f91sA047Ycv+6UgCUpCIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIiACIgIiICIgDgG4DhEQAREQAREQEREBERABEQAREQEREBERAREQEREBERABlox2w5f91sA047Ycv+6UgCUpCIiACIgIiICIvE08UDdr43sByA4sAfxKA94sL9pD/AOlwf/sX90/ak/6XB/8AsX90A3RYX7SH/wBMA/6S/un7SF/p0H/J/dQDRlhftIT/AKXB/wDsf3U/akP6XA/8P7oA3RY37SH/ANMA/wCk/un7SB/0yB/5f3S0A0ZYX3e/pL+5+4vA9kI1w7/cgf8AdE4BoSwws+DTh/EaU6aPZAM/G4I8a3sAD4oUag4xS34yD4oUQ4xS34yE4N2i1qXh/pC1I8G5S2P9IGSAnAtkWtTFFXj1P2U8pC1O/2sDSoI/S93SADL0p3+1hKVDs4NymM6mEa3A4U2n4U86BvC3aIpt4fC1/qSjh4fC2p/S33qAAtY63XfM39SmtxXy4P4m/qQDdFq0/CnnQ3L2084A31X9Tmk2N0p33qU91j4UBI3iAtSmInG99/A39SpC1X3/3sC1/1393A39SR90vS32sDf1f3/eU3/e/oE92BIn/d2H3v6G+6YpBw+m/e3I6A2/e4N5m4v/pY/3tzugI3mU2p019/335I+6/R/qH4O4p329pBve8d9i3/AHpA39Cnv8XUf97Uf3q/p3u4v3r/ALu3ve6fD/C/dvdR+m7v/wBAe9O/A0v+90hS/e6xAnf5U2d43/ve6/vF/f04Tvd/e883e/u73Ue7X3d8xS/ed+2x/vd38z3pAn6b0i/v2fD4t/93u/vcLve7/vd/f83/5vdvh+9qA/o/333333333333+e94/z8/o/3333333x/ve/e6X4v7t/ve7+/e/vf4vvd/m996sB/e9+4pAn+926A2A4I/3v/e6f7vf5vf96RAn4fvdAn4fvdSBOkCdgS/dO4J1A/eO83ve894+9veN5ve983/vS34O8f3i/wB7+m+/e5e4S4Jz39O947ve/ve7+93957pAn6O98A3ve+/f/eN73/4ve97v7/e983vvd4+9A3/dO/ve6A2C2C/A4A2AeI3/d/vd/f/AHjf64G9/fvePve+/fe6XvHe5e8d533ve8S25Anv3z33ve8d53v98A/vd4+/vG8S4H33e8b53u4B39CkB39A/v49+A/42A/42A398A/3SBO8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 In whole, standard Python file code formatted properly into a single snippet.