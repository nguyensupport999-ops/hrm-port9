"""
photo_card_gender.py
---------------------
Module con cho menu "🖼️ Tạo ảnh thẻ NV" trong app.py (HRM-Port).

Cách dùng trong app.py:
    import photo_card_gender
    ...
    elif menu == "🖼️ Tạo ảnh thẻ NV":
        photo_card_gender.render()

Chức năng:
  - User upload 1 ảnh tham chiếu (khuôn mặt + tóc).
  - Chọn 1 mẫu target có sẵn trong thư viện (assets/targets/).
  - Chọn màu nền: "Trắng tinh" hoặc "Phông xanh".
  - Nhấn "Tạo ảnh thẻ" -> gọi Google Gemini (nano banana) ghép mặt/tóc từ ảnh
    tham chiếu vào đúng dáng/trang phục của mẫu target, đổi nền theo lựa chọn.
  - Xem kết quả PNG trong khung cố định + nút tải xuống.

Lưu ý: module này KHÔNG gọi st.set_page_config (đã được gọi 1 lần ở app.py).
Toàn bộ session_state key đều có tiền tố "pcg_" để không đụng với các
màn hình khác trong app chính.
"""

import base64
import io
import mimetypes
from pathlib import Path

import requests
import streamlit as st
from PIL import Image

# ──────────────────────────────────────────────────────────────────────────
# CẤU HÌNH
# ──────────────────────────────────────────────────────────────────────────
MODULE_DIR = Path(__file__).parent
TARGETS_DIR = MODULE_DIR / "assets" / "targets"

# Model image-generation của Google (Gemini "nano banana").
# Nếu Google đổi tên model, chỉ cần sửa dòng dưới đây.
GEMINI_MODEL = "gemini-2.5-flash-image"
GEMINI_ENDPOINT = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"

# Danh sách mẫu target — mỗi mục gồm: file ảnh + nhãn hiển thị
TARGETS = [
    {"file": "t01_nam_dongphuc1.jpg", "label": "Nam · Đồng phục (1)"},
    {"file": "t02_nam_dongphuc2.jpg", "label": "Nam · Đồng phục (2)"},
    {"file": "t03_nu_vest_hong.png", "label": "Nữ · Vest hồng"},
    {"file": "t04_nam_vest_soc_kinh.png", "label": "Nam · Vest sọc, kính"},
    {"file": "t05_nam_vest_xanhdam_kinh.png", "label": "Nam · Vest xanh đậm, kính"},
    {"file": "t06_nam_vest_xanhduong.png", "label": "Nam · Vest xanh dương"},
    {"file": "t07_nu_vest_xam.png", "label": "Nữ · Vest xám"},
    {"file": "t08_nu_aodai_trang.png", "label": "Nữ · Áo dài trắng"},
    {"file": "t09_nam_tocbat_nenxanh.jpg", "label": "Nam · Nền xanh"},
    {"file": "t10_nam_vest_trang.jpg", "label": "Nam · Vest trắng"},
    {"file": "t11_nu_aosomi_trang.jpg", "label": "Nữ · Áo sơ mi trắng"},
    {"file": "t12_nam_vest_den.png", "label": "Nam · Vest đen"},
    {"file": "t13_nu_vest_den.png", "label": "Nữ · Vest đen"},
]

BACKGROUND_OPTIONS = {
    "Trắng tinh": (
        "a pure solid white background (#FFFFFF), even studio lighting, "
        "no shadows, no gradient, standard professional ID-photo white backdrop"
    ),
    "Phông xanh": (
        "a solid standard ID-photo blue background (similar to RGB 0,120,190), "
        "even studio lighting, no shadows, no gradient, standard passport/ID-photo blue backdrop"
    ),
}

_CSS_INJECTED_KEY = "pcg_css_injected"


# ──────────────────────────────────────────────────────────────────────────
# CSS (chỉ inject 1 lần / phiên)
# ──────────────────────────────────────────────────────────────────────────
def _inject_css():
    if st.session_state.get(_CSS_INJECTED_KEY):
        return
    st.session_state[_CSS_INJECTED_KEY] = True
    st.markdown(
        """
        <style>
        .pcg-title{ font-size:26px; font-weight:800; margin-bottom:2px; }
        .pcg-sub{ font-size:14px; color:#6B6B8A; margin-bottom:18px; }
        .pcg-section-label{
            font-size:14px; font-weight:700; margin:18px 0 10px 0;
            display:flex; align-items:center; gap:6px;
        }
        .pcg-target-card{
            border-radius:12px; overflow:hidden; border:2.5px solid transparent;
            box-shadow:0 4px 16px rgba(108,99,255,.10); background:#fff; margin-bottom:6px;
        }
        .pcg-target-card.selected{
            border-color:#6C63FF;
            box-shadow:0 0 0 4px rgba(108,99,255,.18),0 8px 28px rgba(108,99,255,.22);
        }
        .pcg-target-card img{ width:100%; height:150px; object-fit:cover; display:block; }
        .pcg-target-card .foot{
            padding:6px 8px; font-size:11.5px; font-weight:600; color:#6B6B8A;
            text-align:center; border-top:1px solid #F0EFFF;
        }
        .pcg-target-card.selected .foot{ color:#6C63FF; }
        .pcg-render-btn button{
            background:linear-gradient(135deg,#6C63FF 0%,#4FC3F7 100%) !important;
            color:#fff !important; border:none !important; font-size:16px !important;
            font-weight:700 !important; padding:12px 0 !important; border-radius:50px !important;
            box-shadow:0 6px 24px rgba(108,99,255,.38) !important;
        }
        .pcg-result-frame{
            width:100%; max-width:340px; aspect-ratio:3/4; margin:0 auto;
            background:#fff; border-radius:16px; border:1.5px solid #E0DEF7;
            box-shadow:0 8px 32px rgba(108,99,255,.18); overflow:hidden;
            display:flex; align-items:center; justify-content:center;
        }
        .pcg-result-frame img{ width:100%; height:100%; object-fit:cover; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────────────────────────────────
# HÀM XỬ LÝ ẢNH / GỌI GEMINI
# ──────────────────────────────────────────────────────────────────────────
def _to_png_bytes(raw_bytes: bytes) -> bytes:
    """Chuẩn hoá ảnh đầu vào (bất kỳ định dạng) -> PNG bytes."""
    img = Image.open(io.BytesIO(raw_bytes)).convert("RGB")
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def _build_prompt(background_desc: str) -> str:
    return f"""You are given two images. Image 1 is the REFERENCE photo of a real person. Image 2 is the TARGET template photo.

Task: Generate a new photorealistic portrait that keeps EVERYTHING from Image 2 (the target template) exactly as it is — body pose, body shape, outfit/clothing, arm position, framing/crop — EXCEPT the face, the hairstyle, and the background.

1) Face & hair: Replace the face and hairstyle in Image 2 with the face and hairstyle from Image 1.
   - Copy the facial identity from Image 1: face shape, eyes, eyebrows, nose, mouth, skin tone, and any distinguishing features.
   - Copy the hairstyle, hair length, hair color and texture from Image 1, adapted naturally to frame the face/shoulders as shown in Image 2.
   - Adjust head angle, head size and neck alignment so the new face fits naturally onto the body/collar in Image 2 — same head tilt and gaze direction as Image 2.
   - Match lighting direction/color temperature on the face to the studio lighting of Image 2 so the blend looks seamless (no visible seams around jaw, neck or hairline).
   - Keep the same neutral/professional facial expression style as Image 2.

2) Background: Replace the background with {background_desc}.

Output: one single photorealistic ID/corporate headshot photo, portrait orientation (3:4), high detail, sharp focus, studio quality, no text, no watermark, no logo."""


def _call_gemini_image_edit(api_key, ref_bytes, ref_mime, target_bytes, target_mime, background_desc) -> bytes:
    prompt = _build_prompt(background_desc)
    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"inline_data": {"mime_type": ref_mime, "data": base64.b64encode(ref_bytes).decode()}},
                    {"inline_data": {"mime_type": target_mime, "data": base64.b64encode(target_bytes).decode()}},
                    {"text": prompt},
                ],
            }
        ],
        "generationConfig": {"responseModalities": ["IMAGE"]},
    }
    resp = requests.post(GEMINI_ENDPOINT, params={"key": api_key}, json=payload, timeout=120)
    data = resp.json()

    if resp.status_code != 200 or "error" in data:
        msg = data.get("error", {}).get("message", "Lỗi không xác định từ Gemini API.")
        raise RuntimeError(msg)

    candidates = data.get("candidates", [])
    if not candidates:
        raise RuntimeError("Gemini không trả về kết quả. Vui lòng thử lại.")

    for part in candidates[0].get("content", {}).get("parts", []):
        inline = part.get("inlineData") or part.get("inline_data")
        if inline and inline.get("data"):
            return base64.b64decode(inline["data"])

    raise RuntimeError("Không tìm thấy ảnh trong phản hồi của Gemini.")


# ──────────────────────────────────────────────────────────────────────────
# HÀM RENDER CHÍNH — gọi từ app.py: photo_card_gender.render()
# ──────────────────────────────────────────────────────────────────────────
def render():
    _inject_css()

    st.markdown('<div class="pcg-title">🖼️ Tạo ảnh thẻ nhân viên bằng AI</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="pcg-sub">Upload ảnh chân dung, chọn mẫu có sẵn, AI sẽ ghép mặt/tóc vào đúng trang phục và đổi nền.</div>',
        unsafe_allow_html=True,
    )

    # ── SESSION STATE (tiền tố pcg_) ────────────────────────────────────
    st.session_state.setdefault("pcg_selected_target", None)
    st.session_state.setdefault("pcg_result_bytes", None)
    st.session_state.setdefault("pcg_error_msg", None)

    # ── API KEY ──────────────────────────────────────────────────────────
    api_key = st.text_input(
        "🔑 Google Gemini API Key",
        type="password",
        placeholder="AIza...",
        help="Lấy API key tại https://aistudio.google.com/apikey",
        key="pcg_api_key",
    )

    with st.expander("❓ Cách lấy Google Gemini API Key (miễn phí)"):
        st.markdown(
            """
            1. Truy cập **[aistudio.google.com/apikey](https://aistudio.google.com/apikey)** và đăng nhập bằng tài khoản Google.
            2. Nhấn nút **"Create API key"** (hoặc **"Get API key"**).
            3. Chọn 1 project có sẵn hoặc để Google tự tạo project mới.
            4. Copy chuỗi key vừa tạo (dạng `AIzaSy...`).
            5. Dán vào ô **"Google Gemini API Key"** ở trên.

            ⚠️ Lưu ý:
            - Không chia sẻ API key cho người khác.
            - Key chỉ được dùng trong phiên làm việc hiện tại, không được lưu lại trên máy chủ.
            """
        )

    # ── BƯỚC 1: UPLOAD ẢNH THAM CHIẾU ───────────────────────────────────
    st.markdown('<div class="pcg-section-label">1️⃣ Ảnh tham chiếu (khuôn mặt & tóc)</div>', unsafe_allow_html=True)
    ref_file = st.file_uploader(
        "Tải lên ảnh chân dung rõ mặt",
        type=["png", "jpg", "jpeg"],
        label_visibility="collapsed",
        key="pcg_ref_file",
    )
    if ref_file is not None:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.image(ref_file, caption="Ảnh tham chiếu", use_container_width=True)

    # ── BƯỚC 2: CHỌN MẪU TARGET ──────────────────────────────────────────
    st.markdown('<div class="pcg-section-label">2️⃣ Chọn mẫu target</div>', unsafe_allow_html=True)

    cols_per_row = 4
    rows = [TARGETS[i:i + cols_per_row] for i in range(0, len(TARGETS), cols_per_row)]

    for row in rows:
        cols = st.columns(cols_per_row)
        for col, item in zip(cols, row):
            target_path = TARGETS_DIR / item["file"]
            with col:
                is_selected = st.session_state.pcg_selected_target == item["file"]
                if os.path.exists(target_path):
                    with open(target_path, "rb") as f:
                        img_data = base64.b64encode(f.read()).decode()
                        # ... code hiện tại
                else:
                    st.error(f"❌ Không tìm thấy file ảnh: {target_path}")
                    st.info("💡 Vui lòng kiểm tra lại thư mục 'static' hoặc đường dẫn file ảnh.")
                    return
                    b64_thumb = base64.b64encode(f.read()).decode()
                mime = mimetypes.guess_type(item["file"])[0] or "image/jpeg"
                css_class = "pcg-target-card selected" if is_selected else "pcg-target-card"
                st.markdown(
                    f"""
                    <div class="{css_class}">
                        <img src="data:{mime};base64,{b64_thumb}">
                        <div class="foot">{item['label']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                btn_label = "✓ Đã chọn" if is_selected else "Chọn mẫu"
                if st.button(btn_label, key=f"pcg_btn_{item['file']}", use_container_width=True):
                    st.session_state.pcg_selected_target = item["file"]
                    st.session_state.pcg_result_bytes = None
                    st.rerun()

    # ── BƯỚC 3: CHỌN NỀN ─────────────────────────────────────────────────
    st.markdown('<div class="pcg-section-label">3️⃣ Chọn màu nền</div>', unsafe_allow_html=True)
    bg_choice = st.radio(
        "Chọn màu nền",
        options=list(BACKGROUND_OPTIONS.keys()),
        horizontal=True,
        label_visibility="collapsed",
        key="pcg_bg_choice",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── NÚT RENDER ───────────────────────────────────────────────────────
    can_render = ref_file is not None and st.session_state.pcg_selected_target is not None

    st.markdown('<div class="pcg-render-btn">', unsafe_allow_html=True)
    render_clicked = st.button(
        "🎨  Tạo ảnh thẻ",
        disabled=not can_render,
        use_container_width=True,
        key="pcg_render_btn",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    if not can_render:
        st.caption("⚠️ Vui lòng upload ảnh tham chiếu và chọn 1 mẫu target trước khi render.")

    if render_clicked and can_render:
        st.session_state.pcg_error_msg = None
        st.session_state.pcg_result_bytes = None
        if not api_key:
            st.session_state.pcg_error_msg = "Vui lòng nhập Google Gemini API Key trước."
        else:
            try:
                with st.spinner("Đang xử lý ảnh, vui lòng chờ trong giây lát..."):
                    ref_raw = ref_file.getvalue()
                    ref_png = _to_png_bytes(ref_raw)

                    target_path = TARGETS_DIR / st.session_state.pcg_selected_target
                    with open(target_path, "rb") as f:
                        target_raw = f.read()
                    target_mime = mimetypes.guess_type(str(target_path))[0] or "image/jpeg"

                    background_desc = BACKGROUND_OPTIONS[bg_choice]

                    result_bytes = _call_gemini_image_edit(
                        api_key=api_key,
                        ref_bytes=ref_png,
                        ref_mime="image/png",
                        target_bytes=target_raw,
                        target_mime=target_mime,
                        background_desc=background_desc,
                    )
                    st.session_state.pcg_result_bytes = _to_png_bytes(result_bytes)
            except Exception as e:
                st.session_state.pcg_error_msg = f"Lỗi: {e}"

    if st.session_state.pcg_error_msg:
        st.error(st.session_state.pcg_error_msg)

    # ── KẾT QUẢ ──────────────────────────────────────────────────────────
    if st.session_state.pcg_result_bytes:
        st.markdown('<div class="pcg-section-label">✨ Kết quả</div>', unsafe_allow_html=True)
        b64_result = base64.b64encode(st.session_state.pcg_result_bytes).decode()
        st.markdown(
            f"""
            <div class="pcg-result-frame">
                <img src="data:image/png;base64,{b64_result}">
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.download_button(
                "⬇️  Tải xuống PNG",
                data=st.session_state.pcg_result_bytes,
                file_name="anh-the.png",
                mime="image/png",
                use_container_width=True,
                key="pcg_download_btn",
            )
