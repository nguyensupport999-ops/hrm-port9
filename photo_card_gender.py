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
  - Chọn 1 mẫu target có sẵn trong thư viện (assets/targets/) — mỗi mẫu có
    PROMPT RIÊNG (xem khối TARGETS bên dưới), không dùng chung 1 prompt cho
    tất cả các mẫu như bản cũ.
  - Nền luôn mặc định là "Trắng tinh" (#FFFFFF) — không còn tùy chọn màu nền
    nào khác. Kể cả những ảnh target có nền gốc khác màu (vd nền xanh, nền
    olive), prompt riêng của mẫu đó vẫn LUÔN ép kết quả trả về nền trắng.
  - Nhấn "Tạo ảnh thẻ" -> gọi Google Gemini (nano banana) ghép mặt/tóc từ ảnh
    tham chiếu vào đúng dáng/trang phục của mẫu target đã chọn, dùng đúng
    prompt của riêng mẫu đó.
  - Nếu gọi API lỗi (vd sai/thiếu API key, hết quota...): thay vì chỉ báo lỗi
    suông, app sẽ hiện sẵn prompt (dạng có thể copy 1 click) + hướng dẫn từng
    bước để user tự thao tác thủ công trên webchat Gemini/ChatGPT.
  - Ảnh mẫu target hiển thị trong lưới chọn được LÀM MỜ KHUÔN MẶT tự động
    trước khi show lên UI (chỉ áp dụng cho ảnh hiển thị — ảnh gốc gửi cho
    Gemini API vẫn giữ nguyên, không bị mờ, để đảm bảo chất lượng ghép ảnh),
    nhằm tránh vấn đề bản quyền/quyền riêng tư khuôn mặt của người trong ảnh
    mẫu (thường là ảnh sưu tầm, không phải nhân viên thật của công ty).
  - Xem kết quả PNG trong khung cố định + nút tải xuống.

Lưu ý: module này KHÔNG gọi st.set_page_config (đã được gọi 1 lần ở app.py).
Toàn bộ session_state key đều có tiền tố "pcg_" để không đụng với các
màn hình khác trong app chính.
"""

import base64
import io
import mimetypes
import os
from pathlib import Path

import requests
import streamlit as st
from PIL import Image

# Thư viện làm mờ khuôn mặt — không bắt buộc phải cài, nếu thiếu app vẫn chạy
# bình thường (chỉ là ảnh mẫu sẽ không được làm mờ mặt trước khi hiển thị).
# Deploy trên Streamlit Cloud: nhớ thêm "opencv-python-headless" và "numpy"
# vào requirements.txt để tính năng làm mờ hoạt động.
try:
    import cv2
    import numpy as np
    _CV2_OK = True
except Exception:
    _CV2_OK = False


# ══════════════════════════════════════════════════════════════════════════
# ★★★ KHU VỰC ADMIN TÙY CHỈNH: DANH SÁCH MẪU TARGET + PROMPT RIÊNG ★★★
# ══════════════════════════════════════════════════════════════════════════
# Mỗi mẫu (target) trong danh sách TARGETS bên dưới gồm 3 phần:
#   - "file"  : tên file ảnh mẫu, đặt trong thư mục assets/targets/
#   - "label" : nhãn hiển thị dưới ảnh mẫu cho user chọn (Nam/Nữ · loại trang phục)
#   - "prompt": ĐOẠN PROMPT RIÊNG dùng để gọi Gemini khi user chọn đúng mẫu này.
#               Khi cần đổi/thêm/bớt mẫu, chỉ cần sửa trực tiếp trong khối này:
#                 1. Thêm file ảnh mới vào thư mục assets/targets/
#                 2. Thêm 1 dict mới vào TARGETS với "file"/"label"/"prompt" tương ứng
#               KHÔNG cần sửa bất kỳ chỗ nào khác trong file này.
#
# Cấu trúc chuẩn của 1 prompt (nên giữ nguyên khung này khi viết prompt mới):
#   - Đoạn mở đầu: mô tả ảnh 1 (reference) là mặt/tóc thật, ảnh 2 (target) là dáng/trang phục mẫu.
#   - "Keep unchanged": liệt kê CHÍNH XÁC những gì thuộc về ảnh target cần giữ
#     nguyên (loại trang phục, phụ kiện như kính/cà vạt, tư thế...).
#   - "Background": LUÔN LUÔN là nền trắng tinh (#FFFFFF) — KHÔNG được đổi
#     sang màu khác, kể cả khi ảnh target gốc có nền màu (xem ghi chú riêng ở
#     các mẫu t04, t11 — 2 mẫu này có ảnh gốc nền xanh/olive nhưng prompt vẫn
#     ép kết quả về nền trắng).
# ══════════════════════════════════════════════════════════════════════════

MODULE_DIR = Path(__file__).parent
TARGETS_DIR = MODULE_DIR / "assets" / "targets"

# Model image-generation của Google (Gemini "nano banana").
# Nếu Google đổi tên model, chỉ cần sửa dòng dưới đây.
GEMINI_MODEL = "gemini-2.5-flash-image"
GEMINI_ENDPOINT = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"

_WHITE_BG_BLOCK = (
    "Pure white studio background (#FFFFFF).\n"
    "Professional corporate ID portrait.\n"
    "Soft studio lighting.\n"
    "Natural shadows.\n"
    "True-to-life skin rendering.\n"
    "No artificial beauty filters.\n"
    "Photorealistic.\n"
    "Ultra high resolution.\n"
    "8K"
)


def _prompt(keep_unchanged: str, note: str = "") -> str:
    """Helper dựng prompt theo khung chuẩn — dùng khi admin thêm mẫu mới."""
    header = (
        "The image on the left is the 'reference image,' and the image on the right is the 'result image.'\n"
        "Use the following image generation prompt to transfer the head, face, and hair from the reference "
        "image to the 'result image.'\n"
        " \"Replace the complete head of the subject with the head from the provided reference image while "
        "preserving perfect identity consistency.\n"
        "The replacement must include:\n"
        "• forehead\n• hairline\n• hairstyle\n• ears\n• entire face\n• jawline\n• chin\n• neck transition\n"
        "The final portrait must immediately be recognizable as the same person from the reference: Preserve:\n"
        "• natural facial asymmetry\n• skin tone\n• facial proportions\n• expression\n• age\n• realistic pores\n\n"
        "Keep unchanged:\n"
    )
    tail = (
        "• body\n• posture\n• composition\n• framing\n• camera angle\n• perspective\n• shoulder orientation\n"
        f"Background:\n{_WHITE_BG_BLOCK}\"." + (f" {note}" if note else "")
    )
    return header + keep_unchanged + "\n" + tail


TARGETS = [
    {
        "file": "t01_nam_vest_xanhduong.png",
        "label": "Nam · Vest xanh dương",
        "prompt": _prompt(
            "• original clothing\n• business attire (dark blue suit, white shirt, navy tie, wristwatch)\n"
            "• crossed-arms pose"
        ),
    },
    {
        "file": "t02_nam_vest_soc_kinh.png",
        "label": "Nam · Vest sọc, kính",
        "prompt": _prompt(
            "• original clothing\n• business attire (pinstripe navy suit, white shirt, tie)\n"
            "• eyeglasses (frame shape, size and position unchanged)\n• crossed-arms pose",
            note="Do not remove or alter the eyeglasses.",
        ),
    },
    {
        "file": "t03_nu_vest_hong.png",
        "label": "Nữ · Vest hồng",
        "prompt": _prompt(
            "• original clothing\n• pink blazer over black top\n"
            "• loose long-hair framing/crop style of the target"
        ),
    },
    {
        "file": "t04_nam_aosomi_trang.png",
        "label": "Nam · Áo sơ mi trắng",
        "prompt": _prompt(
            "• original clothing\n• plain white collared shirt\n• front-facing head-on pose",
            note=(
                "Note: the ORIGINAL target background is a solid blue studio backdrop — DO NOT keep it; "
                "always replace it with the pure white background specified below, regardless of the "
                "target's original background color."
            ),
        ),
    },
    {
        "file": "t05_nu_vest_xam.png",
        "label": "Nữ · Vest xám",
        "prompt": _prompt(
            "• original clothing\n• grey blazer over white shirt\n• gentle smiling expression style of the target"
        ),
    },
    {
        "file": "t06_nu_aodai_trang.png",
        "label": "Nữ · Áo dài trắng",
        "prompt": _prompt(
            "• original clothing\n• white traditional Ao Dai with mandarin collar\n"
            "• short bob-hair framing/crop style of the target"
        ),
    },
    {
        "file": "t07_nu_vest_den_nho.png",
        "label": "Nữ · Vest đen (1)",
        "prompt": _prompt(
            "• original clothing\n• black blazer over white shirt\n"
            "• bright smiling expression style of the target\n• small ID-photo framing/crop"
        ),
    },
    {
        "file": "t08_nu_aosomi_trang2.png",
        "label": "Nữ · Áo sơ mi trắng (1)",
        "prompt": _prompt(
            "• original clothing\n• plain white collared shirt\n"
            "• long straight-hair framing/crop style of the target"
        ),
    },
    {
        "file": "t09_nam_vest_den_cavatdo.png",
        "label": "Nam · Vest đen, cà vạt sọc",
        "prompt": _prompt(
            "• original clothing\n• black suit, white shirt, red-and-white striped tie"
        ),
    },
    {
        "file": "t10_nam_aosomi_caro.png",
        "label": "Nam · Áo sơ mi caro",
        "prompt": _prompt(
            "• original clothing\n• light plaid/checked casual shirt (no blazer)\n"
            "• relaxed business-casual pose"
        ),
    },
    {
        "file": "t11_nam_vest_den_nenolive.png",
        "label": "Nam · Vest đen (2)",
        "prompt": _prompt(
            "• original clothing\n• black suit, white shirt, black tie",
            note=(
                "Note: the ORIGINAL target background is olive/moss green — DO NOT keep it; always replace "
                "it with the pure white background specified below."
            ),
        ),
    },
    {
        "file": "t12_nam_dongphuc1.png",
        "label": "Nam · Đồng phục (1)",
        "prompt": _prompt(
            "• original clothing\n• company uniform polo/shirt with contrast blue-and-beige "
            "color-block collar and shoulder panels"
        ),
    },
    {
        "file": "t13_nam_vest_kem.png",
        "label": "Nam · Vest kem",
        "prompt": _prompt(
            "• original clothing\n• light cream/beige casual blazer over white crew-neck top, no tie\n"
            "• modern youthful styling"
        ),
    },
    {
        "file": "t14_nu_aosomi_trang3.png",
        "label": "Nữ · Áo sơ mi trắng (2)",
        "prompt": _prompt(
            "• original clothing\n• plain white collared shirt\n"
            "• long straight-hair framing/crop style of the target, center parting"
        ),
    },
    {
        "file": "t15_nu_vest_den_cotrang.png",
        "label": "Nữ · Vest đen, cổ trắng",
        "prompt": _prompt(
            "• original clothing\n• black blazer with white contrast collar\n"
            "• bright smiling expression style of the target"
        ),
    },
    {
        "file": "t16_nam_dongphuc2.png",
        "label": "Nam · Đồng phục (2)",
        "prompt": _prompt(
            "• original clothing\n• company uniform shirt with contrast blue shoulder/side panel piping\n"
            "• straight-on front-facing pose"
        ),
    },
]

_CSS_INJECTED_KEY = "pcg_css_injected"


# ──────────────────────────────────────────────────────────────────────────
# CSS (chỉ inject 1 lần / phiên)
# ──────────────────────────────────────────────────────────────────────────
def _inject_css():
    # Streamlit không giữ lại các phần tử đã render ở lượt chạy trước — mỗi lần
    # script chạy lại (vd sau khi upload ảnh, tick chọn mẫu...) toàn bộ DOM cũ
    # bị thay thế. Vì vậy KHÔNG được chỉ inject CSS 1 lần duy nhất/phiên (nếu
    # không, ở các lượt rerun sau, thẻ <style> biến mất -> layout lưới ảnh mẫu
    # bị vỡ). Việc st.markdown lặp lại cùng 1 khối CSS mỗi lượt chạy là bình
    # thường và không tốn kém.
    st.markdown(
        """
        <style>
        .pcg-title{ font-size:26px; font-weight:800; margin-bottom:2px; }
        .pcg-sub{ font-size:14px; color:#6B6B8A; margin-bottom:18px; }
        .pcg-section-label{
            font-size:14px; font-weight:700; margin:18px 0 10px 0;
            display:flex; align-items:center; gap:6px;
        }
        /* Thẻ mẫu target: tỷ lệ khung 4:6 (width:height), ảnh full-frame (object-fit:cover
           lấp đầy toàn bộ khung, không viền trắng thừa), không còn label tên ảnh bên dưới. */
        .pcg-target-card{
            position:relative; border-radius:12px; overflow:hidden;
            border:2.5px solid transparent; box-shadow:0 4px 16px rgba(108,99,255,.10);
            background:#fff; margin-bottom:6px; aspect-ratio:4/6; width:100%;
        }
        .pcg-target-card.selected{
            border-color:#6C63FF;
            box-shadow:0 0 0 4px rgba(108,99,255,.18),0 8px 28px rgba(108,99,255,.22);
        }
        .pcg-target-card img{ width:100%; height:100%; object-fit:cover; display:block; }
        /* Hàng dưới ảnh chỉ còn tickbox "Chọn" — canh giữa, gọn */
        div[data-testid="stCheckbox"].pcg-tick{ display:flex; justify-content:center; margin-top:2px; }
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


_FACE_CASCADE = None


def _get_face_cascade():
    global _FACE_CASCADE
    if _FACE_CASCADE is None and _CV2_OK:
        cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        _FACE_CASCADE = cv2.CascadeClassifier(cascade_path)
    return _FACE_CASCADE


@st.cache_data(show_spinner=False)
def _blurred_thumb_b64(file_key: str, raw_bytes: bytes) -> str:
    """
    Trả về ảnh (base64) đã làm mờ vùng khuôn mặt — CHỈ dùng để hiển thị trong
    lưới chọn mẫu, để tránh vấn đề bản quyền/quyền riêng tư của người trong
    ảnh mẫu sưu tầm. Ảnh gốc (không mờ) vẫn được dùng khi gọi Gemini API.
    Có cache theo file_key để không phải xử lý lại mỗi lần rerun.
    """
    mime = mimetypes.guess_type(file_key)[0] or "image/jpeg"
    if not _CV2_OK:
        return f"data:{mime};base64,{base64.b64encode(raw_bytes).decode()}"

    try:
        img = Image.open(io.BytesIO(raw_bytes)).convert("RGB")
        arr = np.array(img)
        gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
        cascade = _get_face_cascade()
        faces = cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(40, 40)) if cascade else []

        for (x, y, w, h) in faces:
            # Mở rộng vùng làm mờ ra thêm quanh khung mặt để che trọn (trán, cằm, tai)
            pad_x, pad_y = int(w * 0.30), int(h * 0.35)
            x0, y0 = max(0, x - pad_x), max(0, y - pad_y)
            x1, y1 = min(arr.shape[1], x + w + pad_x), min(arr.shape[0], y + h + pad_y)
            roi = arr[y0:y1, x0:x1]
            if roi.size == 0:
                continue
            k = max(21, (min(roi.shape[:2]) // 2) | 1)  # kernel lẻ, đủ mạnh để mờ hẳn
            arr[y0:y1, x0:x1] = cv2.GaussianBlur(roi, (k, k), 0)

        out_img = Image.fromarray(arr)
        out = io.BytesIO()
        out_img.save(out, format="PNG")
        return f"data:image/png;base64,{base64.b64encode(out.getvalue()).decode()}"
    except Exception:
        # Nếu vì lý do gì đó xử lý lỗi, vẫn hiển thị ảnh gốc thay vì làm sập UI
        return f"data:{mime};base64,{base64.b64encode(raw_bytes).decode()}"


def _call_gemini_image_edit(api_key, ref_bytes, ref_mime, target_bytes, target_mime, prompt_text) -> bytes:
    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"inline_data": {"mime_type": ref_mime, "data": base64.b64encode(ref_bytes).decode()}},
                    {"inline_data": {"mime_type": target_mime, "data": base64.b64encode(target_bytes).decode()}},
                    {"text": prompt_text},
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


def _render_manual_fallback(selected_item: dict):
    """
    Hiển thị hướng dẫn thao tác thủ công + prompt có thể copy 1 click, dùng khi
    gọi Gemini API thất bại (vd sai/thiếu API key, hết quota, lỗi mạng...).
    """
    st.warning(
        "API key chưa đúng, bạn có thể thực hiện trực tiếp trên webchat của "
        "Gemini/ChatGPT theo các bước sau:\n\n"
        "1️⃣ Tải lần 1 là ảnh có khuôn mặt của bạn (ảnh tham chiếu).\n\n"
        "2️⃣ Tải lần 2 là ảnh Target bạn đã chọn.\n\n"
        "3️⃣ Copy đoạn mẫu Prompt bên dưới, dán vào khung chat.\n\n"
        "4️⃣ Nhấn Enter và chờ kết quả nhé."
    )
    st.caption(f"📋 Prompt dùng cho mẫu **{selected_item['label']}** (bấm biểu tượng copy ở góc khối bên dưới):")
    st.code(selected_item["prompt"], language=None)
    col_a, col_b = st.columns(2)
    with col_a:
        st.link_button("🌐 Mở Gemini (webchat)", "https://gemini.google.com/app", use_container_width=True)
    with col_b:
        st.link_button("🌐 Mở ChatGPT (webchat)", "https://chatgpt.com/", use_container_width=True)


def _on_target_checkbox_change(file_key: str):
    """
    Callback cho checkbox "Chọn" của từng thẻ mẫu target.
    Streamlit chạy callback này TRƯỚC khi thân script được render lại ở lượt
    kế tiếp, nên tại đây có thể chỉnh sửa st.session_state của các checkbox
    KHÁC một cách an toàn (chưa bị instantiate trong lượt chạy hiện tại).
    Nếu làm việc này trực tiếp trong thân hàm render() (sau khi các checkbox
    khác đã được vẽ) sẽ bị StreamlitAPIException vì widget đã instantiate.
    """
    is_now_checked = st.session_state.get(f"pcg_chk_{file_key}", False)
    if is_now_checked:
        # User vừa tick chọn mẫu này -> bỏ tick tất cả mẫu khác (chỉ chọn 1)
        st.session_state.pcg_selected_target = file_key
        st.session_state.pcg_result_bytes = None
        st.session_state.pcg_show_manual_fallback = False
        for other in TARGETS:
            if other["file"] != file_key:
                st.session_state[f"pcg_chk_{other['file']}"] = False
    else:
        # User vừa bỏ tick mẫu đang chọn
        if st.session_state.get("pcg_selected_target") == file_key:
            st.session_state.pcg_selected_target = None


# ──────────────────────────────────────────────────────────────────────────
# HÀM RENDER CHÍNH — gọi từ app.py: photo_card_gender.render()
# ──────────────────────────────────────────────────────────────────────────
def render():
    _inject_css()

    st.markdown('<div class="pcg-title">🖼️ Tạo ảnh thẻ nhân viên bằng AI</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="pcg-sub">Upload ảnh chân dung, chọn mẫu có sẵn, AI sẽ ghép mặt/tóc vào đúng trang phục '
        '(nền luôn là trắng tinh).</div>',
        unsafe_allow_html=True,
    )

    # ── SESSION STATE (tiền tố pcg_) ────────────────────────────────────
    st.session_state.setdefault("pcg_selected_target", None)
    st.session_state.setdefault("pcg_result_bytes", None)
    st.session_state.setdefault("pcg_error_msg", None)
    st.session_state.setdefault("pcg_show_manual_fallback", False)

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
                if not target_path.exists():
                    st.error(f"❌ Không tìm thấy file ảnh: {target_path}")
                    st.info("💡 Vui lòng kiểm tra lại thư mục 'assets/targets' hoặc đường dẫn file ảnh.")
                    continue

                with open(target_path, "rb") as f:
                    raw_bytes = f.read()
                # Ảnh hiển thị lên UI đã được làm mờ khuôn mặt (bảo vệ bản quyền/quyền
                # riêng tư) — ảnh gốc raw_bytes chỉ dùng khi thực sự gọi Gemini API.
                thumb_src = _blurred_thumb_b64(item["file"], raw_bytes)

                css_class = "pcg-target-card selected" if is_selected else "pcg-target-card"
                st.markdown(
                    f"""
                    <div class="{css_class}">
                        <img src="{thumb_src}">
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown('<div class="pcg-tick">', unsafe_allow_html=True)
                st.checkbox(
                    "Chọn",
                    value=is_selected,
                    key=f"pcg_chk_{item['file']}",
                    label_visibility="visible",
                    on_change=_on_target_checkbox_change,
                    args=(item["file"],),
                )
                st.markdown("</div>", unsafe_allow_html=True)

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

    selected_item = next((t for t in TARGETS if t["file"] == st.session_state.pcg_selected_target), None)

    if render_clicked and can_render:
        st.session_state.pcg_error_msg = None
        st.session_state.pcg_result_bytes = None
        st.session_state.pcg_show_manual_fallback = False
        if not api_key:
            st.session_state.pcg_error_msg = "Vui lòng nhập Google Gemini API Key trước."
            st.session_state.pcg_show_manual_fallback = True
        else:
            try:
                with st.spinner("Đang xử lý ảnh, vui lòng chờ trong giây lát..."):
                    ref_raw = ref_file.getvalue()
                    ref_png = _to_png_bytes(ref_raw)

                    target_path = TARGETS_DIR / selected_item["file"]
                    with open(target_path, "rb") as f:
                        target_raw = f.read()
                    target_mime = mimetypes.guess_type(str(target_path))[0] or "image/jpeg"

                    result_bytes = _call_gemini_image_edit(
                        api_key=api_key,
                        ref_bytes=ref_png,
                        ref_mime="image/png",
                        target_bytes=target_raw,
                        target_mime=target_mime,
                        prompt_text=selected_item["prompt"],
                    )
                    st.session_state.pcg_result_bytes = _to_png_bytes(result_bytes)
            except Exception as e:
                st.session_state.pcg_error_msg = f"Lỗi: {e}"
                st.session_state.pcg_show_manual_fallback = True

    if st.session_state.pcg_error_msg:
        st.error(st.session_state.pcg_error_msg)

    if st.session_state.pcg_show_manual_fallback and selected_item:
        _render_manual_fallback(selected_item)

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