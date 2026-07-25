"""Module Chấm công Face ID — đăng ký khuôn mặt & nhận diện check-in/out."""
import streamlit as st
import numpy as np
import cv2
import json
from datetime import datetime, date
from deepface import DeepFace
from streamlit_webrtc import webrtc_streamer, VideoProcessorBase
import av

FACE_MODEL_NAME = "Facenet"     # cân bằng tốc độ/độ chính xác cho server free tier
NGUONG_TUONG_DONG = 0.6          # ngưỡng cosine distance để coi là "cùng 1 người" (càng nhỏ càng khắt khe)


@st.cache_resource(show_spinner="Đang tải mô hình nhận diện khuôn mặt (chỉ chạy 1 lần)...")
def _load_model():
    """Ép DeepFace build model 1 lần, cache lại — tránh load lại mỗi lần predict gây chậm/tốn RAM."""
    DeepFace.build_model(FACE_MODEL_NAME)
    return True


def tinh_embedding(img_bgr):
    """Nhận ảnh dạng numpy array (BGR, từ cv2/webrtc), trả về vector embedding khuôn mặt
    hoặc None nếu không phát hiện được mặt rõ ràng trong khung hình."""
    _load_model()
    try:
        result = DeepFace.represent(
            img_path=img_bgr, model_name=FACE_MODEL_NAME,
            enforce_detection=True, detector_backend="opencv"
        )
        return np.array(result[0]["embedding"])
    except Exception:
        return None


def khoang_cach_cosine(vec1, vec2):
    vec1, vec2 = np.array(vec1), np.array(vec2)
    return 1 - np.dot(vec1, vec2) / (np.linalg.norm(vec1) * np.linalg.norm(vec2))


def dang_ky_khuon_mat(conn, nhan_vien_id, img_bgr):
    """Tính embedding từ ảnh và lưu/cập nhật vào bảng nhan_vien_face_id.
    Trả về (True, thông báo) hoặc (False, lý do lỗi)."""
    emb = tinh_embedding(img_bgr)
    if emb is None:
        return False, "Không phát hiện được khuôn mặt rõ ràng trong ảnh. Vui lòng chụp lại, đảm bảo đủ sáng và nhìn thẳng camera."

    try:
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO nhan_vien_face_id (nhan_vien_id, face_encoding, model_name, updated_at)
            VALUES (%s, %s, %s, now())
            ON CONFLICT (nhan_vien_id) DO UPDATE SET
                face_encoding = EXCLUDED.face_encoding,
                model_name = EXCLUDED.model_name,
                updated_at = now()
        """, (nhan_vien_id, json.dumps(emb.tolist()), FACE_MODEL_NAME))
        conn.commit()
        return True, "Đã đăng ký khuôn mặt thành công."
    except Exception as e:
        return False, f"Lỗi lưu dữ liệu: {e}"


def tai_toan_bo_embedding(conn):
    """Tải toàn bộ embedding đã đăng ký (dùng cho check-in) — chỉ nhân viên đang làm việc."""
    cur = conn.cursor()
    cur.execute("""
        SELECT f.nhan_vien_id, f.face_encoding, nv.ho_ten
        FROM nhan_vien_face_id f
        JOIN nhan_vien nv ON nv.id = f.nhan_vien_id
        WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
    """)
    rows = cur.fetchall()
    return [(r[0], np.array(json.loads(r[1]) if isinstance(r[1], str) else r[1]), r[2]) for r in rows]


def nhan_dien(emb, danh_sach_embedding):
    """So sánh embedding vừa chụp với toàn bộ embedding đã đăng ký.
    Trả về (nhan_vien_id, ho_ten, khoang_cach) của người khớp nhất, hoặc (None, None, None) nếu không ai khớp."""
    if not danh_sach_embedding:
        return None, None, None
    tot_nhat = min(danh_sach_embedding, key=lambda x: khoang_cach_cosine(emb, x[1]))
    kc = khoang_cach_cosine(emb, tot_nhat[1])
    if kc <= NGUONG_TUONG_DONG:
        return tot_nhat[0], tot_nhat[2], kc
    return None, None, kc


def ghi_nhan_cham_cong(conn, nhan_vien_id):
    """Ghi giờ vào (nếu chưa có) hoặc giờ ra (nếu đã có giờ vào) cho hôm nay.
    Trả về (loai, gio) với loai = 'VAO' | 'RA' | 'DA_DU_2_LUOT'."""
    hom_nay = date.today()
    gio_hien_tai = datetime.now().time()
    cur = conn.cursor()
    cur.execute("SELECT id, gio_vao, gio_ra FROM cham_cong WHERE nhan_vien_id = %s AND ngay = %s",
                (nhan_vien_id, hom_nay))
    row = cur.fetchone()

    if row is None:
        cur.execute("""
            INSERT INTO cham_cong (nhan_vien_id, ngay, gio_vao, ma_cong, nguon, created_at, updated_at)
            VALUES (%s, %s, %s, 'x', 'FACE_ID', now(), now())
        """, (nhan_vien_id, hom_nay, gio_hien_tai))
        conn.commit()
        return "VAO", gio_hien_tai

    cc_id, gio_vao, gio_ra = row
    if gio_vao and not gio_ra:
        cur.execute("UPDATE cham_cong SET gio_ra=%s, updated_at=now() WHERE id=%s",
                    (gio_hien_tai, cc_id))
        conn.commit()
        return "RA", gio_hien_tai

    return "DA_DU_2_LUOT", None


class FaceIDVideoProcessor(VideoProcessorBase):
    """Xử lý từng khung hình từ webrtc: chỉ nhận diện mỗi 2 giây/lần để tránh quá tải server
    (DeepFace khá nặng, không nên chạy trên MỌI khung hình video)."""
    def __init__(self):
        self.danh_sach_embedding = []
        self.conn = None
        self.ket_qua = {"trang_thai": "DANG_CHO", "ten": None, "thong_bao": None}
        self._lan_cuoi = 0

    def recv(self, frame):
        import time
        img = frame.to_ndarray(format="bgr24")
        now = time.time()
        if now - self._lan_cuoi >= 2:  # throttle: 2 giây/lần nhận diện
            self._lan_cuoi = now
            emb = tinh_embedding(img)
            if emb is not None:
                nv_id, ten, kc = nhan_dien(emb, self.danh_sach_embedding)
                if nv_id:
                    loai, gio = ghi_nhan_cham_cong(self.conn, nv_id)
                    if loai == "VAO":
                        self.ket_qua = {"trang_thai": "THANH_CONG", "ten": ten, "thong_bao": f"✅ {ten} — Check-IN lúc {gio.strftime('%H:%M:%S')}"}
                    elif loai == "RA":
                        self.ket_qua = {"trang_thai": "THANH_CONG", "ten": ten, "thong_bao": f"✅ {ten} — Check-OUT lúc {gio.strftime('%H:%M:%S')}"}
                    else:
                        self.ket_qua = {"trang_thai": "DA_DU", "ten": ten, "thong_bao": f"ℹ️ {ten} đã chấm công đủ 2 lượt hôm nay."}
                else:
                    self.ket_qua = {"trang_thai": "KHONG_KHOP", "ten": None, "thong_bao": "❌ Không nhận diện được — chưa đăng ký khuôn mặt hoặc chưa rõ mặt."}
        return av.VideoFrame.from_ndarray(img, format="bgr24")