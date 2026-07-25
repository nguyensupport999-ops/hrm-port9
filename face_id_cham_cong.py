"""Module Chấm công Face ID — dùng nhận diện khuôn mặt tích hợp sẵn trong OpenCV
(YuNet phát hiện mặt + SFace nhận diện danh tính). KHÔNG cần TensorFlow/DeepFace,
nên chạy được trên Python 3.14 của Streamlit Cloud."""
import os
import json
import time
import threading
import urllib.request
import numpy as np
import cv2
from datetime import datetime, date
from streamlit_webrtc import VideoProcessorBase
import av

# Ngưỡng cosine similarity: >= ngưỡng thì coi là CÙNG 1 người.
# 0.363 là ngưỡng khuyến nghị của OpenCV cho model SFace. Tăng lên (VD 0.40) để khắt khe hơn.
NGUONG_TUONG_DONG = 0.363

_THU_MUC_MODEL = os.path.expanduser("~/.cache/hrm_face_models")

# Nhiều URL dự phòng: link Git LFS của GitHub, sau đó tới HuggingFace.
_MODEL_INFO = {
    "yunet": {
        "ten_file": "face_detection_yunet_2023mar.onnx",
        "kich_thuoc_toi_thieu": 100_000,
        "urls": [
            "https://media.githubusercontent.com/media/opencv/opencv_zoo/main/models/face_detection_yunet/face_detection_yunet_2023mar.onnx",
            "https://huggingface.co/opencv/face_detection_yunet/resolve/main/face_detection_yunet_2023mar.onnx",
        ],
    },
    "sface": {
        "ten_file": "face_recognition_sface_2021dec.onnx",
        "kich_thuoc_toi_thieu": 30_000_000,
        "urls": [
            "https://media.githubusercontent.com/media/opencv/opencv_zoo/main/models/face_recognition_sface/face_recognition_sface_2021dec.onnx",
            "https://huggingface.co/opencv/face_recognition_sface/resolve/main/face_recognition_sface_2021dec.onnx",
        ],
    },
}

_detector = None
_recognizer = None
_lock = threading.Lock()


def _tai_file_model(khoa):
    """Tải 1 file model về cache nếu chưa có. Trả về đường dẫn file.
    Tự kiểm tra file tải về là model thật (không phải con trỏ Git LFS)."""
    info = _MODEL_INFO[khoa]
    os.makedirs(_THU_MUC_MODEL, exist_ok=True)
    duong_dan = os.path.join(_THU_MUC_MODEL, info["ten_file"])

    if os.path.exists(duong_dan) and os.path.getsize(duong_dan) >= info["kich_thuoc_toi_thieu"]:
        return duong_dan

    loi_cuoi = None
    for url in info["urls"]:
        try:
            tam = duong_dan + ".tmp"
            urllib.request.urlretrieve(url, tam)
            if os.path.getsize(tam) >= info["kich_thuoc_toi_thieu"]:
                os.replace(tam, duong_dan)
                return duong_dan
            os.remove(tam)
            loi_cuoi = f"File tải về từ {url} không hợp lệ (quá nhỏ)."
        except Exception as e:
            loi_cuoi = f"Lỗi tải từ {url}: {e}"

    raise RuntimeError(f"Không tải được model '{info['ten_file']}'. {loi_cuoi}")


def chuan_bi_model():
    """Tải model (lần đầu) và khởi tạo bộ nhận diện. Gọi 1 lần từ luồng chính của
    Streamlit TRƯỚC khi bật camera — không gọi trong luồng xử lý video."""
    global _detector, _recognizer
    with _lock:
        if _detector is not None and _recognizer is not None:
            return True
        duong_dan_yunet = _tai_file_model("yunet")
        duong_dan_sface = _tai_file_model("sface")
        _detector = cv2.FaceDetectorYN.create(duong_dan_yunet, "", (320, 320), 0.9, 0.3, 5000)
        _recognizer = cv2.FaceRecognizerSF.create(duong_dan_sface, "")
        return True


def tinh_embedding(img_bgr):
    """Nhận ảnh numpy (BGR), trả về vector đặc trưng khuôn mặt (128 chiều),
    hoặc None nếu không thấy khuôn mặt trong ảnh."""
    if _detector is None or _recognizer is None:
        chuan_bi_model()
    try:
        cao, rong = img_bgr.shape[:2]
        with _lock:
            _detector.setInputSize((rong, cao))
            _, cac_mat = _detector.detect(img_bgr)
            if cac_mat is None or len(cac_mat) == 0:
                return None
            # Nếu có nhiều mặt trong khung hình, chọn mặt to nhất (người đứng gần camera nhất)
            mat = max(cac_mat, key=lambda f: f[2] * f[3])
            anh_can_chinh = _recognizer.alignCrop(img_bgr, mat)
            dac_trung = _recognizer.feature(anh_can_chinh)
        return np.array(dac_trung).flatten()
    except Exception:
        return None


def do_tuong_dong(vec1, vec2):
    """Cosine similarity: càng gần 1 càng giống nhau."""
    vec1, vec2 = np.array(vec1, dtype=np.float64), np.array(vec2, dtype=np.float64)
    return float(np.dot(vec1, vec2) / (np.linalg.norm(vec1) * np.linalg.norm(vec2)))


def dang_ky_khuon_mat(conn, nhan_vien_id, img_bgr):
    """Tính vector đặc trưng từ ảnh và lưu/cập nhật vào bảng nhan_vien_face_id.
    Trả về (True, thông báo) hoặc (False, lý do lỗi)."""
    emb = tinh_embedding(img_bgr)
    if emb is None:
        return False, ("❌ Không phát hiện được khuôn mặt trong ảnh. "
                       "Chụp lại: nhìn thẳng camera, đủ sáng, mặt chiếm phần lớn khung hình.")
    try:
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO nhan_vien_face_id (nhan_vien_id, face_encoding, model_name, updated_at)
            VALUES (%s, %s, %s, NOW())
            ON CONFLICT (nhan_vien_id) DO UPDATE SET
                face_encoding = EXCLUDED.face_encoding,
                model_name = EXCLUDED.model_name,
                updated_at = NOW()
        """, (nhan_vien_id, json.dumps(emb.tolist()), "SFace"))
        conn.commit()
        cur.close()
        return True, "✅ Đã đăng ký khuôn mặt thành công."
    except Exception as e:
        return False, f"❌ Lỗi lưu dữ liệu: {e}"


def tai_toan_bo_embedding(conn):
    """Tải toàn bộ khuôn mặt đã đăng ký của nhân viên đang làm việc/thử việc."""
    cur = conn.cursor()
    cur.execute("""
        SELECT f.nhan_vien_id, f.face_encoding, nv.ho_ten
        FROM nhan_vien_face_id f
        JOIN nhan_vien nv ON nv.id = f.nhan_vien_id
        WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
    """)
    ket_qua = []
    for r in cur.fetchall():
        enc = json.loads(r[1]) if isinstance(r[1], str) else r[1]
        ket_qua.append((r[0], np.array(enc), r[2]))
    cur.close()
    return ket_qua


def nhan_dien(emb, danh_sach_embedding):
    """So khớp với toàn bộ khuôn mặt đã đăng ký.
    Trả về (nhan_vien_id, ho_ten, do_giong) nếu khớp, ngược lại (None, None, do_giong_cao_nhat)."""
    if not danh_sach_embedding:
        return None, None, 0.0
    tot_nhat = max(danh_sach_embedding, key=lambda x: do_tuong_dong(emb, x[1]))
    diem = do_tuong_dong(emb, tot_nhat[1])
    if diem >= NGUONG_TUONG_DONG:
        return tot_nhat[0], tot_nhat[2], diem
    return None, None, diem


def ghi_nhan_cham_cong(conn, nhan_vien_id):
    """Cách B: lần quét đầu tiên trong ngày = giờ VÀO (sớm nhất).
    Mọi lần quét sau (cách ít nhất 60 giây, kiểm tra ở FaceIDVideoProcessor)
    = cập nhật giờ RA muộn nhất. Không bao giờ từ chối ghi khi còn trong ngày.
    Trả về (loai, gio) với loai = 'VAO' | 'RA'."""
    hom_nay = date.today()
    gio_hien_tai = datetime.now().time().replace(microsecond=0)
    cur = conn.cursor()
    cur.execute(
        "SELECT id, gio_vao, gio_ra FROM cham_cong WHERE nhan_vien_id = %s AND ngay = %s",
        (nhan_vien_id, hom_nay)
    )
    row = cur.fetchone()

    if row is None:
        # Lần đầu trong ngày → ghi giờ vào
        cur.execute("""
            INSERT INTO cham_cong
                (nhan_vien_id, ngay, gio_vao, ma_cong, nguon, created_by, created_at, updated_at)
            VALUES (%s, %s, %s, 'x', 'FACE_ID', 'FACE_ID', NOW(), NOW())
        """, (nhan_vien_id, hom_nay, gio_hien_tai))
        conn.commit()
        cur.close()
        return "VAO", gio_hien_tai

    cc_id, gio_vao, gio_ra = row

    if not gio_vao:
        # Có dòng nhưng chưa ghi gio_vao (do chấm thủ công tạo dòng trước)
        cur.execute(
            "UPDATE cham_cong SET gio_vao=%s, nguon='FACE_ID', updated_at=NOW() WHERE id=%s",
            (gio_hien_tai, cc_id)
        )
        conn.commit()
        cur.close()
        return "VAO", gio_hien_tai

    # Đã có gio_vao → luôn cập nhật gio_ra muộn nhất (không từ chối)
    cur.execute(
        "UPDATE cham_cong SET gio_ra=%s, nguon='FACE_ID', updated_at=NOW() WHERE id=%s",
        (gio_hien_tai, cc_id)
    )
    conn.commit()
    cur.close()
    return "RA", gio_hien_tai


class FaceIDVideoProcessor(VideoProcessorBase):
    """Xử lý khung hình từ camera. Chỉ nhận diện mỗi 2 giây/lần để không quá tải server,
    và mỗi nhân viên chỉ ghi chấm công 1 lần trong 60 giây (tránh ghi trùng liên tục)."""

    def __init__(self):
        self.danh_sach_embedding = []
        self.conn = None
        self.ket_qua = {"trang_thai": "DANG_CHO", "ten": None, "thong_bao": None}
        self._lan_cuoi = 0.0
        self._da_ghi = {}  # {nhan_vien_id: thời điểm ghi gần nhất}

    def recv(self, frame):
        img = frame.to_ndarray(format="bgr24")
        bay_gio = time.time()

        if bay_gio - self._lan_cuoi >= 2 and self.conn is not None:
            self._lan_cuoi = bay_gio
            emb = tinh_embedding(img)
            if emb is None:
                self.ket_qua = {"trang_thai": "DANG_CHO", "ten": None,
                                "thong_bao": "📷 Chưa thấy khuôn mặt — đưa mặt vào khung hình."}
            else:
                nv_id, ten, diem = nhan_dien(emb, self.danh_sach_embedding)
                if nv_id is None:
                    self.ket_qua = {"trang_thai": "KHONG_KHOP", "ten": None,
                                    "thong_bao": f"❌ Không nhận diện được (độ giống {diem:.2f}) — "
                                                 "có thể chưa đăng ký khuôn mặt."}
                elif bay_gio - self._da_ghi.get(nv_id, 0) < 60:
                    self.ket_qua = {"trang_thai": "DA_DU", "ten": ten,
                                    "thong_bao": f"ℹ️ {ten} — vừa chấm công xong, vui lòng rời khỏi camera."}
                else:
                    try:
                        loai, gio = ghi_nhan_cham_cong(self.conn, nv_id)
                        self._da_ghi[nv_id] = bay_gio
                        if loai == "VAO":
                            self.ket_qua = {"trang_thai": "THANH_CONG", "ten": ten,
                                            "thong_bao": f"✅ {ten} — Vào lúc {gio.strftime('%H:%M:%S')}"}
                        else:
                            self.ket_qua = {"trang_thai": "THANH_CONG", "ten": ten,
                                            "thong_bao": f"✅ {ten} — Cập nhật giờ ra: {gio.strftime('%H:%M:%S')}"}
                    except Exception as e:
                        self.ket_qua = {"trang_thai": "KHONG_KHOP", "ten": None,
                                        "thong_bao": f"❌ Lỗi ghi chấm công: {e}"}

        return av.VideoFrame.from_ndarray(img, format="bgr24")