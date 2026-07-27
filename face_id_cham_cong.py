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
from datetime import datetime, date, timezone, timedelta
_TZ_VN = timezone(timedelta(hours=7))  # GMT+7 Việt Nam
from streamlit_webrtc import VideoProcessorBase
import av
import chat_noi_bo

_lock = threading.Lock()
_detector = None
_recognizer = None
NGUONG_TUONG_DONG = 0.36  # Độ giống cosine tối thiểu để coi là cùng 1 người (SFace khuyến nghị ~0.36)

_THU_MUC_MODEL = os.path.join(os.path.dirname(__file__), "models_cache")

_MODEL_INFO = {
    "yunet": {
        "ten_file": "face_detection_yunet_2023mar.onnx",
        "kich_thuoc_toi_thieu": 100_000,       # file thật ~228KB, đặt ngưỡng an toàn thấp hơn
        "urls": [
            "https://github.com/opencv/opencv_zoo/raw/main/models/face_detection_yunet/face_detection_yunet_2023mar.onnx",
        ],
    },
    "sface": {
        "ten_file": "face_recognition_sface_2021dec.onnx",
        "kich_thuoc_toi_thieu": 5_000_000,     # file thật ~37MB, đặt ngưỡng an toàn thấp hơn
        "urls": [
            "https://github.com/opencv/opencv_zoo/raw/main/models/face_recognition_sface/face_recognition_sface_2021dec.onnx",
        ],
    },
}


def suy_ma_cong_tu_gio_vao_ra(gio_vao, gio_ra, cfg):
    """
    Diem 2: Suy ma_cong tu dong tu gio vao/ra ghi nhan qua Face ID.
    cfg = get_cau_hinh_cham_cong_full() - dung key: gio_vao, gio_ra (gio chuan), phut_tre.

    Logic di tre:
    - Tre <= phut_tre (mac dinh 15p): 'x' binh thuong, ghi so_phut_di_tre
    - Tre > phut_tre va <= 60p: 'x' + canh bao admin_bcc/truong phong
    - Tre > 60p: ma_cong = None (cho duyet), trang_thai = CHO_DUYET_GIO
    """
    if gio_ra is None:
        return {
            'ma_cong': None,
            'trang_thai_cham_cong': 'THIEU_GIO_RA',
            'so_phut_di_tre': None,
            'so_phut_ve_som': None,
            'canh_bao': None,
        }

    def _phut_tu_nua_dem(t):
        return t.hour * 60 + t.minute

    phut_tre_cfg = int(cfg.get('phut_tre', 15))
    so_phut_di_tre = max(0, _phut_tu_nua_dem(gio_vao) - _phut_tu_nua_dem(cfg['gio_vao']))
    so_phut_ve_som = max(0, _phut_tu_nua_dem(cfg['gio_ra']) - _phut_tu_nua_dem(gio_ra))

    canh_bao = None

    if so_phut_di_tre > 60:
        # Trễ > 60 phút → không tự gán x, chờ admin_bcc duyệt
        return {
            'ma_cong': None,
            'trang_thai_cham_cong': 'CHO_DUYET_GIO',
            'so_phut_di_tre': so_phut_di_tre,
            'so_phut_ve_som': so_phut_ve_som,
            'canh_bao': f'DI_TRE_{so_phut_di_tre}P_CHO_DUYET',
        }
    elif so_phut_di_tre > phut_tre_cfg:
        # Trễ > ngưỡng cho phép (15p) nhưng <= 60p → vẫn gán x + cảnh báo
        canh_bao = f'DI_TRE_{so_phut_di_tre}P_CANH_BAO'

    return {
        'ma_cong': 'x',
        'trang_thai_cham_cong': 'HOP_LE',
        'so_phut_di_tre': so_phut_di_tre,
        'so_phut_ve_som': so_phut_ve_som,
        'canh_bao': canh_bao,
    }


def _cho_phep_tang_ca_phong_ban(conn, ten_phong_ban):
    """Kiem tra phong ban co duoc phep tang ca khong (mac dinh True neu chua
    co cau hinh rieng). Chi lay dung 1 truong can cho buoc cham cong - phan
    he so luong day du thuoc ve Payroll Engine (get_cau_hinh_tang_ca_theo_phong
    trong app.py), khong lap lai o day."""
    if not ten_phong_ban:
        return True
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT cho_phep_tang_ca FROM cau_hinh_tang_ca_phong_ban WHERE ten_phong_ban = %s",
            (ten_phong_ban,)
        )
        row = cur.fetchone()
        cur.close()
        return True if row is None else bool(row[0])
    except Exception:
        return True


def suy_gio_tang_ca(ngay, gio_vao, gio_ra, cfg):
    """
    Diem 3: Suy so gio tang ca tu dong khi lam qua gio ra ca chuan.
    cfg = get_cau_hinh_cham_cong_full() - dung: gio_ra, gio_bat_dau_ca_dem,
    danh_sach_ngay_le, phut_toi_thieu_tang_ca.
    - CHI tinh phan gio VUOT SAU gio ra ca chuan (den som truoc ca KHONG tinh TC).
    - Duoi nguong phut_toi_thieu_tang_ca -> khong tinh la co tang ca.
    - Xac dinh loai_ngay_tang_ca: LE (uu tien cao nhat) > CHU_NHAT > THUONG.
    - Tach rieng phan gio roi vao khung gio dem (tu gio_bat_dau_ca_dem tro di).
    - Luon tra ve trang_thai_duyet_tc = 'CHO_DUYET' neu co gio TC - CHUA tu duyet,
      cho Truong phong xac nhan truoc khi dua vao luong.
    """
    rong = {'tong_gio_tang_ca': 0.0, 'gio_tang_ca_dem': 0.0,
            'loai_ngay_tang_ca': None, 'trang_thai_duyet_tc': None}
    if not gio_vao or not gio_ra:
        return rong

    dt_ra_chuan = datetime.combine(ngay, cfg['gio_ra'])
    dt_ra_thuc_te = datetime.combine(ngay, gio_ra)
    dt_vao_thuc_te = datetime.combine(ngay, gio_vao)
    if dt_ra_thuc_te <= dt_vao_thuc_te:
        dt_ra_thuc_te += timedelta(days=1)  # ca qua đêm
    if dt_ra_thuc_te <= dt_ra_chuan:
        return rong  # chưa vượt giờ ra chuẩn -> không có tăng ca

    so_phut_vuot = (dt_ra_thuc_te - dt_ra_chuan).total_seconds() / 60
    if so_phut_vuot < cfg.get('phut_toi_thieu_tang_ca', 30):
        return rong

    tong_gio_tang_ca = round(so_phut_vuot / 60, 2)

    # Xác định loại ngày: Lễ ưu tiên cao nhất, sau đó Chủ nhật, còn lại Thường
    ngay_str = ngay.strftime('%Y-%m-%d')
    danh_sach_le = {x['ngay'] for x in (cfg.get('danh_sach_ngay_le') or [])}
    if ngay_str in danh_sach_le:
        loai_ngay = 'LE'
    elif ngay.weekday() == 6:  # Chủ nhật
        loai_ngay = 'CHU_NHAT'
    else:
        loai_ngay = 'THUONG'

    # Tách phần giờ rơi vào khung đêm (từ gio_bat_dau_ca_dem trở đi)
    dt_bat_dau_dem = datetime.combine(ngay, cfg['gio_bat_dau_ca_dem'])
    if dt_bat_dau_dem < dt_ra_chuan:
        dt_bat_dau_dem += timedelta(days=1)
    gio_tang_ca_dem = 0.0
    if dt_ra_thuc_te > dt_bat_dau_dem:
        moc_bat_dau_dem = max(dt_ra_chuan, dt_bat_dau_dem)
        gio_tang_ca_dem = round((dt_ra_thuc_te - moc_bat_dau_dem).total_seconds() / 3600, 2)

    return {
        'tong_gio_tang_ca': tong_gio_tang_ca,
        'gio_tang_ca_dem': gio_tang_ca_dem,
        'loai_ngay_tang_ca': loai_ngay,
        'trang_thai_duyet_tc': 'CHO_DUYET',
    }


def quet_va_canh_bao_thieu_gio_ra(conn, ngay_hom_nay):
    """
    Diem 2 - Lazy check: goi 1 lan moi khi co nguoi mo app (ngay sau khi login
    thanh cong, trong app.py). Quet cac ngay TRUOC ngay_hom_nay con o trang thai
    THIEU_GIO_RA va chua gui canh bao -> gui tin nhan qua Chat noi bo + danh dau da gui.
    """
    ID_HE_THONG = 0  # Quy ước có sẵn trong code gốc: id=0 = "Hệ thống"
                      # (giống chat_rooms.created_by=0 dùng cho phòng "Thông báo chung")

    cur = conn.cursor()
    cur.execute("""
        SELECT cc.id, cc.nhan_vien_id, cc.ngay
        FROM cham_cong cc
        WHERE cc.trang_thai_cham_cong = 'THIEU_GIO_RA'
          AND cc.ngay < %s
          AND cc.da_gui_canh_bao_thieu_gio_ra = false
    """, (ngay_hom_nay,))
    danh_sach_thieu = cur.fetchall()

    for cham_cong_id, nhan_vien_id, ngay_thieu in danh_sach_thieu:
        noi_dung = (
            f"⚠️ Bạn quên chấm công RA ngày {ngay_thieu.strftime('%d/%m/%Y')}.\n"
            f"Vui lòng gửi yêu cầu điều chỉnh chấm công qua mục Duyệt yêu cầu "
            f"để HR/Trưởng phòng xác nhận lại giờ ra, nếu không ngày công này "
            f"sẽ chưa được tính vào bảng chấm công tháng."
        )
        room_id = chat_noi_bo.create_private_room(ID_HE_THONG, nhan_vien_id)
        if room_id:
            chat_noi_bo.send_message(room_id, ID_HE_THONG, noi_dung, message_type="system")

        cur.execute("""
            UPDATE cham_cong SET da_gui_canh_bao_thieu_gio_ra = true WHERE id = %s
        """, (cham_cong_id,))

    conn.commit()
    cur.close()
    return len(danh_sach_thieu)
    
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


def ghi_nhan_cham_cong(conn, nhan_vien_id, cfg):
    """Cách B: lần quét đầu tiên trong ngày = giờ VÀO (sớm nhất).
    Mọi lần quét sau (cách ít nhất 60 giây, kiểm tra ở FaceIDVideoProcessor)
    = cập nhật giờ RA muộn nhất. Không bao giờ từ chối ghi khi còn trong ngày.
    cfg = get_cau_hinh_cham_cong_full() (app.py), truyền vào từ
    FaceIDVideoProcessor.cfg để suy ma_cong ngay lúc ghi giờ ra (Điểm 2).
    Trả về (loai, gio) với loai = 'VAO' | 'RA'."""
    bay_gio_vn = datetime.now(_TZ_VN)
    hom_nay = bay_gio_vn.date()
    gio_hien_tai = bay_gio_vn.time().replace(microsecond=0, tzinfo=None)
    cur = conn.cursor()
    cur.execute(
        "SELECT id, gio_vao, gio_ra FROM cham_cong WHERE nhan_vien_id = %s AND ngay = %s",
        (nhan_vien_id, hom_nay)
    )
    row = cur.fetchone()

    if row is None:
        # Lần đầu trong ngày → ghi giờ vào; ma_cong CHƯA xác định (chờ giờ ra)
        cur.execute("""
            INSERT INTO cham_cong
                (nhan_vien_id, ngay, gio_vao, ma_cong, trang_thai_cham_cong, nguon, created_by, created_at, updated_at)
            VALUES (%s, %s, %s, NULL, 'THIEU_GIO_RA', 'FACE_ID', 'FACE_ID', NOW(), NOW())
        """, (nhan_vien_id, hom_nay, gio_hien_tai))
        conn.commit()
        cur.close()
        return "VAO", gio_hien_tai

    cc_id, gio_vao, gio_ra = row

    if not gio_vao:
        # Có dòng nhưng chưa ghi gio_vao (do chấm thủ công tạo dòng trước)
        cur.execute(
            "UPDATE cham_cong SET gio_vao=%s, trang_thai_cham_cong='THIEU_GIO_RA', nguon='FACE_ID', updated_at=NOW() WHERE id=%s",
            (gio_hien_tai, cc_id)
        )
        conn.commit()
        cur.close()
        return "VAO", gio_hien_tai

    # Đã có gio_vao → cập nhật gio_ra muộn nhất + suy ma_cong (Điểm 2) + gio tang ca (Điểm 3)
    ket_qua = suy_ma_cong_tu_gio_vao_ra(gio_vao, gio_hien_tai, cfg)

    cur.execute("SELECT phong_ban_lam_viec, ho_ten, ma_nv FROM nhan_vien WHERE id = %s", (nhan_vien_id,))
    row_nv = cur.fetchone()
    phong_ban = row_nv[0] if row_nv else None
    ho_ten_nv = row_nv[1] if row_nv else ''
    ma_nv = row_nv[2] if row_nv else ''

    tc = {'tong_gio_tang_ca': 0.0, 'gio_tang_ca_dem': 0.0,
          'loai_ngay_tang_ca': None, 'trang_thai_duyet_tc': None}
    if _cho_phep_tang_ca_phong_ban(conn, phong_ban):
        tc = suy_gio_tang_ca(hom_nay, gio_vao, gio_hien_tai, cfg)

    cur.execute("""
        UPDATE cham_cong
        SET gio_ra=%s, ma_cong=%s, trang_thai_cham_cong=%s,
            so_phut_di_tre=%s, so_phut_ve_som=%s,
            gio_tang_ca=%s, gio_tang_ca_dem=%s, loai_ngay_tang_ca=%s, trang_thai_duyet_tc=%s,
            nguon='FACE_ID', updated_at=NOW()
        WHERE id=%s
    """, (gio_hien_tai, ket_qua['ma_cong'], ket_qua['trang_thai_cham_cong'],
          ket_qua['so_phut_di_tre'], ket_qua['so_phut_ve_som'],
          tc['tong_gio_tang_ca'], tc['gio_tang_ca_dem'], tc['loai_ngay_tang_ca'], tc['trang_thai_duyet_tc'],
          cc_id))
    conn.commit()

    # Gửi cảnh báo đi trễ qua Chat nội bộ cho Trưởng phòng
    canh_bao = ket_qua.get('canh_bao')
    if canh_bao and phong_ban:
        try:
            so_phut = ket_qua['so_phut_di_tre']
            loai_cb = 'CHỜ DUYỆT' if 'CHO_DUYET' in canh_bao else 'CẢNH BÁO'

            noi_dung = (f"⚠️ [{loai_cb}] {ma_nv} - {ho_ten_nv} đi trễ {so_phut} phút "
                        f"(ngày {hom_nay.strftime('%d/%m/%Y')}). "
                        f"Giờ vào: {gio_vao.strftime('%H:%M')}, "
                        f"quy định: {cfg['gio_vao'].strftime('%H:%M')}.")
            if 'CHO_DUYET' in canh_bao:
                noi_dung += " Công ngày này CHƯA được tính — cần admin_bcc xác nhận."

            # Tìm trưởng phòng
            cur2 = conn.cursor()
            cur2.execute("""
                SELECT truong_phong_id FROM danh_muc_phong_ban
                WHERE ten_phong_ban = %s AND truong_phong_id IS NOT NULL
            """, (phong_ban,))
            tp_row = cur2.fetchone()
            nguoi_nhan_id = tp_row[0] if tp_row else None

            if not nguoi_nhan_id:
                # Fallback: gửi cho admin/hr đầu tiên
                cur2.execute("""
                    SELECT id FROM nhan_vien
                    WHERE trang_thai = 'DANG_LAM'
                    ORDER BY id ASC LIMIT 1
                """)
                admin_row = cur2.fetchone()
                nguoi_nhan_id = admin_row[0] if admin_row else None

            if nguoi_nhan_id:
                cur2.execute("""
                    INSERT INTO chat_noi_bo (nguoi_gui_id, nguoi_nhan_id, noi_dung, loai, created_at)
                    VALUES (0, %s, %s, 'CANH_BAO_DI_TRE', NOW())
                """, (nguoi_nhan_id, noi_dung))
                conn.commit()
            cur2.close()
        except Exception:
            pass  # Không để lỗi chat phá vỡ luồng chấm công

    cur.close()
    return "RA", gio_hien_tai


class FaceIDVideoProcessor(VideoProcessorBase):
    """Xử lý khung hình từ camera. Chỉ nhận diện mỗi 2 giây/lần để không quá tải server,
    và mỗi nhân viên chỉ ghi chấm công 1 lần trong 60 giây (tránh ghi trùng liên tục)."""

    def __init__(self):
        self.danh_sach_embedding = []
        self.conn = None
        self.cfg = None
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
                        loai, gio = ghi_nhan_cham_cong(self.conn, nv_id, self.cfg)
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

        # Vẽ khung Elip hướng dẫn đặt khuôn mặt
        cao, rong = img.shape[:2]
        tam_x, tam_y = rong // 2, cao // 2
        truc_x, truc_y = int(rong * 0.22), int(cao * 0.35)
        # Elip viền trắng
        cv2.ellipse(img, (tam_x, tam_y), (truc_x, truc_y), 0, 0, 360, (255, 255, 255), 2)
        # Làm tối vùng ngoài elip (giúp user tập trung vào khuôn mặt)
        mask = np.zeros((cao, rong), dtype=np.uint8)
        cv2.ellipse(mask, (tam_x, tam_y), (truc_x, truc_y), 0, 0, 360, 255, -1)
        img[mask == 0] = (img[mask == 0] * 0.4).astype(np.uint8)
        # Chữ hướng dẫn
        cv2.putText(img, "Dat khuon mat vao khung", (tam_x - 130, tam_y + truc_y + 30),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1, cv2.LINE_AA)

        return av.VideoFrame.from_ndarray(img, format="bgr24")