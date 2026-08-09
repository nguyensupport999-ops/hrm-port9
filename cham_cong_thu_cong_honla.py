# -*- coding: utf-8 -*-
"""
cham_cong_thu_cong_honla.py
============================
Module riêng cho tenant CÔNG TY CỔ PHẦN CẢNG HÒN LA (Mã số thuế: 0108872052).

LỊCH SỬ CẬP NHẬT QUAN TRỌNG:
  - Bỏ hẳn cách gọi `from app import get_cau_hinh_tang_ca_theo_phong` vì Streamlit
    chạy app.py như chương trình chính (__main__) — import ngược lại "app" khiến
    Python NẠP LẠI TOÀN BỘ file app.py giữa phiên đang chạy, gây lỗi âm thầm
    (bị except nuốt mất) và làm sai lệch cấu hình tăng ca. Module này giờ tự
    truy vấn thẳng bảng `cau_hinh_tang_ca_phong_ban`, độc lập hoàn toàn với app.py.
  - Thêm đầy đủ 19 ký hiệu chấm công chuẩn (theo TASK_CHAM_CONG_FACE_ID.md) vào
    dropdown "Chấm công hôm nay" (trước đây chỉ có 3 mã rút gọn).
  - Báo cơm tách riêng 3 checkbox Sáng/Trưa/Tối, mặc định TRUE, và LƯU LẠI làm
    mặc định cho các lần chấm công sau của đúng nhân viên đó (không cần chỉnh
    lại từ đầu mỗi ngày).
  - Thêm cảnh báo nghỉ liên tục dài ngày (OD/TN/KL ≥ 14 ngày, hoặc TS) để nhắc
    lập hồ sơ báo giảm BHXH — CHỈ CẢNH BÁO, KHÔNG tự động ghi đè dữ liệu BHXH
    (vì cần xác nhận đúng mã phương án điều chỉnh trước khi ghi vào dữ liệu
    pháp lý — xem ghi chú trong hàm render_canh_bao_bao_giam_bhxh()).

================================================================================
QUYẾT ĐỊNH THIẾT KẾ QUAN TRỌNG — PHÂN QUYỀN "NGƯỜI PHỤ TRÁCH CHẤM CÔNG"
================================================================================
KHÔNG dùng khái niệm "Trưởng phòng" cứng. Thay vào đó:
  - ADMIN vào "⚙️ Danh mục > 🕒 Chấm công" gán TƯỜNG MINH: 1 nhân viên bất kỳ
    được phụ trách chấm công cho 1 hoặc NHIỀU phòng/ban.
  - 1 nhân viên có thể vừa giữ vai trò hệ thống khác (VD: admin_bcc) VỪA được
    gán thêm là người phụ trách chấm công — 2 việc độc lập nhau.
  - Việc gán lưu trong bảng `nguoi_phu_trach_cham_cong` (nhan_vien_id, ten_phong_ban).
  - Vai trò hệ thống `admin` mặc định xem TOÀN BỘ phòng ban.

SCHEMA THẬT ĐÃ XÁC NHẬN QUA SUPABASE (không còn là giả định):
  - `danh_muc_phong_ban`: id, ten_phong_ban, thu_tu, trang_thai, created_at, updated_at
  - `cau_hinh_tang_ca_phong_ban`: id, ten_phong_ban, cho_phep_tang_ca,
        he_so_tc_thuong, he_so_tc_chu_nhat, he_so_tc_le, he_so_tc_dem,
        don_gia_tc_thuong, don_gia_tc_chu_nhat, don_gia_tc_le, don_gia_tc_dem,
        ghi_chu, updated_at
        (các hệ số/đơn giá có thể NULL ở từng phòng → fallback về cấu hình
        chung `_cau_hinh_cache` trong session_state — xem `_lay_cau_hinh_tang_ca()`)
  - `nhan_vien`: id, ma_nv, ho_ten, chuc_danh_nghe, chuc_vu, vi_tri_id,
        phong_ban_lam_viec, trang_thai, so_hdld, ...
  - `cham_cong`: id, nhan_vien_id, ngay, ma_cong, gio_tang_ca, gio_tang_ca_dem,
        loai_ngay_tang_ca, nguon, created_by, updated_at, ...
  - session_state THẬT:
        st.session_state.nhan_vien_id            -> id nhân viên đang đăng nhập
        st.session_state.role                     -> vai trò hệ thống
        st.session_state.tenant['ma_so_thue']     -> mã số thuế tenant
        st.session_state._cau_hinh_cache          -> dict cấu hình chấm công
                                                      CHUNG của tenant (đã cache
                                                      sẵn, key dạng 'cc_he_so_tc_thuong',
                                                      'cc_don_gia_tc_thuong',
                                                      'cc_cach_tinh_tang_ca'...)
        st.session_state.db_engine.get_connection() -> kết nối DB tenant

CÁCH TÍCH HỢP VÀO app.py: xem lại hướng dẫn đã trao đổi trước đó — không đổi
vị trí chèn, chỉ thay nội dung file module này.
"""

import psycopg2
import psycopg2.extras
import streamlit as st
from datetime import date, timedelta

# ============================================================
# HẰNG SỐ
# ============================================================

HON_LA_MA_SO_THUE = "0108872052"

# 19 ký hiệu chấm công theo ngày (KHÔNG gồm nhóm D - tăng ca, vì tăng ca ghi
# riêng ở cột giờ). Theo đúng bảng chuẩn hoá trong TASK_CHAM_CONG_FACE_ID.md.
KY_HIEU_CHAM_CONG_NGAY = {
    "x":     "✅ x — Đi làm ngày thường",
    "x/2":   "🌗 x/2 — Đi làm nửa ngày",
    "P":     "🏖️ P — Nghỉ phép năm (nguyên lương)",
    "1/2P":  "🏖️ 1/2P — Nghỉ phép nửa ngày",
    "NL":    "🎌 NL — Nghỉ lễ",
    "CN":    "📅 CN — Nghỉ hàng tuần (Chủ nhật)",
    "CT":    "🚗 CT — Công tác",
    "NB":    "🔁 NB — Nghỉ bù (đã tăng ca không nhận tiền)",
    "Ro":    "🙋 Ro — Nghỉ việc riêng (hưởng lương)",
    "OD":    "🤒 OD — Nghỉ ốm (có giấy y tế)",
    "CÔ":    "👶 CÔ — Nghỉ con ốm",
    "TS":    "🤰 TS — Thai sản",
    "KT":    "🩺 KT — Khám thai",
    "TN":    "🚑 TN — Tai nạn lao động",
    "DSOD":  "💆 DSOD — Dưỡng sức sau ốm đau",
    "DSTS":  "💆 DSTS — Dưỡng sức sau thai sản",
    "DSTN":  "💆 DSTN — Dưỡng sức sau TNLĐ",
    "KL":    "🚫 KL — Nghỉ không lương (đã duyệt)",
    "KP":    "❌ KP — Nghỉ không phép",
    "":      "⬜ (Chưa chấm / để trống)",
}

# Các mã cần cảnh báo báo giảm BHXH khi nghỉ LIÊN TỤC >= 14 ngày
NHOM_CANH_BAO_14_NGAY = {"OD", "TN", "KL"}
NGUONG_NGAY_CANH_BAO = 14

# Mã báo giảm ngay lập tức khi vừa xuất hiện (không cần đủ 14 ngày)
MA_BAO_GIAM_NGAY_LAP_TUC = {"TS"}

DANH_SACH_BUA_AN = [
    ("an_sang", "🌅 Sáng"),
    ("an_trua", "🍚 Trưa"),
    ("an_toi", "🌙 Tối"),
]


# ============================================================
# TIỆN ÍCH KẾT NỐI DB & SESSION
# ============================================================

def _get_conn():
    return st.session_state.db_engine.get_connection()


def lay_nhan_vien_dang_dang_nhap():
    nv_id = st.session_state.get("nhan_vien_id")
    if not nv_id:
        return None

    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        "SELECT id, ma_nv, ho_ten, phong_ban_lam_viec FROM nhan_vien WHERE id = %s",
        (nv_id,),
    )
    row = c.fetchone()
    c.close()
    db.close()
    if not row:
        return None

    ket_qua = dict(row)
    ket_qua["vai_tro"] = st.session_state.get("role", "nhan_vien")
    return ket_qua


def lay_ma_so_thue_hien_tai():
    tenant = st.session_state.get("tenant") or {}
    return tenant.get("ma_so_thue", "")


# ============================================================
# ĐẢM BẢO SCHEMA (an toàn khi chạy nhiều lần)
# ============================================================

def ensure_bang_phan_cong_cham_cong():
    db = _get_conn()
    c = db.cursor()
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS nguoi_phu_trach_cham_cong (
            id              SERIAL PRIMARY KEY,
            nhan_vien_id    INT4 NOT NULL REFERENCES nhan_vien(id),
            ten_phong_ban   TEXT NOT NULL,
            created_at      TIMESTAMP NOT NULL DEFAULT NOW(),
            created_by      TEXT,
            UNIQUE (nhan_vien_id, ten_phong_ban)
        )
        """
    )
    db.commit()
    c.close()
    db.close()


def ensure_bao_com_table(ma_so_thue_hien_tai: str):
    if (ma_so_thue_hien_tai or "").strip() != HON_LA_MA_SO_THUE:
        return

    db = _get_conn()
    c = db.cursor()
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS bao_com (
            id              SERIAL PRIMARY KEY,
            nhan_vien_id    INT4 NOT NULL REFERENCES nhan_vien(id),
            ngay            DATE NOT NULL,
            an_sang         BOOLEAN NOT NULL DEFAULT TRUE,
            an_trua         BOOLEAN NOT NULL DEFAULT TRUE,
            an_toi          BOOLEAN NOT NULL DEFAULT TRUE,
            trang_thai      TEXT NOT NULL DEFAULT 'BAO',
            ly_do_cat       TEXT,
            phong_ban       TEXT,
            nguoi_bao       TEXT,
            created_at      TIMESTAMP NOT NULL DEFAULT NOW(),
            updated_at      TIMESTAMP NOT NULL DEFAULT NOW(),
            UNIQUE (nhan_vien_id, ngay)
        )
        """
    )
    db.commit()
    c.close()
    db.close()


# ============================================================
# TRUY VẤN: DANH MỤC PHÒNG BAN / NHÂN VIÊN (dùng chung)
# ============================================================

def lay_danh_sach_ten_phong_ban_he_thong():
    db = _get_conn()
    c = db.cursor()
    c.execute("SELECT ten_phong_ban FROM danh_muc_phong_ban ORDER BY thu_tu, ten_phong_ban")
    ds = [r[0] for r in c.fetchall()]
    if not ds:
        c.execute(
            """SELECT DISTINCT phong_ban_lam_viec FROM nhan_vien
               WHERE phong_ban_lam_viec IS NOT NULL AND phong_ban_lam_viec != ''
               ORDER BY phong_ban_lam_viec"""
        )
        ds = [r[0] for r in c.fetchall()]
    c.close()
    db.close()
    return ds


def lay_danh_sach_nhan_vien_toan_bo():
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """SELECT id, ma_nv, ho_ten, phong_ban_lam_viec FROM nhan_vien
           WHERE trang_thai IN ('DANG_LAM','THU_VIEC')
           ORDER BY ho_ten ASC"""
    )
    ds = c.fetchall()
    c.close()
    db.close()
    return ds


# ============================================================
# PHÂN QUYỀN: NGƯỜI PHỤ TRÁCH CHẤM CÔNG (ADMIN quản lý)
# ============================================================

def gan_nguoi_phu_trach(nhan_vien_id: int, danh_sach_ten_phong_ban: list, nguoi_gan: str):
    if not danh_sach_ten_phong_ban:
        return 0
    db = _get_conn()
    c = db.cursor()
    n = 0
    for ten_pb in danh_sach_ten_phong_ban:
        c.execute(
            """INSERT INTO nguoi_phu_trach_cham_cong (nhan_vien_id, ten_phong_ban, created_by)
               VALUES (%s, %s, %s)
               ON CONFLICT (nhan_vien_id, ten_phong_ban) DO NOTHING""",
            (nhan_vien_id, ten_pb, nguoi_gan),
        )
        n += 1
    db.commit()
    c.close()
    db.close()
    return n


def xoa_phan_cong(id_phan_cong: int):
    db = _get_conn()
    c = db.cursor()
    c.execute("DELETE FROM nguoi_phu_trach_cham_cong WHERE id = %s", (id_phan_cong,))
    db.commit()
    c.close()
    db.close()


def lay_bang_phan_cong_hien_tai():
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """SELECT pc.id, pc.nhan_vien_id, nv.ma_nv, nv.ho_ten, pc.ten_phong_ban, pc.created_at
           FROM nguoi_phu_trach_cham_cong pc
           JOIN nhan_vien nv ON nv.id = pc.nhan_vien_id
           ORDER BY nv.ho_ten, pc.ten_phong_ban"""
    )
    ds = c.fetchall()
    c.close()
    db.close()
    return ds


def lay_danh_sach_phong_ban_quan_ly(nhan_vien_hien_tai: dict):
    vai_tro = nhan_vien_hien_tai.get("vai_tro", "")
    if vai_tro == "admin":
        return lay_danh_sach_ten_phong_ban_he_thong()

    db = _get_conn()
    c = db.cursor()
    c.execute(
        "SELECT ten_phong_ban FROM nguoi_phu_trach_cham_cong WHERE nhan_vien_id = %s ORDER BY ten_phong_ban",
        (nhan_vien_hien_tai["id"],),
    )
    ds = [r[0] for r in c.fetchall()]
    c.close()
    db.close()
    return ds


# ============================================================
# GIAO DIỆN ADMIN: GÁN NGƯỜI PHỤ TRÁCH CHẤM CÔNG
# (đặt trong ⚙️ Danh mục > tab 🕒 Chấm công)
# ============================================================

def render_quan_ly_nguoi_phu_trach_cham_cong():
    ensure_bang_phan_cong_cham_cong()

    st.markdown("### 🧑‍💼 Người phụ trách chấm công theo phòng/ban")
    st.caption(
        "Gán 1 nhân viên bất kỳ được phép chấm công thay cho 1 hoặc nhiều phòng/ban. "
        "1 người có thể vừa giữ vai trò khác (VD: admin_bcc) vừa được gán thêm ở đây."
    )

    ds_nhan_vien = lay_danh_sach_nhan_vien_toan_bo()
    ds_phong_ban = lay_danh_sach_ten_phong_ban_he_thong()

    if not ds_nhan_vien or not ds_phong_ban:
        st.warning("Chưa có đủ dữ liệu nhân viên/phòng ban để gán.")
        return

    map_nv = {f"{nv['ho_ten']} ({nv['ma_nv']})": nv["id"] for nv in ds_nhan_vien}

    with st.form("form_gan_nguoi_phu_trach_cham_cong", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nv_chon_label = st.selectbox("Chọn nhân viên", list(map_nv.keys()))
        with col2:
            phong_ban_chon = st.multiselect("Phụ trách phòng/ban", ds_phong_ban)

        submit = st.form_submit_button("➕ Gán phụ trách", type="primary")
        if submit:
            if not phong_ban_chon:
                st.warning("Vui lòng chọn ít nhất 1 phòng/ban.")
            else:
                nv_id = map_nv[nv_chon_label]
                nguoi_gan = st.session_state.get("username", "admin")
                so_luong = gan_nguoi_phu_trach(nv_id, phong_ban_chon, nguoi_gan)
                st.success(f"✅ Đã gán {nv_chon_label} phụ trách chấm công cho {so_luong} phòng/ban.")
                st.rerun()

    st.divider()
    st.markdown("#### 📋 Danh sách đang phân công")
    ds_hien_tai = lay_bang_phan_cong_hien_tai()
    if not ds_hien_tai:
        st.info("Chưa có phân công nào.")
        return

    for row in ds_hien_tai:
        cols = st.columns([3, 3, 2, 1])
        cols[0].write(f"**{row['ho_ten']}** (`{row['ma_nv']}`)")
        cols[1].write(row["ten_phong_ban"])
        cols[2].write(row["created_at"].strftime("%d/%m/%Y") if row["created_at"] else "")
        if cols[3].button("🗑️", key=f"xoa_phancong_{row['id']}"):
            xoa_phan_cong(row["id"])
            st.rerun()


# ============================================================
# TRUY VẤN NGHIỆP VỤ CHẤM CÔNG / OT / BÁO CƠM
# ============================================================

def lay_danh_sach_nhan_vien_theo_phong(ten_phong_ban: str):
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """SELECT id, ma_nv, ho_ten, chuc_danh_nghe FROM nhan_vien
           WHERE trang_thai IN ('DANG_LAM','THU_VIEC')
             AND phong_ban_lam_viec = %s
             AND so_hdld IS NOT NULL
           ORDER BY ma_nv ASC""",
        (ten_phong_ban,),
    )
    ds = c.fetchall()
    c.close()
    db.close()
    return ds


def lay_cham_cong_ngay(nhan_vien_ids, ngay: date):
    if not nhan_vien_ids:
        return {}
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """SELECT nhan_vien_id, ma_cong, gio_tang_ca, gio_tang_ca_dem, loai_ngay_tang_ca
           FROM cham_cong WHERE nhan_vien_id = ANY(%s) AND ngay = %s""",
        (nhan_vien_ids, ngay),
    )
    ket_qua = {r["nhan_vien_id"]: r for r in c.fetchall()}
    c.close()
    db.close()
    return ket_qua


def lay_bao_com_ngay_cu_the(nhan_vien_ids, ngay: date):
    """Bản ghi báo cơm ĐÚNG ngày cần xem (nếu người phụ trách đã chấm ngày đó rồi)."""
    if not nhan_vien_ids:
        return {}
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """SELECT nhan_vien_id, an_sang, an_trua, an_toi
           FROM bao_com WHERE nhan_vien_id = ANY(%s) AND ngay = %s""",
        (nhan_vien_ids, ngay),
    )
    ket_qua = {r["nhan_vien_id"]: r for r in c.fetchall()}
    c.close()
    db.close()
    return ket_qua


def lay_bao_com_mac_dinh_gan_nhat(nhan_vien_ids):
    """
    Bản ghi báo cơm GẦN NHẤT (bất kỳ ngày nào trước đó) của từng nhân viên —
    dùng làm giá trị MẶC ĐỊNH khi chưa có dữ liệu cho đúng ngày cần chấm.
    Nhờ đó: NV nào mọi khi chỉ báo ăn trưa thì lần chấm mới cũng tự động giữ
    nguyên "chỉ ăn trưa", không cần người phụ trách tick lại từ đầu.
    """
    if not nhan_vien_ids:
        return {}
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """SELECT DISTINCT ON (nhan_vien_id)
               nhan_vien_id, an_sang, an_trua, an_toi
           FROM bao_com
           WHERE nhan_vien_id = ANY(%s)
           ORDER BY nhan_vien_id, ngay DESC""",
        (nhan_vien_ids,),
    )
    ket_qua = {r["nhan_vien_id"]: r for r in c.fetchall()}
    c.close()
    db.close()
    return ket_qua


def _lay_cau_hinh_tang_ca(ten_phong_ban: str):
    """
    Tự truy vấn thẳng bảng `cau_hinh_tang_ca_phong_ban` (KHÔNG import app.py —
    xem ghi chú đầu file lý do vì sao cách cũ gây lỗi âm thầm).
    Field nào bị NULL ở cấu hình riêng phòng ban sẽ fallback về cấu hình
    CHUNG của tenant (đọc từ st.session_state._cau_hinh_cache, đã được app.py
    cache sẵn khi tải trang — đúng dữ liệu hiển thị trong Danh mục > Chấm công).
    """
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        "SELECT * FROM cau_hinh_tang_ca_phong_ban WHERE ten_phong_ban = %s",
        (ten_phong_ban,),
    )
    row = c.fetchone()
    c.close()
    db.close()

    if not row or not row.get("cho_phep_tang_ca"):
        return {"cho_phep_tang_ca": False}

    cfg_chung = st.session_state.get("_cau_hinh_cache", {}) or {}

    def _so(gia_tri_rieng, khoa_chung, mac_dinh):
        if gia_tri_rieng is not None:
            return float(gia_tri_rieng)
        try:
            return float(cfg_chung.get(khoa_chung) or mac_dinh)
        except (TypeError, ValueError):
            return mac_dinh

    return {
        "cho_phep_tang_ca": True,
        "cach_tinh_tang_ca": cfg_chung.get("cc_cach_tinh_tang_ca", "HE_SO"),
        "he_so_tc_thuong": _so(row.get("he_so_tc_thuong"), "cc_he_so_tc_thuong", 1.5),
        "he_so_tc_chu_nhat": _so(row.get("he_so_tc_chu_nhat"), "cc_he_so_tc_chu_nhat", 2.0),
        "he_so_tc_le": _so(row.get("he_so_tc_le"), "cc_he_so_tc_le", 3.0),
        "he_so_tc_dem": _so(row.get("he_so_tc_dem"), "cc_he_so_tc_dem", 1.3),
        "don_gia_tc_thuong": _so(row.get("don_gia_tc_thuong"), "cc_don_gia_tc_thuong", 0),
        "don_gia_tc_chu_nhat": _so(row.get("don_gia_tc_chu_nhat"), "cc_don_gia_tc_chu_nhat", 0),
        "don_gia_tc_le": _so(row.get("don_gia_tc_le"), "cc_don_gia_tc_le", 0),
        "don_gia_tc_dem": _so(row.get("don_gia_tc_dem"), "cc_don_gia_tc_dem", 0),
    }


def tinh_so_ngay_nghi_lien_tuc(nhan_vien_id: int, ma_cong_kiem_tra: str, ngay_hien_tai: date):
    """
    Đếm số ngày LIÊN TỤC (tính đến và bao gồm ngay_hien_tai) mà nhân viên có
    đúng ma_cong_kiem_tra — dùng để cảnh báo báo giảm BHXH (OD/TN/KL >= 14 ngày).
    Giá trị của ngay_hien_tai được coi là ma_cong_kiem_tra (vì đây là mã đang
    được người phụ trách CHỌN, có thể chưa lưu vào DB).
    """
    db = _get_conn()
    c = db.cursor()
    c.execute(
        """SELECT ngay, ma_cong FROM cham_cong
           WHERE nhan_vien_id = %s AND ngay < %s AND ngay >= %s
           ORDER BY ngay DESC""",
        (nhan_vien_id, ngay_hien_tai, ngay_hien_tai - timedelta(days=45)),
    )
    lich_su = {r[0]: r[1] for r in c.fetchall()}
    c.close()
    db.close()

    so_ngay = 1  # tính luôn ngày hiện tại đang chọn
    ngay_kiem = ngay_hien_tai - timedelta(days=1)
    while lich_su.get(ngay_kiem) == ma_cong_kiem_tra:
        so_ngay += 1
        ngay_kiem -= timedelta(days=1)
    return so_ngay


# ============================================================
# GHI DỮ LIỆU
# ============================================================

def luu_cham_cong_va_bao_com(ten_phong_ban, ngay_cham_cong, du_lieu_nv, is_hon_la, nguoi_luu):
    db = _get_conn()
    c = db.cursor()
    n_cham_cong = 0
    n_bao_com = 0
    ngay_bao_com = ngay_cham_cong + timedelta(days=1)

    for nv_id, d in du_lieu_nv.items():
        ma_cong = (d.get("ma_cong") or "").strip()
        gio_tc = float(d.get("gio_tang_ca") or 0)
        gio_tc_dem = float(d.get("gio_tang_ca_dem") or 0)

        if ma_cong or gio_tc > 0 or gio_tc_dem > 0:
            c.execute(
                """
                INSERT INTO cham_cong
                    (nhan_vien_id, ngay, ma_cong, gio_tang_ca, gio_tang_ca_dem,
                     nguon, created_by, updated_at)
                VALUES (%s, %s, %s, %s, %s, 'THU_CONG', %s, NOW())
                ON CONFLICT (nhan_vien_id, ngay) DO UPDATE SET
                    ma_cong = CASE WHEN EXCLUDED.ma_cong != '' THEN EXCLUDED.ma_cong
                                   ELSE cham_cong.ma_cong END,
                    gio_tang_ca = EXCLUDED.gio_tang_ca,
                    gio_tang_ca_dem = EXCLUDED.gio_tang_ca_dem,
                    nguon = CASE WHEN cham_cong.nguon = 'FACE_ID'
                                 THEN cham_cong.nguon ELSE EXCLUDED.nguon END,
                    updated_at = NOW()
                """,
                (nv_id, ngay_cham_cong, ma_cong, gio_tc, gio_tc_dem, nguoi_luu),
            )
            n_cham_cong += 1

        if is_hon_la:
            an_sang = bool(d.get("an_sang", True))
            an_trua = bool(d.get("an_trua", True))
            an_toi = bool(d.get("an_toi", True))
            trang_thai_bc = "BAO" if (an_sang or an_trua or an_toi) else "CAT"

            c.execute(
                """
                INSERT INTO bao_com
                    (nhan_vien_id, ngay, an_sang, an_trua, an_toi,
                     trang_thai, phong_ban, nguoi_bao, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, NOW())
                ON CONFLICT (nhan_vien_id, ngay) DO UPDATE SET
                    an_sang = EXCLUDED.an_sang,
                    an_trua = EXCLUDED.an_trua,
                    an_toi = EXCLUDED.an_toi,
                    trang_thai = EXCLUDED.trang_thai,
                    phong_ban = EXCLUDED.phong_ban,
                    nguoi_bao = EXCLUDED.nguoi_bao,
                    updated_at = NOW()
                """,
                (nv_id, ngay_bao_com, an_sang, an_trua, an_toi, trang_thai_bc, ten_phong_ban, nguoi_luu),
            )
            n_bao_com += 1

    db.commit()
    c.close()
    db.close()
    return n_cham_cong, n_bao_com


# ============================================================
# GIAO DIỆN CHÍNH — NGƯỜI PHỤ TRÁCH CHẤM CÔNG
# ============================================================

def render_tab_thu_cong(nhan_vien_hien_tai: dict, ma_so_thue_hien_tai: str):
    if not nhan_vien_hien_tai:
        st.warning("⚠️ Không xác định được thông tin nhân viên đang đăng nhập.")
        return

    is_hon_la = (ma_so_thue_hien_tai or "").strip() == HON_LA_MA_SO_THUE
    ensure_bang_phan_cong_cham_cong()
    if is_hon_la:
        ensure_bao_com_table(ma_so_thue_hien_tai)

    danh_sach_phong = lay_danh_sach_phong_ban_quan_ly(nhan_vien_hien_tai)

    if not danh_sach_phong:
        st.info(
            "ℹ️ Bạn hiện chưa được ADMIN phân công phụ trách chấm công cho phòng/ban nào. "
            "Liên hệ Admin để được gán qua **⚙️ Danh mục > 🕒 Chấm công**."
        )
        return

    st.markdown("### 🧑‍💼 Chấm công cho phòng/ban bạn phụ trách")

    col_chon1, col_chon2 = st.columns([1, 2])
    with col_chon1:
        ngay_cham = st.date_input("Ngày cần chấm công", value=date.today(), key="cc_tp_ngay_cham")
    with col_chon2:
        ten_phong_chon = st.selectbox("Phòng ban", danh_sach_phong, key="cc_tp_phong_chon")

    if is_hon_la:
        ngay_bao_com_hien_thi = ngay_cham + timedelta(days=1)
        st.caption(
            f"🍚 Báo cơm áp dụng cho ngày **{ngay_bao_com_hien_thi.strftime('%d/%m/%Y')}**. "
            f"Mặc định giữ nguyên như lần chấm gần nhất của từng nhân viên — chỉ cần "
            f"bỏ tick bữa nào NV không ăn."
        )

    ds_nv = lay_danh_sach_nhan_vien_theo_phong(ten_phong_chon)
    if not ds_nv:
        st.warning("Phòng ban này hiện không có nhân viên nào đang làm việc.")
        return

    nv_ids = [nv["id"] for nv in ds_nv]
    cham_cong_hien_co = lay_cham_cong_ngay(nv_ids, ngay_cham)

    if is_hon_la:
        ngay_bc = ngay_cham + timedelta(days=1)
        bao_com_dung_ngay = lay_bao_com_ngay_cu_the(nv_ids, ngay_bc)
        bao_com_mac_dinh = lay_bao_com_mac_dinh_gan_nhat(nv_ids)
    else:
        bao_com_dung_ngay = {}
        bao_com_mac_dinh = {}

    cfg_tc = _lay_cau_hinh_tang_ca(ten_phong_chon)
    cho_phep_tang_ca = cfg_tc.get("cho_phep_tang_ca", False)

    st.divider()

    if cho_phep_tang_ca and is_hon_la:
        so_cot = [3, 3, 2, 4]
        tieu_de = ["Nhân viên", "Chấm công hôm nay", "Tăng ca (giờ)", "🍚 Báo cơm hôm sau"]
    elif cho_phep_tang_ca:
        so_cot = [3, 4, 3]
        tieu_de = ["Nhân viên", "Chấm công hôm nay", "Tăng ca (giờ)"]
    elif is_hon_la:
        so_cot = [3, 4, 4]
        tieu_de = ["Nhân viên", "Chấm công hôm nay", "🍚 Báo cơm hôm sau"]
    else:
        so_cot = [4, 4]
        tieu_de = ["Nhân viên", "Chấm công hôm nay"]

    header_cols = st.columns(so_cot)
    for col, tt in zip(header_cols, tieu_de):
        col.markdown(f"**{tt}**")

    key_state = f"cc_tp_data_{ten_phong_chon}_{ngay_cham}"
    if key_state not in st.session_state:
        st.session_state[key_state] = {}

    ds_canh_bao = []  # gom cảnh báo báo giảm BHXH hiển thị cuối bảng
    ma_cong_options = list(KY_HIEU_CHAM_CONG_NGAY.keys())

    for nv in ds_nv:
        nv_id = nv["id"]
        rec_cc = cham_cong_hien_co.get(nv_id, {})

        rec_bc_dung_ngay = bao_com_dung_ngay.get(nv_id)
        rec_bc_mac_dinh = bao_com_mac_dinh.get(nv_id)
        if rec_bc_dung_ngay is not None:
            an_sang_mac_dinh = rec_bc_dung_ngay["an_sang"]
            an_trua_mac_dinh = rec_bc_dung_ngay["an_trua"]
            an_toi_mac_dinh = rec_bc_dung_ngay["an_toi"]
        elif rec_bc_mac_dinh is not None:
            an_sang_mac_dinh = rec_bc_mac_dinh["an_sang"]
            an_trua_mac_dinh = rec_bc_mac_dinh["an_trua"]
            an_toi_mac_dinh = rec_bc_mac_dinh["an_toi"]
        else:
            an_sang_mac_dinh = an_trua_mac_dinh = an_toi_mac_dinh = True

        gia_tri_luu = st.session_state[key_state].setdefault(nv_id, {
            "ma_cong": rec_cc.get("ma_cong") or "x",
            "gio_tang_ca": float(rec_cc.get("gio_tang_ca") or 0),
            "gio_tang_ca_dem": float(rec_cc.get("gio_tang_ca_dem") or 0),
            "an_sang": an_sang_mac_dinh,
            "an_trua": an_trua_mac_dinh,
            "an_toi": an_toi_mac_dinh,
        })

        cols = st.columns(so_cot)
        cidx = 0

        with cols[cidx]:
            st.markdown(f"**{nv['ho_ten']}**  \n`{nv['ma_nv']}` · {nv.get('chuc_danh_nghe') or ''}")
        cidx += 1

        with cols[cidx]:
            gia_tri_hien_tai = gia_tri_luu["ma_cong"] if gia_tri_luu["ma_cong"] in KY_HIEU_CHAM_CONG_NGAY else "x"
            ma_chon = st.selectbox(
                "Trạng thái", ma_cong_options,
                format_func=lambda k: KY_HIEU_CHAM_CONG_NGAY[k],
                index=ma_cong_options.index(gia_tri_hien_tai),
                key=f"cc_tp_ma_{nv_id}_{ngay_cham}",
                label_visibility="collapsed",
            )
            gia_tri_luu["ma_cong"] = ma_chon
        cidx += 1

        # ---- Cảnh báo báo giảm BHXH ----
        if ma_chon in NHOM_CANH_BAO_14_NGAY:
            so_ngay = tinh_so_ngay_nghi_lien_tuc(nv_id, ma_chon, ngay_cham)
            if so_ngay >= NGUONG_NGAY_CANH_BAO:
                ds_canh_bao.append(
                    f"🔴 **{nv['ho_ten']}** ({nv['ma_nv']}) — nghỉ **{ma_chon}** liên tục "
                    f"**{so_ngay} ngày** (tính đến {ngay_cham.strftime('%d/%m/%Y')}) → "
                    f"cần lập hồ sơ **BÁO GIẢM BHXH**."
                )
        elif ma_chon in MA_BAO_GIAM_NGAY_LAP_TUC:
            ds_canh_bao.append(
                f"🟡 **{nv['ho_ten']}** ({nv['ma_nv']}) — mã **{ma_chon}** (Thai sản) → "
                f"theo quy định, BÁO GIẢM BHXH áp dụng NGAY từ tháng bắt đầu nghỉ."
            )

        if cho_phep_tang_ca:
            with cols[cidx]:
                sub1, sub2 = st.columns(2)
                with sub1:
                    gio_ngay = st.number_input(
                        "OT ca ngày", min_value=0.0, max_value=12.0, step=0.5,
                        value=gia_tri_luu["gio_tang_ca"],
                        key=f"cc_tp_otngay_{nv_id}_{ngay_cham}",
                        label_visibility="collapsed",
                    )
                    gia_tri_luu["gio_tang_ca"] = gio_ngay
                with sub2:
                    gio_dem = st.number_input(
                        "OT ca đêm", min_value=0.0, max_value=12.0, step=0.5,
                        value=gia_tri_luu["gio_tang_ca_dem"],
                        key=f"cc_tp_otdem_{nv_id}_{ngay_cham}",
                        label_visibility="collapsed",
                    )
                    gia_tri_luu["gio_tang_ca_dem"] = gio_dem
                st.caption("☀️ Ngày · 🌙 Đêm (giờ)")
            cidx += 1

        if is_hon_la:
            with cols[cidx]:
                sub_bc = st.columns(3)
                for (khoa, nhan), sub in zip(DANH_SACH_BUA_AN, sub_bc):
                    with sub:
                        gia_tri_luu[khoa] = st.checkbox(
                            nhan, value=gia_tri_luu[khoa],
                            key=f"cc_tp_{khoa}_{nv_id}_{ngay_cham}",
                        )

        st.session_state[key_state][nv_id] = gia_tri_luu

    # ---- Khối cảnh báo báo giảm BHXH (nếu có) ----
    if ds_canh_bao:
        st.divider()
        st.warning(
            "⚠️ **Cảnh báo báo giảm BHXH** — CHỈ mang tính nhắc nhở, hệ thống KHÔNG tự "
            "động ghi vào hồ sơ BHXH. Vào **📋 BHXH > Báo cáo tăng/giảm D02-LT** để xử lý:\n\n"
            + "\n\n".join(ds_canh_bao)
        )

    st.divider()
    col_luu1, col_luu2 = st.columns([1, 3])
    with col_luu1:
        luu_clicked = st.button("💾 Lưu chấm công" + (" & báo cơm" if is_hon_la else ""),
                                 type="primary", use_container_width=True)
    with col_luu2:
        if is_hon_la:
            so_co_an = sum(
                1 for v in st.session_state[key_state].values()
                if v.get("an_sang") or v.get("an_trua") or v.get("an_toi")
            )
            st.caption(f"🍚 Có báo ăn (≥1 bữa): **{so_co_an}**/{len(st.session_state[key_state])} người")

    if luu_clicked:
        n_cc, n_bc = luu_cham_cong_va_bao_com(
            ten_phong_ban=ten_phong_chon,
            ngay_cham_cong=ngay_cham,
            du_lieu_nv=st.session_state[key_state],
            is_hon_la=is_hon_la,
            nguoi_luu=st.session_state.get("username", nhan_vien_hien_tai.get("ma_nv", "")),
        )
        msg = f"✅ Đã lưu chấm công cho {n_cc} nhân viên"
        if is_hon_la:
            msg += f", báo cơm cho {n_bc} nhân viên (ngày {(ngay_cham + timedelta(days=1)).strftime('%d/%m/%Y')})"
        st.success(msg + ".")
        del st.session_state[key_state]
        st.rerun()


# ============================================================
# GIAO DIỆN TỔNG HỢP — DÀNH CHO admin_bcc (chỉ có ý nghĩa với Hòn La)
# ============================================================

def render_bang_tong_hop_bao_com_admin(ngay_xem: date = None):
    if ngay_xem is None:
        ngay_xem = date.today() + timedelta(days=1)

    db = _get_conn()
    c = db.cursor()
    c.execute(
        """SELECT EXISTS (
               SELECT 1 FROM information_schema.tables WHERE table_name = 'bao_com'
           )"""
    )
    co_bang = c.fetchone()[0]
    if not co_bang:
        c.close()
        db.close()
        st.info("ℹ️ Tenant này chưa sử dụng tính năng Báo cơm.")
        return

    ngay_chon = st.date_input("Ngày báo cơm cần xem", value=ngay_xem, key="cc_admin_bc_ngay")

    c2 = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c2.execute(
        """
        SELECT bc.phong_ban,
               COUNT(*) FILTER (WHERE bc.an_sang) AS so_suat_sang,
               COUNT(*) FILTER (WHERE bc.an_trua) AS so_suat_trua,
               COUNT(*) FILTER (WHERE bc.an_toi) AS so_suat_toi,
               COUNT(*) FILTER (WHERE NOT bc.an_sang AND NOT bc.an_trua AND NOT bc.an_toi) AS so_nguoi_cat_het
        FROM bao_com bc
        WHERE bc.ngay = %s
        GROUP BY bc.phong_ban
        ORDER BY bc.phong_ban
        """,
        (ngay_chon,),
    )
    rows = c2.fetchall()
    c2.close()
    c.close()
    db.close()

    if not rows:
        st.warning(f"Chưa có dữ liệu báo cơm cho ngày {ngay_chon.strftime('%d/%m/%Y')}.")
        return

    import pandas as pd
    df = pd.DataFrame(rows)
    df = df.rename(columns={
        "phong_ban": "Phòng ban",
        "so_suat_sang": "🌅 Sáng",
        "so_suat_trua": "🍚 Trưa",
        "so_suat_toi": "🌙 Tối",
        "so_nguoi_cat_het": "🚫 Cắt cả 3 bữa",
    })
    tong = {
        "Phòng ban": "TỔNG CỘNG",
        "🌅 Sáng": df["🌅 Sáng"].sum(),
        "🍚 Trưa": df["🍚 Trưa"].sum(),
        "🌙 Tối": df["🌙 Tối"].sum(),
        "🚫 Cắt cả 3 bữa": df["🚫 Cắt cả 3 bữa"].sum(),
    }
    df = pd.concat([df, pd.DataFrame([tong])], ignore_index=True)

    st.markdown(f"### 🍚 Tổng hợp báo cơm ngày {ngay_chon.strftime('%d/%m/%Y')}")
    st.dataframe(df, hide_index=True, use_container_width=True)
