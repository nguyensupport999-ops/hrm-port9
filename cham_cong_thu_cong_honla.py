# -*- coding: utf-8 -*-
"""
cham_cong_thu_cong_honla.py
============================
Module riêng cho tenant CÔNG TY CỔ PHẦN CẢNG HÒN LA (Mã số thuế: 0108872052).

MỤC TIÊU (theo yêu cầu trong Hoàn_thiện_menu_Chấm_Công.txt):
  Hoàn thiện Tab "📝 Thủ công" trong menu "🕒 Chấm công" để:
  1. Người được ADMIN phân công phụ trách chấm công cho 1 hoặc nhiều phòng/ban
     có thể chấm công cho toàn bộ nhân viên thuộc (các) phòng/ban đó (theo ngày).
  2. Với phòng được cấu hình cho phép tăng ca (OT): người phụ trách chọn số
     giờ OT, phân biệt ca ngày / ca đêm.
  3. Báo cơm ngày hôm sau: mặc định TẤT CẢ nhân viên được BÁO ăn 3 bữa
     (sáng / trưa / tối). Người phụ trách CHỈ cần bấm "Báo cắt" cho nhân viên
     nào nghỉ (off) ngày hôm sau — không cần tick chọn từng bữa khi ăn bình thường.
  4. Dữ liệu chấm công/OT ghi vào bảng `cham_cong` dùng chung toàn hệ thống.
     Dữ liệu báo cơm ghi vào bảng `bao_com` — bảng này CHỈ tạo/dùng cho tenant
     Hòn La (theo đúng yêu cầu: "bảng báo cơm trước mắt chỉ add cho ma_cty Hòn La").
  5. admin_bcc xem bảng tổng hợp báo cơm hàng ngày (toàn công ty) để theo dõi.

================================================================================
QUYẾT ĐỊNH THIẾT KẾ QUAN TRỌNG — PHÂN QUYỀN "NGƯỜI PHỤ TRÁCH CHẤM CÔNG"
================================================================================
KHÔNG dùng khái niệm "Trưởng phòng" cứng (không có cột truong_phong_id, không
suy luận qua chức vụ/chức danh). Thay vào đó:

  - ADMIN vào "⚙️ Danh mục > 🕒 Chấm công" gán TƯỜNG MINH: 1 nhân viên bất kỳ
    được phụ trách chấm công cho 1 hoặc NHIỀU phòng/ban (VD: gán Văn thư phụ
    trách chấm công cho toàn khối "Bộ phận văn phòng"; gán 1 người khác phụ
    trách chấm công cho TẤT CẢ các phòng công nhân).
  - 1 nhân viên có thể vừa giữ vai trò hệ thống khác (VD: admin_bcc) VỪA được
    gán thêm là người phụ trách chấm công — 2 việc này độc lập nhau, không
    xung đột, không cần đổi `role` trong bảng nhan_vien.
  - Việc gán này lưu trong bảng mới `nguoi_phu_trach_cham_cong`
    (nhan_vien_id, ten_phong_ban) — quan hệ NHIỀU-NHIỀU.
  - Vai trò hệ thống `admin` (quản trị cấp cao nhất) mặc định được xem TOÀN
    BỘ phòng ban (để hỗ trợ/khắc phục sự cố), không cần gán riêng.

QUY ƯỚC / GIẢ ĐỊNH KHÁC (đã xác nhận qua Supabase — KHÔNG còn là giả định):
  - Bảng phòng ban thật: `danh_muc_phong_ban` (id, ten_phong_ban, thu_tu,
    trang_thai, created_at, updated_at) — đã XÁC NHẬN qua information_schema.
  - Bảng `nhan_vien`: id, ma_nv, ho_ten, chuc_danh_nghe, chuc_vu, vi_tri_id,
    phong_ban_lam_viec, trang_thai, so_hdld, ...
  - Bảng `cham_cong`: id, nhan_vien_id, ngay, ma_cong, gio_tang_ca,
    gio_tang_ca_dem, loai_ngay_tang_ca, nguon, created_by, updated_at, ...
  - Hàm `get_cau_hinh_tang_ca_theo_phong(ten_phong_ban)` đã có sẵn trong app.py.
  - session_state THẬT (đã xác nhận qua debug thực tế):
        st.session_state.nhan_vien_id      -> id nhân viên đang đăng nhập
        st.session_state.role              -> vai trò hệ thống (admin/hr/
                                               admin_bcc/nhan_vien...)
        st.session_state.tenant['ma_so_thue'] -> mã số thuế tenant hiện tại
        st.session_state.db_engine.get_connection() -> kết nối DB tenant

CÁCH TÍCH HỢP VÀO app.py:
  1. Đặt file này cùng thư mục với app.py.
  2. Thêm dòng import ở đầu app.py:
         import cham_cong_thu_cong_honla as cc_honla

  3. TRONG MENU "🕒 Chấm công" (nơi trước đó anh đã chèn khối expander):
         st.divider()
         st.caption("💡 Thay đổi phương thức chấm công → vào **⚙️ Danh mục** > tab **🕒 Chấm công**")

         if phuong_thuc_cfg == 'THU_CONG':
             with st.expander("🧑‍💼 Chấm công theo phòng/ban (người phụ trách)", expanded=False):
                 cc_honla.render_tab_thu_cong(
                     nhan_vien_hien_tai=cc_honla.lay_nhan_vien_dang_dang_nhap(),
                     ma_so_thue_hien_tai=cc_honla.lay_ma_so_thue_hien_tai(),
                 )

  4. TRONG MENU "⚙️ Danh mục" > tab "🕒 Chấm công" (nơi đang cấu hình
     cc_phuong_thuc, cc_gio_vao, cc_gio_ra...), thêm 1 khu vực mới để ADMIN
     gán người phụ trách chấm công:
         st.divider()
         cc_honla.render_quan_ly_nguoi_phu_trach_cham_cong()

     (Chỉ admin/hr mới nên thấy khu vực này — có thể bọc thêm
      `if st.session_state.get("role") in ("admin", "hr"):` tuỳ quy ước
      phân quyền hiện có trong app.py của anh.)

  5. Nếu admin_bcc cần xem tổng hợp báo cơm:
         cc_honla.render_bang_tong_hop_bao_com_admin()
"""

import psycopg2
import psycopg2.extras
import streamlit as st
from datetime import date, timedelta

# ============================================================
# HẰNG SỐ
# ============================================================

HON_LA_MA_SO_THUE = "0108872052"

DANH_SACH_BUA_AN = [
    ("an_sang", "🌅 Sáng"),
    ("an_trua", "🍚 Trưa"),
    ("an_toi", "🌙 Tối"),
]

MA_CHAM_CONG_NHANH = {
    "x": "✅ Đi làm (đủ công)",
    "x/2": "🌗 Đi làm nửa ngày",
    "CT": "🚗 Công tác",
    "": "⬜ Chưa chấm / để trống",
}

CA_TANG_CA = {"NGAY": "☀️ Ca ngày", "DEM": "🌙 Ca đêm"}


# ============================================================
# TIỆN ÍCH KẾT NỐI DB & SESSION
# ============================================================

def _get_conn():
    """Lấy kết nối DB của tenant hiện tại (đúng cơ chế multi-tenant control_plane)."""
    return st.session_state.db_engine.get_connection()


def lay_nhan_vien_dang_dang_nhap():
    """
    Trả về dict thông tin nhân viên đang đăng nhập:
        {'id': int, 'ma_nv': str, 'ho_ten': str, 'phong_ban_lam_viec': str,
         'vai_tro': str}

    ĐÃ KHỚP session_state thực tế:
        st.session_state.nhan_vien_id -> id nhân viên
        st.session_state.role         -> vai trò hệ thống
    """
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
    """Lấy mã số thuế tenant hiện tại từ st.session_state.tenant['ma_so_thue']."""
    tenant = st.session_state.get("tenant") or {}
    return tenant.get("ma_so_thue", "")


# ============================================================
# ĐẢM BẢO SCHEMA (an toàn khi chạy nhiều lần)
# ============================================================

def ensure_bang_phan_cong_cham_cong():
    """
    Tạo bảng `nguoi_phu_trach_cham_cong` — quan hệ NHIỀU-NHIỀU giữa nhân viên
    và phòng/ban mà họ được ADMIN phân công phụ trách chấm công.
    Bảng này dùng chung cho MỌI tenant (không riêng Hòn La), vì đây là cơ chế
    phân quyền chấm công tổng quát.
    """
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
    """Tạo bảng `bao_com` NẾU tenant hiện tại là Hòn La. Không tạo cho tenant khác."""
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
    """
    Danh sách tên phòng ban của tenant — ưu tiên lấy từ `danh_muc_phong_ban`
    (bảng danh mục chính thức). Nếu bảng trống, fallback lấy DISTINCT từ cột
    `phong_ban_lam_viec` trong `nhan_vien` để không bị chặn thao tác.
    """
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
    """Danh sách toàn bộ nhân viên đang làm việc — dùng cho ADMIN chọn khi gán phân công."""
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
    """Gán 1 nhân viên phụ trách chấm công cho danh sách phòng/ban (nhiều-nhiều)."""
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
    """Danh sách phân công hiện có, kèm tên nhân viên — hiển thị cho admin xem/xoá."""
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
    """
    Trả về danh sách tên phòng ban mà nhân viên hiện tại ĐƯỢC PHÉP chấm công thay.
    - role = 'admin': toàn bộ phòng ban của tenant (hỗ trợ/khắc phục sự cố).
    - Còn lại: CHỈ theo bảng `nguoi_phu_trach_cham_cong` (ADMIN đã gán tường minh).
    """
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
    """UI để ADMIN gán 1 nhân viên phụ trách chấm công cho 1 hoặc nhiều phòng/ban."""
    ensure_bang_phan_cong_cham_cong()

    st.markdown("### 🧑‍💼 Người phụ trách chấm công theo phòng/ban")
    st.caption(
        "Gán 1 nhân viên bất kỳ được phép chấm công thay cho 1 hoặc nhiều phòng/ban. "
        "1 người có thể vừa giữ vai trò khác (VD: admin_bcc) vừa được gán thêm ở đây — "
        "2 việc độc lập nhau. VD: gán Văn thư phụ trách toàn bộ khối Văn phòng; "
        "gán 1 người khác phụ trách toàn bộ các phòng công nhân."
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
# TRUY VẤN NGHIỆP VỤ CHẤM CÔNG / BÁO CƠM
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


def lay_bao_com_ngay(nhan_vien_ids, ngay: date):
    if not nhan_vien_ids:
        return {}
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """SELECT nhan_vien_id, an_sang, an_trua, an_toi, trang_thai
           FROM bao_com WHERE nhan_vien_id = ANY(%s) AND ngay = %s""",
        (nhan_vien_ids, ngay),
    )
    ket_qua = {r["nhan_vien_id"]: r for r in c.fetchall()}
    c.close()
    db.close()
    return ket_qua


def _lay_cau_hinh_tang_ca_an_toan(ten_phong_ban: str):
    try:
        from app import get_cau_hinh_tang_ca_theo_phong
        return get_cau_hinh_tang_ca_theo_phong(ten_phong_ban)
    except Exception:
        return {"cho_phep_tang_ca": False}


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
            trang_thai_bc = d.get("bao_com_trang_thai", "BAO")
            an = trang_thai_bc == "BAO"
            ly_do_cat = d.get("ly_do_cat") or None

            c.execute(
                """
                INSERT INTO bao_com
                    (nhan_vien_id, ngay, an_sang, an_trua, an_toi,
                     trang_thai, ly_do_cat, phong_ban, nguoi_bao, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, NOW())
                ON CONFLICT (nhan_vien_id, ngay) DO UPDATE SET
                    an_sang = EXCLUDED.an_sang,
                    an_trua = EXCLUDED.an_trua,
                    an_toi = EXCLUDED.an_toi,
                    trang_thai = EXCLUDED.trang_thai,
                    ly_do_cat = EXCLUDED.ly_do_cat,
                    phong_ban = EXCLUDED.phong_ban,
                    nguoi_bao = EXCLUDED.nguoi_bao,
                    updated_at = NOW()
                """,
                (nv_id, ngay_bao_com, an, an, an, trang_thai_bc, ly_do_cat, ten_phong_ban, nguoi_luu),
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
            f"Mặc định TẤT CẢ được BÁO ĂN — chỉ cần bấm **Báo cắt** cho nhân viên nghỉ."
        )

    ds_nv = lay_danh_sach_nhan_vien_theo_phong(ten_phong_chon)
    if not ds_nv:
        st.warning("Phòng ban này hiện không có nhân viên nào đang làm việc.")
        return

    nv_ids = [nv["id"] for nv in ds_nv]
    cham_cong_hien_co = lay_cham_cong_ngay(nv_ids, ngay_cham)
    bao_com_hien_co = lay_bao_com_ngay(nv_ids, ngay_cham + timedelta(days=1)) if is_hon_la else {}

    cfg_tc = _lay_cau_hinh_tang_ca_an_toan(ten_phong_chon)
    cho_phep_tang_ca = cfg_tc.get("cho_phep_tang_ca", False)

    st.divider()

    if cho_phep_tang_ca and is_hon_la:
        so_cot = [3, 2, 2, 2]
        tieu_de = ["Nhân viên", "Đi làm hôm nay", "Tăng ca (giờ)", "🍚 Báo cơm hôm sau"]
    elif cho_phep_tang_ca:
        so_cot = [3, 3, 3]
        tieu_de = ["Nhân viên", "Đi làm hôm nay", "Tăng ca (giờ)"]
    elif is_hon_la:
        so_cot = [3, 3, 3]
        tieu_de = ["Nhân viên", "Đi làm hôm nay", "🍚 Báo cơm hôm sau"]
    else:
        so_cot = [4, 4]
        tieu_de = ["Nhân viên", "Đi làm hôm nay"]

    header_cols = st.columns(so_cot)
    for col, tt in zip(header_cols, tieu_de):
        col.markdown(f"**{tt}**")

    key_state = f"cc_tp_data_{ten_phong_chon}_{ngay_cham}"
    if key_state not in st.session_state:
        st.session_state[key_state] = {}

    for nv in ds_nv:
        nv_id = nv["id"]
        rec_cc = cham_cong_hien_co.get(nv_id, {})
        rec_bc = bao_com_hien_co.get(nv_id, {})

        gia_tri_luu = st.session_state[key_state].setdefault(nv_id, {
            "ma_cong": rec_cc.get("ma_cong") or "x",
            "gio_tang_ca": float(rec_cc.get("gio_tang_ca") or 0),
            "gio_tang_ca_dem": float(rec_cc.get("gio_tang_ca_dem") or 0),
            "bao_com_trang_thai": rec_bc.get("trang_thai") or "BAO",
            "ly_do_cat": "",
        })

        cols = st.columns(so_cot)
        cidx = 0

        with cols[cidx]:
            st.markdown(f"**{nv['ho_ten']}**  \n`{nv['ma_nv']}` · {nv.get('chuc_danh_nghe') or ''}")
        cidx += 1

        with cols[cidx]:
            ma_chon = st.selectbox(
                "Trạng thái", list(MA_CHAM_CONG_NHANH.keys()),
                format_func=lambda k: MA_CHAM_CONG_NHANH[k],
                index=list(MA_CHAM_CONG_NHANH.keys()).index(gia_tri_luu["ma_cong"])
                if gia_tri_luu["ma_cong"] in MA_CHAM_CONG_NHANH else 0,
                key=f"cc_tp_ma_{nv_id}_{ngay_cham}",
                label_visibility="collapsed",
            )
            gia_tri_luu["ma_cong"] = ma_chon
        cidx += 1

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
                dang_bao = gia_tri_luu["bao_com_trang_thai"] == "BAO"
                nhan_nut = "🍚 Đang BÁO ĂN — bấm để BÁO CẮT" if dang_bao else "🚫 ĐÃ BÁO CẮT — bấm để báo ăn lại"
                loai_nut = "secondary" if dang_bao else "primary"
                if st.button(nhan_nut, key=f"cc_tp_baocom_{nv_id}_{ngay_cham}",
                             type=loai_nut, use_container_width=True):
                    gia_tri_luu["bao_com_trang_thai"] = "CAT" if dang_bao else "BAO"
                    st.rerun()

                if gia_tri_luu["bao_com_trang_thai"] == "CAT":
                    gia_tri_luu["ly_do_cat"] = st.text_input(
                        "Lý do cắt cơm (tuỳ chọn)",
                        value=gia_tri_luu.get("ly_do_cat", ""),
                        key=f"cc_tp_lydocat_{nv_id}_{ngay_cham}",
                        label_visibility="collapsed",
                        placeholder="VD: nghỉ phép, đi công tác...",
                    )

        st.session_state[key_state][nv_id] = gia_tri_luu

    st.divider()
    col_luu1, col_luu2 = st.columns([1, 3])
    with col_luu1:
        luu_clicked = st.button("💾 Lưu chấm công" + (" & báo cơm" if is_hon_la else ""),
                                 type="primary", use_container_width=True)
    with col_luu2:
        if is_hon_la:
            so_bao = sum(1 for v in st.session_state[key_state].values() if v["bao_com_trang_thai"] == "BAO")
            so_cat = len(st.session_state[key_state]) - so_bao
            st.caption(f"🍚 Đang báo ăn: **{so_bao}** người · 🚫 Báo cắt: **{so_cat}** người")

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
               COUNT(*) FILTER (WHERE bc.trang_thai = 'BAO' AND bc.an_sang) AS so_suat_sang,
               COUNT(*) FILTER (WHERE bc.trang_thai = 'BAO' AND bc.an_trua) AS so_suat_trua,
               COUNT(*) FILTER (WHERE bc.trang_thai = 'BAO' AND bc.an_toi) AS so_suat_toi,
               COUNT(*) FILTER (WHERE bc.trang_thai = 'CAT') AS so_nguoi_cat
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
        "so_nguoi_cat": "🚫 Đã báo cắt",
    })
    tong = {
        "Phòng ban": "TỔNG CỘNG",
        "🌅 Sáng": df["🌅 Sáng"].sum(),
        "🍚 Trưa": df["🍚 Trưa"].sum(),
        "🌙 Tối": df["🌙 Tối"].sum(),
        "🚫 Đã báo cắt": df["🚫 Đã báo cắt"].sum(),
    }
    df = pd.concat([df, pd.DataFrame([tong])], ignore_index=True)

    st.markdown(f"### 🍚 Tổng hợp báo cơm ngày {ngay_chon.strftime('%d/%m/%Y')}")
    st.dataframe(df, hide_index=True, use_container_width=True)