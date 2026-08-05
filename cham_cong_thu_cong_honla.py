# -*- coding: utf-8 -*-
"""
cham_cong_thu_cong_honla.py
============================
Module riêng cho tenant CÔNG TY CỔ PHẦN CẢNG HÒN LA (Mã số thuế: 0108872052).

MỤC TIÊU (theo yêu cầu trong Hoàn_thiện_menu_Chấm_Công.txt):
  Hoàn thiện Tab "📝 Thủ công" trong menu "🕒 Chấm công" để:
  1. Trưởng phòng chấm công cho nhân viên trong phòng mình quản lý (theo ngày).
  2. Với phòng được cấu hình cho phép tăng ca (OT): Trưởng phòng chọn số giờ OT,
     phân biệt ca ngày / ca đêm.
  3. Báo cơm ngày hôm sau: mặc định TẤT CẢ nhân viên được BÁO ăn 3 bữa
     (sáng / trưa / tối). Trưởng phòng CHỈ cần bấm "Báo cắt" cho nhân viên nào
     nghỉ (off) ngày hôm sau — không cần tick chọn từng bữa khi ăn bình thường.
  4. Dữ liệu chấm công/OT ghi vào bảng `cham_cong` dùng chung toàn hệ thống.
     Dữ liệu báo cơm ghi vào bảng `bao_com` — bảng này CHỈ tạo/dùng cho tenant
     Hòn La (theo đúng yêu cầu: "bảng báo cơm trước mắt chỉ add cho ma_cty Hòn La").
  5. admin_bcc xem bảng tổng hợp báo cơm hàng ngày (toàn công ty) để theo dõi.

QUY ƯỚC / GIẢ ĐỊNH (đối chiếu context.md, memory.md, TASK_CHAM_CONG_FACE_ID.md):
  - Bảng `nhan_vien`: id, ma_nv, ho_ten, chuc_danh_nghe, phong_ban_lam_viec,
    trang_thai, so_hdld, ...
  - Bảng `phong_ban`: id, ten_phong_ban, truong_phong_id (FK -> nhan_vien.id).
    (Cột truong_phong_id đã được nêu trong memory.md mục 5.5 — nếu DB thật của
    bạn chưa có cột này, xem hàm `ensure_cot_truong_phong()` bên dưới.)
  - Bảng `cham_cong`: id, nhan_vien_id, ngay, ma_cong, gio_tang_ca,
    gio_tang_ca_dem, loai_ngay_tang_ca, nguon, created_by, updated_at, ...
  - Hàm `get_cau_hinh_tang_ca_theo_phong(ten_phong_ban)` đã có sẵn trong app.py
    (trả về dict gồm khoá 'cho_phep_tang_ca', 'cach_tinh_tang_ca',
    'he_so_tc_thuong', 'don_gia_tc_thuong'...) — import lại từ app.py.
  - `st.session_state.db_engine.get_connection()` dùng để lấy kết nối Postgres
    của đúng tenant đang đăng nhập (đã resolve qua control_plane.py).
  - `st.session_state.username`: username đang đăng nhập (dùng cho created_by).
  - Thông tin nhân viên hiện tại (id, mã NV, vai trò...) giả định lưu ở
    `st.session_state.nv_hien_tai` dạng dict — NẾU project của bạn dùng tên
    biến khác, chỉ cần sửa lại hàm `lay_nhan_vien_dang_dang_nhap()` bên dưới,
    KHÔNG cần sửa phần còn lại của file.

CÁCH TÍCH HỢP VÀO app.py:
  1. Đặt file này cùng thư mục với app.py.
  2. Thêm dòng import ở đầu app.py (cạnh các import module khác):
         import cham_cong_thu_cong_honla as cc_honla
  3. Trong khối xử lý menu "🕒 Chấm công", đoạn có 3 nút phương thức
     (📝 Thủ công / 📥 Máy vân tay / 👤 Face ID), thêm 1 st.tabs con để
     tách "BCC tháng" (đã có) và "Trưởng phòng chấm công hôm nay" (mới):

         if phuong_thuc_cfg == 'THU_CONG':
             sub_tab_bcc, sub_tab_tp = st.tabs(["📅 Bảng chấm công tháng", "🧑‍💼 Trưởng phòng chấm công"])
             with sub_tab_tp:
                 cc_honla.render_tab_thu_cong(
                     nhan_vien_hien_tai=cc_honla.lay_nhan_vien_dang_dang_nhap(),
                     ma_so_thue_hien_tai=st.session_state.get("ma_so_thue_hien_tai", ""),
                 )
             with sub_tab_bcc:
                 ... (giữ nguyên toàn bộ code BCC tháng hiện có) ...

     (Vị trí chính xác: ngay sau đoạn `st.caption("💡 Thay đổi phương thức...")`
      và trước dòng `ensure_cham_cong_table()`.)

  4. Nếu admin_bcc cần xem tổng hợp báo cơm, gọi thêm:
         cc_honla.render_bang_tong_hop_bao_com_admin()
     ở trang/menu phù hợp (VD: cuối tab "📋 Báo cáo định kỳ").
"""

import psycopg2
import psycopg2.extras
import streamlit as st
from datetime import date, timedelta, datetime

# ============================================================
# HẰNG SỐ
# ============================================================

# Mã số thuế của Cảng Hòn La — dùng để bật/tắt tính năng "Báo cơm"
HON_LA_MA_SO_THUE = "0108872052"

# Danh sách 3 bữa ăn trong ngày (khoá lưu DB : nhãn hiển thị)
DANH_SACH_BUA_AN = [
    ("an_sang", "🌅 Sáng"),
    ("an_trua", "🍚 Trưa"),
    ("an_toi", "🌙 Tối"),
]

# Các mã chấm công nhanh mà Trưởng phòng có thể chọn trực tiếp (không cần
# quy trình xin phép/duyệt). Các mã cần duyệt (P, OD, KL, Ro, NB, CÔ...) KHÔNG
# đưa vào đây — Trưởng phòng phải hướng dẫn nhân viên xin phép qua "💬 Chat nội bộ"
# như đã thiết kế trong TASK_CHAM_CONG_FACE_ID.md mục 5.3.
MA_CHAM_CONG_NHANH = {
    "x": "✅ Đi làm (đủ công)",
    "x/2": "🌗 Đi làm nửa ngày",
    "CT": "🚗 Công tác",
    "": "⬜ Chưa chấm / để trống",
}

# Ca tăng ca
CA_TANG_CA = {"NGAY": "☀️ Ca ngày", "DEM": "🌙 Ca đêm"}


# ============================================================
# TIỆN ÍCH KẾT NỐI DB
# ============================================================

def _get_conn():
    """Lấy kết nối DB của tenant hiện tại (đúng cơ chế multi-tenant control_plane)."""
    return st.session_state.db_engine.get_connection()


def lay_nhan_vien_dang_dang_nhap():
    """
    Trả về dict thông tin nhân viên đang đăng nhập, tối thiểu gồm:
        {'id': int, 'ma_nv': str, 'ho_ten': str, 'vai_tro': str}

    GHI CHÚ: hàm này giả định project đã lưu sẵn thông tin này ở
    st.session_state.nv_hien_tai. Nếu project của bạn lưu ở nơi khác
    (VD: st.session_state.user_id + truy vấn lại), hãy sửa DUY NHẤT hàm này.
    """
    nv = st.session_state.get("nv_hien_tai")
    if nv:
        return nv

    # Phương án dự phòng: nếu chỉ có user_id, tự truy vấn nhan_vien
    user_id = st.session_state.get("user_id")
    if not user_id:
        return None
    db = _get_conn()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        "SELECT id, ma_nv, ho_ten, vai_tro, phong_ban_lam_viec FROM nhan_vien WHERE id = %s",
        (user_id,),
    )
    row = c.fetchone()
    c.close()
    db.close()
    return dict(row) if row else None


# ============================================================
# ĐẢM BẢO SCHEMA (chạy an toàn nhiều lần — CREATE TABLE IF NOT EXISTS)
# ============================================================

def ensure_bao_com_table(ma_so_thue_hien_tai: str):
    """
    Tạo bảng `bao_com` NẾU tenant hiện tại là Hòn La.
    KHÔNG tạo bảng cho các tenant khác — đúng yêu cầu "trước mắt chỉ add cho
    ma_cty Hòn la thôi".
    """
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
            trang_thai      TEXT NOT NULL DEFAULT 'BAO',   -- 'BAO' | 'CAT'
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


def ensure_cot_truong_phong():
    """
    Đảm bảo bảng `phong_ban` có cột `truong_phong_id` (đã quyết định trong
    memory.md mục 5.5, dùng chung cho cả tính năng "phê duyệt vị trí chấm công"
    và tính năng Trưởng phòng chấm công ở đây). An toàn khi chạy nhiều lần.
    """
    db = _get_conn()
    c = db.cursor()
    c.execute(
        """
        ALTER TABLE phong_ban
        ADD COLUMN IF NOT EXISTS truong_phong_id INT4 REFERENCES nhan_vien(id)
        """
    )
    db.commit()
    c.close()
    db.close()


# ============================================================
# TRUY VẤN NGHIỆP VỤ
# ============================================================

def lay_danh_sach_phong_ban_quan_ly(nhan_vien_id: int, vai_tro: str = ""):
    """
    Trả về danh sách tên phòng ban mà nhân viên này là Trưởng phòng.
    Với vai trò admin/hr/admin_bcc: trả về TOÀN BỘ phòng ban (để có thể xem/hỗ trợ
    chấm công thay khi cần), có đánh dấu riêng để UI hiển thị khác.
    """
    db = _get_conn()
    c = db.cursor()
    if vai_tro in ("admin", "hr", "admin_bcc"):
        c.execute(
            """SELECT DISTINCT phong_ban_lam_viec FROM nhan_vien
               WHERE trang_thai IN ('DANG_LAM','THU_VIEC')
                 AND phong_ban_lam_viec IS NOT NULL AND phong_ban_lam_viec != ''
               ORDER BY phong_ban_lam_viec"""
        )
        ten_phong_bans = [r[0] for r in c.fetchall()]
    else:
        c.execute(
            "SELECT ten_phong_ban FROM phong_ban WHERE truong_phong_id = %s ORDER BY ten_phong_ban",
            (nhan_vien_id,),
        )
        ten_phong_bans = [r[0] for r in c.fetchall()]
    c.close()
    db.close()
    return ten_phong_bans


def lay_danh_sach_nhan_vien_theo_phong(ten_phong_ban: str):
    """Danh sách nhân viên đang làm việc thuộc 1 phòng ban."""
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
    """Dữ liệu cham_cong hiện có của danh sách NV cho đúng 1 ngày."""
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
    """Dữ liệu bao_com hiện có của danh sách NV cho đúng 1 ngày (Hòn La)."""
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
    """
    Gọi hàm get_cau_hinh_tang_ca_theo_phong() đã có sẵn trong app.py (đã import
    ở nơi gọi module này). Nếu vì lý do nào đó chưa import được, trả về mặc định
    KHÔNG cho phép tăng ca để tránh lỗi vỡ trang.
    """
    try:
        from app import get_cau_hinh_tang_ca_theo_phong  # import trễ, tránh vòng lặp import
        return get_cau_hinh_tang_ca_theo_phong(ten_phong_ban)
    except Exception:
        return {"cho_phep_tang_ca": False}


# ============================================================
# GHI DỮ LIỆU
# ============================================================

def luu_cham_cong_va_bao_com(
    ten_phong_ban: str,
    ngay_cham_cong: date,
    du_lieu_nv: dict,
    is_hon_la: bool,
    nguoi_luu: str,
):
    """
    Ghi dữ liệu chấm công/OT (bảng cham_cong) + báo cơm (bảng bao_com, chỉ Hòn La)
    cho toàn bộ nhân viên trong `du_lieu_nv`.

    du_lieu_nv: dict {
        nhan_vien_id: {
            'ma_cong': str,             # '' | 'x' | 'x/2' | 'CT' ...
            'gio_tang_ca': float,       # số giờ OT ca ngày
            'gio_tang_ca_dem': float,   # số giờ OT ca đêm
            'bao_com_trang_thai': 'BAO' | 'CAT',   # chỉ có khi is_hon_la=True
            'ly_do_cat': str,
        },
        ...
    }

    Trả về: số nhân viên đã lưu chấm công, số nhân viên đã lưu báo cơm.
    """
    db = _get_conn()
    c = db.cursor()
    n_cham_cong = 0
    n_bao_com = 0
    ngay_bao_com = ngay_cham_cong + timedelta(days=1)

    for nv_id, d in du_lieu_nv.items():
        # ---- 1) Ghi chấm công + OT ----
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

        # ---- 2) Ghi báo cơm ngày hôm sau (chỉ Hòn La) ----
        if is_hon_la:
            trang_thai_bc = d.get("bao_com_trang_thai", "BAO")
            an_sang = trang_thai_bc == "BAO"
            an_trua = trang_thai_bc == "BAO"
            an_toi = trang_thai_bc == "BAO"
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
                (nv_id, ngay_bao_com, an_sang, an_trua, an_toi,
                 trang_thai_bc, ly_do_cat, ten_phong_ban, nguoi_luu),
            )
            n_bao_com += 1

    db.commit()
    c.close()
    db.close()
    return n_cham_cong, n_bao_com


# ============================================================
# GIAO DIỆN CHÍNH — DÀNH CHO TRƯỞNG PHÒNG
# ============================================================

def render_tab_thu_cong(nhan_vien_hien_tai: dict, ma_so_thue_hien_tai: str):
    """
    Giao diện Trưởng phòng chấm công + OT + báo cơm ngày hôm sau.
    Gọi hàm này bên trong tab con "🧑‍💼 Trưởng phòng chấm công" (xem hướng dẫn
    tích hợp ở đầu file).
    """
    if not nhan_vien_hien_tai:
        st.warning("⚠️ Không xác định được thông tin nhân viên đang đăng nhập.")
        return

    is_hon_la = (ma_so_thue_hien_tai or "").strip() == HON_LA_MA_SO_THUE
    ensure_cot_truong_phong()
    if is_hon_la:
        ensure_bao_com_table(ma_so_thue_hien_tai)

    vai_tro = nhan_vien_hien_tai.get("vai_tro", "")
    danh_sach_phong = lay_danh_sach_phong_ban_quan_ly(nhan_vien_hien_tai["id"], vai_tro)

    if not danh_sach_phong:
        st.info("ℹ️ Bạn hiện không được gán làm Trưởng phòng của phòng ban nào, "
                 "nên không có quyền chấm công thay nhân viên tại đây.")
        return

    st.markdown("### 🧑‍💼 Chấm công cho nhân viên phòng bạn quản lý")

    col_chon1, col_chon2 = st.columns([1, 2])
    with col_chon1:
        ngay_cham = st.date_input(
            "Ngày cần chấm công", value=date.today(), key="cc_tp_ngay_cham"
        )
    with col_chon2:
        ten_phong_chon = st.selectbox(
            "Phòng ban", danh_sach_phong, key="cc_tp_phong_chon"
        )

    if is_hon_la:
        ngay_bao_com_hien_thi = ngay_cham + timedelta(days=1)
        st.caption(
            f"🍚 Báo cơm áp dụng cho ngày **{ngay_bao_com_hien_thi.strftime('%d/%m/%Y')}** "
            f"(ngày hôm sau của ngày chấm công đã chọn). Mặc định TẤT CẢ được BÁO ĂN — "
            f"chỉ cần bấm **Báo cắt** cho nhân viên nghỉ."
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

    # Header của "bảng"
    if cho_phep_tang_ca and is_hon_la:
        so_cot = [3, 2, 2, 2, 2]
        tieu_de = ["Nhân viên", "Đi làm hôm nay", "Tăng ca (giờ / ca)", "", "🍚 Báo cơm hôm sau"]
    elif cho_phep_tang_ca:
        so_cot = [3, 3, 3]
        tieu_de = ["Nhân viên", "Đi làm hôm nay", "Tăng ca (giờ / ca)"]
    elif is_hon_la:
        so_cot = [3, 3, 3]
        tieu_de = ["Nhân viên", "Đi làm hôm nay", "🍚 Báo cơm hôm sau"]
    else:
        so_cot = [4, 4]
        tieu_de = ["Nhân viên", "Đi làm hôm nay"]

    header_cols = st.columns(so_cot)
    for col, tt in zip(header_cols, tieu_de):
        col.markdown(f"**{tt}**")

    # Dữ liệu tạm giữ trong session_state để giữ lựa chọn khi rerun
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

        # --- Cột tên nhân viên ---
        with cols[cidx]:
            st.markdown(f"**{nv['ho_ten']}**  \n`{nv['ma_nv']}` · {nv.get('chuc_danh_nghe') or ''}")
        cidx += 1

        # --- Cột chấm công nhanh ---
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

        # --- Cột OT (nếu phòng cho phép) ---
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
                st.caption("☀️ Ca ngày · 🌙 Ca đêm (giờ)")
            cidx += 1

        # --- Cột báo cơm (chỉ Hòn La) ---
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
            so_bao = sum(1 for v in st.session_state[key_state].values()
                         if v["bao_com_trang_thai"] == "BAO")
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
    """
    Bảng tổng hợp số suất ăn (sáng/trưa/tối) theo phòng ban cho 1 ngày,
    dùng cho admin_bcc theo dõi hàng ngày. Chỉ hiển thị dữ liệu nếu bảng
    bao_com đã tồn tại (tức tenant hiện tại là Hòn La).
    """
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
