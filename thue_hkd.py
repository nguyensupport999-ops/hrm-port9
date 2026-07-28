# -*- coding: utf-8 -*-
"""
thue_hkd.py
============
Module quản lý Thuế cho Hộ kinh doanh (HKD).
Gồm 3 tab chính:
  - ⚙️ Cấu hình HKD: thông tin hộ, ngành nghề, tỷ lệ thuế, BHXH chủ hộ
  - 📊 Theo dõi Doanh thu & Thuế: nhập DT theo kỳ, tự động tính thuế
  - 📋 Tổng hợp & Cảnh báo: báo cáo lũy kế, nhắc hạn kê khai

Cách dùng trong app.py:
    from thue_hkd import render_thue_hkd
    if menu == "🧾 Thuế HKD":
        render_thue_hkd(db_engine)
"""

import datetime
import streamlit as st

# ═══════════════════════════════════════════════════════════════════
#  HẰNG SỐ
# ═══════════════════════════════════════════════════════════════════

NGUONG_MIEN_THUE = 500_000_000       # 500 triệu
NGUONG_NHOM_2 = 3_000_000_000        # 3 tỷ
NGUONG_NHOM_3 = 50_000_000_000_000   # 50 tỷ — dùng cho phân nhóm, thực tế HKD hiếm khi đạt

# Tỷ lệ thuế theo ngành nghề (tham chiếu Thông tư 40/2021/TT-BTC)
TY_LE_THUE_THEO_NGANH = {
    "THUONG_MAI":   {"ten": "Phân phối, cung cấp hàng hóa",                      "gtgt": 1.0, "tncn": 0.5},
    "DICH_VU":      {"ten": "Dịch vụ, xây dựng không bao thầu NVL",              "gtgt": 5.0, "tncn": 2.0},
    "SAN_XUAT":     {"ten": "Sản xuất, vận tải, dịch vụ gắn hàng hóa, XD có NVL","gtgt": 3.0, "tncn": 1.5},
    "KHAC":         {"ten": "Hoạt động kinh doanh khác",                          "gtgt": 2.0, "tncn": 1.0},
}

THUE_SUAT_LOI_NHUAN = {
    "500TR_3TY":  15.0,
    "3TY_50TY":   17.0,
    "TREN_50TY":  20.0,
}

MUC_THAM_CHIEU_BHXH = 2_340_000  # Mức tham chiếu BHXH 2025-2026 (= lương cơ sở)


# ═══════════════════════════════════════════════════════════════════
#  HÀM TIỆN ÍCH
# ═══════════════════════════════════════════════════════════════════

def _can_edit():
    """Kiểm tra quyền chỉnh sửa — tái sử dụng logic can_edit() từ app.py."""
    role = st.session_state.get("role", "")
    if role in ("xem_toan_bo", "demo_readonly", "viewer"):
        return False
    return True


def _fmt_tien(so):
    """Định dạng số tiền VND: 1,234,567"""
    if so is None:
        return "0"
    return f"{int(so):,}".replace(",", ".")


def _fmt_tien_trieu(so):
    """Hiển thị số tiền dạng triệu: 500.000.000 → 500 triệu"""
    if so is None or so == 0:
        return "0"
    if so >= 1_000_000_000:
        return f"{so / 1_000_000_000:.1f} tỷ"
    if so >= 1_000_000:
        return f"{so / 1_000_000:.0f} triệu"
    return _fmt_tien(so)


def _phan_nhom_doanh_thu(doanh_thu_nam):
    """Phân nhóm HKD theo doanh thu lũy kế năm."""
    if doanh_thu_nam <= NGUONG_MIEN_THUE:
        return "MIEN_THUE"
    elif doanh_thu_nam <= NGUONG_NHOM_2:
        return "500TR_3TY"
    elif doanh_thu_nam <= NGUONG_NHOM_3:
        return "3TY_50TY"
    else:
        return "TREN_50TY"


def _ten_nhom_doanh_thu(nhom):
    MAP = {
        "MIEN_THUE": "≤ 500 triệu — Miễn thuế",
        "500TR_3TY": "500 triệu – 3 tỷ",
        "3TY_50TY":  "3 tỷ – 50 tỷ",
        "TREN_50TY": "> 50 tỷ",
    }
    return MAP.get(nhom, nhom)


def _tinh_thue_ty_le_doanh_thu(doanh_thu, ty_le_gtgt, ty_le_tncn):
    """Tính thuế theo tỷ lệ % trên doanh thu (sau khi trừ 500 triệu)."""
    phan_chiu_thue = max(0, doanh_thu - NGUONG_MIEN_THUE)
    thue_gtgt = round(phan_chiu_thue * ty_le_gtgt / 100)
    thue_tncn = round(phan_chiu_thue * ty_le_tncn / 100)
    return thue_gtgt, thue_tncn


def _tinh_thue_loi_nhuan(doanh_thu, chi_phi, nhom):
    """Tính thuế theo lợi nhuận (DT - CP) × thuế suất."""
    loi_nhuan = max(0, doanh_thu - chi_phi)
    thue_suat = THUE_SUAT_LOI_NHUAN.get(nhom, 15.0)
    # Thuế TNCN = lợi nhuận × thuế suất
    thue_tncn = round(loi_nhuan * thue_suat / 100)
    # Thuế GTGT: vẫn tính theo tỷ lệ trên doanh thu (không đổi)
    return thue_tncn


def _ky_ke_khai_options(nam=None):
    """Tạo danh sách kỳ kê khai cho 1 năm."""
    if nam is None:
        nam = datetime.date.today().year
    return [f"Q1/{nam}", f"Q2/{nam}", f"Q3/{nam}", f"Q4/{nam}", str(nam)]


def _han_ke_khai(ky_thue):
    """Xác định hạn kê khai dựa trên kỳ thuế."""
    if "/" in ky_thue:
        # Kỳ quý: Q1/2026 → hạn 30/4/2026
        quy, nam = ky_thue.split("/")
        nam = int(nam)
        quy_num = int(quy.replace("Q", ""))
        thang_han = quy_num * 3 + 1
        if thang_han > 12:
            thang_han = 1
            nam += 1
        return datetime.date(nam, thang_han, 30 if thang_han in (4, 6, 9, 11) else 31 if thang_han in (1, 3, 5, 7, 8, 10, 12) else 28)
    else:
        # Kỳ năm: 2026 → hạn 31/01/2027
        nam = int(ky_thue)
        return datetime.date(nam + 1, 1, 31)


# ═══════════════════════════════════════════════════════════════════
#  DB HELPERS
# ═══════════════════════════════════════════════════════════════════

def _get_cau_hinh(conn):
    """Lấy cấu hình HKD (1 dòng duy nhất). Trả None nếu chưa có."""
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM cau_hinh_thue_hkd ORDER BY id LIMIT 1")
        cols = [desc[0] for desc in c.description]
        row = c.fetchone()
        if row:
            return dict(zip(cols, row))
    except Exception:
        pass
    return None


def _luu_cau_hinh(conn, data):
    """Insert hoặc Update cấu hình HKD."""
    c = conn.cursor()
    # Kiểm tra đã có chưa
    c.execute("SELECT id FROM cau_hinh_thue_hkd LIMIT 1")
    existing = c.fetchone()

    if existing:
        c.execute("""
            UPDATE cau_hinh_thue_hkd SET
                ten_hkd = %s, dia_chi = %s, so_dkkd = %s, ngay_cap_dkkd = %s,
                co_quan_cap = %s, chu_ho_ten = %s, chu_ho_cccd = %s,
                nganh_nghe = %s, ty_le_thue_gtgt = %s, ty_le_thue_tncn = %s,
                phuong_phap_tinh_thue = %s, ky_ke_khai = %s,
                chu_ho_dong_bhxh = %s, muc_luong_dong_bhxh_chu_ho = %s,
                phuong_thuc_dong_bhxh = %s,
                updated_at = now()
            WHERE id = %s
        """, (
            data["ten_hkd"], data["dia_chi"], data["so_dkkd"], data["ngay_cap_dkkd"],
            data["co_quan_cap"], data["chu_ho_ten"], data["chu_ho_cccd"],
            data["nganh_nghe"], data["ty_le_thue_gtgt"], data["ty_le_thue_tncn"],
            data["phuong_phap_tinh_thue"], data["ky_ke_khai"],
            data["chu_ho_dong_bhxh"], data["muc_luong_dong_bhxh_chu_ho"],
            data["phuong_thuc_dong_bhxh"],
            existing[0],
        ))
    else:
        c.execute("""
            INSERT INTO cau_hinh_thue_hkd (
                ten_hkd, dia_chi, so_dkkd, ngay_cap_dkkd, co_quan_cap,
                chu_ho_ten, chu_ho_cccd,
                nganh_nghe, ty_le_thue_gtgt, ty_le_thue_tncn,
                phuong_phap_tinh_thue, ky_ke_khai,
                chu_ho_dong_bhxh, muc_luong_dong_bhxh_chu_ho, phuong_thuc_dong_bhxh
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            data["ten_hkd"], data["dia_chi"], data["so_dkkd"], data["ngay_cap_dkkd"],
            data["co_quan_cap"], data["chu_ho_ten"], data["chu_ho_cccd"],
            data["nganh_nghe"], data["ty_le_thue_gtgt"], data["ty_le_thue_tncn"],
            data["phuong_phap_tinh_thue"], data["ky_ke_khai"],
            data["chu_ho_dong_bhxh"], data["muc_luong_dong_bhxh_chu_ho"],
            data["phuong_thuc_dong_bhxh"],
        ))
    conn.commit()


def _get_ds_thue(conn, nam=None):
    """Lấy danh sách theo dõi thuế của 1 năm."""
    c = conn.cursor()
    if nam:
        c.execute("""SELECT * FROM theo_doi_thue_hkd
                     WHERE ky_thue LIKE %s OR ky_thue = %s
                     ORDER BY id""", (f"%/{nam}", str(nam)))
    else:
        c.execute("SELECT * FROM theo_doi_thue_hkd ORDER BY id DESC LIMIT 20")
    cols = [desc[0] for desc in c.description]
    rows = c.fetchall()
    return [dict(zip(cols, r)) for r in rows]


def _luu_ky_thue(conn, data):
    """Insert hoặc Update 1 kỳ thuế."""
    c = conn.cursor()
    # Kiểm tra kỳ đã tồn tại chưa
    c.execute("SELECT id FROM theo_doi_thue_hkd WHERE ky_thue = %s", (data["ky_thue"],))
    existing = c.fetchone()

    if existing:
        c.execute("""
            UPDATE theo_doi_thue_hkd SET
                doanh_thu_ky = %s, doanh_thu_luy_ke_nam = %s, chi_phi_ky = %s,
                nhom_doanh_thu = %s, thue_gtgt_phai_nop = %s, thue_tncn_phai_nop = %s,
                trang_thai = %s, han_ke_khai = %s, ngay_ke_khai = %s, ngay_nop = %s,
                ghi_chu = %s, updated_at = now()
            WHERE id = %s
        """, (
            data["doanh_thu_ky"], data["doanh_thu_luy_ke_nam"], data["chi_phi_ky"],
            data["nhom_doanh_thu"], data["thue_gtgt_phai_nop"], data["thue_tncn_phai_nop"],
            data["trang_thai"], data["han_ke_khai"], data.get("ngay_ke_khai"),
            data.get("ngay_nop"), data.get("ghi_chu", ""),
            existing[0],
        ))
    else:
        c.execute("""
            INSERT INTO theo_doi_thue_hkd (
                ky_thue, doanh_thu_ky, doanh_thu_luy_ke_nam, chi_phi_ky,
                nhom_doanh_thu, thue_gtgt_phai_nop, thue_tncn_phai_nop,
                trang_thai, han_ke_khai, ngay_ke_khai, ngay_nop, ghi_chu
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            data["ky_thue"], data["doanh_thu_ky"], data["doanh_thu_luy_ke_nam"],
            data["chi_phi_ky"], data["nhom_doanh_thu"],
            data["thue_gtgt_phai_nop"], data["thue_tncn_phai_nop"],
            data["trang_thai"], data["han_ke_khai"], data.get("ngay_ke_khai"),
            data.get("ngay_nop"), data.get("ghi_chu", ""),
        ))
    conn.commit()


def _xoa_ky_thue(conn, ky_thue):
    """Xóa 1 kỳ thuế."""
    c = conn.cursor()
    c.execute("DELETE FROM theo_doi_thue_hkd WHERE ky_thue = %s", (ky_thue,))
    conn.commit()


# ═══════════════════════════════════════════════════════════════════
#  TAB 1: CẤU HÌNH HKD
# ═══════════════════════════════════════════════════════════════════

def _render_tab_cau_hinh(db_engine):
    """Tab cấu hình thông tin Hộ kinh doanh + ngành nghề + BHXH chủ hộ."""
    conn = db_engine.get_connection()
    try:
        cfg = _get_cau_hinh(conn)
    finally:
        conn.close()

    if cfg is None:
        cfg = {}

    st.subheader("📝 Thông tin Hộ kinh doanh")

    col1, col2 = st.columns(2)
    with col1:
        ten_hkd = st.text_input("Tên HKD", value=cfg.get("ten_hkd", ""), key="hkd_ten")
        so_dkkd = st.text_input("Số ĐKKD", value=cfg.get("so_dkkd", ""), key="hkd_dkkd")
        chu_ho_ten = st.text_input("Họ tên chủ hộ", value=cfg.get("chu_ho_ten", ""), key="hkd_chuten")
    with col2:
        dia_chi = st.text_input("Địa chỉ", value=cfg.get("dia_chi", ""), key="hkd_diachi")
        co_quan_cap = st.text_input("Cơ quan cấp ĐKKD", value=cfg.get("co_quan_cap", ""), key="hkd_coquan")
        chu_ho_cccd = st.text_input("Số CCCD chủ hộ", value=cfg.get("chu_ho_cccd", ""), key="hkd_cccd")

    ngay_cap_val = cfg.get("ngay_cap_dkkd")
    if ngay_cap_val and isinstance(ngay_cap_val, str):
        try:
            ngay_cap_val = datetime.datetime.strptime(ngay_cap_val, "%Y-%m-%d").date()
        except Exception:
            ngay_cap_val = None
    ngay_cap_dkkd = st.date_input("Ngày cấp ĐKKD", value=ngay_cap_val, key="hkd_ngaycap")

    st.divider()
    st.subheader("🏷️ Ngành nghề & Tỷ lệ thuế")

    nganh_nghe_list = list(TY_LE_THUE_THEO_NGANH.keys())
    nganh_hien_tai = cfg.get("nganh_nghe", "THUONG_MAI")
    idx_nganh = nganh_nghe_list.index(nganh_hien_tai) if nganh_hien_tai in nganh_nghe_list else 0

    nganh_nghe = st.selectbox(
        "Ngành nghề chính",
        nganh_nghe_list,
        index=idx_nganh,
        format_func=lambda k: TY_LE_THUE_THEO_NGANH[k]["ten"],
        key="hkd_nganh",
    )

    # Tỷ lệ thuế auto theo ngành nghề (luật định — không cho sửa tay)
    ty_le_mac_dinh = TY_LE_THUE_THEO_NGANH[nganh_nghe]
    ty_le_gtgt = ty_le_mac_dinh["gtgt"]
    ty_le_tncn = ty_le_mac_dinh["tncn"]
    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("Tỷ lệ thuế GTGT", f"{ty_le_gtgt}%")
    with col_b:
        st.metric("Tỷ lệ thuế TNCN", f"{ty_le_tncn}%")
    st.caption("📌 Tỷ lệ thuế theo Thông tư 40/2021/TT-BTC, tự động theo ngành nghề.")

    # Phương pháp tính thuế: HKD 500tr–3tỷ được chọn 1 trong 2; nhóm khác fix theo luật
    phuong_phap_options = ["TY_LE_DOANH_THU", "LOI_NHUAN"]
    phuong_phap_labels = {
        "TY_LE_DOANH_THU": "Theo tỷ lệ % trên doanh thu (không cần chứng từ chi phí)",
        "LOI_NHUAN": "Theo lợi nhuận: (DT − Chi phí) × thuế suất (cần hóa đơn đầu vào)",
    }
    pp_hien_tai = cfg.get("phuong_phap_tinh_thue", "TY_LE_DOANH_THU")
    idx_pp = phuong_phap_options.index(pp_hien_tai) if pp_hien_tai in phuong_phap_options else 0
    phuong_phap = st.selectbox(
        "Phương pháp tính thuế (HKD DT 500tr–3tỷ được chọn)",
        phuong_phap_options,
        index=idx_pp,
        format_func=lambda k: phuong_phap_labels[k],
        key="hkd_phuongphap",
    )

    # Kỳ kê khai: auto theo luật — DT ≤ 500tr → năm, DT > 500tr → quý
    # Lần đầu admin tự chọn dự kiến; sau đó Tab "Theo dõi" sẽ tự điều chỉnh khi có DT thực tế
    ky_options = ["NAM", "QUY"]
    ky_labels = {"NAM": "Kê khai theo năm (DT ≤ 500 triệu)", "QUY": "Kê khai theo quý (DT > 500 triệu)"}
    ky_hien_tai = cfg.get("ky_ke_khai", "NAM")
    idx_ky = ky_options.index(ky_hien_tai) if ky_hien_tai in ky_options else 0
    ky_ke_khai = st.selectbox(
        "Kỳ kê khai (dự kiến — sẽ tự điều chỉnh theo DT thực tế)",
        ky_options,
        index=idx_ky,
        format_func=lambda k: ky_labels[k],
        key="hkd_kykekhai",
    )

    st.divider()
    st.subheader("🏥 BHXH Chủ hộ kinh doanh")
    st.caption("Chủ HKD đóng BHXH (25%) + BHYT (4,5%) = 29,5%. Không đóng BHTN. Mức lương tự chọn.")

    chu_ho_dong_bhxh = st.checkbox(
        "Chủ hộ tham gia BHXH bắt buộc",
        value=cfg.get("chu_ho_dong_bhxh", True),
        key="hkd_bhxh",
    )

    if chu_ho_dong_bhxh:
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            _ml_hien_tai = int(cfg.get("muc_luong_dong_bhxh_chu_ho", MUC_THAM_CHIEU_BHXH))
            _ml_text = st.text_input(
                f"Mức lương đóng BHXH (tối thiểu {_fmt_tien(MUC_THAM_CHIEU_BHXH)}đ)",
                value=_fmt_tien(_ml_hien_tai),
                key="hkd_mucluong",
            )
            # Parse: bỏ dấu chấm phân hàng → số nguyên
            try:
                muc_luong = int(_ml_text.replace(".", "").replace(",", "").strip())
            except ValueError:
                muc_luong = _ml_hien_tai
                st.caption("⚠️ Nhập số, VD: 2.340.000")
            # Validate
            if muc_luong < MUC_THAM_CHIEU_BHXH:
                st.caption(f"⚠️ Tối thiểu {_fmt_tien(MUC_THAM_CHIEU_BHXH)}đ")
                muc_luong = MUC_THAM_CHIEU_BHXH
            elif muc_luong > MUC_THAM_CHIEU_BHXH * 20:
                st.caption(f"⚠️ Tối đa {_fmt_tien(MUC_THAM_CHIEU_BHXH * 20)}đ")
                muc_luong = MUC_THAM_CHIEU_BHXH * 20
        with col_m2:
            pt_options = ["HANG_THANG", "3_THANG", "6_THANG"]
            pt_labels = {"HANG_THANG": "Hàng tháng", "3_THANG": "3 tháng/lần", "6_THANG": "6 tháng/lần"}
            pt_hien_tai = cfg.get("phuong_thuc_dong_bhxh", "HANG_THANG")
            idx_pt = pt_options.index(pt_hien_tai) if pt_hien_tai in pt_options else 0
            phuong_thuc_dong = st.selectbox(
                "Phương thức đóng",
                pt_options,
                index=idx_pt,
                format_func=lambda k: pt_labels[k],
                key="hkd_ptdong",
            )

        # Hiển thị ước tính
        tien_bhxh = round(muc_luong * 0.25)
        tien_bhyt = round(muc_luong * 0.045)
        tong = tien_bhxh + tien_bhyt
        st.info(f"💰 Ước tính: BHXH {_fmt_tien(tien_bhxh)}đ + BHYT {_fmt_tien(tien_bhyt)}đ = **{_fmt_tien(tong)}đ/tháng**")
    else:
        muc_luong = MUC_THAM_CHIEU_BHXH
        phuong_thuc_dong = "HANG_THANG"

    st.divider()

    # Nút Lưu
    if st.button("💾 Lưu cấu hình", disabled=not _can_edit(), type="primary", key="hkd_luu"):
        data = {
            "ten_hkd": ten_hkd,
            "dia_chi": dia_chi,
            "so_dkkd": so_dkkd,
            "ngay_cap_dkkd": ngay_cap_dkkd,
            "co_quan_cap": co_quan_cap,
            "chu_ho_ten": chu_ho_ten,
            "chu_ho_cccd": chu_ho_cccd,
            "nganh_nghe": nganh_nghe,
            "ty_le_thue_gtgt": ty_le_gtgt,
            "ty_le_thue_tncn": ty_le_tncn,
            "phuong_phap_tinh_thue": phuong_phap,
            "ky_ke_khai": ky_ke_khai,
            "chu_ho_dong_bhxh": chu_ho_dong_bhxh,
            "muc_luong_dong_bhxh_chu_ho": muc_luong,
            "phuong_thuc_dong_bhxh": phuong_thuc_dong,
        }
        try:
            conn = db_engine.get_connection()
            _luu_cau_hinh(conn, data)
            conn.close()
            st.success("✅ Đã lưu cấu hình HKD!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Lỗi: {e}")


# ═══════════════════════════════════════════════════════════════════
#  TAB 2: THEO DÕI DOANH THU & THUẾ
# ═══════════════════════════════════════════════════════════════════

def _render_tab_theo_doi(db_engine):
    """Tab nhập doanh thu theo kỳ, tự động tính thuế."""
    # Đọc cấu hình
    conn = db_engine.get_connection()
    try:
        cfg = _get_cau_hinh(conn)
    finally:
        conn.close()

    if not cfg:
        st.warning("⚠️ Chưa có cấu hình HKD. Vui lòng vào tab **⚙️ Cấu hình HKD** để thiết lập trước.")
        return

    nam_hien_tai = datetime.date.today().year
    nam = st.selectbox("Năm", [nam_hien_tai - 1, nam_hien_tai, nam_hien_tai + 1],
                       index=1, key="hkd_nam_theodoi")

    # Đọc danh sách kỳ thuế đã nhập
    conn = db_engine.get_connection()
    try:
        ds_thue = _get_ds_thue(conn, nam)
    finally:
        conn.close()

    ky_da_nhap = {d["ky_thue"] for d in ds_thue}

    st.divider()
    st.subheader("➕ Nhập doanh thu kỳ mới")

    # Tạo danh sách kỳ có thể nhập
    if cfg["ky_ke_khai"] == "QUY":
        ky_chon_list = [f"Q{i}/{nam}" for i in range(1, 5)]
    else:
        ky_chon_list = [str(nam)]

    ky_chua_nhap = [k for k in ky_chon_list if k not in ky_da_nhap]

    if not ky_chua_nhap:
        st.info("✅ Đã nhập đủ các kỳ thuế trong năm này.")
    else:
        ky_thue = st.selectbox("Chọn kỳ thuế", ky_chua_nhap, key="hkd_ky_moi")

        col1, col2 = st.columns(2)
        with col1:
            _dt_text = st.text_input("Doanh thu kỳ này (VNĐ)", value="0", key="hkd_dt_ky")
            try:
                doanh_thu_ky = int(_dt_text.replace(".", "").replace(",", "").strip())
                if doanh_thu_ky < 0:
                    doanh_thu_ky = 0
                # Hiển thị lại dạng có dấu chấm để user thấy
                if doanh_thu_ky > 0:
                    st.caption(f"= {_fmt_tien(doanh_thu_ky)}đ")
            except ValueError:
                doanh_thu_ky = 0
                st.caption("⚠️ Nhập số, VD: 150.000.000")
        with col2:
            _is_loi_nhuan = cfg["phuong_phap_tinh_thue"] == "LOI_NHUAN"
            _cp_text = st.text_input(
                "Chi phí kỳ này (VNĐ)",
                value="0",
                key="hkd_cp_ky",
                disabled=not _is_loi_nhuan,
            )
            if _is_loi_nhuan:
                try:
                    chi_phi_ky = int(_cp_text.replace(".", "").replace(",", "").strip())
                    if chi_phi_ky < 0:
                        chi_phi_ky = 0
                    if chi_phi_ky > 0:
                        st.caption(f"= {_fmt_tien(chi_phi_ky)}đ")
                except ValueError:
                    chi_phi_ky = 0
                    st.caption("⚠️ Nhập số, VD: 80.000.000")
            else:
                chi_phi_ky = 0

        # Tính doanh thu lũy kế
        dt_luy_ke = sum(d.get("doanh_thu_ky", 0) or 0 for d in ds_thue) + doanh_thu_ky

        # Phân nhóm
        nhom = _phan_nhom_doanh_thu(dt_luy_ke)

        # Tính thuế
        if cfg["phuong_phap_tinh_thue"] == "LOI_NHUAN" and nhom != "MIEN_THUE":
            thue_tncn = _tinh_thue_loi_nhuan(doanh_thu_ky, chi_phi_ky, nhom)
            thue_gtgt = round((doanh_thu_ky - NGUONG_MIEN_THUE / 4 if "/" in ky_thue else doanh_thu_ky - NGUONG_MIEN_THUE) * float(cfg["ty_le_thue_gtgt"]) / 100)
            thue_gtgt = max(0, thue_gtgt)
        elif nhom != "MIEN_THUE":
            thue_gtgt, thue_tncn = _tinh_thue_ty_le_doanh_thu(
                doanh_thu_ky, float(cfg["ty_le_thue_gtgt"]), float(cfg["ty_le_thue_tncn"])
            )
        else:
            thue_gtgt, thue_tncn = 0, 0

        han = _han_ke_khai(ky_thue)

        # Hiển thị kết quả tính
        st.markdown(f"""
        | Chỉ tiêu | Giá trị |
        |---|---|
        | Doanh thu lũy kế năm {nam} | **{_fmt_tien(dt_luy_ke)}đ** |
        | Nhóm doanh thu | **{_ten_nhom_doanh_thu(nhom)}** |
        | Thuế GTGT phải nộp | **{_fmt_tien(thue_gtgt)}đ** |
        | Thuế TNCN phải nộp | **{_fmt_tien(thue_tncn)}đ** |
        | **Tổng thuế** | **{_fmt_tien(thue_gtgt + thue_tncn)}đ** |
        | Hạn kê khai | **{han.strftime('%d/%m/%Y')}** |
        """)

        if st.button("💾 Lưu kỳ thuế", disabled=not _can_edit(), type="primary", key="hkd_luu_ky"):
            data = {
                "ky_thue": ky_thue,
                "doanh_thu_ky": doanh_thu_ky,
                "doanh_thu_luy_ke_nam": dt_luy_ke,
                "chi_phi_ky": chi_phi_ky,
                "nhom_doanh_thu": nhom,
                "thue_gtgt_phai_nop": thue_gtgt,
                "thue_tncn_phai_nop": thue_tncn,
                "trang_thai": "CHUA_KE_KHAI",
                "han_ke_khai": han,
                "ngay_ke_khai": None,
                "ngay_nop": None,
                "ghi_chu": "",
            }
            try:
                conn = db_engine.get_connection()
                _luu_ky_thue(conn, data)
                conn.close()
                st.success(f"✅ Đã lưu kỳ thuế {ky_thue}!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Lỗi: {e}")

    # Hiển thị danh sách kỳ đã nhập
    if ds_thue:
        st.divider()
        st.subheader(f"📋 Các kỳ thuế năm {nam}")

        for d in ds_thue:
            trang_thai_icon = {"CHUA_KE_KHAI": "⏳", "DA_KE_KHAI": "📝", "DA_NOP": "✅"}.get(
                d.get("trang_thai", ""), "❓"
            )
            trang_thai_ten = {"CHUA_KE_KHAI": "Chưa kê khai", "DA_KE_KHAI": "Đã kê khai", "DA_NOP": "Đã nộp"}.get(
                d.get("trang_thai", ""), d.get("trang_thai", "")
            )

            with st.expander(
                f"{trang_thai_icon} **{d['ky_thue']}** — DT: {_fmt_tien(d.get('doanh_thu_ky', 0))}đ | "
                f"Thuế: {_fmt_tien((d.get('thue_gtgt_phai_nop', 0) or 0) + (d.get('thue_tncn_phai_nop', 0) or 0))}đ | "
                f"{trang_thai_ten}"
            ):
                st.markdown(f"""
                - **Doanh thu kỳ:** {_fmt_tien(d.get('doanh_thu_ky', 0))}đ
                - **DT lũy kế năm:** {_fmt_tien(d.get('doanh_thu_luy_ke_nam', 0))}đ
                - **Nhóm:** {_ten_nhom_doanh_thu(d.get('nhom_doanh_thu', ''))}
                - **Thuế GTGT:** {_fmt_tien(d.get('thue_gtgt_phai_nop', 0))}đ
                - **Thuế TNCN:** {_fmt_tien(d.get('thue_tncn_phai_nop', 0))}đ
                - **Hạn kê khai:** {d.get('han_ke_khai', '')}
                """)

                # Cập nhật trạng thái
                col_s1, col_s2, col_s3 = st.columns(3)
                with col_s1:
                    if d.get("trang_thai") == "CHUA_KE_KHAI":
                        if st.button("📝 Đánh dấu Đã kê khai", key=f"kk_{d['ky_thue']}",
                                     disabled=not _can_edit()):
                            conn = db_engine.get_connection()
                            d_copy = dict(d)
                            d_copy["trang_thai"] = "DA_KE_KHAI"
                            d_copy["ngay_ke_khai"] = datetime.date.today()
                            _luu_ky_thue(conn, d_copy)
                            conn.close()
                            st.rerun()
                with col_s2:
                    if d.get("trang_thai") in ("CHUA_KE_KHAI", "DA_KE_KHAI"):
                        if st.button("✅ Đánh dấu Đã nộp", key=f"nop_{d['ky_thue']}",
                                     disabled=not _can_edit()):
                            conn = db_engine.get_connection()
                            d_copy = dict(d)
                            d_copy["trang_thai"] = "DA_NOP"
                            d_copy["ngay_nop"] = datetime.date.today()
                            if not d_copy.get("ngay_ke_khai"):
                                d_copy["ngay_ke_khai"] = datetime.date.today()
                            _luu_ky_thue(conn, d_copy)
                            conn.close()
                            st.rerun()
                with col_s3:
                    if st.button("🗑️ Xóa", key=f"xoa_{d['ky_thue']}",
                                 disabled=not _can_edit()):
                        conn = db_engine.get_connection()
                        _xoa_ky_thue(conn, d["ky_thue"])
                        conn.close()
                        st.rerun()


# ═══════════════════════════════════════════════════════════════════
#  TAB 3: TỔNG HỢP & CẢNH BÁO
# ═══════════════════════════════════════════════════════════════════

def _render_tab_tong_hop(db_engine):
    """Tab báo cáo tổng hợp và cảnh báo hạn kê khai."""
    conn = db_engine.get_connection()
    try:
        cfg = _get_cau_hinh(conn)
    finally:
        conn.close()

    if not cfg:
        st.warning("⚠️ Chưa có cấu hình HKD.")
        return

    nam = datetime.date.today().year
    conn = db_engine.get_connection()
    try:
        ds_thue = _get_ds_thue(conn, nam)
    finally:
        conn.close()

    hom_nay = datetime.date.today()

    # ── Tổng hợp năm ──
    st.subheader(f"📊 Tổng hợp năm {nam}")

    tong_dt = sum(d.get("doanh_thu_ky", 0) or 0 for d in ds_thue)
    tong_gtgt = sum(d.get("thue_gtgt_phai_nop", 0) or 0 for d in ds_thue)
    tong_tncn = sum(d.get("thue_tncn_phai_nop", 0) or 0 for d in ds_thue)
    tong_thue = tong_gtgt + tong_tncn

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Doanh thu lũy kế", _fmt_tien_trieu(tong_dt))
    col2.metric("Thuế GTGT", _fmt_tien(tong_gtgt) + "đ")
    col3.metric("Thuế TNCN", _fmt_tien(tong_tncn) + "đ")
    col4.metric("Tổng thuế", _fmt_tien(tong_thue) + "đ")

    nhom = _phan_nhom_doanh_thu(tong_dt)
    st.info(f"📌 Nhóm doanh thu hiện tại: **{_ten_nhom_doanh_thu(nhom)}** | "
            f"Phương pháp: **{cfg.get('phuong_phap_tinh_thue', '')}** | "
            f"Kỳ kê khai: **{cfg.get('ky_ke_khai', '')}**")

    # ── BHXH chủ hộ ──
    if cfg.get("chu_ho_dong_bhxh"):
        st.divider()
        st.subheader("🏥 BHXH Chủ hộ")
        muc_luong = int(cfg.get("muc_luong_dong_bhxh_chu_ho", MUC_THAM_CHIEU_BHXH))
        tien_thang = round(muc_luong * 0.295)
        pt = cfg.get("phuong_thuc_dong_bhxh", "HANG_THANG")
        pt_label = {"HANG_THANG": "hàng tháng", "3_THANG": "3 tháng/lần", "6_THANG": "6 tháng/lần"}.get(pt, pt)
        he_so = {"HANG_THANG": 1, "3_THANG": 3, "6_THANG": 6}.get(pt, 1)
        tien_moi_lan = tien_thang * he_so

        col_b1, col_b2, col_b3 = st.columns(3)
        col_b1.metric("Mức lương đóng", _fmt_tien(muc_luong) + "đ")
        col_b2.metric(f"Số tiền đóng ({pt_label})", _fmt_tien(tien_moi_lan) + "đ")
        col_b3.metric("Đóng cả năm", _fmt_tien(tien_thang * 12) + "đ")

    # ── Cảnh báo hạn kê khai ──
    st.divider()
    st.subheader("⏰ Cảnh báo hạn kê khai")

    canh_bao = []
    for d in ds_thue:
        if d.get("trang_thai") in ("CHUA_KE_KHAI", "DA_KE_KHAI"):
            han = d.get("han_ke_khai")
            if han:
                if isinstance(han, str):
                    try:
                        han = datetime.datetime.strptime(han, "%Y-%m-%d").date()
                    except Exception:
                        continue
                so_ngay_con = (han - hom_nay).days
                canh_bao.append({
                    "ky": d["ky_thue"],
                    "han": han,
                    "so_ngay": so_ngay_con,
                    "trang_thai": d["trang_thai"],
                })

    if not canh_bao:
        st.success("✅ Không có kỳ nào đang chờ kê khai/nộp.")
    else:
        for cb in sorted(canh_bao, key=lambda x: x["so_ngay"]):
            if cb["so_ngay"] < 0:
                st.error(f"🔴 **{cb['ky']}** — QUÁ HẠN {abs(cb['so_ngay'])} ngày! (hạn {cb['han'].strftime('%d/%m/%Y')})")
            elif cb["so_ngay"] <= 7:
                st.warning(f"🟡 **{cb['ky']}** — còn **{cb['so_ngay']} ngày** (hạn {cb['han'].strftime('%d/%m/%Y')})")
            elif cb["so_ngay"] <= 30:
                st.info(f"🔵 **{cb['ky']}** — còn {cb['so_ngay']} ngày (hạn {cb['han'].strftime('%d/%m/%Y')})")
            else:
                st.caption(f"⚪ **{cb['ky']}** — còn {cb['so_ngay']} ngày (hạn {cb['han'].strftime('%d/%m/%Y')})")


# ═══════════════════════════════════════════════════════════════════
#  HÀM CHÍNH — gọi từ app.py
# ═══════════════════════════════════════════════════════════════════

def render_thue_hkd(db_engine):
    """Entry point — gọi từ app.py khi menu == '🧾 Thuế HKD'."""
    st.title("🧾 Thuế Hộ kinh doanh")

    tab1, tab2, tab3 = st.tabs([
        "⚙️ Cấu hình HKD",
        "📊 Theo dõi Doanh thu & Thuế",
        "📋 Tổng hợp & Cảnh báo",
    ])

    with tab1:
        _render_tab_cau_hinh(db_engine)
    with tab2:
        _render_tab_theo_doi(db_engine)
    with tab3:
        _render_tab_tong_hop(db_engine)
