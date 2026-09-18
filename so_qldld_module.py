# -*- coding: utf-8 -*-
"""
Module: Sổ quản lý lao động (theo Nghị định 145/2020/NĐ-CP, mẫu ban hành kèm
Thông tư hướng dẫn của Bộ LĐTBXH).

Gồm 2 phần:
  1. build_so_qldld_excel(...)   -> hàm điền dữ liệu vào file mẫu Excel
  2. render_tab_so_qldld(...)    -> code Streamlit cho tab t4, dán trực tiếp
     vào khối "with t4:" trong menu 📋 BHXH.

CÁCH TÍCH HỢP VÀO FILE APP HIỆN TẠI
------------------------------------
1) Đổi dòng khai báo tab từ:
       t1, t2, t3 = st.tabs(["📊 Tổng quan", "📝 Báo cáo tăng/giảm D02-LT", "📥 Xuất BC trích nộp BH"])
   thành:
       t1, t2, t3, t4 = st.tabs([
           "📊 Tổng quan", "📝 Báo cáo tăng/giảm D02-LT",
           "📥 Xuất BC trích nộp BH", "📔 Sổ quản lý lao động",
       ])

2) Thêm ở cuối file (sau khối "with t3: render_xuat_bao_cao_bhxh(...)"):
       with t4:
           render_tab_so_qldld(st.session_state.db_engine)

3) Copy file mẫu gốc (đính kèm) vào:
       excel_templates/SoQLLD_template_goc.xlsx
   (giữ nguyên tên/khớp với TEMPLATE_SO_QLLD bên dưới, hoặc sửa lại đường dẫn).

4) import các hàm bên dưới vào file app chính:
       from so_qldld_module import build_so_qldld_excel, render_tab_so_qldld
   (hoặc copy thẳng nội dung 2 hàm vào file app nếu app đang gộp 1 file).
"""

import os
from datetime import date

import openpyxl

TEMPLATE_SO_QLLD = "excel_templates/SoQLLD_template_goc.xlsx"

# Dòng đầu tiên có sẵn định dạng (viền + font) trong file mẫu là dòng 7,
# và mẫu có sẵn định dạng cho tới dòng 12 (6 dòng mẫu). Nếu số lao động
# nhiều hơn, ta copy định dạng của dòng 7 xuống các dòng tiếp theo.
FIRST_DATA_ROW = 7
LAST_PREFORMATTED_ROW = 12


def _gioi_tinh_display(gt):
    if not gt:
        return ""
    gt = str(gt).strip().lower()
    if gt in ("nam", "m", "male"):
        return "Nam"
    if gt in ("nữ", "nu", "f", "female"):
        return "Nữ"
    return gt


def _fmt_ngay(d):
    if not d:
        return ""
    if isinstance(d, str):
        return d
    return d.strftime("%d/%m/%Y")


def _copy_row_style(ws, src_row, dst_row, n_cols=24):
    """Copy font/border/fill/alignment từ src_row sang dst_row (dùng khi cần
    thêm dòng vượt quá số dòng đã định dạng sẵn trong mẫu)."""
    from copy import copy as _copy

    for col in range(1, n_cols + 1):
        src_cell = ws.cell(row=src_row, column=col)
        dst_cell = ws.cell(row=dst_row, column=col)
        dst_cell.font = _copy(src_cell.font)
        dst_cell.border = _copy(src_cell.border)
        dst_cell.fill = _copy(src_cell.fill)
        dst_cell.alignment = _copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def build_so_qldld_excel(template_path, output_path, employees, tu_ngay, den_ngay, company_config):
    """
    Điền danh sách lao động vào file mẫu "Sổ quản lý lao động".

    employees: list[dict] - mỗi dict là 1 bản ghi lấy từ bảng nhan_vien, cần các
        khoá (lấy None/"" nếu không có, hàm tự bỏ qua):
        ho_ten, gioi_tinh, ngay_sinh, quoc_tich, thuong_tru, so_cccd,
        trinh_do, bac_trinh_do_nghe, chuc_danh_nghe, loai_hop_dong,
        ngay_vao_lam, thang_bat_dau_bh, luong_bao_hiem, ngay_ket_thuc, ly_do_nghi
    tu_ngay, den_ngay: date - khoảng thời gian thống kê (in vào tiêu đề sổ)
    company_config: dict - {"ten_doanh_nghiep":..., "mst":..., "dia_chi":...}
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb["Sheet1"]

    # ----- Tiêu đề doanh nghiệp -----
    ws["A1"] = f"DOANH NGHIỆP: {company_config.get('ten_doanh_nghiep', '')}"
    ws["A2"] = f"Mã số thuế: {company_config.get('mst', '')}"
    ws["A3"] = f"Địa chỉ: {company_config.get('dia_chi', '')}"
    ws["A4"] = (
        f"SỔ QUẢN LÝ LAO ĐỘNG (Từ ngày {tu_ngay.strftime('%d/%m/%Y')} "
        f"đến ngày {den_ngay.strftime('%d/%m/%Y')})"
    )

    # ----- Điền dữ liệu từng lao động -----
    for idx, nv in enumerate(employees, start=1):
        r = FIRST_DATA_ROW + idx - 1

        if r > LAST_PREFORMATTED_ROW:
            _copy_row_style(ws, FIRST_DATA_ROW, r)

        dang_dong_bh = bool(nv.get("thang_bat_dau_bh"))

        ws.cell(row=r, column=1, value=idx)                                   # STT
        ws.cell(row=r, column=2, value=(nv.get("ho_ten") or "").strip())       # Họ và tên
        ws.cell(row=r, column=3, value=_gioi_tinh_display(nv.get("gioi_tinh")))  # Giới tính
        ws.cell(row=r, column=4, value=_fmt_ngay(nv.get("ngay_sinh")))          # Ngày sinh
        ws.cell(row=r, column=5, value=nv.get("quoc_tich") or "Việt Nam")       # Quốc tịch
        ws.cell(row=r, column=6, value=nv.get("thuong_tru") or "")             # Nơi cư trú
        ws.cell(row=r, column=7, value=nv.get("so_cccd") or "")                # Số CCCD/CMND/hộ chiếu
        ws.cell(row=r, column=8, value=nv.get("trinh_do") or "")               # Trình độ (CMKT)
        ws.cell(row=r, column=9, value=nv.get("bac_trinh_do_nghe") or "")      # Bậc trình độ kỹ năng nghề
        ws.cell(row=r, column=10, value=nv.get("chuc_danh_nghe") or "")        # Vị trí làm việc
        ws.cell(row=r, column=11, value=nv.get("loai_hop_dong") or "")         # Loại HĐLĐ
        ws.cell(row=r, column=12, value=_fmt_ngay(nv.get("ngay_vao_lam")))     # Bắt đầu làm việc
        ws.cell(row=r, column=13, value="x" if dang_dong_bh else "")           # BHXH
        ws.cell(row=r, column=14, value="x" if dang_dong_bh else "")           # BHYT
        ws.cell(row=r, column=15, value="x" if dang_dong_bh else "")           # BHTN
        ws.cell(row=r, column=16, value=nv.get("luong_bao_hiem") or "")        # Tiền lương
        # Cột Q..W (Nâng bậc lương, Số ngày nghỉ, Giờ làm thêm, Hưởng chế độ BHXH,
        # Học nghề đào tạo, Kỷ luật lao động, TNLĐ-BNN): hệ thống hiện chưa lưu các
        # sự kiện này dưới dạng trường riêng -> để trống, cập nhật thủ công theo
        # phát sinh thực tế trong kỳ.
        thoi_diem_cham_dut = ""
        if nv.get("ngay_ket_thuc") and nv.get("ly_do_nghi"):
            thoi_diem_cham_dut = _fmt_ngay(nv.get("ngay_ket_thuc"))
            thoi_diem_cham_dut += f" - {nv.get('ly_do_nghi')}"
        ws.cell(row=r, column=24, value=thoi_diem_cham_dut)                    # Chấm dứt HĐLĐ và lý do

    wb.save(output_path)
    return output_path


def render_tab_so_qldld(db_engine, format_date, company_config, auto_download_excel):
    """Nội dung tab '📔 Sổ quản lý lao động' — dán vào khối `with t4:`.

    LƯU Ý QUAN TRỌNG: hàm này KHÔNG tự import format_date/COMPANY_CONFIG/
    _auto_download_excel từ app.py nữa. Lý do: khi Streamlit chạy app.py làm
    script chính, app.py được nạp dưới tên module "__main__", KHÔNG phải
    "app" -> nếu ta viết `from app import ...` bên trong module này, Python
    sẽ không tìm thấy module "app" đã nạp sẵn, và sẽ NẠP LẠI TOÀN BỘ app.py
    từ đầu như một bản chạy song song -> gây lỗi (vì các lệnh st.* bị gọi
    trùng lặp, ví dụ st.set_page_config gọi 2 lần, hoặc trùng key widget).

    Thay vào đó, 3 tham số này phải được TRUYỀN VÀO từ nơi gọi hàm (app.py),
    nơi chúng vốn đã tồn tại sẵn trong cùng file:
        render_tab_so_qldld(
            st.session_state.db_engine,
            format_date=format_date,
            company_config=COMPANY_CONFIG,
            auto_download_excel=_auto_download_excel,
        )
    """
    import streamlit as st
    import pandas as pd
    import psycopg2.extras

    st.subheader("📔 Sổ quản lý lao động")
    st.caption(
        "Theo Nghị định 145/2020/NĐ-CP - Ghi nhận thông tin lao động đang làm việc "
        "trong khoảng thời gian được chọn."
    )

    col1, col2 = st.columns(2)
    with col1:
        tu_ngay = st.date_input(
            "📅 Từ ngày:",
            value=date(date.today().year, 1, 1),
            key="soqldld_tu_ngay",
        )
    with col2:
        den_ngay = st.date_input(
            "📅 Đến ngày:",
            value=date.today(),
            key="soqldld_den_ngay",
        )

    if tu_ngay > den_ngay:
        st.error("⚠️ 'Từ ngày' phải nhỏ hơn hoặc bằng 'Đến ngày'.")
        return

    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        export_clicked = st.button(
            "📥 XUẤT EXCEL SỔ QUẢN LÝ LAO ĐỘNG",
            type="primary",
            width="stretch",
            use_container_width=True,
            key="soqldld_export_btn",
        )

    st.divider()

    # Lao động "có mặt" trong kỳ: đã vào làm trước/trong kỳ, và (chưa nghỉ việc
    # HOẶC nghỉ việc sau khi kỳ đã bắt đầu) -> đúng nghiệp vụ của Sổ QLLĐ (ghi nhận
    # toàn bộ lao động từng làm việc trong khoảng thời gian, không chỉ lao động
    # đang làm tại thời điểm hiện tại).
    db = db_engine.get_connection()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute(
        """
        SELECT
            nv.ma_nv, nv.ho_ten, nv.gioi_tinh, nv.ngay_sinh, nv.quoc_tich,
            nv.thuong_tru, nv.so_cccd, nv.chuc_danh_nghe, nv.loai_hop_dong,
            nv.ngay_vao_lam, nv.thang_bat_dau_bh, nv.luong_bao_hiem,
            nv.ngay_ket_thuc, nv.ly_do_nghi, nv.trang_thai,
            nv.trinh_do, nv.so_hdld
        FROM nhan_vien nv
        WHERE nv.ngay_vao_lam <= %s
        AND (nv.ngay_ket_thuc IS NULL OR nv.ngay_ket_thuc >= %s)
        AND nv.so_hdld IS NOT NULL AND nv.so_hdld != ''
        ORDER BY nv.ngay_vao_lam ASC
        """,
        (den_ngay, tu_ngay),
    )
    ds_lao_dong = c.fetchall()
    db.close()

    st.markdown(f"### 👥 Danh sách lao động trong kỳ ({len(ds_lao_dong)})")
    if ds_lao_dong:
        df = pd.DataFrame(ds_lao_dong)
        for col in df.columns:
            if "ngay" in col.lower():
                df[col] = df[col].apply(format_date)
        preview_cols = [
            "ma_nv", "ho_ten", "gioi_tinh", "chuc_danh_nghe",
            "loai_hop_dong", "ngay_vao_lam", "ngay_ket_thuc",
        ]
        available_cols = [c for c in preview_cols if c in df.columns]
        df_preview = df[available_cols]
        df_preview.columns = [
            "Mã NV", "Họ tên", "Giới tính", "Vị trí làm việc",
            "Loại HĐLĐ", "Ngày vào làm", "Ngày kết thúc",
        ][: len(available_cols)]
        st.dataframe(df_preview, width="stretch", hide_index=True)
    else:
        st.info("📭 Không có lao động nào trong khoảng thời gian đã chọn.")

    if export_clicked:
        if ds_lao_dong:
            with st.spinner("Đang tạo Sổ quản lý lao động... Vui lòng chờ..."):
                try:
                    filename = (
                        f"SoQLLD_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
                    )
                    build_so_qldld_excel(
                        template_path=TEMPLATE_SO_QLLD,
                        output_path=filename,
                        employees=ds_lao_dong,
                        tu_ngay=tu_ngay,
                        den_ngay=den_ngay,
                        company_config=company_config,
                    )

                    with open(filename, "rb") as f:
                        file_data = f.read()

                    st.success(f"✅ Đã tạo Sổ quản lý lao động thành công! {len(ds_lao_dong)} lao động.")
                    st.cache_data.clear()

                    auto_download_excel(file_data, filename)

                    if os.path.exists(filename):
                        os.remove(filename)

                except Exception as e:
                    st.error(f"❌ Lỗi khi tạo Sổ quản lý lao động: {str(e)}")
                    st.exception(e)
        else:
            st.warning("⚠️ Không có lao động nào trong kỳ để xuất báo cáo!")
