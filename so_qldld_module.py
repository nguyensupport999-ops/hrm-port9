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
import openpyxl.styles

TEMPLATE_SO_QLLD = "excel_templates/SoQLLD_template_goc.xlsx"

# Dòng đầu tiên có sẵn định dạng (viền + font) trong file mẫu là dòng 7,
# và mẫu có sẵn định dạng cho tới dòng 12 (6 dòng mẫu). Nếu số lao động
# nhiều hơn, ta copy định dạng của dòng 7 xuống các dòng tiếp theo.
FIRST_DATA_ROW = 7
# Không còn dùng để rẽ nhánh nữa (từ khi có chèn dòng tiêu đề nhóm, mọi dòng
# dữ liệu đều được áp style thống nhất từ _capture_template_row_style/
# _apply_template_row_style) - giữ lại chỉ để tham khảo lịch sử mẫu gốc.
LAST_PREFORMATTED_ROW = 12

# Dòng tiêu đề cột (header) NGAY PHÍA TRÊN vùng dữ liệu. Đây là GIẢ ĐỊNH
# (mẫu gốc không có trong repo để kiểm tra chính xác) -> nếu dòng tiêu đề
# thật trong file mẫu nằm ở vị trí khác FIRST_DATA_ROW - 1, hãy sửa lại
# hằng số này cho khớp, nếu không tiêu đề "Mã NV" sẽ bị ghi sai dòng.
HEADER_ROW = FIRST_DATA_ROW - 1

# Bảng gốc theo mẫu Nghị định 145 có 24 cột (A..X). Ta bổ sung thêm 1 cột
# cuối bảng "Mã NV" -> cột 25 (Y).
N_COLS_GOC = 24
MA_NV_COL = 25
N_COLS = MA_NV_COL

SALARY_COL = 16   # Cột P - Tiền lương
TERMINATION_COL = 24  # Cột X - Chấm dứt HĐLĐ và lý do

# ----- Thứ tự nhóm bắt buộc theo yêu cầu -----
GROUP_LABELS = [
    "Không xác định thời hạn",
    "Hợp đồng có thời hạn",
    "Thử việc",
    "Khác",                              # nhóm dự phòng, chỉ xuất hiện nếu có
    "Đã chấm dứt hợp đồng lao động",      # luôn ở cuối cùng
]
GRP_KHONG_XD = 0
GRP_CO_THOI_HAN = 1
GRP_THU_VIEC = 2
GRP_KHAC = 3
GRP_CHAM_DUT = 4


def _phan_nhom(nv):
    """Xác định nhóm của 1 lao động theo đúng thứ tự yêu cầu:
    1) Không xác định thời hạn
    2) Hợp đồng có thời hạn  (DB lưu là "Xác định thời hạn", kể cả biến thể
       "Xác định thời hạn 12/24/36 tháng")
    3) Thử việc
    4) (dự phòng) Khác - loai_hop_dong không khớp mẫu nào ở trên
    5) Đã chấm dứt hợp đồng lao động - ĐƯỢC ƯU TIÊN CAO NHẤT: bất kể loại
       hợp đồng gì, hễ đã có ngày kết thúc thì luôn xếp vào nhóm cuối này.
    """
    if nv.get("ngay_ket_thuc"):
        return GRP_CHAM_DUT
    lhd = (nv.get("loai_hop_dong") or "").strip().lower()
    if lhd == "không xác định thời hạn":
        return GRP_KHONG_XD
    if "xác định thời hạn" in lhd:  # khớp cả "xác định thời hạn 12 tháng"...
        return GRP_CO_THOI_HAN
    if lhd == "thử việc":
        return GRP_THU_VIEC
    return GRP_KHAC


def _write_group_header(ws, row, label, n_cols=N_COLS):
    """Ghi 1 dòng tiêu đề nhóm, merge từ cột A đến cột cuối bảng, in đậm,
    căn giữa, có nền xám nhạt để phân biệt với dòng dữ liệu."""
    from openpyxl.styles import Alignment, Font, PatternFill

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    cell = ws.cell(row=row, column=1, value=label)
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    ws.row_dimensions[row].height = 20


def _chars_per_line(ws, col_idx, fallback=28):
    """Ước lượng số ký tự vừa 1 dòng trong cột, dựa vào độ rộng cột đã đặt
    trong file mẫu (nếu có), dùng để tính chiều cao dòng khi wrap text."""
    from openpyxl.utils import get_column_letter

    dim = ws.column_dimensions.get(get_column_letter(col_idx))
    if dim and dim.width:
        return max(8, int(dim.width) - 2)
    return fallback


def _autosize_row_for_wrap(ws, row, col, text, line_height=15, min_height=15):
    """Tăng chiều cao dòng (nếu cần) để text wrap trong 1 ô hiển thị đủ,
    thay cho việc merge nhiều dòng (xem giải thích ở cuối phản hồi)."""
    if not text:
        return
    chars_per_line = _chars_per_line(ws, col)
    n_lines = max(1, -(-len(str(text)) // chars_per_line))  # ceil division
    needed = max(min_height, n_lines * line_height)
    current = ws.row_dimensions[row].height or min_height
    if needed > current:
        ws.row_dimensions[row].height = needed


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


def _capture_template_row_style(ws, src_row, n_cols=N_COLS_GOC):
    """Chụp (snapshot) style của dòng mẫu (font/border/fill/alignment/
    number_format + chiều cao dòng) vào bộ nhớ.

    QUAN TRỌNG: phải gọi hàm này 1 LẦN DUY NHẤT, TRƯỚC KHI ghi bất cứ nội
    dung nào (kể cả tiêu đề nhóm) đè lên dòng mẫu (FIRST_DATA_ROW). Vì việc
    phân nhóm có thể khiến chính dòng FIRST_DATA_ROW trở thành dòng tiêu đề
    nhóm đầu tiên (nếu nhóm đó có lao động) -> nếu không chụp trước, các
    dòng dữ liệu phía sau sẽ vô tình copy nhầm định dạng "tiêu đề nhóm"
    (in đậm, nền xám, merge) thay vì định dạng dòng dữ liệu bình thường.
    """
    from copy import copy as _copy

    styles = []
    for col in range(1, n_cols + 1):
        src_cell = ws.cell(row=src_row, column=col)
        styles.append({
            "font": _copy(src_cell.font),
            "border": _copy(src_cell.border),
            "fill": _copy(src_cell.fill),
            "alignment": _copy(src_cell.alignment),
            "number_format": src_cell.number_format,
        })
    row_height = ws.row_dimensions[src_row].height
    return {"cols": styles, "row_height": row_height}


def _apply_template_row_style(ws, dst_row, template_style, n_cols=N_COLS_GOC):
    """Áp style đã chụp (từ _capture_template_row_style) vào dst_row. Cột
    thứ N_COLS_GOC + 1 trở đi (vd. cột "Mã NV" mới thêm) không có style mẫu
    sẵn -> lấy tạm style của cột cuối cùng trong mẫu (n_cols) cho đồng bộ
    viền/font với các cột khác trên cùng dòng."""
    cols_style = template_style["cols"]
    last_style = cols_style[-1] if cols_style else None
    for col in range(1, max(n_cols, N_COLS) + 1):
        style = cols_style[col - 1] if col <= len(cols_style) else last_style
        if style is None:
            continue
        dst_cell = ws.cell(row=dst_row, column=col)
        dst_cell.font = style["font"]
        dst_cell.border = style["border"]
        dst_cell.fill = style["fill"]
        dst_cell.alignment = style["alignment"]
        dst_cell.number_format = style["number_format"]
    if template_style.get("row_height"):
        ws.row_dimensions[dst_row].height = template_style["row_height"]


def build_so_qldld_excel(template_path, output_path, employees, tu_ngay, den_ngay, company_config):
    """
    Điền danh sách lao động vào file mẫu "Sổ quản lý lao động".

    employees: list[dict] - mỗi dict là 1 bản ghi lấy từ bảng nhan_vien, cần các
        khoá (lấy None/"" nếu không có, hàm tự bỏ qua):
        ma_nv, ho_ten, gioi_tinh, ngay_sinh, quoc_tich, thuong_tru, so_cccd,
        trinh_do, bac_trinh_do_nghe, chuc_danh_nghe, loai_hop_dong,
        ngay_vao_lam, thang_bat_dau_bh, luong_bao_hiem, ngay_ket_thuc, ly_do_nghi
    tu_ngay, den_ngay: date - khoảng thời gian thống kê (in vào tiêu đề sổ)
    company_config: dict - {"ten_doanh_nghiep":..., "mst":..., "dia_chi":...}

    Danh sách được TỰ ĐỘNG PHÂN NHÓM và in theo đúng thứ tự:
      1) Không xác định thời hạn
      2) Hợp đồng có thời hạn
      3) Thử việc
      4) (nếu có phát sinh) Khác
      5) Đã chấm dứt hợp đồng lao động - luôn ở cuối, bất kể loại HĐLĐ gì.
    Mỗi nhóm có 1 dòng tiêu đề (merge toàn bảng, in đậm). Nhóm không có lao
    động nào thì không in dòng tiêu đề của nhóm đó. STT đánh số liên tục
    xuyên suốt toàn bộ sổ (không reset lại theo từng nhóm).
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

    # ----- Tiêu đề cột bổ sung "Mã NV" (xem ghi chú tại hằng số HEADER_ROW) -----
    ws.cell(row=HEADER_ROW, column=MA_NV_COL, value="Mã NV")

    # ----- Chụp style dòng mẫu TRƯỚC khi ghi đè bất cứ nội dung nào -----
    template_style = _capture_template_row_style(ws, FIRST_DATA_ROW)

    # ----- Phân nhóm lao động theo đúng thứ tự yêu cầu -----
    groups = {i: [] for i in range(len(GROUP_LABELS))}
    for nv in employees:
        groups[_phan_nhom(nv)].append(nv)

    r = FIRST_DATA_ROW
    stt = 1
    for g_idx, g_label in enumerate(GROUP_LABELS):
        ds_nhom = groups.get(g_idx) or []
        if not ds_nhom:
            continue  # nhóm rỗng -> không in dòng tiêu đề nhóm

        # ----- Dòng tiêu đề nhóm -----
        _write_group_header(ws, r, g_label)
        r += 1

        # ----- Các dòng lao động trong nhóm -----
        for nv in ds_nhom:
            _apply_template_row_style(ws, r, template_style)

            dang_dong_bh = bool(nv.get("thang_bat_dau_bh"))

            ws.cell(row=r, column=1, value=stt)                                   # STT
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

            # Cột P - Tiền lương: định dạng number, phân tách hàng nghìn.
            luong_cell = ws.cell(row=r, column=SALARY_COL, value=nv.get("luong_bao_hiem") or None)
            luong_cell.number_format = "#,##0"

            # Cột Q..W (Nâng bậc lương, Số ngày nghỉ, Giờ làm thêm, Hưởng chế độ BHXH,
            # Học nghề đào tạo, Kỷ luật lao động, TNLĐ-BNN): hệ thống hiện chưa lưu các
            # sự kiện này dưới dạng trường riêng -> để trống, cập nhật thủ công theo
            # phát sinh thực tế trong kỳ.
            thoi_diem_cham_dut = ""
            if nv.get("ngay_ket_thuc") and nv.get("ly_do_nghi"):
                thoi_diem_cham_dut = _fmt_ngay(nv.get("ngay_ket_thuc"))
                thoi_diem_cham_dut += f" - {nv.get('ly_do_nghi')}"
            # Cột X - Chấm dứt HĐLĐ và lý do: căn giữa ngang/dọc, tự xuống dòng
            # (wrap text), tự giãn chiều cao dòng theo độ dài nội dung.
            cham_dut_cell = ws.cell(row=r, column=TERMINATION_COL, value=thoi_diem_cham_dut)
            cham_dut_cell.alignment = openpyxl.styles.Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            _autosize_row_for_wrap(ws, r, TERMINATION_COL, thoi_diem_cham_dut)

            # Cột cuối bảng - Mã NV
            ws.cell(row=r, column=MA_NV_COL, value=nv.get("ma_nv") or "")

            stt += 1
            r += 1

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
