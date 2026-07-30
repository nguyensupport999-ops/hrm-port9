# ═══════════════════════════════════════════════════════════════════
# PHASE 5 — 2 yêu cầu:
# (A) Báo cáo trích nộp BHXH cho DN + HKD (Excel)
# (B) Gợi ý lương BH theo chức danh khi nhập NV
# ═══════════════════════════════════════════════════════════════════


# ─────────────────────────────────────────────────────────────────
# (A) BÁO CÁO TRÍCH NỘP BHXH — Dùng chung cho DN & HKD
# ─────────────────────────────────────────────────────────────────
#
# Đặt hàm này trong app.py hoặc tạo file riêng (vd: bao_cao_bhxh.py)
# rồi import vào app.py. Em đề xuất đặt trong app.py, khu vực tab BHXH.
#
# Mẫu theo ảnh: DANH SÁCH LAO ĐỘNG NỘP BẢO HIỂM
# Cột: STT | Họ và Tên | Tháng bắt đầu nộp BH | Tiền lương nộp BH |
#       Số tiền BH phải nộp/tháng 32% | DN phải nộp 21,5% | NLĐ phải nộp 10,5% |
#       Số HĐLĐ | Mã số BH | CCCD | Ngày cấp CCCD | Chức vụ | Nơi ĐK KCB | Ghi chú
# ─────────────────────────────────────────────────────────────────

import io
import datetime
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


def xuat_bao_cao_trich_nop_bhxh(db_engine, tu_ngay=None, den_ngay=None):
    """Xuất báo cáo trích nộp BHXH cho tất cả NV đang tham gia BH.
    Dùng chung cho cả DN và HKD.
    tu_ngay, den_ngay: datetime.date — lọc theo khoảng thời gian (nếu cần)
    Trả về bytes Excel.
    """
    conn = db_engine.get_connection()
    try:
        c = conn.cursor()
        # Lấy danh sách NV đang làm + thử việc, có lương BH > 0
        c.execute("""
            SELECT
                ho_ten,
                ngay_bat_dau_bh,
                luong_bao_hiem,
                so_hdld,
                ma_bhxh,
                so_cccd,
                ngay_cap_cccd,
                chuc_danh_nghe,
                noi_dang_ky_kcb,
                ghi_chu,
                phong_ban_lam_viec
            FROM nhan_vien
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
              AND luong_bao_hiem IS NOT NULL
              AND luong_bao_hiem != ''
              AND CAST(luong_bao_hiem AS NUMERIC) > 0
            ORDER BY phong_ban_lam_viec, ho_ten
        """)
        cols = [desc[0] for desc in c.description]
        rows = [dict(zip(cols, r)) for r in c.fetchall()]
    finally:
        conn.close()

    # Lọc theo khoảng thời gian (nếu có)
    if tu_ngay or den_ngay:
        filtered = []
        for r in rows:
            ngay_bh = r.get("ngay_bat_dau_bh")
            if ngay_bh:
                if isinstance(ngay_bh, str):
                    try:
                        ngay_bh = datetime.datetime.strptime(ngay_bh, "%Y-%m-%d").date()
                    except Exception:
                        ngay_bh = None
                if ngay_bh:
                    if tu_ngay and ngay_bh < tu_ngay:
                        continue
                    if den_ngay and ngay_bh > den_ngay:
                        continue
            filtered.append(r)
        rows = filtered

    # ── Tạo Excel ──
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DS Nop BH"

    # Styles
    font_title = Font(name="Times New Roman", size=13, bold=True)
    font_header = Font(name="Times New Roman", size=10, bold=True)
    font_normal = Font(name="Times New Roman", size=10)
    font_header_white = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    fill_header = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    fill_total = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    # Column widths
    col_widths = {
        "A": 5, "B": 22, "C": 14, "D": 16, "E": 16, "F": 16, "G": 16,
        "H": 14, "I": 14, "J": 15, "K": 14, "L": 18, "M": 16, "N": 16
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # Title
    ws.merge_cells("A1:N1")
    ws["A1"] = "DANH SÁCH LAO ĐỘNG NỘP BẢO HIỂM"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ky_label = ""
    if tu_ngay and den_ngay:
        ky_label = f"Kỳ: từ {tu_ngay.strftime('%d/%m/%Y')} đến {den_ngay.strftime('%d/%m/%Y')}"
    elif tu_ngay:
        ky_label = f"Từ ngày: {tu_ngay.strftime('%d/%m/%Y')}"
    elif den_ngay:
        ky_label = f"Đến ngày: {den_ngay.strftime('%d/%m/%Y')}"
    if ky_label:
        ws.merge_cells("A2:N2")
        ws["A2"] = ky_label
        ws["A2"].font = font_normal
        ws["A2"].alignment = align_center

    # Header row
    headers = [
        "STT", "HỌ VÀ TÊN", "THÁNG\nBẮT ĐẦU\nNỘP BH", "TIỀN\nLƯƠNG\nNỘP BH",
        "SỐ TIỀN\nBH PHẢI\nNỘP/THÁNG\n32%",
        "DN PHẢI\nNỘP\n21,5%", "NLĐ PHẢI\nNỘP\n10,5%",
        "SỐ HĐLĐ", "MÃ SỐ BH", "CCCD", "NGÀY CẤP\nCCCD",
        "CHỨC VỤ", "NƠI ĐK\nKCB", "GHI CHÚ"
    ]

    header_row = 4
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = font_header_white
        cell.alignment = align_center
        cell.border = thin_border
        cell.fill = fill_header
    ws.row_dimensions[header_row].height = 50

    # Data rows
    data_row = header_row + 1
    tong_luong_bh = 0
    tong_32 = 0
    tong_dn = 0
    tong_nld = 0

    for idx, r in enumerate(rows, 1):
        luong_bh = float(r.get("luong_bao_hiem", 0) or 0)
        so_tien_32 = round(luong_bh * 0.32)
        dn_215 = round(luong_bh * 0.215)
        nld_105 = round(luong_bh * 0.105)

        tong_luong_bh += luong_bh
        tong_32 += so_tien_32
        tong_dn += dn_215
        tong_nld += nld_105

        ngay_bh = r.get("ngay_bat_dau_bh", "")
        if isinstance(ngay_bh, datetime.date):
            ngay_bh = ngay_bh.strftime("%m/%Y")
        elif isinstance(ngay_bh, str) and len(ngay_bh) >= 7:
            try:
                ngay_bh = datetime.datetime.strptime(ngay_bh[:10], "%Y-%m-%d").strftime("%m/%Y")
            except Exception:
                pass

        ngay_cap = r.get("ngay_cap_cccd", "")
        if isinstance(ngay_cap, datetime.date):
            ngay_cap = ngay_cap.strftime("%d/%m/%Y")
        elif isinstance(ngay_cap, str) and len(ngay_cap) >= 10:
            try:
                ngay_cap = datetime.datetime.strptime(ngay_cap[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                pass

        values = [
            idx,
            r.get("ho_ten", ""),
            ngay_bh,
            luong_bh,
            so_tien_32,
            dn_215,
            nld_105,
            r.get("so_hdld", ""),
            r.get("ma_bhxh", ""),
            r.get("so_cccd", ""),
            ngay_cap,
            r.get("chuc_danh_nghe", ""),
            r.get("noi_dang_ky_kcb", ""),
            r.get("ghi_chu", ""),
        ]

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=data_row, column=col_idx, value=val)
            cell.font = font_normal
            cell.border = thin_border
            if col_idx == 1:
                cell.alignment = align_center
            elif col_idx in (4, 5, 6, 7):
                cell.alignment = align_right
                cell.number_format = '#,##0'
            elif col_idx in (3, 8, 9, 10, 11):
                cell.alignment = align_center
            else:
                cell.alignment = align_left

        data_row += 1

    # Dòng tổng
    ws.cell(row=data_row, column=1).border = thin_border
    ws.merge_cells(f"A{data_row}:C{data_row}")
    cell_tong = ws.cell(row=data_row, column=1, value="TỔNG CỘNG")
    cell_tong.font = font_header
    cell_tong.alignment = align_center
    cell_tong.border = thin_border
    cell_tong.fill = fill_total

    for col_idx, val in [(4, tong_luong_bh), (5, tong_32), (6, tong_dn), (7, tong_nld)]:
        cell = ws.cell(row=data_row, column=col_idx, value=val)
        cell.font = font_header
        cell.alignment = align_right
        cell.border = thin_border
        cell.number_format = '#,##0'
        cell.fill = fill_total

    for col_idx in range(8, 15):
        ws.cell(row=data_row, column=col_idx).border = thin_border
        ws.cell(row=data_row, column=col_idx).fill = fill_total

    # Dòng ghi chú
    data_row += 2
    ws.merge_cells(f"A{data_row}:N{data_row}")
    ws[f"A{data_row}"] = "Ghi chú: DN phải nộp 21,5% = BHXH 14% + BHYT 3% + BHTN 1% + BHTNLĐ-BNN 0,5% + Quỹ HT 3%"
    ws[f"A{data_row}"].font = Font(name="Times New Roman", size=9, italic=True)
    data_row += 1
    ws.merge_cells(f"A{data_row}:N{data_row}")
    ws[f"A{data_row}"] = "         NLĐ phải nộp 10,5% = BHXH 8% + BHYT 1,5% + BHTN 1%"
    ws[f"A{data_row}"].font = Font(name="Times New Roman", size=9, italic=True)

    # Ký tên
    data_row += 2
    ws.merge_cells(f"I{data_row}:N{data_row}")
    ws[f"I{data_row}"] = f"Ngày ..... tháng ..... năm {datetime.date.today().year}"
    ws[f"I{data_row}"].font = font_normal
    ws[f"I{data_row}"].alignment = align_center
    data_row += 1
    ws.merge_cells(f"A{data_row}:D{data_row}")
    ws[f"A{data_row}"] = "KẾ TOÁN"
    ws[f"A{data_row}"].font = font_header
    ws[f"A{data_row}"].alignment = align_center
    ws.merge_cells(f"I{data_row}:N{data_row}")
    ws[f"I{data_row}"] = "GIÁM ĐỐC"
    ws[f"I{data_row}"].font = font_header
    ws[f"I{data_row}"].alignment = align_center

    # Xuất bytes
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue(), len(rows)


def render_xuat_bao_cao_bhxh(db_engine):
    """UI xuất báo cáo trích nộp BHXH — thêm vào tab BHXH."""
    st.subheader("📥 Xuất báo cáo trích nộp BHXH")

    col1, col2 = st.columns(2)
    with col1:
        tu_ngay = st.date_input("Từ ngày (tháng bắt đầu nộp BH)",
                                value=datetime.date(datetime.date.today().year, 1, 1),
                                key="bhxh_tu_ngay")
    with col2:
        den_ngay = st.date_input("Đến ngày",
                                 value=datetime.date.today(),
                                 key="bhxh_den_ngay")

    loc_tat_ca = st.checkbox("Xuất tất cả NV đang đóng BH (bỏ lọc ngày)", value=True, key="bhxh_tat_ca")

    if st.button("📥 Xuất báo cáo BHXH", type="primary", key="btn_xuat_bhxh"):
        if loc_tat_ca:
            excel_bytes, so_nv = xuat_bao_cao_trich_nop_bhxh(db_engine)
        else:
            excel_bytes, so_nv = xuat_bao_cao_trich_nop_bhxh(db_engine, tu_ngay, den_ngay)

        if so_nv == 0:
            st.warning("⚠️ Không có NV nào thỏa điều kiện.")
        else:
            ten_file = f"DS_Nop_BH_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                label=f"⬇️ Tải {ten_file} ({so_nv} người)",
                data=excel_bytes,
                file_name=ten_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success(f"✅ Đã tạo báo cáo cho {so_nv} NV!")
