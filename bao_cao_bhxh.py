"""
Module xuất báo cáo trích nộp BHXH — dùng chung cho DN & HKD.
Mẫu: DANH SÁCH LAO ĐỘNG NỘP BẢO HIỂM
"""

import io
import datetime
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


def _safe_float(val):
    """Chuyển giá trị (có thể là str/None) sang float an toàn."""
    if val is None:
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _safe_date_format(val, fmt="%d/%m/%Y"):
    """Format ngày an toàn."""
    if val is None:
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime(fmt)
    if isinstance(val, str) and len(val) >= 10:
        try:
            return datetime.datetime.strptime(val[:10], "%Y-%m-%d").strftime(fmt)
        except Exception:
            return str(val)
    return str(val)


def xuat_bao_cao_trich_nop_bhxh(db_engine, tu_ngay=None, den_ngay=None):
    """Xuất báo cáo trích nộp BHXH.
    Trả về (excel_bytes, so_nv).
    """
    conn = db_engine.get_connection()
    try:
        c = conn.cursor()

        # Lấy danh sách cột thực tế của bảng nhan_vien
        c.execute("""
            SELECT column_name FROM information_schema.columns
            WHERE table_name = 'nhan_vien'
        """)
        ds_cot = {r[0] for r in c.fetchall()}

        # Build SELECT an toàn — chỉ lấy cột tồn tại
        cot_chon = ["ho_ten"]
        cot_chon.append("ngay_bat_dau_bh" if "ngay_bat_dau_bh" in ds_cot else "NULL AS ngay_bat_dau_bh")
        cot_chon.append("luong_bao_hiem")
        cot_chon.append("so_hdld" if "so_hdld" in ds_cot else "NULL AS so_hdld")
        cot_chon.append("ma_bhxh" if "ma_bhxh" in ds_cot else "NULL AS ma_bhxh")
        cot_chon.append("so_cccd" if "so_cccd" in ds_cot else "NULL AS so_cccd")
        cot_chon.append("ngay_cap_cccd" if "ngay_cap_cccd" in ds_cot else "NULL AS ngay_cap_cccd")
        cot_chon.append("chuc_danh_nghe" if "chuc_danh_nghe" in ds_cot else ("chuc_danh" if "chuc_danh" in ds_cot else "NULL AS chuc_danh_nghe"))
        cot_chon.append("noi_dang_ky_kcb" if "noi_dang_ky_kcb" in ds_cot else "NULL AS noi_dang_ky_kcb")
        cot_chon.append("ghi_chu" if "ghi_chu" in ds_cot else "NULL AS ghi_chu")

        cot_sort = "phong_ban_lam_viec" if "phong_ban_lam_viec" in ds_cot else "ho_ten"

        sql = f"""
            SELECT {', '.join(cot_chon)}
            FROM nhan_vien
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
              AND luong_bao_hiem IS NOT NULL
              AND luong_bao_hiem != ''
              AND CAST(luong_bao_hiem AS NUMERIC) > 0
            ORDER BY {cot_sort}, ho_ten
        """
        c.execute(sql)
        col_names = [desc[0] for desc in c.description]
        rows = [dict(zip(col_names, r)) for r in c.fetchall()]
    except Exception as e:
        conn.close()
        raise e
    finally:
        try:
            conn.close()
        except Exception:
            pass

    # Lọc theo khoảng thời gian (nếu có)
    if tu_ngay or den_ngay:
        filtered = []
        for r in rows:
            ngay_bh = r.get("ngay_bat_dau_bh")
            if ngay_bh:
                if isinstance(ngay_bh, str):
                    try:
                        ngay_bh = datetime.datetime.strptime(ngay_bh[:10], "%Y-%m-%d").date()
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
    font_note = Font(name="Times New Roman", size=9, italic=True)
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
    col_widths = {"A": 5, "B": 22, "C": 14, "D": 16, "E": 16, "F": 16, "G": 16,
                  "H": 14, "I": 14, "J": 15, "K": 14, "L": 20, "M": 16, "N": 16}
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
    if ky_label:
        ws.merge_cells("A2:N2")
        ws["A2"] = ky_label
        ws["A2"].font = font_normal
        ws["A2"].alignment = align_center

    # Header
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
    tong_luong = 0
    tong_32 = 0
    tong_dn = 0
    tong_nld = 0

    for idx, r in enumerate(rows, 1):
        luong_bh = _safe_float(r.get("luong_bao_hiem", 0))
        so_tien_32 = round(luong_bh * 0.32)
        dn_215 = round(luong_bh * 0.215)
        nld_105 = round(luong_bh * 0.105)

        tong_luong += luong_bh
        tong_32 += so_tien_32
        tong_dn += dn_215
        tong_nld += nld_105

        ngay_bh_str = _safe_date_format(r.get("ngay_bat_dau_bh"), "%m/%Y")
        ngay_cap_str = _safe_date_format(r.get("ngay_cap_cccd"), "%d/%m/%Y")

        values = [
            idx,
            r.get("ho_ten", ""),
            ngay_bh_str,
            luong_bh,
            so_tien_32,
            dn_215,
            nld_105,
            r.get("so_hdld", "") or "",
            r.get("ma_bhxh", "") or "",
            r.get("so_cccd", "") or "",
            ngay_cap_str,
            r.get("chuc_danh_nghe", "") or "",
            r.get("noi_dang_ky_kcb", "") or "",
            r.get("ghi_chu", "") or "",
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
    ws.merge_cells(f"A{data_row}:C{data_row}")
    cell_tong = ws.cell(row=data_row, column=1, value="TỔNG CỘNG")
    cell_tong.font = font_header
    cell_tong.alignment = align_center
    cell_tong.border = thin_border
    cell_tong.fill = fill_total
    # Border cho B, C (đã merge)
    for mc in [2, 3]:
        ws.cell(row=data_row, column=mc).border = thin_border

    for col_idx, val in [(4, tong_luong), (5, tong_32), (6, tong_dn), (7, tong_nld)]:
        cell = ws.cell(row=data_row, column=col_idx, value=val)
        cell.font = font_header
        cell.alignment = align_right
        cell.border = thin_border
        cell.number_format = '#,##0'
        cell.fill = fill_total

    for col_idx in range(8, 15):
        ws.cell(row=data_row, column=col_idx).border = thin_border
        ws.cell(row=data_row, column=col_idx).fill = fill_total

    # Ghi chú
    data_row += 2
    ws.merge_cells(f"A{data_row}:N{data_row}")
    ws[f"A{data_row}"] = "Ghi chú: DN phải nộp 21,5% = BHXH 14% + BHYT 3% + BHTN 1% + BHTNLĐ-BNN 0,5% + Quỹ HT 3%"
    ws[f"A{data_row}"].font = font_note
    data_row += 1
    ws.merge_cells(f"A{data_row}:N{data_row}")
    ws[f"A{data_row}"] = "         NLĐ phải nộp 10,5% = BHXH 8% + BHYT 1,5% + BHTN 1%"
    ws[f"A{data_row}"].font = font_note

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
    """UI xuất báo cáo trích nộp BHXH."""
    st.subheader("📥 Xuất báo cáo trích nộp BHXH")
    st.caption("Mẫu: Danh sách lao động nộp bảo hiểm (32% = DN 21,5% + NLĐ 10,5%)")

    col1, col2 = st.columns(2)
    with col1:
        tu_ngay = st.date_input("Từ ngày (tháng bắt đầu nộp BH)",
                                value=datetime.date(datetime.date.today().year, 1, 1),
                                key="bhxh_bc_tu")
    with col2:
        den_ngay = st.date_input("Đến ngày",
                                 value=datetime.date.today(),
                                 key="bhxh_bc_den")

    loc_tat_ca = st.checkbox("Xuất tất cả NV đang đóng BH (bỏ lọc ngày)", value=True, key="bhxh_bc_all")

    if st.button("📥 Xuất báo cáo", type="primary", key="btn_xuat_bc_bhxh"):
        try:
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
        except Exception as e:
            st.error(f"❌ Lỗi: {e}")
