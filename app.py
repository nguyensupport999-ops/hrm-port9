'''Tóm lại bài học rút ra từ ca này: khi HTML nằm trong components.html (iframe), việc giao tiếp với trang Streamlit cha luôn phải dùng window.top thay vì window.parent — đặc biệt trên Streamlit Cloud nơi có thể có nhiều tầng iframe lồng nhau. Và không bao giờ dùng replaceState rồi lại location.href cùng lúc vì chúng triệt tiêu nhau.
Nếu sau này cần thêm tính năng hay gặp bug mới, cứ ping lại nhé!
'''
import streamlit as st
import psycopg2
import psycopg2.extras
from datetime import datetime, date, timedelta
import calendar
import random
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import tempfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import requests
from openpyxl.styles import Font, Alignment
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import qrcode
from io import BytesIO
import os
import sys
import subprocess
import pathlib
import streamlit.components.v1 as components
import urllib.parse
import re
import json
import unicodedata
import control_plane
from control_plane import DatabaseEngine, resolve_tenant
import bcrypt
import chat_utils
import photo_card_gender
import base64
import mimetypes
from io import BytesIO

# Import config - ưu tiên config.py (local), fallback to config_template (cloud)
try:
    from config import COMPANY_CONFIG, BHXH_CONFIG, EMAIL_CONFIG, TELEGRAM_CONFIG, USERS
    print("Using local config.py")
except ImportError:
    from config_template import COMPANY_CONFIG, BHXH_CONFIG, EMAIL_CONFIG, TELEGRAM_CONFIG, USERS
    print("Using config_template.py")

@st.cache_data(ttl=3600, show_spinner=False)
def get_avatar_bytes_cached(storage_path: str) -> bytes:
    """Tải ảnh avatar có cache 1 giờ"""
    if not storage_path:
        return None
    try:
        sb = get_supabase_storage()
        if not sb:
            return None
        return sb.storage.from_(SUPABASE_BUCKET).download(storage_path)
    except Exception as e:
        print(f"Lỗi tải avatar: {e}")
        return None

@st.cache_data(ttl=3600, show_spinner=False)
def get_chat_image_bytes_cached(file_url: str) -> bytes:
    """Tải ảnh chat có cache 1 giờ"""
    if not file_url:
        return None
    try:
        sb = get_supabase_storage()
        if not sb:
            return None
        return sb.storage.from_(SUPABASE_BUCKET).download(file_url)
    except Exception as e:
        print(f"Lỗi tải ảnh chat: {e}")
        return None

@st.cache_data(ttl=300, show_spinner=False)  # Cache 5 phút
def get_dashboard_stats():
    """Lấy toàn bộ thống kê cho Dashboard trong 1 lần query duy nhất"""
    db = st.session_state.db_engine.get_connection()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    
    # Gộp tất cả query vào 1 lần
    stats = {}
    
    # Tổng ứng viên
    c.execute("SELECT COUNT(*) as count FROM ung_vien")
    stats['tong_uv'] = c.fetchone()['count']
    
    # Tổng nhân viên đang làm
    c.execute("""
        SELECT COUNT(*) as count 
        FROM nhan_vien 
        WHERE trang_thai IN ('DANG_LAM','THU_VIEC') 
        AND so_hdld IS NOT NULL AND so_hdld != ''
    """)
    stats['tong_nv'] = c.fetchone()['count']
    
    # Ứng viên theo trạng thái
    c.execute("""
        SELECT trang_thai, COUNT(*) as count 
        FROM ung_vien 
        GROUP BY trang_thai
    """)
    stats['uv_by_status'] = {row['trang_thai']: row['count'] for row in c.fetchall()}
    
    # Phân bố nhân viên theo phòng ban
    c.execute("""
        SELECT phong_ban_lam_viec as phong_ban, COUNT(*) as count
        FROM nhan_vien 
        WHERE trang_thai IN ('DANG_LAM','THU_VIEC') 
        AND so_hdld IS NOT NULL AND so_hdld != ''
        GROUP BY phong_ban_lam_viec
        ORDER BY count DESC
    """)
    stats['nv_by_dept'] = c.fetchall()
    
    # Phân bố theo giới tính
    c.execute("""
        SELECT gioi_tinh, COUNT(*) as count
        FROM nhan_vien 
        WHERE trang_thai IN ('DANG_LAM','THU_VIEC') 
        AND so_hdld IS NOT NULL AND so_hdld != ''
        GROUP BY gioi_tinh
    """)
    stats['nv_by_gender'] = c.fetchall()
    
    # ... thêm các query khác tương tự ...
    
    db.close()
    return stats

# ========== HÀM TIỆN ÍCH MỚI ==========
def format_date_thang_nam(date_obj):
    """Định dạng ngày thành MM/YYYY"""
    if not date_obj:
        return ""
    if hasattr(date_obj, 'strftime'):
        return date_obj.strftime('%m/%Y')
    return str(date_obj)

def get_ma_tinh_from_name(tinh_name):
    """Lấy mã tỉnh từ tên tỉnh - Có thể mở rộng từ database"""
    if not tinh_name:
        return "44"  # Mặc định Quảng Trị
    
    # Map tên tỉnh -> mã tỉnh (có thể load từ database)
    ma_tinh_map = {
        'Quảng Trị': '44',
        'Quảng Bình': '43',
        'Thừa Thiên Huế': '45',
        'Huế': '45',
        'Đà Nẵng': '48',
        'Hà Nội': '01',
        'TP.HCM': '79',
        'Hồ Chí Minh': '79',
    }
    
    # Thử tìm kiếm trong map
    for key, value in ma_tinh_map.items():
        if key in tinh_name:
            return value
    
    return "44"  # Mặc định

def chuan_hoa_ten_phong_ban(ten):
    """Chuẩn hóa tên phòng ban.
    QUAN TRỌNG: tiếng Việt không viết hoa chữ cái đầu của MỌI từ (kiểu .title()/
    .capitalize() theo từng từ sẽ biến "Phòng KT-Cơ điện" thành "Phòng Kt-Cơ Điện",
    sai chính tả và không khớp với PHONG_BAN_THU_TU/PHONG_BAN_LANH_DAO_CAO_CAP).
    Thay vào đó:
    1) Nếu tên (không phân biệt hoa/thường, khoảng trắng thừa) trùng với 1 tên
       trong danh sách chuẩn PHONG_BAN_THU_TU -> trả về đúng tên chuẩn đó.
    2) Nếu không khớp -> chỉ viết hoa chữ cái đầu tiên của cả chuỗi, giữ nguyên
       phần còn lại như người dùng đã nhập (đúng kiểu viết hoa tên riêng tiếng Việt).
    """
    if not ten:
        return ""
    ten_sach = " ".join(ten.strip().split())
    if not ten_sach:
        return ""
    for chuan in PHONG_BAN_THU_TU:
        if ten_sach.lower() == chuan.lower():
            return chuan
    return ten_sach[0].upper() + ten_sach[1:]


def la_phong_ban_lanh_dao_cao_cap(ten):
    """So khớp phòng ban HĐQT/BTGĐ không phân biệt hoa/thường và khoảng trắng thừa,
    để không bị lệ thuộc vào cách viết hoa của dữ liệu cũ đã lưu trong DB."""
    ten_sach = " ".join((ten or "").strip().split()).lower()
    return any(ten_sach == pb.lower() for pb in PHONG_BAN_LANH_DAO_CAO_CAP)

def can_edit():
    """Kiểm tra xem user hiện tại có quyền chỉnh sửa dữ liệu không"""
    # Admin, HR, Văn thư, Kế toán lương có quyền chỉnh sửa
    edit_roles = ['admin', 'hr', 'van_thu', 'kt_luong']
    return st.session_state.get('role') in edit_roles

def can_delete():
    """Kiểm tra xem user hiện tại có quyền xóa dữ liệu không"""
    # Chỉ Admin mới có quyền xóa
    return st.session_state.get('role') == 'admin'

def can_export():
    """Kiểm tra xem user hiện tại có quyền xuất báo cáo không"""
    # Admin, HR, Văn thư, Kế toán lương có quyền xuất
    export_roles = ['admin', 'hr', 'van_thu', 'kt_luong']
    return st.session_state.get('role') in export_roles

def can_manage_users():
    """Kiểm tra xem user hiện tại có quyền quản lý người dùng không"""
    # Chỉ Admin mới có quyền quản lý người dùng
    return st.session_state.get('role') == 'admin'

def get_chu_ho_info(nhan_vien_id):
    """Lấy thông tin chủ hộ từ bảng phu_luc_gia_dinh"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT ho_ten, so_cccd, dien_thoai 
            FROM phu_luc_gia_dinh 
            WHERE nhan_vien_id = %s AND (quan_he_voi_chu_ho = 'Chủ hộ' OR quan_he_voi_chu_ho = 'Chủ hộ gia đình')
            LIMIT 1
        """, (nhan_vien_id,))
        result = c.fetchone()
        db.close()
        return result
    except Exception as e:
        print(f"Lỗi lấy thông tin chủ hộ: {e}")
        return None

def tao_bao_cao_bhxh_d02_lt(tang_list, giam_list, tu_ngay, den_ngay, ten_cong_ty, ma_don_vi_bhxh):
    """Tạo báo cáo tăng/giảm BHXH mẫu D02-LT đúng 100% theo file mẫu chuẩn (130+ cột)"""
    
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    
    wb = Workbook()
    ws = wb.active
    ws.title = "D02-LT"
    
    # Định nghĩa border
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Định nghĩa tất cả các cột theo đúng thứ tự file mẫu (130+ cột)
    columns = [
        "STT", "Họ và tên", "Mã số BHXH", "Loại phương án", "Mã loại PA",
        "Loại ngày sinh", "Ngày Sinh", "Giới tính", "Số CMND/ CCCD/Hộ chiếu",
        "Cấp bậc, chức vụ, chức danh nghề", "Phòng ban làm việc", "Nơi Làm Việc",
        "Mức lương", "Phụ cấp lương", "Các khoản bổ sung", "Hệ số lương",
        "Phụ cấp CV", "Phụ cấp TNVK (%)", "Phụ cấp TN nghề (%)", "Phương án điều chỉnh",
        "Mã PA", "Tháng/ năm bắt đầu", "Tháng/ năm kết thúc",
        "Nghỉ ốm đau/Thai sản/không lương", "Ghi chú", "Số sổ BHXH",
        "Mức hưởng BHYT", "Tỷ lệ đóng (%)", "Mã vùng sinh sống", "Mã vùng lương tối thiểu",
        "Có giảm chết", "Ngày chết", "Tính lãi", "Nhóm vị trí việc làm",
        "Ngày bắt đầu giữ vị trí", "Ngày kết thúc giữ vị trí", "Hợp đồng lao động",
        "Hiệu lực từ ngày", "Hiệu lực đến ngày", "Ngày bắt đầu", "Ngày kết thúc",
        "Số", "Ngày ký", "Ngành nghề nặng nhọc, độc hại", "Ngày bắt đầu",
        "Ngày kết thúc", "Hợp đồng lao động", "Số", "Ngày ký", "Quốc tịch",
        "Mã QT", "Dân tộc", "Mã DT", "Điện thoại liên hệ", "Email liên hệ",
        "Tỉnh / Thành phố (Khai sinh)", "Mã Tỉnh (Khai sinh)", "Phường/ Xã (Khai sinh)",
        "Mã xã (Khai sinh)", "Địa chỉ khai sinh", "Tỉnh / Thành phố (Nhận HS)",
        "Mã Tỉnh (Nhận HS)", "Phường/ Xã (Nhận HS)", "Mã xã (Nhận HS)",
        "Địa chỉ nhận hồ sơ", "Tỉnh nơi KCB", "Mã tỉnh (KCB)", "Nơi đăng ký KCB",
        "Mã BV", "Đăng ký nhận sổ và thẻ", "Tỉnh / Thành phố (Nhận sổ thẻ)",
        "Mã Tỉnh (Nhận sổ thẻ)", "Phường/ Xã (Nhận sổ thẻ)", "Mã Xã (Nhận sổ thẻ)",
        "Địa chỉ nhận Sổ thẻ", "Mức tiền đóng", "Phương thức đóng", "Nội dung thay đổi",
        "Hồ sơ kèm theo", "Họ tên người giám hộ", "Mã số hộ gia đình",
        "Họ Tên chủ hộ", "Số CMND/ CCCD/Hộ chiếu (chủ hộ)", "Điện thoại (chủ hộ)",
        "Loại giấy tờ", "Số giấy tờ", "Tỉnh / Thành phố (hộ khẩu)", "Mã Tỉnh (hộ khẩu)",
        "Phường/ Xã (hộ khẩu)", "Mã xã (hộ khẩu)", "Tổ/ Thôn/ Xóm", "Địa chỉ hộ khẩu",
        "Tỉnh / Thành phố thường trú", "Mã Tỉnh (thường trú)", "Phường/ Xã thường trú",
        "Mã xã (thường trú)", "Địa chỉ thường trú", "Mã số hộ gia đình (PL)",
        "Họ và tên (PL)", "Mã số BHXH (PL)", "Loại ngày sinh (PL)", "Ngày sinh (PL)",
        "Giới tính (PL)", "Quốc tịch (PL)", "Mã Quốc tịch (PL)", "Dân tộc (PL)",
        "Mã Dân tộc (PL)", "Số CMND (PL)", "Mối quan hệ với chủ hộ", "Mã MQH",
        "Tỉnh / Thành phố (PL)", "Mã Tỉnh (PL)", "Phường/ Xã (PL)", "Mã xã (PL)",
        "Địa chỉ khai sinh (PL)", "Người tham gia", "Ghi chú (PL)"
    ]
    
    # Thiết lập độ rộng cột (điều chỉnh cho phù hợp)
    col_widths = [5, 25, 18, 15, 12, 12, 15, 10, 20, 25, 20, 25, 15, 15, 15, 12, 
                  15, 12, 12, 15, 12, 15, 15, 15, 20, 15, 15, 12, 12, 12, 12, 12, 
                  12, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 
                  15, 12, 12, 12, 12, 15, 20, 15, 12, 15, 12, 20, 15, 12, 15, 12, 
                  20, 15, 12, 15, 12, 15, 12, 15, 12, 20, 15, 15, 20, 15, 20, 20, 
                  15, 15, 15, 15, 15, 15, 15, 15, 15, 20, 15, 15, 15, 20, 20, 15, 
                  15, 15, 20, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 20, 20]
    
    # Thiết lập độ rộng cột
    for idx, width in enumerate(col_widths[:len(columns)]):
        ws.column_dimensions[get_column_letter(idx + 1)].width = width
    
    # ===== HEADER =====
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    ws['A1'] = ten_cong_ty
    ws['A1'].font = Font(bold=True, size=13, name='Times New Roman')
    ws['A1'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(columns))
    ws['A2'] = f"Mã đơn vị BHXH: {ma_don_vi_bhxh}"
    ws['A2'].font = Font(size=11, name='Times New Roman')
    ws['A2'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(columns))
    ws['A3'] = "DANH SÁCH LAO ĐỘNG THAM GIA BHXH, BHYT, BHTN, BHTNLĐ, BNN (Mẫu D02-LT TK1)"
    ws['A3'].font = Font(bold=True, size=12, name='Times New Roman')
    ws['A3'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(columns))
    ws['A4'] = f"Kỳ báo cáo: Tháng {tu_ngay.strftime('%m/%Y')} - {den_ngay.strftime('%m/%Y')}"
    ws['A4'].font = Font(size=11, name='Times New Roman')
    ws['A4'].alignment = Alignment(horizontal='center')
    
    # ===== HEADER BẢNG 2 DÒNG =====
    header_row_main = 6
    header_row_sub = 7
    
    # Tạo header 2 dòng phức tạp
    header_config = [
        (1, 1, "STT"), (2, 2, "Họ và tên"), (3, 3, "Mã số BHXH"),
        (4, 4, "Loại phương án"), (5, 5, "Mã loại PA"), (6, 6, "Loại ngày sinh"),
        (7, 7, "Ngày Sinh"), (8, 8, "Giới tính"), (9, 9, "Số CMND/ CCCD/Hộ chiếu"),
        (10, 10, "Cấp bậc, chức vụ, chức danh nghề"), (11, 11, "Phòng ban làm việc"),
        (12, 12, "Nơi Làm Việc"), (13, 17, "Tiền lương"), (18, 19, "Ngành nghề nặng nhọc, độc hại"),
        (20, 24, "Loại và hiệu lực hợp đồng"), (25, 25, "Thời điểm bắt đầu đóng BHXH"),
        (26, 26, "Thời điểm kết thúc đóng BHXH")
    ]
    
    # Dòng header chính
    for start_col, end_col, text in header_config:
        if start_col == end_col:
            cell = ws.cell(row=header_row_main, column=start_col, value=text)
        else:
            ws.merge_cells(start_row=header_row_main, start_column=start_col, 
                          end_row=header_row_main, end_column=end_col)
            cell = ws.cell(row=header_row_main, column=start_col, value=text)
        cell.font = Font(bold=True, size=10, name='Times New Roman')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
    
    # Dòng header phụ
    sub_headers = {
        13: "Mức lương", 14: "Phụ cấp lương", 15: "Các khoản bổ sung",
        16: "Hệ số lương", 17: "Phụ cấp CV", 18: "Phụ cấp TNVK (%)",
        19: "Phụ cấp TN nghề (%)", 20: "Loại HĐLĐ", 21: "Hiệu lực từ ngày",
        22: "Hiệu lực đến ngày", 23: "Ngày bắt đầu", 24: "Ngày kết thúc",
        25: "Thời điểm bắt đầu", 26: "Thời điểm kết thúc"
    }
    
    for col, value in sub_headers.items():
        cell = ws.cell(row=header_row_sub, column=col, value=value)
        cell.font = Font(bold=True, size=9, name='Times New Roman')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
    
    # Các cột đơn giản còn lại
    for col in range(1, len(columns) + 1):
        if col not in [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26]:
            cell = ws.cell(row=header_row_sub, column=col, value=columns[col-1])
            cell.font = Font(bold=True, size=9, name='Times New Roman')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
            cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
    
    # ===== DỮ LIỆU =====
    all_data = []
    
    # Xử lý dữ liệu tăng
    for nv in tang_list:
        row_data = {}
        row_data['STT'] = len(all_data) + 1
        row_data['Họ và tên'] = nv.get('ho_ten', '')
        row_data['Mã số BHXH'] = nv.get('ma_so_bhxh', '')
        row_data['Loại phương án'] = "Tăng lao động"
        row_data['Mã loại PA'] = "1"
        row_data['Loại ngày sinh'] = "0"
        row_data['Ngày Sinh'] = format_date(nv.get('ngay_sinh'))
        row_data['Giới tính'] = "1" if nv.get('gioi_tinh') == 'Nam' else "2" if nv.get('gioi_tinh') == 'Nữ' else "3"
        row_data['Số CMND/ CCCD/Hộ chiếu'] = nv.get('so_cccd', '')
        row_data['Cấp bậc, chức vụ, chức danh nghề'] = nv.get('chuc_danh_nghe', '')
        row_data['Phòng ban làm việc'] = nv.get('phong_ban_lam_viec', '')
        row_data['Nơi Làm Việc'] = nv.get('noi_lam_viec', 'Cảng THQT Hòn La')
        row_data['Mức lương'] = nv.get('luong_bao_hiem', '')
        row_data['Phụ cấp lương'] = ""
        row_data['Các khoản bổ sung'] = ""
        row_data['Hệ số lương'] = nv.get('he_so_luong', '')
        row_data['Phụ cấp CV'] = nv.get('phu_cap_chuc_vu', '')
        row_data['Phụ cấp TNVK (%)'] = nv.get('phu_cap_tnvk', '')
        row_data['Phụ cấp TN nghề (%)'] = nv.get('phu_cap_tnn', '')
        
        # Xác định phương án điều chỉnh
        if nv.get('ma_so_bhxh'):
            row_data['Phương án điều chỉnh'] = "TD-Tăng đến đã có số sổ, di chuyển trong địa bàn tỉnh"
            row_data['Mã PA'] = "TD"
        else:
            row_data['Phương án điều chỉnh'] = "TM-Tăng mới chưa có số sổ"
            row_data['Mã PA'] = "TM"
        
        row_data['Tháng/ năm bắt đầu'] = format_date_thang_nam(nv.get('thang_bat_dau_bh')) if nv.get('thang_bat_dau_bh') else ""
        row_data['Tháng/ năm kết thúc'] = ""
        row_data['Nghỉ ốm đau/Thai sản/không lương'] = ""
        row_data['Ghi chú'] = nv.get('ghi_chu', '')
        row_data['Số sổ BHXH'] = nv.get('ma_so_bhxh', '')
        row_data['Mức hưởng BHYT'] = nv.get('muc_huong_bhyt', '100')
        row_data['Tỷ lệ đóng (%)'] = nv.get('ty_le_dong', '')
        row_data['Mã vùng sinh sống'] = ""
        row_data['Mã vùng lương tối thiểu'] = "03"
        row_data['Có giảm chết'] = ""
        row_data['Ngày chết'] = ""
        row_data['Tính lãi'] = ""
        row_data['Nhóm vị trí việc làm'] = ""
        row_data['Ngày bắt đầu giữ vị trí'] = format_date(nv.get('ngay_vao_lam'))
        row_data['Ngày kết thúc giữ vị trí'] = ""
        row_data['Loại HĐLĐ'] = nv.get('loai_hop_dong', '')
        row_data['Hiệu lực từ ngày'] = format_date(nv.get('ngay_ky_hd')) or format_date(nv.get('ngay_vao_lam'))
        row_data['Hiệu lực đến ngày'] = format_date(nv.get('ngay_ket_thuc')) if nv.get('ngay_ket_thuc') else ""
        row_data['Ngày bắt đầu'] = format_date(nv.get('thang_bat_dau_bh'))
        row_data['Ngày kết thúc'] = ""
        row_data['Số'] = nv.get('so_hdld', '')
        row_data['Ngày ký'] = format_date(nv.get('ngay_ky_hd')) or format_date(nv.get('ngay_vao_lam'))
        row_data['Quốc tịch'] = nv.get('quoc_tich', 'VIET NAM')
        row_data['Mã QT'] = "VN"
        row_data['Dân tộc'] = nv.get('dan_toc', 'Kinh')
        row_data['Mã DT'] = "1"
        row_data['Điện thoại liên hệ'] = nv.get('dien_thoai', '')
        row_data['Email liên hệ'] = nv.get('email_lien_he', '')
        
        # Thông tin địa chỉ
        row_data['Tỉnh / Thành phố (Khai sinh)'] = "Tỉnh Quảng Trị"
        row_data['Mã Tỉnh (Khai sinh)'] = "44"
        row_data['Phường/ Xã (Khai sinh)'] = nv.get('phuong_xa_khai_sinh', '')
        row_data['Mã xã (Khai sinh)'] = nv.get('ma_xa_khai_sinh', '')
        row_data['Địa chỉ khai sinh'] = nv.get('noi_sinh', '')
        
        row_data['Tỉnh / Thành phố (Nhận HS)'] = nv.get('tinh_nhan_hs', 'Tỉnh Quảng Trị')
        row_data['Mã Tỉnh (Nhận HS)'] = get_ma_tinh_from_name(nv.get('tinh_nhan_hs', 'Quảng Trị'))
        row_data['Phường/ Xã (Nhận HS)'] = nv.get('phuong_nhan_hs', '')
        row_data['Mã xã (Nhận HS)'] = nv.get('ma_xa_nhan_hs', '')
        row_data['Địa chỉ nhận hồ sơ'] = nv.get('dia_chi_nhan_hs', '')
        
        row_data['Tỉnh nơi KCB'] = nv.get('tinh_kcb', 'Tỉnh Quảng Trị')
        row_data['Mã tỉnh (KCB)'] = "44"
        row_data['Nơi đăng ký KCB'] = nv.get('noi_dang_ky_kcb', 'Bệnh viện đa khoa khu vực Bắc Quảng Trị')
        row_data['Mã BV'] = "44003"
        
        row_data['Đăng ký nhận sổ và thẻ'] = nv.get('dang_ky_nhan_so', 'Có')
        row_data['Tỉnh / Thành phố (Nhận sổ thẻ)'] = nv.get('tinh_nhan_hs', '')
        row_data['Mã Tỉnh (Nhận sổ thẻ)'] = get_ma_tinh_from_name(nv.get('tinh_nhan_hs', ''))
        row_data['Phường/ Xã (Nhận sổ thẻ)'] = nv.get('phuong_nhan_hs', '')
        row_data['Mã Xã (Nhận sổ thẻ)'] = nv.get('ma_xa_nhan_hs', '')
        row_data['Địa chỉ nhận Sổ thẻ'] = nv.get('dia_chi_nhan_hs', '')
        
        row_data['Mức tiền đóng'] = nv.get('muc_tien_dong', '')
        row_data['Phương thức đóng'] = nv.get('phuong_thuc_dong', 'Hàng tháng')
        
        # Lấy thông tin chủ hộ
        chu_ho = get_chu_ho_info(nv.get('id'))
        if chu_ho:
            row_data['Họ Tên chủ hộ'] = chu_ho.get('ho_ten', '')
            row_data['Số CMND/ CCCD/Hộ chiếu (chủ hộ)'] = chu_ho.get('so_cccd', '')
            row_data['Điện thoại (chủ hộ)'] = chu_ho.get('dien_thoai', '')
        
        # Nếu chưa có mã BHXH, lấy thông tin phụ lục gia đình
        if not nv.get('ma_so_bhxh'):
            family_members = get_family_members(nv.get('id'))
            if family_members:
                first_member = family_members[0] if family_members else {}
                row_data['Họ và tên (PL)'] = first_member.get('ho_ten', '')
                row_data['Ngày sinh (PL)'] = format_date(first_member.get('ngay_sinh'))
                row_data['Giới tính (PL)'] = "1" if first_member.get('gioi_tinh') == 'Nam' else "2"
                row_data['Mối quan hệ với chủ hộ'] = first_member.get('quan_he', '')
        
        all_data.append(row_data)
    
    # Xử lý dữ liệu giảm
    for nv in giam_list:
        row_data = {}
        row_data['STT'] = len(all_data) + 1
        row_data['Họ và tên'] = nv.get('ho_ten', '')
        row_data['Mã số BHXH'] = nv.get('ma_so_bhxh', '')
        row_data['Loại phương án'] = "Giảm lao động"
        row_data['Mã loại PA'] = "2"
        row_data['Ngày Sinh'] = format_date(nv.get('ngay_sinh'))
        row_data['Giới tính'] = "1" if nv.get('gioi_tinh') == 'Nam' else "2"
        row_data['Tháng/ năm kết thúc'] = format_date_thang_nam(nv.get('thang_ket_thuc_bh')) if nv.get('thang_ket_thuc_bh') else ""
        row_data['Ghi chú'] = nv.get('ly_do_nghi', '')
        all_data.append(row_data)
    
    # Ghi dữ liệu vào Excel
    start_row = header_row_sub + 1
    for idx, row_data in enumerate(all_data):
        current_row = start_row + idx
        for col_idx, col_name in enumerate(columns, 1):
            value = row_data.get(col_name, '')
            cell = ws.cell(row=current_row, column=col_idx, value=value)
            cell.font = Font(size=10, name='Times New Roman')
            cell.border = thin_border
            # Căn giữa cho các cột số và mã
            if col_idx in [1, 4, 5, 6, 8, 21, 22, 27, 28, 29, 30, 31, 32, 33]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
    
    # Footer
    total_row = start_row + len(all_data)
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=1, value=f"Tổng số: {len(all_data)} lao động")
    ws.cell(row=total_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
    
    # Ký tên
    sign_row = total_row + 3
    ws.merge_cells(start_row=sign_row, start_column=len(columns)-3, end_row=sign_row, end_column=len(columns))
    ws.cell(row=sign_row, column=len(columns)-3, value="NGƯỜI LẬP BÁO CÁO")
    ws.cell(row=sign_row, column=len(columns)-3).font = Font(bold=True, size=11, name='Times New Roman')
    ws.cell(row=sign_row, column=len(columns)-3).alignment = Alignment(horizontal='center')
    
    sign_row += 1
    ws.merge_cells(start_row=sign_row, start_column=len(columns)-3, end_row=sign_row, end_column=len(columns))
    ws.cell(row=sign_row, column=len(columns)-3, value="(Ký, ghi rõ họ tên)")
    ws.cell(row=sign_row, column=len(columns)-3).font = Font(size=10, name='Times New Roman', italic=True)
    ws.cell(row=sign_row, column=len(columns)-3).alignment = Alignment(horizontal='center')
    
    sign_row += 1
    ws.merge_cells(start_row=sign_row, start_column=len(columns)-3, end_row=sign_row, end_column=len(columns))
    ws.cell(row=sign_row, column=len(columns)-3, value=COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC').upper())
    ws.cell(row=sign_row, column=len(columns)-3).font = Font(bold=True, size=11, name='Times New Roman')
    ws.cell(row=sign_row, column=len(columns)-3).alignment = Alignment(horizontal='center')
    
    # Lưu file
    filename = f"D02-LT_BHXH_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
    wb.save(filename)
    return filename

def get_family_members(nhan_vien_id):
    """Lấy danh sách thành viên gia đình của nhân viên"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT ho_ten, ngay_sinh, gioi_tinh, quan_he_voi_chu_ho as quan_he
            FROM phu_luc_gia_dinh 
            WHERE nhan_vien_id = %s
            ORDER BY id ASC
        """, (nhan_vien_id,))
        result = c.fetchall()
        db.close()
        return result
    except Exception as e:
        print(f"Lỗi lấy thông tin gia đình: {e}")
        return []


def format_date_thang_nam(date_obj):
    """Định dạng ngày thành MM/YYYY"""
    if not date_obj:
        return ""
    if hasattr(date_obj, 'strftime'):
        return date_obj.strftime('%m/%Y')
    return str(date_obj)
        
def remove_accents(text):
    """Bỏ dấu tiếng Việt, chuyển về chữ hoa không dấu"""
    if not text:
        return ""
    # Chuẩn hóa unicode và loại bỏ dấu
    text = unicodedata.normalize('NFKD', text)
    text = re.sub(r'[\u0300-\u036f]', '', text)
    # Chuyển thành chữ hoa, chỉ giữ chữ cái và số
    text = text.upper()
    text = re.sub(r'[^A-Z0-9\s]', '', text)
    return text.strip()

def generate_ten_don_vi_thu_huong(ho_ten):
    """Tạo tên đơn vị thụ hưởng từ họ tên (bỏ dấu, in hoa)"""
    return remove_accents(ho_ten)

# Load danh sách ngân hàng từ file Excel
def load_bank_list():
    """Đọc danh sách ngân hàng từ Bank_list.xlsx"""
    banks = []
    bank_file_path = os.path.join(os.path.dirname(__file__), "Bank_list.xlsx")
    
    if os.path.exists(bank_file_path):
        try:
            df_banks = pd.read_excel(bank_file_path, sheet_name=0)
            # Tìm cột chứa tên ngân hàng
            for col in df_banks.columns:
                if 'NGÂN' in col.upper() or 'BANK' in col.upper() or 'TÊN' in col.upper():
                    banks = df_banks[col].dropna().tolist()
                    break
            if not banks:
                banks = df_banks.iloc[:, 0].dropna().tolist()
        except Exception as e:
            print(f"Lỗi đọc Bank_list.xlsx: {e}")
            banks = []
    
    # Fallback: danh sách ngân hàng mặc định
    if not banks:
        banks = [
            "MB - Ngân hàng TMCP Quân Đội",
            "TCB - Ngân hàng TMCP Kỹ Thương Việt Nam",
            "ABBANK - Ngân hàng TMCP An Bình",
            "EIB - Ngân hàng TMCP Xuất Nhập Khẩu Việt Nam",
            "HDB - Ngân hàng TMCP Phát triển TP Hồ Chí Minh",
            "BVB - Ngân hàng TMCP Bảo Việt",
            "VAB - Ngân hàng TMCP Việt Á",
            "SEAB - Ngân hàng TMCP Đông Nam Á",
            "SCB - Ngân hang TMCP Sài Gòn",
            "NASB - Ngan hang TMCP Bac A",
            "VBA - Ngan hang Nong Nghiep va Phat Trien Nong Thon Viet Nam",
            "VCB - Ngân hàng TMCP Ngoại Thương Việt Nam",
            "BIDV - Ngân hàng TMCP Đầu Tư và Phát Triển Việt Nam",
            "VIETINBANK - Ngân hàng TMCP Công Thương Việt Nam",
            "ACB - Ngân hàng TMCP Á Châu",
            "VPB - Ngân hàng TMCP Việt Nam Thịnh Vượng",
            "STB - Ngân hàng TMCP Sài Gòn Thương Tín",
            "HDB - Ngân hàng TMCP Phát triển TP Hồ Chí Minh",
            "TPB - Ngân hàng TMCP Tiên Phong",
            "SHB - Ngân hàng TMCP Sài Gòn - Hà Nội",
            "MSB - Ngân hàng TMCP Hàng Hải Việt Nam",
            "VIB - Ngân hàng TMCP Quốc Tế",
            "OCB - Ngân hàng TMCP Phương Đông",
            "LPB - Ngân hàng TMCP Bưu Điện Liên Việt",
            "VIETBANK - Ngân hàng TMCP Việt Nam Thương Tín",
        ]
    return banks

BANK_LIST = load_bank_list()

# Danh mục Trình độ học vấn/chuyên môn (dùng cho form Thêm/Sửa nhân viên)
TRINH_DO_LIST = ["THPT", "Chứng chỉ nghề", "Cao đẳng", "Đại học", "Thạc sỹ", "Tiến sĩ"]

# Thứ tự ưu tiên CHUẨN — dùng thống nhất cho mọi biểu đồ / bảng / dropdown / tìm kiếm
PHONG_BAN_THU_TU = [
    "Hội đồng Quản trị",
    "Ban Tổng Giám đốc",
    "Phòng Hành chính Nhân sự",
    "Phòng Tài chính",
    "Phòng Kinh doanh",
    "Phòng Điều độ",
    "Phòng KT-Cơ điện",
    "Tổ Cơ giới",
    "Đội Bốc xếp",
    "Đội Bảo vệ",
]

CHUC_VU_THU_TU = [
    "Chủ tịch HĐQT",
    "Phó Chủ tịch HĐQT",
    "Thành viên HĐQT",
    "Tổng Giám đốc",
    "Phó Tổng Giám đốc",
    "Chánh Văn phòng",
    "Phó Chánh VP",
    "Trưởng phòng",
    "Phó Trưởng phòng",
]

def sap_xep_phong_ban(danh_sach_phong_ban):
    """Sắp xếp tên phòng ban theo thứ tự ưu tiên chuẩn; phòng ban lạ xếp cuối theo alpha bê."""
    def key_fn(ten):
        try:
            return (0, PHONG_BAN_THU_TU.index(ten))
        except ValueError:
            return (1, ten or "")
    return sorted(danh_sach_phong_ban, key=key_fn)

def sap_xep_nhan_vien(ds_nv):
    """Sắp xếp list nhân viên (dict/RealDictRow) theo: Phòng ban -> Chức vụ -> Tên (alpha bê)."""
    def key_fn(nv):
        pb = nv.get('phong_ban_lam_viec') or ''
        try:
            pb_key = (0, PHONG_BAN_THU_TU.index(pb))
        except ValueError:
            pb_key = (1, pb)
        cv = nv.get('chuc_vu') or 'Nhân viên'
        try:
            cv_key = (0, CHUC_VU_THU_TU.index(cv))
        except ValueError:
            cv_key = (1, 0)
        return (pb_key, cv_key, nv.get('ho_ten') or '')
    return sorted(ds_nv, key=key_fn)

def sap_xep_phong_ban_rows(rows, key_field="Phòng ban"):
    """Sắp xếp list các dict/RealDictRow theo thứ tự phòng ban chuẩn (PHONG_BAN_THU_TU)."""
    def key_fn(row):
        ten = row.get(key_field) or ""
        try:
            return (0, PHONG_BAN_THU_TU.index(ten))
        except ValueError:
            return (1, ten)
    return sorted(rows, key=key_fn)

def get_phong_ban_options():
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT ten_phong_ban FROM danh_muc_phong_ban WHERE trang_thai = TRUE")
        # Chuẩn hoá viết hoa chữ cái đầu mỗi từ (Title Case) khi hiển thị, bất kể dữ liệu
        # gốc trong danh mục viết hoa/thường không nhất quán. Dùng dict để loại trùng lặp
        # phát sinh do khác biệt viết hoa (VD: "phòng kế toán" và "Phòng Kế Toán" từng là
        # 2 dòng riêng biệt trong danh mục cũ).
        ds_raw = [row[0] for row in c.fetchall() if row[0]]
        db.close()
        ds = list(dict.fromkeys(chuan_hoa_ten_phong_ban(t) for t in ds_raw))
        if not ds:
            ds = list(PHONG_BAN_THU_TU)  # fallback khi danh mục chưa cấu hình
        return sap_xep_phong_ban(ds)
    except Exception:
        return list(PHONG_BAN_THU_TU)

# ============================================================
# 🤖 CHATBOT GIẢI ĐÁP — AI Tư vấn Hành chính Nhân sự
# (chuyển thể từ bản HTML/JS gốc sang Python để chạy trong Streamlit,
#  gọi Anthropic API từ phía server bằng API key lưu trong st.secrets)
# ============================================================
CHATBOT_LAW_DB = {
    "bhxh": [
        {"id": "L01", "ref": "Điều 54, Luật BHXH 2014", "text": "Người lao động được hưởng lương hưu khi đủ tuổi đời (nam đủ 60, nữ đủ 55 — đang lộ trình tăng lên 62/60 tuổi) và đóng BHXH từ đủ 20 năm trở lên."},
        {"id": "L02", "ref": "Điều 39, Luật BHXH 2014", "text": "Thời gian hưởng chế độ ốm đau tối đa: 30 ngày/năm nếu đóng BHXH dưới 15 năm; 40 ngày nếu đóng 15–30 năm; 60 ngày nếu đóng từ 30 năm trở lên. Mắc bệnh dài ngày trong danh mục được hưởng tối đa 180 ngày."},
        {"id": "L03", "ref": "NĐ 115/2015/NĐ-CP, Điều 6", "text": "Mức đóng BHXH bắt buộc: NLĐ đóng 8% tiền lương vào quỹ hưu trí, tử tuất. NSDLĐ đóng 17,5%: ốm đau 3%, thai sản 0,5%, TNLĐ-BNN 0,5%, hưu trí tử tuất 14%."},
        {"id": "L04", "ref": "Khoản 1 Điều 40, Luật BHXH 2014", "text": "Mức hưởng chế độ ốm đau bằng 75% mức tiền lương đóng BHXH của tháng liền kề trước khi nghỉ việc."},
        {"id": "L05", "ref": "Khoản 2 Điều 56, Luật BHXH 2014", "text": "Mức lương hưu tính bằng 45% bình quân tiền lương tháng đóng BHXH tương ứng với 20 năm đóng (nam), sau đó cứ thêm 1 năm đóng thêm 2%, tối đa 75%. Nữ tính từ 15 năm = 45%."},
    ],
    "thaisan": [
        {"id": "T01", "ref": "Điều 34, Luật BHXH 2014", "text": "Lao động nữ sinh con được nghỉ thai sản 6 tháng (gồm trước và sau sinh). Sinh đôi trở lên cứ mỗi con từ con thứ hai được nghỉ thêm 01 tháng."},
        {"id": "T02", "ref": "Khoản 1 Điều 39, Luật BHXH 2014", "text": "Mức hưởng thai sản bằng 100% mức bình quân tiền lương tháng đóng BHXH của 06 tháng trước khi nghỉ việc hưởng chế độ."},
        {"id": "T03", "ref": "Điều 38, Luật BHXH 2014", "text": "Điều kiện hưởng thai sản: đóng BHXH từ đủ 06 tháng trở lên trong 12 tháng trước khi sinh con hoặc nhận nuôi con nuôi dưới 06 tháng tuổi."},
        {"id": "T04", "ref": "Điều 36, Luật BHXH 2014", "text": "Lao động nữ được nghỉ hưởng thai sản khi khám thai: 05 lần, mỗi lần 01 ngày. Ở xa hoặc thai không bình thường được nghỉ 02 ngày mỗi lần khám."},
        {"id": "T05", "ref": "Khoản 1 Điều 38, Luật BHXH 2014", "text": "Trợ cấp một lần khi sinh con: 02 lần mức lương cơ sở tại tháng sinh con (hiện lương cơ sở = 2.340.000 đồng/tháng từ 1/7/2024 → trợ cấp một lần = 4.680.000 đồng)."},
    ],
    "thatnghiep": [
        {"id": "TN01", "ref": "Điều 49, Luật Việc làm 2013", "text": "Điều kiện hưởng BHTN: đã chấm dứt HĐLĐ; đã đóng BHTN từ đủ 12 tháng trở lên trong 24 tháng trước khi chấm dứt HĐLĐ (với hợp đồng không xác định thời hạn và xác định thời hạn từ 3 tháng–36 tháng)."},
        {"id": "TN02", "ref": "Khoản 1 Điều 50, Luật Việc làm 2013", "text": "Mức hưởng TCTN hàng tháng = 60% × bình quân tiền lương tháng đóng BHTN của 06 tháng liền kề trước khi thất nghiệp, tối đa không quá 05 lần mức lương cơ sở."},
        {"id": "TN03", "ref": "Khoản 2 Điều 50, Luật Việc làm 2013", "text": "Thời gian hưởng TCTN: đóng 12–35 tháng → 03 tháng; cứ đóng thêm đủ 12 tháng → hưởng thêm 01 tháng, tối đa không quá 12 tháng."},
        {"id": "TN04", "ref": "NĐ 28/2015/NĐ-CP, Điều 17", "text": "Người lao động phải nộp hồ sơ hưởng TCTN trong thời hạn 03 tháng kể từ ngày chấm dứt HĐLĐ tại trung tâm dịch vụ việc làm nơi cư trú."},
    ],
    "thuetncn": [
        {"id": "TAX01", "ref": "Nghị quyết 954/2020/UBTVQH14", "text": "Mức giảm trừ gia cảnh: bản thân người nộp thuế 11 triệu/tháng (132 triệu/năm); mỗi người phụ thuộc 4,4 triệu/tháng. Áp dụng từ kỳ tính thuế năm 2020."},
        {"id": "TAX02", "ref": "Điều 22, Luật Thuế TNCN 2007 (sửa đổi 2012)", "text": "Biểu thuế lũy tiến từng phần: Bậc 1 ≤5tr: 5% | Bậc 2 (5–10tr): 10% | Bậc 3 (10–18tr): 15% | Bậc 4 (18–32tr): 20% | Bậc 5 (32–52tr): 25% | Bậc 6 (52–80tr): 30% | Bậc 7 >80tr: 35%."},
        {"id": "TAX03", "ref": "Điều 9, Thông tư 111/2013/TT-BTC", "text": "Người phụ thuộc được đăng ký giảm trừ gồm: con dưới 18 tuổi; con từ 18 tuổi bị khuyết tật không có khả năng lao động; cha mẹ, vợ/chồng không có khả năng lao động hoặc không có thu nhập hoặc thu nhập ≤1 triệu/tháng."},
        {"id": "TAX04", "ref": "Điều 25, Thông tư 111/2013/TT-BTC", "text": "Tổ chức trả thu nhập khấu trừ 10% đối với HĐLĐ dưới 3 tháng từ 2 triệu/lần. Với HĐLĐ từ 3 tháng trở lên, khấu trừ theo biểu lũy tiến từng phần hàng tháng."},
    ],
    "bhyt": [
        {"id": "BHYT01", "ref": "Điều 13, Luật BHYT 2008 (sửa đổi 2014)", "text": "Mức đóng BHYT = 4,5% tiền lương: NLĐ đóng 1,5%, NSDLĐ đóng 3%."},
        {"id": "BHYT02", "ref": "Điều 22, Luật BHYT 2008 (sửa đổi 2014)", "text": "Mức hưởng BHYT: đúng tuyến 80% chi phí; tuyến tỉnh không đúng tuyến 60%; tuyến TW không đúng tuyến 40%; cấp cứu 100% đến khi ổn định. Người có công, người nghèo: 100%."},
        {"id": "BHYT03", "ref": "Khoản 1 Điều 23, Luật BHYT sửa đổi 2014", "text": "Người tham gia BHYT liên tục từ 05 năm trở lên, tổng tiền cùng chi trả trong năm vượt 6 tháng lương cơ sở thì chỉ cùng chi trả tối đa bằng 6 tháng lương cơ sở."},
    ],
    "hopdong": [
        {"id": "HD01", "ref": "Điều 34, Bộ luật Lao động 2019", "text": "HĐLĐ xác định thời hạn không quá 36 tháng. Hết hạn 01 lần nếu tiếp tục sử dụng thì phải ký HĐLĐ không xác định thời hạn, trừ lao động cao tuổi và lao động nước ngoài."},
        {"id": "HD02", "ref": "Điều 46, Bộ luật Lao động 2019", "text": "Trợ cấp thôi việc: NLĐ làm việc đủ 12 tháng, mỗi năm làm việc được trợ cấp ½ tháng lương. Tiền lương tính là bình quân 06 tháng liền kề. Thời gian tính trợ cấp là tổng thời gian làm việc trừ thời gian đã hưởng TCTN từ BHTN."},
        {"id": "HD03", "ref": "Điều 36, Bộ luật Lao động 2019", "text": "NLĐ đơn phương chấm dứt HĐLĐ phải báo trước: không xác định thời hạn ≥45 ngày; xác định thời hạn 12–36 tháng ≥30 ngày; dưới 12 tháng ≥03 ngày làm việc."},
        {"id": "HD04", "ref": "Điều 41, Bộ luật Lao động 2019", "text": "NSDLĐ đơn phương chấm dứt HĐLĐ trái pháp luật: phải nhận NLĐ trở lại, trả lương những ngày không được làm việc, và bồi thường thêm ít nhất 02 tháng tiền lương theo HĐLĐ."},
    ],
}

def _chatbot_detect_laws(q):
    """Dò các điều luật liên quan tới câu hỏi (dựa vào từ khoá), tương đương hàm dLaws() bản JS."""
    import re as _re
    t = q.lower()
    laws = []
    patterns = [
        (r"bhxh|bảo hiểm xã hội|hưu|ốm đau|đóng bảo hiểm|lương hưu", "bhxh"),
        (r"thai sản|sinh con|nghỉ thai|mang thai|khám thai", "thaisan"),
        (r"thất nghiệp|bhtn|mất việc", "thatnghiep"),
        (r"thuế|tncn|thu nhập|giảm trừ|gia cảnh|phụ thuộc", "thuetncn"),
        (r"bhyt|bảo hiểm y tế|khám bệnh|viện phí", "bhyt"),
        (r"hợp đồng|hđlđ|thôi việc|sa thải|chấm dứt|báo trước|trợ cấp", "hopdong"),
    ]
    for pattern, key in patterns:
        if _re.search(pattern, t):
            laws.extend(CHATBOT_LAW_DB[key])
    if not laws:
        for arr in CHATBOT_LAW_DB.values():
            laws.extend(arr[:2])
    return laws

def _chatbot_all_laws():
    ket_qua = []
    for arr in CHATBOT_LAW_DB.values():
        ket_qua.extend(arr)
    return ket_qua

def _chatbot_system_prompt(laws):
    laws_text = "\n".join(f'[{l["id"]}] {l["ref"]}: "{l["text"]}"' for l in laws)
    return f"""Bạn là chuyên gia tư vấn pháp luật hành chính nhân sự Việt Nam với 15+ năm kinh nghiệm. Tư vấn chuyên nghiệp, cụ thể, có căn cứ pháp lý.

ĐIỀU LUẬT ĐÃ TRUY XUẤT:
{laws_text}

Trả về JSON thuần (KHÔNG markdown, KHÔNG text ngoài JSON):
{{"summary":"Tóm tắt 1-2 câu","analysis":"Phân tích chi tiết 3-5 câu","options":[{{"label":"Phương án A – tên ngắn","detail":"Mô tả cụ thể","risk":"Rủi ro hoặc rỗng","type":"recommended|alternative|risky"}}],"citations":["ID1","ID2"],"calculations":[{{"label":"Tên khoản","formula":"Công thức","result":"Kết quả nếu đủ số liệu"}}],"note":"Lưu ý quan trọng hoặc thông tin cần bổ sung"}}"""

def _chatbot_get_api_key():
    try:
        return st.secrets.get("ANTHROPIC_API_KEY") or st.secrets.get("anthropic", {}).get("api_key")
    except Exception:
        return None

def _chatbot_call_claude(system_prompt, history):
    """Gọi Anthropic Messages API từ phía server (Streamlit backend), trả về dict đã parse JSON."""
    api_key = _chatbot_get_api_key()
    if not api_key:
        return {"summary": "⚠️ Chưa cấu hình ANTHROPIC_API_KEY trong Secrets của app.", "analysis": "",
                "options": [], "citations": [], "calculations": [], "note": "Vào Streamlit Cloud → Manage app → Secrets, thêm dòng: ANTHROPIC_API_KEY = \"sk-ant-...\""}
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-sonnet-4-5-20250929",
                "max_tokens": 1200,
                "system": system_prompt,
                "messages": history,
            },
            timeout=30,
        )
        data = resp.json()
        if "content" not in data:
            err_msg = data.get("error", {}).get("message", str(data))
            return {"summary": f"❌ Lỗi gọi API: {err_msg}", "analysis": "", "options": [],
                    "citations": [], "calculations": [], "note": ""}
        raw = "".join(b.get("text", "") for b in data["content"])
        raw_clean = raw.replace("```json", "").replace("```", "").strip()
        try:
            return json.loads(raw_clean)
        except Exception:
            return {"summary": raw[:400], "analysis": "", "options": [], "citations": [], "calculations": [], "note": ""}
    except Exception as e:
        return {"summary": f"❌ Lỗi kết nối: {e}", "analysis": "", "options": [], "citations": [], "calculations": [], "note": ""}

def _chatbot_badge_html(loai):
    return {
        "recommended": '<span style="font-size:10px;padding:2px 8px;border-radius:20px;background:#e6f4ea;color:#1e6e3a;border:1px solid #8fc8a3;font-weight:600">Khuyến nghị</span>',
        "alternative": '<span style="font-size:10px;padding:2px 8px;border-radius:20px;background:#e8f0fe;color:#1a56c4;border:1px solid #93b4f5;font-weight:600">Thay thế</span>',
        "risky": '<span style="font-size:10px;padding:2px 8px;border-radius:20px;background:#fce8e6;color:#b91c1c;border:1px solid #f5a3a3;font-weight:600">Rủi ro cao</span>',
    }.get(loai, '<span style="font-size:10px;padding:2px 8px;border-radius:20px;background:#e8f0fe;color:#1a56c4;border:1px solid #93b4f5;font-weight:600">Thay thế</span>')

def _chatbot_render_answer_html(data):
    """Dựng HTML thẻ trả lời có cấu trúc (tương đương hàm render() bản JS), hiển thị qua st.markdown."""
    import html as _html
    al = _chatbot_all_laws()
    h = '<div style="padding:14px 16px;background:#fff;border:1px solid #e5e7eb;border-radius:4px 14px 14px 14px;box-shadow:0 1px 4px rgba(0,0,0,.05)">'
    h += '<div style="display:flex;align-items:center;gap:6px;margin-bottom:11px;padding-bottom:8px;border-bottom:1px solid #f1f5f9">' \
         '<div style="width:22px;height:22px;border-radius:6px;background:#1e3a5f;display:flex;align-items:center;justify-content:center;font-size:11px">⚖️</div>' \
         '<span style="font-size:10.5px;font-weight:700;color:#1e3a5f;letter-spacing:.06em">CHUYÊN GIA TƯ VẤN HCNS</span></div>'
    if data.get("summary"):
        h += f'<div style="padding:10px 13px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;margin-bottom:11px;font-size:13px;color:#0c4a6e;font-weight:500;line-height:1.65">{_html.escape(data["summary"])}</div>'
    if data.get("analysis"):
        h += f'<div style="font-size:13.5px;color:#374151;line-height:1.8;margin-bottom:13px">{_html.escape(data["analysis"])}</div>'
    if data.get("options"):
        h += '<div style="font-size:10px;font-weight:700;color:#6b7280;letter-spacing:.07em;text-transform:uppercase;margin-bottom:7px">⚖️ Phương án xử lý</div>'
        for o in data["options"]:
            risk_html = f'<div style="margin-top:5px;font-size:12px;color:#b91c1c;background:#fef2f2;padding:4px 9px;border-radius:5px">⚠️ {_html.escape(o.get("risk",""))}</div>' if o.get("risk") else ""
            h += f'<div style="margin-bottom:7px;padding:10px 13px;background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px">' \
                 f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;flex-wrap:wrap">' \
                 f'<span style="font-size:13px;font-weight:700;color:#111827">{_html.escape(o.get("label",""))}</span>{_chatbot_badge_html(o.get("type"))}</div>' \
                 f'<div style="font-size:13px;color:#4b5563;line-height:1.65">{_html.escape(o.get("detail",""))}</div>{risk_html}</div>'
    if data.get("calculations"):
        h += '<div style="font-size:10px;font-weight:700;color:#6b7280;letter-spacing:.07em;text-transform:uppercase;margin:10px 0 7px">🔢 Tính toán cụ thể</div>'
        for c in data["calculations"]:
            ket_qua = f'<span style="font-weight:700;color:#14532d"> = {_html.escape(str(c.get("result","")))}</span>' if c.get("result") else ""
            h += f'<div style="margin-bottom:5px;padding:8px 12px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;font-size:13px">' \
                 f'<span style="font-weight:700;color:#166534">{_html.escape(c.get("label",""))}: </span>' \
                 f'<span style="color:#15803d">{_html.escape(c.get("formula",""))}</span>{ket_qua}</div>'
    if data.get("citations"):
        matched = [l for l in al if l["id"] in data["citations"]]
        if matched:
            h += '<div style="margin-top:13px;border-top:1px solid #f3f4f6;padding-top:11px">' \
                 '<div style="font-size:10px;font-weight:700;color:#6b7280;letter-spacing:.07em;text-transform:uppercase;margin-bottom:7px">📚 Căn cứ pháp lý</div>'
            for law in matched:
                h += f'<div style="margin-bottom:8px;padding:8px 12px;background:#fffbeb;border:1px solid #fde68a;border-left:3px solid #f59e0b;border-radius:6px">' \
                     f'<div style="font-size:11px;font-weight:700;color:#92400e;margin-bottom:3px">{_html.escape(law["ref"])}</div>' \
                     f'<div style="font-size:12px;color:#78350f;line-height:1.65">{_html.escape(law["text"])}</div></div>'
            h += '</div>'
    if data.get("note"):
        h += f'<div style="margin-top:9px;padding:8px 12px;background:#faf5ff;border:1px solid #e9d5ff;border-radius:6px;font-size:12.5px;color:#6b21a8;line-height:1.65">💡 <strong>Lưu ý:</strong> {_html.escape(data["note"])}</div>'
    h += '</div>'
    return h

def upload_anh_ho_so(ma_nv_or_id, ho_ten, uploaded_file):
    """Upload ảnh hồ sơ nhân viên lên Supabase Storage (dùng chung bucket hồ sơ,
    lưu trong thư mục con 'avatars/'). Trả về storage_path đã lưu, hoặc None nếu lỗi."""
    if not uploaded_file:
        return None
    sb = get_supabase_storage()
    if not sb:
        st.warning("⚠️ Chưa cấu hình Supabase Storage nên không lưu được ảnh hồ sơ (các thông tin khác vẫn được lưu bình thường).")
        return None
    try:
        safe_name = sanitize_storage_filename(uploaded_file.name)
        ten_folder = sanitize_storage_filename(f"{ma_nv_or_id}_{ho_ten}")
        base_path = f"avatars/{ten_folder}/{safe_name}"
        return upload_to_storage_unique(
            sb, SUPABASE_BUCKET, base_path,
            uploaded_file.getvalue(), uploaded_file.type
        )
    except Exception as e:
        st.warning(f"⚠️ Lỗi upload ảnh hồ sơ (các thông tin khác vẫn được lưu): {e}")
        return None

@st.cache_data(ttl=600, show_spinner=False)
def get_anh_ho_so_bytes(storage_path):
    """Tải bytes ảnh hồ sơ từ Supabase Storage để hiển thị (bucket riêng tư nên
    không dùng URL public trực tiếp được). Cache 10 phút để đỡ tải lại liên tục."""
    if not storage_path:
        return None
    try:
        sb = get_supabase_storage()
        if not sb:
            return None
        return sb.storage.from_(SUPABASE_BUCKET).download(storage_path)
    except Exception:
        return None

# ===== CHẤM CÔNG - HẰNG SỐ MỚI =====
CHAM_CONG_MA_CODE = {
    "":    "(Trống) - Chưa chấm công",
    "X":   "X - Ngày công thường (đủ ca)",
    "P":   "P - Nghỉ phép hưởng lương",
    "V":   "V - Vắng mặt/nghỉ không lương",
    "N":   "N - Ca làm việc 8T - ngày",
    "D":   "D - Ca làm việc 8T - đêm",
    "L":   "L - Đi làm ngày lễ",
    "0.5": "0.5 - Làm nửa ngày công thường",
    "NL":  "NL - Nghỉ lễ hưởng nguyên lương",
    "CN":  "CN - Chủ nhật (nghỉ) - auto đánh dấu",
}
CHAM_CONG_MA_OPTIONS = list(CHAM_CONG_MA_CODE.keys())
CHAM_CONG_CELL_REGEX = r"^$|^[XxPpVvNnDdLl]$|^[Nn][Ll]$|^[Cc][Nn]$|^\d{1,2}(\.\d)?$"

# Bộ phận chỉ chấm công 1 dòng (Văn phòng)
CHAM_CONG_DEPT_MOT_DONG = ["VP"]

# Bộ phận có 2 dòng (Ca chính + Tăng ca)
CHAM_CONG_DEPT_HAI_DONG = ["QL", "SX", "LDPT"]

# === THÊM MỚI: Từ khóa nhận diện Văn phòng ===
VAN_PHONG_KEYWORDS = ["VP", "VĂN PHÒNG", "ADMIN", "HÀNH CHÍNH"]

def is_van_phong(dept):
    """Kiểm tra xem bộ phận có thuộc Văn phòng không"""
    if not dept:
        return False
    dept_upper = dept.upper()
    for kw in VAN_PHONG_KEYWORDS:
        if kw in dept_upper:
            return True
    return False

CHAM_CONG_DEPT_LABEL = {
    "QL": "QL - Quản lý",
    "VP": "VP - Văn phòng",
    "SX": "SX - Sản xuất/Vận hành",
    "LDPT": "LDPT - Lao động phổ thông",
}

# Xử lý đổi ngôn ngữ từ request
def handle_language_change():
    """Xử lý thay đổi ngôn ngữ từ query params"""
    query_params = st.query_params
    if 'lang' in query_params:
        new_lang = query_params['lang']
        if new_lang in ['vi', 'en']:
            st.session_state.language = new_lang
            st.query_params.clear()
            st.rerun()
    
    # Cũng kiểm tra POST request (cho fetch từ client)
    try:
        # Lấy dữ liệu từ request (nếu có)
        import sys
        if hasattr(st, 'context') and hasattr(st.context, 'headers'):
            content_length = int(st.context.headers.get('content-length', 0))
            if content_length > 0:
                body = sys.stdin.read(content_length) if content_length else ''
                if 'set_language=' in body:
                    new_lang = body.replace('set_language=', '').strip()
                    if new_lang in ['vi', 'en']:
                        st.session_state.language = new_lang
                        st.rerun()
    except:
        pass

# Gọi hàm xử lý ngôn ngữ trước khi hiển thị landing page
handle_language_change()

def show_landing_page():
    """Hiển thị Landing Page với chuyển ngữ Việt/Anh"""
    
    # Import languages
    try:
        from languages import LANGUAGES
    except ImportError:
        # Fallback nếu chưa có file languages.py
        LANGUAGES = {'vi': {}, 'en': {}}
    
    lang = st.session_state.get('language', 'vi')
    text = LANGUAGES.get(lang, LANGUAGES.get('vi', {}))
    
    # Ẩn UI chrome của Streamlit
    st.markdown("""
        <style>
            [data-testid="stSidebar"],
            [data-testid="collapsedControl"],
            [data-testid="stDecoration"],
            [data-testid="stHeader"],
            header[data-testid],
            footer[data-testid],
            .stAppDeployButton,
            .stToolbar,
            .stStatusWidget,
            .stApp > header,
            .stApp > div[data-testid="stToolbar"] {
                display: none !important;
                height: 0 !important;
            }
            html, body, .stApp, .stApp > div {
                margin: 0 !important;
                padding: 0 !important;
            }
            .main > div {
                padding: 0 !important;
                margin: 0 !important;
            }
            .block-container {
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                padding-left: 0 !important;
                padding-right: 0 !important;
                max-width: 100% !important;
            }
            section[data-testid="stMain"] > div {
                padding-top: 0 !important;
            }
            iframe {
                border: none !important;
                display: block !important;
                margin: 0 !important;
                padding: 0 !important;
                width: 100% !important;
            }
            body {
                overflow-x: hidden;
            }
            [data-testid="stDataFrame"] {
                max-height: 700px !important;
            }
            [data-testid="stDataFrame"] > div {
                max-height: 700px !important;
            }
            [data-testid="stDataEditor"] {
                max-height: 700px !important;
            }
            [data-testid="stDataEditor"] > div {
                max-height: 700px !important;
            }
            [data-testid="stDataFrame"] td:last-child,
            [data-testid="stDataEditor"] td:last-child {
                background-color: #E8F5E9 !important;
                font-weight: bold !important;
                cursor: pointer !important;
            }
            
            /* ===== Tăng chiều cao bảng chấm công ===== */
            [data-testid="stDataFrame"] {
                max-height: 800px !important;
            }
            [data-testid="stDataFrame"] > div {
                max-height: 800px !important;
            }
            [data-testid="stDataEditor"] {
                max-height: 800px !important;
            }
            [data-testid="stDataEditor"] > div {
                max-height: 800px !important;
            }
        </style>
    """, unsafe_allow_html=True)
    
    # Đọc file logo động
    import base64
    import requests
    logo_base64 = ""
    logo_src = COMPANY_CONFIG.get("logo_url")
    if logo_src:
        if logo_src.startswith("http://") or logo_src.startswith("https://"):
            try:
                response = requests.get(logo_src, timeout=3)
                if response.status_code == 200:
                    logo_base64 = base64.b64encode(response.content).decode()
            except Exception:
                pass
        elif os.path.exists(logo_src):
            try:
                with open(logo_src, "rb") as f:
                    logo_base64 = base64.b64encode(f.read()).decode()
            except Exception:
                pass
    
    if not logo_base64:
        # Fallback về logo_cty.png mặc định
        logo_path = os.path.join(os.path.dirname(__file__), "logo_cty.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                logo_base64 = base64.b64encode(f.read()).decode()
    
    # Đọc ảnh slider
    def load_img_b64(filename):
        path = os.path.join(os.path.dirname(__file__), "static", filename)
        if os.path.exists(path):
            with open(path, "rb") as f:
                ext = filename.rsplit(".", 1)[-1].lower()
                mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
                return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
        return ""
    
    slide1_src = load_img_b64("anh1.jpeg")
    slide2_src = load_img_b64("anh2.jpeg")
    slide3_src = load_img_b64("anh3.jpeg")
    chu_tich_img = load_img_b64("333.png")
    chu_ky_img = load_img_b64("123456.png")
    
    # Active class cho language buttons
    vi_active = 'active' if lang == 'vi' else ''
    en_active = 'active' if lang == 'en' else ''
    
    # JavaScript riêng biệt (không có {} để tránh xung đột)
    landing_js = """
    <script>
        // Hàm cuộn mượt đến section bằng window.top
        function scrollToSection(sectionId) {
            var topWin = window.top || window.parent || window;
            var targetElement = document.getElementById(sectionId);
            if (targetElement) {
                var targetRect = targetElement.getBoundingClientRect();
                var offsetTop = targetRect.top + (topWin.scrollY || topWin.pageYOffset);
                topWin.scrollTo({
                    top: offsetTop - 80,
                    behavior: 'smooth'
                });
            }
        }
        
        function handleNavClick(e) {
            e.preventDefault();
            var section = this.getAttribute('data-section');
            if (section) {
                scrollToSection(section);
            }
        }
        
        function initNavigation() {
            var navLinks = document.querySelectorAll('.nav-link');
            for (var i = 0; i < navLinks.length; i++) {
                navLinks[i].removeEventListener('click', handleNavClick);
                navLinks[i].addEventListener('click', handleNavClick);
            }
        }
        
        // Slider
        var currentSlide = 0;
        var slides = document.querySelectorAll('.slide');
        var dots = document.querySelectorAll('.slider-dot');
        var totalSlides = slides.length;
        var autoSlideInterval;
        var progressInterval;
        var progressValue = 0;
        var SLIDE_DURATION = 5000;
        var progressBar = document.getElementById('sliderProgress');
        
        function showSlide(index) {
            for (var i = 0; i < slides.length; i++) {
                slides[i].classList.remove('active');
                if (dots[i]) dots[i].classList.remove('active');
            }
            slides[index].classList.add('active');
            if (dots[index]) dots[index].classList.add('active');
            currentSlide = index;
            resetProgress();
        }
        
        function nextSlide() {
            showSlide((currentSlide + 1) % totalSlides);
        }
        
        function prevSlide() {
            showSlide((currentSlide - 1 + totalSlides) % totalSlides);
        }
        
        function resetProgress() {
            progressValue = 0;
            if (progressBar) progressBar.style.width = '0%';
        }
        
        function startProgress() {
            if (progressInterval) clearInterval(progressInterval);
            progressValue = 0;
            progressInterval = setInterval(function() {
                progressValue += 100 / (SLIDE_DURATION / 100);
                if (progressBar) progressBar.style.width = Math.min(progressValue, 100) + '%';
                if (progressValue >= 100) resetProgress();
            }, 100);
        }
        
        function startAutoSlide() {
            if (autoSlideInterval) clearInterval(autoSlideInterval);
            autoSlideInterval = setInterval(nextSlide, SLIDE_DURATION);
            startProgress();
        }
        
        if (totalSlides > 0) {
            for (var i = 0; i < dots.length; i++) {
                dots[i].addEventListener('click', (function(idx) {
                    return function() {
                        showSlide(idx);
                        if (autoSlideInterval) clearInterval(autoSlideInterval);
                        if (progressInterval) clearInterval(progressInterval);
                        startAutoSlide();
                    };
                })(i));
            }
            
            var prevBtn = document.getElementById('prevBtn');
            var nextBtn = document.getElementById('nextBtn');
            if (prevBtn) {
                prevBtn.addEventListener('click', function() {
                    prevSlide();
                    if (autoSlideInterval) clearInterval(autoSlideInterval);
                    if (progressInterval) clearInterval(progressInterval);
                    startAutoSlide();
                });
            }
            if (nextBtn) {
                nextBtn.addEventListener('click', function() {
                    nextSlide();
                    if (autoSlideInterval) clearInterval(autoSlideInterval);
                    if (progressInterval) clearInterval(progressInterval);
                    startAutoSlide();
                });
            }
            
            var touchStartX = 0;
            var heroSlider = document.querySelector('.hero-slider');
            if (heroSlider) {
                heroSlider.addEventListener('touchstart', function(e) {
                    touchStartX = e.touches[0].clientX;
                });
                heroSlider.addEventListener('touchend', function(e) {
                    var diff = touchStartX - e.changedTouches[0].clientX;
                    if (Math.abs(diff) > 50) {
                        diff > 0 ? nextSlide() : prevSlide();
                        startAutoSlide();
                    }
                });
            }
            startAutoSlide();
        }
        
        // Scroll reveal
        var revealObserver = new IntersectionObserver(function(entries) {
            entries.forEach(function(entry) {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                    revealObserver.unobserve(entry.target);
                }
            });
        }, { threshold: 0.15 });
        document.querySelectorAll('.reveal').forEach(function(el) {
            revealObserver.observe(el);
        });
        
        // Navbar scroll effect
        window.addEventListener('scroll', function() {
            var navbar = document.getElementById('navbar');
            if (navbar) {
                if (window.scrollY > 50) {
                    navbar.style.background = 'rgba(15, 59, 92, 0.98)';
                    navbar.style.boxShadow = '0 4px 20px rgba(0,0,0,0.2)';
                } else {
                    navbar.style.background = 'rgba(15, 59, 92, 0.95)';
                    navbar.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';
                }
            }
        });
        
        // Modal
        var modal = document.getElementById('thuNgoModal');
        var thuNgoBtn = document.getElementById('thuNgoBtn');
        var closeModalBtn = document.getElementById('closeModalBtn');
        
        if (thuNgoBtn && modal) {
            thuNgoBtn.addEventListener('click', function(e) {
                e.preventDefault();
                modal.classList.add('active');
                document.body.style.overflow = 'hidden';
            });
            
            var closeModal = function() {
                modal.classList.remove('active');
                document.body.style.overflow = '';
            };
            
            if (closeModalBtn) closeModalBtn.addEventListener('click', closeModal);
            modal.addEventListener('click', function(e) {
                if (e.target === modal) closeModal();
            });
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape' && modal.classList.contains('active')) closeModal();
            });
        }
        
        // Career link
        var careerLink = document.getElementById('careerLink');
        if (careerLink) {
            careerLink.addEventListener('click', function(e) {
                e.preventDefault();
                alert('Vui lòng liên hệ HR qua email: hr@honlaport.com.vn');
            });
        }
        
        // Language switcher
        function switchLanguage(lang) {
            var topWin = window.top || window.parent || window;
            var url = new URL(topWin.location.href);
            url.searchParams.set('lang', lang);
            topWin.history.replaceState(null, '', url.toString());
            topWin.location.reload();
        }
        
        // Khởi tạo navigation
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', function() {
                initNavigation();
            });
        } else {
            initNavigation();
        }
    </script>
    """
    
    landing_html = f"""
    <!DOCTYPE html>
    <html lang="{lang}">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
        <title>Cảng Quốc tế Hòn La</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;14..32,400;14..32,500;14..32,600;14..32,700;14..32,800&display=swap" rel="stylesheet">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            body {{
                font-family: 'Inter', sans-serif;
                background-color: #ffffff;
                color: #1e293b;
                line-height: 1.5;
                overflow-x: hidden;
                width: 100%;
                padding-top: 100px;
            }}
            ::-webkit-scrollbar {{
                width: 8px;
            }}
            ::-webkit-scrollbar-track {{
                background: #f1f1f1;
            }}
            ::-webkit-scrollbar-thumb {{
                background: #0f3b5c;
                border-radius: 4px;
            }}
            
            /* ===== NAVIGATION ===== */
            .navbar {{
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                z-index: 1000;
                padding: 0.8rem 30px;                                       
                background: rgba(15, 59, 92, 0.95);
                backdrop-filter: blur(10px);
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            .nav-container {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                max-width: 1400px;
                margin: 0 auto;
            }}
            
            .logo-circle {{
                width: 86px;
                height: 86px;
                border-radius: 50%;
                overflow: hidden;
                box-shadow: 0 8px 25px rgba(0,0,0,0.25), 0 0 0 3px rgba(255,255,255,0.3);
                background: white;
                display: flex;
                align-items: center;
                justify-content: center;
                transition: transform 0.3s, box-shadow 0.3s;
                cursor: pointer;
            }}
            .logo-circle:hover {{
                transform: scale(1.02);
                box-shadow: 0 12px 30px rgba(0,0,0,0.3);
            }}
            .logo-circle img {{
                width: 100%;
                height: 100%;
                object-fit: cover;
            }}
            
            .nav-links {{
                display: flex;
                gap: 0.5rem;
                align-items: center;
                background: rgba(0,0,0,0.35);
                padding: 5px 15px;
                border-radius: 50px;
                backdrop-filter: blur(5px);
            }}
            .nav-links a {{
                text-decoration: none;
                color: white;
                font-weight: 500;
                font-size: 0.9rem;
                padding: 8px 16px;
                border-radius: 40px;
                transition: all 0.3s;
                cursor: pointer;
            }}
            .nav-links a:hover {{
                background: #f59e0b;
                color: #0f3b5c;
            }}
            
            /* ===== DIVIDER VÀ LANGUAGE SWITCH ===== */
            .nav-links .nav-divider {{
                color: rgba(255,255,255,0.4);
                margin: 0 5px;
                font-size: 14px;
            }}
            .lang-switch {{
                display: inline-flex;
                align-items: center;
                gap: 5px;
                margin-left: 5px;
            }}
            .lang-link {{
                text-decoration: none !important;
                color: white !important;
                font-weight: 500;
                font-size: 0.85rem;
                padding: 8px 8px !important;
                border-radius: 40px;
                transition: all 0.3s;
                background: transparent !important;
                cursor: pointer;
            }}
            .lang-link:hover {{
                background: #f59e0b !important;
                color: #0f3b5c !important;
            }}
            .lang-link.active {{
                background: #f59e0b !important;
                color: #0f3b5c !important;
            }}
            .lang-sep {{
                color: rgba(255,255,255,0.5);
                font-size: 12px;
            }}
            
            /* ===== MOBILE LANGUAGE ===== */
            .mobile-lang {{
                position: fixed;
                top: 15px;
                right: 15px;
                z-index: 10001;
                background: rgba(0,0,0,0.6);
                backdrop-filter: blur(8px);
                border-radius: 30px;
                padding: 6px 12px;
                display: none;
                gap: 8px;
                border: 1px solid rgba(255,255,255,0.2);
            }}
            .mobile-lang a {{
                color: white;
                text-decoration: none;
                font-size: 12px;
                font-weight: 600;
                padding: 4px 8px;
                border-radius: 20px;
                transition: all 0.2s;
                cursor: pointer;
            }}
            .mobile-lang a:hover {{
                background: rgba(255,255,255,0.2);
            }}
            .mobile-lang a.active {{
                background: #f59e0b;
                color: #0f3b5c;
            }}
            
            .dropdown {{
                position: relative;
            }}
            .dropdown-content {{
                display: none;
                position: absolute;
                background: white;
                min-width: 200px;
                box-shadow: 0 8px 16px rgba(0,0,0,0.1);
                border-radius: 8px;
                padding: 0.5rem 0;
                top: 100%;
                left: 0;
                z-index: 1;
            }}
            .dropdown:hover .dropdown-content {{
                display: block;
            }}
            .dropdown-content a {{
                color: #333 !important;
                padding: 8px 16px;
                display: block;
                font-size: 0.85rem;
                background: transparent;
                cursor: pointer;
            }}
            .dropdown-content a:hover {{
                background: #f8fafc;
                color: #f59e0b !important;
            }}
            
            /* ===== HERO SLIDER ===== */
            .hero-slider {{
                height: 550px;
                width: 100%;
                position: relative;
                overflow: hidden;
                margin-top: 0;
                background-color: #0a2a3a;
            }}
            .slides-container {{
                position: relative;
                width: 100%;
                height: 100%;
            }}
            .slide {{
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                opacity: 0;
                transition: opacity 0.8s ease-in-out;
                background-color: #0a2a3a;
            }}
            .slide.active {{
                opacity: 1;
                z-index: 2;
            }}
            .slide-layout {{
                display: flex;
                width: 100%;
                height: 100%;
                align-items: center;
                justify-content: center;
            }}
            .slide-image {{
                flex: 1;
                height: 100%;
                background-size: cover;
                background-position: center center;
                background-repeat: no-repeat;
            }}
            .slide-content {{
                flex: 1;
                padding: 50px 40px;
                color: white;
                z-index: 3;
                text-align: center;
            }}
            .slide-content h1 {{
                font-size: 2.8rem;
                font-weight: 800;
                margin-bottom: 1.2rem;
                letter-spacing: -0.5px;
                text-shadow: 0 2px 5px rgba(0,0,0,0.4);
                line-height: 1.3;
            }}
            .slide-content p {{
                font-size: 1.15rem;
                margin-bottom: 0.8rem;
                line-height: 1.5;
                text-shadow: 0 1px 3px rgba(0,0,0,0.3);
            }}
            .slide-content .highlight {{
                font-size: clamp(0.85rem, 1.1vw, 1.2rem);
                font-weight: 700;
                color: #f59e0b;
                margin-top: 1.2rem;
                display: inline-block;
                background: rgba(0,0,0,0.25);
                padding: 6px 18px;
                border-radius: 40px;
                backdrop-filter: blur(4px);
                white-space: nowrap;
            }}
            .slider-progress {{
                position: absolute;
                bottom: 0;
                left: 0;
                height: 4px;
                background: #f59e0b;
                width: 0%;
                z-index: 10;
                transition: width 0.1s linear;
            }}
            .slider-nav {{
                position: absolute;
                bottom: 20px;
                left: 50%;
                transform: translateX(-50%);
                display: flex;
                gap: 15px;
                z-index: 10;
            }}
            .slider-dot {{
                width: 12px;
                height: 12px;
                border-radius: 50%;
                background: rgba(255,255,255,0.5);
                cursor: pointer;
                transition: all 0.3s;
            }}
            .slider-dot.active {{
                background: #f59e0b;
                width: 30px;
                border-radius: 10px;
            }}
            .slider-arrow {{
                position: absolute;
                top: 50%;
                transform: translateY(-50%);
                z-index: 10;
                background: rgba(255,255,255,0.15);
                border: 2px solid rgba(255,255,255,0.4);
                color: white;
                width: 45px;
                height: 45px;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                cursor: pointer;
                font-size: 1.2rem;
                transition: all 0.3s;
                backdrop-filter: blur(4px);
            }}
            .slider-arrow:hover {{
                background: #f59e0b;
                border-color: #f59e0b;
                color: #1e293b;
            }}
            .slider-arrow.prev {{ left: 20px; }}
            .slider-arrow.next {{ right: 20px; }}
            
            /* ===== SCROLL REVEAL ===== */
            .reveal {{
                opacity: 0;
                transform: translateY(30px);
                transition: opacity 0.7s ease, transform 0.7s ease;
            }}
            .reveal.visible {{
                opacity: 1;
                transform: translateY(0);
            }}
            
            /* ===== STATS SECTION ===== */
            .stats-section {{
                padding: 60px 30px;                                 
                background: #0f3b5c;
                color: white;
            }}
            .stats-grid {{
                display: flex;
                justify-content: space-between;
                max-width: 1200px;
                margin: 0 auto;
                gap: 0;
                flex-wrap: nowrap;
            }}
            .stat-card {{
                text-align: center;
                flex: 1;
                min-width: 0;
                padding: 28px 12px;
                border-right: 1px solid rgba(255,255,255,0.2);
                transition: transform 0.3s;
            }}
            .stat-card:hover {{ transform: translateY(-5px); }}
            .stat-card:last-child {{ border-right: none; }}
            .stat-number {{
                font-size: clamp(1.4rem, 2.2vw, 2.4rem);
                font-weight: 800;
                color: #f59e0b;
                margin-bottom: 8px;
                white-space: nowrap;
                line-height: 1.2;
            }}
            .stat-label {{
                font-size: clamp(0.7rem, 1vw, 0.85rem);
                text-transform: uppercase;
                letter-spacing: 1px;
                font-weight: 500;
                white-space: nowrap;
            }}
            
            /* ===== ABOUT & SERVICES ===== */
            .about-section {{
                padding: 80px 30px;                                                       
                background: #f8fafc;
            }}
            .about-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 60px;
                max-width: 1280px;
                margin: 0 auto;
            }}
            .about-tag {{
                color: #f59e0b;
                font-weight: 700;
                letter-spacing: 2px;
                margin-bottom: 1rem;
                font-size: 0.8rem;
            }}
            .about-title {{
                font-size: 2.5rem;
                font-weight: 700;
                color: #0f3b5c;
                margin-bottom: 1.5rem;
            }}
            .about-text {{
                color: #475569;
                line-height: 1.7;
                margin-bottom: 1.5rem;
            }}
            .about-highlight {{
                background: white;
                padding: 20px;
                border-radius: 16px;
                border-left: 4px solid #f59e0b;
            }}
            .about-img {{
                width: 100%;
                border-radius: 24px;
                box-shadow: 0 20px 30px -15px rgba(0,0,0,0.15);
            }}
            .services-section {{
                padding: 80px 30px;                                         
                background: white;
            }}
            .section-header {{
                text-align: center;
                margin-bottom: 50px;
            }}
            .section-header h2 {{
                font-size: 2.2rem;
                color: #0f3b5c;
            }}
            .services-grid {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 30px;
                max-width: 1280px;
                margin: 0 auto;
            }}
            .service-card {{
                background: white;
                padding: 30px 20px;
                border-radius: 20px;
                text-align: center;
                box-shadow: 0 4px 15px rgba(0,0,0,0.05);
                border: 1px solid #e2e8f0;
                transition: all 0.3s;
            }}
            .service-card:hover {{
                transform: translateY(-8px);
                border-color: #f59e0b;
            }}
            .service-icon {{
                font-size: 3rem;
                color: #f59e0b;
                margin-bottom: 20px;
            }}
            .infra-section {{
                padding: 80px 30px;                                         
                background: #f8fafc;
            }}
            .infra-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 50px;
                max-width: 1280px;
                margin: 0 auto;
            }}
            .infra-feature {{
                display: flex;
                gap: 15px;
                margin-bottom: 25px;
            }}
            .infra-feature i {{
                font-size: 1.8rem;
                color: #f59e0b;
            }}
            .careers-section {{
                padding: 80px 30px;                                             
                background: linear-gradient(135deg, #0f3b5c 0%, #1e4a76 100%);
                color: white;
                text-align: center;
            }}
            .btn-white {{
                background: white;
                color: #0f3b5c;
                padding: 12px 35px;
                border-radius: 40px;
                font-weight: 700;
                display: inline-block;
                margin-top: 20px;
                text-decoration: none;
                cursor: pointer;
            }}
            
            /* ===== FOOTER ===== */
            .footer {{
                background: #0f172a;
                color: #cbd5e1;
                padding: 50px 30px 30px;                                        
                width: 100%;
                clear: both;
            }}
            .footer-grid {{
                display: grid;
                grid-template-columns: repeat(4, 1fr);
                gap: 40px;
                max-width: 1280px;
                margin: 0 auto;
                padding-bottom: 40px;
                border-bottom: 1px solid #334155;
            }}
            .footer-col h4 {{
                color: white;
                margin-bottom: 20px;
                font-size: 1.1rem;
            }}
            .footer-col p, .footer-col a {{
                color: #94a3b8;
                text-decoration: none;
                line-height: 1.8;
                font-size: 0.9rem;
                display: block;
                cursor: pointer;
            }}
            .footer-col a:hover {{
                color: #f59e0b;
            }}
            .copyright {{
                text-align: center;
                padding-top: 30px;
                font-size: 0.8rem;
                color: #64748b;
            }}
            
            /* ===== MODAL ===== */
            .modal {{
                display: none;
                position: fixed;
                top: 0;
                left: 0;
                width: 100vw;
                height: 100vh;
                z-index: 99999;
                background: rgba(0,0,0,0.85);
                align-items: flex-start;
                justify-content: center;
                overflow-y: auto;
                padding: 80px 20px 20px 20px;
            }}
            .modal.active {{
                display: flex;
            }}
            .modal-content {{
                max-width: 900px;
                width: 100%;
                background: white;
                border-radius: 8px;
                box-shadow: 0 25px 60px rgba(0,0,0,0.4);
                overflow: hidden;
                animation: modalFadeIn 0.3s ease;
            }}
            @keyframes modalFadeIn {{
                from {{ opacity: 0; transform: scale(0.95); }}
                to {{ opacity: 1; transform: scale(1); }}
            }}
            .modal-header {{
                display: flex;
                justify-content: flex-end;
                padding: 10px 15px;
                background: #f0f2f5;
                border-bottom: 1px solid #ddd;
            }}
            .modal-close {{
                background: none;
                border: none;
                font-size: 24px;
                cursor: pointer;
                color: #666;
                transition: color 0.2s;
            }}
            .modal-close:hover {{
                color: #f59e0b;
            }}
            .modal-body {{
                padding: 30px 40px;
                max-height: 80vh;
                overflow-y: auto;
            }}
            .a4-chairman {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 20px;
                margin-bottom: 30px;
                padding: 20px;
                background: #f8f9fa;
                border-radius: 16px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            }}
            .a4-chairman-info {{
                flex: 2;
            }}
            .a4-chairman-avatar {{
                flex: 0 0 auto;
                width: 150px;
                height: 150px;
            }}
            .a4-chairman-avatar img {{
                width: 100%;
                height: 100%;
                border-radius: 50%;
                object-fit: cover;
                border: 3px solid #f59e0b;
                box-shadow: 0 4px 15px rgba(0,0,0,0.15);
            }}
            .a4-chairman-logo {{
                flex: 0 0 auto;
                width: 80px;
                height: 80px;
            }}
            .a4-chairman-logo img {{
                width: 100%;
                height: 100%;
                border-radius: 50%;
                object-fit: cover;
                border: 2px solid #ddd;
                background: white;
                padding: 5px;
            }}
            .a4-chairman-info h2 {{
                font-size: 1.3rem;
                color: #0f3b5c;
                margin-bottom: 5px;
            }}
            .a4-chairman-info .title {{
                color: #f59e0b;
                font-weight: 600;
                margin-bottom: 8px;
            }}
            .a4-chairman-info .company {{
                font-size: 0.8rem;
                color: #666;
                line-height: 1.4;
                font-weight: 500;
            }}
            .a4-body {{
                line-height: 1.7;
                color: #333;
            }}
            .modal-body .a4-body p {{
                text-align: justify;
                text-justify: inter-ideograph;
            }}
            .a4-date {{
                text-align: right !important;
                font-style: italic;
                margin-bottom: 20px;
                color: #666;
            }}
            .vision-box, .mission-box {{
                background: #f0f7ff;
                padding: 20px;
                border-radius: 12px;
                margin: 20px 0;
                border-left: 4px solid #f59e0b;
            }}
            .vision-box h3, .mission-box h3 {{
                color: #0f3b5c;
                margin-bottom: 12px;
                font-size: 1.1rem;
            }}
            .a4-signature-left {{
                margin-top: 40px;
                text-align: left;
            }}
            .sig-block-left {{
                display: inline-block;
                text-align: center;
            }}
            .sig-block-left .sig-image img {{
                max-width: 386px;
                height: auto;
            }}
            .a4-footer {{
                margin-top: 30px;
                padding-top: 15px;
                border-top: 1px solid #ddd;
                text-align: center;
                font-size: 0.75rem;
                color: #999;
                display: flex;
                justify-content: space-between;
                flex-wrap: wrap;
            }}
            
            /* ===== RESPONSIVE ===== */
            @media (max-width: 768px) {{
                .slide-layout {{
                    flex-direction: column;
                }}
                .slide-image {{
                    width: 100%;
                    height: 35%;
                    flex: none;
                }}
                .slide-content {{
                    padding: 25px 20px;
                }}
                .slide-content h1 {{
                    font-size: 1.6rem;
                }}
                .slide-content p {{
                    font-size: 0.9rem;
                }}
                .hero-slider {{
                    height: auto;
                    min-height: 480px;
                }}
                .stats-grid, .about-grid, .services-grid, .infra-grid, .footer-grid {{
                    grid-template-columns: 1fr;
                }}
                .stat-card {{
                    border-right: none;
                    border-bottom: 1px solid rgba(255,255,255,0.2);
                }}
                .nav-links {{
                    display: none;
                }}
                .logo-circle {{
                    width: 60px;
                    height: 60px;
                }}
                .a4-chairman {{
                    flex-wrap: wrap;
                    justify-content: center;
                    text-align: center;
                }}
                .modal-body {{
                    padding: 20px;
                }}
                .mobile-lang {{
                    display: flex !important;
                }}
                .nav-links .lang-switch {{
                    display: none;
                }}
            }}
        </style>
    </head>
    <body>
    
    <!-- Navigation -->
    <nav class="navbar" id="navbar">
        <div class="nav-container">
            <div class="logo-circle">
                <img src="data:image/png;base64,{logo_base64}" alt="Cảng Hòn La">
            </div>
            <div class="nav-links">
                <a class="nav-link" data-section="home">{text.get('nav_home', 'Trang chủ')}</a>
                <div class="dropdown">
                    <a class="nav-link" data-section="about">{text.get('nav_about', 'Giới thiệu')} <i class="fas fa-chevron-down"></i></a>
                    <div class="dropdown-content">
                        <a class="nav-link" data-section="about">{text.get('about_us', 'Về chúng tôi')}</a>
                        <a href="#" id="thuNgoBtn">{text.get('chairman_letter', 'Thư ngỏ của Chủ tịch HĐQT')}</a>
                    </div>
                </div>
                <a class="nav-link" data-section="services">{text.get('nav_services', 'Dịch vụ')}</a>
                <a class="nav-link" data-section="infrastructure">{text.get('nav_infrastructure', 'Vị trí & Hạ tầng')}</a>
                <a class="nav-link" data-section="careers">{text.get('nav_careers', 'Tuyển dụng')}</a>
                <a class="nav-link" data-section="contact">{text.get('nav_contact', 'Liên hệ')}</a>
                <span class="nav-divider">|</span>
                <div class="lang-switch">
                    <a href="#" class="lang-link {vi_active}" onclick="switchLanguage('vi'); return false;">🇻🇳 VI</a>
                    <span class="lang-sep">/</span>
                    <a href="#" class="lang-link {en_active}" onclick="switchLanguage('en'); return false;">🇬🇧 EN</a>
                </div>
            </div>
        </div>
    </nav>

    <!-- Mobile Language Switcher -->
    <div class="mobile-lang">
        <a href="#" class="{vi_active}" onclick="switchLanguage('vi'); return false;">🇻🇳 VI</a>
        <span style="color:white; opacity:0.5;">|</span>
        <a href="#" class="{en_active}" onclick="switchLanguage('en'); return false;">🇬🇧 EN</a>
    </div>

    <!-- Modal Thư ngỏ -->
    <div id="thuNgoModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <button class="modal-close" id="closeModalBtn">✕</button>
            </div>
            <div class="modal-body">
                <div class="a4-chairman">
                    <div class="a4-chairman-avatar">
                        <img src="{chu_tich_img}" alt="Chủ tịch HĐQT">
                    </div>
                    <div class="a4-chairman-info">
                        <h2>Ông Phùng Gia Phát</h2>
                        <p class="title">{text.get('modal_chairman_title', 'Chủ tịch Hội đồng Quản trị')}</p>
                        <p class="company">Công ty Cổ phần Cảng Hòn La</p>
                        <p class="company">Khu kinh tế Hòn La, Xã Quảng Đông, Huyện Quảng Trạch, Tỉnh Quảng Bình</p>
                    </div>                  
                    <div class="a4-chairman-logo">
                        <img src="data:image/png;base64,{logo_base64}" alt="Logo Cảng Hòn La">
                    </div>
                </div>
                <div class="a4-body">
                    <p class="a4-date">Quảng Bình, ngày 21 tháng 3 năm 2025</p>
                    <p class="a4-greeting" style="font-weight: bold; font-size: 1rem;">{text.get('modal_greeting', 'Kính gửi: Quý đối tác, nhà đầu tư và toàn thể cán bộ nhân viên,')}</p>
                    <p>{text.get('modal_content_1', 'Với niềm tự hào sâu sắc, Tôi xin thay mặt Hội đồng Quản trị Công ty Cổ phần Cảng Hòn La gửi lời chào trân trọng nhất đến Quý đối tác, nhà đầu tư và toàn thể cán bộ nhân viên — những người đã và đang đồng hành cùng chúng tôi trên hành trình kiến tạo một cảng biển tầm cỡ quốc tế giữa lòng đất nước Việt Nam.')}</p>
                    <p>{text.get('modal_content_2', 'Ngày 21 tháng 3 năm 2025 là một mốc son lịch sử — ngày chính thức khởi công Dự án Cảng tổng hợp quốc tế Hòn La, dự án được Chính phủ công nhận là Dự án trọng điểm Quốc gia. Đây không chỉ là thành quả của nhiều năm nỗ lực không ngừng, mà còn là khởi đầu của một chương mới trong lịch sử phát triển kinh tế hàng hải miền Trung Việt Nam.')}</p>
                    <div class="vision-box">
                        <h3>{text.get('modal_vision_title', '🎯 Tầm nhìn — Vision 2035')}</h3>
                        <p>{text.get('modal_vision_text', 'Trở thành cảng biển quốc tế hiện đại hàng đầu Đông Nam Á trên tuyến hành lang kinh tế Đông–Tây (EWEC) — nơi kết nối Việt Nam với thế giới, thúc đẩy thương mại, logistic và du lịch tàu biển, đóng góp thiết thực vào chiến lược phát triển kinh tế biển bền vững của Việt Nam đến năm 2035 và tầm nhìn 2045.')}</p>
                    </div>
                    <p>{text.get('modal_content_3', 'Với vị trí địa chiến lược độc đáo, hệ thống hạ tầng quy mô 39,22 ha, năng lực tiếp nhận tàu trọng tải lên đến 70.000 DWT và tàu du lịch quốc tế 225.000 GT, Cảng tổng hợp quốc tế Hòn La sẽ là cửa ngõ hàng hải chiến lược, cầu nối giữa các nền kinh tế trong khu vực và toàn cầu.')}</p>
                    <div class="mission-box">
                        <h3>{text.get('modal_mission_title', '💡 Sứ mệnh - Nhắn gửi đến mỗi thành viên')}</h3>
                        <p>{text.get('modal_mission_text', 'Mỗi cán bộ nhân viên của Công ty cổ phần Cảng Hòn La là một đại sứ của sự chuyên nghiệp và tận tâm. Sứ mệnh của chúng ta là xây dựng một môi trường làm việc đẳng cấp, nơi năng lực được trọng dụng, sáng tạo được khuyến khích và mỗi cá nhân đều tự hào khi đặt bàn tay mình vào công trình lịch sử này. Hãy làm việc với trái tim của người kiến tạo — bởi di sản chúng ta để lại không chỉ là những cầu bến vững chắc, mà còn là những thế hệ nhân lực xuất sắc của đất nước.')}</p>
                    </div>
                    <p>{text.get('modal_content_4', 'Chúng tôi hiểu rằng con đường phía trước còn không ít thách thức. Song Tôi tin tưởng sâu sắc rằng với trí tuệ tập thể, khí phách dân tộc và khát vọng vươn ra biển lớn, Cảng tổng hợp quốc tế Hòn La sẽ hoàn thành xuất sắc sứ mệnh lịch sử được giao phó.')}</p>
                    <p>{text.get('modal_thanks', 'Xin trân trọng cảm ơn sự tin tưởng, đồng hành và cống hiến của tất cả Quý vị.')}<br>{text.get('modal_wishes', 'Chúc Quý đối tác thịnh vượng, toàn thể cán bộ nhân viên sức khỏe và thành công!')}</p>
                </div>
                <div class="a4-signature-left">
                    <div class="sig-block-left">
                        <div class="sig-image">
                            <img src="{chu_ky_img}" alt="Chữ ký Chủ tịch">
                        </div>
                    </div>
                </div>
                <div class="a4-footer">
                    <span>🌐 honlaport.com.vn</span>
                    <span>✉ info@honlaport.com.vn</span>
                    <span>📞 0232.xxxx.xxx</span>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Hero Slider -->
    <section id="home" class="hero-slider">
        <div class="slides-container">
            <div class="slide active">
                <div class="slide-layout">
                    <div class="slide-image" style="background-image: url('{slide1_src}');"></div>
                    <div class="slide-content">
                        <h1>{text.get('hero_title_1', 'CẢNG TỔNG HỢP QUỐC TẾ HÒN LA')}</h1>
                        <p>{text.get('hero_desc_1_1', 'Chính thức khởi công ngày 21 tháng 3 năm 2025')}</p>
                        <p>{text.get('hero_desc_1_2', 'Đưa vào khai thác từ Tháng 5 năm 2026')}</p>
                        <div class="highlight">{text.get('hero_tag_1', '🚢 Cửa ngõ hàng hải chiến lược của Miền Trung')}</div>
                    </div>
                </div>
            </div>
            <div class="slide">
                <div class="slide-layout">
                    <div class="slide-content">
                        <h1>{text.get('hero_title_2', 'KẾT NỐI TOÀN CẦU')}</h1>
                        <p>{text.get('hero_desc_2_1', 'Vị trí chiến lược trên tuyến hành lang kinh tế Đông - Tây (EWEC)')}</p>
                        <p>{text.get('hero_desc_2_2', 'Kết nối trực tiếp với các cảng biển lớn trong khu vực và quốc tế')}</p>
                        <div class="highlight">{text.get('hero_tag_2', '🌏 Hành lang thương mại huyết mạch')}</div>
                    </div>
                    <div class="slide-image" style="background-image: url('{slide2_src}');"></div>
                </div>
            </div>
            <div class="slide">
                <div class="slide-layout">
                    <div class="slide-image" style="background-image: url('{slide3_src}');"></div>
                    <div class="slide-content">
                        <h1>{text.get('hero_title_3', 'HẠ TẦNG ĐẲNG CẤP QUỐC TẾ')}</h1>
                        <p>{text.get('hero_desc_3_1', '04 Cầu Tàu | Tổng chiều dài 970m | Tiếp nhận tàu 70.000 DWT')}</p>
                        <p>{text.get('hero_desc_3_2', 'Tàu du lịch quốc tế 225.000 GT')}</p>
                        <div class="highlight">{text.get('hero_tag_3', '⚓ Hiện đại - Đồng bộ - Chuyên nghiệp')}</div>
                    </div>
                </div>
            </div>
        </div>
        <div class="slider-arrow prev" id="prevBtn">&#8592;</div>
        <div class="slider-arrow next" id="nextBtn">&#8594;</div>
        <div class="slider-nav">
            <div class="slider-dot active" data-slide="0"></div>
            <div class="slider-dot" data-slide="1"></div>
            <div class="slider-dot" data-slide="2"></div>
        </div>
        <div class="slider-progress" id="sliderProgress"></div>
    </section>
    
    <!-- Statistics -->
    <section id="stats" class="stats-section">
        <div class="stats-grid">
            <div class="stat-card reveal"><div class="stat-number">39,22 ha</div><div class="stat-label">{text.get('stat_total_area', 'Tổng diện tích')}</div></div>
            <div class="stat-card reveal"><div class="stat-number">70.000 DWT</div><div class="stat-label">{text.get('stat_max_capacity', 'Trọng tải tàu tối đa')}</div></div>
            <div class="stat-card reveal"><div class="stat-number">970 m</div><div class="stat-label">{text.get('stat_berth_length', 'Chiều dài cầu cảng')}</div></div>
            <div class="stat-card reveal"><div class="stat-number">225.000 GT</div><div class="stat-label">{text.get('stat_cruise_ship', 'Tàu du lịch quốc tế')}</div></div>
        </div>
    </section>
    
    <!-- About -->
    <section id="about" class="about-section">
        <div class="about-grid">
            <div>
                <div class="about-tag">{text.get('about_tag', 'CHÀO MỪNG ĐẾN VỚI CẢNG QUỐC TẾ HÒN LA')}</div>
                <h2 class="about-title">{text.get('about_title', 'Cửa ngõ hàng hải chiến lược của Miền Trung')}</h2>
                <p class="about-text">{text.get('about_text', 'Cảng tổng hợp Quốc tế Hòn La được đầu tư bài bản với hệ thống cơ sở hạ tầng đồng bộ, hiện đại, đáp ứng nhu cầu bốc xếp hàng hóa, trung chuyển container và đón tàu du lịch quốc tế.')}</p>
                <div class="about-highlight"><i class="fas fa-trophy" style="color:#f59e0b"></i> <strong>{text.get('about_highlight', 'Dự án trọng điểm Quốc gia')}</strong></div>
            </div>
            <div><img src="https://images.unsplash.com/photo-1562329264-a2c2d4112b8d?q=80&w=2070" class="about-img"></div>
        </div>
    </section>
    
    <!-- Services -->
    <section id="services" class="services-section">
        <div class="section-header"><h2>{text.get('services_title', 'Dịch vụ của chúng tôi')}</h2></div>
        <div class="services-grid">
            <div class="service-card"><div class="service-icon"><i class="fas fa-ship"></i></div><h3>{text.get('service_bulk', 'Hàng rời & Hàng khô')}</h3></div>
            <div class="service-card"><div class="service-icon"><i class="fas fa-boxes"></i></div><h3>{text.get('service_container', 'Hàng container')}</h3></div>
            <div class="service-card"><div class="service-icon"><i class="fas fa-umbrella-beach"></i></div><h3>{text.get('service_cruise', 'Du lịch tàu biển')}</h3></div>
            <div class="service-card"><div class="service-icon"><i class="fas fa-warehouse"></i></div><h3>{text.get('service_logistics', 'Logistics & Kho bãi')}</h3></div>
        </div>
    </section>
    
    <!-- Infrastructure -->
    <section id="infrastructure" class="infra-section">
        <div class="infra-grid">
            <div>
                <div class="about-tag">{text.get('infra_tag', 'HẠ TẦNG & VỊ TRÍ')}</div>
                <h2 class="about-title">{text.get('infra_title', 'Vị thế vàng trên bản đồ logistics')}</h2>
                <div class="infra-feature"><i class="fas fa-map-marker-alt"></i><div><strong>Quảng Trạch, Quảng Bình</strong><br>{text.get('infra_location', 'Khu kinh tế Hòn La')}</div></div>
                <div class="infra-feature"><i class="fas fa-road"></i><div><strong>{text.get('infra_connection', 'Kết nối hành lang Đông - Tây (EWEC)')}</strong></div></div>
                <div class="infra-feature"><i class="fas fa-anchor"></i><div><strong>{text.get('infra_berths', '04 bến cấp tàu')}</strong><br>Tổng chiều dài 970m</div></div>
            </div>
            <div><img src="https://images.unsplash.com/photo-1578575437130-527eed3abbec?q=80&w=2070" class="about-img"></div>
        </div>
    </section>
    
    <!-- Careers -->
    <section id="careers" class="careers-section">
        <h2>{text.get('careers_title', 'GIA NHẬP ĐỘI NGŨ NHÂN SỰ CỦA CHÚNG TÔI')}</h2>
        <p>{text.get('careers_subtitle', 'Chúng tôi luôn tìm kiếm những nhân tài')}</p>
        <a href="#" class="btn-white" id="careerLink">{text.get('careers_button', '📢 Xem cơ hội việc làm tại đây')}</a>
    </section>
    
    <!-- Footer -->
    <footer id="contact" class="footer">
        <div class="footer-grid">
            <div class="footer-col"><h4 style="font-size:0.95rem; white-space:nowrap;">{text.get('footer_company', 'CÔNG TY CỔ PHẦN CẢNG HÒN LA')}</h4><p>Khu kinh tế Hòn La, Xã Phú Trạch, Tỉnh Quảng Trị</p><p>📞 0232.xxxx.xxx</p><p>📧 info@honlaport.com.vn</p></div>
            <div class="footer-col"><h4>{text.get('footer_quick_links', 'Liên kết nhanh')}</h4>
                <a class="nav-link" data-section="home">{text.get('nav_home', 'Trang chủ')}</a>
                <a class="nav-link" data-section="about">{text.get('nav_about', 'Về chúng tôi')}</a>
                <a class="nav-link" data-section="services">{text.get('nav_services', 'Dịch vụ')}</a>
                <a class="nav-link" data-section="infrastructure">{text.get('nav_infrastructure', 'Hạ tầng')}</a>
                <a class="nav-link" data-section="careers">{text.get('nav_careers', 'Tuyển dụng')}</a>
            </div>
            <div class="footer-col"><h4>{text.get('footer_support', 'Hỗ trợ')}</h4><a href="#">{text.get('footer_faq', 'Câu hỏi thường gặp')}</a><a href="#">{text.get('footer_privacy', 'Chính sách bảo mật')}</a><a href="#">{text.get('footer_terms', 'Điều khoản sử dụng')}</a></div>
            <div class="footer-col"><h4>{text.get('footer_working_hours', 'Giờ làm việc')}</h4><p>🚢 {text.get('footer_working_hours_port', 'Bến cảng: 24/7')}</p><p>🏢 {text.get('footer_working_hours_office', 'Văn phòng: 7:30 - 17:00')}</p><p>📅 {text.get('footer_working_days', 'Thứ 2 - Thứ 7')}</p></div>
        </div>
        <div class="copyright">
            <p>© 2026 - Công ty Cổ phần Cảng Hòn La. All rights reserved.</p>
            <p style="margin-top: 10px;">{text.get('footer_copyright', 'PHÁT TRIỂN BỀN VỮNG - KẾT NỐI TOÀN CẦU')}</p>
        </div>
    </footer>
    
    {landing_js}
    </body>
    </html>
    """
    
    # Render landing page
    components.html(landing_html, height=3150, scrolling=False)
    
    # Nút HRM dùng components.html (giữ nguyên phần còn lại)
    hrm_html = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    background: linear-gradient(135deg, #0f3b5c 0%, #1a4a6e 100%);
    display: flex; justify-content: center; align-items: center;
    min-height: 100px;
    border-top: 3px solid #f59e0b;
    border-bottom: 3px solid #f59e0b;
    padding: 20px;
}
.hrm-button {
    background: linear-gradient(135deg, #f59e0b 0%, #e67e22 100%);
    color: #0f3b5c; font-weight: 800; font-size: 1.2rem;
    border: none; border-radius: 60px; padding: 18px 60px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.3); letter-spacing: 1px;
    cursor: pointer; transition: all 0.3s ease; min-width: 420px;
    font-family: sans-serif;
}
.hrm-button:hover {
    background: linear-gradient(135deg, #e67e22 0%, #d35400 100%);
    transform: translateY(-3px); box-shadow: 0 12px 30px rgba(0,0,0,0.4);
}
@media (max-width: 768px) {
    .hrm-button { font-size: 0.9rem; padding: 14px 30px; min-width: 260px; }
}
</style>
</head>
<body>
    <button class="hrm-button" id="hrmBtn">
        🔐 HRM - QUẢN LÝ NHÂN SỰ / Chỉ dành cho Nhân viên
    </button>
    <script>
    document.getElementById('hrmBtn').addEventListener('click', function() {
        var topWin = window.top || window.parent || window;
        var url = new URL(topWin.location.href);
        url.searchParams.set('goto', 'hrm');
        topWin.location.href = url.toString();
    });
    </script>
</body>
</html>"""
 
    st.markdown("""
        <style>
            .hrm-button-container {
                background: linear-gradient(135deg, #0f3b5c 0%, #1a4a6e 100%);
                border-top: 3px solid #f59e0b;
                border-bottom: 3px solid #f59e0b;
                padding: 20px;
                text-align: center;
            }
            .stButton > button {
                background: linear-gradient(135deg, #f59e0b 0%, #e67e22 100%);
                color: #0f3b5c !important;
                font-weight: 800;
                font-size: 1.2rem;
                border: none;
                border-radius: 60px;
                padding: 18px 60px;
                box-shadow: 0 8px 25px rgba(0,0,0,0.3);
                min-width: 420px;
                transition: all 0.3s ease;
                width: auto !important;
            }
            .stButton > button:hover {
                background: linear-gradient(135deg, #e67e22 0%, #d35400 100%);
                transform: translateY(-3px);
                box-shadow: 0 12px 30px rgba(0,0,0,0.4);
            }
            @media (max-width: 768px) {
                .stButton > button {
                    font-size: 0.9rem;
                    padding: 14px 30px;
                    min-width: 260px;
                }
            }
        </style>
        <div class="hrm-button-container">
    """, unsafe_allow_html=True)

    # Nút HRM thuần Python
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🔐 HRM - QUẢN LÝ NHÂN SỰ / Chỉ dành cho Nhân viên", width='stretch'):
            st.session_state.show_hrm = True
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


st.set_page_config(page_title="HRM-Port", page_icon="🏗️", layout="wide")

# Gọi định danh tenant
resolve_tenant()

# Nạp động cấu hình thương hiệu (Branding) từ tenant
if st.session_state.get('tenant'):
    tenant_data = st.session_state.tenant
    mapping = {
        "ten_cong_ty": "ten_cty",
        "dai_dien": "dai_dien",
        "chuc_vu": "chuc_vu",
        "ma_so_thue": "ma_so_thue",
        "dien_thoai_cty": "dien_thoai_cty",
        "ma_don_vi_BHXH": "ma_don_vi_BHXH",
        "ma_vung_luong": "ma_vung_luong",
        "dia_chi": "dia_chi",
        "loi_nhan_zalo": "loi_nhan_zalo",
        "zalo_group_link": "zalo_group_link",
        "zalo_group_name": "zalo_group_name",
        "logo_url": "logo_url"
    }
    for config_key, tenant_key in mapping.items():
        if tenant_key in tenant_data and tenant_data[tenant_key]:
            COMPANY_CONFIG[config_key] = tenant_data[tenant_key]

    # Banner chế độ Demo: dữ liệu dùng chung, chỉ xem thử — không lưu/sửa/xoá được
    if str(tenant_data.get('ma_cty', '')).upper() == 'DEMO':
        st.info(
            "🧪 **Chế độ Demo** — Bạn đang xem dữ liệu mẫu dùng chung. "
            "Mọi thao tác Lưu/Sửa/Xoá sẽ bị chặn để bảo vệ dữ liệu chung. "
            "Liên hệ để đăng ký dùng thử với dữ liệu riêng của công ty bạn.",
            icon="🧪"
        )


# ========== XỬ LÝ ĐA NGÔN NGỮ ==========
def init_language():
    """Khởi tạo và xử lý chuyển đổi ngôn ngữ"""
    if 'language' not in st.session_state:
        st.session_state.language = 'vi'
    
    # Kiểm tra query params để đổi ngôn ngữ
    query_params = st.query_params
    if 'lang' in query_params:
        new_lang = query_params['lang']
        if new_lang in ['vi', 'en']:
            st.session_state.language = new_lang
            # Xóa param để tránh lặp
            st.query_params.clear()
            st.rerun()

# Gọi hàm khởi tạo
init_language()

# ========== QUẢN LÝ CÔNG VĂN & HĐ KINH TẾ ==========

# === Hàm khởi tạo bảng nếu chưa có ===
def init_cong_van_tables():
    """Khởi tạo các bảng cho module Quản lý Công văn & HĐ kinh tế"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        
        # Tạo bảng cấu hình công văn
        c.execute("""
            CREATE TABLE IF NOT EXISTS cau_hinh_cong_van (
                id SERIAL PRIMARY KEY,
                loai VARCHAR(20) NOT NULL,
                so_max INTEGER NOT NULL DEFAULT 0,
                prefix VARCHAR(10),
                nam_hien_tai INTEGER NOT NULL DEFAULT EXTRACT(YEAR FROM CURRENT_DATE),
                updated_at TIMESTAMP DEFAULT NOW(),
                UNIQUE(loai, nam_hien_tai)
            )
        """)
        
        # Tạo bảng công văn đến
        c.execute("""
            CREATE TABLE IF NOT EXISTS cong_van_den (
                id SERIAL PRIMARY KEY,
                so_cong_van VARCHAR(50) NOT NULL,
                co_quan_phat_hanh VARCHAR(200) NOT NULL,
                ngay_den DATE NOT NULL DEFAULT CURRENT_DATE,
                tieu_de TEXT NOT NULL,
                trich_yeu TEXT,
                file_url TEXT,
                ghi_chu TEXT,
                nguoi_tao VARCHAR(100),
                created_at TIMESTAMP DEFAULT NOW(),
                updated_at TIMESTAMP DEFAULT NOW()
            )
        """)
        c.execute("ALTER TABLE cong_van_den ADD COLUMN IF NOT EXISTS ma_vach_buu_dien VARCHAR(100)")
        
        # Tạo bảng công văn đi
        c.execute("""
            CREATE TABLE IF NOT EXISTS cong_van_di (
                id SERIAL PRIMARY KEY,
                so_cong_van VARCHAR(50) NOT NULL,
                phong_phat_hanh VARCHAR(100) NOT NULL,
                ngay_phat_hanh DATE NOT NULL DEFAULT CURRENT_DATE,
                tieu_de TEXT NOT NULL,
                trich_yeu TEXT,
                file_url TEXT,
                loai_cong_van VARCHAR(20) NOT NULL,
                ghi_chu TEXT,
                nguoi_tao VARCHAR(100),
                created_at TIMESTAMP DEFAULT NOW(),
                updated_at TIMESTAMP DEFAULT NOW()
            )
        """)
        c.execute("ALTER TABLE cong_van_di ADD COLUMN IF NOT EXISTS ma_vach_buu_dien VARCHAR(100)")
        
        # Tạo bảng hợp đồng kinh tế
        c.execute("""
            CREATE TABLE IF NOT EXISTS hop_dong_kinh_te (
                id SERIAL PRIMARY KEY,
                so_hop_dong VARCHAR(50) NOT NULL,
                ten_doi_tac VARCHAR(200) NOT NULL,
                ngay_ky DATE NOT NULL DEFAULT CURRENT_DATE,
                trich_yeu TEXT,
                file_url TEXT,
                ghi_chu TEXT,
                nguoi_tao VARCHAR(100),
                created_at TIMESTAMP DEFAULT NOW(),
                updated_at TIMESTAMP DEFAULT NOW()
            )
        """)
        
        # Tạo bảng danh mục loại công văn
        c.execute("""
            CREATE TABLE IF NOT EXISTS danh_muc_loai_cong_van (
                id SERIAL PRIMARY KEY,
                ma_loai VARCHAR(10) NOT NULL UNIQUE,
                ten_loai VARCHAR(50) NOT NULL,
                thu_tu INTEGER DEFAULT 0,
                trang_thai BOOLEAN DEFAULT TRUE
            )
        """)
        
        # Insert dữ liệu mặc định cho danh mục loại công văn
        c.execute("""
            INSERT INTO danh_muc_loai_cong_van (ma_loai, ten_loai, thu_tu) VALUES
            ('QĐ', 'Quyết định', 1),
            ('CV', 'Công văn', 2),
            ('BC', 'Báo cáo', 3),
            ('TB', 'Thông báo', 4),
            ('TTr', 'Tờ trình', 5)
            ON CONFLICT (ma_loai) DO NOTHING
        """)
        
        # Tạo bảng cấu hình hệ thống
        c.execute("""
            CREATE TABLE IF NOT EXISTS cau_hinh_he_thong (
                id SERIAL PRIMARY KEY,
                ten_cau_hinh VARCHAR(50) NOT NULL UNIQUE,
                gia_tri VARCHAR(100),
                mo_ta TEXT,
                updated_at TIMESTAMP DEFAULT NOW()
            )
        """)
        
        # Insert cấu hình mặc định
        c.execute("""
            INSERT INTO cau_hinh_he_thong (ten_cau_hinh, gia_tri, mo_ta) VALUES
            ('cv_danh_so_option', 'RIENG', 'CHUNG hoac RIENG - Cách đánh số công văn')
            ON CONFLICT (ten_cau_hinh) DO NOTHING
        """)
        
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi khởi tạo bảng công văn: {e}")
        return False

# === Hàm lấy cấu hình đánh số ===
def get_cv_danh_so_option():
    """Lấy option đánh số công văn: CHUNG hoặc RIENG"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT gia_tri FROM cau_hinh_he_thong WHERE ten_cau_hinh = 'cv_danh_so_option'")
        result = c.fetchone()
        db.close()
        return result[0] if result else 'RIENG'
    except:
        return 'RIENG'

# === Hàm cập nhật cấu hình đánh số ===
def update_cv_danh_so_option(option):
    """Cập nhật option đánh số công văn"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            UPDATE cau_hinh_he_thong 
            SET gia_tri = %s, updated_at = NOW() 
            WHERE ten_cau_hinh = 'cv_danh_so_option'
        """, (option,))
        db.commit()
        db.close()
        return True
    except:
        return False

# === Hàm lấy số max hiện tại ===
def get_so_max_cong_van(loai=None):
    """Lấy số max hiện tại cho loại công văn (hoặc chung nếu loai=None)"""
    nam_hien_tai = datetime.now().year
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        
        if loai:
            c.execute("""
                SELECT so_max FROM cau_hinh_cong_van 
                WHERE loai = %s AND nam_hien_tai = %s
            """, (loai, nam_hien_tai))
        else:
            c.execute("""
                SELECT so_max FROM cau_hinh_cong_van 
                WHERE loai = 'CHUNG' AND nam_hien_tai = %s
            """, (nam_hien_tai,))
        
        result = c.fetchone()
        db.close()
        return result[0] if result else 0
    except:
        return 0

# === Hàm cập nhật số max ===
def update_so_max_cong_van(loai, so_moi):
    """Cập nhật số max cho loại công văn"""
    nam_hien_tai = datetime.now().year
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            INSERT INTO cau_hinh_cong_van (loai, so_max, nam_hien_tai, updated_at)
            VALUES (%s, %s, %s, NOW())
            ON CONFLICT (loai, nam_hien_tai) 
            DO UPDATE SET so_max = EXCLUDED.so_max, updated_at = NOW()
        """, (loai, so_moi, nam_hien_tai))
        db.commit()
        db.close()
        return True
    except:
        return False

# === Cấu hình + đánh số Hợp đồng kinh tế (HĐKT): mẫu "stt/năm/Prefix-ma_cty" ===
# === Cấu hình hạn nộp Báo cáo Tăng/Giảm BHXH hàng tháng (mỗi doanh nghiệp 1 ngày riêng) ===
def get_han_nop_bhxh():
    """Ngày trong tháng phải nộp BC Tăng/Giảm BHXH (mặc định 20, VD CHL nộp trước ngày 20)."""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT gia_tri FROM cau_hinh_he_thong WHERE ten_cau_hinh = 'han_nop_bhxh_ngay'")
        r = c.fetchone()
        db.close()
        return int(r[0]) if r and r[0] else 20
    except:
        return 20

def update_han_nop_bhxh(ngay):
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            INSERT INTO cau_hinh_he_thong (ten_cau_hinh, gia_tri, ghi_chu)
            VALUES ('han_nop_bhxh_ngay', %s, 'Ngày trong tháng phải nộp BC Tăng/Giảm BHXH')
            ON CONFLICT (ten_cau_hinh) DO UPDATE SET gia_tri = EXCLUDED.gia_tri, updated_at = NOW()
        """, (str(ngay),))
        db.commit(); db.close()
        return True
    except:
        return False

# === Cấu hình Chấm công (khung sườn - sẽ tích hợp logic tính công dần theo nhu cầu) ===
def get_cau_hinh_cham_cong():
    from datetime import time as _time
    mac_dinh = {'gio_vao': _time(8, 0), 'gio_ra': _time(17, 0), 'phut_tre': 15}
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT ten_cau_hinh, gia_tri FROM cau_hinh_he_thong WHERE ten_cau_hinh IN "
                   "('cc_gio_vao', 'cc_gio_ra', 'cc_phut_tre')")
        rows = dict(c.fetchall())
        db.close()
        if rows.get('cc_gio_vao'):
            h, m = map(int, rows['cc_gio_vao'].split(':'))
            mac_dinh['gio_vao'] = _time(h, m)
        if rows.get('cc_gio_ra'):
            h, m = map(int, rows['cc_gio_ra'].split(':'))
            mac_dinh['gio_ra'] = _time(h, m)
        if rows.get('cc_phut_tre'):
            mac_dinh['phut_tre'] = int(rows['cc_phut_tre'])
        return mac_dinh
    except:
        return mac_dinh

def update_cau_hinh_cham_cong(gio_vao, gio_ra, phut_tre):
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        for ten, gia_tri in [('cc_gio_vao', gio_vao.strftime('%H:%M')),
                              ('cc_gio_ra', gio_ra.strftime('%H:%M')),
                              ('cc_phut_tre', str(phut_tre))]:
            c.execute("""
                INSERT INTO cau_hinh_he_thong (ten_cau_hinh, gia_tri, ghi_chu)
                VALUES (%s, %s, 'Cấu hình chấm công')
                ON CONFLICT (ten_cau_hinh) DO UPDATE SET gia_tri = EXCLUDED.gia_tri, updated_at = NOW()
            """, (ten, gia_tri))
        db.commit(); db.close()
        return True
    except:
        return False

# === Cấu hình công thức tính lương đang áp dụng (khung sườn - salary/salary_{key}.py) ===
def get_cau_hinh_luong_key():
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT gia_tri FROM cau_hinh_he_thong WHERE ten_cau_hinh = 'luong_plugin_key'")
        r = c.fetchone()
        db.close()
        return r[0] if r and r[0] else 'salary_1'
    except:
        return 'salary_1'

def update_cau_hinh_luong_key(key):
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            INSERT INTO cau_hinh_he_thong (ten_cau_hinh, gia_tri, ghi_chu)
            VALUES ('luong_plugin_key', %s, 'Công thức tính lương đang áp dụng (salary/salary_{key}.py)')
            ON CONFLICT (ten_cau_hinh) DO UPDATE SET gia_tri = EXCLUDED.gia_tri, updated_at = NOW()
        """, (key,))
        db.commit(); db.close()
        return True
    except:
        return False

def get_hdkt_prefix():
    """Lấy prefix đánh số HĐKT hiện tại (mặc định 'HĐKT'), cho phép admin tuỳ chỉnh."""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT gia_tri FROM cau_hinh_he_thong WHERE ten_cau_hinh = 'hdkt_prefix'")
        result = c.fetchone()
        db.close()
        return result[0] if result and result[0] else 'HĐKT'
    except:
        return 'HĐKT'

def update_hdkt_prefix(prefix):
    """Cập nhật prefix đánh số HĐKT (chỉ Admin)."""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            INSERT INTO cau_hinh_he_thong (ten_cau_hinh, gia_tri, ghi_chu)
            VALUES ('hdkt_prefix', %s, 'Prefix đánh số Hợp đồng kinh tế')
            ON CONFLICT (ten_cau_hinh) DO UPDATE SET gia_tri = EXCLUDED.gia_tri, updated_at = NOW()
        """, (prefix,))
        db.commit(); db.close()
        return True
    except:
        return False

def preview_so_hdkt():
    """Chỉ XEM TRƯỚC số HĐKT tiếp theo (dạng stt/năm/Prefix-ma_cty) - KHÔNG ghi/tăng số max."""
    prefix = get_hdkt_prefix()
    ma_cty = st.session_state.tenant.get('ma_cty', 'CHL') if st.session_state.get('tenant') else 'CHL'
    nam_hien_tai = datetime.now().year
    so_max = get_so_max_cong_van('HDKT')
    so_moi = so_max + 1
    return f"{so_moi:02d}/{nam_hien_tai}/{prefix}-{ma_cty}", prefix

def generate_so_hdkt():
    """Sinh số HĐKT CHÍNH THỨC + cập nhật số max. CHỈ gọi khi bấm 'Lưu Hợp đồng'."""
    prefix = get_hdkt_prefix()
    ma_cty = st.session_state.tenant.get('ma_cty', 'CHL') if st.session_state.get('tenant') else 'CHL'
    nam_hien_tai = datetime.now().year
    so_max = get_so_max_cong_van('HDKT')
    so_moi = so_max + 1
    update_so_max_cong_van('HDKT', so_moi)
    return f"{so_moi:02d}/{nam_hien_tai}/{prefix}-{ma_cty}"

# === Ánh xạ mã loại công văn (bảng danh_muc_loai_cong_van) <-> mã loại nội bộ dùng cho cấu hình/đánh số ===
# Lưu ý: cột ma_loai trong danh_muc_loai_cong_van lưu chính là ký hiệu hiển thị (QĐ, CV, BC, TB, TTr),
# trong khi bảng cau_hinh_cong_van và bộ lọc loại công văn dùng mã nội bộ (QUYET_DINH, CONG_VAN, ...).
# Trước đây 2 bộ mã này KHÔNG khớp nhau khiến prefix luôn rơi về mặc định 'CV'.
MA_LOAI_TO_CODE = {
    'QĐ': 'QUYET_DINH',
    'CV': 'CONG_VAN',
    'BC': 'BAO_CAO',
    'TB': 'THONG_BAO',
    'TTr': 'TO_TRINH',
}
PREFIX_MAP = {
    'QUYET_DINH': 'QĐ',
    'CONG_VAN': 'CV',
    'BAO_CAO': 'BC',
    'THONG_BAO': 'TB',
    'TO_TRINH': 'TTr',
}

def chuan_hoa_loai_cong_van(ma_loai_raw):
    """Chuyển mã loại lấy từ dropdown (vd 'QĐ') về mã nội bộ chuẩn (vd 'QUYET_DINH').
    Nếu đã là mã chuẩn (hoặc không nhận diện được) thì trả về nguyên giá trị."""
    return MA_LOAI_TO_CODE.get(ma_loai_raw, ma_loai_raw)

# === Hàm XEM TRƯỚC số công văn (KHÔNG ghi/cập nhật cấu hình) ===
def preview_so_cong_van(loai_cv):
    """Chỉ tính toán số công văn TIẾP THEO sẽ được cấp để hiển thị cho user xem trước.
    Hàm này KHÔNG được phép ghi gì vào DB (không tăng số max), vì nó bị gọi lại
    ở MỌI lần rerun của trang (mỗi khi user click bất kỳ đâu trên màn hình)."""
    option = get_cv_danh_so_option()
    ma_cty = st.session_state.tenant.get('ma_cty', 'CHL') if st.session_state.get('tenant') else 'CHL'
    nam_hien_tai = datetime.now().year

    prefix = PREFIX_MAP.get(loai_cv, 'CV')

    loai_tim = 'CHUNG' if option == 'CHUNG' else loai_cv
    so_max = get_so_max_cong_van(loai_tim)
    so_moi = so_max + 1

    so_cv = f"{so_moi:02d}/{nam_hien_tai}/{prefix}-{ma_cty}"
    return so_cv

# === Hàm SINH SỐ CHÍNH THỨC + cập nhật số max (chỉ gọi khi user bấm "Lưu công văn đi") ===
def generate_so_cong_van(loai_cv):
    """Sinh số công văn chính thức theo cấu hình VÀ cập nhật số max trong DB.
    CHỈ được gọi trong handler của nút 'Lưu công văn đi', không được gọi khi
    chỉ render lại form (tránh nhảy số mỗi lần rerun)."""
    option = get_cv_danh_so_option()
    ma_cty = st.session_state.tenant.get('ma_cty', 'CHL') if st.session_state.get('tenant') else 'CHL'
    nam_hien_tai = datetime.now().year

    prefix = PREFIX_MAP.get(loai_cv, 'CV')

    # Xác định loại để lấy số max
    loai_tim = 'CHUNG' if option == 'CHUNG' else loai_cv
    so_max = get_so_max_cong_van(loai_tim)
    so_moi = so_max + 1

    # Cập nhật số max - CHỈ xảy ra khi hàm này được gọi (tức là khi Lưu)
    update_so_max_cong_van(loai_tim, so_moi)

    # Tạo số công văn
    so_cv = f"{so_moi:02d}/{nam_hien_tai}/{prefix}-{ma_cty}"
    return so_cv

# === Hàm upload file cho công văn ===
def upload_cong_van_file(uploaded_file, folder_name):
    """Upload file công văn lên Supabase Storage"""
    if not uploaded_file:
        return None
    
    sb = get_supabase_storage()
    if not sb:
        return None
    
    try:
        safe_name = sanitize_storage_filename(uploaded_file.name)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = f"cong_van/{folder_name}/{timestamp}_{safe_name}"
        
        # Upload file
        result = sb.storage.from_(SUPABASE_BUCKET).upload(
            path=path,
            file=uploaded_file.getvalue(),
            file_options={"content-type": uploaded_file.type or "application/octet-stream"}
        )
        return path
    except Exception as e:
        print(f"Lỗi upload file: {e}")
        return None

# === Hàm lấy danh sách công văn đến ===
def get_cong_van_den(tu_ngay=None, den_ngay=None, search_text=None):
    """Lấy danh sách công văn đến với bộ lọc"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        sql = "SELECT * FROM cong_van_den WHERE 1=1"
        params = []
        
        if tu_ngay:
            sql += " AND ngay_den >= %s"
            params.append(tu_ngay)
        if den_ngay:
            sql += " AND ngay_den <= %s"
            params.append(den_ngay)
        if search_text:
            sql += """ AND (so_cong_van ILIKE %s OR tieu_de ILIKE %s 
                     OR co_quan_phat_hanh ILIKE %s OR trich_yeu ILIKE %s OR ma_vach_buu_dien ILIKE %s)"""
            search_pattern = f"%{search_text}%"
            params.extend([search_pattern] * 5)
        
        sql += " ORDER BY ngay_den DESC, id DESC"
        c.execute(sql, tuple(params))
        result = c.fetchall()
        db.close()
        return result
    except Exception as e:
        print(f"Lỗi lấy công văn đến: {e}")
        return []

# === Hàm lấy danh sách công văn đi ===
def get_cong_van_di(tu_ngay=None, den_ngay=None, search_text=None, loai_cv=None):
    """Lấy danh sách công văn đi với bộ lọc"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        sql = "SELECT * FROM cong_van_di WHERE 1=1"
        params = []
        
        if tu_ngay:
            sql += " AND ngay_phat_hanh >= %s"
            params.append(tu_ngay)
        if den_ngay:
            sql += " AND ngay_phat_hanh <= %s"
            params.append(den_ngay)
        if loai_cv:
            sql += " AND loai_cong_van = %s"
            params.append(loai_cv)
        if search_text:
            sql += """ AND (so_cong_van ILIKE %s OR tieu_de ILIKE %s 
                     OR phong_phat_hanh ILIKE %s OR trich_yeu ILIKE %s OR ma_vach_buu_dien ILIKE %s)"""
            search_pattern = f"%{search_text}%"
            params.extend([search_pattern] * 5)
        
        sql += " ORDER BY ngay_phat_hanh DESC, id DESC"
        c.execute(sql, tuple(params))
        result = c.fetchall()
        db.close()
        return result
    except Exception as e:
        print(f"Lỗi lấy công văn đi: {e}")
        return []

# === Hàm lấy danh sách hợp đồng kinh tế ===
def get_hop_dong_kinh_te(tu_ngay=None, den_ngay=None, search_text=None):
    """Lấy danh sách hợp đồng kinh tế với bộ lọc"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        sql = "SELECT * FROM hop_dong_kinh_te WHERE 1=1"
        params = []
        
        if tu_ngay:
            sql += " AND ngay_ky >= %s"
            params.append(tu_ngay)
        if den_ngay:
            sql += " AND ngay_ky <= %s"
            params.append(den_ngay)
        if search_text:
            sql += """ AND (so_hop_dong ILIKE %s OR ten_doi_tac ILIKE %s 
                     OR trich_yeu ILIKE %s)"""
            search_pattern = f"%{search_text}%"
            params.extend([search_pattern] * 3)
        
        sql += " ORDER BY ngay_ky DESC, id DESC"
        c.execute(sql, tuple(params))
        result = c.fetchall()
        db.close()
        return result
    except Exception as e:
        print(f"Lỗi lấy hợp đồng kinh tế: {e}")
        return []

# === Hàm xuất Excel công văn ===
def export_cong_van_excel(data, ten_file, headers, col_widths=None, title=None):
    """Xuất dữ liệu ra file Excel với format chuẩn"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Danh sách"
    
    # Border
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Header fill
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    
    # Tiêu đề
    if title:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        ws['A1'] = title
        ws['A1'].font = Font(bold=True, size=14, name='Times New Roman')
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Dòng trống
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
        ws['A2'] = f"Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws['A2'].font = Font(size=10, name='Times New Roman', italic=True)
        ws['A2'].alignment = Alignment(horizontal='center')
        start_row = 4
    else:
        start_row = 1
    
    # Header
    header_row = start_row
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = Font(bold=True, size=11, name='Times New Roman', color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    
    # Dữ liệu
    data_row = header_row + 1
    for idx, row in enumerate(data):
        for col_idx, key in enumerate(headers, 1):
            value = row.get(key, '')
            # Format ngày tháng
            if 'ngay' in key.lower() and value:
                if hasattr(value, 'strftime'):
                    value = value.strftime('%d/%m/%Y')
            cell = ws.cell(row=data_row + idx, column=col_idx, value=value)
            cell.font = Font(size=10, name='Times New Roman')
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center' if col_idx <= 3 else 'left', vertical='center')
    
    # Footer
    footer_row = data_row + len(data) + 2
    ws.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=len(headers))
    ws.cell(row=footer_row, column=1, value=f"Tổng cộng: {len(data)} bản ghi")
    ws.cell(row=footer_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
    
    # Độ rộng cột
    if col_widths:
        for idx, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
    
    # Lưu file
    wb.save(ten_file)
    return ten_file

# === UI: Quản lý Công văn & HĐ kinh tế ===
def show_quan_ly_cong_van():
    """Hiển thị giao diện Quản lý Công văn & HĐ kinh tế"""
    st.title("📄 Quản lý Công văn & HĐ kinh tế")
    
    # Khởi tạo bảng nếu chưa có
    init_cong_van_tables()
    
    # Kiểm tra quyền
    role = st.session_state.get('role', '')
    if role not in ['admin', 'van_thu']:
        st.warning("🔒 Chỉ Admin và Văn thư mới có quyền truy cập module này!")
        st.stop()
    
    # Cấu hình (chỉ Admin)
    if role == 'admin':
        with st.expander("⚙️ Cấu hình đánh số công văn", expanded=False):
            st.markdown("**Cấu hình cách đánh số công văn đi**")
            
            current_option = get_cv_danh_so_option()
            new_option = st.radio(
                "Chọn phương án đánh số:",
                options=['CHUNG', 'RIENG'],
                index=0 if current_option == 'CHUNG' else 1,
                format_func=lambda x: "📌 Số chung cho tất cả loại công văn" if x == 'CHUNG' else "📌 Mỗi loại công văn có số riêng",
                key="cv_option_radio"
            )
            
            if new_option != current_option:
                if st.button("✅ Cập nhật cấu hình", type="primary"):
                    if update_cv_danh_so_option(new_option):
                        st.success(f"✅ Đã cập nhật cấu hình sang: {new_option}")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("❌ Cập nhật thất bại!")
            
            st.divider()
            st.markdown("**📊 Trạng thái đánh số hiện tại**")
            
            # Lấy dữ liệu từ database
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            
            # Xác định loại cần hiển thị dựa trên option
            option = get_cv_danh_so_option()
            if option == 'CHUNG':
                # Chỉ hiển thị CHUNG
                c.execute("""
                    SELECT loai, so_max, prefix, nam_hien_tai, updated_at 
                    FROM cau_hinh_cong_van 
                    WHERE loai = 'CHUNG'
                    ORDER BY loai, nam_hien_tai
                """)
            else:
                # Hiển thị các loại riêng (không bao gồm CHUNG và QĐ trùng)
                c.execute("""
                    SELECT loai, so_max, prefix, nam_hien_tai, updated_at 
                    FROM cau_hinh_cong_van 
                    WHERE loai IN ('CONG_VAN', 'QUYET_DINH', 'BAO_CAO', 'THONG_BAO', 'TO_TRINH')
                    ORDER BY 
                        CASE loai
                            WHEN 'QUYET_DINH' THEN 1
                            WHEN 'CONG_VAN' THEN 2
                            WHEN 'BAO_CAO' THEN 3
                            WHEN 'THONG_BAO' THEN 4
                            WHEN 'TO_TRINH' THEN 5
                        END
                """)
            
            configs = c.fetchall()
            db.close()
            
            if configs:
                df_config = pd.DataFrame(configs)
                df_config['updated_at'] = df_config['updated_at'].apply(
                    lambda x: x.strftime('%d/%m/%Y %H:%M') if x else ''
                )
                
                # Đổi tên loại cho đẹp
                loai_name_map = {
                    'CHUNG': 'CHUNG (Tất cả loại)',
                    'QUYET_DINH': 'QUYẾT ĐỊNH',
                    'CONG_VAN': 'CÔNG VĂN',
                    'BAO_CAO': 'BÁO CÁO',
                    'THONG_BAO': 'THÔNG BÁO',
                    'TO_TRINH': 'TỜ TRÌNH'
                }
                df_config['loai'] = df_config['loai'].map(loai_name_map)
                df_config.columns = ['Loại', 'Số hiện tại', 'Prefix', 'Năm', 'Cập nhật lúc']
                st.dataframe(df_config, width='stretch', hide_index=True)
            else:
                st.info("Chưa có dữ liệu cấu hình. Hãy tạo công văn đi đầu tiên để khởi tạo.")
            
            st.divider()
            st.markdown("**🔄 Đặt lại số**")
            col_reset1, col_reset2 = st.columns(2)
            with col_reset1:
                # Xác định danh sách loại cho dropdown dựa trên option
                if option == 'CHUNG':
                    loai_list = ['CHUNG']
                    loai_display = {'CHUNG': 'CHUNG (Tất cả loại)'}
                else:
                    loai_list = ['QUYET_DINH', 'CONG_VAN', 'BAO_CAO', 'THONG_BAO', 'TO_TRINH']
                    loai_display = {
                        'QUYET_DINH': 'QUYẾT ĐỊNH',
                        'CONG_VAN': 'CÔNG VĂN',
                        'BAO_CAO': 'BÁO CÁO',
                        'THONG_BAO': 'THÔNG BÁO',
                        'TO_TRINH': 'TỜ TRÌNH'
                    }
                
                selected_loai_display = st.selectbox(
                    "Chọn loại cần đặt lại:",
                    [loai_display.get(l, l) for l in loai_list],
                    key="reset_loai_display"
                )
                
                # Lấy lại mã loại thực tế
                if option == 'CHUNG':
                    loai_reset = 'CHUNG'
                else:
                    # Tìm key từ display name
                    for key, value in loai_display.items():
                        if value == selected_loai_display:
                            loai_reset = key
                            break
                
                # Lấy số hiện tại để hiển thị
                current_so = get_so_max_cong_van(loai_reset)
                st.caption(f"📌 Số hiện tại: **{current_so}**")
                
                so_moi = st.number_input(
                    "Số bắt đầu mới:", 
                    min_value=0, 
                    value=current_so, 
                    step=1, 
                    key="reset_so"
                )
            
            with col_reset2:
                # Nguyên tắc: cấu hình (so_max) CHỈ được cập nhật khi user bấm nút "Đặt lại số",
                # và luôn lấy đúng giá trị đang có trong ô "Số bắt đầu mới:" tại thời điểm bấm.
                # (Trước đây dùng checkbox xác nhận lồng bên trong if st.button(...) — nhưng vì
                # Streamlit rerun lại toàn bộ script sau mỗi tương tác, việc tick checkbox ở lần
                # rerun sau đó lại rơi vào nhánh st.button() == False nên state bị "treo" và
                # không đáng tin cậy. Dùng session_state để giữ yêu cầu đặt lại qua các lần rerun.)
                st.markdown("&nbsp;")  # căn cho nút thẳng hàng với ô nhập bên trái
                if st.button("🔄 Đặt lại số", type="secondary", key="btn_dat_lai_so"):
                    st.session_state['cv_pending_reset'] = {'loai': loai_reset, 'so_moi': so_moi}

                pending = st.session_state.get('cv_pending_reset')
                if pending and pending['loai'] == loai_reset:
                    st.warning(
                        f"⚠️ Xác nhận đặt lại số cho **{selected_loai_display}** "
                        f"từ **{current_so}** thành **{pending['so_moi']}**?"
                    )
                    col_confirm, col_cancel = st.columns(2)
                    with col_confirm:
                        if st.button("✅ Xác nhận đặt lại số", type="primary", key="confirm_dat_lai_so"):
                            if update_so_max_cong_van(pending['loai'], pending['so_moi']):
                                st.success(
                                    f"✅ Đã đặt lại số cho {selected_loai_display} "
                                    f"từ {current_so} thành {pending['so_moi']}"
                                )
                                del st.session_state['cv_pending_reset']
                                st.cache_data.clear()
                                st.rerun()
                            else:
                                st.error("❌ Đặt lại số thất bại!")
                    with col_cancel:
                        if st.button("✖️ Hủy", key="cancel_dat_lai_so"):
                            del st.session_state['cv_pending_reset']
                            st.rerun()

        with st.expander("⚙️ Cấu hình đánh số Hợp đồng kinh tế (HĐKT)", expanded=False):
            st.caption("Số HĐKT tự sinh theo mẫu: **stt/năm/Prefix-ma_cty** (VD: 04/2026/HĐKT-CHL)")
            prefix_hdkt_hien_tai = get_hdkt_prefix()
            prefix_hdkt_moi = st.text_input(
                "Prefix đánh số HĐKT:", value=prefix_hdkt_hien_tai, key="hdkt_prefix_input"
            )
            if st.button("✅ Cập nhật prefix HĐKT", key="btn_update_hdkt_prefix"):
                if prefix_hdkt_moi.strip() and update_hdkt_prefix(prefix_hdkt_moi.strip()):
                    st.success(f"✅ Đã cập nhật prefix HĐKT sang: {prefix_hdkt_moi.strip()}")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("❌ Cập nhật thất bại!")
    
    # Tabs
    tab1, tab2, tab3 = st.tabs(["📥 Công văn đến", "📤 Công văn đi", "📑 Hợp đồng kinh tế"])
    
    # === TAB 1: CÔNG VĂN ĐẾN ===
    with tab1:
        st.subheader("📥 Quản lý Công văn đến")
        
        # Form thêm mới
        with st.expander("➕ Thêm công văn đến mới", expanded=False):
            with st.form("add_cong_van_den"):
                col1, col2 = st.columns(2)
                with col1:
                    so_cv = st.text_input("Số công văn *", placeholder="VD: 123/BQP-2026")
                    co_quan = st.text_input("Cơ quan phát hành *", placeholder="VD: Bộ Quốc phòng")
                    ngay_den = st.date_input("Ngày đến *", value=date.today())
                    ma_vach = st.text_input("📦 Mã vạch Bưu điện", placeholder="VD: EV123456789VN")
                with col2:
                    tieu_de = st.text_input("Tiêu đề *", placeholder="Nhập tiêu đề công văn...")
                    trich_yeu = st.text_area("Trích yếu", placeholder="Tóm tắt nội dung chính...", height=80)
                    ghi_chu = st.text_area("Ghi chú", height=60)
                
                uploaded_file = st.file_uploader("📎 Upload file", type=['pdf', 'doc', 'docx', 'jpg', 'png', 'jpeg'], key="cv_den_upload")
                
                col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
                with col_btn2:
                    if st.form_submit_button("💾 Lưu công văn đến", width='stretch', type="primary"):
                        if not so_cv or not co_quan or not tieu_de:
                            st.error("⚠️ Vui lòng nhập đầy đủ các trường bắt buộc (*)")
                        else:
                            try:
                                # Upload file
                                file_url = None
                                if uploaded_file:
                                    file_url = upload_cong_van_file(uploaded_file, "den")
                                
                                # Lưu vào database
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                c.execute("""
                                    INSERT INTO cong_van_den (so_cong_van, co_quan_phat_hanh, ngay_den, 
                                    tieu_de, trich_yeu, file_url, ghi_chu, nguoi_tao, ma_vach_buu_dien)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                                """, (so_cv, co_quan, ngay_den, tieu_de, trich_yeu, file_url, ghi_chu, 
                                      st.session_state.username, ma_vach))
                                db.commit()
                                db.close()
                                
                                st.success(f"✅ Đã thêm công văn đến: {so_cv}")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")
        
        # Tìm kiếm và lọc
        st.divider()
        col_search1, col_search2, col_search3 = st.columns([2, 1, 1])
        with col_search1:
            search_text_cv_den = st.text_input("🔍 Tìm kiếm", placeholder="Theo số, tiêu đề, cơ quan...", key="search_cv_den")
        with col_search2:
            tu_ngay_cv_den = st.date_input("Từ ngày", value=None, key="tu_ngay_cv_den")
        with col_search3:
            den_ngay_cv_den = st.date_input("Đến ngày", value=None, key="den_ngay_cv_den")
        
        # Lấy dữ liệu
        data_cv_den = get_cong_van_den(tu_ngay_cv_den, den_ngay_cv_den, search_text_cv_den)
        
        # Hiển thị bảng
        if data_cv_den:
            df_cv_den = pd.DataFrame(data_cv_den)
            
            # Format ngày
            for col in ['ngay_den', 'created_at', 'updated_at']:
                if col in df_cv_den.columns:
                    df_cv_den[col] = df_cv_den[col].apply(format_date)
            
            display_cols = ['so_cong_van', 'co_quan_phat_hanh', 'ngay_den', 'tieu_de', 'trich_yeu', 'ma_vach_buu_dien', 'file_url', 'ghi_chu']
            available_cols = [c for c in display_cols if c in df_cv_den.columns]
            df_display = df_cv_den[available_cols]
            
            col_map = {
                'so_cong_van': 'Số công văn',
                'co_quan_phat_hanh': 'Cơ quan phát hành',
                'ngay_den': 'Ngày đến',
                'tieu_de': 'Tiêu đề',
                'trich_yeu': 'Trích yếu',
                'ma_vach_buu_dien': 'Mã vạch BĐ',
                'file_url': 'File',
                'ghi_chu': 'Ghi chú'
            }
            df_display.rename(columns=col_map, inplace=True)
            
            st.caption(f"📌 Tổng số: {len(data_cv_den)} công văn đến")
            st.dataframe(df_display, width='stretch', hide_index=True, height=400)
            
            # Nút xuất Excel
            col_export1, col_export2, col_export3 = st.columns([1, 2, 1])
            with col_export2:
                if st.button("📥 Xuất Excel công văn đến", width='stretch', type="primary"):
                    headers = ['so_cong_van', 'co_quan_phat_hanh', 'ngay_den', 'tieu_de', 'trich_yeu', 'ghi_chu']
                    col_widths = [15, 25, 12, 35, 30, 25]
                    title = f"BÁO CÁO CÔNG VĂN ĐẾN (Từ {tu_ngay_cv_den or '...'} đến {den_ngay_cv_den or '...'})"
                    filename = f"Cong_van_den_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    # Chuyển đổi data cho Excel
                    excel_data = []
                    for row in data_cv_den:
                        excel_row = {}
                        for key in headers:
                            val = row.get(key)
                            if key == 'ngay_den' and val:
                                val = format_date(val)
                            excel_row[key] = val
                        excel_data.append(excel_row)
                    
                    export_cong_van_excel(excel_data, filename, headers, col_widths, title)
                    
                    with open(filename, "rb") as f:
                        st.download_button(
                            label="📥 TẢI FILE EXCEL",
                            data=f,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch'
                        )
                    st.success(f"✅ Đã xuất {len(data_cv_den)} công văn đến")
        else:
            st.info("📭 Không có công văn đến nào")
    
    # === TAB 2: CÔNG VĂN ĐI ===
    with tab2:
        st.subheader("📤 Quản lý Công văn đi")
        
        # Form thêm mới
        with st.expander("➕ Thêm công văn đi mới", expanded=False):
            # Lấy danh sách loại công văn
            db_loai = st.session_state.db_engine.get_connection()
            c_loai = db_loai.cursor()
            c_loai.execute("SELECT ma_loai, ten_loai FROM danh_muc_loai_cong_van WHERE trang_thai = TRUE ORDER BY thu_tu")
            loai_cv_list = c_loai.fetchall()
            db_loai.close()

            loai_options = {f"{loai[1]} ({loai[0]})": loai[0] for loai in loai_cv_list}

            # QUAN TRỌNG: đặt selectbox "Loại công văn" NGOÀI st.form. Bên trong st.form,
            # đổi giá trị widget không làm rerun app (chỉ form_submit_button mới rerun),
            # nên trước đây đổi dropdown KHÔNG cập nhật được ký hiệu/prefix hiển thị.
            # Đặt ngoài form giúp phần xem trước số/ký hiệu phản ứng ngay khi đổi lựa chọn.
            selected_loai = st.selectbox("Loại công văn *", list(loai_options.keys()), key="cv_di_loai")
            loai_cv = chuan_hoa_loai_cong_van(loai_options[selected_loai])

            # CHỈ xem trước số/ký hiệu - hàm này không ghi gì vào cấu hình/DB.
            so_cv_xem_truoc = preview_so_cong_van(loai_cv)
            prefix_hien_tai = PREFIX_MAP.get(loai_cv, 'CV')
            st.info(
                f"📄 **Số công văn dự kiến:** `{so_cv_xem_truoc}` (Prefix: **{prefix_hien_tai}**) "
                f"— số chính thức sẽ được cấp khi bấm **Lưu công văn đi**"
            )

            with st.form("add_cong_van_di"):
                col1, col2 = st.columns(2)
                with col1:
                    phong_phat_hanh = st.text_input("Phòng phát hành *", placeholder="VD: Phòng Hành chính")
                    ngay_phat_hanh = st.date_input("Ngày phát hành *", value=date.today())
                    ma_vach = st.text_input("📦 Mã vạch Bưu điện", placeholder="VD: EV123456789VN")
                with col2:
                    tieu_de = st.text_input("Tiêu đề *", placeholder="Nhập tiêu đề công văn...")
                    trich_yeu = st.text_area("Trích yếu", placeholder="Tóm tắt nội dung chính...", height=80)
                    ghi_chu = st.text_area("Ghi chú", height=60)
                
                uploaded_file = st.file_uploader("📎 Upload file", type=['pdf', 'doc', 'docx', 'jpg', 'png', 'jpeg'], key="cv_di_upload")
                
                col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
                with col_btn2:
                    if st.form_submit_button("💾 Lưu công văn đi", width='stretch', type="primary"):
                        if not phong_phat_hanh or not tieu_de:
                            st.error("⚠️ Vui lòng nhập đầy đủ các trường bắt buộc (*)")
                        else:
                            try:
                                # Upload file
                                file_url = None
                                if uploaded_file:
                                    file_url = upload_cong_van_file(uploaded_file, "di")
                                
                                # Nguyên tắc: số công văn CHỈ được sinh chính thức (và cấu hình
                                # so_max CHỈ được cập nhật) tại đây, khi user bấm "Lưu công văn đi".
                                so_cv_chinh_thuc = generate_so_cong_van(loai_cv)

                                # Lưu vào database
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                c.execute("""
                                    INSERT INTO cong_van_di (so_cong_van, phong_phat_hanh, ngay_phat_hanh, 
                                    tieu_de, trich_yeu, file_url, loai_cong_van, ghi_chu, nguoi_tao, ma_vach_buu_dien)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                """, (so_cv_chinh_thuc, phong_phat_hanh, ngay_phat_hanh, tieu_de, trich_yeu, 
                                      file_url, loai_cv, ghi_chu, st.session_state.username, ma_vach))
                                db.commit()
                                db.close()
                                
                                st.success(f"✅ Đã thêm công văn đi: {so_cv_chinh_thuc}")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")
        
        # Tìm kiếm và lọc
        st.divider()
        col_search1, col_search2 = st.columns([2, 1])
        with col_search1:
            search_text_cv_di = st.text_input("🔍 Tìm kiếm", placeholder="Theo số, tiêu đề, phòng...", key="search_cv_di")
        with col_search2:
            loai_filter = st.selectbox("Loại", ["Tất cả", "Quyết định", "Công văn", "Báo cáo", "Thông báo", "Tờ trình"], key="loai_filter_cv_di")
        col_search3, col_search4 = st.columns(2)
        with col_search3:
            tu_ngay_cv_di = st.date_input("Từ ngày", value=None, key="tu_ngay_cv_di")
        with col_search4:
            den_ngay_cv_di = st.date_input("Đến ngày", value=None, key="den_ngay_cv_di")
        
        # Lấy dữ liệu
        loai_map = {
            "Tất cả": None,
            "Quyết định": "QUYET_DINH",
            "Công văn": "CONG_VAN",
            "Báo cáo": "BAO_CAO",
            "Thông báo": "THONG_BAO",
            "Tờ trình": "TO_TRINH"
        }
        data_cv_di = get_cong_van_di(tu_ngay_cv_di, den_ngay_cv_di, search_text_cv_di, loai_map.get(loai_filter))
        
        # Hiển thị bảng
        if data_cv_di:
            df_cv_di = pd.DataFrame(data_cv_di)
            
            # Format ngày
            for col in ['ngay_phat_hanh', 'created_at', 'updated_at']:
                if col in df_cv_di.columns:
                    df_cv_di[col] = df_cv_di[col].apply(format_date)
            
            display_cols = ['so_cong_van', 'loai_cong_van', 'phong_phat_hanh', 'ngay_phat_hanh', 'tieu_de', 'trich_yeu', 'ma_vach_buu_dien', 'file_url', 'ghi_chu']
            available_cols = [c for c in display_cols if c in df_cv_di.columns]
            df_display = df_cv_di[available_cols]
            
            # Map loại
            loai_name_map = {
                'QUYET_DINH': 'Quyết định',
                'CONG_VAN': 'Công văn',
                'BAO_CAO': 'Báo cáo',
                'THONG_BAO': 'Thông báo',
                'TO_TRINH': 'Tờ trình'
            }
            df_display['loai_cong_van'] = df_display['loai_cong_van'].map(loai_name_map)
            
            col_map = {
                'so_cong_van': 'Số công văn',
                'loai_cong_van': 'Loại',
                'phong_phat_hanh': 'Phòng phát hành',
                'ngay_phat_hanh': 'Ngày phát hành',
                'tieu_de': 'Tiêu đề',
                'trich_yeu': 'Trích yếu',
                'ma_vach_buu_dien': 'Mã vạch BĐ',
                'file_url': 'File',
                'ghi_chu': 'Ghi chú'
            }
            df_display.rename(columns=col_map, inplace=True)
            
            st.caption(f"📌 Tổng số: {len(data_cv_di)} công văn đi")
            st.dataframe(df_display, width='stretch', hide_index=True, height=400)

            # ----- Sửa / Xóa công văn đi -----
            with st.expander("✏️ Sửa / 🗑️ Xóa công văn đi", expanded=False):
                tuy_chon_cvd = {f"{r['so_cong_van']} - {r.get('tieu_de') or ''}": r for r in data_cv_di}
                chon_cvd = st.selectbox("Chọn công văn đi:", ["-- Chọn --"] + list(tuy_chon_cvd.keys()), key="chon_sua_cvd")
                if chon_cvd != "-- Chọn --":
                    bg_sua = tuy_chon_cvd[chon_cvd]
                    col_s1, col_s2 = st.columns(2)
                    with col_s1:
                        tieu_de_sua_cvd = st.text_input("Tiêu đề:", value=bg_sua.get('tieu_de') or '', key=f"sua_td_cvd_{bg_sua['id']}")
                        ma_vach_sua_cvd = st.text_input("Mã vạch Bưu điện:", value=bg_sua.get('ma_vach_buu_dien') or '', key=f"sua_mv_cvd_{bg_sua['id']}")
                    with col_s2:
                        trich_yeu_sua_cvd = st.text_area("Trích yếu:", value=bg_sua.get('trich_yeu') or '', key=f"sua_ty_cvd_{bg_sua['id']}", height=80)
                    col_luu_cvd, col_xoa_cvd = st.columns(2)
                    with col_luu_cvd:
                        if st.button("💾 Lưu thay đổi", key=f"btn_luu_cvd_{bg_sua['id']}", type="primary", width='stretch'):
                            try:
                                db_s = st.session_state.db_engine.get_connection()
                                c_s = db_s.cursor()
                                c_s.execute("""
                                    UPDATE cong_van_di SET tieu_de=%s, trich_yeu=%s, ma_vach_buu_dien=%s
                                    WHERE id=%s
                                """, (tieu_de_sua_cvd, trich_yeu_sua_cvd, ma_vach_sua_cvd, bg_sua['id']))
                                db_s.commit(); db_s.close()
                                st.success("✅ Đã cập nhật công văn đi")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")
                    with col_xoa_cvd:
                        if st.button("🗑️ Xóa công văn này", key=f"btn_xoa_cvd_{bg_sua['id']}", width='stretch'):
                            try:
                                db_x = st.session_state.db_engine.get_connection()
                                c_x = db_x.cursor()
                                c_x.execute("DELETE FROM cong_van_di WHERE id=%s", (bg_sua['id'],))
                                db_x.commit(); db_x.close()
                                st.success("✅ Đã xóa công văn đi")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")

            # Nút xuất Excel
            col_export1, col_export2, col_export3 = st.columns([1, 2, 1])
            with col_export2:
                if st.button("📥 Xuất Excel công văn đi", width='stretch', type="primary"):
                    headers = ['so_cong_van', 'loai_cong_van', 'phong_phat_hanh', 'ngay_phat_hanh', 'tieu_de', 'trich_yeu', 'ghi_chu']
                    col_widths = [20, 12, 20, 12, 35, 30, 25]
                    title = f"BÁO CÁO CÔNG VĂN ĐI (Từ {tu_ngay_cv_di or '...'} đến {den_ngay_cv_di or '...'})"
                    filename = f"Cong_van_di_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    excel_data = []
                    for row in data_cv_di:
                        excel_row = {}
                        for key in headers:
                            val = row.get(key)
                            if key == 'ngay_phat_hanh' and val:
                                val = format_date(val)
                            if key == 'loai_cong_van':
                                val = loai_name_map.get(val, val)
                            excel_row[key] = val
                        excel_data.append(excel_row)
                    
                    export_cong_van_excel(excel_data, filename, headers, col_widths, title)
                    
                    with open(filename, "rb") as f:
                        st.download_button(
                            label="📥 TẢI FILE EXCEL",
                            data=f,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch'
                        )
                    st.success(f"✅ Đã xuất {len(data_cv_di)} công văn đi")
        else:
            st.info("📭 Không có công văn đi nào")
    
    # === TAB 3: HỢP ĐỒNG KINH TẾ ===
    with tab3:
        st.subheader("📑 Quản lý Hợp đồng kinh tế")
        
        # Form thêm mới
        with st.expander("➕ Thêm hợp đồng kinh tế mới", expanded=False):
            so_hd_xem_truoc, prefix_hdkt_dang_dung = preview_so_hdkt()
            st.info(
                f"📄 **Số HĐKT dự kiến:** `{so_hd_xem_truoc}` (Prefix: **{prefix_hdkt_dang_dung}**) "
                f"— số chính thức sẽ được cấp khi bấm **Lưu Hợp đồng**"
            )
            with st.form("add_hop_dong_kt"):
                col1, col2 = st.columns(2)
                with col1:
                    so_hd_tuy_chinh = st.text_input(
                        "Số hợp đồng (để trống sẽ tự cấp số theo cấu hình trên)",
                        placeholder=so_hd_xem_truoc
                    )
                    ten_doi_tac = st.text_input("Tên đối tác *", placeholder="VD: Công ty TNHH ABC")
                    ngay_ky = st.date_input("Ngày ký *", value=date.today())
                with col2:
                    trich_yeu = st.text_area("Trích yếu", placeholder="Tóm tắt nội dung hợp đồng...", height=80)
                    ghi_chu = st.text_area("Ghi chú", height=60)
                
                uploaded_file = st.file_uploader("📎 Upload file", type=['pdf', 'doc', 'docx', 'jpg', 'png', 'jpeg'], key="hd_kt_upload")
                
                col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
                with col_btn2:
                    if st.form_submit_button("💾 Lưu hợp đồng", width='stretch', type="primary"):
                        if not ten_doi_tac:
                            st.error("⚠️ Vui lòng nhập đầy đủ các trường bắt buộc (*)")
                        else:
                            try:
                                # Số hợp đồng: nếu admin nhập tay thì dùng số đó (không tăng số max),
                                # nếu để trống thì cấp số CHÍNH THỨC theo cấu hình (và tăng số max)
                                so_hd = so_hd_tuy_chinh.strip() if so_hd_tuy_chinh.strip() else generate_so_hdkt()

                                # Upload file
                                file_url = None
                                if uploaded_file:
                                    file_url = upload_cong_van_file(uploaded_file, "hop_dong_kt")
                                
                                # Lưu vào database
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                c.execute("""
                                    INSERT INTO hop_dong_kinh_te (so_hop_dong, ten_doi_tac, ngay_ky, 
                                    trich_yeu, file_url, ghi_chu, nguoi_tao)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                                """, (so_hd, ten_doi_tac, ngay_ky, trich_yeu, file_url, ghi_chu, 
                                      st.session_state.username))
                                db.commit()
                                db.close()
                                
                                st.success(f"✅ Đã thêm hợp đồng: {so_hd}")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")
        
        # Tìm kiếm và lọc
        st.divider()
        col_search1, col_search2, col_search3 = st.columns([2, 1, 1])
        with col_search1:
            search_text_hd = st.text_input("🔍 Tìm kiếm", placeholder="Theo số HĐ, đối tác...", key="search_hd_kt")
        with col_search2:
            tu_ngay_hd = st.date_input("Từ ngày", value=None, key="tu_ngay_hd")
        with col_search3:
            den_ngay_hd = st.date_input("Đến ngày", value=None, key="den_ngay_hd")
        
        # Lấy dữ liệu
        data_hd = get_hop_dong_kinh_te(tu_ngay_hd, den_ngay_hd, search_text_hd)
        
        # Hiển thị bảng
        if data_hd:
            df_hd = pd.DataFrame(data_hd)
            
            # Format ngày
            for col in ['ngay_ky', 'created_at', 'updated_at']:
                if col in df_hd.columns:
                    df_hd[col] = df_hd[col].apply(format_date)
            
            display_cols = ['so_hop_dong', 'ten_doi_tac', 'ngay_ky', 'trich_yeu', 'file_url', 'ghi_chu']
            available_cols = [c for c in display_cols if c in df_hd.columns]
            df_display = df_hd[available_cols]
            
            col_map = {
                'so_hop_dong': 'Số hợp đồng',
                'ten_doi_tac': 'Đối tác',
                'ngay_ky': 'Ngày ký',
                'trich_yeu': 'Trích yếu',
                'file_url': 'File',
                'ghi_chu': 'Ghi chú'
            }
            df_display.rename(columns=col_map, inplace=True)
            
            st.caption(f"📌 Tổng số: {len(data_hd)} hợp đồng kinh tế")
            st.dataframe(df_display, width='stretch', hide_index=True, height=400)

            # ----- Sửa / Xóa hợp đồng kinh tế -----
            with st.expander("✏️ Sửa / 🗑️ Xóa hợp đồng kinh tế", expanded=False):
                tuy_chon_hd = {f"{r['so_hop_dong']} - {r.get('ten_doi_tac') or ''}": r for r in data_hd}
                chon_hd = st.selectbox("Chọn hợp đồng:", ["-- Chọn --"] + list(tuy_chon_hd.keys()), key="chon_sua_hd")
                if chon_hd != "-- Chọn --":
                    hd_sua = tuy_chon_hd[chon_hd]
                    col_s1, col_s2 = st.columns(2)
                    with col_s1:
                        ten_doi_tac_sua = st.text_input("Tên đối tác:", value=hd_sua.get('ten_doi_tac') or '', key=f"sua_dt_hd_{hd_sua['id']}")
                    with col_s2:
                        trich_yeu_sua_hd = st.text_area("Trích yếu:", value=hd_sua.get('trich_yeu') or '', key=f"sua_ty_hd_{hd_sua['id']}", height=80)
                    col_luu_hd, col_xoa_hd = st.columns(2)
                    with col_luu_hd:
                        if st.button("💾 Lưu thay đổi", key=f"btn_luu_hd_{hd_sua['id']}", type="primary", width='stretch'):
                            try:
                                db_s = st.session_state.db_engine.get_connection()
                                c_s = db_s.cursor()
                                c_s.execute("""
                                    UPDATE hop_dong_kinh_te SET ten_doi_tac=%s, trich_yeu=%s
                                    WHERE id=%s
                                """, (ten_doi_tac_sua, trich_yeu_sua_hd, hd_sua['id']))
                                db_s.commit(); db_s.close()
                                st.success("✅ Đã cập nhật hợp đồng")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")
                    with col_xoa_hd:
                        if st.button("🗑️ Xóa hợp đồng này", key=f"btn_xoa_hd_{hd_sua['id']}", width='stretch'):
                            try:
                                db_x = st.session_state.db_engine.get_connection()
                                c_x = db_x.cursor()
                                c_x.execute("DELETE FROM hop_dong_kinh_te WHERE id=%s", (hd_sua['id'],))
                                db_x.commit(); db_x.close()
                                st.success("✅ Đã xóa hợp đồng")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")

            # Nút xuất Excel
            col_export1, col_export2, col_export3 = st.columns([1, 2, 1])
            with col_export2:
                if st.button("📥 Xuất Excel hợp đồng", width='stretch', type="primary"):
                    headers = ['so_hop_dong', 'ten_doi_tac', 'ngay_ky', 'trich_yeu', 'ghi_chu']
                    col_widths = [20, 30, 12, 35, 25]
                    title = f"BÁO CÁO HỢP ĐỒNG KINH TẾ (Từ {tu_ngay_hd or '...'} đến {den_ngay_hd or '...'})"
                    filename = f"Hop_dong_kinh_te_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    excel_data = []
                    for row in data_hd:
                        excel_row = {}
                        for key in headers:
                            val = row.get(key)
                            if key == 'ngay_ky' and val:
                                val = format_date(val)
                            excel_row[key] = val
                        excel_data.append(excel_row)
                    
                    export_cong_van_excel(excel_data, filename, headers, col_widths, title)
                    
                    with open(filename, "rb") as f:
                        st.download_button(
                            label="📥 TẢI FILE EXCEL",
                            data=f,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch'
                        )
                    st.success(f"✅ Đã xuất {len(data_hd)} hợp đồng kinh tế")
        else:
            st.info("📭 Không có hợp đồng kinh tế nào")

# ========== KHỞI TẠO SESSION STATE ==========
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None

# show_hrm=False → hiện landing | show_hrm=True → vào HRM (sidebar login)
if 'show_hrm' not in st.session_state:
    st.session_state.show_hrm = False

# ========== KIỂM TRA URL PARAMS (Từ nút Nhân viên trên Landing Page) ==========
query_params = st.query_params
if query_params.get('goto') == 'hrm':
    st.session_state.show_hrm = True   # Chỉ thoát landing, KHÔNG tự đăng nhập
    st.query_params.clear()
    st.rerun()


# ========== HIỂN THỊ LANDING PAGE NẾU CHƯA VÀO HRM ==========
logo_url = COMPANY_CONFIG.get("logo_url")
if logo_url:
    with st.sidebar:
        st.image(logo_url, width='stretch')
        st.divider()
elif os.path.exists("logo_cty.png"):
    with st.sidebar:
        st.image("logo_cty.png", width='stretch')
        st.divider()

# ========== ĐÃ BỎ LANDING PAGE (mỗi khách hàng có domain riêng, vào thẳng màn hình đăng nhập) ==========
# show_hrm được ép luôn True để các đoạn code phía dưới (vốn kiểm tra show_hrm) không bị ảnh hưởng.
st.session_state.show_hrm = True

# ========== PHẦN CODE HRM BẮT ĐẦU TỪ ĐÂY ==========

st.markdown("""
    <style>
        /* ===== ẨN MANAGE APP - dùng mọi selector có thể ===== */
        [data-testid="stToolbar"],
        [data-testid="manage-app-button"],
        [data-testid="stAppDeployButton"],
        .stDeployButton,
        #MainMenu,
        div[class*="toolbar"],
        div[class*="StatusWidget"],
        div[class*="viewerBadge"],
        div[class*="manage-app"],
        button[kind="managedApp"],
        [data-testid="stBottom"] > div:last-child { 
            display: none !important; 
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
        }
        footer[data-testid] { display: none !important; }

        /* ===== PADDING TOP / BOTTOM = 5px ===== */
        .stApp > div[data-testid="stAppViewContainer"] > section[data-testid="stMain"] > div {
            padding-top: 5px !important;
            padding-bottom: 5px !important;
        }
        .block-container {
            padding-top: 5px !important;
            padding-bottom: 5px !important;
        }

        /* ===== LOGO SIDEBAR: hình tròn đổ bóng, sát top, căn giữa ===== */
        [data-testid="stSidebar"] > div:first-child {
            padding-top: 0 !important;
        }
        [data-testid="stSidebar"] [data-testid="stImage"] {
            display: flex !important;
            justify-content: center !important;
            margin-top: 0 !important;
            padding-top: 0 !important;
        }
        [data-testid="stSidebar"] [data-testid="stImage"] img {
            width: 150px !important;
            height: 150px !important;
            border-radius: 50% !important;
            object-fit: cover !important;
            box-shadow: 0 4px 20px rgba(0,0,0,0.25) !important;
            display: block !important;
            margin: 4px auto 0 auto !important;
        }
    </style>
    <script>
        // MutationObserver: ẩn Manage App ngay khi DOM thay đổi
        function hideManageApp() {
            const selectors = [
                '[data-testid="manage-app-button"]',
                '[data-testid="stToolbar"]',
                '[data-testid="stAppDeployButton"]',
                '.stDeployButton',
                'button[kind="managedApp"]'
            ];
            selectors.forEach(sel => {
                document.querySelectorAll(sel).forEach(el => {
                    el.style.cssText = 'display:none!important;visibility:hidden!important';
                });
            });
        }
        hideManageApp();
        // Quan sát DOM liên tục — bắt được dù Streamlit render trễ
        const observer = new MutationObserver(hideManageApp);
        observer.observe(document.body, { childList: true, subtree: true });
    </script>
""", unsafe_allow_html=True)

def to_float_or_none(val):
    """Chuyển đổi giá trị sang float hoặc None, tránh lỗi numeric"""
    if val is None or str(val).strip() == '':
        return None
    try:
        return float(val)
    except:
        return None

def format_date(d):
    if d is None or pd.isna(d): return ''
    try: return d.strftime('%d/%m/%Y') if hasattr(d,'strftime') else str(d)[:10]
    except: return str(d)

def parse_date(s):
    """Chuyển đổi nhiều định dạng ngày tháng về date object"""
    if not s or str(s).strip() == '':
        return None
    
    # Nếu đã là date object
    if hasattr(s, 'strftime'):
        return s
    
    s = str(s).strip()
    
    # Các định dạng cần thử
    formats = [
        '%d/%m/%Y',      # 18/04/2026
        '%d-%m-%Y',      # 18-04-2026
        '%Y-%m-%d',      # 2026-04-18
        '%Y/%m/%d',      # 2026/04/18
        '%d.%m.%Y',      # 18.04.2026
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    
    # Nếu không có định dạng nào phù hợp
    print(f"⚠️ Không thể parse ngày: {s}")
    return None

def get_xung_ho_trang_trong(gioi_tinh):
    """Xưng hô trang trọng theo giới tính, dùng cho Hợp đồng và các card hiển thị lãnh đạo
    cấp cao (Hội đồng Quản trị, Ban Tổng Giám đốc...).
    - Nam -> "Ông"; Nữ -> "Bà"; None/rỗng -> "Ông/Bà"
    """
    if gioi_tinh == "Nam":
        return "Ông"
    elif gioi_tinh == "Nữ":
        return "Bà"
    else:
        return "Ông/Bà"

def get_xung_ho(gioi_tinh, ho_ten=""):
    """
    Lấy cách xưng hô phù hợp dựa trên giới tính
    - Nếu giới tính là Nam -> trả về "Anh"
    - Nếu giới tính là Nữ -> trả về "Chị"
    - Nếu giới tính là None hoặc rỗng -> trả về "Anh/Chị"
    """
    if gioi_tinh == "Nam":
        return "Anh"
    elif gioi_tinh == "Nữ":
        return "Chị"
    else:
        return "Anh/Chị"

def get_loi_chuc_sinh_nhat(ho_ten, gioi_tinh, tuoi=None):
    """
    Tạo lời chúc sinh nhật có xưng hô phù hợp
    """
    xung_ho = get_xung_ho(gioi_tinh, ho_ten)
    
    loi_chuc = f"""
🎉🎂 CHÚC MỪNG SINH NHẬT {xung_ho.upper()} {ho_ten.upper()} 🎂🎉

Thân gửi {xung_ho}: {ho_ten},

Nhân dịp sinh nhật của {xung_ho}, thay mặt Ban Lãnh đạo Công ty CP Cảng Hòn La, 
xin gửi đến {xung_ho} những lời chúc tốt đẹp nhất.

Chúc {xung_ho} luôn mạnh khỏe, hạnh phúc và thành công trong công việc 
cũng như trong cuộc sống.

"""
    
    if tuoi:
        loi_chuc += f"Chúc mừng {xung_ho} tròn {tuoi} tuổi! 🎂\n\n"
    
    loi_chuc += f"""
Cảm ơn {xung_ho} đã luôn đồng hành và đóng góp cho sự phát triển của Công ty.

Trân trọng!

🏗️ CÔNG TY CP CẢNG HÒN LA
    """
    
    return loi_chuc

def auto_check_birthday():
    if 'sinh_nhat_hom_nay_list' not in st.session_state:
        st.session_state.sinh_nhat_hom_nay_list = []

    today_str = date.today().strftime('%Y-%m-%d')
    
    # Key mới: gắn với cả ngày lẫn trạng thái login
    # → mỗi lần login lại đều query DB, không bị cache nhầm
    check_key = f"{today_str}_{st.session_state.get('username', 'guest')}"
    
    if st.session_state.get('last_birthday_check') == check_key:
        return  # Đã check cho user này hôm nay rồi
    
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, dien_thoai, email_lien_he
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_sinh IS NOT NULL
            AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
            AND EXTRACT(DAY FROM ngay_sinh) = EXTRACT(DAY FROM CURRENT_DATE)
        """)
        birthday_today = c.fetchall()
        db.close()
        
        st.session_state.sinh_nhat_hom_nay_list = [
            {
                'ho_ten': nv['ho_ten'],
                'ma_nv': nv['ma_nv'],
                'xung_ho': get_xung_ho(nv.get('gioi_tinh'), nv['ho_ten'])
            }
            for nv in birthday_today
        ]
        
        for nv in birthday_today:
            xung_ho = get_xung_ho(nv.get('gioi_tinh'), nv['ho_ten'])
            st.toast(f"🎂 Sinh nhật {xung_ho} {nv['ho_ten']} hôm nay!", icon="🎂")
        
        # Đánh dấu đã check — dùng check_key gắn với username
        st.session_state.last_birthday_check = check_key
        
    except Exception as e:
        st.warning(f"⚠️ Không thể kiểm tra sinh nhật: {e}")

def da_chuyen_doi_chinh_thuc(nv_id):
    """Kiểm tra xem nhân viên đã có quyết định chuyển từ thử việc sang chính thức chưa"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT * FROM quyet_dinh_nhan_su 
            WHERE nhan_vien_id = %s AND loai_quyet_dinh = 'CHINH_THUC'
            ORDER BY ngay_quyet_dinh DESC LIMIT 1
        """, (nv_id,))
        result = c.fetchone()
        db.close()
        
        # Debug: in ra log để kiểm tra
        print(f"Checking nv_id={nv_id}, found={result is not None}")
        if result:
            print(f"Quyet dinh: {result}")
        
        return result is not None, result
    except Exception as e:
        print(f"Error in da_chuyen_doi_chinh_thuc: {e}")
        return False, None

def lay_thong_tin_truoc_chuyen_doi(nv_id):
    """Lấy thông tin nhân viên trước khi chuyển đổi (từ lich_su_cong_tac)"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        # Lấy lịch sử công tác cũ nhất (thời gian thử việc)
        c.execute("""
            SELECT * FROM lich_su_cong_tac 
            WHERE nhan_vien_id = %s 
            ORDER BY tu_ngay ASC LIMIT 1
        """, (nv_id,))
        result = c.fetchone()
        db.close()
        return result
    except:
        return None
# ========== DATABASE CONNECTION (SUPABASE) — ĐA KHÁCH HÀNG (MULTI-TENANT) ==========
def get_connection():
    """Wrapper tương thích ngược, tự động gọi db_engine từ st.session_state."""
    if 'db_engine' not in st.session_state:
        st.session_state.db_engine = DatabaseEngine(st.session_state.get('tenant'))
    return st.session_state.db_engine.get_connection()

# ========== SUPABASE STORAGE (lưu trữ file hồ sơ) ==========
# Tên bucket Storage trên Supabase dùng để lưu hồ sơ nhân viên.
# Cần tạo trước trên Supabase Dashboard > Storage (khuyến nghị để Private).
SUPABASE_BUCKET = "ho-so-nhan-vien"

def sanitize_storage_filename(filename):
    """Chuẩn hóa tên file để làm 'key' hợp lệ trên Supabase Storage:
    - Bỏ dấu tiếng Việt (Lộ_trình_học -> Lo_trinh_hoc)
    - Thay khoảng trắng bằng '_'
    - Chỉ giữ lại chữ cái không dấu, số, '_', '-', '.'
    Supabase Storage sẽ báo lỗi 'InvalidKey' nếu key chứa ký tự có dấu/unicode."""
    # Bỏ dấu: chuẩn hóa Unicode rồi loại bỏ các ký tự dấu (combining marks)
    normalized = unicodedata.normalize('NFD', filename)
    no_accent = ''.join(ch for ch in normalized if unicodedata.category(ch) != 'Mn')
    # Xử lý riêng chữ đ/Đ vì NFD không tách được
    no_accent = no_accent.replace('đ', 'd').replace('Đ', 'D')
    # Thay khoảng trắng bằng '_'
    no_accent = re.sub(r'\s+', '_', no_accent)
    # Chỉ giữ ký tự an toàn
    safe = re.sub(r'[^A-Za-z0-9_.\-]', '', no_accent)
    return safe or "file"

def upload_to_storage_unique(sb, bucket, base_path, file_bytes, content_type, max_tries=50):
    """Upload file lên Supabase Storage tại base_path. Nếu path đã tồn tại (trùng
    Loại hồ sơ + ngày upload + tên file trong cùng ngày), tự thêm hậu tố _2, _3...
    trước phần mở rộng để không bị lỗi/ghi đè. Trả về path thực tế đã dùng để upload."""
    path = base_path
    root, ext = os.path.splitext(base_path)
    tries = 0
    while True:
        try:
            sb.storage.from_(bucket).upload(
                path=path,
                file=file_bytes,
                file_options={"content-type": content_type or "application/octet-stream"}
            )
            return path
        except Exception as e:
            msg = str(e).lower()
            is_duplicate = ('duplicate' in msg or 'exists' in msg or 'already' in msg or '409' in msg)
            if is_duplicate and tries < max_tries:
                tries += 1
                path = f"{root}_{tries + 1}{ext}"
                continue
            raise

def get_supabase_storage():
    """Khởi tạo Supabase Client dùng cho Storage (ảnh NV, hồ sơ, file chat...).
    Ưu tiên url/key của TENANT đang đăng nhập (mô hình SaaS đa khách hàng).
    Fallback sang st.secrets['supabase'] / .env khi chạy chế độ đơn khách hàng.
    Không dùng @st.cache_resource nữa vì client giờ có thể khác nhau theo từng tenant
    trong cùng 1 tiến trình app (nhiều khách hàng dùng chung 1 deployment)."""
    try:
        from supabase import create_client
    except ImportError:
        print("Chưa cài thư viện supabase. Chạy: pip install supabase")
        return None

    tenant = st.session_state.get('tenant')
    if tenant:
        url, key = tenant['supabase_url'], tenant['supabase_key']
    else:
        url, key = None, None
        try:
            if 'supabase' in st.secrets:
                url = st.secrets.supabase.get('url')
                key = st.secrets.supabase.get('key')
        except Exception:
            pass
        if not url or not key:
            from dotenv import load_dotenv
            load_dotenv()
            url = url or os.getenv('SUPABASE_URL')
            key = key or os.getenv('SUPABASE_KEY')

    if not url or not key:
        return None

    cache_key = f"_sb_client_{url}"
    if cache_key in st.session_state:
        return st.session_state[cache_key]

    try:
        client = create_client(url, key)
        st.session_state[cache_key] = client
        return client
    except Exception as e:
        print(f"Lỗi khởi tạo Supabase Storage: {e}")
        return None

# ========== CHẤM CÔNG THỦ CÔNG - HẰNG SỐ & HÀM DÙNG CHUNG ==========
# Mã công theo đúng bảng chấm công mẫu (sheet T4-2026 / T5-2026 file HLP)
# Đây cũng chính là danh sách ký hiệu hợp lệ được liệt kê trong "Chú giải" và
# dùng để giới hạn dữ liệu nhập vào các ô ly (row Ca ngày / Ca đêm).
CHAM_CONG_MA_CODE = {
    "":    "(Trống) - Chưa chấm công",
    "X":   "X - Ngày công thường (đủ ca)",
    "P":   "P - Nghỉ phép hưởng lương",
    "V":   "V - Vắng mặt/nghỉ không lương",
    "N":   "N - Ca làm việc 8T - ngày",
    "D":   "D - Ca làm việc 8T - đêm",
    "L":   "L - Đi làm ngày lễ",
    "0.5": "0.5 - Làm nửa ngày công thường",
    "NL":  "NL - Nghỉ lễ hưởng nguyên lương",
}
CHAM_CONG_MA_OPTIONS = list(CHAM_CONG_MA_CODE.keys())
# Regex (dùng bởi Streamlit data_editor) giới hạn ký tự được phép nhập vào ô ly
# trong bảng chấm công dạng lịch. Vì 1 cột ngày dùng chung cho cả 3 loại dòng
# (Ca ngày / Ca đêm / Tăng ca) nên regex gộp cả 2 nhóm: mã chữ quy ước
# (X/P/V/N/D/L/NL/0.5) và số giờ tăng ca 0-9, để chặn ký tự rác mà không chặn
# nhầm dữ liệu hợp lệ của bất kỳ loại dòng nào.
CHAM_CONG_CELL_REGEX = r"^$|^[XxPpVvNnDdLl]$|^[Nn][Ll]$|^\d{1,2}(\.\d)?$"


def cc_pin_col(col_type, **kwargs):
    """Tạo column_config, cố gắng ghim (pin) cột vào bên trái khi cuộn ngang.
    Một số phiên bản Streamlit cũ chưa hỗ trợ tham số `pinned` -> fallback bỏ qua."""
    try:
        return col_type(pinned=True, **kwargs)
    except TypeError:
        return col_type(**kwargs)

# Những bộ phận (theo mã phong_ban_lam_viec trong bảng nhan_vien) có phát sinh tăng ca
# theo đúng bản mẫu (nhóm LX-M/LDPT thực tế đang lưu mã "SX" và "LDPT")
CHAM_CONG_DEPT_TANG_CA = ["SX", "LDPT"]

CHAM_CONG_DEPT_LABEL = {
    "QL": "QL - Quản lý",
    "VP": "VP - Văn phòng",
    "SX": "SX - Sản xuất/Vận hành",
    "LDPT": "LDPT - Lao động phổ thông",
}

# Bộ phận chỉ chấm công 1 dòng/nhân viên (giờ hành chính, không tách ca ngày/đêm/tăng ca)
CHAM_CONG_DEPT_MOT_DONG = ["VP"]

CC_ROW_HEIGHT = 24  # giảm size dòng (px) để bảng chấm công hiển thị gọn, nhiều dữ liệu hơn

def cc_render_grid(data, edit=False, **kwargs):
    """Wrapper cho st.dataframe/st.data_editor, cố gắng thu nhỏ chiều cao dòng (row_height)
    nếu phiên bản Streamlit đang chạy hỗ trợ; nếu không thì bỏ qua tham số đó."""
    fn = st.data_editor if edit else st.dataframe
    try:
        return fn(data, row_height=CC_ROW_HEIGHT, **kwargs)
    except TypeError:
        return fn(data, **kwargs)

def cc_normalize_marker(v):
    """Chuẩn hoá ký hiệu chấm công theo bảng CHAM_CONG_MA_CODE (không phân biệt hoa/thường).
    x/X = có công, P = nghỉ phép có lương, V = nghỉ không lương/vắng, N/D/L/NL/0.5 = các mã ca khác."""
    v = (v or "").strip()
    if not v:
        return None
    vu = v.upper()
    for code in CHAM_CONG_MA_CODE:
        if code and code.upper() == vu:
            return code  # trả về đúng dạng chuẩn đã khai báo, vd "X", "P", "V", "0.5"
    return v  # giữ nguyên ký hiệu lạ, không chặn để tránh mất dữ liệu người dùng đã nhập

def cc_is_cong(v):
    return isinstance(v, str) and v.strip().upper() == "X"

def cc_marker_is(v, target):
    return isinstance(v, str) and v.strip().upper() == target


def ensure_cham_cong_table():
    """Tạo bảng cham_cong trên Supabase nếu chưa có, và tự nâng cấp thêm cột mới (idempotent)."""
    db = st.session_state.db_engine.get_connection()
    c = db.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS cham_cong (
            id SERIAL PRIMARY KEY,
            nhan_vien_id INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE,
            ngay DATE NOT NULL,
            ma_cong VARCHAR(10),
            ca_ngay VARCHAR(10),
            ca_dem VARCHAR(10),
            gio_tang_ca NUMERIC(5,2) DEFAULT 0,
            gio_tang_ca_le NUMERIC(5,2) DEFAULT 0,
            ghi_chu TEXT,
            nguon VARCHAR(20) DEFAULT 'THU_CONG',
            created_by VARCHAR(100),
            created_at TIMESTAMP DEFAULT NOW(),
            updated_at TIMESTAMP DEFAULT NOW(),
            UNIQUE(nhan_vien_id, ngay)
        )
    """)
    # Nâng cấp cho DB đã tạo bảng từ phiên bản trước (chưa có ca_ngay/ca_dem)
    c.execute("ALTER TABLE cham_cong ADD COLUMN IF NOT EXISTS ca_ngay VARCHAR(10)")
    c.execute("ALTER TABLE cham_cong ADD COLUMN IF NOT EXISTS ca_dem VARCHAR(10)")
    db.commit()
    c.close()
    db.close()


def ensure_qdns_columns():
    """Bổ sung cột 'chuc_vu' và 'ngay_qd_ns' vào bảng nhan_vien nếu chưa có (idempotent).
    - chuc_vu: mặc định 'Nhân viên'
    - ngay_qd_ns: mặc định lấy theo ngay_vao_lam; sau này được cập nhật theo ngày ban hành
      các Quyết định nhân sự (bổ nhiệm/miễn nhiệm/đổi chức danh/điều chuyển)."""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS chuc_vu VARCHAR(100)")
        c.execute("ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ngay_qd_ns DATE")
        c.execute("ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS so_luong_npt INTEGER DEFAULT 0")
        c.execute("UPDATE nhan_vien SET chuc_vu = 'Nhân viên' WHERE chuc_vu IS NULL OR chuc_vu = ''")
        c.execute("UPDATE nhan_vien SET ngay_qd_ns = ngay_vao_lam WHERE ngay_qd_ns IS NULL")
        c.execute("UPDATE nhan_vien SET so_luong_npt = 0 WHERE so_luong_npt IS NULL")
        db.commit()
        c.close()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi ensure_qdns_columns: {e}")
        return False


def ensure_qdns_table():
    """Tạo bảng quyet_dinh_nhan_su lưu lịch sử các Quyết định nhân sự đã ban hành."""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS quyet_dinh_nhan_su (
                id SERIAL PRIMARY KEY,
                so_qd VARCHAR(50) NOT NULL,
                loai_qd VARCHAR(30) NOT NULL,
                nhan_vien_id INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE,
                ngay_qd DATE NOT NULL DEFAULT CURRENT_DATE,
                noi_dung TEXT,
                gia_tri_truoc VARCHAR(150),
                gia_tri_sau VARCHAR(150),
                file_url TEXT,
                nguoi_tao VARCHAR(100),
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        # Nâng cấp cho DB đã tạo bảng này từ phiên bản trước khi có đủ các cột trên
        # (CREATE TABLE IF NOT EXISTS không tự thêm cột còn thiếu vào bảng đã tồn tại)
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS so_qd VARCHAR(50)")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS loai_qd VARCHAR(30)")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS ngay_qd DATE DEFAULT CURRENT_DATE")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS noi_dung TEXT")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS gia_tri_truoc VARCHAR(150)")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS gia_tri_sau VARCHAR(150)")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS file_url TEXT")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS nguoi_tao VARCHAR(100)")
        c.execute("ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS created_at TIMESTAMP DEFAULT NOW()")
        db.commit()
        c.close()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi ensure_qdns_table: {e}")
        return False


# Danh sách chức vụ dùng cho QĐ Bổ nhiệm / Miễn nhiệm
DANH_SACH_CHUC_VU = ["Phó Tổng Giám đốc", "Trưởng phòng", "Phó Trưởng phòng", "Đội Trưởng", "Tổ Trưởng", "Quản đốc"]

LOAI_QDNS_LABEL = {
    'BO_NHIEM': 'QĐ Bổ nhiệm',
    'MIEN_NHIEM': 'QĐ Miễn nhiệm',
    'DOI_CHUC_DANH': 'QĐ Thay đổi chức danh',
    'DIEU_CHUYEN': 'QĐ Điều chuyển công tác',
    'CHAM_DUT_HD': 'QĐ Chấm dứt HĐTV/HĐLĐ',
}


def tao_quyet_dinh_nhan_su(nv, so_qd, ngay_qd, tieu_de, dieu1_lines, hieu_luc_text=None):
    """Tạo file Word Quyết định nhân sự dùng chung cho các loại: bổ nhiệm, miễn nhiệm,
    thay đổi chức danh, điều chuyển công tác, chấm dứt HĐTV/HĐLĐ."""
    CC = COMPANY_CONFIG
    doc = Document()
    s = doc.styles['Normal']; s.font.name = 'Times New Roman'; s.font.size = Pt(13)
    s.paragraph_format.space_after = Pt(4); s.paragraph_format.space_before = Pt(0)
    s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sec = doc.sections[0]; sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(3); sec.right_margin = Cm(2)

    ht = doc.add_table(rows=3, cols=2); ht.alignment = WD_TABLE_ALIGNMENT.CENTER; ht.autofit = False; remove_table_border(ht)
    for row in ht.rows:
        row.cells[0].width = Cm(7); row.cells[1].width = Cm(10)
    c = ht.rows[0].cells[0]; p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(CC.get('ten_cong_ty', 'CÔNG TY').upper()); r.bold = True; r.font.size = Pt(13)
    c = ht.rows[0].cells[1]; p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'); r.bold = True; r.font.size = Pt(13)
    c = ht.rows[1].cells[0]; p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f'Số: {so_qd}'); r.italic = True; r.font.size = Pt(12)
    c = ht.rows[1].cells[1]; p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Độc lập - Tự do - Hạnh phúc'); r.bold = True; r.italic = True; r.font.size = Pt(13)
    c = ht.rows[2].cells[1]; p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after = Pt(20)
    dia_diem = CC.get('dia_diem', 'Quảng Trị')
    ns = f'{dia_diem}, ngày {ngay_qd.day} tháng {ngay_qd.month:02d} năm {ngay_qd.year}' if hasattr(ngay_qd, 'day') else f'{dia_diem}, ngày ... tháng ... năm ......'
    r = p.add_run(ns); r.italic = True; r.font.size = Pt(13)

    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
    r = p.add_run('QUYẾT ĐỊNH'); r.bold = True; r.font.size = Pt(18)
    force_center(p)
    p = doc.add_paragraph()
    r = p.add_run(f'Về việc: {tieu_de}'); r.bold = True; r.italic = True; r.font.size = Pt(14)
    force_center(p)
    p = doc.add_paragraph()
    r = p.add_run(f"GIÁM ĐỐC {CC.get('ten_cong_ty', '').upper()}"); r.bold = True
    force_center(p)

    doc.add_paragraph('- Căn cứ Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019;')
    doc.add_paragraph('- Căn cứ Điều lệ tổ chức và hoạt động của Công ty;')
    doc.add_paragraph('- Căn cứ nhu cầu công tác và năng lực cán bộ, nhân viên;')
    doc.add_paragraph('- Xét đề nghị của Phòng Tổ chức - Hành chính,')

    p = doc.add_paragraph(); r = p.add_run('QUYẾT ĐỊNH:'); r.bold = True
    force_center(p)

    p = doc.add_paragraph(); r = p.add_run('Điều 1. '); r.bold = True
    r2 = p.add_run(dieu1_lines[0] if dieu1_lines else '')
    for extra_line in dieu1_lines[1:]:
        doc.add_paragraph(extra_line)

    p = doc.add_paragraph(); r = p.add_run('Điều 2. '); r.bold = True
    r2 = p.add_run(hieu_luc_text or f'Quyết định này có hiệu lực kể từ ngày {ngay_qd.day}/{ngay_qd.month:02d}/{ngay_qd.year}.')

    p = doc.add_paragraph(); r = p.add_run('Điều 3. '); r.bold = True
    r2 = p.add_run(f"Ông/Bà {nv.get('ho_ten', '')}, Trưởng phòng Tổ chức - Hành chính và các bộ phận có liên quan chịu trách nhiệm thi hành Quyết định này./.")

    doc.add_paragraph('')
    t2 = doc.add_table(rows=2, cols=2); remove_table_border(t2)
    p_nn = t2.rows[0].cells[0].paragraphs[0]
    r = p_nn.add_run('Nơi nhận:'); r.bold = True; r.font.size = Pt(11)
    p_nn2 = t2.rows[0].cells[0].add_paragraph()
    r = p_nn2.add_run('- Như Điều 3;\n- Lưu: VT, HSNV.'); r.font.size = Pt(11)
    p_kh = t2.rows[0].cells[1].paragraphs[0]; p_kh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_kh.add_run(CC.get('chuc_vu', 'GIÁM ĐỐC').upper()); r.bold = True
    p_kh2 = t2.rows[1].cells[1].add_paragraph(); p_kh2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_kh2.add_run('\n\n\n' + CC.get('dai_dien', ''))

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(tmp.name)
    return tmp.name


if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None
if 'selected_nv_id' not in st.session_state:
    st.session_state.selected_nv_id = None
if 'edit_uv_id' not in st.session_state:
    st.session_state.edit_uv_id = None
if 'bhxh_family_nv_id' not in st.session_state:
    st.session_state.bhxh_family_nv_id = None
if 'bhxh_family_nv_name' not in st.session_state:
    st.session_state.bhxh_family_nv_name = None
if 'bhxh_family_members' not in st.session_state:
    st.session_state.bhxh_family_members = []
if 'show_chuyen_nv_form' not in st.session_state:
    st.session_state.show_chuyen_nv_form = False
if 'chuyen_uv_id' not in st.session_state:
    st.session_state.chuyen_uv_id = None
if 'chuyen_uv_data' not in st.session_state:
    st.session_state.chuyen_uv_data = {}

def force_center(p):
    pPr = p._p.get_or_add_pPr()
    for jc in pPr.findall(qn('w:jc')):
        pPr.remove(jc)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)

st.markdown("""
<style>
    [data-testid="stDataFrame"] > div {
        overflow-x: auto !important;
    }
    [data-testid="stDataFrame"] table {
        min-width: 2000px !important;
        width: max-content !important;
    }
    /* ===== Bảng chấm công (BCC): auto center Horizontal + Vertical, giảm size chữ ===== */
    [data-testid="stDataFrame"] th,
    [data-testid="stDataFrame"] td {
        text-align: center !important;
        vertical-align: middle !important;
        font-size: 12px !important;
    }
    [data-testid="stDataFrame"] [data-testid="stElementToolbar"] { font-size: 12px !important; }
    /* st.data_editor (bảng chấm công dạng edit) dùng cùng component nền glide-data-grid */
    [data-testid="stDataEditor"] th,
    [data-testid="stDataEditor"] td {
        text-align: center !important;
        vertical-align: middle !important;
        font-size: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def to_int_or_none(val):
    """Chuyển đổi giá trị sang int hoặc None"""
    if val is None or str(val).strip() == '':
        return None
    try:
        return int(float(val))
    except:
        return None
        
def tao_noi_dung_zalo(nv):
    ZC = COMPANY_CONFIG
    return f"""Gửi anh/chị: {nv.get('ho_ten','')},

Thông tin đã cập nhật:
- Họ tên: {nv.get('ho_ten','')}
- Ngày sinh: {format_date(nv.get('ngay_sinh'))}
- CCCD: {nv.get('so_cccd','')}
- Ngày cấp: {format_date(nv.get('ngay_cap_cccd'))}
- Thường trú: {nv.get('thuong_tru','')}
- Số BHXH: {nv.get('ma_so_bhxh','')}
- TK NH: {nv.get('so_tai_khoan_nh','')}
- CN NH: {nv.get('chi_nhanh_nh','')}
- Tên đơn vị thụ hưởng: {nv.get('ten_don_vi_thu_huong','')}

{ZC.get('loi_nhan_zalo','Vui lòng kiểm tra và phản hồi nếu có sai sót. Xin Cảm ơn!')}"""

def remove_table_border(tbl):
    for row in tbl.rows:
        for cell in row.cells:
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            b = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders')
            if b is not None: tcPr.remove(b)

# ========== CÁC HÀM TẠO HỢP ĐỒNG (GIỮ NGUYÊN) ==========
def ensure_mau_dieu_hop_dong_table():
    """Bảng lưu nội dung tuỳ chỉnh của từng Điều trong HĐLĐ/HĐTV, do admin cấu hình.
    Mỗi dòng = 1 Điều của 1 loại hợp đồng. Nếu chưa có dòng nào cho 1 mã Điều,
    hệ thống dùng nội dung mặc định (DEFAULT_DIEU_HDLD / DEFAULT_DIEU_HDTV)."""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS mau_dieu_hop_dong (
                id SERIAL PRIMARY KEY,
                loai_hd VARCHAR(10) NOT NULL,
                ma_dieu VARCHAR(30) NOT NULL,
                tieu_de TEXT,
                noi_dung TEXT,
                thu_tu INT DEFAULT 0,
                updated_at TIMESTAMP DEFAULT NOW(),
                UNIQUE(loai_hd, ma_dieu)
            )
        """)
        db.commit()
        c.close(); db.close()
        return True
    except Exception as e:
        print(f"Lỗi ensure_mau_dieu_hop_dong_table: {e}")
        return False

def ensure_chuc_danh_ung_vien_table():
    """Danh mục 'Vị trí dự tuyển' RIÊNG cho Ứng viên - độc lập với vi_tri_cong_tac (chức danh
    Nhân viên), để đổi danh mục chức danh Nhân viên không làm mất khả năng tìm/nhập ứng viên
    theo các chức danh cũ. Tự động khởi tạo dữ liệu từ các giá trị vi_tri_du_tuyen đã có sẵn
    trong bảng ung_vien (chỉ chạy 1 lần, khi bảng còn trống)."""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS chuc_danh_ung_vien (
                id SERIAL PRIMARY KEY,
                ten_chuc_danh VARCHAR(150) UNIQUE NOT NULL,
                ghi_chu TEXT,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        db.commit()
        c.execute("SELECT COUNT(*) FROM chuc_danh_ung_vien")
        so_luong = c.fetchone()[0]
        if so_luong == 0:
            # Nạp dữ liệu ban đầu từ các vị trí dự tuyển đã có trong ung_vien (chức danh cũ)
            c.execute("""SELECT DISTINCT vi_tri_du_tuyen FROM ung_vien 
                         WHERE vi_tri_du_tuyen IS NOT NULL AND vi_tri_du_tuyen != ''""")
            ds_cu = [row[0] for row in c.fetchall()]
            for ten in ds_cu:
                c.execute("INSERT INTO chuc_danh_ung_vien (ten_chuc_danh) VALUES (%s) ON CONFLICT DO NOTHING", (ten,))
            db.commit()
        c.close(); db.close()
        return True
    except Exception as e:
        print(f"Lỗi ensure_chuc_danh_ung_vien_table: {e}")
        return False

# Nội dung MẶC ĐỊNH của từng Điều — dùng khi admin CHƯA tuỳ chỉnh gì.
# Dòng bắt đầu bằng "## " sẽ được in đậm (tiêu đề phụ, VD "1. Nghĩa vụ:").
# Có thể dùng {vi_tri}, {ngay_hieu_luc} trong nội dung Điều 1 — sẽ tự thay bằng thông tin nhân viên.
DEFAULT_DIEU_HDLD = {
    "dieu1": ("Điều 1. Thời hạn và công việc hợp đồng:",
        "-    Bên B làm việc theo chế độ hợp đồng lao động không xác định thời hạn;\n"
        "-    Thời gian: Từ ngày {ngay_hieu_luc};\n"
        "-    Địa điểm làm việc: Tại Cảng tổng hợp quốc tế Hòn La và các địa điểm khác theo sự sắp xếp của Công ty;\n"
        "-    Vị trí: {vi_tri};\n"
        "-    Công việc phải làm: Thực hiện công việc theo đúng chuyên môn dưới sự quản lý, điều hành của cấp trên;\n"
        "-    Mức lương và phụ cấp: Theo thỏa thuận;\n"
        "-    Hình thức trả lương: Tiền mặt hoặc chuyển khoản, theo lần chi trả;\n"
        "-    Kỳ hạn trả lương: Theo quy định Công ty;\n"
        "-    Chế độ nâng lương: Theo thỏa thuận."),
    "dieu2": ("Điều 2. Chế độ làm việc:",
        "-    Thời gian làm việc: Theo tính chất công việc, do nhu cầu kinh doanh của Công ty nên thời gian làm việc của bên B là linh hoạt nhưng phải đảm bảo hoàn thành công việc được giao;\n"
        "-    Thời gian nghỉ ngơi của người lao động: Theo thỏa thuận và phù hợp với quy định của pháp luật;\n"
        "-    Ngoài giờ làm việc: Người lao động phải tự chịu trách nhiệm về các hoạt động cá nhân của mình."),
    "dieu3": ("Điều 3. Nghĩa vụ, quyền lợi NLĐ:",
        "## 1. Nghĩa vụ:\n"
        "-    Hoàn thành những công việc được giao và sẵn sàng chấp nhận mọi sự điều động khi có yêu cầu;\n"
        "-    Chấp hành nghiêm túc nội quy, kỷ luật lao động, an toàn lao động và các quy định của Công ty và pháp luật của Nhà nước;\n"
        "-    Người lao động có trách nhiệm tuân thủ đầy đủ quy định về an toàn lao động, quy trình vận hành thiết bị và hướng dẫn của Công ty. Trường hợp NLĐ cố ý vi phạm hoặc vi phạm nghiêm trọng quy định an toàn lao động gây thiệt hại thì phải chịu trách nhiệm theo quy định pháp luật và nội quy Công ty;\n"
        "-    Bồi thường vi phạm vật chất : Phải bồi thường vật chất do cá nhân vi phạm quy định của Công ty về bảo quản trang thiết bị được giao.\n"
        "## 2. Quyền Lợi:\n"
        "-    Phương tiện đi lại: Tự túc;\n"
        "-    Được Công ty đóng Bảo hiểm xã hội, bảo hiểm y tế, BHTN: theo chế độ hiện hành của Nhà nước và Quy định của Công ty;\n"
        "-    Được Công ty cấp đầy đủ bảo hộ lao động theo đúng vị trí làm việc;\n"
        "-    Được phân công công việc theo yêu cầu của Công ty phù hợp với khả năng và trình độ chuyên môn mà người lao động đáp ứng;\n"
        "-    Các quyền lợi khác thực hiện theo quy định của Pháp luật Lao động như tạm dừng, chấm dứt hợp đồng."),
    "dieu4": ("Điều 4. Nghĩa vụ, quyền hạn NSDLĐ:",
        "-    Bảo đảm việc làm và thực hiện đầy đủ những điều đã cam kết trong hợp đồng;\n"
        "-    Thanh toán đầy đủ, đúng hạn các chế độ và quyền lợi cho người lao động theo hợp đồng;\n"
        "-    Điều hành người lao động hoàn thành công việc theo hợp đồng;\n"
        "-    Tạm hoãn, chấm dứt hợp đồng, kỷ luật người lao động theo quy định của pháp luật, và nội quy lao động của Công ty."),
    "dieu5": ("Điều 5. Điều khoản chung:",
        "-    Những nội dung về quan hệ lao động không ghi trong hợp đồng này thì được áp dụng theo pháp luật lao động;\n"
        "-    Những thoả thuận khác (nếu có): không;\n"
        "-    Hợp đồng này có hiệu lực từ ngày ký và được làm thành 02 bản, Bên A giữ 01 bản, Bên B giữ 01 có giá trị pháp lý như nhau, để làm căn cứ thực hiện."),
}

DEFAULT_DIEU_HDTV = {
    "dieu1": ("Điều 1. Thời hạn và công việc hợp đồng:",
        "-    Bên B làm việc theo chế độ hợp đồng thử việc, có thời hạn 01 tháng;\n"
        "-    Bắt đầu: {ngay_bat_dau};\n"
        "-    Kết thúc: {ngay_ket_thuc};\n"
        "-    Địa điểm làm việc: Tại Cảng tổng hợp quốc tế Hòn La và các địa điểm khác theo sự sắp xếp của Công ty;\n"
        "-    Vị trí: {vi_tri};\n"
        "-    Công việc phải làm: Thực hiện công việc theo đúng chuyên môn dưới sự quản lý, điều hành của cấp trên;\n"
        "-    Mức lương và phụ cấp: Theo thỏa thuận;\n"
        "-    Hình thức trả lương: Tiền mặt hoặc chuyển khoản, theo lần chi trả;\n"
        "-    Kỳ hạn trả lương: Theo quy định Công ty."),
    "dieu2": ("Điều 2. Chế độ làm việc:",
        "-    Thời gian làm việc: Theo tính chất công việc, do nhu cầu kinh doanh của Công ty nên thời gian làm việc của bên B là linh hoạt nhưng phải đảm bảo hoàn thành công việc được giao;\n"
        "-    Thời gian nghỉ ngơi của người lao động: Theo thỏa thuận và phù hợp với quy định của pháp luật;\n"
        "-    Ngoài giờ làm việc: Người lao động phải tự chịu trách nhiệm về các hoạt động cá nhân của mình."),
    "dieu3": ("Điều 3. Nghĩa vụ, quyền lợi NLĐ:",
        "## 1. Nghĩa vụ:\n"
        "-    Hoàn thành những công việc được giao và sẵn sàng chấp nhận mọi sự điều động khi có yêu cầu;\n"
        "-    Chấp hành nghiêm túc nội quy, kỷ luật lao động, an toàn lao động và các quy định của Công ty và pháp luật của Nhà nước;\n"
        "-    Người lao động có trách nhiệm tuân thủ đầy đủ quy định về an toàn lao động, quy trình vận hành thiết bị và hướng dẫn của Công ty. Trường hợp NLĐ cố ý vi phạm hoặc vi phạm nghiêm trọng quy định an toàn lao động gây thiệt hại thì phải chịu trách nhiệm theo quy định pháp luật và nội quy Công ty;\n"
        "-    Bồi thường vi phạm vật chất : Phải bồi thường vật chất do cá nhân vi phạm quy định của Công ty về bảo quản trang thiết bị được giao.\n"
        "## 2. Quyền Lợi:\n"
        "-    Phương tiện đi lại: Tự túc;\n"
        "-    Được Công ty cấp đầy đủ bảo hộ lao động theo đúng vị trí làm việc;\n"
        "-    Được phân công công việc theo yêu cầu của Công ty phù hợp với khả năng và trình độ chuyên môn mà người lao động đáp ứng;\n"
        "-    Các quyền lợi khác thực hiện theo quy định của Pháp luật Lao động như tạm dừng, chấm dứt hợp đồng."),
    "dieu4": ("Điều 4. Nghĩa vụ, quyền hạn NSDLĐ:",
        "-    Bảo đảm việc làm và thực hiện đầy đủ những điều đã cam kết trong hợp đồng;\n"
        "-    Thanh toán đầy đủ, đúng hạn các chế độ và quyền lợi cho người lao động theo hợp đồng;\n"
        "-    Điều hành người lao động hoàn thành công việc theo hợp đồng;\n"
        "-    Tạm hoãn, chấm dứt hợp đồng theo quy định của pháp luật, và nội quy lao động của Công ty;"),
    "dieu5": ("Điều 5. Điều khoản chung:",
        "-    Những nội dung về quan hệ lao động không ghi trong hợp đồng này thì được áp dụng theo pháp luật lao động;\n"
        "-    Những thoả thuận khác (nếu có): không;\n"
        "-    Hợp đồng này có hiệu lực từ ngày ký và được làm thành 02 bản, Bên A giữ 01 bản, Bên B giữ 01 có giá trị pháp lý như nhau, để làm căn cứ thực hiện."),
}

class _SafeDict(dict):
    """Dict để .format_map() không lỗi khi thiếu placeholder — giữ nguyên {ten} nếu không có dữ liệu."""
    def __missing__(self, key):
        return '{' + key + '}'

@st.cache_data(ttl=60, show_spinner=False)
def get_all_dieu_hop_dong(loai_hd):
    """Lấy toàn bộ nội dung Điều đã tuỳ chỉnh (nếu có) cho 1 loại hợp đồng ('HDLD'/'HDTV').
    Trả về dict {ma_dieu: (tieu_de, noi_dung, thu_tu)}."""
    ket_qua = {}
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT ma_dieu, tieu_de, noi_dung, thu_tu FROM mau_dieu_hop_dong WHERE loai_hd=%s ORDER BY thu_tu ASC, ma_dieu ASC", (loai_hd,))
        for ma_dieu, tieu_de, noi_dung, thu_tu in c.fetchall():
            ket_qua[ma_dieu] = (tieu_de, noi_dung, thu_tu)
        db.close()
    except Exception:
        pass
    return ket_qua

def get_ds_ma_dieu(tuy_chinh):
    """Trả về danh sách mã Điều theo đúng thứ tự hiển thị/in ấn: 5 Điều mặc định (dieu1..dieu5)
    luôn giữ nguyên thứ tự gốc, cộng thêm các Điều admin tự thêm mới (không giới hạn số lượng)
    được sắp xếp chèn theo cột thu_tu."""
    mac_dinh_keys = ["dieu1", "dieu2", "dieu3", "dieu4", "dieu5"]
    them_moi = [(md, (info[2] if len(info) > 2 and info[2] else 999)) 
                for md, info in tuy_chinh.items() if md not in mac_dinh_keys]
    them_moi.sort(key=lambda x: (x[1], x[0]))
    return mac_dinh_keys + [md for md, _ in them_moi]

def sinh_ma_dieu_moi(tuy_chinh_hdld, tuy_chinh_hdtv):
    """Sinh mã Điều mới tự động (dieu6, dieu7, ...) không trùng với bất kỳ Điều nào đã có
    ở CẢ 2 loại hợp đồng, để tránh nhầm lẫn khi admin chuyển qua lại giữa HĐLĐ/HĐTV."""
    so_hien_co = [5]
    for md in list(tuy_chinh_hdld.keys()) + list(tuy_chinh_hdtv.keys()):
        if md.startswith("dieu") and md[4:].isdigit():
            so_hien_co.append(int(md[4:]))
    return f"dieu{max(so_hien_co) + 1}"

def get_dieu_content(loai_hd, ma_dieu, tuy_chinh, mac_dinh):
    """Trả về (tieu_de, noi_dung) — ưu tiên bản admin đã tuỳ chỉnh, nếu chưa có thì dùng mặc định."""
    if ma_dieu in tuy_chinh:
        info = tuy_chinh[ma_dieu]
        return info[0], info[1]
    return mac_dinh.get(ma_dieu, ("", ""))

def render_dieu(doc, add_p, tieu_de, noi_dung, context=None):
    """In 1 Điều ra file Word: tiêu đề in đậm, các dòng nội dung xuống dòng theo \\n.
    Dòng bắt đầu bằng '## ' sẽ in đậm (tiêu đề phụ như '1. Nghĩa vụ:')."""
    if tieu_de:
        p = doc.add_paragraph(); r = p.add_run(tieu_de); r.bold = True
    if not noi_dung:
        return
    text = noi_dung.format_map(_SafeDict(context or {}))
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        if line.startswith('## '):
            p = doc.add_paragraph(); r = p.add_run(line[3:]); r.bold = True
        else:
            add_p(line)


def tao_hop_dong(nv):
    """In Hợp đồng lao động (không xác định thời hạn). Nội dung 5 Điều lấy từ bảng
    mau_dieu_hop_dong nếu admin đã tuỳ chỉnh (Danh mục → Mẫu Điều khoản Hợp đồng),
    nếu chưa có thì dùng nội dung mặc định DEFAULT_DIEU_HDLD."""
    CC = COMPANY_CONFIG; doc = Document()
    s = doc.styles['Normal']; s.font.name='Times New Roman'; s.font.size=Pt(13)
    s.paragraph_format.space_after=Pt(0); s.paragraph_format.space_before=Pt(0)
    sec = doc.sections[0]; sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(3.5); sec.right_margin=Cm(2)
    def add_p(text='', bold=False, size=Pt(13)):
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        if bold and p.runs:
            p.runs[0].bold = True
        if p.runs:
            p.runs[0].font.size = size
        return p
    def al(label,value):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(5))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r=p.add_run(f'{label}'); r.font.size=Pt(13)
        r=p.add_run('\t: '); r.font.size=Pt(13)
        r=p.add_run(f'{value}'); r.font.size=Pt(13)
    ht=doc.add_table(rows=4,cols=2); ht.alignment=WD_TABLE_ALIGNMENT.CENTER; ht.autofit=False; remove_table_border(ht)
    for row in ht.rows: row.cells[0].width=Cm(6); row.cells[1].width=Cm(11)
    c=ht.rows[0].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CÔNG TY CỔ PHẦN'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[0].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CẢNG HÒN LA'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Độc lập - Tự do - Hạnh phúc'); r.bold=True; r.italic=True; r.font.size=Pt(13)
    c=ht.rows[2].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*12); r.font.size=Pt(9)
    c=ht.rows[2].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*20); r.font.size=Pt(9)
    c=ht.rows[3].cells[1]; p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(20)
    ngay_ky = nv.get("ngay_ky_hd")
    ngay_vao = nv.get("ngay_vao_lam")
    nk = ngay_ky if ngay_ky else ngay_vao
    ns = 'Quảng Trị, ngày ... tháng ... năm ......'
    if nk and hasattr(nk, 'day'):
        ns = f'Quảng Trị, ngày {nk.day} tháng {nk.month:02d} năm {nk.year}'
    run = p.add_run(ns)
    run.font.size = Pt(13)
    run.italic = True
    c=ht.rows[3].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f'Số: {nv.get("so_hdld","...")}'); r.italic=True; r.font.size=Pt(12)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run('HỢP ĐỒNG LAO ĐỘNG')
    r.bold = True
    r.font.size = Pt(18)
    force_center(p)
    p2 = doc.add_paragraph('- Căn cứ thông tư 10/2020/TT-LĐTBXH ngày 12/11/2020 hướng dẫn thi hành một số điều của Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019 về nội dung của hợp đồng lao động;')
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p2 = doc.add_paragraph('- Căn cứ nhu cầu sử dụng lao động trong đơn vị.')
    doc.add_paragraph('Chúng tôi gồm:')
    p=doc.add_paragraph(); r=p.add_run(f'BÊN A: {CC["ten_cong_ty"]} (Người sử dụng LĐ)'); r.bold=True
    al('Đại diện',f"Ông {CC['dai_dien']}"); al('Chức vụ',CC['chuc_vu']); al('Mã số thuế',CC['ma_so_thue'])
    al('Điện thoại',CC['dien_thoai_cty']); al('Địa chỉ',CC['dia_chi']); doc.add_paragraph()
    p=doc.add_paragraph(); r=p.add_run('BÊN B: (Người lao động)'); r.bold=True
    sk=nv.get('so_tai_khoan_nh','')
    if nv.get('chi_nhanh_nh'): sk+=f' - {nv.get("chi_nhanh_nh")}'
    gt = nv.get('gioi_tinh','')
    xung_ho = get_xung_ho_trang_trong(gt)
    al(xung_ho, nv.get('ho_ten',''))
    al('Ngày sinh',format_date(nv.get('ngay_sinh')))
    al('Số CMND/CCCD',nv.get('so_cccd','')); al('Ngày cấp',format_date(nv.get('ngay_cap_cccd')))
    al('Nơi cấp',nv.get('noi_cap_cccd','')); al('Số TKNH',sk)
    al('Điện thoại',nv.get('dien_thoai','')); al('Thường trú',nv.get('thuong_tru',''))
    doc.add_paragraph('Thoả thuận ký kết Hợp đồng lao động với những điều khoản dưới đây:')
    ngay_hieu_luc = nv.get("ngay_ky_hd") or nv.get("ngay_vao_lam")
    ns2 = '.../.../..........'
    if ngay_hieu_luc and hasattr(ngay_hieu_luc, 'day'):
        ns2 = f'{ngay_hieu_luc.day} tháng {ngay_hieu_luc.month:02d} năm {ngay_hieu_luc.year}'
    elif ngay_hieu_luc:
        ns2 = str(ngay_hieu_luc)
    # ===== NỘI DUNG CÁC ĐIỀU: lấy bản admin đã tuỳ chỉnh (nếu có), fallback về mặc định =====
    tuy_chinh_hdld = get_all_dieu_hop_dong('HDLD')
    ctx_hdld = {"vi_tri": nv.get("chuc_danh_nghe", ""), "ngay_hieu_luc": ns2}
    for ma_dieu in get_ds_ma_dieu(tuy_chinh_hdld):
        tieu_de, noi_dung = get_dieu_content("HDLD", ma_dieu, tuy_chinh_hdld, DEFAULT_DIEU_HDLD)
        if not tieu_de and not noi_dung:
            continue
        render_dieu(doc, add_p, tieu_de, noi_dung, context=ctx_hdld)
    add_p('Bản HĐ này lập tại văn phòng Công ty CP Cảng Hòn La.'); doc.add_paragraph()
    ts=doc.add_table(rows=3,cols=2); ts.alignment=WD_TABLE_ALIGNMENT.CENTER; remove_table_border(ts)
    c=ts.rows[0].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[0].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI SỬ DỤNG LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[1].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[1].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[2].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(nv.get('ho_ten','').upper()); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[2].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(CC['dai_dien'].upper()); r.bold=True; r.font.size=Pt(13)
    tf=tempfile.NamedTemporaryFile(delete=False,suffix='.docx'); doc.save(tf.name); return tf.name

def tao_hop_dong_thu_viec(nv):
    # ... giữ nguyên code cũ (tương tự)
    CC = COMPANY_CONFIG; doc = Document()
    s = doc.styles['Normal']; s.font.name='Times New Roman'; s.font.size=Pt(13)
    s.paragraph_format.space_after=Pt(0); s.paragraph_format.space_before=Pt(0)
    s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sec = doc.sections[0]; sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(3.5); sec.right_margin=Cm(2)
    def al(label,value):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(5))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r=p.add_run(f'{label}'); r.font.size=Pt(13)
        r=p.add_run('\t: '); r.font.size=Pt(13)
        r=p.add_run(f'{value}'); r.font.size=Pt(13)
    def add_p(text='', bold=False, size=Pt(13)):
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        if bold and p.runs:
            p.runs[0].bold = True
        if p.runs:
            p.runs[0].font.size = size
        return p
    ht=doc.add_table(rows=4,cols=2); ht.alignment=WD_TABLE_ALIGNMENT.CENTER; ht.autofit=False; remove_table_border(ht)
    for row in ht.rows: row.cells[0].width=Cm(6); row.cells[1].width=Cm(11)
    c=ht.rows[0].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CÔNG TY CỔ PHẦN'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[0].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CẢNG HÒN LA'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Độc lập - Tự do - Hạnh phúc'); r.bold=True; r.italic=True; r.font.size=Pt(13)
    c=ht.rows[2].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*12); r.font.size=Pt(9)
    c=ht.rows[2].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*20); r.font.size=Pt(9)
    c=ht.rows[3].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f'Số: {nv.get("so_hdld",".../.../HĐTV-CHL")}'); r.italic=True; r.font.size=Pt(12)
    c=ht.rows[3].cells[1]; p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(20)
    nk=nv.get("ngay_vao_lam") or nv.get("ngay_ky_hd")
    ns='Quảng Trị, ngày ... tháng ... năm ......'
    if nk and hasattr(nk,'day'): ns=f'Quảng Trị, ngày {nk.day} tháng {nk.month:02d} năm {nk.year}'
    run = p.add_run(ns)
    run.font.size = Pt(13)
    run.italic = True
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run('HỢP ĐỒNG THỬ VIỆC')
    r.bold = True
    r.font.size = Pt(18)
    force_center(p)
    p2 = doc.add_paragraph('- Căn cứ thông tư 10/2020/TT-LĐTBXH ngày 12/11/2020 hướng dẫn thi hành một số điều của Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019 về nội dung của hợp đồng lao động;')
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph('- Căn cứ nhu cầu sử dụng lao động trong đơn vị.')
    doc.add_paragraph('Chúng tôi gồm:')
    p=doc.add_paragraph(); r=p.add_run(f'BÊN A: {CC["ten_cong_ty"]} (Người sử dụng LĐ)'); r.bold=True
    al('Đại diện',f"Ông {CC['dai_dien']}"); al('Chức vụ',CC['chuc_vu']); al('Mã số thuế',CC['ma_so_thue'])
    al('Điện thoại',CC['dien_thoai_cty']); al('Địa chỉ',CC['dia_chi'])
    p=doc.add_paragraph(); r=p.add_run('BÊN B: (Người lao động)'); r.bold=True
    sk=nv.get('so_tai_khoan_nh','')
    if nv.get('chi_nhanh_nh'): sk+=f' - {nv.get("chi_nhanh_nh")}'
    gt = nv.get('gioi_tinh','')
    xung_ho = get_xung_ho_trang_trong(gt)
    al(xung_ho, nv.get('ho_ten',''))
    al('Ngày sinh',format_date(nv.get('ngay_sinh')))
    al('Số CMND/CCCD',nv.get('so_cccd','')); al('Ngày cấp',format_date(nv.get('ngay_cap_cccd')))
    al('Nơi cấp',nv.get('noi_cap_cccd','')); al('Số TKNH',sk)
    al('Điện thoại',nv.get('dien_thoai','')); al('Thường trú',nv.get('thuong_tru',''))
    doc.add_paragraph('Thoả thuận ký kết Hợp đồng Thử việc với những điều khoản dưới đây:')
    nkt = nk + timedelta(days=30) if (nk and hasattr(nk, 'day')) else None
    ns_bd = f'{nk.day:02d}/{nk.month:02d}/{nk.year}' if (nk and hasattr(nk, 'day')) else '.../.../......'
    ns_kt = f'{nkt.day:02d}/{nkt.month:02d}/{nkt.year}' if nkt else '.../.../......'
    # ===== NỘI DUNG CÁC ĐIỀU: lấy bản admin đã tuỳ chỉnh (nếu có), fallback về mặc định =====
    tuy_chinh_hdtv = get_all_dieu_hop_dong('HDTV')
    ctx_hdtv = {"vi_tri": nv.get("chuc_danh_nghe", ""), "ngay_bat_dau": ns_bd, "ngay_ket_thuc": ns_kt}
    for ma_dieu in get_ds_ma_dieu(tuy_chinh_hdtv):
        tieu_de, noi_dung = get_dieu_content("HDTV", ma_dieu, tuy_chinh_hdtv, DEFAULT_DIEU_HDTV)
        if not tieu_de and not noi_dung:
            continue
        render_dieu(doc, add_p, tieu_de, noi_dung, context=ctx_hdtv)
    add_p('Bản HĐ này lập tại văn phòng Công ty CP Cảng Hòn La.'); doc.add_paragraph()
    ts=doc.add_table(rows=3,cols=2); ts.alignment=WD_TABLE_ALIGNMENT.CENTER; remove_table_border(ts)
    c=ts.rows[0].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[0].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI SỬ DỤNG LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[1].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[1].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[2].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(nv.get('ho_ten','').upper()); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[2].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(CC['dai_dien'].upper()); r.bold=True; r.font.size=Pt(13)
    tf = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(tf.name)
    return tf.name

def gui_email_don(to_email, subject, html_body):
    """Gửi 1 email đơn (dùng cho OTP reset mật khẩu, thông báo cá nhân...) - dùng chung EMAIL_CONFIG."""
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG['email']
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))
        srv = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
        srv.starttls()
        srv.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
        srv.send_message(msg)
        srv.quit()
        return True
    except Exception as e:
        print(f"Lỗi gửi email OTP: {e}")
        return False

def gui_email(loai, ds, file=None):
    # Không import trong hàm nữa, dùng EMAIL_CONFIG đã có sẵn
    # from config import EMAIL_CONFIG as EC  <-- XÓA DÒNG NÀY
    
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG['email']
        msg['To'] = EMAIL_CONFIG['nguoi_nhan']
        tn = f"{datetime.now().month:02d}/{datetime.now().year}"
        msg['Subject'] = f"[HRM-Port] Báo cáo {loai} lao động tháng {tn}"
        nd = f"<h3>BÁO CÁO {loai.upper()} LĐ</h3><p>Tháng: <b>{tn}</b></p><p>SL: <b>{len(ds)}</b></p><hr><ul>"
        for nv in ds[:10]:
            nd += f"<li>{nv.get('ho_ten','')} - {nv.get('chuc_danh_nghe','')}</li>"
        if len(ds) > 10:
            nd += f"<li>... và {len(ds)-10} người khác</li>"
        nd += "</ul><p><b>File Excel đính kèm.</b></p>"
        msg.attach(MIMEText(nd, 'html', 'utf-8'))
        if file and os.path.exists(file):
            with open(file, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file)}"')
                msg.attach(part)
        srv = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
        srv.starttls()
        srv.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
        srv.send_message(msg)
        srv.quit()
        return True
    except Exception as e:
        st.error(f"Lỗi email: {e}")
        return False

def gui_telegram(msg):
    # from config import TELEGRAM_CONFIG as TC  <-- XÓA DÒNG NÀY
    
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_CONFIG['bot_token']}/sendMessage"
        r = requests.post(url, data={"chat_id": TELEGRAM_CONFIG['chat_id'], "text": msg, "parse_mode": "HTML"}, timeout=10)
        return r.status_code == 200
    except:
        return False

def tao_bao_cao_tang_giam(tang_list, giam_list, tu_ngay, den_ngay):
    """Tạo báo cáo Word tăng/giảm nhân sự"""
    from docx import Document
    from docx.shared import Pt, Cm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    
    doc = Document()
    
    # Tiêu đề
    title = doc.add_heading('BÁO CÁO TĂNG/GIẢM NHÂN SỰ', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    period = doc.add_paragraph(f'Thời gian: Từ {tu_ngay.strftime("%d/%m/%Y")} đến {den_ngay.strftime("%d/%m/%Y")}')
    period.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    
    # Danh sách tăng
    doc.add_heading('I. LAO ĐỘNG TĂNG MỚI', level=1)
    if tang_list:
        table = doc.add_table(rows=1 + len(tang_list), cols=5)
        table.style = 'Table Grid'
        # Header
        headers = ['STT', 'Họ tên', 'Chức danh', 'Loại HĐ', 'Ngày vào làm']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, nv in enumerate(tang_list, 1):
            row = table.rows[idx]
            row.cells[0].text = str(idx)
            row.cells[1].text = nv.get('ho_ten', '')
            row.cells[2].text = nv.get('chuc_danh_nghe', '')
            row.cells[3].text = nv.get('loai_hop_dong', '')
            row.cells[4].text = format_date(nv.get('ngay_vao_lam'))
    else:
        doc.add_paragraph('Không có lao động tăng trong kỳ.')
    
    doc.add_paragraph()
    
    # Danh sách giảm
    doc.add_heading('II. LAO ĐỘNG GIẢM (NGHỈ VIỆC)', level=1)
    if giam_list:
        table = doc.add_table(rows=1 + len(giam_list), cols=5)
        table.style = 'Table Grid'
        headers = ['STT', 'Họ tên', 'Chức danh', 'Loại HĐ', 'Ngày nghỉ việc']
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        
        for idx, nv in enumerate(giam_list, 1):
            row = table.rows[idx]
            row.cells[0].text = str(idx)
            row.cells[1].text = nv.get('ho_ten', '')
            row.cells[2].text = nv.get('chuc_danh_nghe', '')
            row.cells[3].text = nv.get('loai_hop_dong', '')
            row.cells[4].text = format_date(nv.get('ngay_ket_thuc'))
    else:
        doc.add_paragraph('Không có lao động giảm trong kỳ.')
    
    # Footer
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(f'Ngày {date.today().day} tháng {date.today().month} năm {date.today().year}')
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run('NGƯỜI LẬP BÁO CÁO')
    
    tf = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(tf.name)
    return tf.name
    
def show_super_admin_page():
    """Trang QUẢN TRỊ HỆ THỐNG — chỉ đội vận hành App dùng để thêm/sửa/khoá khách hàng (tenant).
    Hoàn toàn tách biệt với dữ liệu nhân sự của từng khách hàng."""
    st.title("⚙️ Quản trị hệ thống — Danh sách khách hàng (Tenants)")
    if st.button("🚪 Thoát trang quản trị"):
        st.session_state.super_admin_mode = False
        st.rerun()
    st.divider()

    # Hộp hướng dẫn bước THỦ CÔNG còn lại sau khi thêm 1 tenant mới (tạo Streamlit app riêng).
    # Lưu vào session_state để hiển thị BỀN sau st.rerun() (nếu không sẽ mất ngay lập tức).
    _vua_tao = st.session_state.get('_tenant_vua_tao')
    if _vua_tao:
        st.success(f"✅ Đã thêm khách hàng **{_vua_tao['ten_cty']}** (mã: **{_vua_tao['ma_cty']}**) "
                   f"và tự động chạy migration `schema.sql` thành công.")
        with st.container(border=True):
            st.markdown("### 📌 Bước tiếp theo (BẮT BUỘC — thực hiện thủ công trên Streamlit Cloud)")
            st.markdown(f"""
Mỗi khách hàng cần **1 app Streamlit Cloud riêng** để vào thẳng màn hình đăng nhập
(không cần chọn công ty). Thực hiện theo đúng thứ tự:

1. Vào [share.streamlit.io](https://share.streamlit.io) → **"New app"**
2. Chọn đúng repo GitHub hiện tại (repo chứa `app.py` này), nhánh `main`, file chính `app.py`
3. Đặt tên app theo **đúng quy chuẩn**: `hrm-{_vua_tao['ma_cty'].lower()}`
   → URL sẽ là: `https://hrm-{_vua_tao['ma_cty'].lower()}.streamlit.app`
4. Trước khi bấm Deploy, vào **"Advanced settings" → "Secrets"**, dán y hệt nội dung Secrets
   của app hiện tại, rồi **thêm thêm 1 dòng mới** vào cuối:
   ```
   tenant_code = "{_vua_tao['ma_cty']}"
   ```
   (Dòng này giúp app tự nhận diện đúng công ty, khách vào thẳng màn hình đăng nhập,
   không cần gõ mã công ty.)
5. Bấm **Deploy** và gửi link `https://hrm-{_vua_tao['ma_cty'].lower()}.streamlit.app` cho khách hàng
6. Tạo sẵn 1 dòng nhân viên đầu tiên (admin) trong bảng `nhan_vien` của DB khách hàng này —
   họ sẽ đăng nhập lần đầu bằng chính số điện thoại (xem hướng dẫn ở màn hình "Đổi mật khẩu lần đầu")
""")
            if st.button("✅ Đã tạo app xong, đóng thông báo này"):
                del st.session_state['_tenant_vua_tao']
                st.rerun()

    with st.expander("➕ Thêm khách hàng mới (SaaS)", expanded=False):
        with st.form("add_tenant_form"):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("##### 🏢 Thông tin Kết nối & Hệ thống")
                ma_cty = st.text_input("Mã công ty * (VD: CHL)")
                ten_cty = st.text_input("Tên công ty *")
                logo_url = st.text_input("Link logo (tuỳ chọn)")
                db_host = st.text_input("Supabase DB Host *")
                db_port = st.text_input("Supabase DB Port", value="5432")
                db_user = st.text_input("Supabase DB User", value="postgres")
                db_password = st.text_input("Supabase DB Password *", type="password")
                db_name = st.text_input("Supabase DB Name", value="postgres")
                supabase_url = st.text_input("Supabase Project URL *")
                supabase_key = st.text_input("Supabase API Key *", type="password")
            with col2:
                st.markdown("##### 🎨 Cấu hình Thương hiệu & Metadata")
                dai_dien = st.text_input("Người đại diện (Ký hợp đồng)", placeholder="VD: Nguyễn Đình Thi")
                chuc_vu = st.text_input("Chức vụ người ký", placeholder="VD: Tổng Giám Đốc")
                ma_so_thue = st.text_input("Mã số thuế")
                dien_thoai_cty = st.text_input("Điện thoại công ty")
                ma_don_vi_BHXH = st.text_input("Mã đơn vị BHXH")
                ma_vung_luong = st.text_input("Mã vùng lương")
                dia_chi = st.text_input("Địa chỉ công ty")
                loi_nhan_zalo = st.text_input("Lời nhắn Zalo sinh nhật")
                zalo_group_link = st.text_input("Link nhóm Zalo")
                zalo_group_name = st.text_input("Tên nhóm Zalo")
            
            if st.form_submit_button("💾 Lưu khách hàng & Tự động chạy Migration"):
                if not all([ma_cty, ten_cty, db_host, db_password, supabase_url, supabase_key]):
                    st.error("❌ Vui lòng điền đầy đủ các trường bắt buộc (*)")
                else:
                    try:
                        # Đọc file schema.sql từ thư mục hiện tại
                        migration_sql = None
                        import os
                        schema_path = os.path.join(os.path.dirname(__file__), "schema.sql")
                        if os.path.exists(schema_path):
                            with open(schema_path, "r", encoding="utf-8") as sf:
                                migration_sql = sf.read()
                        
                        control_plane.add_tenant(
                            ma_cty=ma_cty, ten_cty=ten_cty, db_host=db_host, db_port=db_port,
                            db_user=db_user, db_password=db_password, db_name=db_name,
                            supabase_url=supabase_url, supabase_key=supabase_key, logo_url=logo_url,
                            dai_dien=dai_dien, chuc_vu=chuc_vu, ma_so_thue=ma_so_thue,
                            dien_thoai_cty=dien_thoai_cty, ma_don_vi_BHXH=ma_don_vi_BHXH,
                            ma_vung_luong=ma_vung_luong, dia_chi=dia_chi, loi_nhan_zalo=loi_nhan_zalo,
                            zalo_group_link=zalo_group_link, zalo_group_name=zalo_group_name,
                            migration_sql=migration_sql
                        )
                        st.session_state['_tenant_vua_tao'] = {'ma_cty': ma_cty.strip().upper(), 'ten_cty': ten_cty}
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Lỗi khi thêm khách hàng hoặc chạy migration: {e}")

    st.subheader("📋 Danh sách khách hàng hiện có")
    try:
        tenants = control_plane.list_tenants()
    except Exception as e:
        tenants = []
        st.error(f"❌ Không kết nối được Control Plane. Kiểm tra lại st.secrets['control_plane']. Chi tiết: {e}")

    if tenants:
        df = pd.DataFrame(tenants)
        st.dataframe(df, width='stretch', hide_index=True)
        st.divider()
        col_a, col_b = st.columns(2)
        with col_a:
            ma_toggle = st.text_input("Mã công ty cần Khoá/Mở khoá")
            trang_thai_moi = st.selectbox("Trạng thái mới", ["active", "suspended"])
            if st.button("🔄 Cập nhật trạng thái"):
                if ma_toggle:
                    control_plane.update_tenant_status(ma_toggle, trang_thai_moi)
                    st.success("✅ Đã cập nhật!"); st.rerun()
        with col_b:
            ma_xoa = st.text_input("Mã công ty cần XOÁ vĩnh viễn khỏi hệ thống")
            if st.button("🗑️ Xoá khách hàng", type="primary"):
                if ma_xoa:
                    control_plane.delete_tenant(ma_xoa)
                    st.success("✅ Đã xoá!"); st.rerun()

        st.divider()
        st.markdown("##### 🖼️ Upload logo cho khách hàng")
        col_logo1, col_logo2 = st.columns([1, 2])
        with col_logo1:
            ma_cty_logo = st.text_input("Mã công ty", key="ma_cty_upload_logo")
        with col_logo2:
            logo_file = st.file_uploader("Chọn file logo (PNG/JPG)", type=["png", "jpg", "jpeg"], key="logo_file_uploader")
        if st.button("📤 Upload logo", key="btn_upload_logo"):
            if not ma_cty_logo or not logo_file:
                st.warning("⚠️ Vui lòng nhập Mã công ty và chọn file logo.")
            else:
                sb = get_supabase_storage()
                if not sb:
                    st.error("❌ Chưa cấu hình Supabase Storage.")
                else:
                    try:
                        safe_name = sanitize_storage_filename(logo_file.name)
                        storage_path = f"logos/{sanitize_storage_filename(ma_cty_logo)}/{safe_name}"
                        upload_to_storage_unique(
                            sb, SUPABASE_BUCKET, storage_path,
                            logo_file.getvalue(), logo_file.type
                        )
                        # Lấy public URL để lưu vào tenant.logo_url (bucket/đường dẫn logo cần để PUBLIC
                        # vì logo hiển thị cả ở màn hình đăng nhập, trước khi xác thực người dùng)
                        public_url = sb.storage.from_(SUPABASE_BUCKET).get_public_url(storage_path)
                        control_plane.update_tenant_logo(ma_cty_logo.strip().upper(), public_url)
                        st.success(f"✅ Đã upload logo và cập nhật cho công ty {ma_cty_logo.strip().upper()}. Link: {public_url}")
                        st.image(public_url, width=160)
                        st.caption("ℹ️ Lưu ý: cần bật chế độ **Public** cho bucket/đường dẫn "
                                   f"`{SUPABASE_BUCKET}` (hoặc riêng thư mục `logos/`) trên Supabase Dashboard, "
                                   "nếu không link này sẽ không tải được vì bucket hồ sơ nhân viên mặc định là riêng tư.")
                        st.rerun()
                    except AttributeError:
                        st.error("❌ Chưa có hàm `update_tenant_logo()` trong control_plane.py. "
                                 "Cần thêm hàm này (UPDATE tenants SET logo_url=%s WHERE ma_cty=%s) để nút này hoạt động.")
                    except Exception as e:
                        st.error(f"❌ Lỗi upload logo: {e}")
    else:
        st.info("Chưa có khách hàng nào. Thêm khách hàng đầu tiên ở form phía trên.")


# ========== SIDEBAR + LOGIN (ĐA KHÁCH HÀNG) ==========
if not st.session_state.get('tenant'):
    st.sidebar.title("🏗️ HRM Master")
    st.sidebar.caption("Nền tảng Quản lý Hồ sơ nhân sự đa doanh nghiệp")


def check_login(username, password):
    """Xác thực đăng nhập của NHÂN VIÊN thuộc tenant (công ty) đã chọn.
    Tài khoản = số điện thoại (dien_thoai), mật khẩu hash bằng bcrypt trong cột mat_khau_hash.
    Trả về (success, role, nhan_vien_row) — nhan_vien_row là dict thông tin NV nếu thành công."""
    tenant = st.session_state.get('tenant')

    if tenant:
        try:
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            # Tài khoản đăng nhập = ten_dang_nhap nếu admin đã set riêng, mặc định = số điện thoại.
            # COALESCE(ten_dang_nhap, dien_thoai): nếu ten_dang_nhap NULL/rỗng thì so sánh theo SĐT.
            c.execute("""SELECT id, ho_ten, dien_thoai, ten_dang_nhap, mat_khau_hash, vai_tro, phai_doi_mat_khau
                         FROM nhan_vien
                         WHERE NULLIF(TRIM(ten_dang_nhap), '') = %s
                            OR (NULLIF(TRIM(ten_dang_nhap), '') IS NULL AND dien_thoai = %s)""",
                      (username.strip(), username.strip()))
            rows = c.fetchall()
            db.close()
            row = rows[0] if rows else None
            if not row:
                # Không khớp nhân viên nào trong DB — thử tài khoản khai báo sẵn trong
                # Secrets [users] (dành cho đội vận hành/thử nghiệm theo vai trò,
                # KHÔNG gắn với 1 nhân viên cụ thể nên không có nhan_vien_id).
                try:
                    if 'users' in st.secrets and username in st.secrets.users:
                        if st.secrets.users[username]['password'] == password:
                            return True, st.secrets.users[username]['role'], None
                except Exception:
                    pass
                return False, None, None
            if not row.get('mat_khau_hash'):
                khop = password.strip() == (row.get('dien_thoai') or '').strip()
                if khop:
                    row['phai_doi_mat_khau'] = True
                    return True, row.get('vai_tro') or 'nhan_vien', row
                return False, None, None
            if bcrypt.checkpw(password.encode(), row['mat_khau_hash'].encode()):
                return True, row.get('vai_tro') or 'nhan_vien', row
        except Exception:
            pass
        return False, None, None

    # ---- Chế độ KHÔNG có tenant (chạy đơn lẻ / dev local) — giữ cách cũ để không phá vỡ ----
    try:
        if 'users' in st.secrets and username in st.secrets.users:
            if st.secrets.users[username]['password'] == password:
                return True, st.secrets.users[username]['role'], None
    except Exception:
        pass
    try:
        if username in USERS:
            return USERS[username]['password'] == password, USERS[username]['role'], None
    except Exception:
        pass
    return False, None, None


if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None

if not st.session_state.logged_in:

    # ---------- Trang quản trị hệ thống (super-admin quản lý danh sách khách hàng) ----------
    if st.session_state.get('super_admin_mode'):
        show_super_admin_page()
        st.stop()

    # ---------- App bị khoá cứng vào 1 tenant (Secret tenant_code) nhưng mã sai/bị khoá ----------
    if st.session_state.get('_tenant_locked_error'):
        st.error("⚠️ App này được cấu hình riêng cho 1 khách hàng, nhưng không tìm thấy "
                 "hoặc tài khoản khách hàng đang bị tạm khoá. Vui lòng liên hệ đơn vị triển khai App.")
        st.stop()

    # ---------- BƯỚC 1: Không xác định được tenant qua domain/subdomain ----------
    # Mỗi khách hàng nay có domain/subdomain riêng, được resolve_tenant() tự nhận diện.
    # Nếu chạy đến đây nghĩa là đang ở domain gốc (chưa gán cho khách nào) hoặc
    # subdomain không khớp tenant nào — KHÔNG còn cho phép tự nhập mã công ty nữa.
    # Domain gốc chỉ còn dùng làm cổng vào cho Super Admin (đội vận hành App).
    if not st.session_state.get('tenant'):
        st.title("🏗️ HRM-Port")
        st.info(
            "🔒 Tên miền này chưa được gán cho khách hàng nào. "
            "Nếu bạn là nhân viên của một công ty đang dùng HRM-Port, vui lòng truy cập "
            "đúng địa chỉ riêng của công ty bạn (ví dụ: `hangcuaban.kendu-ai.com`)."
        )
        with st.expander("⚙️ Quản trị hệ thống (chỉ dành cho đội vận hành App)"):
            sa_u = st.text_input("Tài khoản", key="sa_user")
            sa_p = st.text_input("Mật khẩu", type="password", key="sa_pass")
            if st.button("Đăng nhập quản trị", key="sa_login"):
                if control_plane.check_super_admin(sa_u, sa_p):
                    st.session_state.super_admin_mode = True
                    st.rerun()
                else:
                    st.error("❌ Sai tài khoản/mật khẩu quản trị hệ thống!")
        st.stop()

    # ---------- BƯỚC 2: Đăng nhập nhân viên của công ty đã chọn ----------
    tenant = st.session_state.tenant
    if tenant.get('logo_url'):
        st.sidebar.image(tenant['logo_url'], width='stretch')
    st.sidebar.success(f"🏢 **{tenant['ten_cty']}**")

    st.sidebar.subheader("🔐 Đăng nhập")
    u = st.sidebar.text_input("Số điện thoại hoặc Tên đăng nhập")
    p = st.sidebar.text_input("Mật khẩu", type="password")
    st.sidebar.caption("💡 Mật khẩu mặc định = số điện thoại của bạn. Đổi lại sau khi đăng nhập lần đầu.")
    if st.sidebar.button("Đăng nhập", width='stretch'):
        success, role, nv_row = check_login(u, p)
        if success:
            st.session_state.logged_in = True
            st.session_state.role = role
            st.session_state.username = u
            st.session_state.nhan_vien_id = nv_row['id'] if nv_row else None
            st.session_state.ho_ten_dang_nhap = nv_row['ho_ten'] if nv_row else u
            st.session_state.phai_doi_mat_khau = bool(nv_row and nv_row.get('phai_doi_mat_khau'))
            st.rerun()
        else:
            st.sidebar.error("❌ Sai tài khoản hoặc mật khẩu!")

    with st.sidebar.expander("🔑 Quên mật khẩu?"):
        try:
            db_qmk = st.session_state.db_engine.get_connection()
            c_qmk = db_qmk.cursor()
            c_qmk.execute("""
                CREATE TABLE IF NOT EXISTS yeu_cau_reset_mk (
                    id SERIAL PRIMARY KEY,
                    nhan_vien_id INT NOT NULL,
                    otp_code VARCHAR(10) NOT NULL,
                    het_han TIMESTAMP NOT NULL,
                    da_dung BOOLEAN DEFAULT FALSE,
                    created_at TIMESTAMP DEFAULT NOW()
                )
            """)
            db_qmk.commit(); db_qmk.close()
        except Exception:
            pass

        buoc_qmk = st.session_state.get('qmk_buoc', 1)
        if buoc_qmk == 1:
            st.caption("Nhập SĐT hoặc Tên đăng nhập — mã xác nhận (OTP) sẽ gửi về Email liên hệ đã đăng ký.")
            tk_qmk = st.text_input("SĐT / Tên đăng nhập:", key="qmk_tk")
            if st.button("📧 Gửi mã OTP", key="qmk_gui_otp"):
                db_q = st.session_state.db_engine.get_connection()
                c_q = db_q.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                c_q.execute("""SELECT id, ho_ten, email_lien_he FROM nhan_vien
                               WHERE dien_thoai=%s OR ten_dang_nhap=%s""", (tk_qmk, tk_qmk))
                nv_qmk = c_q.fetchone()
                db_q.close()
                if not nv_qmk:
                    st.error("❌ Không tìm thấy tài khoản này.")
                elif not nv_qmk.get('email_lien_he'):
                    st.error("❌ Tài khoản chưa có Email liên hệ. Liên hệ Admin/HR để được hỗ trợ đặt lại mật khẩu.")
                else:
                    otp = f"{random.randint(0, 999999):06d}"
                    het_han = datetime.now() + timedelta(minutes=10)
                    db_o = st.session_state.db_engine.get_connection()
                    c_o = db_o.cursor()
                    c_o.execute("""INSERT INTO yeu_cau_reset_mk (nhan_vien_id, otp_code, het_han)
                                   VALUES (%s, %s, %s)""", (nv_qmk['id'], otp, het_han))
                    db_o.commit(); db_o.close()
                    da_gui = gui_email_don(
                        nv_qmk['email_lien_he'],
                        "🔑 Mã xác nhận đặt lại mật khẩu - HRM",
                        f"<p>Xin chào {nv_qmk['ho_ten']},</p><p>Mã xác nhận (OTP) của bạn là:</p>"
                        f"<h2 style='letter-spacing:4px;'>{otp}</h2>"
                        f"<p>Mã có hiệu lực trong 10 phút. Nếu không phải bạn yêu cầu, vui lòng bỏ qua email này.</p>"
                    )
                    if da_gui:
                        st.session_state['qmk_buoc'] = 2
                        st.session_state['qmk_nv_id'] = nv_qmk['id']
                        st.success(f"✅ Đã gửi mã OTP về {nv_qmk['email_lien_he'][:3]}***@...")
                        st.rerun()
                    else:
                        st.error("❌ Gửi email thất bại. Vui lòng thử lại hoặc liên hệ Admin.")
        elif buoc_qmk == 2:
            st.caption("Nhập mã OTP đã gửi về Email và đặt mật khẩu mới.")
            otp_nhap = st.text_input("Mã OTP:", key="qmk_otp_nhap")
            mk_moi_qmk = st.text_input("Mật khẩu mới:", type="password", key="qmk_mk_moi")
            mk_moi_qmk2 = st.text_input("Nhập lại mật khẩu mới:", type="password", key="qmk_mk_moi2")
            col_qmk1, col_qmk2 = st.columns(2)
            with col_qmk1:
                if st.button("✅ Xác nhận đặt lại", key="qmk_xac_nhan"):
                    if len(mk_moi_qmk) < 6:
                        st.error("Mật khẩu mới phải có ít nhất 6 ký tự.")
                    elif mk_moi_qmk != mk_moi_qmk2:
                        st.error("Hai mật khẩu nhập lại không khớp.")
                    else:
                        db_v = st.session_state.db_engine.get_connection()
                        c_v = db_v.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                        c_v.execute("""
                            SELECT id FROM yeu_cau_reset_mk
                            WHERE nhan_vien_id=%s AND otp_code=%s AND da_dung=FALSE AND het_han > NOW()
                            ORDER BY id DESC LIMIT 1
                        """, (st.session_state['qmk_nv_id'], otp_nhap))
                        yc = c_v.fetchone()
                        if not yc:
                            db_v.close()
                            st.error("❌ Mã OTP không đúng hoặc đã hết hạn.")
                        else:
                            c_v2 = db_v.cursor()
                            new_hash = bcrypt.hashpw(mk_moi_qmk.encode(), bcrypt.gensalt()).decode()
                            c_v2.execute("UPDATE nhan_vien SET mat_khau_hash=%s, phai_doi_mat_khau=FALSE WHERE id=%s",
                                         (new_hash, st.session_state['qmk_nv_id']))
                            c_v2.execute("UPDATE yeu_cau_reset_mk SET da_dung=TRUE WHERE id=%s", (yc['id'],))
                            db_v.commit(); db_v.close()
                            st.success("✅ Đặt lại mật khẩu thành công! Vui lòng đăng nhập lại.")
                            st.session_state.pop('qmk_buoc', None)
                            st.session_state.pop('qmk_nv_id', None)
                            st.rerun()
            with col_qmk2:
                if st.button("✖️ Hủy", key="qmk_huy"):
                    st.session_state.pop('qmk_buoc', None)
                    st.session_state.pop('qmk_nv_id', None)
                    st.rerun()
    st.stop()

# ---------- Bắt buộc đổi mật khẩu lần đầu (đang dùng mật khẩu mặc định = SĐT) ----------
if st.session_state.get('phai_doi_mat_khau'):
    st.title("🔑 Đổi mật khẩu lần đầu")
    st.warning("Đây là lần đăng nhập đầu tiên (mật khẩu mặc định = số điện thoại của bạn). "
               "Vui lòng đặt mật khẩu mới trước khi tiếp tục sử dụng hệ thống.")
    mk_moi = st.text_input("Mật khẩu mới", type="password", key="mk_moi_lan_dau")
    mk_moi2 = st.text_input("Nhập lại mật khẩu mới", type="password", key="mk_moi_lan_dau_2")
    if st.button("✅ Xác nhận đổi mật khẩu"):
        if len(mk_moi) < 6:
            st.error("Mật khẩu mới phải có ít nhất 6 ký tự.")
        elif mk_moi != mk_moi2:
            st.error("Hai mật khẩu nhập lại không khớp.")
        else:
            try:
                db = st.session_state.db_engine.get_connection()
                c = db.cursor()
                new_hash = bcrypt.hashpw(mk_moi.encode(), bcrypt.gensalt()).decode()
                c.execute(
                    "UPDATE nhan_vien SET mat_khau_hash=%s, phai_doi_mat_khau=FALSE WHERE id=%s",
                    (new_hash, st.session_state.nhan_vien_id)
                )
                db.commit()
                db.close()
                st.session_state.phai_doi_mat_khau = False
                st.success("✅ Đổi mật khẩu thành công! Đang vào hệ thống...")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Lỗi khi đổi mật khẩu: {e}")
    st.stop()

# Menu theo role — 4 vai trò cố định: admin / hr / kt_luong / viewer (+ 'nhan_vien' tự phục vụ)
if st.session_state.role == "admin":
    # Toàn quyền
    menu_options = ["📊 Dashboard","👤 Ứng viên","✅ Nhân viên","📁 Upload hồ sơ","⚙️ Danh mục","📋 BHXH","📋 Báo cáo định kỳ","🕒 Chấm công","💰 Tính thu nhập","📄 Quản lý Công văn & HĐ kinh tế","💬 Chat nội bộ","🤖 Chatbot Giải đáp","🔑 Quản lý MK","🖼️ Tạo ảnh thẻ NV","🔍 Audit Dashboard","📘 Hướng dẫn sử dụng",]
elif st.session_state.role in ["văn thư", "hr"]:
    # HR: như admin trừ Upload hồ sơ, Danh mục — và KHÔNG được xem Tính thu nhập (dữ liệu lương)
    menu_options = ["📊 Dashboard","✅ Nhân viên","📋 BHXH","📋 Báo cáo định kỳ","🕒 Chấm công","📄 Quản lý Công văn & HĐ kinh tế","💬 Chat nội bộ","🤖 Chatbot Giải đáp","🔑 Quản lý MK","🖼️ Tạo ảnh thẻ NV","📘 Hướng dẫn sử dụng",]
elif st.session_state.role == "kt_luong":
    # Kế toán lương: tập trung vào Chấm công + Tính thu nhập, không có Upload hồ sơ/Danh mục
    menu_options = ["📊 Dashboard","✅ Nhân viên","📋 BHXH","🕒 Chấm công","💰 Tính thu nhập","💬 Chat nội bộ","🤖 Chatbot Giải đáp","🔑 Quản lý MK","🖼️ Tạo ảnh thẻ NV","📘 Hướng dẫn sử dụng",]
elif st.session_state.role == "van_thu":
    menu_options = ["📊 Dashboard","✅ Nhân viên","🕒 Chấm công","📄 Quản lý Công văn & HĐ kinh tế","💬 Chat nội bộ","🤖 Chatbot Giải đáp","🔑 Quản lý MK","🖼️ Tạo ảnh thẻ NV","📘 Hướng dẫn sử dụng",]
elif st.session_state.role == "viewer":
    # Viewer: chỉ xem, thu hẹp — không có BHXH, không có Tính thu nhập
    menu_options = ["📊 Dashboard","✅ Nhân viên","📋 Báo cáo định kỳ","🕒 Chấm công","💬 Chat nội bộ","🤖 Chatbot Giải đáp","🔑 Quản lý MK","🖼️ Tạo ảnh thẻ NV","📘 Hướng dẫn sử dụng",]
else:  # 'nhan_vien' thường — chỉ xem hồ sơ bản thân + chat nội bộ
    menu_options = ["📊 Dashboard","✅ Nhân viên","🕒 Chấm công","💬 Chat nội bộ","🤖 Chatbot Giải đáp","🔑 Quản lý MK","🖼️ Tạo ảnh thẻ NV","📘 Hướng dẫn sử dụng"]
menu = st.sidebar.radio("📋 Menu", menu_options)
st.sidebar.divider()
st.sidebar.caption(f"👤 {st.session_state.get('ho_ten_dang_nhap', st.session_state.username)} ({st.session_state.role})")
# MỚI:
if st.sidebar.button("🚪 Đăng xuất", width='stretch'):
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None
    st.session_state.show_hrm = False
    st.session_state.pop('last_birthday_check', None)
    st.session_state.pop('sinh_nhat_hom_nay_list', None)
    st.cache_data.clear()
    st.rerun()

# ========== HÀM DÙNG CHUNG: CARD THÔNG TIN NHÂN VIÊN ==========
PHONG_BAN_LANH_DAO_CAO_CAP = ('Hội đồng Quản trị', 'Ban Tổng Giám đốc')

def render_employee_info_card(nv, key_prefix, on_close=None):
    """Hiển thị card '👤 THÔNG TIN NHÂN VIÊN' với avatar load on-demand"""
    st.subheader("👤 THÔNG TIN NHÂN SỰ")

    col_avatar, col_info = st.columns([1, 2])

    with col_avatar:
        st.markdown("""
        <style>
        .avatar-wrapper {
            display: flex;
            align-items: center;
            justify-content: center;
            height: 100%;
            min-height: 250px;
        }
        .avatar-img {
            width: 200px;
            height: 200px;
            border-radius: 50%;
            object-fit: cover;
            border: 4px solid #f59e0b;
            box-shadow: 0 4px 15px rgba(0,0,0,0.15);
        }
        </style>
        """, unsafe_allow_html=True)

        # ===== CẢI TIẾN: Chỉ tải ảnh khi có hành động =====
        anh_path = nv.get('anh_ho_so')
        avatar_key = f"avatar_loaded_{nv['id']}"
        
        if anh_path:
            # Dùng cache để tải ảnh (chỉ tải 1 lần, cache 1 giờ)
            anh_bytes = get_avatar_bytes_cached(anh_path)
            if anh_bytes:
                img_base64 = base64.b64encode(anh_bytes).decode()
                st.markdown(f"""
                <div class="avatar-wrapper">
                    <img src="data:image/jpeg;base64,{img_base64}" class="avatar-img" loading="lazy">
                </div>
                """, unsafe_allow_html=True)
            else:
                # Fallback: ảnh mặc định
                gioi_tinh = nv.get('gioi_tinh', '')
                ho_ten = nv.get('ho_ten', '')
                avatar_file = "avatar_male.png" if gioi_tinh == "Nam" else "avatar_female.png"
                avatar_path = os.path.join(os.path.dirname(__file__), "static", avatar_file)
                if os.path.exists(avatar_path):
                    with open(avatar_path, "rb") as f:
                        img_data = base64.b64encode(f.read()).decode()
                    st.markdown(f"""
                    <div class="avatar-wrapper">
                        <img src="data:image/png;base64,{img_data}" class="avatar-img">
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="avatar-wrapper">
                        <img src="https://ui-avatars.com/api/?name={ho_ten.replace(' ', '+')}&size=200&background=f59e0b&color=fff" class="avatar-img">
                    </div>
                    """, unsafe_allow_html=True)
        else:
            # Fallback khi không có ảnh
            gioi_tinh = nv.get('gioi_tinh', '')
            ho_ten = nv.get('ho_ten', '')
            avatar_file = "avatar_male.png" if gioi_tinh == "Nam" else "avatar_female.png"
            avatar_path = os.path.join(os.path.dirname(__file__), "static", avatar_file)
            if os.path.exists(avatar_path):
                with open(avatar_path, "rb") as f:
                    img_data = base64.b64encode(f.read()).decode()
                st.markdown(f"""
                <div class="avatar-wrapper">
                    <img src="data:image/png;base64,{img_data}" class="avatar-img">
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="avatar-wrapper">
                    <img src="https://ui-avatars.com/api/?name={ho_ten.replace(' ', '+')}&size=200&background=f59e0b&color=fff" class="avatar-img">
                </div>
                """, unsafe_allow_html=True)

    with col_info:
        # Xác định xem có phải lãnh đạo cấp cao không
        la_lanh_dao_cc = la_phong_ban_lanh_dao_cao_cap(nv.get('phong_ban_lam_viec'))
        
        if la_lanh_dao_cc:
            # Lãnh đạo cấp cao: chỉ hiển thị Ông/Bà + Họ tên, KHÔNG có ma_nv
            xung_ho_card = get_xung_ho_trang_trong(nv.get('gioi_tinh'))
            st.markdown(f"### {xung_ho_card} {nv['ho_ten']}".strip())
        else:
            # Nhân viên thường: hiển thị Họ tên (Mã NV)
            st.markdown(f"### {nv['ho_ten']} ({nv['ma_nv']})")

        info_col1, info_col2 = st.columns(2)

        with info_col1:
            st.markdown(f"**📅 Ngày sinh:** {format_date(nv.get('ngay_sinh'))}")
            st.markdown(f"**⚧ Giới tính:** {nv.get('gioi_tinh', 'Chưa cập nhật')}")
            if nv.get('chuc_danh_nghe'):
                st.markdown(f"**💼 Chức danh:** {nv.get('chuc_danh_nghe')}")
            st.markdown(f"**🏢 Phòng:** {nv.get('phong_ban_lam_viec', 'Chưa cập nhật')}")
            # Chỉ hiển thị chức vụ nếu có
            if nv.get('chuc_vu'):
                st.markdown(f"**🎖️ Chức vụ:** {nv.get('chuc_vu')}")
            st.markdown(f"**📞 SĐT:** {nv.get('dien_thoai', 'Chưa cập nhật')}")

        with info_col2:
            if nv.get('so_hdld'):
                st.markdown(f"**Số Hợp đồng:** {nv.get('so_hdld')}")
            # Chỉ hiển thị Loại HĐ cho nhân viên thường (không phải HĐQT/BTGĐ)
            if not la_lanh_dao_cc:
                st.markdown(f"**📋 Loại HĐ:** {nv.get('loai_hop_dong', 'Chưa cập nhật')}")
            if nv.get('ngay_vao_lam'):
                st.markdown(f"**📅 Ngày vào làm:** {format_date(nv.get('ngay_vao_lam'))}")
            st.markdown(f"**🎓 Trình độ:** {nv.get('trinh_do', 'Chưa cập nhật')}")
            st.markdown(f"**📇 Mã BHXH:** {nv.get('ma_so_bhxh', 'Chưa có')}")
            # Chỉ hiển thị trạng thái cho nhân viên thường
            if not la_lanh_dao_cc:
                trang_thai_text = {
                    'DANG_LAM': '🟢 Đang làm',
                    'THU_VIEC': '🔵 Thử việc',
                    'NGHI_VIEC': '🔴 Đã nghỉ'
                }
                status = trang_thai_text.get(nv.get('trang_thai'), nv.get('trang_thai', 'Chưa xác định'))
                st.markdown(f"**📊 Trạng thái:** {status}")
            # Nếu là lãnh đạo cấp cao, hiển thị chức vụ thay vì trạng thái
            else:
                st.markdown(f"**🎖️ Chức vụ:** {nv.get('chuc_vu', 'Thành viên')}")

    # ===== Nút hành động (thêm nút "Đóng" ở cuối) =====
    st.divider()
    col_btn_action1, col_btn_action2, col_btn_action3, col_btn_action4, col_btn_action5 = st.columns(5)
    if st.session_state.role == "admin":
        with col_btn_action1:
            if st.button("✏️ SỬA NHÂN VIÊN", width='stretch', type="primary", key=f"edit_nv_btn_{key_prefix}"):
                st.session_state['selected_nv_id'] = int(nv['id'])
                st.rerun()

        with col_btn_action3:
            if nv.get('trang_thai') == 'DANG_LAM':
                if st.button("🖨️ IN HĐLĐ", width='stretch', key=f"print_hdld_card_{key_prefix}"):
                    db = st.session_state.db_engine.get_connection()
                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(nv['id']),))
                    nv_full = c.fetchone()
                    db.close()
                    if nv_full:
                        fp = tao_hop_dong(nv_full)
                        with open(fp, "rb") as f:
                            st.download_button(
                                label="📥 TẢI HĐLĐ",
                                data=f,
                                file_name=f"HDLD_{nv_full['ho_ten']}_{datetime.now().strftime('%Y%m%d')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"download_hdld_{key_prefix}"
                            )
            elif nv.get('trang_thai') == 'THU_VIEC':
                if st.button("🖨️ IN HĐTV", width='stretch', key=f"print_hdtv_card_{key_prefix}"):
                    db = st.session_state.db_engine.get_connection()
                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(nv['id']),))
                    nv_full = c.fetchone()
                    db.close()
                    if nv_full:
                        fp = tao_hop_dong_thu_viec(nv_full)
                        with open(fp, "rb") as f:
                            st.download_button(
                                label="📥 TẢI HĐTV",
                                data=f,
                                file_name=f"HDTV_{nv_full['ho_ten']}_{datetime.now().strftime('%Y%m%d')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"download_hdtv_{key_prefix}"
                            )
            else:
                st.button("📄 KHÔNG THỂ IN HĐ", disabled=True, width='stretch', key=f"no_hd_{key_prefix}")

        with col_btn_action5:
            if st.button("❌ Đóng", width='stretch', key=f"close_profile_{key_prefix}"):
                if callable(on_close):
                    on_close()
                st.rerun()
        
        with col_btn_action4:
            ma_bhxh = nv.get('ma_so_bhxh', '')
            chua_co_bhxh = not bool(ma_bhxh and str(ma_bhxh).strip())
            if chua_co_bhxh:
                if st.button("🏠 NHẬP T.TIN HỘ GĐ", width='stretch', type="primary", key=f"bhxh_family_{key_prefix}"):
                    st.session_state['bhxh_family_nv_id'] = int(nv['id'])
                    st.session_state['bhxh_family_nv_name'] = nv['ho_ten']
                    st.rerun()


# ========== DASHBOARD ==========
if menu == "📊 Dashboard":
    st.title("📊 Dashboard")
    
    # Lấy dữ liệu từ cache
    stats = get_dashboard_stats()
    
    # Hiển thị metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Tổng UV", stats['tong_uv'])
    col2.metric("Nhân viên", stats['tong_nv'])
    col3.metric("Chờ duyệt", stats['uv_by_status'].get('CHO_DUYET', 0))
    col4.metric("Đã nhận", stats['uv_by_status'].get('DA_NHAN_VIEC', 0))
    col5.metric("Từ chối", stats['uv_by_status'].get('TU_CHOI', 0))

    # ===== CẢNH BÁO 1: Chuẩn bị 6 Báo cáo định kỳ (30/06-05/07 và 31/12-05/01) =====
    hom_nay = date.today()
    md_hien_tai = (hom_nay.month, hom_nay.day)
    trong_khoang_giua_nam = (6, 30) <= md_hien_tai <= (7, 5)
    trong_khoang_cuoi_nam = md_hien_tai >= (12, 31) or md_hien_tai <= (1, 5)
    if trong_khoang_giua_nam or trong_khoang_cuoi_nam:
        han_nop_bc = "05/7" if trong_khoang_giua_nam else "05/01"
        st.warning(f"📋 **Chuẩn bị các 6 BC Định kỳ và nộp trước {han_nop_bc} HR nhé!**")

    # ===== CẢNH BÁO 2: Hạn nộp BC Tăng/Giảm BHXH hàng tháng (cấu hình riêng theo DN) =====
    han_ngay_bhxh = get_han_nop_bhxh()
    so_ngay_trong_thang = calendar.monthrange(hom_nay.year, hom_nay.month)[1]
    ngay_han_hop_le = min(han_ngay_bhxh, so_ngay_trong_thang)
    han_thang_nay = date(hom_nay.year, hom_nay.month, ngay_han_hop_le)
    so_ngay_con_lai_bhxh = (han_thang_nay - hom_nay).days
    if 0 <= so_ngay_con_lai_bhxh <= 5:
        thong_diep_bhxh = (f"Chuẩn bị nộp BC Tăng/Giảm BHXH tháng này nhé HR! "
                            f"(Hạn nộp: {han_thang_nay.strftime('%d/%m/%Y')}, còn {so_ngay_con_lai_bhxh} ngày)")
        if so_ngay_con_lai_bhxh <= 1:
            st.error(f"🚨 **KHẨN:** {thong_diep_bhxh}")
        elif so_ngay_con_lai_bhxh <= 3:
            st.warning(f"⚠️ {thong_diep_bhxh}")
        else:
            st.info(f"📌 {thong_diep_bhxh}")

    st.divider()
        
    
    
    # ── Phân bố chức danh ──
    import plotly.express as px
    import plotly.graph_objects as go

    db2 = st.session_state.db_engine.get_connection()
    c2 = db2.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c2.execute("""
        SELECT chuc_danh_nghe, COUNT(*) t 
        FROM nhan_vien 
        WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND so_hdld IS NOT NULL AND so_hdld != ''
        GROUP BY chuc_danh_nghe 
        ORDER BY t DESC
    """)
    data = c2.fetchall()
    db2.close()

    if data:
    # ========== PHẦN DASHBOARD NÂNG CAO ==========
        st.subheader("📊 TỔNG QUAN PHÂN BỐ NHÂN SỰ")

        db_dash = st.session_state.db_engine.get_connection()
        c_dash = db_dash.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        # Tiêu chuẩn thống kê DUY NHẤT cho toàn bộ Dashboard: chỉ tính nhân sự đang làm
        # (đang làm hoặc thử việc) VÀ đã có số hợp đồng lao động (so_hdld) — loại bỏ
        # nhân viên đã nghỉ việc và nhân viên chưa có so_hdld (hồ sơ chưa hoàn thiện)
        # để số liệu giữa các biểu đồ luôn khớp nhau.
        DK_CHUAN_NV = "trang_thai IN ('DANG_LAM', 'THU_VIEC') AND so_hdld IS NOT NULL AND so_hdld != ''"

        # 1. Dữ liệu cho Table "Trạng thái nhân sự các phòng ban"
        c_dash.execute(f"""
            SELECT 
                phong_ban_lam_viec as "Phòng ban",
                COUNT(*) as "Tổng số",
                SUM(CASE WHEN trang_thai = 'DANG_LAM' THEN 1 ELSE 0 END) as "Đang làm",
                SUM(CASE WHEN trang_thai = 'THU_VIEC' THEN 1 ELSE 0 END) as "Thử việc"
            FROM nhan_vien
            WHERE {DK_CHUAN_NV}
            GROUP BY phong_ban_lam_viec
            ORDER BY "Tổng số" DESC
        """)
        table_data = c_dash.fetchall()
        table_data = sap_xep_phong_ban_rows(table_data, "Phòng ban")
        
        # 2. Dữ liệu cho các biểu đồ
        # a. Tỷ lệ nhân sự mỗi phòng ban
        c_dash.execute(f"""
            SELECT phong_ban_lam_viec as "Phòng ban", COUNT(*) as "Số lượng"
            FROM nhan_vien WHERE {DK_CHUAN_NV}
            GROUP BY phong_ban_lam_viec
            ORDER BY "Số lượng" DESC
        """)
        dept_data = c_dash.fetchall()
        dept_data = sap_xep_phong_ban_rows(dept_data, "Phòng ban")

        # b. Cơ cấu theo giới tính
        c_dash.execute(f"""
            SELECT gioi_tinh, COUNT(*) as "Số lượng"
            FROM nhan_vien WHERE {DK_CHUAN_NV}
            GROUP BY gioi_tinh
        """)
        gender_data = c_dash.fetchall()

        # c. Cơ cấu theo Trình độ học vấn
        c_dash.execute(f"""
            SELECT trinh_do, COUNT(*) as "Số lượng"
            FROM nhan_vien WHERE {DK_CHUAN_NV}
            GROUP BY trinh_do
            ORDER BY "Số lượng" DESC
        """)
        education_data = c_dash.fetchall()

        # d. Cơ cấu theo Chức danh (Top 10) — LẤY TOÀN BỘ chức danh ở SQL (không LIMIT ở đây),
        # việc gộp "Top 9 + Khác" để hiển thị được xử lý bên dưới bằng Python, để đảm bảo
        # Tổng của biểu đồ này luôn khớp với các biểu đồ khác (không bị mất nhân sự do LIMIT).
        c_dash.execute(f"""
            SELECT chuc_danh_nghe, COUNT(*) as "Số lượng"
            FROM nhan_vien WHERE {DK_CHUAN_NV}
            AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != ''
            GROUP BY chuc_danh_nghe
            ORDER BY "Số lượng" DESC
        """)
        role_data = c_dash.fetchall()

        # e. Cơ cấu theo Độ tuổi
        c_dash.execute(f"""
            SELECT 
                CASE 
                    WHEN EXTRACT(YEAR FROM age(CURRENT_DATE, ngay_sinh)) < 25 THEN 'Dưới 25 tuổi'
                    WHEN EXTRACT(YEAR FROM age(CURRENT_DATE, ngay_sinh)) BETWEEN 25 AND 34 THEN '25-34 tuổi'
                    WHEN EXTRACT(YEAR FROM age(CURRENT_DATE, ngay_sinh)) BETWEEN 35 AND 44 THEN '35-44 tuổi'
                    WHEN EXTRACT(YEAR FROM age(CURRENT_DATE, ngay_sinh)) BETWEEN 45 AND 54 THEN '45-54 tuổi'
                    ELSE 'Từ 55 tuổi trở lên'
                END as "Độ tuổi",
                COUNT(*) as "Số lượng"
            FROM nhan_vien
            WHERE {DK_CHUAN_NV} AND ngay_sinh IS NOT NULL
            GROUP BY "Độ tuổi"
            ORDER BY MIN(EXTRACT(YEAR FROM age(CURRENT_DATE, ngay_sinh)))
        """)
        seniority_data = c_dash.fetchall()

        # f. Biểu đồ đường: Xu hướng tuyển dụng theo tháng (6 tháng gần nhất)
        c_dash.execute(f"""
            SELECT 
                TO_CHAR(DATE_TRUNC('month', ngay_vao_lam), 'MM/YYYY') as "Tháng",
                COUNT(*) as "Số lượng"
            FROM nhan_vien
            WHERE ngay_vao_lam >= (CURRENT_DATE - INTERVAL '6 months') AND {DK_CHUAN_NV}
            GROUP BY DATE_TRUNC('month', ngay_vao_lam)
            ORDER BY DATE_TRUNC('month', ngay_vao_lam) ASC
        """)
        trend_data = c_dash.fetchall()

        db_dash.close()

        # --- RENDER BIỂU ĐỒ ĐA DẠNG ---
        import plotly.express as px
        import plotly.graph_objects as go

        MODERN_PALETTE = ['#0f3b5c', '#2196F3', '#4FC3F7', '#00BFA5', '#66BB6A', '#FFB74D', '#FF7043', '#AB47BC', '#78909C']
        CHART_HEIGHT = 300

        # Hàng 1: Table + Biểu đồ thanh + Biểu đồ tròn
        row1_col1, row1_col2, row1_col3 = st.columns(3)

        with row1_col1:
            st.markdown("**💼 Cơ cấu theo Chức danh (Top 10)**")
            if role_data:
                import plotly.express as px
                import plotly.graph_objects as go
                
                df_role = pd.DataFrame(role_data)
                # Tổng luôn tính trên TOÀN BỘ chức danh (không bị ảnh hưởng bởi việc
                # gộp Top 9 + Khác bên dưới), để khớp với các biểu đồ khác trong Dashboard.
                total = df_role['Số lượng'].sum()

                # Nếu có nhiều hơn 10 chức danh: giữ 9 chức danh đông nhất, gộp phần còn
                # lại thành 1 dòng "Khác" — vừa gọn (đúng tinh thần "Top 10"), vừa không
                # làm mất người khỏi biểu đồ/tổng như khi dùng LIMIT 10 ở SQL.
                if len(df_role) > 10:
                    df_top = df_role.iloc[:9].copy()
                    so_luong_khac = df_role.iloc[9:]['Số lượng'].sum()
                    cac_chuc_danh_khac = df_role.iloc[9:]['chuc_danh_nghe'].tolist()
                    df_khac = pd.DataFrame([{
                        'chuc_danh_nghe': f"Khác ({len(cac_chuc_danh_khac)} chức danh)",
                        'Số lượng': so_luong_khac
                    }])
                    df_role = pd.concat([df_top, df_khac], ignore_index=True)
                
                # Tạo labels với format: "Chức danh\nSố lượng (tỷ lệ%)"
                labels_with_stats = []
                for _, row in df_role.iterrows():
                    pct = (row['Số lượng'] / total * 100)
                    labels_with_stats.append(f"{row['chuc_danh_nghe']}\n{row['Số lượng']} ({pct:.1f}%)")
                
                # Sử dụng biểu đồ hình tròn với labels đã format
                fig_role = go.Figure(data=[go.Pie(
                    labels=labels_with_stats,
                    values=df_role['Số lượng'],
                    hole=0.55,
                    textinfo='label',
                    textposition='outside',
                    textfont=dict(size=11, color='#1e293b'),
                    marker=dict(
                        colors=px.colors.qualitative.Safe,
                        line=dict(color='white', width=2)
                    ),
                    hovertemplate='<b>%{label}</b><br>Số lượng: %{value}<br>Tỷ lệ: %{percent:.1f}%<extra></extra>'
                )])
                fig_role.update_layout(
                    title=dict(
                        text=f"<b>Tổng: {total} nhân viên</b>",
                        x=0.5, y=0.5,
                        xanchor='center', yanchor='middle',
                        font=dict(size=14, color='#0f3b5c')
                    ),
                    showlegend=False,
                    margin=dict(t=40, b=40, l=10, r=10),
                    height=280,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                )
                st.plotly_chart(fig_role, use_container_width=True)
            else:
                st.info("Không có dữ liệu")  
        
        with row1_col2:
            st.markdown("**🎓 Cơ cấu theo Trình độ học vấn**")
            if education_data:
                df_edu = pd.DataFrame(education_data)
                df_edu['trinh_do'] = df_edu['trinh_do'].fillna('Chưa cập nhật')
                # Sử dụng biểu đồ thanh đứng thay vì tròn
                fig_edu = px.bar(
                    df_edu,
                    x='trinh_do',
                    y='Số lượng',
                    color='trinh_do',
                    text='Số lượng',
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                fig_edu.update_layout(
                    margin=dict(t=0, b=0, l=0, r=0),
                    height=280,
                    xaxis_title="",
                    yaxis_title="Số lượng",
                    showlegend=False
                )
                fig_edu.update_traces(textposition='outside')
                st.plotly_chart(fig_edu, use_container_width=True)
            else:
                st.info("Không có dữ liệu")
        
        with row1_col3:
            st.markdown("**🥧 Cơ cấu theo Phòng ban**")
            if dept_data:
                df_dept = pd.DataFrame(dept_data)
                fig_dept = px.pie(
                    df_dept, 
                    names='Phòng ban', 
                    values='Số lượng',
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                # QUAN TRỌNG: Plotly Pie mặc định tự sắp xếp lát cắt theo giá trị giảm dần
                # (sort=True), làm mất thứ tự ưu tiên phòng ban đã chuẩn hóa ở dept_data
                # (sap_xep_phong_ban_rows theo PHONG_BAN_THU_TU). Tắt sort để giữ đúng thứ tự.
                fig_dept.update_traces(sort=False)
                fig_dept.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=280)
                st.plotly_chart(fig_dept, use_container_width=True)
            else:
                st.info("Không có dữ liệu")

        # Hàng 2: Biểu đồ tròn + Biểu đồ đường + Biểu đồ thanh
        row2_col1, row2_col2, row2_col3 = st.columns(3)

        with row2_col1:
            st.markdown("**👫 Cơ cấu theo Giới tính**")
            if gender_data:
                df_gender = pd.DataFrame(gender_data)
                
                # Màu sắc nổi bật cho từng giới tính
                color_map = {
                    'Nam': '#2196F3',  # Xanh dương đẹp
                    'Nữ': '#FF6B6B',   # Đỏ hồng
                    'Khác': '#FFD93D'  # Vàng
                }
                colors = [color_map.get(g, '#95a5a6') for g in df_gender['gioi_tinh']]
                
                # Tạo donut chart với hiệu ứng đẹp
                fig_gender = go.Figure(data=[go.Pie(
                    labels=df_gender['gioi_tinh'],
                    values=df_gender['Số lượng'],
                    hole=0.4,
                    marker=dict(
                        colors=colors,
                        line=dict(color='white', width=3)
                    ),
                    textinfo='label+value+percent',
                    textposition='auto',
                    textfont=dict(size=12, color='#2c3e50', family='Arial Black'),
                    insidetextorientation='radial',
                    hovertemplate='<b>%{label}</b><br>Số lượng: %{value}<br>Tỷ lệ: %{percent:.1f}%<extra></extra>',
                    pull=[0.05 if i == 0 else 0 for i in range(len(df_gender))],  # Tách nhẹ phần tử đầu tiên
                    sort=False
                )])
                
                # Thêm vòng tròn bên trong với tổng số
                total = sum(df_gender['Số lượng'])
                fig_gender.add_annotation(
                    x=0.5, y=0.5,
                    text=f"<b>{total}</b>",
                    showarrow=False,
                    font=dict(size=24, color='#2c3e50', family='Arial Black'),
                    align='center'
                )
                fig_gender.add_annotation(
                    x=0.5, y=0.42,
                    text="Tổng",
                    showarrow=False,
                    font=dict(size=12, color='#7f8c8d', family='Arial'),
                    align='center'
                )
                
                fig_gender.update_layout(
                    margin=dict(t=10, b=10, l=10, r=10),
                    height=280,
                    showlegend=True,
                    legend=dict(
                        orientation='h',
                        yanchor='bottom',
                        y=-0.15,
                        xanchor='center',
                        x=0.5,
                        font=dict(size=12, color='#2c3e50')
                    ),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                
                st.plotly_chart(fig_gender, use_container_width=True)
            else:
                st.info("Không có dữ liệu")

        with row2_col2:
            st.markdown("**📈 Xu hướng tuyển dụng 6 tháng**")
            if trend_data:
                df_trend = pd.DataFrame(trend_data)
                fig_trend = px.line(
                    df_trend,
                    x='Tháng',
                    y='Số lượng',
                    markers=True,
                    line_shape='spline'
                )
                fig_trend.update_layout(
                    margin=dict(t=0, b=0, l=0, r=0),
                    height=280,
                    xaxis_title="",
                    yaxis_title="Số lượng",
                    showlegend=False
                )
                fig_trend.update_traces(
                    line=dict(color='#f59e0b', width=3),
                    marker=dict(size=10, color='#0f3b5c')
                )
                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.info("Không có dữ liệu")

        with row2_col3:
            st.markdown("**🎂 Cơ cấu theo Độ tuổi**")
            if seniority_data:
                df_sen = pd.DataFrame(seniority_data)
                order = ['Dưới 25 tuổi', '25-34 tuổi', '35-44 tuổi', '45-54 tuổi', 'Từ 55 tuổi trở lên']
                df_sen['Độ tuổi'] = pd.Categorical(df_sen['Độ tuổi'], categories=order, ordered=True)
                df_sen = df_sen.sort_values('Độ tuổi')
                
                # Sử dụng biểu đồ tròn với màu sắc gradient
                colors = ['#FFEAA7', '#FDCB6E', '#E17055', '#D63031', '#6C5CE7']
                fig_sen = go.Figure(data=[go.Pie(
                    labels=df_sen['Độ tuổi'],
                    values=df_sen['Số lượng'],
                    marker=dict(colors=colors[:len(df_sen)]),
                    textinfo='percent',
                    textposition='inside',
                    textfont=dict(size=11, color='white'),
                    hole=0.3,
                    # QUAN TRỌNG: mặc định Pie tự sắp xếp theo giá trị giảm dần (sort=True),
                    # phá vỡ thứ tự tuổi tăng dần đã set ở df_sen. Tắt sort để giữ đúng thứ tự nhóm tuổi.
                    sort=False
                )])
                fig_sen.update_layout(
                    margin=dict(t=10, b=40, l=10, r=10),
                    height=280,
                    showlegend=True,
                    legend=dict(
                        orientation='h',
                        yanchor='bottom',
                        y=-0.25,
                        xanchor='center',
                        x=0.5,
                        font=dict(size=10)
                    ),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                )
                st.plotly_chart(fig_sen, use_container_width=True)
            else:
                st.info("Không có dữ liệu")

        # Hàng 3: 2 biểu đồ còn lại
        row3_col1, row3_col2, row3_col3 = st.columns(3)

        with row3_col1:
            st.markdown("**📋 Trạng thái nhân sự các phòng ban**")
            if table_data:
                df_table = pd.DataFrame(table_data)
                # Định dạng số và hiển thị
                st.dataframe(df_table, hide_index=True, width='stretch', height=280)
            else:
                st.info("Không có dữ liệu")

        
            

        with row3_col3:
            st.markdown("**📊 Tổng hợp nhân sự**")
            # Hiển thị các chỉ số KPI quan trọng
            if dept_data:
                total_employees = sum([d['Số lượng'] for d in dept_data])
                total_depts = len(dept_data)
                avg_per_dept = total_employees / total_depts if total_depts > 0 else 0
                
                st.metric("🏢 Tổng số phòng ban", total_depts)
                st.metric("👥 Tổng nhân viên", f"{total_employees:,}")
                st.metric("📊 Trung bình/phòng", f"{avg_per_dept:.1f}")
                
                # Thêm thông tin phòng ban đông nhất
                if dept_data:
                    max_dept = max(dept_data, key=lambda x: x['Số lượng'])
                    st.info(f"🏆 Phòng đông nhất: **{max_dept['Phòng ban']}** ({max_dept['Số lượng']} NV)")
            else:
                st.info("Không có dữ liệu")

        # ========== KẾT THÚC PHẦN DASHBOARD NÂNG CAO ==========

# Gọi kiểm tra sinh nhật (đã xóa 2 dòng debug)
    auto_check_birthday()

    # 👇 Hiển thị banner cố định nếu có sinh nhật hôm nay
    sinh_nhat_list = st.session_state.get('sinh_nhat_hom_nay_list', [])
    if sinh_nhat_list:
        for sn in sinh_nhat_list:
            st.success(
                f"🎂 **Chúc mừng sinh nhật {sn['xung_ho']} {sn['ho_ten']} ({sn['ma_nv']})** — Hôm nay là sinh nhật của {sn['xung_ho']}! 🎉",
                icon="🎂"
            )
    
    # ========== PHẦN SINH NHẬT HOÀN CHỈNH ==========
    st.subheader("🎂 SINH NHẬT")

    # Tạo tabs cho sinh nhật
    tab_trong_thang, tab_hom_nay, tab_lich_su = st.tabs(["📅 Sinh nhật trong tháng", "🎉 Hôm nay", "📜 Lịch sử đã gửi"])

    with tab_trong_thang:
        # Lấy danh sách sinh nhật trong tháng
        c.execute("""
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, dien_thoai, email_lien_he, chuc_danh_nghe
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_sinh IS NOT NULL
            AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
            ORDER BY EXTRACT(DAY FROM ngay_sinh) ASC
        """)
        sinh_nhat_trong_thang = c.fetchall()
        
        if sinh_nhat_trong_thang:
            st.success(f"📅 Tháng {datetime.now().month} có **{len(sinh_nhat_trong_thang)}** nhân viên có sinh nhật:")
            
            # Hiển thị dạng grid
            cols = st.columns(3)
            for idx, sn in enumerate(sinh_nhat_trong_thang):
                with cols[idx % 3]:
                    ngay_sinh = sn.get('ngay_sinh')
                    if ngay_sinh:
                        # Tính tuổi
                        today = date.today()
                        tuoi = today.year - ngay_sinh.year
                        if today.month < ngay_sinh.month or (today.month == ngay_sinh.month and today.day < ngay_sinh.day):
                            tuoi -= 1
                        
                        # Tính ngày sinh nhật trong năm nay
                        sinh_nhat_nam_nay = date(today.year, ngay_sinh.month, ngay_sinh.day)
                        is_today = sinh_nhat_nam_nay == today
                        da_qua = sinh_nhat_nam_nay < today
                        
                        xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                        
                        if is_today:
                            st.markdown(f"""
                            <div style='background: linear-gradient(135deg, #ffd700 0%, #ffed4e 100%); 
                                        padding: 15px; border-radius: 15px; margin: 10px 0; 
                                        border: 2px solid #ff9800; box-shadow: 0 4px 8px rgba(0,0,0,0.1);'>
                                <div style='font-size: 30px; text-align: center;'>🎉🎂</div>
                                <h4 style='text-align: center; color: #d32f2f; margin: 5px 0;'>
                                    <b>HÔM NAY LÀ SINH NHẬT!</b>
                                </h4>
                                <h3 style='text-align: center; color: #333; margin: 5px 0;'>
                                    {xung_ho} <b>{sn['ho_ten']}</b>
                                </h3>
                                <p style='text-align: center; color: #666;'>
                                    📅 {format_date(ngay_sinh)} (🎂 {tuoi} tuổi)<br>
                                    💼 {sn.get('chuc_danh_nghe', 'Chưa cập nhật')}<br>
                                    📞 {sn.get('dien_thoai', 'Chưa cập nhật')}
                                </p>
                            </div>
                            """, unsafe_allow_html=True)
                        elif da_qua:
                            st.markdown(f"""
                            <div style='background-color: #e0e0e0; padding: 12px; border-radius: 10px; margin: 8px 0;'>
                                <b>✅ {sn['ho_ten']}</b><br>
                                🎂 Sinh ngày: {format_date(ngay_sinh)} (đã qua)<br>
                                💼 {sn.get('chuc_danh_nghe', '')}
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            ngay_con_lai = (sinh_nhat_nam_nay - today).days
                            st.markdown(f"""
                            <div style='background-color: #e3f2fd; padding: 12px; border-radius: 10px; margin: 8px 0;'>
                                <b>🎂 {sn['ho_ten']}</b><br>
                                📅 Sinh ngày: {format_date(ngay_sinh)}<br>
                                ⏰ Còn {ngay_con_lai} ngày nữa<br>
                                💼 {sn.get('chuc_danh_nghe', '')}
                            </div>
                            """, unsafe_allow_html=True)
        else:
            st.info("📭 Tháng này không có ai sinh nhật.")
            
            # Hiển thị sinh nhật tháng sau
            c.execute("""
                SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, chuc_danh_nghe
                FROM nhan_vien 
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                AND ngay_sinh IS NOT NULL
                AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE + INTERVAL '1 month')
                ORDER BY EXTRACT(DAY FROM ngay_sinh) ASC
                LIMIT 10
            """)
            sinh_nhat_thang_sau = c.fetchall()
            if sinh_nhat_thang_sau:
                st.caption("📅 Sinh nhật tháng sau:")
                for sn in sinh_nhat_thang_sau:
                    xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                    st.caption(f"🎂 {xung_ho} **{sn['ho_ten']}** - {format_date(sn['ngay_sinh'])}")

    with tab_hom_nay:
        # Lấy danh sách sinh nhật hôm nay
        c.execute("""
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, dien_thoai, email_lien_he, 
                   chuc_danh_nghe, phong_ban_lam_viec
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_sinh IS NOT NULL
            AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
            AND EXTRACT(DAY FROM ngay_sinh) = EXTRACT(DAY FROM CURRENT_DATE)
        """)
        sinh_nhat_hom_nay = c.fetchall()
        
        if sinh_nhat_hom_nay:
            st.balloons()
            for sn in sinh_nhat_hom_nay:
                ngay_sinh = sn.get('ngay_sinh')
                today = date.today()
                tuoi = today.year - ngay_sinh.year
                
                xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                
                st.markdown(f"""
                <div style='background: linear-gradient(135deg, #ff6b6b 0%, #ff8e8e 100%);
                            padding: 25px; border-radius: 20px; margin: 15px 0;
                            text-align: center; color: white;'>
                    <div style='font-size: 50px;'>🎉🎂🎉</div>
                    <h1 style='color: white; margin: 10px 0;'>CHÚC MỪNG SINH NHẬT!</h1>
                    <h2 style='color: #fff3e0; margin: 10px 0;'>{xung_ho} {sn['ho_ten']}</h2>
                    <p style='font-size: 18px;'>
                        🎂 {tuoi} tuổi - Một tuổi mới thật nhiều niềm vui! 🎂
                    </p>
                    <p style='margin-top: 15px;'>
                        📅 {format_date(ngay_sinh)} | 💼 {sn.get('chuc_danh_nghe', '')} | 
                        🏢 {sn.get('phong_ban_lam_viec', '')}
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                # Hiển thị thông tin liên hệ
                # Hiển thị thông tin liên hệ (tất cả đều thấy)
                col_phone, col_email_info = st.columns(2)
                with col_phone:
                    if sn.get('dien_thoai'):
                        st.markdown(f"📞 **SĐT:** {sn['dien_thoai']}")
                    else:
                        st.warning("⚠️ Chưa cập nhật số điện thoại")
                with col_email_info:
                    if sn.get('email_lien_he'):
                        st.markdown(f"📧 **Email:** {sn['email_lien_he']}")
                    else:
                        st.warning("⚠️ Chưa cập nhật email")

                # Nút gửi — CHỈ ADMIN mới thấy
                if st.session_state.get('role') == 'admin':
                    col_btn_zalo, col_btn_email = st.columns(2)
                    
                    with col_btn_zalo:
                        if sn.get('dien_thoai'):
                            sdt = sn['dien_thoai'].replace('+84', '0').replace(' ', '').strip()
                            if st.button(f"📱 Gửi Zalo cho {sn['ho_ten']}", 
                                         key=f"zalo_sn_{sn['id']}", 
                                         width='stretch', 
                                         type="primary"):
                                tuoi_nv = date.today().year - sn['ngay_sinh'].year
                                loi_chuc_nv = get_loi_chuc_sinh_nhat(sn['ho_ten'], sn.get('gioi_tinh'), tuoi_nv)
                                st.code(loi_chuc_nv)
                                st.markdown(f"[👉 NHẤN ĐỂ GỬI QUA ZALO CHO {sn['ho_ten']}](https://zalo.me/{sdt})")
                        else:
                            st.button("📱 Gửi Zalo", disabled=True, 
                                      key=f"zalo_sn_disabled_{sn['id']}", 
                                      width='stretch',
                                      help="Chưa có số điện thoại")

                    with col_btn_email:
                        if sn.get('email_lien_he'):
                            if st.button(f"📧 Gửi Email cho {sn['ho_ten']}", 
                                         key=f"email_sn_{sn['id']}", 
                                         width='stretch'):
                                st.info(f"📧 Email: {sn['email_lien_he']}")
                                st.toast(f"Chức năng gửi email đang phát triển!", icon="📧")
                        else:
                            st.button("📧 Gửi Email", disabled=True, 
                                      key=f"email_sn_disabled_{sn['id']}", 
                                      width='stretch',
                                      help="Chưa có email")
                else:
                    # Tài khoản thường — hiện thông báo thay vì nút
                    st.caption("💡 Liên hệ admin để gửi lời chúc sinh nhật.")
        else:
            st.info("🎉 Hôm nay không có ai sinh nhật.")
            
            # Gợi ý sinh nhật sắp tới
            c.execute("""
                SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh
                FROM nhan_vien 
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                AND ngay_sinh IS NOT NULL
                AND EXTRACT(MONTH FROM ngay_sinh) = EXTRACT(MONTH FROM CURRENT_DATE)
                AND EXTRACT(DAY FROM ngay_sinh) > EXTRACT(DAY FROM CURRENT_DATE)
                ORDER BY EXTRACT(DAY FROM ngay_sinh) ASC
                LIMIT 3
            """)
            sinh_nhat_sap_toi = c.fetchall()
            if sinh_nhat_sap_toi:
                st.subheader("📅 Sinh nhật sắp tới trong tháng:")
                for sn in sinh_nhat_sap_toi:
                    xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                    st.info(f"🎂 {xung_ho} **{sn['ho_ten']}** - {format_date(sn['ngay_sinh'])}")

    with tab_lich_su:
        if st.session_state.role == "admin":
            st.subheader("📜 Lịch sử đã gửi lời chúc sinh nhật")
            
            # Kiểm tra bảng lịch sử tồn tại chưa
            try:
                c.execute("""
                    SELECT ls.*, nv.ho_ten, nv.ma_nv
                    FROM lich_su_gui_loi_chuc ls
                    JOIN nhan_vien nv ON ls.nhan_vien_id = nv.id
                    WHERE ls.loai_chuc = 'SINH_NHAT'
                    ORDER BY ls.ngay_gui DESC
                    LIMIT 50
                """)
                lich_su = c.fetchall()
                
                if lich_su:
                    ls_data = []
                    for ls in lich_su:
                        ls_data.append({
                            "Ngày gửi": format_date(ls['ngay_gui']),
                            "Mã NV": ls['ma_nv'],
                            "Họ tên": ls['ho_ten'],
                            "Kênh gửi": ls['kenh_gui'],
                            "Trạng thái": "✅ Đã gửi" if ls['trang_thai'] == 'DA_GUI' else ls['trang_thai']
                        })
                    df_ls = pd.DataFrame(ls_data)
                    st.dataframe(df_ls, width='stretch', hide_index=True)
                else:
                    st.info("📭 Chưa có lịch sử gửi lời chúc nào.")
            except Exception as e:
                st.info("📭 Chưa có dữ liệu lịch sử. Bảng lịch sử có thể chưa được tạo.")
        else:
            st.info("🔒 Chỉ Admin mới xem được lịch sử gửi lời chúc.")

    st.divider()

    # Nút gửi lời chúc sinh nhật (chỉ admin)
    if st.session_state.role == "admin" and sinh_nhat_trong_thang:
        with st.expander("💌 GỬI LỜI CHÚC SINH NHẬT", expanded=False):
            st.subheader("Gửi lời chúc sinh nhật đến nhân viên")
            
            # Chọn nhân viên để gửi
            sn_options = {}
            for sn in sinh_nhat_trong_thang:
                ngay_sinh = sn.get('ngay_sinh')
                xung_ho = get_xung_ho(sn.get('gioi_tinh'), sn['ho_ten'])
                label = f"{xung_ho} {sn['ho_ten']} - {format_date(ngay_sinh)}"
                sn_options[label] = sn
            
            # SAU KHI SỬA (ĐÚNG) — dùng key động theo từng nhân viên
            selected_label = st.selectbox("Chọn nhân viên:", list(sn_options.keys()), key="chon_sn_gui", help="💡 Gõ mã NV hoặc tên để tìm nhanh trong danh sách")
            selected_sn = sn_options[selected_label]

            # Tính tuổi
            if selected_sn.get('ngay_sinh'):
                today = date.today()
                tuoi = today.year - selected_sn['ngay_sinh'].year
                if today.month < selected_sn['ngay_sinh'].month or (
                    today.month == selected_sn['ngay_sinh'].month and 
                    today.day < selected_sn['ngay_sinh'].day
                ):
                    tuoi -= 1
            else:
                tuoi = None

            default_chuc = get_loi_chuc_sinh_nhat(
                selected_sn['ho_ten'], 
                selected_sn.get('gioi_tinh'), 
                tuoi
            )

            # ✅ KEY ĐỘNG theo nv_id — buộc Streamlit re-render lại text_area mỗi khi chọn người mới
            loi_chuc = st.text_area(
                "📝 Lời chúc sinh nhật:", 
                value=default_chuc, 
                height=250, 
                key=f"loi_chuc_sn_{selected_sn['id']}"  # ← thay "loi_chuc_sn_gui" bằng key động này
            )
            
            col_zalo, col_email, col_cancel = st.columns(3)
            
            with col_zalo:
                if st.button("📱 GỬI QUA ZALO", width='stretch', type="primary"):
                    sdt = selected_sn.get('dien_thoai', '')
                    if sdt:
                        sdt = sdt.replace('+84', '0').replace(' ', '').strip()
                        st.code(loi_chuc)
                        st.markdown(f"[👉 NHẤN VÀO ĐÂY ĐỂ GỬI QUA ZALO CHO {selected_sn['ho_ten']}](https://zalo.me/{sdt})")
                        st.success(f"✅ Đã sao chép nội dung! Vui lòng nhấn link Zalo để gửi.")
                        
                        # Lưu lịch sử
                        try:
                            db_log = st.session_state.db_engine.get_connection()
                            cur_log = db_log.cursor()
                            cur_log.execute("""
                                INSERT INTO lich_su_gui_loi_chuc (nhan_vien_id, loai_chuc, noi_dung, kenh_gui, trang_thai)
                                VALUES (%s, %s, %s, %s, %s)
                            """, (selected_sn['id'], 'SINH_NHAT', loi_chuc[:500], 'ZALO', 'DA_GUI'))
                            db_log.commit()
                            db_log.close()
                            st.toast("Đã lưu lịch sử gửi!", icon="✅")
                        except:
                            pass
                    else:
                        st.error("❌ Nhân viên chưa có số điện thoại! Vui lòng cập nhật SĐT trước.")
            
            with col_email:
                if st.button("📧 GỬI QUA EMAIL", width='stretch'):
                    email = selected_sn.get('email_lien_he', '')
                    if email:
                        try:
                            xung_ho = get_xung_ho(selected_sn.get('gioi_tinh'), selected_sn['ho_ten'])
                            
                            msg = MIMEMultipart()
                            msg['From'] = EMAIL_CONFIG['email']
                            msg['To'] = email
                            msg['Subject'] = f"🎂 Chúc mừng sinh nhật {xung_ho} {selected_sn['ho_ten']} - Công ty CP Cảng Hòn La"
                            
                            html_content = f"""
                            <html>
                            <head>
                                <meta charset="UTF-8">
                            </head>
                            <body style='font-family: "Times New Roman", Arial, sans-serif;'>
                                <div style='background: linear-gradient(135deg, #ffd700 0%, #ff9800 100%); 
                                            padding: 20px; text-align: center; border-radius: 10px;'>
                                    <h1 style='color: white;'>🎂 CHÚC MỪNG SINH NHẬT 🎂</h1>
                                </div>
                                <div style='padding: 20px; line-height: 1.6;'>
                                    {loi_chuc.replace(chr(10), '<br>')}
                                </div>
                                <hr>
                                <p style='color: #999; font-size: 11px; text-align: center;'>
                                    Email được gửi tự động từ hệ thống HRM-Port Công ty CP Cảng Hòn La<br>
                                    Địa chỉ: {COMPANY_CONFIG.get('dia_chi', '')} | Điện thoại: {COMPANY_CONFIG.get('dien_thoai_cty', '')}
                                </p>
                            </body>
                            </html>
                            """
                            msg.attach(MIMEText(html_content, 'html', 'utf-8'))
                            
                            server = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
                            server.starttls()
                            server.login(EMAIL_CONFIG['email'], EMAIL_CONFIG['password'])
                            server.send_message(msg)
                            server.quit()
                            
                            st.success(f"✅ Đã gửi lời chúc sinh nhật qua email cho {xung_ho} {selected_sn['ho_ten']}!")
                            
                            # Lưu lịch sử
                            db_log = st.session_state.db_engine.get_connection()
                            cur_log = db_log.cursor()
                            cur_log.execute("""
                                INSERT INTO lich_su_gui_loi_chuc (nhan_vien_id, loai_chuc, noi_dung, kenh_gui, trang_thai)
                                VALUES (%s, %s, %s, %s, %s)
                            """, (selected_sn['id'], 'SINH_NHAT', loi_chuc[:500], 'EMAIL', 'DA_GUI'))
                            db_log.commit()
                            db_log.close()
                        except Exception as e:
                            st.error(f"❌ Lỗi gửi email: {e}")
                    else:
                        st.error("❌ Nhân viên chưa có email! Vui lòng cập nhật email trước.")
            
            with col_cancel:
                st.write("")  # Placeholder

    # ========== KẾT THÚC PHẦN SINH NHẬT ==========


    # ── Thông báo ──
    st.subheader("📌 Thông báo")
    db3 = st.session_state.db_engine.get_connection()
    c3 = db3.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c3.execute("SELECT ho_ten FROM nhan_vien WHERE DATE(ngay_vao_lam)=CURRENT_DATE")
    hn = c3.fetchall()
    c3.execute("SELECT ho_ten FROM nhan_vien WHERE DATE(ngay_vao_lam)=CURRENT_DATE - INTERVAL '1 day'")
    hq = c3.fetchall()
    if hn:
        st.success(f"🟢 Hôm nay có thêm: **{', '.join([x['ho_ten'] for x in hn])}**")
    if hq:
        st.info(f"🔵 Hôm qua có thêm: **{', '.join([x['ho_ten'] for x in hq])}**")
    if st.session_state.role == "admin":
        c3.execute("""
            SELECT STT, ma_nv, ho_ten, ngay_vao_lam, 
                   (ngay_vao_lam + INTERVAL '30 days')::DATE as ngay_ket_thuc_tv,
                   ((ngay_vao_lam + INTERVAL '30 days')::DATE - CURRENT_DATE) as ngay_con_lai
            FROM nhan_vien 
            WHERE trang_thai = 'THU_VIEC' 
            AND (ngay_vao_lam + INTERVAL '30 days')::DATE <= CURRENT_DATE + INTERVAL '5 days'
            ORDER BY ngay_con_lai ASC
        """)
        tv_sap_het = c3.fetchall()
        for x in tv_sap_het:
            ngay_con_lai = x['ngay_con_lai']
            if isinstance(ngay_con_lai, timedelta):
                ngay_con_lai = ngay_con_lai.days
            if ngay_con_lai < 0:
                st.error(f"🔴 **{x.get('ma_nv','')} {x['ho_ten']}** - ĐÃ QUÁ THỜI HẠN THỬ VIỆC {abs(ngay_con_lai)} NGÀY!")
            elif ngay_con_lai == 0:
                st.error(f"⚠️ **{x.get('ma_nv','')} {x['ho_ten']}** - HÔM NAY LÀ NGÀY CUỐI HỢP ĐỒNG THỬ VIỆC!")
            else:
                st.warning(f"⚠️ **{x.get('ma_nv','')} {x['ho_ten']}** còn **{ngay_con_lai}** ngày sẽ kết thúc hợp đồng thử việc!")
    db3.close()
    
    
    if st.session_state.role == "admin":
        st.markdown("#### 💾 Sao lưu dữ liệu")
        col_bk1, col_bk2 = st.columns(2)

        with col_bk1:
            if st.button("💾 BACKUP DỮ LIỆU NGAY", width='stretch'):
                try:
                    from backup_data import backup_all
                    with st.spinner("⏳ Đang backup bảng Ứng viên, Nhân viên và hồ sơ trên Supabase Storage..."):
                        result = backup_all()

                    if result["mode"] == "local":
                        st.success(f"✅ Đã backup xong! Thư mục: {result['dest_folder']}")
                    else:
                        st.success("✅ Đã backup xong! App đang chạy trên môi trường Cloud (không có ổ D: của bạn), "
                                    "nên kết quả được nén thành file zip — bấm nút bên dưới để tải về máy:")
                        st.download_button(
                            label="📥 TẢI FILE BACKUP (.zip)",
                            data=result["zip_bytes"],
                            file_name=result["zip_filename"],
                            mime="application/zip",
                            width='stretch'
                        )

                    for table, res in result['db'].items():
                        if res[0]:
                            st.caption(f"✔️ Bảng `{table}`: {res[1]} dòng")
                        else:
                            st.caption(f"❌ Bảng `{table}`: {res[1]}")
                    if result['storage']['ok']:
                        st.caption(f"✔️ Storage: đã tải {result['storage']['count']} file hồ sơ")
                    else:
                        st.caption(f"❌ Storage: {result['storage']['error']} (đã tải {result['storage']['count']} file)")
                except ImportError:
                    st.error("❌ Không tìm thấy `backup_data.py`. Hãy đặt file này cùng thư mục với app.py.")
                except Exception as e:
                    st.error(f"❌ Lỗi khi backup: {e}")
            st.caption("Backup dữ liệu bảng `ung_vien`, `nhan_vien` (Excel) + toàn bộ file hồ sơ trên Supabase Storage. "
                       "Nếu chạy trên máy Windows local → lưu vào `D:\\hrm-port9\\backup`. Nếu chạy trên Cloud → tải về dạng file zip.")

        with col_bk2:
            with st.popover("🗓️ Lịch backup tự động", width='stretch'):
                is_windows = (os.name == 'nt')
                st.caption("Dùng **Windows Task Scheduler** để tự động chạy backup vào **02:00 sáng Thứ 7 hàng tuần**. "
                           "Chỉ tạo được lịch khi bấm nút này **ngay trên máy Windows** nơi bạn muốn lưu file backup — "
                           "không thể bật từ xa qua Streamlit Cloud.")
                if not is_windows:
                    st.warning("⚠️ App hiện đang chạy trên môi trường Cloud (không phải Windows), nên không thể tạo "
                               "lịch Task Scheduler tại đây. Cách làm đúng: copy file `backup_data.py` cùng file cấu "
                               "hình kết nối (.env) xuống máy Windows của bạn, rồi mở app này **chạy local** "
                               "(`streamlit run app.py` ngay trên máy đó) để bấm nút BẬT lịch — lúc đó Task Scheduler "
                               "sẽ được tạo đúng trên máy bạn và tự backup vào D:\\ hàng tuần dù sau đó bạn tắt app đi.")
                else:
                    if st.button("✅ BẬT lịch backup tự động", width='stretch'):
                        try:
                            python_exe = sys.executable
                            script_path = os.path.abspath("backup_data.py")
                            task_cmd = (
                                'schtasks /Create /TN "HRM_Port_Backup_Weekly" '
                                f'/TR "\\"{python_exe}\\" \\"{script_path}\\"" '
                                '/SC WEEKLY /D SAT /ST 02:00 /F'
                            )
                            result = subprocess.run(task_cmd, shell=True, capture_output=True, text=True)
                            if result.returncode == 0:
                                st.success("✅ Đã tạo lịch: tự động backup 02:00 sáng Thứ 7 hàng tuần.")
                            else:
                                st.error(f"❌ Không tạo được lịch: {result.stderr or result.stdout}")
                        except Exception as e:
                            st.error(f"❌ Lỗi khi tạo lịch: {e}")

                    if st.button("🗑️ TẮT lịch backup tự động", width='stretch'):
                        try:
                            result = subprocess.run(
                                'schtasks /Delete /TN "HRM_Port_Backup_Weekly" /F',
                                shell=True, capture_output=True, text=True
                            )
                            if result.returncode == 0:
                                st.success("✅ Đã tắt lịch backup tự động.")
                            else:
                                st.warning(f"⚠️ {result.stderr or result.stdout}")
                        except Exception as e:
                            st.error(f"❌ Lỗi: {e}")
    
    
    
# ========== ỨNG VIÊN ==========
elif menu == "👤 Ứng viên":
    st.title("👤 Ứng viên")
    ensure_chuc_danh_ung_vien_table()
    su = st.text_input("🔍 Tìm kiếm", key="suv")
    
    # Kiểm tra nếu đang chuyển từ ứng viên sang nhân viên
    if 'show_chuyen_nv_form' in st.session_state and st.session_state.show_chuyen_nv_form:
        st.subheader("📝 CHUYỂN ỨNG VIÊN THÀNH NHÂN VIÊN")
        uv_data = st.session_state.get('chuyen_uv_data', {})
        
        # Lấy danh sách chức danh từ database
        db_chuc = st.session_state.db_engine.get_connection()
        c_chuc = db_chuc.cursor()
        c_chuc.execute("SELECT DISTINCT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
        dschucdanh = [row[0] for row in c_chuc.fetchall()]
        db_chuc.close()
        
        with st.form("chuyen_uv_to_nv_form"):
            st.markdown(f"**Ứng viên:** {uv_data.get('ho_ten', '')}")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                ho_ten_nv = st.text_input("Họ và tên *", value=uv_data.get('ho_ten', ''))
                ngay_sinh_nv = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(uv_data.get('ngay_sinh')), placeholder="dd/mm/yyyy", max_chars=10)
                gioi_tinh_nv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"], index=["", "Nam", "Nữ", "Khác"].index(uv_data.get('gioi_tinh', '')) if uv_data.get('gioi_tinh') in ["Nam", "Nữ", "Khác"] else 0)
                quoc_tich_nv = st.text_input("Quốc tịch", value="Việt Nam")
                dan_toc_nv = st.text_input("Dân tộc", value="Kinh")
            with col2:
                so_cccd_nv = st.text_input("CCCD")
                ngay_cap_cccd_nv = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                noi_cap_cccd_nv = st.text_input("Nơi cấp CCCD")
                nguyen_quan_nv = st.text_input("Nguyên quán")
                thuong_tru_nv = st.text_area("Thường trú", value=uv_data.get('ghi_chu', ''), height=68)
            with col3:
                dien_thoai_nv = st.text_input("SĐT", value=uv_data.get('dien_thoai', ''))
                email_nv = st.text_input("Email")
                chuc_danh_nv = st.selectbox("Chức danh", [""] + dschucdanh, index=([""] + dschucdanh).index(uv_data.get('vi_tri', '')) if uv_data.get('vi_tri', '') in dschucdanh else 0)
                phong_ban_nv = st.text_input("Phòng ban")
                noi_lam_viec_nv = st.text_input("Nơi làm việc", value="Cảng THQT Hòn La")
                trinh_do_nv = st.selectbox("Trình độ", [""] + TRINH_DO_LIST)
                anh_ho_so_nv = st.file_uploader("Ảnh hồ sơ", type=["png", "jpg", "jpeg"], key="anh_ho_so_chuyen")
            
            st.divider()
            st.caption("💼 Hợp đồng & BHXH")
            col4, col5, col6 = st.columns(3)
            with col4:
                loai_hd_chuyen = st.selectbox("Loại HĐ *", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"])
                ngay_vao_lam_chuyen = st.date_input("Ngày vào làm", value=uv_data.get('ngay_vao_lam', date.today()))
                ngay_ket_thuc_chuyen = st.text_input("Ngày kết thúc (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                ma_bhxh_chuyen = st.text_input("Mã BHXH")
                bat_dau_bh_chuyen = st.text_input("Bắt đầu BH (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
            with col5:
                luong_bh_chuyen = st.text_input("Lương BH")
                he_so_luong_chuyen = st.text_input("Hệ số lương")
                pc_chuc_vu_chuyen = st.text_input("PC chức vụ")
                pc_tnvk_chuyen = st.text_input("PC TNVK (%)")
                pc_tnn_chuyen = st.text_input("PC TNN (%)")
            with col6:
                muc_huong_bhyt_chuyen = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"])
                ty_le_dong_chuyen = st.text_input("Tỷ lệ đóng (%)")
                muc_tien_dong_chuyen = st.text_input("Mức tiền đóng")
                phuong_thuc_dong_chuyen = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"])
                nhom_bhxh_chuyen = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"])
            
            st.divider()
            st.caption("🏦 Ngân hàng & KCB")
            col7, col8, col9 = st.columns(3)
            with col7:
                stk_chuyen = st.text_input("STK")
                # Tạo dropdown cho chi nhánh ngân hàng
                bank_chuyen_index = 0
                chi_nhanh_nh_chuyen = st.selectbox("Chi nhánh NH", options=[""] + BANK_LIST, index=bank_chuyen_index, key="chuyen_cnh")
                tinh_kcb_chuyen = st.text_input("Tỉnh KCB")
                noi_kcb_chuyen = st.text_input("Nơi KCB")
            with col8:
                tinh_nhan_hs_chuyen = st.text_input("Tỉnh/TP nhận HS")
                phuong_nhan_hs_chuyen = st.text_input("Phường/Xã nhận HS")
                dia_chi_nhan_hs_chuyen = st.text_area("Địa chỉ nhận HS", height=100)
            with col9:
                dk_nhan_so_chuyen = st.selectbox("ĐK nhận sổ", ["Có", "Không"])
                ho_so_chuyen = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"])
            
            col_confirm1, col_confirm2 = st.columns(2)
            with col_confirm1:
                if st.form_submit_button("✅ XÁC NHẬN CHUYỂN", width='stretch', type="primary"):
                    if ho_ten_nv:
                        # Kiểm tra định dạng ngày
                        ngay_loi = []
                        if ngay_sinh_nv and not parse_date(ngay_sinh_nv): 
                            ngay_loi.append("Ngày sinh")
                        if ngay_cap_cccd_nv and not parse_date(ngay_cap_cccd_nv): 
                            ngay_loi.append("Ngày cấp CCCD")
                        if ngay_ket_thuc_chuyen and not parse_date(ngay_ket_thuc_chuyen): 
                            ngay_loi.append("Ngày kết thúc")
                        if bat_dau_bh_chuyen and not parse_date(bat_dau_bh_chuyen): 
                            ngay_loi.append("Bắt đầu BH")
                        if ngay_loi:
                            st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                        else:
                            try:
                                # Tạo ten_don_vi_thu_huong từ ho_ten
                                ten_don_vi_thu_huong = generate_ten_don_vi_thu_huong(ho_ten_nv)
                                
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                
                                c.execute("SELECT COALESCE(MAX(STT), 0) + 1 FROM nhan_vien")
                                stt_moi = c.fetchone()[0]

                                # Tạo STT và mã nhân viên mới
                                c.execute("SELECT COALESCE(MAX(CAST(SUBSTRING(ma_nv FROM 2) AS INTEGER)), 0)+1 FROM nhan_vien WHERE ma_nv LIKE 'C%'")
                                so_moi = c.fetchone()[0]
                                ma_nv = f"C{so_moi:03d}"
                                c.execute("SELECT COALESCE(MAX(STT),0)+1 FROM nhan_vien")
                                stt_moi = c.fetchone()[0]
                                
                                nhl = ngay_vao_lam_chuyen
                                tbd_val = parse_date(bat_dau_bh_chuyen) if bat_dau_bh_chuyen and bat_dau_bh_chuyen.strip() else None
                                
                                # Tạo số hợp đồng theo loại
                                if loai_hd_chuyen == "Thử việc":
                                    trang_thai_nv = 'THU_VIEC'
                                    trang_thai_bhxh = 'CHUA_DONG'
                                    c.execute("""
                                        SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) 
                                        FROM nhan_vien 
                                        WHERE so_hdld LIKE '%/HĐTV-CHL' 
                                        AND trang_thai IN ('THU_VIEC', 'DANG_LAM')
                                    """)
                                    tv_cnt = c.fetchone()[0] + 1
                                    so_hd = f"{tv_cnt:02d}/{nhl.year}/HĐTV-CHL"
                                else:
                                    trang_thai_nv = 'DANG_LAM'
                                    trang_thai_bhxh = 'DANG_DONG'
                                    if not tbd_val:
                                        tbd_val = nhl
                                    c.execute("""
                                        SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) 
                                        FROM nhan_vien 
                                        WHERE so_hdld LIKE '%/HĐLĐ-CHL' 
                                        AND trang_thai = 'DANG_LAM'
                                        AND loai_hop_dong != 'Thử việc'
                                    """)
                                    so_hd_cnt = c.fetchone()[0] or 0
                                    so_hd = f"{so_hd_cnt + 1:02d}/{nhl.year}/HĐLĐ-CHL"
                                
                                # Thêm nhân viên mới (đã thêm trường ten_don_vi_thu_huong, trinh_do)
                                c.execute("""
                                    INSERT INTO nhan_vien (STT, ma_nv, so_hdld, ho_ten, chuc_danh_nghe, 
                                        ngay_sinh, gioi_tinh, so_cccd, ngay_cap_cccd, noi_cap_cccd,
                                        nguyen_quan, thuong_tru, dien_thoai, email, email_lien_he, ho_so,
                                        luong_bao_hiem, ma_so_bhxh, ngay_vao_lam, noi_lam_viec,
                                        so_tai_khoan_nh, chi_nhanh_nh, ngay_ky_hd, loai_hop_dong,
                                        nhom_bhxh, thang_bat_dau_bh, thang_ket_thuc_bh, trang_thai, trang_thai_bhxh,
                                        phong_ban_lam_viec, ngay_ket_thuc, quoc_tich, dan_toc, 
                                        he_so_luong, phu_cap_chuc_vu, phu_cap_tnvk, phu_cap_tnn,
                                        muc_huong_bhyt, ty_le_dong, muc_tien_dong, phuong_thuc_dong,
                                        tinh_nhan_hs, phuong_nhan_hs, dia_chi_nhan_hs, 
                                        tinh_kcb, noi_dang_ky_kcb, dang_ky_nhan_so, ten_don_vi_thu_huong, trinh_do)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                    RETURNING id
                                """, (
                                    stt_moi, ma_nv, so_hd, ho_ten_nv, chuc_danh_nv,
                                    parse_date(ngay_sinh_nv), gioi_tinh_nv, so_cccd_nv, parse_date(ngay_cap_cccd_nv), noi_cap_cccd_nv,
                                    nguyen_quan_nv, thuong_tru_nv, dien_thoai_nv, email_nv, email_nv, ho_so_chuyen,
                                    luong_bh_chuyen, ma_bhxh_chuyen, ngay_vao_lam_chuyen, noi_lam_viec_nv,
                                    stk_chuyen, chi_nhanh_nh_chuyen, ngay_vao_lam_chuyen, loai_hd_chuyen,
                                    nhom_bhxh_chuyen, tbd_val, parse_date(ngay_ket_thuc_chuyen), trang_thai_nv, trang_thai_bhxh,
                                    phong_ban_nv, parse_date(ngay_ket_thuc_chuyen), quoc_tich_nv, dan_toc_nv,
                                    to_float_or_none(he_so_luong_chuyen), to_float_or_none(pc_chuc_vu_chuyen),
                                    to_float_or_none(pc_tnvk_chuyen), to_float_or_none(pc_tnn_chuyen),
                                    muc_huong_bhyt_chuyen, to_float_or_none(ty_le_dong_chuyen), to_float_or_none(muc_tien_dong_chuyen),
                                    phuong_thuc_dong_chuyen, tinh_nhan_hs_chuyen, phuong_nhan_hs_chuyen, dia_chi_nhan_hs_chuyen,
                                    tinh_kcb_chuyen, noi_kcb_chuyen, dk_nhan_so_chuyen, ten_don_vi_thu_huong, trinh_do_nv
                                ))
                                nhan_vien_id_moi = c.fetchone()[0]

                                # Upload ảnh hồ sơ (nếu có) — cần id vừa tạo để đặt tên thư mục trên Storage
                                if anh_ho_so_nv is not None:
                                    storage_path_anh = upload_anh_ho_so(ma_nv, ho_ten_nv, anh_ho_so_nv)
                                    if storage_path_anh:
                                        c.execute("UPDATE nhan_vien SET anh_ho_so=%s WHERE id=%s", (storage_path_anh, nhan_vien_id_moi))

                                # Cập nhật trạng thái ứng viên
                                c.execute("UPDATE ung_vien SET trang_thai='DA_NHAN_VIEC', ma_nv=%s WHERE id=%s", 
                                         (ma_nv, st.session_state['chuyen_uv_id']))
                                
                                # Thêm lịch sử công tác
                                c.execute("""
                                    INSERT INTO lich_su_cong_tac (nhan_vien_id, tu_ngay, chuc_danh, phong_ban, noi_lam_viec, loai_hop_dong, he_so_luong, so_hop_dong)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                                """, (nhan_vien_id_moi, ngay_vao_lam_chuyen, chuc_danh_nv, phong_ban_nv, noi_lam_viec_nv, loai_hd_chuyen, 
                                      to_float_or_none(he_so_luong_chuyen), so_hd))
                                
                                db.commit()
                                db.close()
                                
                                st.success(f"✅ Đã chuyển {ho_ten_nv} thành nhân viên! Mã NV: {ma_nv}")
                                st.cache_data.clear()
                                # Xóa session state
                                del st.session_state['show_chuyen_nv_form']
                                del st.session_state['chuyen_uv_id']
                                del st.session_state['chuyen_uv_data']
                                st.rerun()
                                
                            except Exception as e:
                                db.rollback()
                                db.close()
                                st.error(f"❌ Lỗi khi chuyển: {e}")
                    else:
                        st.error("Họ tên không được để trống!")
            
            with col_confirm2:
                if st.form_submit_button("❌ HỦY", width='stretch'):
                    del st.session_state['show_chuyen_nv_form']
                    del st.session_state['chuyen_uv_id']
                    del st.session_state['chuyen_uv_data']
                    st.rerun()
        
        st.divider()
        st.stop()  # Dừng lại để không hiển thị danh sách ứng viên phía dưới
    
    db_f = st.session_state.db_engine.get_connection()
    c_f = db_f.cursor()
    c_f.execute("SELECT ten_chuc_danh FROM chuc_danh_ung_vien ORDER BY ten_chuc_danh")
    ds_vi_tri = [row[0] for row in c_f.fetchall()]
    c_f.execute("SELECT DISTINCT vi_tri_du_tuyen FROM ung_vien WHERE vi_tri_du_tuyen IS NOT NULL AND vi_tri_du_tuyen != '' ORDER BY vi_tri_du_tuyen")
    for row in c_f.fetchall():
        if row[0] not in ds_vi_tri:
            ds_vi_tri.append(row[0])
    db_f.close()
    
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        filter_vi_tri = st.selectbox("🔍 Lọc Vị trí dự tuyển:", ["Tất cả"] + ds_vi_tri)
    
    # Chỉ admin mới thấy nút thêm ứng viên
    if st.session_state.role == "admin":
        with st.expander("➕ THÊM ỨNG VIÊN MỚI", expanded=False):
            with st.form("add_uv_form"):
                db_f = st.session_state.db_engine.get_connection()
                c_f = db_f.cursor()
                c_f.execute("SELECT ten_chuc_danh FROM chuc_danh_ung_vien ORDER BY ten_chuc_danh")
                ds_vt_uv = [row[0] for row in c_f.fetchall()]
                db_f.close()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    ho_ten_uv = st.text_input("Họ và tên *")
                    vi_tri_uv = st.selectbox("Vị trí dự tuyển", [""] + ds_vt_uv)
                    dien_thoai_uv = st.text_input("SĐT")
                with col2:
                    ngay_sinh_uv = st.text_input("Ngày sinh (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    gioi_tinh_uv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"])
                with col3:
                    ngay_vao_lam_uv = st.text_input("Ngày vào làm (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    ghi_chu_uv = st.text_area("Ghi chú")
                
                if st.form_submit_button("💾 LƯU", width='stretch'):
                    if not can_edit():
                        st.error("❌ Bạn không có quyền thực hiện thao tác này!")
                    else:
                        if ho_ten_uv:
                            ngay_loi = []
                            if ngay_sinh_uv and not parse_date(ngay_sinh_uv): 
                                ngay_loi.append("Ngày sinh")
                            if ngay_vao_lam_uv and not parse_date(ngay_vao_lam_uv): 
                                ngay_loi.append("Ngày vào làm")
                            if ngay_loi:
                                st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                            else:
                                try:
                                    db = st.session_state.db_engine.get_connection()
                                    c = db.cursor()
                                    c.execute("""INSERT INTO ung_vien (ho_ten, vi_tri_du_tuyen, dien_thoai, 
                                        ngay_sinh, gioi_tinh, ngay_vao_lam, luong_bao_hiem, trang_thai)
                                        VALUES (%s, %s, %s, %s, %s, %s, %s, 'CHO_DUYET')""",
                                        (ho_ten_uv, vi_tri_uv, dien_thoai_uv, parse_date(ngay_sinh_uv),
                                         gioi_tinh_uv, parse_date(ngay_vao_lam_uv), ghi_chu_uv))
                                    new_id = c.lastrowid
                                    ma_uv = f"UV{new_id:04d}"
                                    c.execute("UPDATE ung_vien SET ma_uv = %s WHERE id = %s", (ma_uv, new_id))
                                    db.commit()
                                    db.close()
                                    st.success(f"✅ Đã thêm ứng viên: {ho_ten_uv} (Mã: {ma_uv})")
                                    st.cache_data.clear()
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Lỗi khi thêm ứng viên: {e}")
                        else:
                            st.error("Họ tên không được để trống!")
    
    st.divider()
    
    t1, t2, t3, t4 = st.tabs(["📋 Tất cả", "⏳ Chờ duyệt", "✅ Đã nhận", "❌ Từ chối"])
    tm = {"📋 Tất cả": "", "⏳ Chờ duyệt": "CHO_DUYET", "✅ Đã nhận": "DA_NHAN_VIEC", "❌ Từ chối": "TU_CHOI"}
    
    for tn, tab in zip(tm.keys(), [t1, t2, t3, t4]):
        with tab:
            tt = tm[tn]
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            sql = "SELECT id, ma_uv, ho_ten, vi_tri_du_tuyen, dien_thoai, ngay_vao_lam, luong_bao_hiem, ngay_sinh, trang_thai FROM ung_vien WHERE 1=1"
            params = []
            if tt:
                sql += " AND trang_thai = %s"
                params.append(tt)
            if su:
                sql += " AND (ho_ten LIKE %s OR dien_thoai LIKE %s OR ma_uv LIKE %s)"
                params.extend([f'%{su}%', f'%{su}%', f'%{su}%'])
            if filter_vi_tri != "Tất cả":
                sql += " AND vi_tri_du_tuyen = %s"
                params.append(filter_vi_tri)
            sql += " ORDER BY id ASC"
            c.execute(sql, tuple(params))
            ds = c.fetchall()
            db.close()
            
            if ds:
                df = pd.DataFrame(ds)
                for col in df.columns:
                    if 'ngay' in col.lower():
                        df[col] = df[col].apply(format_date)
                
                display_cols = ['ma_uv', 'ho_ten', 'vi_tri_du_tuyen', 'dien_thoai', 'ngay_vao_lam', 'luong_bao_hiem', 'ngay_sinh', 'trang_thai']
                available_cols = [c for c in display_cols if c in df.columns]
                df_show = df[available_cols]
                
                col_map = {
                    'ma_uv': 'Mã UV',
                    'ho_ten': 'Họ tên',
                    'vi_tri_du_tuyen': 'Vị trí dự tuyển',
                    'dien_thoai': 'SĐT',
                    'ngay_vao_lam': 'Ngày vào làm',
                    'luong_bao_hiem': 'Ghi chú',
                    'ngay_sinh': 'Ngày sinh',
                    'trang_thai': 'Trạng thái',
                }
                df_show.rename(columns=col_map, inplace=True)
                
                st.caption(f"📌 {len(ds)} kết quả.")
                
                if st.session_state.role == "admin":
                    # Admin: hiển thị bảng có checkbox và nút chức năng
                    if 'selected' not in df.columns:
                        df.insert(0, 'selected', False)
                    df_show_with_checkbox = df[['selected'] + [c for c in df.columns if c in display_cols]]
                    df_show_with_checkbox.rename(columns={'selected': 'Chọn'}, inplace=True)
                    
                    edited_df = st.data_editor(
                        df_show_with_checkbox,
                        column_config={"Chọn": st.column_config.CheckboxColumn("Chọn", default=False)},
                        disabled=[col for col in df_show_with_checkbox.columns if col != 'Chọn'],
                        hide_index=True,
                        height=400,
                        key=f"uv_editor_{tn}"
                    )
                    
                    if edited_df is not None:
                        selected_rows = edited_df[edited_df['Chọn'] == True]
                        if len(selected_rows) > 1:
                            st.error("⚠️ Chỉ được chọn 1 ứng viên!")
                        elif len(selected_rows) == 1:
                            selected_idx = selected_rows.index[0]
                            selected_nv = df.iloc[selected_idx]
                            col_btn1, col_btn2 = st.columns(2)
                            with col_btn1:
                                # Chỉ hiển thị nút Sửa khi đã chọn 1 ứng viên
                                if st.button(f"✏️ SỬA", key=f"edit_sel_{tn}", width='stretch'):
                                    st.session_state['edit_uv_id'] = int(selected_nv['id'])
                                    st.rerun()
                            if tn == "⏳ Chờ duyệt":
                                with col_btn2:
                                    if st.button(f"👥 CHUYỂN SANG NHÂN VIÊN", type="primary", key=f"chuyen_uv_{tn}"):
                                        try:
                                            db = st.session_state.db_engine.get_connection()
                                            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                            uv_id = int(selected_nv['id'])
                                            c.execute("SELECT * FROM ung_vien WHERE id = %s", (uv_id,))
                                            uv = c.fetchone()
                                            db.close()
                                            
                                            if uv:
                                                # Lưu thông tin ứng viên vào session_state
                                                st.session_state['chuyen_uv_id'] = uv_id
                                                st.session_state['chuyen_uv_data'] = {
                                                    'ho_ten': uv['ho_ten'],
                                                    'vi_tri': uv['vi_tri_du_tuyen'],
                                                    'dien_thoai': uv['dien_thoai'],
                                                    'ngay_sinh': uv['ngay_sinh'],
                                                    'gioi_tinh': uv['gioi_tinh'],
                                                    'ngay_vao_lam': uv['ngay_vao_lam'] or date.today(),
                                                    'ghi_chu': uv['luong_bao_hiem']
                                                }
                                                st.session_state['show_chuyen_nv_form'] = True
                                                st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi: {e}")
                else:
                    # Viewer: chỉ hiển thị bảng
                    st.dataframe(df_show, width='stretch', hide_index=True, height=400)
            else:
                st.info("Không có dữ liệu")
    
    # Form sửa ứng viên (chỉ admin)
    if st.session_state.get('edit_uv_id') and st.session_state.role == "admin":
        st.divider()
        st.subheader(f"✏️ Sửa ứng viên")
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT * FROM ung_vien WHERE id = %s", (st.session_state['edit_uv_id'],))
        uv_data = c.fetchone()
        db.close()
        if uv_data:
            with st.form("edit_uv_direct"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    ho_ten_e = st.text_input("Họ và tên *", value=uv_data['ho_ten'] or '')
                    vi_tri_e = st.selectbox("Vị trí dự tuyển", [""] + ds_vi_tri,
                        index=([""] + ds_vi_tri).index(uv_data['vi_tri_du_tuyen']) if uv_data['vi_tri_du_tuyen'] in ds_vi_tri else 0)
                    dien_thoai_e = st.text_input("SĐT", value=uv_data['dien_thoai'] or '')
                with col2:
                    ngay_sinh_e = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(uv_data['ngay_sinh']))
                    gioi_tinh_e = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"],
                        index=["", "Nam", "Nữ", "Khác"].index(uv_data['gioi_tinh']) if uv_data['gioi_tinh'] in ["Nam", "Nữ", "Khác"] else 0)
                with col3:
                    ngay_vao_lam_e = st.text_input("Ngày vào làm (dd/mm/yyyy)", value=format_date(uv_data['ngay_vao_lam']))
                    ghi_chu_e = st.text_area("Ghi chú", value=uv_data['luong_bao_hiem'] or '')
                    trang_thai_e = st.selectbox("Trạng thái", ["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"],
                        index=["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"].index(uv_data['trang_thai']) if uv_data['trang_thai'] in ["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"] else 0)
                
                col_save, col_del, col_cancel = st.columns(3)
                with col_save:
                    if st.form_submit_button("💾 CẬP NHẬT"):
                        if not can_edit():
                            st.error("❌ Bạn không có quyền thực hiện thao tác này!")
                        else:
                            ngay_loi = []
                            if ngay_sinh_e and not parse_date(ngay_sinh_e): 
                                ngay_loi.append("Ngày sinh")
                            if ngay_vao_lam_e and not parse_date(ngay_vao_lam_e): 
                                ngay_loi.append("Ngày vào làm")
                            if ngay_loi:
                                st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                            else:
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                c.execute("""UPDATE ung_vien SET ho_ten=%s, vi_tri_du_tuyen=%s, dien_thoai=%s,
                                    ngay_sinh=%s, gioi_tinh=%s, ngay_vao_lam=%s, luong_bao_hiem=%s, trang_thai=%s
                                    WHERE id=%s""",
                                    (ho_ten_e, vi_tri_e, dien_thoai_e, parse_date(ngay_sinh_e), gioi_tinh_e,
                                     parse_date(ngay_vao_lam_e), ghi_chu_e, trang_thai_e, uv_data['id']))
                                db.commit()
                                db.close()
                                st.success("✅ Đã cập nhật!")
                                st.cache_data.clear()
                                del st.session_state['edit_uv_id']
                                st.rerun()

                with col_del:
                    if st.form_submit_button("🗑️ XÓA"):
                        if not can_delete():
                            st.error("❌ Bạn không có quyền xóa dữ liệu!")
                        else:
                            db = st.session_state.db_engine.get_connection()
                            c = db.cursor()
                            c.execute("DELETE FROM ung_vien WHERE id = %s", (uv_data['id'],))
                            db.commit()
                            db.close()
                            st.success("🗑️ Đã xóa!")
                            st.cache_data.clear()
                            del st.session_state['edit_uv_id']
                            st.rerun()
                with col_cancel:
                    if st.form_submit_button("❌ HỦY"):
                        del st.session_state['edit_uv_id']
                        st.rerun()
    
    # Quản lý danh mục vị trí dự tuyển (chỉ admin) - bảng RIÊNG chuc_danh_ung_vien,
    # độc lập với danh mục chức danh Nhân viên (vi_tri_cong_tac)
    if st.session_state.role == "admin":
        st.divider()
        with st.expander("⚙️ Quản lý danh mục Vị trí dự tuyển (riêng cho Ứng viên)", expanded=False):
            st.caption("Danh mục này độc lập với danh mục Chức danh của Nhân viên — "
                       "đổi chức danh Nhân viên sẽ không ảnh hưởng đến danh mục và dữ liệu Ứng viên.")
            with st.form("add_vi_tri_uv"):
                ten_vt_moi = st.text_input("Tên vị trí dự tuyển mới *")
                if st.form_submit_button("➕ Thêm"):
                    if ten_vt_moi:
                        db = st.session_state.db_engine.get_connection()
                        c = db.cursor()
                        c.execute("SELECT COUNT(*) FROM chuc_danh_ung_vien WHERE ten_chuc_danh = %s", (ten_vt_moi,))
                        if c.fetchone()[0] == 0:
                            c.execute("INSERT INTO chuc_danh_ung_vien (ten_chuc_danh) VALUES (%s)", (ten_vt_moi,))
                            db.commit()
                            st.success(f"✅ Đã thêm: {ten_vt_moi}")
                            st.rerun()
                        else:
                            st.warning("Vị trí này đã tồn tại!")
                        db.close()
                    else:
                        st.error("Tên không được để trống!")
            
            db = st.session_state.db_engine.get_connection()
            c = db.cursor()
            c.execute("SELECT id, ten_chuc_danh FROM chuc_danh_ung_vien ORDER BY ten_chuc_danh")
            ds_vt = c.fetchall()
            db.close()
            if ds_vt:
                st.caption("📋 Danh sách vị trí dự tuyển:")
                for row in ds_vt:
                    col_ten, col_xoa = st.columns([4, 1])
                    with col_ten:
                        st.write(f"- {row[1]}")
                    with col_xoa:
                        if st.button("🗑️", key=f"xoa_cdv_{row[0]}"):
                            db = st.session_state.db_engine.get_connection()
                            c = db.cursor()
                            c.execute("DELETE FROM chuc_danh_ung_vien WHERE id=%s", (row[0],))
                            db.commit(); db.close()
                            st.success("🗑️ Đã xóa!")
                            st.rerun()

# ========== NHÂN VIÊN ==========
elif menu == "✅ Nhân viên":
    st.title("✅ Quản lý nhân viên")
    ensure_qdns_columns()
    ensure_qdns_table()
    ensure_mau_dieu_hop_dong_table()

    tab_dang_lam, tab_da_nghi, tab_qtct, tab_qdns, tab_co_cau = st.tabs(["📌 ĐANG LÀM VIỆC", "📋 ĐÃ NGHỈ VIỆC", "📜 LỊCH SỬ CÔNG TÁC", "📜 QUYẾT ĐỊNH NHÂN SỰ", "🏢 CƠ CẤU PHÒNG BAN"])
    
    with tab_dang_lam:
        st.caption("👥 Danh sách nhân viên đang làm việc (bao gồm thử việc)")
        # Xử lý yêu cầu reset ô tìm kiếm (đến từ nút "Đóng" của card thông tin nhân viên)
        # Phải làm TRƯỚC khi widget text_input được khởi tạo, nếu không Streamlit sẽ báo lỗi
        if st.session_state.pop('_reset_snv_dang_lam', False):
            st.session_state['snv_dang_lam'] = ''
        if st.session_state.role == "admin":
            st.session_state.setdefault('add_nv_reset_ctr', 0)
            with st.expander("➕ THÊM NHÂN VIÊN MỚI", expanded=False, key=f"add_nv_expander_{st.session_state.add_nv_reset_ctr}"):
                with st.form(f"add_nv_{st.session_state.add_nv_reset_ctr}"):
                    db = st.session_state.db_engine.get_connection()
                    c = db.cursor()
                    c.execute("SELECT DISTINCT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
                    dcv = [row[0] for row in c.fetchall()]
                    db.close()
                    dpb = get_phong_ban_options()
                    st.caption("📝 Thông tin cá nhân")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        htn = st.text_input("Họ và tên *", key="htn")
                        nsn = st.text_input("Ngày sinh (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10, key="nsn")
                        gtn = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"], key="gtn")
                        scc = st.text_input("CCCD", key="scc")
                        ncc = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10, key="ncc")
                        ncc2 = st.text_input("Nơi cấp CCCD", key="ncc2")
                    with c2:
                        nqn = st.text_input("Nguyên quán", key="nqn")
                        ttn = st.text_input("Thường trú", key="ttn")
                        qtn = st.text_input("Quốc tịch", value="Việt Nam", key="qtn")
                        dtn = st.text_input("Dân tộc", value="Kinh", key="dtn")
                        so_luong_npt = st.number_input("Số người phụ thuộc", min_value=0, value=0, step=1, key="so_luong_npt_add")
                        trinh_do_moi = st.selectbox("Trình độ", [""] + TRINH_DO_LIST, key="trinh_do_add")
                    with c3:
                        dtn2 = st.text_input("SĐT", key="dtn2")
                        emn = st.text_input("Email", key="emn")
                        cdn = st.selectbox("Chức danh", [""] + dcv, key="cdn")
                        pbn = st.selectbox("Phòng ban", [""] + dpb, key="pbn")
                        pbn_chuan = chuan_hoa_ten_phong_ban(pbn)
                        nlv = st.text_input("Nơi làm việc", value="Cảng THQT Hòn La", key="nlv")
                        anh_ho_so_moi = st.file_uploader("Ảnh hồ sơ", type=["png", "jpg", "jpeg"], key="anh_ho_so_add")
                    st.divider()
                    st.caption("💼 Hợp đồng & BHXH")
                    c4, c5, c6 = st.columns(3)
                    # ... (các trường c4, c5, c6 giữ nguyên)
                    with c4:
                        lhd = st.selectbox("Loại HĐ *", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"], key="lhd")
                        nvl = st.text_input("Ngày vào làm (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10, key="nvl")
                        nkt = st.text_input("Ngày kết thúc", placeholder="dd/mm/yyyy", max_chars=10, key="nkt")
                        mbh = st.text_input("Mã BHXH", key="mbh")
                        tbd = st.text_input("Bắt đầu BH (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10, key="tbd")
                    with c5:
                        lbh = st.text_input("Lương BH", key="lbh")
                        hsl = st.text_input("Hệ số lương", key="hsl")
                        pcv = st.text_input("PC chức vụ", key="pcv")
                        ptv = st.text_input("PC TNVK (%)", key="ptv")
                        ptn = st.text_input("PC TNN (%)", key="ptn")
                    with c6:
                        mhb = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"], key="mhb")
                        tld = st.text_input("Tỷ lệ đóng (%)", key="tld")
                        mtd = st.text_input("Mức tiền đóng", key="mtd")
                        ptd = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"], key="ptd")
                        nbh = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"], key="nbh")
                    st.divider()
                    st.caption("🏦 Ngân hàng & KCB")
                    c7, c8, c9 = st.columns(3)
                    with c7:
                        stk = st.text_input("STK", key="stk")
                        bank_index = 0
                        cnh = st.selectbox("Chi nhánh NH", options=[""] + BANK_LIST, index=bank_index, key="add_cnh")
                        tkb = st.text_input("Tỉnh KCB", value="Tỉnh Quảng Trị", key="tkb")
                    with c8:
                        nkb = st.text_input("Nơi KCB", value="Bệnh viện đa khoa khu vực Bắc Quảng Trị", key="nkb")
                        ths = st.text_input("Tỉnh/TP nhận HS", value="Tỉnh Quảng Trị", key="ths")
                        phs = st.text_input("Phường/Xã nhận HS", value="Xã Phú Trạch", key="phs")
                    with c9:
                        dhs = st.text_input("Địa chỉ nhận HS", value="Công ty cổ phần Cảng Hòn La", key="dhs")
                        dks = st.selectbox("ĐK nhận sổ", ["Có", "Không"], key="dks")
                        hso = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"], key="hso")
                    
                    col_save_exit1, col_save_exit2 = st.columns(2)
                    with col_save_exit1:
                        if st.form_submit_button("💾 LƯU", width='stretch'):
                            if not can_edit():
                                st.error("❌ Bạn không có quyền thực hiện thao tác này!")
                            else:
                                if htn:
                                    ngay_loi = []
                                    if nsn and not parse_date(nsn):
                                        ngay_loi.append("Ngày sinh")
                                    # ... (các kiểm tra ngày khác)
                                    if ngay_loi:
                                        st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                                    else:
                                        try:
                                            ten_don_vi_thu_huong = generate_ten_don_vi_thu_huong(htn)
                                            db = st.session_state.db_engine.get_connection()
                                            c = db.cursor()

                                            c.execute("SELECT COALESCE(MAX(CAST(SUBSTRING(ma_nv FROM 2) AS INTEGER)), 0)+1 FROM nhan_vien WHERE ma_nv LIKE 'C%'")
                                            so_moi = c.fetchone()[0]
                                            ma_nv = f"C{so_moi:03d}"
                                            c.execute("SELECT COALESCE(MAX(STT),0)+1 FROM nhan_vien")
                                            stt_moi = c.fetchone()[0]

                                            nhl = parse_date(nvl)
                                            tbd_val = parse_date(tbd) if tbd and tbd.strip() else None

                                            if lhd == "Thử việc":
                                                ttnv = 'THU_VIEC'
                                                ttbh = 'CHUA_DONG'
                                                c.execute("""
                                                    SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0)
                                                    FROM nhan_vien
                                                    WHERE so_hdld LIKE '%/HĐTV-CHL'
                                                    AND trang_thai IN ('THU_VIEC', 'DANG_LAM')
                                                """)
                                                tv_cnt = c.fetchone()[0] + 1
                                                so_hd = f"{tv_cnt:02d}/{nhl.year}/HĐTV-CHL"
                                            else:
                                                ttnv = 'DANG_LAM'
                                                ttbh = 'DANG_DONG'
                                                if not tbd_val:
                                                    tbd_val = nhl
                                                c.execute("""
                                                    SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0)
                                                    FROM nhan_vien
                                                    WHERE so_hdld LIKE '%/HĐLĐ-CHL'
                                                    AND trang_thai = 'DANG_LAM'
                                                    AND loai_hop_dong != 'Thử việc'
                                                """)
                                                so_hd_cnt = c.fetchone()[0] or 0
                                                so_hd = f"{so_hd_cnt + 1:02d}/{nhl.year}/HĐLĐ-CHL"
                                            
                                            # Chuẩn hóa tên phòng ban
                                            pbn_chuan = chuan_hoa_ten_phong_ban(pbn)
                                            
                                            c.execute("""INSERT INTO nhan_vien (STT, ma_nv, so_hdld, ho_ten, chuc_danh_nghe, ngay_sinh, gioi_tinh,
                                                so_cccd, ngay_cap_cccd, noi_cap_cccd, nguyen_quan, thuong_tru,
                                                dien_thoai, email, email_lien_he, ho_so, luong_bao_hiem, ma_so_bhxh, ngay_vao_lam,
                                                noi_lam_viec, so_tai_khoan_nh, chi_nhanh_nh, ngay_ky_hd, loai_hop_dong,
                                                nhom_bhxh, thang_bat_dau_bh, thang_ket_thuc_bh, trang_thai, trang_thai_bhxh,
                                                phong_ban_lam_viec, ngay_ket_thuc, quoc_tich, dan_toc, he_so_luong, phu_cap_chuc_vu,
                                                phu_cap_tnvk, phu_cap_tnn, muc_huong_bhyt, ty_le_dong, muc_tien_dong, phuong_thuc_dong,
                                                tinh_nhan_hs, phuong_nhan_hs, dia_chi_nhan_hs, tinh_kcb, noi_dang_ky_kcb, dang_ky_nhan_so,
                                                ten_don_vi_thu_huong, trinh_do, so_luong_npt)
                                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                                %s, %s, %s, %s, %s, %s, %s, %s, %s) RETURNING id""",
                                                (stt_moi, ma_nv, so_hd, htn, cdn, parse_date(nsn), gtn, scc, parse_date(ncc), ncc2, nqn, ttn,
                                                 dtn2, emn, emn, hso, lbh, mbh, parse_date(nvl), nlv, stk, cnh, parse_date(nvl), lhd,
                                                 nbh, tbd_val, None, ttnv, ttbh, pbn_chuan, parse_date(nkt), qtn, dtn, 
                                                 to_float_or_none(hsl), to_float_or_none(pcv), to_float_or_none(ptv), to_float_or_none(ptn),
                                                 mhb, to_float_or_none(tld), to_float_or_none(mtd), ptd, ths, phs, dhs, tkb, nkb, dks,
                                                 ten_don_vi_thu_huong, trinh_do_moi, so_luong_npt))
                                            new_nv_id = c.fetchone()[0]
                                            
                                            if anh_ho_so_moi is not None:
                                                storage_path_anh = upload_anh_ho_so(ma_nv, htn, anh_ho_so_moi)
                                                if storage_path_anh:
                                                    c.execute("UPDATE nhan_vien SET anh_ho_so=%s WHERE id=%s", (storage_path_anh, new_nv_id))
                                            
                                            db.commit()
                                            db.close()
                                            st.success(f"✅ Đã lưu nhân viên mới thành công! {htn} - {ma_nv}")
                                            st.cache_data.clear()
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi: {e}")
                                else:
                                    st.error("Họ tên không được để trống!")
                    with col_save_exit2:
                        if st.form_submit_button("❌ THOÁT", width='stretch'):
                            # Reset tất cả các trường trong form
                            # Xóa các session state liên quan đến form thêm nhân viên
                            keys_to_clear = [
                                'htn', 'nsn', 'gtn', 'scc', 'ncc', 'ncc2', 'nqn', 'ttn', 
                                'dtn2', 'emn', 'cdn', 'pbn', 'nlv', 'lhd', 'nvl', 'nkt',
                                'mbh', 'tbd', 'lbh', 'hsl', 'pcv', 'ptv', 'ptn', 'mhb',
                                'tld', 'mtd', 'ptd', 'nbh', 'stk', 'add_cnh', 'tkb', 'nkb',
                                'ths', 'phs', 'dhs', 'dks', 'hso', 'trinh_do_moi', 'trinh_do_add',
                                'so_luong_npt_add', 'qtn', 'dtn', 'anh_ho_so_add'
                            ]
                            for key in keys_to_clear:
                                if key in st.session_state:
                                    del st.session_state[key]
                            # Đổi key của form & expander để buộc Streamlit khởi tạo lại
                            # toàn bộ widget (thu gọn expander, xóa sạch mọi giá trị đã nhập)
                            st.session_state.add_nv_reset_ctr += 1
                            st.success("✅ Đã thoát form thêm nhân viên")
                            st.rerun()                
                st.divider()

        sn = st.text_input("🔍 Tìm kiếm", key="snv_dang_lam")

        
        db_f = st.session_state.db_engine.get_connection()
        c_f = db_f.cursor()
        c_f.execute("SELECT DISTINCT chuc_danh_nghe FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != '' ORDER BY chuc_danh_nghe")
        ds_chuc_danh = [row[0] for row in c_f.fetchall()]
        c_f.execute("SELECT DISTINCT loai_hop_dong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND loai_hop_dong IS NOT NULL AND loai_hop_dong != '' ORDER BY loai_hop_dong")
        ds_loai_hd = [row[0] for row in c_f.fetchall()]
        c_f.execute("SELECT DISTINCT trinh_do FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND trinh_do IS NOT NULL AND trinh_do != '' ORDER BY trinh_do")
        ds_bang_cap = [row[0] for row in c_f.fetchall()]
        db_f.close()
        
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filter_chuc_danh = st.selectbox("🔍 Lọc Chức danh:", ["Tất cả"] + ds_chuc_danh, key="filter_cd_danglam")
        with col_f2:
            filter_loai_hd = st.selectbox("🔍 Lọc Loại HĐ:", ["Tất cả"] + ds_loai_hd, key="filter_lhd_danglam")
        with col_f3:
            filter_bang_cap = st.selectbox("🔍 Lọc theo Bằng cấp:", ["Tất cả"] + ds_bang_cap, key="filter_bc_danglam")
        
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        sql = "SELECT * FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC')"
        params = []
        if sn:
            sql += " AND (ho_ten LIKE %s OR dien_thoai LIKE %s OR so_cccd LIKE %s OR ma_nv LIKE %s)"
            params.extend([f'%{sn}%'] * 4)
        if filter_chuc_danh != "Tất cả":
            sql += " AND chuc_danh_nghe = %s"
            params.append(filter_chuc_danh)
        if filter_loai_hd != "Tất cả":
            sql += " AND loai_hop_dong = %s"
            params.append(filter_loai_hd)
        if filter_bang_cap != "Tất cả":
            sql += " AND trinh_do = %s"
            params.append(filter_bang_cap)
        sql += " ORDER BY id DESC"
        c.execute(sql, tuple(params))
        ds = c.fetchall()
        db.close()
        
        if ds:
            # ===== KIỂM TRA NẾU CHỈ CÓ 1 KẾT QUẢ TÌM KIẾM =====
            if len(ds) == 1:
                nv = ds[0]  # Lấy nhân viên duy nhất
                st.success(f"🎯 Tìm thấy 1 nhân viên: {nv['ho_ten']}")
                render_employee_info_card(
                    nv,
                    key_prefix=f"single_{nv['id']}",
                    on_close=lambda: st.session_state.update({'_reset_snv_dang_lam': True})
                )

                # Thêm tùy chọn hiển thị bảng
                st.divider()
                if st.checkbox("📊 Hiển thị danh sách đầy đủ", value=False, key="show_full_list_card"):
                    # Hiển thị bảng bên dưới
                    pass
                else:
                    # Nếu không hiển thị bảng, vẫn cần render các phần bên dưới
                    # nhưng chúng ta sẽ bỏ qua phần bảng
                    # Để không bị lỗi, chúng ta sẽ đặt một flag
                    st.session_state['skip_table_display'] = True
                    # Vẫn cần giữ các form sửa nhân viên ở phía sau
                    # nhưng chúng sẽ không hiển thị nếu không có selected_nv_id
                    pass
            
            # ===== PHẦN HIỂN THỊ BẢNG (CHẠY KHI CÓ NHIỀU KẾT QUẢ HOẶC USER CHỌN HIỂN THỊ) =====
            # Chỉ hiển thị bảng nếu có nhiều hơn 1 kết quả HOẶC user chọn hiển thị đầy đủ
            show_table = (len(ds) > 1) or (len(ds) == 1 and st.session_state.get('show_full_list_card', False))
            
            if show_table or len(ds) > 1:
                # Reset flag nếu có
                st.session_state['skip_table_display'] = False
                
                df = pd.DataFrame(ds)
                for col in df.columns:
                    if 'ngay' in col.lower():
                        df[col] = df[col].apply(format_date)
                
                if 'selected' not in df.columns:
                    df.insert(0, 'selected', False)
                
                display_cols = ['selected', 'ma_nv', 'ho_ten', 'ngay_sinh', 'gioi_tinh', 'so_hdld', 'so_cccd', 'dien_thoai',
                                'thuong_tru', 'chuc_danh_nghe', 'loai_hop_dong', 'ngay_vao_lam', 'ma_so_bhxh', 'thang_bat_dau_bh',
                                'ten_don_vi_thu_huong']
                # viewer và kt_luong: ẩn thông tin nhạy cảm (CCCD, STK ngân hàng) trên bảng danh sách
                SENSITIVE_COLS = {'so_cccd', 'so_tai_khoan_nh'}
                if st.session_state.role in ("viewer", "kt_luong"):
                    display_cols = [c for c in display_cols if c not in SENSITIVE_COLS]
                available_cols = [c for c in display_cols if c in df.columns]
                df_show = df[available_cols]
                
                col_map = {
                    'selected': 'Chọn',
                    'ma_nv': 'Mã NV',
                    'ho_ten': 'Họ và tên',
                    'ngay_sinh': 'Ngày sinh',
                    'gioi_tinh': 'Giới tính',
                    'so_hdld': 'Số HĐLĐ',
                    'so_cccd': 'CCCD',
                    'dien_thoai': 'SĐT',
                    'thuong_tru': 'Thường trú',
                    'chuc_danh_nghe': 'Chức danh',
                    'loai_hop_dong': 'Loại HĐ',
                    'ngay_vao_lam': 'Ngày vào làm',
                    'ma_so_bhxh': 'Mã số BHXH',
                    'thang_bat_dau_bh': 'Bắt đầu BH',
                    'ten_don_vi_thu_huong': 'Tên đơn vị thụ hưởng',
                }
                df_show.rename(columns=col_map, inplace=True)
                
                if len(ds) > 1:
                    st.caption(f"📌 {len(ds)} kết quả. Tick chọn 1 nhân viên để thao tác.")
                else:
                    st.caption(f"📌 Danh sách đầy đủ ({len(ds)} kết quả). Tick chọn 1 nhân viên để thao tác.")
                
                # Nếu là viewer, hiển thị bảng không có checkbox chọn
                if st.session_state.role == "admin":
                    # Xử lý yêu cầu reset lựa chọn (đến từ nút "Đóng" của card thông tin nhân viên)
                    # Phải làm TRƯỚC khi widget data_editor được khởi tạo
                    if st.session_state.pop('_reset_nv_editor_danglam', False):
                        st.session_state.pop('nv_editor_danglam', None)
                    edited_df = st.data_editor(
                        df_show,
                        column_config={
                            "Chọn": st.column_config.CheckboxColumn("Profile", default=False)
                        },
                        disabled=[col for col in df_show.columns if col != 'Chọn'],
                        hide_index=True,
                        height=400,
                        key="nv_editor_danglam"
                    )
                else:
                    # Viewer: hiển thị bảng đơn thuần, không có checkbox
                    st.dataframe(df_show.drop(columns=['Chọn'], errors='ignore'), width='stretch', hide_index=True, height=400)
                    edited_df = None
                
                selected_nv = None
                if edited_df is not None and st.session_state.role == "admin" and 'Chọn' in edited_df.columns:
                    selected_rows = edited_df[edited_df['Chọn'] == True]
                    if len(selected_rows) > 0:
                        if len(selected_rows) > 1:
                            st.error("⚠️ Chỉ được chọn 1 nhân viên!")
                        else:
                            selected_idx = selected_rows.index[0]
                            selected_nv = df.iloc[selected_idx]
                            nv_id_key = selected_nv['id']
                            
                            # Hiển thị các nút chức năng (chỉ admin mới thấy và mới click được)
                            render_employee_info_card(
                                selected_nv,
                                key_prefix=f"multi_{nv_id_key}",
                                on_close=lambda: st.session_state.update({'_reset_nv_editor_danglam': True})
                            )
                            col_btn5 = st.container()

                            with col_btn5:
                                trang_thai_nv = selected_nv.get('trang_thai', '')
                                if trang_thai_nv == 'THU_VIEC':
                                    if f'convert_open_{nv_id_key}' not in st.session_state:
                                        st.session_state[f'convert_open_{nv_id_key}'] = False
                                    
                                    if not st.session_state[f'convert_open_{nv_id_key}']:
                                        if st.button(f"🔄 CHUYỂN HĐLĐ KHÔNG XĐTH - {selected_nv['ho_ten']}", 
                                                    key=f"convert_hdld_btn_{nv_id_key}", 
                                                    width='stretch', type="primary"):
                                            st.session_state[f'convert_open_{nv_id_key}'] = True
                                            st.rerun()
                                    else:
                                        st.markdown("---")
                                        st.markdown("### 📝 CHUYỂN ĐỔI HỢP ĐỒNG LAO ĐỘNG")
                                        st.caption("Vui lòng nhập đầy đủ thông tin cho quyết định chuyển đổi")
                                        
                                        db_temp = st.session_state.db_engine.get_connection()
                                        c_temp = db_temp.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                        c_temp.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                        nv_data = c_temp.fetchone()
                                        db_temp.close()
                                        
                                        if nv_data:
                                            ngay_quyet_dinh = st.date_input(
                                                "📅 Ngày quyết định:", 
                                                value=date.today(),
                                                key=f"ngay_qd_{nv_id_key}"
                                            )
                                            
                                            current_year = datetime.now().year
                                            db_temp2 = st.session_state.db_engine.get_connection()
                                            c_temp2 = db_temp2.cursor()
                                            pattern = f'%/{current_year}/HĐLĐ-CHL'
                                            c_temp2.execute("""
                                                SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) as max_stt
                                                FROM nhan_vien 
                                                WHERE so_hdld LIKE %s 
                                                AND trang_thai = 'DANG_LAM'
                                                AND loai_hop_dong != 'Thử việc'
                                            """, (f'%/{current_year}/HĐLĐ-CHL',))
                                            result = c_temp2.fetchone()
                                            max_stt = result[0] if result else 0
                                            db_temp2.close()
                                            
                                            next_stt = max_stt + 1
                                            stt_str = str(next_stt).zfill(2)
                                            so_hd_moi = f"{stt_str}/{current_year}/HĐLĐ-CHL"
                                            
                                            st.info(f"📄 **Số HĐLĐ mới:** {so_hd_moi} (tự động sinh)")
                                            
                                            ngay_hieu_luc = st.date_input(
                                                "📅 Ngày hiệu lực (bắt đầu HĐLĐ):", 
                                                value=ngay_quyet_dinh,
                                                key=f"ngay_hl_{nv_id_key}"
                                            )
                                            
                                            ngay_bat_dau_bh = st.date_input(
                                                "📅 Ngày bắt đầu đóng BHXH:", 
                                                value=ngay_hieu_luc,
                                                key=f"ngay_bhxh_{nv_id_key}",
                                                help="⚠️ Ngày bắt đầu tham gia BHXH. Thường là ngày hiệu lực HĐLĐ chính thức."
                                            )
                                            
                                            ly_do_chuyen = st.text_area(
                                                "📝 Lý do/ Nội dung quyết định:", 
                                                value="Hoàn thành thời gian thử việc, chuyển sang hợp đồng lao động không xác định thời hạn",
                                                key=f"ly_do_{nv_id_key}",
                                                height=80
                                            )
                                            
                                            col_confirm1, col_confirm2, col_confirm3 = st.columns([1, 2, 1])
                                            with col_confirm2:
                                                if st.button("✅ XÁC NHẬN CHUYỂN ĐỔI", key=f"confirm_convert_{nv_id_key}", width='stretch', type="primary"):
                                                    try:
                                                        db = st.session_state.db_engine.get_connection()
                                                        c = db.cursor()
                                                        
                                                        so_hd_tv_cu = selected_nv.get('so_hdld', '')
                                                        ngay_vao_lam_cu = selected_nv.get('ngay_vao_lam')
                                                        
                                                        if ngay_vao_lam_cu:
                                                            if hasattr(ngay_vao_lam_cu, 'strftime'):
                                                                pass
                                                            else:
                                                                ngay_vao_lam_cu = parse_date(ngay_vao_lam_cu)
                                                                if not ngay_vao_lam_cu:
                                                                    ngay_vao_lam_cu = date.today()
                                                        else:
                                                            ngay_vao_lam_cu = date.today()
                                                        
                                                        current_year = datetime.now().year
                                                        c.execute("""
                                                            SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) as max_stt
                                                            FROM nhan_vien 
                                                            WHERE so_hdld LIKE %s 
                                                            AND trang_thai = 'DANG_LAM'
                                                            AND loai_hop_dong != 'Thử việc'
                                                        """, (f'%/{current_year}/HĐLĐ-CHL',))
                                                        result = c.fetchone()
                                                        max_stt = result[0] if result else 0
                                                        next_stt = max_stt + 1
                                                        stt_str = str(next_stt).zfill(2)
                                                        so_hd_moi = f"{stt_str}/{current_year}/HĐLĐ-CHL"
                                                        
                                                        c.execute("""
                                                            UPDATE nhan_vien SET 
                                                                trang_thai = 'DANG_LAM',
                                                                loai_hop_dong = 'Không xác định thời hạn',
                                                                so_hdld = %s,
                                                                ngay_ky_hd = %s,
                                                                ngay_chinh_thuc = %s,
                                                                thang_bat_dau_bh = %s,
                                                                trang_thai_bhxh = 'DANG_DONG',
                                                                ngay_ket_thuc = NULL
                                                            WHERE id = %s
                                                        """, (so_hd_moi, ngay_quyet_dinh, ngay_hieu_luc, ngay_bat_dau_bh, int(selected_nv['id'])))
                                                        
                                                        c.execute("""
                                                            INSERT INTO quyet_dinh_nhan_su (
                                                                nhan_vien_id, loai_quyet_dinh, ngay_quyet_dinh, ngay_hieu_luc,
                                                                noi_dung, so_quyet_dinh, loai_hop_dong_cu, loai_hop_dong_moi,
                                                                he_so_luong_cu, he_so_luong_moi, so_hd_cu
                                                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                                        """, (
                                                            int(selected_nv['id']),
                                                            'CHINH_THUC',
                                                            ngay_quyet_dinh,
                                                            ngay_hieu_luc,
                                                            ly_do_chuyen,
                                                            f"QD{ngay_quyet_dinh.strftime('%Y%m%d')}_{selected_nv['ma_nv']}",
                                                            nv_data.get('loai_hop_dong', 'Thử việc'),
                                                            'Không xác định thời hạn',
                                                            nv_data.get('he_so_luong', 0),
                                                            nv_data.get('he_so_luong', 0),
                                                            so_hd_tv_cu
                                                        ))
                                                        
                                                        c.execute("""
                                                            UPDATE lich_su_cong_tac 
                                                            SET den_ngay = %s,
                                                                so_hop_dong = %s
                                                            WHERE nhan_vien_id = %s 
                                                            AND loai_hop_dong = 'Thử việc'
                                                            AND den_ngay IS NULL
                                                        """, (ngay_hieu_luc - timedelta(days=1), so_hd_tv_cu, int(selected_nv['id'])))
                                                        
                                                        c.execute("""
                                                            INSERT INTO lich_su_cong_tac (
                                                                nhan_vien_id, tu_ngay, chuc_danh, phong_ban, 
                                                                noi_lam_viec, loai_hop_dong, he_so_luong, so_hop_dong
                                                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                                                        """, (
                                                            int(selected_nv['id']),
                                                            ngay_hieu_luc,
                                                            nv_data.get('chuc_danh_nghe', ''),
                                                            nv_data.get('phong_ban_lam_viec', ''),
                                                            nv_data.get('noi_lam_viec', 'Cảng THQT Hòn La'),
                                                            'Không xác định thời hạn',
                                                            nv_data.get('he_so_luong', 0),
                                                            so_hd_moi
                                                        ))
                                                        
                                                        db.commit()
                                                        db.close()
                                                        
                                                        st.success(f"✅ Đã chuyển {nv_data['ho_ten']} sang HĐLĐ không xác định thời hạn!")
                                                        st.info(f"📄 Số HĐTV cũ: {so_hd_tv_cu}")
                                                        st.info(f"📄 Số HĐLĐ mới: {so_hd_moi}")
                                                        st.cache_data.clear()
                                                        st.session_state[f'convert_open_{nv_id_key}'] = False
                                                        st.rerun()
                                                        
                                                    except Exception as e:
                                                        db.rollback()
                                                        db.close()
                                                        st.error(f"❌ Lỗi: {str(e)}")

                                            if st.button("❌ HỦY", key=f"cancel_convert_{nv_id_key}", width='stretch'):
                                                st.session_state[f'convert_open_{nv_id_key}'] = False
                                                st.rerun()

                            st.divider()
            
            # Form sửa nhân viên (chỉ admin)
            if 'selected_nv_id' in st.session_state and st.session_state.selected_nv_id is not None and st.session_state.role == "admin":
                try:
                    nid = int(st.session_state['selected_nv_id'])
                    db = st.session_state.db_engine.get_connection()
                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c.execute("SELECT * FROM nhan_vien WHERE id=%s", (nid,))
                    nd = c.fetchone()
                    db.close()
                    db_cd = st.session_state.db_engine.get_connection()
                    c_cd = db_cd.cursor()
                    c_cd.execute("SELECT DISTINCT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
                    dcv_edit = [row[0] for row in c_cd.fetchall()]
                    db_cd.close()
                    dpb_edit = get_phong_ban_options()
                    
                    if nd:
                        st.subheader(f"✏️ Cập nhật: {nd.get('ho_ten', '')} ({nd.get('ma_nv', '')})")
                        with st.form("edit_nv"):
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                hnv = st.text_input("Họ và tên *", value=nd.get('ho_ten', ''))
                                nsnv = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(nd.get('ngay_sinh')), placeholder="dd/mm/yyyy", max_chars=10)
                                gtnv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"], index=["", "Nam", "Nữ", "Khác"].index(nd.get('gioi_tinh', '')) if nd.get('gioi_tinh') in ["Nam", "Nữ", "Khác"] else 0)
                                sccv = st.text_input("CCCD", value=nd.get('so_cccd', ''))
                                nccv = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", value=format_date(nd.get('ngay_cap_cccd')), placeholder="dd/mm/yyyy", max_chars=10)
                                ncv = st.text_input("Nơi cấp CCCD", value=nd.get('noi_cap_cccd', ''))
                                dtnv2 = st.text_input("SĐT", value=nd.get('dien_thoai', ''))
                            with col2:
                                nqnv = st.text_input("Nguyên quán", value=nd.get('nguyen_quan', ''))
                                ttnv = st.text_input("Thường trú", value=nd.get('thuong_tru', ''))
                                qtnv = st.text_input("Quốc tịch", value=nd.get('quoc_tich', 'Việt Nam'))
                                dtnv = st.text_input("Dân tộc", value=nd.get('dan_toc', 'Kinh'))
                                so_luong_npt_edit = st.number_input("Số người phụ thuộc", min_value=0, value=int(nd.get('so_luong_npt') or 0), step=1, key=f"so_luong_npt_edit_{nid}")
                                trinh_do_v = st.selectbox("Trình độ", [""] + TRINH_DO_LIST, index=([""] + TRINH_DO_LIST).index(nd.get('trinh_do', '')) if nd.get('trinh_do') in TRINH_DO_LIST else 0)
                                cdnv = st.selectbox("Chức danh", [""] + dcv_edit, index=([""] + dcv_edit).index(nd.get('chuc_danh_nghe', '')) if nd.get('chuc_danh_nghe') in dcv_edit else 0)
                            with col3:
                                pb_hien_tai_chuan = chuan_hoa_ten_phong_ban(nd.get('phong_ban_lam_viec'))
                                pbnv = st.selectbox("Phòng ban", [""] + dpb_edit, index=([""] + dpb_edit).index(pb_hien_tai_chuan) if pb_hien_tai_chuan in dpb_edit else 0)
                                nlv2 = st.text_input("Nơi làm việc", value=nd.get('noi_lam_viec', 'Cảng THQT Hòn La'))
                                emnv = st.text_input("Email", value=nd.get('email_lien_he', ''))
                                anh_hien_tai = nd.get('anh_ho_so')
                                if anh_hien_tai:
                                    anh_bytes_ht = get_anh_ho_so_bytes(anh_hien_tai)
                                    if anh_bytes_ht:
                                        st.image(anh_bytes_ht, caption="Ảnh hồ sơ hiện tại", width=120)
                                anh_ho_so_v = st.file_uploader("Đổi ảnh hồ sơ (bỏ trống nếu giữ nguyên)", key=f"anh_ho_so_edit_{nid}")
                            
                            st.divider()
                            st.caption("💼 Hợp đồng & BHXH")
                            col4, col5, col6 = st.columns(3)
                            with col4:
                                lhdv = st.selectbox("Loại HĐ", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"], index=["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"].index(nd.get('loai_hop_dong', 'Thử việc')) if nd.get('loai_hop_dong') in ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"] else 0)
                                nvlv = st.text_input("Ngày vào làm (dd/mm/yyyy)", value=format_date(nd.get('ngay_vao_lam')), placeholder="dd/mm/yyyy", max_chars=10)
                                nktv = st.text_input("Ngày kết thúc (dd/mm/yyyy)", value=format_date(nd.get('ngay_ket_thuc')), placeholder="dd/mm/yyyy", max_chars=10)
                                mbhv = st.text_input("Mã BHXH", value=nd.get('ma_so_bhxh', ''))
                                tbdv = st.text_input("Bắt đầu BH (dd/mm/yyyy)", value=format_date(nd.get('thang_bat_dau_bh')), placeholder="dd/mm/yyyy", max_chars=10)
                            with col5:
                                lbhv = st.text_input("Lương BH", value=nd.get('luong_bao_hiem', ''))
                                hslv = st.text_input("Hệ số lương", value=str(nd.get('he_so_luong', '')))
                                pcvv = st.text_input("PC chức vụ", value=str(nd.get('phu_cap_chuc_vu', '')))
                                ptvv = st.text_input("PC TNVK (%)", value=str(nd.get('phu_cap_tnvk', '')))
                                ptnv = st.text_input("PC TNN (%)", value=str(nd.get('phu_cap_tnn', '')))
                            with col6:
                                mhbv = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"], index=["80%", "95%", "100%"].index(nd.get('muc_huong_bhyt', '80%')) if nd.get('muc_huong_bhyt') in ["80%", "95%", "100%"] else 0)
                                tldv = st.text_input("Tỷ lệ đóng (%)", value=str(nd.get('ty_le_dong', '')))
                                mtdv = st.text_input("Mức tiền đóng", value=str(nd.get('muc_tien_dong', '')))
                                ptdv = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"], index=["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"].index(nd.get('phuong_thuc_dong', 'Hàng tháng')) if nd.get('phuong_thuc_dong') in ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"] else 0)
                                nbhv = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"], index=["", "Văn phòng", "Lao động trực tiếp"].index(nd.get('nhom_bhxh', '')) if nd.get('nhom_bhxh') in ["Văn phòng", "Lao động trực tiếp"] else 0)
                            
                            st.divider()
                            st.caption("🏦 Ngân hàng & KCB")
                            col7, col8, col9 = st.columns(3)
                            with col7:
                                stkv = st.text_input("STK", value=nd.get('so_tai_khoan_nh', ''))
                                # Tạo dropdown cho chi nhánh ngân hàng
                                bank_edit_index = 0
                                old_bank = nd.get('chi_nhanh_nh', '')
                                if old_bank in BANK_LIST:
                                    bank_edit_index = BANK_LIST.index(old_bank) + 1
                                cnhv = st.selectbox("Chi nhánh NH", options=[""] + BANK_LIST, index=bank_edit_index, key="edit_cnh")
                                tkbv = st.text_input("Tỉnh KCB", value=nd.get('tinh_kcb', ''))
                            with col8:
                                nkbv = st.text_input("Nơi KCB", value=nd.get('noi_dang_ky_kcb', ''))
                                thsv = st.text_input("Tỉnh/TP nhận HS", value=nd.get('tinh_nhan_hs', ''))
                                phsv = st.text_input("Phường/Xã nhận HS", value=nd.get('phuong_nhan_hs', ''))
                            with col9:
                                dhsv = st.text_input("Địa chỉ nhận HS", value=nd.get('dia_chi_nhan_hs', ''))
                                dksv = st.selectbox("ĐK nhận sổ", ["Có", "Không"], index=["Có", "Không"].index(nd.get('dang_ky_nhan_so', 'Có')) if nd.get('dang_ky_nhan_so') in ["Có", "Không"] else 0)
                                hsov = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"], index=["", "Đã có HS", "Chưa có"].index(nd.get('ho_so', '')) if nd.get('ho_so') in ["Đã có HS", "Chưa có"] else 0)
                            col_save, col_cancel = st.columns(2)
                            with col_save:
                                if st.form_submit_button("💾 CẬP NHẬT", width='stretch'):
                                    if not can_edit():
                                        st.error("❌ Bạn không có quyền thực hiện thao tác này!")
                                    else:
                                        if hnv:
                                            ngay_loi = []
                                            if nsnv and not parse_date(nsnv):
                                                ngay_loi.append("Ngày sinh")
                                            if nccv and not parse_date(nccv):
                                                ngay_loi.append("Ngày cấp CCCD")
                                            if nvlv and not parse_date(nvlv):
                                                ngay_loi.append("Ngày vào làm")
                                            if nktv and not parse_date(nktv):
                                                ngay_loi.append("Ngày kết thúc")
                                            if tbdv and not parse_date(tbdv):
                                                ngay_loi.append("Bắt đầu BH")
                                            if ngay_loi:
                                                st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                                            else:
                                                try:
                                                    ten_don_vi_thu_huong = generate_ten_don_vi_thu_huong(hnv)
                                                    
                                                    db_upd = st.session_state.db_engine.get_connection()
                                                    c_upd = db_upd.cursor()
                                                    nhl = parse_date(nvlv) or date.today()
                                                    ngay_bat_dau_bh = parse_date(tbdv) if tbdv and tbdv.strip() else None
                                                    if lhdv == "Thử việc":
                                                        tt_nv, tt_bh, tbd_val = 'THU_VIEC', 'CHUA_DONG', None
                                                    else:
                                                        tt_nv, tt_bh = 'DANG_LAM', 'DANG_DONG'
                                                        if ngay_bat_dau_bh:
                                                            tbd_val = ngay_bat_dau_bh
                                                        else:
                                                            tbd_val = parse_date(nvlv)
                                                    
                                                    # Chuẩn hóa tên phòng ban
                                                    pbnv_chuan = chuan_hoa_ten_phong_ban(pbnv)
                                                    
                                                    c_upd.execute("""UPDATE nhan_vien SET ho_ten=%s,chuc_danh_nghe=%s,ngay_sinh=%s,gioi_tinh=%s,
                                                        so_cccd=%s,ngay_cap_cccd=%s,noi_cap_cccd=%s,nguyen_quan=%s,thuong_tru=%s,dien_thoai=%s,
                                                        email=%s,email_lien_he=%s,ho_so=%s,luong_bao_hiem=%s,ma_so_bhxh=%s,ngay_vao_lam=%s,noi_lam_viec=%s,
                                                        so_tai_khoan_nh=%s,chi_nhanh_nh=%s,ngay_ky_hd=%s,loai_hop_dong=%s,nhom_bhxh=%s,
                                                        thang_bat_dau_bh=%s,trang_thai=%s,trang_thai_bhxh=%s,phong_ban_lam_viec=%s,
                                                        ngay_ket_thuc=%s,quoc_tich=%s,dan_toc=%s,he_so_luong=%s,phu_cap_chuc_vu=%s,
                                                        phu_cap_tnvk=%s,phu_cap_tnn=%s,muc_huong_bhyt=%s,ty_le_dong=%s,muc_tien_dong=%s,
                                                        phuong_thuc_dong=%s,tinh_nhan_hs=%s,phuong_nhan_hs=%s,dia_chi_nhan_hs=%s,
                                                        tinh_kcb=%s,noi_dang_ky_kcb=%s,dang_ky_nhan_so=%s, ten_don_vi_thu_huong=%s, trinh_do=%s,
                                                        so_luong_npt=%s WHERE id=%s""",
                                                        (hnv, cdnv, parse_date(nsnv), gtnv, sccv, parse_date(nccv), ncv, nqnv, ttnv, dtnv2,
                                                         emnv, emnv, hsov, lbhv, mbhv, parse_date(nvlv), nlv2, stkv, cnhv, parse_date(nvlv), lhdv,
                                                         nbhv, tbd_val, tt_nv, tt_bh, pbnv_chuan, parse_date(nktv), qtnv, dtnv,
                                                         to_float_or_none(hslv), to_float_or_none(pcvv), to_float_or_none(ptvv), to_float_or_none(ptnv),
                                                         mhbv, to_float_or_none(tldv), to_float_or_none(mtdv), ptdv, thsv, phsv, dhsv,
                                                         tkbv, nkbv, dksv, ten_don_vi_thu_huong, trinh_do_v, so_luong_npt_edit, nid))
                                                    
                                                    if anh_ho_so_v is not None:
                                                        storage_path_anh_v = upload_anh_ho_so(nd.get('ma_nv', nid), hnv, anh_ho_so_v)
                                                        if storage_path_anh_v:
                                                            c_upd.execute("UPDATE nhan_vien SET anh_ho_so=%s WHERE id=%s", (storage_path_anh_v, nid))
                                                    
                                                    db_upd.commit()
                                                    db_upd.close()
                                                    st.success(f"✅ Đã cập nhật: {hnv}")
                                                    st.cache_data.clear()
                                                    del st.session_state['selected_nv_id']
                                                    st.rerun()
                                                except Exception as e:
                                                    st.error(f"❌ Lỗi: {e}")
                                        else:
                                            st.error("Họ tên không được để trống!")

                            with col_cancel:
                                if st.form_submit_button("❌ HỦY SỬA", width='stretch'):
                                    # Xóa session state để đóng form sửa
                                    if 'selected_nv_id' in st.session_state:
                                        del st.session_state['selected_nv_id']
                                    st.success("✅ Đã hủy sửa nhân viên")
                                    st.rerun()
            
                except Exception as e:
                    st.error(f"Lỗi khi tải thông tin nhân viên: {e}")
                    st.session_state.selected_nv_id = None
                    st.rerun()
            
            # Form nhập thông tin hộ gia đình (chỉ admin) - Đặt NGOÀI form sửa nhân viên
            if 'bhxh_family_nv_id' in st.session_state and st.session_state.bhxh_family_nv_id is not None and st.session_state.role == "admin":
                nv_id = st.session_state['bhxh_family_nv_id']
                nv_name = st.session_state['bhxh_family_nv_name']
                st.divider()
                st.subheader(f"🏠 NHẬP THÔNG TIN HỘ GIA ĐÌNH CHO: {nv_name}")
                st.caption("Vui lòng nhập đầy đủ thông tin chủ hộ và các thành viên trong hộ gia đình")
                
                db = st.session_state.db_engine.get_connection()
                c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                c.execute("SELECT * FROM nhan_vien WHERE id = %s", (nv_id,))
                nv_data = c.fetchone()
                db.close()
                
                if 'bhxh_family_members' not in st.session_state:
                    db_temp = st.session_state.db_engine.get_connection()
                    c_temp = db_temp.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c_temp.execute("SELECT * FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_id,))
                    existing_members = c_temp.fetchall()
                    db_temp.close()
                    st.session_state['bhxh_family_members'] = []
                    for tv in existing_members:
                        st.session_state['bhxh_family_members'].append({
                            'ho_ten': tv['ho_ten'], 'ngay_sinh': tv['ngay_sinh'], 'gioi_tinh': tv['gioi_tinh'],
                            'quoc_tich': tv['quoc_tich'], 'dan_toc': tv['dan_toc'], 'quan_he': tv['quan_he_voi_chu_ho'],
                            'tinh': tv['tinh_thanh_pho'], 'phuong_xa': tv['phuong_xa']
                        })
                
                db_temp = st.session_state.db_engine.get_connection()
                c_temp = db_temp.cursor()
                c_temp.execute("SELECT ma_tinh, ten_tinh FROM danh_muc_tinh ORDER BY ten_tinh")
                ds_tinh = c_temp.fetchall()
                db_temp.close()
                tinh_options = {ten: ma for ma, ten in ds_tinh}
                
                if st.session_state.bhxh_family_members:
                    st.markdown("**Danh sách thành viên đã thêm:**")
                    tv_data = []
                    for i, tv in enumerate(st.session_state.bhxh_family_members):
                        tv_data.append({"STT": i+1, "Họ và tên": tv['ho_ten'], "Ngày sinh": format_date(tv['ngay_sinh']),
                                        "Giới tính": tv['gioi_tinh'], "Quốc tịch": tv['quoc_tich'], "Dân tộc": tv['dan_toc'],
                                        "Quan hệ chủ hộ": tv['quan_he'], "Tỉnh/TP": tv['tinh'], "Phường/Xã": tv['phuong_xa']})
                    df_tv = pd.DataFrame(tv_data)
                    st.dataframe(df_tv, width='stretch', hide_index=True)
                    col_del1, col_del2, col_del3 = st.columns([1,1,1])
                    with col_del2:
                        tv_to_delete = st.number_input("Nhập STT thành viên cần xóa:", min_value=1, max_value=len(st.session_state.bhxh_family_members), step=1, key="tv_delete_family")
                        if st.button("🗑️ Xóa thành viên", key="btn_del_tv_family"):
                            st.session_state.bhxh_family_members.pop(tv_to_delete - 1)
                            st.rerun()
                
                with st.form(key=f"bhxh_family_form_{nv_id}"):
                    st.markdown("**I. THÔNG TIN CHỦ HỘ:**")
                    col1, col2 = st.columns(2)
                    with col1:
                        ho_ten_chu_ho = st.text_input("Họ và tên chủ hộ", value=nv_data.get('ho_ten_chu_ho', '') if nv_data else '')
                        so_cccd_chu_ho = st.text_input("Số CCCD chủ hộ", value=nv_data.get('so_cccd_chu_ho', '') if nv_data else '')
                        tinh_chu_ho_index = 0
                        tinh_chu_ho_current = nv_data.get('tinh_thanh_pho_chu_ho', '') if nv_data else ''
                        if tinh_chu_ho_current in tinh_options:
                            tinh_chu_ho_index = list(tinh_options.keys()).index(tinh_chu_ho_current) + 1
                        tinh_chu_ho = st.selectbox("Tỉnh/Thành phố (chủ hộ)", options=[""] + list(tinh_options.keys()), index=tinh_chu_ho_index)
                    with col2:
                        phuong_xa_options = []
                        phuong_xa_current = nv_data.get('phuong_xa_chu_ho', '') if nv_data else ''
                        if tinh_chu_ho and tinh_chu_ho != "":
                            ma_tinh = tinh_options.get(tinh_chu_ho)
                            db_temp2 = st.session_state.db_engine.get_connection()
                            c_temp2 = db_temp2.cursor()
                            c_temp2.execute("SELECT ten_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh,))
                            phuong_xa_options = [row[0] for row in c_temp2.fetchall()]
                            db_temp2.close()
                        phuong_xa_index = 0
                        if phuong_xa_current in phuong_xa_options:
                            phuong_xa_index = phuong_xa_options.index(phuong_xa_current) + 1
                        phuong_xa_chu_ho = st.selectbox("Phường/Xã (chủ hộ)", options=[""] + phuong_xa_options, index=phuong_xa_index)
                    
                    st.markdown("**Thông tin thường trú:**")
                    col_tt1, col_tt2 = st.columns(2)
                    with col_tt1:
                        tinh_thuong_tru_index = 0
                        tinh_thuong_tru_current = nv_data.get('tinh_thanh_pho_thuong_tru', '') if nv_data else ''
                        if tinh_thuong_tru_current in tinh_options:
                            tinh_thuong_tru_index = list(tinh_options.keys()).index(tinh_thuong_tru_current) + 1
                        tinh_thuong_tru = st.selectbox("Tỉnh/Thành phố thường trú", options=[""] + list(tinh_options.keys()), index=tinh_thuong_tru_index)
                        ma_tinh_thuong_tru = tinh_options.get(tinh_thuong_tru, "") if tinh_thuong_tru else ""
                    with col_tt2:
                        phuong_xa_tt_options = []
                        phuong_xa_tt_current = nv_data.get('phuong_xa_thuong_tru', '') if nv_data else ''
                        if tinh_thuong_tru and tinh_thuong_tru != "":
                            ma_tinh_tt = tinh_options.get(tinh_thuong_tru)
                            db_temp3 = st.session_state.db_engine.get_connection()
                            c_temp3 = db_temp3.cursor()
                            c_temp3.execute("SELECT ten_xa, ma_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh_tt,))
                            phuong_xa_tt_options = c_temp3.fetchall()
                            db_temp3.close()
                        phuong_xa_tt_index = 0
                        ma_phuong_xa_thuong_tru = ""
                        for i, px in enumerate(phuong_xa_tt_options):
                            if px[0] == phuong_xa_tt_current:
                                phuong_xa_tt_index = i + 1
                                ma_phuong_xa_thuong_tru = px[1]
                                break
                        phuong_xa_display_list = [""] + [px[0] for px in phuong_xa_tt_options]
                        phuong_xa_thuong_tru = st.selectbox("Phường/Xã thường trú", options=phuong_xa_display_list, index=phuong_xa_tt_index)
                        for px in phuong_xa_tt_options:
                            if px[0] == phuong_xa_thuong_tru:
                                ma_phuong_xa_thuong_tru = px[1]
                                break
                    
                    st.markdown("**II. THÊM THÀNH VIÊN MỚI:**")
                    st.caption("Điền thông tin vào các cột bên dưới, sau đó bấm '➕ Thêm thành viên'")
                    col_tv1, col_tv2, col_tv3, col_tv4, col_tv5, col_tv6, col_tv7, col_tv8 = st.columns([2,1.3,1,1,1,1.5,1.8,1.8])
                    with col_tv1:
                        ho_ten_tv = st.text_input("Họ và tên", key="tv_ho_ten_family", placeholder="Nguyễn Văn A")
                    with col_tv2:
                        ngay_sinh_tv = st.text_input("Ngày sinh", key="tv_ngay_sinh_family", placeholder="dd/mm/yyyy")
                    with col_tv3:
                        gioi_tinh_tv = st.selectbox("Giới tính", ["Nam", "Nữ"], key="tv_gioi_tinh_family")
                    with col_tv4:
                        quoc_tich_tv = st.text_input("Quốc tịch", value="Việt Nam", key="tv_quoc_tich_family")
                    with col_tv5:
                        dan_toc_tv = st.text_input("Dân tộc", value="Kinh", key="tv_dan_toc_family")
                    with col_tv6:
                        quan_he_tv = st.selectbox("Quan hệ chủ hộ", ["", "Vợ", "Chồng", "Con", "Bố", "Mẹ", "Anh", "Chị", "Em", "Ông", "Bà", "Khác"], key="tv_quan_he_family")
                    with col_tv7:
                        tinh_tv = st.selectbox("Tỉnh/Thành phố", options=[""] + list(tinh_options.keys()), key="tv_tinh_family")
                    with col_tv8:
                        phuong_xa_tv_options = []
                        if tinh_tv and tinh_tv != "":
                            ma_tinh_tv = tinh_options.get(tinh_tv)
                            db_temp4 = st.session_state.db_engine.get_connection()
                            c_temp4 = db_temp4.cursor()
                            c_temp4.execute("SELECT ten_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh_tv,))
                            phuong_xa_tv_options = [row[0] for row in c_temp4.fetchall()]
                            db_temp4.close()
                        phuong_xa_tv = st.selectbox("Phường/Xã", options=[""] + phuong_xa_tv_options, key="tv_phuong_xa_family")
                    
                    col_btn_add1, col_btn_add2, col_btn_add3 = st.columns([1,1,1])
                    with col_btn_add2:
                        if st.form_submit_button("➕ Thêm thành viên vào danh sách", width='stretch'):
                            if ho_ten_tv:
                                st.session_state.bhxh_family_members.append({
                                    'ho_ten': ho_ten_tv, 'ngay_sinh': parse_date(ngay_sinh_tv), 'gioi_tinh': gioi_tinh_tv,
                                    'quoc_tich': quoc_tich_tv, 'dan_toc': dan_toc_tv, 'quan_he': quan_he_tv,
                                    'tinh': tinh_tv, 'phuong_xa': phuong_xa_tv
                                })
                                st.rerun()
                            else:
                                st.error("Vui lòng nhập họ tên thành viên")
                    
                    st.markdown("---")
                    col_save1, col_save2, col_save3 = st.columns([1,2,1])
                    with col_save2:
                        if st.form_submit_button("💾 LƯU THÔNG TIN CHỦ HỘ", width='stretch', type="primary"):
                            try:
                                db_luu = st.session_state.db_engine.get_connection()
                                c_luu = db_luu.cursor()
                                c_luu.execute("""UPDATE nhan_vien SET ho_ten_chu_ho=%s, so_cccd_chu_ho=%s, tinh_thanh_pho_chu_ho=%s, phuong_xa_chu_ho=%s,
                                    tinh_thanh_pho_thuong_tru=%s, ma_tinh_thuong_tru=%s, phuong_xa_thuong_tru=%s, ma_phuong_xa_thuong_tru=%s WHERE id=%s""",
                                    (ho_ten_chu_ho, so_cccd_chu_ho, tinh_chu_ho, phuong_xa_chu_ho, tinh_thuong_tru, ma_tinh_thuong_tru, phuong_xa_thuong_tru, ma_phuong_xa_thuong_tru, nv_id))
                                c_luu.execute("DELETE FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_id,))
                                for tv in st.session_state.bhxh_family_members:
                                    c_luu.execute("""INSERT INTO phu_luc_gia_dinh (nhan_vien_id, ho_ten, ngay_sinh, gioi_tinh, quoc_tich, dan_toc, quan_he_voi_chu_ho, tinh_thanh_pho, phuong_xa) 
                                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                                        (nv_id, tv['ho_ten'], tv['ngay_sinh'], tv['gioi_tinh'], tv['quoc_tich'], tv['dan_toc'], tv['quan_he'], tv['tinh'], tv['phuong_xa']))
                                db_luu.commit()
                                db_luu.close()
                                del st.session_state['bhxh_family_nv_id']
                                del st.session_state['bhxh_family_nv_name']
                                del st.session_state['bhxh_family_members']
                                st.success(f"✅ Đã lưu thông tin hộ gia đình cho nhân viên {nv_name}")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi lưu: {e}")
                
                col_cancel1, col_cancel2, col_cancel3 = st.columns([1,2,1])
                with col_cancel2:
                    if st.button("❌ HỦY BỎ", width='stretch'):
                        del st.session_state['bhxh_family_nv_id']
                        del st.session_state['bhxh_family_nv_name']
                        if 'bhxh_family_members' in st.session_state:
                            del st.session_state['bhxh_family_members']
                        st.rerun()
    
    with tab_da_nghi:
        st.caption("📋 Danh sách nhân viên đã nghỉ việc (có thông tin ngày nghỉ)")
        
        col_filter1, col_filter2, col_filter3 = st.columns([2, 2, 1])
        with col_filter1:
            search_nghi = st.text_input("🔍 Tìm kiếm (Tên, Mã NV, SĐT, CCCD)", key="search_da_nghi")
        with col_filter2:
            db_temp = st.session_state.db_engine.get_connection()
            c_temp = db_temp.cursor()
            c_temp.execute("SELECT DISTINCT EXTRACT(YEAR FROM ngay_ket_thuc) as nam FROM nhan_vien WHERE trang_thai='NGHI_VIEC' AND ngay_ket_thuc IS NOT NULL ORDER BY nam DESC")
            years = [int(row[0]) for row in c_temp.fetchall() if row[0] is not None]
            db_temp.close()
            filter_nam = st.selectbox("📅 Lọc theo năm nghỉ:", ["Tất cả"] + [str(y) for y in years] if years else ["Tất cả"])
        
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        sql = """
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, so_cccd, dien_thoai, 
                   chuc_danh_nghe, loai_hop_dong, so_hdld, ngay_vao_lam, ngay_ket_thuc,
                   ma_so_bhxh, thang_bat_dau_bh, ly_do_nghi, ten_don_vi_thu_huong
            FROM nhan_vien 
            WHERE trang_thai = 'NGHI_VIEC'
        """
        params = []
        
        if search_nghi:
            sql += " AND (ho_ten LIKE %s OR ma_nv LIKE %s OR dien_thoai LIKE %s OR so_cccd LIKE %s)"
            params.extend([f'%{search_nghi}%'] * 4)
        
        if filter_nam != "Tất cả" and filter_nam.isdigit():
            sql += " AND EXTRACT(YEAR FROM ngay_ket_thuc) = %s"
            params.append(int(filter_nam))
        
        sql += " ORDER BY ngay_ket_thuc DESC, id DESC"
        c.execute(sql, tuple(params))
        ds_nghi = c.fetchall()
        db.close()
        
        if ds_nghi:
            df_nghi = pd.DataFrame(ds_nghi)
            for col in df_nghi.columns:
                if 'ngay' in col.lower():
                    df_nghi[col] = df_nghi[col].apply(format_date)
            
            display_cols_nghi = ['ma_nv', 'ho_ten', 'ngay_sinh', 'gioi_tinh', 'chuc_danh_nghe', 
                                 'so_hdld', 'ngay_vao_lam', 'ngay_ket_thuc', 'dien_thoai', 'ma_so_bhxh', 'ten_don_vi_thu_huong']
            available_cols_nghi = [c for c in display_cols_nghi if c in df_nghi.columns]
            df_show_nghi = df_nghi[available_cols_nghi]
            
            col_map_nghi = {
                'ma_nv': 'Mã NV',
                'ho_ten': 'Họ và tên',
                'ngay_sinh': 'Ngày sinh',
                'gioi_tinh': 'Giới tính',
                'chuc_danh_nghe': 'Chức danh',
                'so_hdld': 'Số HĐLĐ',
                'ngay_vao_lam': 'Ngày vào làm',
                'ngay_ket_thuc': '📅 Ngày nghỉ việc',
                'dien_thoai': 'SĐT',
                'ma_so_bhxh': 'Mã BHXH',
                'ten_don_vi_thu_huong': 'Tên đơn vị thụ hưởng',
            }
            df_show_nghi.rename(columns=col_map_nghi, inplace=True)
            
            st.caption(f"📌 Tổng số: **{len(ds_nghi)}** nhân viên đã nghỉ việc")
            st.dataframe(df_show_nghi, width='stretch', hide_index=True, height=400)
            
            st.divider()
            st.subheader("🔍 Xem chi tiết / Khôi phục nhân viên")
            
            nv_options = {f"{nv['ma_nv']} - {nv['ho_ten']} (Nghỉ: {format_date(nv.get('ngay_ket_thuc'))})": nv['id'] for nv in ds_nghi}
            selected_nghi_name = st.selectbox("Chọn nhân viên đã nghỉ:", list(nv_options.keys()), help="💡 Gõ mã NV hoặc tên để tìm nhanh trong danh sách")
            selected_nghi_id = nv_options[selected_nghi_name]
            
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("SELECT * FROM nhan_vien WHERE id = %s", (selected_nghi_id,))
            nv_nghi_detail = c.fetchone()
            db.close()
            
            if nv_nghi_detail:
                col_detail1, col_detail2 = st.columns(2)
                with col_detail1:
                    st.markdown("**📋 Thông tin cá nhân**")
                    st.write(f"- **Mã NV:** {nv_nghi_detail.get('ma_nv', '')}")
                    st.write(f"- **Họ tên:** {nv_nghi_detail.get('ho_ten', '')}")
                    st.write(f"- **Ngày sinh:** {format_date(nv_nghi_detail.get('ngay_sinh'))}")
                    st.write(f"- **Giới tính:** {nv_nghi_detail.get('gioi_tinh', '')}")
                    st.write(f"- **CCCD:** {nv_nghi_detail.get('so_cccd', '')}")
                    st.write(f"- **SĐT:** {nv_nghi_detail.get('dien_thoai', '')}")
                    st.write(f"- **Chức danh:** {nv_nghi_detail.get('chuc_danh_nghe', '')}")
                    st.write(f"- **Tên đơn vị thụ hưởng:** {nv_nghi_detail.get('ten_don_vi_thu_huong', '')}")
                
                with col_detail2:
                    st.markdown("**📅 Thông tin hợp đồng & nghỉ việc**")
                    st.write(f"- **Số HĐLĐ:** {nv_nghi_detail.get('so_hdld', '')}")
                    st.write(f"- **Loại HĐ:** {nv_nghi_detail.get('loai_hop_dong', '')}")
                    st.write(f"- **Ngày vào làm:** {format_date(nv_nghi_detail.get('ngay_vao_lam'))}")
                    st.write(f"- **📅 Ngày nghỉ việc:** **{format_date(nv_nghi_detail.get('ngay_ket_thuc'))}**")
                    st.write(f"- **Mã BHXH:** {nv_nghi_detail.get('ma_so_bhxh', '')}")
                    st.write(f"- **Lý do nghỉ:** {nv_nghi_detail.get('ly_do_nghi', 'Chưa có thông tin')}")
                
                if st.session_state.role == "admin":
                    st.divider()
                    col_restore1, col_restore2, col_restore3 = st.columns([1, 2, 1])
                    with col_restore2:
                        if st.button(f"🔄 KHÔI PHỤC NHÂN VIÊN - {nv_nghi_detail['ho_ten']}", width='stretch', type="primary"):
                            try:
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                loai_hd = nv_nghi_detail.get('loai_hop_dong', '')
                                if loai_hd == 'Thử việc':
                                    trang_thai_moi = 'THU_VIEC'
                                else:
                                    trang_thai_moi = 'DANG_LAM'
                                c.execute("""
                                    UPDATE nhan_vien 
                                    SET trang_thai = %s, 
                                        ngay_ket_thuc = NULL
                                    WHERE id = %s
                                """, (trang_thai_moi, selected_nghi_id))
                                db.commit()
                                db.close()
                                st.success(f"✅ Đã khôi phục nhân viên {nv_nghi_detail['ho_ten']}!")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi khôi phục: {e}")
        else:
            st.info("📭 Không có nhân viên nào đã nghỉ việc")
    
    with tab_qtct:
        st.caption("📜 Lịch sử công tác và quyết định nhân sự")
        
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien ORDER BY id DESC")
        all_nv = c.fetchall()
        db.close()
        
        if all_nv:
            nv_options = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in all_nv}
            selected_nv_history = st.selectbox("🔍 Chọn nhân viên:", list(nv_options.keys()), key="history_nv", help="💡 Gõ mã NV hoặc tên để tìm nhanh trong danh sách")
            nv_id_history = nv_options[selected_nv_history]
            
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("SELECT * FROM nhan_vien WHERE id = %s", (nv_id_history,))
            nv_current = c.fetchone()
            
            st.markdown(f"""
            ### 📌 Thông tin hiện tại của {nv_current['ho_ten']} ({nv_current['ma_nv']})
            | Trường | Giá trị |
            |--------|---------|
            | Trạng thái | {'🟢 Đang làm' if nv_current['trang_thai'] == 'DANG_LAM' else '🔵 Thử việc' if nv_current['trang_thai'] == 'THU_VIEC' else '🔴 Đã nghỉ'} |
            | Loại hợp đồng | {nv_current['loai_hop_dong']} |
            | Ngày vào làm | {format_date(nv_current['ngay_vao_lam'])} |
            | Ngày chính thức | {format_date(nv_current.get('ngay_chinh_thuc')) or 'Chưa có'} |
            | Chức danh | {nv_current['chuc_danh_nghe']} |
            | Phòng ban | {nv_current['phong_ban_lam_viec']} |
            """)
            
            c.execute("""
                SELECT * FROM quyet_dinh_nhan_su 
                WHERE nhan_vien_id = %s 
                ORDER BY ngay_quyet_dinh DESC
            """, (nv_id_history,))
            quyet_dinh_list = c.fetchall()
            
            if quyet_dinh_list:
                st.markdown("### 📋 Các quyết định nhân sự")
                qd_data = []
                for i, qd in enumerate(quyet_dinh_list, 1):
                    loai_qd_map = {
                        'THU_VIEC': '📝 Quyết định thử việc',
                        'CHINH_THUC': '✅ Quyết định chính thức',
                        'DIEU_CHUYEN': '🔄 Quyết định điều chuyển',
                        'BO_NHIEM': '⭐ Quyết định bổ nhiệm',
                        'TANG_LUONG': '💰 Quyết định tăng lương',
                        'NGHI_VIEC': '🚫 Quyết định nghỉ việc'
                    }
                    qd_data.append({
                        "STT": i,
                        "Loại quyết định": loai_qd_map.get(qd['loai_quyet_dinh'], qd['loai_quyet_dinh']),
                        "Số quyết định": qd['so_quyet_dinh'] or '...',
                        "Ngày quyết định": format_date(qd['ngay_quyet_dinh']),
                        "Ngày hiệu lực": format_date(qd['ngay_hieu_luc']),
                        "Nội dung": (qd['noi_dung'][:50] + "...") if qd['noi_dung'] and len(qd['noi_dung']) > 50 else qd['noi_dung']
                    })
                df_qd = pd.DataFrame(qd_data)
                st.dataframe(df_qd, width='stretch', hide_index=True)
                
                with st.expander("🔍 Xem chi tiết quyết định"):
                    qd_options = {f"{format_date(qd['ngay_quyet_dinh'])} - {qd['loai_quyet_dinh']}": qd for qd in quyet_dinh_list}
                    selected_qd_name = st.selectbox("Chọn quyết định:", list(qd_options.keys()), key="qd_detail")
                    selected_qd = qd_options[selected_qd_name]
                    st.markdown(f"""
                    **📄 Chi tiết quyết định:**
                    - **Số quyết định:** {selected_qd.get('so_quyet_dinh', '...')}
                    - **Ngày quyết định:** {format_date(selected_qd.get('ngay_quyet_dinh'))}
                    - **Ngày hiệu lực:** {format_date(selected_qd.get('ngay_hieu_luc'))}
                    - **Loại quyết định:** {selected_qd.get('loai_quyet_dinh')}
                    - **Nội dung:** {selected_qd.get('noi_dung', '...')}
                    - **Người ký:** {selected_qd.get('nguoi_ky', COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC'))}
                    
                    **📊 Thay đổi:**
                    - Chức danh: {selected_qd.get('chuc_danh_cu', '...')} → {selected_qd.get('chuc_danh_moi', '...')}
                    - Phòng ban: {selected_qd.get('phong_ban_cu', '...')} → {selected_qd.get('phong_ban_moi', '...')}
                    - Loại HĐ: {selected_qd.get('loai_hop_dong_cu', '...')} → {selected_qd.get('loai_hop_dong_moi', '...')}
                    """)
            else:
                st.info("📭 Nhân viên này chưa có quyết định nào.")
            
            c.execute("""
                SELECT * FROM lich_su_cong_tac 
                WHERE nhan_vien_id = %s 
                ORDER BY tu_ngay ASC
            """, (nv_id_history,))
            lich_su_list = c.fetchall()
            
            if lich_su_list:
                st.markdown("### 📅 Lịch sử công tác")
                ls_data = []
                for i, ls in enumerate(lich_su_list, 1):
                    ls_data.append({
                        "STT": i,
                        "Từ ngày": format_date(ls['tu_ngay']),
                        "Đến ngày": format_date(ls['den_ngay']) if ls['den_ngay'] else "Đang làm",
                        "Chức danh": ls['chuc_danh'] or '',
                        "Phòng ban": ls['phong_ban'] or '',
                        "Loại HĐ": ls['loai_hop_dong'] or '',
                        "Hệ số lương": ls['he_so_luong'] or ''
                    })
                df_ls = pd.DataFrame(ls_data)
                st.dataframe(df_ls, width='stretch', hide_index=True, height=400)
            else:
                st.info("📭 Chưa có lịch sử công tác. Đang tạo từ dữ liệu hiện tại...")
                loai_hd_dung = nv_current['loai_hop_dong']
                if nv_current['trang_thai'] == 'THU_VIEC':
                    loai_hd_dung = 'Thử việc'
                c.execute("""
                    INSERT INTO lich_su_cong_tac (nhan_vien_id, tu_ngay, chuc_danh, phong_ban, noi_lam_viec, loai_hop_dong, he_so_luong)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (nv_id_history, nv_current['ngay_vao_lam'], nv_current['chuc_danh_nghe'], 
                      nv_current['phong_ban_lam_viec'], nv_current['noi_lam_viec'], 
                      loai_hd_dung, nv_current['he_so_luong']))
                db.commit()
                st.rerun()
            db.close()
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")
    
    # ========== PHẦN XÓA NHÂN VIÊN THEO SỐ HĐ ==========
    st.divider()
    
    if st.session_state.role == "admin":
        with st.expander("🗑️ CÔNG CỤ XÓA NHÂN VIÊN (CHỈ DÀNH CHO ADMIN)", expanded=False):
            st.warning("⚠️ **CẢNH BÁO:** Thao tác này sẽ XÓA VĨNH VIỄN nhân viên và tất cả dữ liệu liên quan!")
        
            col_hd1, col_hd2 = st.columns([2, 1])
            with col_hd1:
                so_hd_can_xoa = st.text_input("📝 Nhập số hợp đồng cần xóa (VD: 21/2026/HĐTV-CHL):", key="so_hd_xoa")
            with col_hd2:
                st.write("")
                st.write("")
                xac_nhan_xoa = st.checkbox("✅ Tôi xác nhận muốn xóa vĩnh viễn", key="xac_nhan_xoa_nv")
        
            if so_hd_can_xoa and xac_nhan_xoa:
                try:
                    db_check = st.session_state.db_engine.get_connection()
                    c_check = db_check.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c_check.execute("SELECT id, ho_ten, ma_nv, trang_thai FROM nhan_vien WHERE so_hdld = %s", (so_hd_can_xoa,))
                    nv_info = c_check.fetchone()
                    db_check.close()
                
                    if nv_info:
                        st.warning(f"⚠️ Nhân viên: **{nv_info['ho_ten']}** (Mã: {nv_info['ma_nv']}) - Trạng thái: {nv_info['trang_thai']}")
                    
                        # Đếm số bản ghi liên quan
                        db_count = st.session_state.db_engine.get_connection()
                        c_count = db_count.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                        c_count.execute("SELECT COUNT(*) as count FROM lich_su_cong_tac WHERE nhan_vien_id = %s", (nv_info['id'],))
                        ls_count = c_count.fetchone()['count']
                        c_count.execute("SELECT COUNT(*) as count FROM quyet_dinh_nhan_su WHERE nhan_vien_id = %s", (nv_info['id'],))
                        qd_count = c_count.fetchone()['count']
                        c_count.execute("SELECT COUNT(*) as count FROM ho_so_nhan_vien WHERE nhan_vien_id = %s", (nv_info['id'],))
                        hs_count = c_count.fetchone()['count']
                        c_count.execute("SELECT COUNT(*) as count FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_info['id'],))
                        pl_count = c_count.fetchone()['count']
                        db_count.close()
                    
                        st.info(f"📊 Sẽ xóa: {ls_count} lịch sử công tác, {qd_count} quyết định, {hs_count} hồ sơ, {pl_count} phụ lục gia đình")
                    
                        # Xác nhận lần cuối
                        xac_nhan_cuoi = st.checkbox("⚠️ Tôi hiểu rủi ro và muốn xóa VĨNH VIỄN nhân viên này", key="xac_nhan_cuoi_xoa")
                    
                        # Tìm phần xóa nhân viên (khoảng dòng 1900)
                        if xac_nhan_cuoi:
                            col_confirm1, col_confirm2, col_confirm3 = st.columns([1, 2, 1])
                            with col_confirm2:
                                if st.button("🗑️ XÁC NHẬN XÓA VĨNH VIỄN", type="primary", key="btn_confirm_xoa"):
                                    if not can_delete():
                                        st.error("❌ Bạn không có quyền xóa dữ liệu!")
                                    else:
                                        try:
                                            db = st.session_state.db_engine.get_connection()
                                            cur = db.cursor()
                                        
                                            cur.execute("DELETE FROM lich_su_cong_tac WHERE nhan_vien_id = %s", (nv_info['id'],))
                                            cur.execute("DELETE FROM quyet_dinh_nhan_su WHERE nhan_vien_id = %s", (nv_info['id'],))
                                            cur.execute("DELETE FROM ho_so_nhan_vien WHERE nhan_vien_id = %s", (nv_info['id'],))
                                            cur.execute("DELETE FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_info['id'],))
                                            cur.execute("DELETE FROM nhan_vien WHERE id = %s", (nv_info['id'],))
                                        
                                            db.commit()
                                            db.close()
                                        
                                            st.success(f"✅ Đã XÓA VĨNH VIỄN nhân viên {nv_info['ho_ten']} (Mã: {nv_info['ma_nv']})")
                                            st.balloons()
                                            st.cache_data.clear()
                                            st.rerun()
                                        
                                        except Exception as e:
                                            st.error(f"❌ Lỗi khi xóa: {str(e)}")
                                            try:
                                                db.rollback()
                                                db.close()
                                            except:
                                                pass
                        else:
                            st.info("🔒 **Vui lòng tick vào ô xác nhận 'Tôi hiểu rủi ro...' để kích hoạt nút xóa**")
                    else:
                        st.error(f"❌ Không tìm thấy nhân viên có số hợp đồng: {so_hd_can_xoa}")
                    
                except Exception as e:
                    st.error(f"❌ Lỗi khi tìm kiếm: {e}")
                
            elif so_hd_can_xoa and not xac_nhan_xoa:
                st.info("🔒 Vui lòng tick xác nhận 'Tôi xác nhận muốn xóa vĩnh viễn' để tiếp tục")
    
    # ===== TAB: QUYẾT ĐỊNH NHÂN SỰ =====
    with tab_qdns:
        st.caption("📜 Ra các Quyết định nhân sự: Bổ nhiệm, Miễn nhiệm, Thay đổi chức danh, Điều chuyển công tác, Chấm dứt HĐTV/HĐLĐ")

        # Thông báo file vừa tạo (nếu có), hiển thị TRƯỚC form để không bị mất sau khi rerun
        if st.session_state.get('qdns_last_file'):
            st.success(f"✅ Đã tạo {st.session_state.get('qdns_last_label','Quyết định')} số {st.session_state.get('qdns_last_so')}")
            try:
                with open(st.session_state['qdns_last_file'], "rb") as f:
                    st.download_button(
                        label="📥 TẢI QUYẾT ĐỊNH (Word)",
                        data=f,
                        file_name=f"QDNS_{st.session_state.get('qdns_last_so','').replace('/', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="dl_qdns_last"
                    )
            except Exception:
                pass
            if st.button("✖️ Đóng thông báo", key="close_qdns_notice"):
                for k in ['qdns_last_file', 'qdns_last_label', 'qdns_last_so']:
                    st.session_state.pop(k, None)
                st.rerun()
            st.divider()

        db_qd = st.session_state.db_engine.get_connection()
        c_qd = db_qd.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c_qd.execute("""
            SELECT id, ma_nv, ho_ten, chuc_vu, chuc_danh_nghe, phong_ban_lam_viec, loai_hop_dong, trang_thai
            FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY ho_ten
        """)
        nv_qd_list = c_qd.fetchall()
        c_qd.execute("""
            SELECT DISTINCT phong_ban_lam_viec FROM nhan_vien
            WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND phong_ban_lam_viec IS NOT NULL AND phong_ban_lam_viec != ''
            ORDER BY phong_ban_lam_viec
        """)
        ds_phong_ban = [r['phong_ban_lam_viec'] for r in c_qd.fetchall()]
        db_qd.close()

        if not nv_qd_list:
            st.info("Không có nhân viên đang làm việc.")
        else:
            nv_qd_options = {f"{nv['ma_nv']} - {nv['ho_ten']}": nv for nv in nv_qd_list}
            chon_nv_label = st.selectbox("👤 Chọn nhân viên:", list(nv_qd_options.keys()), key="qdns_chon_nv",
                                          help="💡 Bấm vào ô rồi gõ tên/mã NV để lọc nhanh — không cần scroll (Ctrl+F trình duyệt không lọc được ô này).")
            nv_qd = nv_qd_options[chon_nv_label]

            col_info1, col_info2, col_info3 = st.columns(3)
            col_info1.markdown(f"**Chức vụ hiện tại:** {nv_qd.get('chuc_vu') or 'Nhân viên'}")
            col_info2.markdown(f"**Chức danh hiện tại:** {nv_qd.get('chuc_danh_nghe') or '-'}")
            col_info3.markdown(f"**Phòng ban hiện tại:** {nv_qd.get('phong_ban_lam_viec') or '-'}")

            col_qd1, col_qd2 = st.columns(2)
            with col_qd1:
                loai_qd = st.selectbox(
                    "📋 Loại quyết định:",
                    list(LOAI_QDNS_LABEL.keys()),
                    format_func=lambda k: LOAI_QDNS_LABEL[k],
                    key="qdns_loai"
                )
            with col_qd2:
                ngay_qd = st.date_input("📅 Ngày ban hành quyết định:", value=date.today(), key="qdns_ngay")

            dieu1_lines = []
            tieu_de = ""
            hieu_luc_text = None
            gia_tri_truoc = None
            gia_tri_sau = None
            ok_to_submit = True

            if loai_qd == 'BO_NHIEM':
                chuc_vu_moi = st.selectbox("🏷️ Chức vụ được bổ nhiệm:", DANH_SACH_CHUC_VU, key="qdns_cv_bonhiem")
                tieu_de = f"Bổ nhiệm chức vụ {chuc_vu_moi}"
                dieu1_lines = [f"Bổ nhiệm Ông/Bà {nv_qd['ho_ten']} ({nv_qd['ma_nv']}) giữ chức vụ {chuc_vu_moi} kể từ ngày {ngay_qd.strftime('%d/%m/%Y')}."]
                gia_tri_truoc = nv_qd.get('chuc_vu') or 'Nhân viên'
                gia_tri_sau = chuc_vu_moi

            elif loai_qd == 'MIEN_NHIEM':
                cv_hien_tai = nv_qd.get('chuc_vu') or 'Nhân viên'
                idx_mn = DANH_SACH_CHUC_VU.index(cv_hien_tai) if cv_hien_tai in DANH_SACH_CHUC_VU else 0
                chuc_vu_mien = st.selectbox("🏷️ Chức vụ bị miễn nhiệm:", DANH_SACH_CHUC_VU, index=idx_mn, key="qdns_cv_miennhiem")
                tieu_de = f"Miễn nhiệm chức vụ {chuc_vu_mien}"
                dieu1_lines = [f"Miễn nhiệm chức vụ {chuc_vu_mien} đối với Ông/Bà {nv_qd['ho_ten']} ({nv_qd['ma_nv']}) kể từ ngày {ngay_qd.strftime('%d/%m/%Y')}."]
                gia_tri_truoc = cv_hien_tai
                gia_tri_sau = 'Nhân viên'
                if cv_hien_tai == 'Nhân viên':
                    st.warning("⚠️ Nhân viên này hiện đang giữ chức vụ 'Nhân viên' (không có chức vụ quản lý để miễn nhiệm).")

            elif loai_qd == 'DOI_CHUC_DANH':
                chuc_danh_moi = st.text_input("💼 Chức danh mới:", value=nv_qd.get('chuc_danh_nghe') or '', key="qdns_cd_moi")
                tieu_de = f"Thay đổi chức danh - {nv_qd['ho_ten']}"
                dieu1_lines = [f"Thay đổi chức danh của Ông/Bà {nv_qd['ho_ten']} ({nv_qd['ma_nv']}) từ '{nv_qd.get('chuc_danh_nghe') or ''}' thành '{chuc_danh_moi}' kể từ ngày {ngay_qd.strftime('%d/%m/%Y')}."]
                gia_tri_truoc = nv_qd.get('chuc_danh_nghe') or ''
                gia_tri_sau = chuc_danh_moi
                if not chuc_danh_moi.strip():
                    ok_to_submit = False
                    st.error("⚠️ Vui lòng nhập chức danh mới.")

            elif loai_qd == 'DIEU_CHUYEN':
                phong_hien_tai = nv_qd.get('phong_ban_lam_viec') or ''
                st.text_input("🏢 Từ phòng ban:", value=phong_hien_tai, disabled=True, key="qdns_pb_tu")
                tuy_chon_pb = ds_phong_ban + ["➕ Nhập phòng ban khác..."]
                chon_pb = st.selectbox("🏢 Đến phòng ban:", tuy_chon_pb, key="qdns_pb_den_select")
                if chon_pb == "➕ Nhập phòng ban khác...":
                    phong_moi = st.text_input("Nhập tên phòng ban mới:", key="qdns_pb_den_moi")
                else:
                    phong_moi = chon_pb
                tieu_de = f"Điều chuyển công tác - {nv_qd['ho_ten']}"
                dieu1_lines = [f"Điều chuyển Ông/Bà {nv_qd['ho_ten']} ({nv_qd['ma_nv']}) từ {phong_hien_tai or '(chưa xác định)'} sang {phong_moi or '(chưa xác định)'} kể từ ngày {ngay_qd.strftime('%d/%m/%Y')}."]
                gia_tri_truoc = phong_hien_tai
                gia_tri_sau = phong_moi
                if not (phong_moi or '').strip():
                    ok_to_submit = False
                    st.error("⚠️ Vui lòng chọn hoặc nhập phòng ban đến.")

            elif loai_qd == 'CHAM_DUT_HD':
                loai_hd_hien_tai = nv_qd.get('loai_hop_dong') or ''
                if loai_hd_hien_tai == 'Thử việc':
                    nhan_hd = "Hợp đồng thử việc (HĐTV)"
                else:
                    nhan_hd = "Hợp đồng lao động (HĐLĐ)"
                st.info(f"🔎 Loại hợp đồng hiện tại: **{loai_hd_hien_tai or 'Chưa xác định'}** → Sẽ ban hành: **QĐ Chấm dứt {nhan_hd}**")
                ly_do_cd = st.text_area("📝 Lý do chấm dứt:", key="qdns_lydo_cd", height=80,
                                          placeholder="VD: Hết hạn hợp đồng, Xin nghỉ theo nguyện vọng cá nhân, Chuyển công tác...")
                tieu_de = f"Chấm dứt {nhan_hd} - {nv_qd['ho_ten']}"
                dieu1_lines = [
                    f"Chấm dứt {nhan_hd} đối với Ông/Bà {nv_qd['ho_ten']} ({nv_qd['ma_nv']}) kể từ ngày {ngay_qd.strftime('%d/%m/%Y')}."
                ]
                if ly_do_cd.strip():
                    dieu1_lines.append(f"Lý do: {ly_do_cd.strip()}.")
                hieu_luc_text = f"Ông/Bà {nv_qd['ho_ten']} có trách nhiệm bàn giao công việc, tài sản (nếu có) trước ngày {ngay_qd.strftime('%d/%m/%Y')}."
                gia_tri_truoc = loai_hd_hien_tai
                gia_tri_sau = 'NGHI_VIEC'

            st.divider()
            if st.button("💾 TẠO QUYẾT ĐỊNH & LƯU", type="primary", width='stretch', key="qdns_submit", disabled=not ok_to_submit):
                try:
                    so_qd = generate_so_cong_van('QUYET_DINH')

                    file_path = tao_quyet_dinh_nhan_su(nv_qd, so_qd, ngay_qd, tieu_de, dieu1_lines, hieu_luc_text)
                    file_url = None
                    # (File Word được tạo để tải về ngay lập tức; bản ghi vẫn được lưu để tra cứu)

                    db_s = st.session_state.db_engine.get_connection()
                    c_s = db_s.cursor()
                    c_s.execute("""
                        INSERT INTO quyet_dinh_nhan_su (so_qd, loai_qd, nhan_vien_id, ngay_qd, noi_dung, gia_tri_truoc, gia_tri_sau, file_url, nguoi_tao)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (so_qd, loai_qd, nv_qd['id'], ngay_qd, " ".join(dieu1_lines), gia_tri_truoc, gia_tri_sau, file_url, st.session_state.username))

                    # Đăng ký vào hệ thống Quản lý công văn đi để cùng theo dõi số thứ tự
                    c_s.execute("""
                        INSERT INTO cong_van_di (so_cong_van, phong_phat_hanh, ngay_phat_hanh, tieu_de, trich_yeu, file_url, loai_cong_van, ghi_chu, nguoi_tao)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (so_qd, "Phòng Tổ chức - Hành chính", ngay_qd, f"{LOAI_QDNS_LABEL[loai_qd]}: {tieu_de}",
                          " ".join(dieu1_lines), file_url, 'QUYET_DINH', f"Quyết định nhân sự - NV: {nv_qd['ho_ten']}", st.session_state.username))

                    # Cập nhật hồ sơ nhân viên theo đúng logic từng loại quyết định
                    if loai_qd == 'BO_NHIEM':
                        c_s.execute("UPDATE nhan_vien SET chuc_vu = %s, ngay_qd_ns = %s WHERE id = %s", (gia_tri_sau, ngay_qd, nv_qd['id']))
                    elif loai_qd == 'MIEN_NHIEM':
                        c_s.execute("UPDATE nhan_vien SET chuc_vu = %s, ngay_qd_ns = %s WHERE id = %s", ('Nhân viên', ngay_qd, nv_qd['id']))
                    elif loai_qd == 'DOI_CHUC_DANH':
                        c_s.execute("UPDATE nhan_vien SET chuc_danh_nghe = %s, ngay_qd_ns = %s WHERE id = %s", (gia_tri_sau, ngay_qd, nv_qd['id']))
                    elif loai_qd == 'DIEU_CHUYEN':
                        c_s.execute("UPDATE nhan_vien SET phong_ban_lam_viec = %s, ngay_qd_ns = %s WHERE id = %s", (gia_tri_sau, ngay_qd, nv_qd['id']))
                    elif loai_qd == 'CHAM_DUT_HD':
                        c_s.execute("""
                            UPDATE nhan_vien SET trang_thai = 'NGHI_VIEC', ngay_ket_thuc = %s, ly_do_nghi = %s WHERE id = %s
                        """, (ngay_qd, ly_do_cd if ly_do_cd.strip() else None, nv_qd['id']))

                    db_s.commit()
                    db_s.close()

                    st.session_state['qdns_last_file'] = file_path
                    st.session_state['qdns_last_label'] = LOAI_QDNS_LABEL[loai_qd]
                    st.session_state['qdns_last_so'] = so_qd
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    try:
                        db_s.rollback()
                        db_s.close()
                    except Exception:
                        pass
                    st.error(f"❌ Lỗi khi tạo quyết định: {e}")

        # ===== Lịch sử các quyết định nhân sự đã ban hành =====
        st.divider()
        st.subheader("📚 Lịch sử Quyết định nhân sự")
        search_qd = st.text_input("🔍 Tìm Quyết định (theo Số QĐ, Mã NV, Họ tên, Nội dung):", key="search_qdns")

        db_h = st.session_state.db_engine.get_connection()
        c_h = db_h.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        try:
            db_h.rollback()  # dọn transaction lỡ bị abort từ thao tác trước đó
            sql_qd = """
                SELECT q.id, q.so_qd, q.loai_qd, q.ngay_qd, q.noi_dung, q.gia_tri_truoc, q.gia_tri_sau,
                       n.ho_ten, n.ma_nv
                FROM quyet_dinh_nhan_su q
                JOIN nhan_vien n ON n.id = q.nhan_vien_id
                WHERE 1=1
            """
            params_qd = []
            if search_qd.strip():
                sql_qd += """ AND (q.so_qd ILIKE %s OR n.ma_nv ILIKE %s OR n.ho_ten ILIKE %s OR q.noi_dung ILIKE %s)"""
                p = f"%{search_qd.strip()}%"
                params_qd.extend([p, p, p, p])
            sql_qd += " ORDER BY q.id DESC LIMIT 200"
            c_h.execute(sql_qd, params_qd)
            lich_su_qd = c_h.fetchall()
        except Exception as e:
            st.error(f"❌ Lỗi tải lịch sử quyết định: {e}")
            lich_su_qd = []
        finally:
            db_h.close()

        if lich_su_qd:
            # Header
            h1, h2, h3, h4, h5, h6, h7 = st.columns([1.2, 1.3, 1, 1.6, 2.2, 0.6, 0.6])
            for h, txt in zip((h1, h2, h3, h4, h5), ("Số QĐ", "Loại QĐ", "Ngày QĐ", "Nhân viên", "Nội dung")):
                h.markdown(f"**{txt}**")
            st.divider()
            for qd in lich_su_qd:
                r1, r2, r3, r4, r5, r6, r7 = st.columns([1.2, 1.3, 1, 1.6, 2.2, 0.6, 0.6])
                r1.write(qd['so_qd'])
                r2.write(LOAI_QDNS_LABEL.get(qd['loai_qd'], qd['loai_qd']))
                r3.write(format_date(qd['ngay_qd']))
                r4.write(f"{qd['ho_ten']} ({qd['ma_nv']})")
                r5.write(qd['noi_dung'] or '')
                if r6.button("✏️", key=f"sua_qd_{qd['id']}", help="Sửa Quyết định"):
                    st.session_state['qdns_dang_sua'] = qd['id']
                    st.rerun()
                if r7.button("🗑️", key=f"xoa_qd_{qd['id']}", help="Xóa Quyết định"):
                    st.session_state['qdns_dang_xoa'] = qd['id']
                    st.rerun()

            # ----- Form SỬA Quyết định -----
            if st.session_state.get('qdns_dang_sua'):
                qd_id_sua = st.session_state['qdns_dang_sua']
                qd_sua = next((q for q in lich_su_qd if q['id'] == qd_id_sua), None)
                if qd_sua:
                    with st.expander(f"✏️ Sửa Quyết định số {qd_sua['so_qd']}", expanded=True):
                        st.caption("⚠️ Chỉ chỉnh sửa thông tin lưu trữ của Quyết định. Việc sửa KHÔNG tự động "
                                   "hoàn tác/áp dụng lại thay đổi tương ứng trên hồ sơ nhân viên (chức vụ, chức danh, phòng ban...).")
                        so_qd_moi = st.text_input("Số QĐ:", value=qd_sua['so_qd'], key=f"edit_so_{qd_id_sua}")
                        ngay_qd_moi = st.date_input("Ngày QĐ:", value=qd_sua['ngay_qd'], key=f"edit_ngay_{qd_id_sua}")
                        noi_dung_moi = st.text_area("Nội dung:", value=qd_sua['noi_dung'] or '', key=f"edit_nd_{qd_id_sua}", height=100)
                        col_luu_qd, col_huy_qd = st.columns(2)
                        with col_luu_qd:
                            if st.button("💾 Lưu thay đổi", key=f"btn_luu_sua_qd_{qd_id_sua}", type="primary", width='stretch'):
                                try:
                                    db_u = st.session_state.db_engine.get_connection()
                                    c_u = db_u.cursor()
                                    c_u.execute("""
                                        UPDATE quyet_dinh_nhan_su SET so_qd=%s, ngay_qd=%s, noi_dung=%s
                                        WHERE id=%s
                                    """, (so_qd_moi, ngay_qd_moi, noi_dung_moi, qd_id_sua))
                                    db_u.commit(); db_u.close()
                                    st.session_state.pop('qdns_dang_sua', None)
                                    st.success("✅ Đã cập nhật Quyết định")
                                    st.cache_data.clear()
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Lỗi: {e}")
                        with col_huy_qd:
                            if st.button("✖️ Hủy", key=f"btn_huy_sua_qd_{qd_id_sua}", width='stretch'):
                                st.session_state.pop('qdns_dang_sua', None)
                                st.rerun()

            # ----- Xác nhận XÓA Quyết định -----
            if st.session_state.get('qdns_dang_xoa'):
                qd_id_xoa = st.session_state['qdns_dang_xoa']
                qd_xoa = next((q for q in lich_su_qd if q['id'] == qd_id_xoa), None)
                if qd_xoa:
                    st.warning(f"⚠️ Xác nhận xóa Quyết định số **{qd_xoa['so_qd']}** ({LOAI_QDNS_LABEL.get(qd_xoa['loai_qd'], qd_xoa['loai_qd'])} — {qd_xoa['ho_ten']})? "
                               f"Việc xóa KHÔNG tự động hoàn tác thay đổi đã áp dụng trên hồ sơ nhân viên.")
                    col_xn_xoa, col_huy_xoa = st.columns(2)
                    with col_xn_xoa:
                        if st.button("🗑️ Xác nhận xóa", key=f"btn_xn_xoa_qd_{qd_id_xoa}", type="primary", width='stretch'):
                            try:
                                db_d = st.session_state.db_engine.get_connection()
                                c_d = db_d.cursor()
                                c_d.execute("DELETE FROM quyet_dinh_nhan_su WHERE id=%s", (qd_id_xoa,))
                                db_d.commit(); db_d.close()
                                st.session_state.pop('qdns_dang_xoa', None)
                                st.success("✅ Đã xóa Quyết định")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")
                    with col_huy_xoa:
                        if st.button("✖️ Hủy", key=f"btn_huy_xoa_qd_{qd_id_xoa}", width='stretch'):
                            st.session_state.pop('qdns_dang_xoa', None)
                            st.rerun()
        else:
            st.info("Chưa có Quyết định nhân sự nào được tạo.")

    with tab_co_cau:
        st.subheader("📋 CƠ CẤU NHÂN SỰ THEO PHÒNG BAN")

        db_ct = st.session_state.db_engine.get_connection()
        c_ct = db_ct.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        # Chỉ lấy nhân sự đang làm việc (đang làm hoặc thử việc) - bỏ nhân viên đã nghỉ việc
        c_ct.execute("""
            SELECT * FROM nhan_vien 
            WHERE phong_ban_lam_viec IS NOT NULL AND phong_ban_lam_viec != ''
            AND trang_thai IN ('DANG_LAM', 'THU_VIEC')
        """)
        tat_ca_nv_ct = c_ct.fetchall()
        db_ct.close()

        # Chuẩn hóa tên phòng ban của từng nhân viên trước khi gộp nhóm — dữ liệu cũ có thể
        # đã bị lưu với cách viết hoa khác nhau (VD: "Tổ Cơ Giới" vs "Tổ Cơ giới"), nếu so
        # khớp nguyên văn sẽ khiến 1 phòng ban bị tách thành nhiều nhóm và thiếu nhân viên.
        for nv_norm in tat_ca_nv_ct:
            nv_norm['phong_ban_lam_viec'] = chuan_hoa_ten_phong_ban(nv_norm.get('phong_ban_lam_viec'))

        cac_phong_ban_ct = sap_xep_phong_ban(list({nv['phong_ban_lam_viec'] for nv in tat_ca_nv_ct}))

        # Hàng search + 2 chỉ số tổng quan (giữ nguyên logic cũ, chỉ bỏ cột tiêu đề vì
        # tiêu đề đã tách thành subheader riêng full-row ở trên)
        col_search_ct, col_tong_ct, col_dl_ct = st.columns([2, 1, 1])
        with col_search_ct:
            pb_chon_ct = st.selectbox("🔍 Chọn tìm kiếm theo phòng ban:", cac_phong_ban_ct, key="pb_thongke_chitiet")

        ds_nv_ct = sap_xep_nhan_vien([nv for nv in tat_ca_nv_ct if nv['phong_ban_lam_viec'] == pb_chon_ct])
        tong_so = len(ds_nv_ct)
        dang_lam_so = len([nv for nv in ds_nv_ct if nv['trang_thai'] in ('DANG_LAM', 'THU_VIEC')])

        with col_tong_ct:
            st.metric("Tổng số nhân sự", tong_so)
        with col_dl_ct:
            st.metric("Nhân sự đang làm", dang_lam_so)

        st.divider()

        # CSS avatar: to bằng ảnh profile (200px, viền cam) khi cột đủ rộng; khi màn hình hẹp
        # (5 cột co lại) tự chuyển sang ảnh chữ nhật bo góc, rộng tối đa theo cột, vẫn giữ viền cam.
        st.markdown("""
        <style>
        .co-cau-avatar-wrap { display:flex; justify-content:center; }
        .co-cau-avatar-img {
            width: 100%;
            max-width: 200px;
            aspect-ratio: 1 / 1;
            border-radius: 50%;
            object-fit: cover;
            border: 4px solid #f59e0b;
            box-shadow: 0 4px 15px rgba(0,0,0,0.15);
        }
        @media (max-width: 1400px) {
            .co-cau-avatar-img {
                aspect-ratio: 4 / 3;
                border-radius: 12px;
                max-width: 100%;
            }
        }
        </style>
        """, unsafe_allow_html=True)

        # Phòng ban đặc biệt: không hiện Lao động chính thức/Thử việc và Đang làm/Nghỉ việc,
        # thay bằng chức vụ (không đúng bản chất & hình thức đối với 2 phòng ban này)
        PHONG_BAN_KHONG_HIEN_TT = PHONG_BAN_LANH_DAO_CAO_CAP
        # Từ khóa nhận diện "người đứng đầu" phòng ban -> luôn xếp hàng đầu, cột giữa.
        # QUAN TRỌNG: so khớp theo TỪ KHÓA (không phân biệt hoa/thường), KHÔNG so khớp
        # chính xác nguyên chuỗi — vì dữ liệu chuc_vu thực tế có thể viết khác đi chút
        # (VD: "Trưởng Phòng", "Tổ trưởng", "Giám đốc", "Trưởng Ban"...). So khớp chính xác
        # từng ký tự trước đây khiến nhiều Trưởng phòng/Tổ trưởng/Đội trưởng không được
        # nhận diện đúng, làm họ bị rơi khỏi hàng đầu (bug đã sửa).
        TU_KHOA_DUNG_DAU = ['chủ tịch', 'tổng giám đốc', 'giám đốc', 'trưởng phòng',
                            'tổ trưởng', 'đội trưởng', 'trưởng ban', 'trưởng bộ phận', 'phụ trách']

        def _la_cap_pho(nv):
            cv = (nv.get('chuc_vu') or '').strip().lower()
            return cv.startswith('phó')

        def _la_dung_dau(nv):
            cv = (nv.get('chuc_vu') or '').strip().lower()
            if not cv or cv.startswith('phó'):
                return False
            return any(tk in cv for tk in TU_KHOA_DUNG_DAU)

        def _vi_tri_hang_cuoi(so_luong, so_cot=5):
            """Vị trí cột (0..4) cho 1 hàng có `so_luong` người (< so_cot, tức hàng cuối chưa đủ).
            Ưu tiên cân xứng quanh cột giữa (index 2 = 'col 3'); người đầu danh sách (đã ưu tiên
            cấp phó) sẽ rơi vào vị trí bên trái nhất trong bộ vị trí được chọn."""
            if so_luong == 1:
                return [2]
            elif so_luong == 2:
                return [1, 3]
            elif so_luong == 3:
                return [1, 2, 3]
            else:
                return list(range(so_luong))  # 4 hoặc 5 người -> bố trí tự do, lấp đầy từ trái

        # ===== ĐOẠN CODE MỚI CHO _lay_anh_src =====
        # Đã có hàm get_avatar_bytes_cached ở trên

        def _lay_anh_src(nv_ct):
            """Trả về src cho thẻ <img>: ảnh hồ sơ nếu có, không thì ảnh mẫu trong static/
            SỬ DỤNG CACHE để không tải lại ảnh mỗi lần render.
            """
            anh_path_ct = nv_ct.get('anh_ho_so')
            
            # ===== CẢI TIẾN: Dùng cache =====
            if anh_path_ct:
                anh_bytes_ct = get_avatar_bytes_cached(anh_path_ct)
                if anh_bytes_ct:
                    img_b64 = base64.b64encode(anh_bytes_ct).decode()
                    return f"data:image/jpeg;base64,{img_b64}"
            
            # Fallback: ảnh mẫu trong static/
            gioi_tinh_ct = nv_ct.get('gioi_tinh', '')
            avatar_file = "avatar_male.png" if gioi_tinh_ct == "Nam" else "avatar_female.png"
            avatar_path = os.path.join(os.path.dirname(__file__), "static", avatar_file)
            if os.path.exists(avatar_path):
                with open(avatar_path, "rb") as f:
                    img_b64 = base64.b64encode(f.read()).decode()
                return f"data:image/png;base64,{img_b64}"
            
            # Fallback cuối cùng: ui-avatars.com
            ten_url = (nv_ct.get('ho_ten') or 'NV').replace(' ', '+')
            return f"https://ui-avatars.com/api/?name={ten_url}&size=200&background=f59e0b&color=fff"

        def _render_the_nv(nv_ct, cols_ct, idx_c):
            with cols_ct[idx_c]:
                img_src = _lay_anh_src(nv_ct)
                st.markdown(f"""
                <div class="co-cau-avatar-wrap">
                    <img src="{img_src}" class="co-cau-avatar-img">
                </div>
                """, unsafe_allow_html=True)

                if pb_chon_ct in PHONG_BAN_KHONG_HIEN_TT:
                    # HĐQT/BTGĐ: gắn xưng hô Ông/Bà theo giới tính, KHÔNG hiện mã NV
                    xung_ho_ct = get_xung_ho_trang_trong(nv_ct.get('gioi_tinh'))
                    ten_hien_thi_ct = f"{xung_ho_ct} {nv_ct['ho_ten']}".strip()
                else:
                    ten_hien_thi_ct = f"{nv_ct['ho_ten']}-{nv_ct['ma_nv']}"
                st.markdown(f"<p style='text-align:center;margin-bottom:0;'><b>{ten_hien_thi_ct}</b></p>", unsafe_allow_html=True)
                if pb_chon_ct not in PHONG_BAN_KHONG_HIEN_TT:
                    # Nhóm BHXH = 'Văn phòng' -> hiện chức vụ; ngược lại hiện chức danh nghề (như trước đây)
                    if (nv_ct.get('nhom_bhxh') or '') == 'Văn phòng':
                        dong_phu = nv_ct.get('chuc_vu') or ''
                    else:
                        dong_phu = nv_ct.get('chuc_danh_nghe') or ''
                    st.markdown(f"<p style='text-align:center;color:gray;font-size:0.85em;'>{dong_phu}</p>", unsafe_allow_html=True)

                if pb_chon_ct in PHONG_BAN_KHONG_HIEN_TT:
                    # Không đúng bản chất/hình thức với HĐQT & BTGĐ -> hiện chức vụ thay vì loại HĐ/trạng thái
                    st.markdown(f"<p style='text-align:center;'>🏷️ {nv_ct.get('chuc_vu') or 'Thành viên'}</p>", unsafe_allow_html=True)
                else:
                    if nv_ct.get('loai_hop_dong') == 'Thử việc':
                        st.markdown("<p style='text-align:center;color:red;'>Thử việc</p>", unsafe_allow_html=True)
                    elif nv_ct.get('loai_hop_dong'):
                        st.markdown("<p style='text-align:center;color:green;'>Lao động chính thức</p>", unsafe_allow_html=True)

                if st.button("Xem chi tiết>>", key=f"xem_ct_{nv_ct['id']}", width='stretch'):
                    st.session_state['_nv_xem_chi_tiet_dashboard'] = nv_ct['id']

        if ds_nv_ct:
            so_cot = 5

            # Tách người đứng đầu phòng ban (nếu có) -> luôn ở hàng đầu tiên, cột giữa (index 2/5).
            # Nếu không có người đứng đầu -> bỏ qua hàng riêng này, các hàng sau tịnh tiến lên.
            nguoi_dung_dau = next((nv for nv in ds_nv_ct if _la_dung_dau(nv)), None)
            if not nguoi_dung_dau:
                # Không có Tổng/Giám đốc/Trưởng/Phụ trách -> đôn "Phó" đầu tiên (theo alpha) lên hàng 1
                nguoi_dung_dau = next((nv for nv in sorted(ds_nv_ct, key=lambda x: x.get('ho_ten') or '') if _la_cap_pho(nv)), None)
            ds_con_lai = [nv for nv in ds_nv_ct if nv is not nguoi_dung_dau]
            # Cấp phó ưu tiên lên đầu (bên trái), sau đó xếp theo alpha bê tên
            ds_con_lai = sorted(ds_con_lai, key=lambda nv: (0 if _la_cap_pho(nv) else 1, nv.get('ho_ten') or ''))

            if nguoi_dung_dau:
                cols_ct = st.columns(so_cot)
                _render_the_nv(nguoi_dung_dau, cols_ct, 2)  # cột giữa trong 5 cột (0,1,[2],3,4)
                st.divider()

            for i in range(0, len(ds_con_lai), so_cot):
                hang = ds_con_lai[i:i + so_cot]
                cols_ct = st.columns(so_cot)
                vi_tri_cot = _vi_tri_hang_cuoi(len(hang), so_cot) if len(hang) < so_cot else list(range(so_cot))
                for nv_ct, idx_c in zip(hang, vi_tri_cot):
                    _render_the_nv(nv_ct, cols_ct, idx_c)

            if st.session_state.get('_nv_xem_chi_tiet_dashboard'):
                nv_id_xem = st.session_state['_nv_xem_chi_tiet_dashboard']
                nv_xem = next((nv for nv in ds_nv_ct if nv['id'] == nv_id_xem), None)
                if nv_xem:
                    st.divider()
                    render_employee_info_card(
                        nv_xem,
                        key_prefix=f"nv_co_cau_{nv_xem['id']}",
                        on_close=lambda: st.session_state.pop('_nv_xem_chi_tiet_dashboard', None)
                    )
        else:
            st.info("Không có nhân sự nào trong phòng ban này.")

# ========== CHẤM CÔNG ==========
elif menu == "🕒 Chấm công":
    st.title("🕒 Chấm công")
    
    # 3 nút lựa chọn phương thức
    col_method1, col_method2, col_method3 = st.columns(3)
    with col_method1:
        if st.button("📝 Thủ công", use_container_width=True, type="primary" if st.session_state.get('cc_method') == 'manual' else "secondary"):
            st.session_state.cc_method = 'manual'
            st.rerun()
    with col_method2:
        if st.button("📥 Máy vân tay", use_container_width=True, type="primary" if st.session_state.get('cc_method') == 'fingerprint' else "secondary"):
            st.session_state.cc_method = 'fingerprint'
            st.rerun()
    with col_method3:
        if st.button("👤 Face ID", use_container_width=True, type="primary" if st.session_state.get('cc_method') == 'faceid' else "secondary"):
            st.session_state.cc_method = 'faceid'
            st.rerun()
    
    st.divider()
    
    # ========== 1. CHẤM CÔNG THỦ CÔNG ==========
    if st.session_state.get('cc_method', 'manual') == 'manual':
        ensure_cham_cong_table()
        
        # Bố cục chọn tháng/năm/bộ phận
        if not st.session_state.get('cc_full_open', False):
            col_m1, col_m2, col_m3, col_m4 = st.columns([1, 1, 2, 1.5])
            with col_m1:
                thang_nhap = st.selectbox("Tháng", list(range(1, 13)), index=date.today().month - 1, key="cc_thang_nhap", label_visibility="collapsed")
            with col_m2:
                nam_nhap = st.number_input("Năm", min_value=2020, max_value=2100, value=date.today().year, step=1, key="cc_nam_nhap", label_visibility="collapsed")
            with col_m3:
                db_bp = st.session_state.db_engine.get_connection()
                c_bp = db_bp.cursor()
                c_bp.execute("""SELECT DISTINCT phong_ban_lam_viec FROM nhan_vien
                                WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND phong_ban_lam_viec IS NOT NULL
                                AND phong_ban_lam_viec != '' ORDER BY phong_ban_lam_viec""")
                all_depts = [r[0] for r in c_bp.fetchall()]
                c_bp.close(); db_bp.close()
                bo_phan_nhap = st.multiselect(
                    "Bộ phận", all_depts, default=[],
                    format_func=lambda d: CHAM_CONG_DEPT_LABEL.get(d, d), 
                    key="cc_bp_nhap",
                    placeholder="Tất cả bộ phận",
                    label_visibility="collapsed"
                )
            with col_m4:
                if st.button("📂 Mở BCC", type="primary", use_container_width=True):
                    st.session_state.cc_full_open = True
                    st.session_state.cc_view_thang = thang_nhap
                    st.session_state.cc_view_nam = int(nam_nhap)
                    st.session_state.cc_view_bo_phan = bo_phan_nhap
                    st.session_state.cc_edit_mode = False
                    st.session_state.cc_data_loaded = False
                    st.rerun()
            
            st.caption("💡 Chọn tháng/năm và bộ phận (để trống = tất cả), sau đó bấm 'Mở BCC'")
        
        # ===== Bảng chấm công full-width =====
        else:
            import calendar as _cal
            thang_v = st.session_state.cc_view_thang
            nam_v = st.session_state.cc_view_nam
            bp_v = st.session_state.cc_view_bo_phan
            so_ngay = _cal.monthrange(nam_v, thang_v)[1]
            day_list = [date(nam_v, thang_v, d) for d in range(1, so_ngay + 1)]
            WD_ABBR = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]
            col_titles = [f"{d.day:02d} {WD_ABBR[d.weekday()]}" for d in day_list]
            sunday_cols = [t for d, t in zip(day_list, col_titles) if d.weekday() == 6]
            
            # Thanh điều khiển
            col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns([2.5, 0.7, 0.7, 0.7, 0.7])
            with col_h1:
                ten_bp = ", ".join(CHAM_CONG_DEPT_LABEL.get(b, b) for b in bp_v) if bp_v else "Tất cả bộ phận"
                st.markdown(f"**📅 {thang_v}/{nam_v} — {ten_bp}**")
            with col_h2:
                if st.button("◀️ Đóng", key="cc_close_btn", use_container_width=True):
                    st.session_state.cc_full_open = False
                    st.session_state.cc_pending_missing = None
                    st.rerun()
            with col_h3:
                edit_label = "👁️" if st.session_state.get('cc_edit_mode') else "✏️"
                if st.button(edit_label, key="cc_toggle_edit_btn", use_container_width=True):
                    st.session_state.cc_edit_mode = not st.session_state.get('cc_edit_mode', False)
                    st.session_state.cc_pending_missing = None
                    st.rerun()
            with col_h4:
                save_clicked = st.button(
                    "💾", key="cc_save_month_btn", type="primary", use_container_width=True,
                    disabled=not st.session_state.get('cc_edit_mode', False)
                )
            with col_h5:
                if st.button("📤 Xuất file", key="cc_export_btn", use_container_width=True):
                    st.session_state.cc_export_trigger = True
                    st.rerun()
            
            # Hướng dẫn
            with st.expander("📖 Hướng dẫn", expanded=False):
                st.markdown("""
                **Cách sử dụng bảng chấm công:**
                
                1. **Đánh dấu ngày nghỉ:** Nhập các ký hiệu vào ô tương ứng:
                   - `P` - Nghỉ phép hưởng lương
                   - `V` - Vắng mặt/Nghỉ không lương
                   - `NL` - Nghỉ lễ hưởng nguyên lương
                   - `0.5` - Làm nửa ngày công
                   - `CN` - Chủ nhật (tự động đánh dấu)
                
                2. **Đánh dấu ca làm việc:**
                   - `N` - Ca ngày (8 tiếng)
                   - `D` - Ca đêm (8 tiếng)
                   - `X` - Công thường (đủ ca)
                
                3. **Tăng ca:** Nhập số giờ vào dòng Tăng ca (VD: 4, 2.5)
                
                4. **🔥 Chấm công full:** Tại cột cuối cùng **"Chấm công full"**, click vào ô để chuyển thành `✅`, hệ thống sẽ tự động đánh dấu `X` cho tất cả các ô còn trống trên dòng Ca chính của nhân viên đó (trừ ngày Chủ nhật đã có `CN`).
                
                5. **Chủ nhật:** Tất cả ngày Chủ nhật được tự động đánh dấu `CN`.
                """)
            
            # Lấy danh sách nhân viên
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            if bp_v:
                c.execute("""SELECT id, ma_nv, ho_ten, chuc_danh_nghe, phong_ban_lam_viec FROM nhan_vien
                             WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND phong_ban_lam_viec = ANY(%s)
                             ORDER BY ma_nv ASC""", (bp_v,))
            else:
                c.execute("""SELECT id, ma_nv, ho_ten, chuc_danh_nghe, phong_ban_lam_viec FROM nhan_vien
                             WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY ma_nv ASC""")
            nv_list = c.fetchall()

            # Lấy dữ liệu chấm công hiện có
            existing = {}
            if nv_list:
                nv_ids = [nv['id'] for nv in nv_list]
                c.execute("""SELECT nhan_vien_id, ngay, ca_ngay, ca_dem, gio_tang_ca FROM cham_cong
                             WHERE nhan_vien_id = ANY(%s) AND EXTRACT(MONTH FROM ngay) = %s AND EXTRACT(YEAR FROM ngay) = %s""",
                          (nv_ids, thang_v, nam_v))
                for r in c.fetchall():
                    existing[(r['nhan_vien_id'], r['ngay'])] = r
            c.close(); db.close()

            if not nv_list:
                st.warning("Không có nhân viên nào phù hợp với bộ phận đã chọn.")
            else:
                # Xây dựng dữ liệu - Mỗi nhân viên có 1 hoặc 2 dòng
                flat_rows = []
                nv_row_indices = {}  # Lưu index của từng nhân viên
                
                for nv in nv_list:
                    dept = nv['phong_ban_lam_viec']
                    
                    # Xác định số dòng cho nhân viên này
                    if is_van_phong(dept):  # Dùng hàm mới thay vì so sánh in
                        # Văn phòng: chỉ 1 dòng
                        loai_list = ["Ca chính"]
                    else:
                        # Các bộ phận khác: 2 dòng (Ca chính + Tăng ca)
                        loai_list = ["Ca chính", "Tăng ca"]
                    
                    nv_indices = []
                    for idx, loai in enumerate(loai_list):
                        row = {
                            "Mã NV": nv['ma_nv'] if idx == 0 else "",
                            "Họ tên": nv['ho_ten'] if idx == 0 else "",
                            "Chức danh": (nv.get('chuc_danh_nghe') or "") if idx == 0 else "",
                            "Loại": loai,
                        }
                        # Thêm cột Chấm công full (chỉ ở dòng Ca chính)
                        row["Chấm công full"] = "" if idx > 0 else "⬜"
                        
                        for d, title in zip(day_list, col_titles):
                            rec = existing.get((nv['id'], d))
                            if loai == "Ca chính":
                                # Lấy ca_ngay hoặc ca_dem (ưu tiên ca_ngay)
                                val = rec.get('ca_ngay') if rec is not None else ""
                                if not val and rec is not None:
                                    val = rec.get('ca_dem') or ""
                                # Nếu là Chủ nhật và chưa có dữ liệu -> tự động đánh dấu CN
                                if d.weekday() == 6 and not val:
                                    val = "CN"
                                row[title] = val
                            else:  # Tăng ca
                                tc_val = rec.get('gio_tang_ca') if rec is not None else None
                                row[title] = "" if not tc_val else str(tc_val)
                        flat_rows.append(row)
                        nv_indices.append(len(flat_rows) - 1)
                    
                    nv_row_indices[nv['ma_nv']] = {
                        'ca_main': nv_indices[0],  # Index của dòng Ca chính
                        'tc': nv_indices[1] if len(nv_indices) > 1 else None  # Index của dòng Tăng ca
                    }

                flat_rows = [row for row in flat_rows if row.get("Họ tên", "").strip() != ""]
                
                # Đảm bảo cột "Chấm công full" luôn có mặt
                df_month = pd.DataFrame(flat_rows)
                if "Chấm công full" not in df_month.columns:
                    # Thêm cột vào vị trí sau cột Loại
                    cols = df_month.columns.tolist()
                    # Tìm vị trí của cột Loại
                    loai_idx = cols.index("Loại") if "Loại" in cols else 3
                    cols.insert(loai_idx + 1, "Chấm công full")
                    df_month = df_month.reindex(columns=cols)
                    # Điền giá trị mặc định
                    df_month["Chấm công full"] = df_month.apply(
                        lambda row: "" if row["Loại"] == "Tăng ca" else "⬜", axis=1
                    )
                
                CC_MAX_VISIBLE_ROWS = 30
                CC_HEADER_H = 38
                table_height = CC_HEADER_H + CC_ROW_HEIGHT * min(len(df_month), CC_MAX_VISIBLE_ROWS)

                # Xử lý export file
                if st.session_state.get('cc_export_trigger', False):
                    st.session_state.cc_export_trigger = False
                    try:
                        from openpyxl import Workbook
                        from openpyxl.styles import Font, Alignment, Border, Side
                        from openpyxl.utils import get_column_letter
                        
                        wb = Workbook()
                        ws = wb.active
                        ws.title = f"Cham_cong_{thang_v}_{nam_v}"
                        
                        # Ghi header
                        for col_idx, col_name in enumerate(df_month.columns, 1):
                            cell = ws.cell(row=1, column=col_idx, value=col_name)
                            cell.font = Font(bold=True, size=10)
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            # Đặt độ rộng cột
                            if col_name in ["Mã NV", "Loại"]:
                                ws.column_dimensions[get_column_letter(col_idx)].width = 12
                            elif col_name == "Họ tên":
                                ws.column_dimensions[get_column_letter(col_idx)].width = 25
                            elif col_name == "Chức danh":
                                ws.column_dimensions[get_column_letter(col_idx)].width = 20
                            elif col_name == "Chấm công full":
                                ws.column_dimensions[get_column_letter(col_idx)].width = 15
                            else:
                                ws.column_dimensions[get_column_letter(col_idx)].width = 8
                        
                        # Ghi dữ liệu
                        for row_idx, row_data in enumerate(df_month.itertuples(index=False), 2):
                            for col_idx, val in enumerate(row_data, 1):
                                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                        filename = f"Cham_cong_{thang_v}_{nam_v}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        wb.save(filename)
                        
                        with open(filename, "rb") as f:
                            st.download_button(
                                label="📥 TẢI FILE EXCEL",
                                data=f,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                width='stretch'
                            )
                        st.success(f"✅ Đã xuất file: {filename}")
                    except Exception as e:
                        st.error(f"❌ Lỗi xuất file: {e}")

                if not st.session_state.get('cc_edit_mode', False):
                    # Chế độ XEM - LOẠI BỎ cột "Chấm công full"
                    def _highlight_sunday(s):
                        if s.name in sunday_cols:
                            return ['background-color: #FFF2CC; font-weight: bold; color: #999;'] * len(s)
                        return [''] * len(s)
                    
                    def _center_style(s):
                        return ['text-align: center; vertical-align: middle;' for _ in s]
                    
                    # Tạo bảng không có cột "Chấm công full"
                    df_view = df_month.drop(columns=["Chấm công full"], errors='ignore')

                    view_col_cfg = {
                        "Mã NV": cc_pin_col(st.column_config.TextColumn, width="small"),
                        "Họ tên": cc_pin_col(st.column_config.TextColumn, width=180),
                        "Chức danh": cc_pin_col(st.column_config.TextColumn, width=140),
                        "Loại": cc_pin_col(st.column_config.TextColumn, width="small"),
                        # KHÔNG CÓ "Chấm công full" ở chế độ xem
                    }
                    styled = (
                        df_view.style
                        .apply(_highlight_sunday, axis=0)
                        .apply(_center_style, axis=0)
                        .set_properties(**{"text-align": "center", "vertical-align": "middle"})
                        .hide(axis="index")
                    )
                    cc_render_grid(
                        styled, edit=False, width='stretch', height=table_height,
                        column_config=view_col_cfg,
                    )
                    st.caption("👁️ Xem | ✏️ Bấm nút bút chì để sửa | 📤 Xuất file Excel")
                else:
                    # Chế độ SỬA - sử dụng data_editor với cột Chấm công full là checkbox
                    col_cfg = {
                        "Mã NV": cc_pin_col(st.column_config.TextColumn, disabled=True, width="small"),
                        "Họ tên": cc_pin_col(st.column_config.TextColumn, disabled=True, width=180),
                        "Chức danh": cc_pin_col(st.column_config.TextColumn, disabled=True, width=140),
                        "Loại": cc_pin_col(st.column_config.TextColumn, disabled=True, width="small"),
                        "Chấm công full": cc_pin_col(st.column_config.CheckboxColumn, width="small"),
                    }
                    for t in col_titles:
                        col_cfg[t] = st.column_config.TextColumn(width="small", validate=CHAM_CONG_CELL_REGEX)

                    edit_key = f"cc_month_editor_{thang_v}_{nam_v}_{'-'.join(bp_v) if bp_v else 'all'}"
                    
                    # Hiển thị bảng chỉnh sửa
                    edited_df = st.data_editor(
                        df_month,
                        column_config=col_cfg,
                        hide_index=True,
                        num_rows="fixed",
                        width='stretch',
                        height=table_height,
                        key=edit_key,
                        use_container_width=True,
                    )
                    
                    # Xử lý khi user click vào checkbox "Chấm công full"
                    if edited_df is not None:
                        # Kiểm tra từng dòng Ca chính
                        for idx, row in edited_df.iterrows():
                            if row["Loại"] == "Ca chính":
                                nv_ma = row["Mã NV"]
                                # So sánh giá trị checkbox
                                old_val = df_month.iloc[idx]["Chấm công full"]
                                new_val = row["Chấm công full"]
                                
                                # Nếu checkbox được tick (chuyển từ False sang True)
                                if new_val is True and old_val is not True:  # SỬA: dùng "is not True"
                                    # Tự động đánh dấu X cho tất cả ô trống trên dòng Ca chính
                                    ca_main_idx = idx
                                    for d, title in zip(day_list, col_titles):
                                        if d.weekday() == 6:  # Bỏ qua Chủ nhật (đã có CN)
                                            continue
                                        current_val = str(edited_df.iloc[ca_main_idx][title] or "").strip()
                                        if not current_val:
                                            edited_df.at[ca_main_idx, title] = "X"
                                    
                                    # Reset checkbox sau khi đã xử lý
                                    edited_df.at[idx, "Chấm công full"] = False
                                    st.rerun()
                                    
                            # Nếu checkbox được bỏ tick (chuyển từ True sang False) thì không làm gì

                    # Xử lý lưu
                    if save_clicked:
                        # Kiểm tra thiếu dữ liệu
                        missing = []
                        for nv_ma, indices in nv_row_indices.items():
                            if 'ca_main' not in indices:
                                continue
                            nv_ten = edited_df.iloc[indices['ca_main']]["Họ tên"]
                            ca_main_idx = indices['ca_main']
                            for d, title in zip(day_list, col_titles):
                                if d.weekday() == 6:  # Bỏ qua Chủ nhật
                                    continue
                                if d >= date.today():  # Bỏ qua ngày tương lai
                                    continue
                                val = str(edited_df.iloc[ca_main_idx][title] or "").strip()
                                if not val:
                                    missing.append((nv_ma, nv_ten, d))

                        if missing and not st.session_state.get('cc_force_save_approved', False):
                            st.warning(f"⚠️ Có {len(missing)} lượt chưa chấm công")
                            with st.expander("Xem chi tiết"):
                                for ma_nv, ho_ten, d in missing[:100]:
                                    st.caption(f"- {ma_nv} - {ho_ten}: {d.strftime('%d/%m/%Y')}")
                                if len(missing) > 100:
                                    st.caption(f"... và {len(missing) - 100} lượt khác")
                            col_cf1, col_cf2 = st.columns(2)
                            with col_cf1:
                                if st.button("✅ Vẫn lưu", key="cc_force_save_btn", type="primary", use_container_width=True):
                                    st.session_state.cc_force_save_approved = True
                                    st.rerun()
                            with col_cf2:
                                if st.button("✏️ Sửa tiếp", key="cc_cancel_save_btn", use_container_width=True):
                                    st.session_state.cc_force_save = False
                                    st.session_state.cc_force_save_approved = False
                                    st.rerun()
                        else:
                            # Thực hiện lưu
                            db2 = st.session_state.db_engine.get_connection()
                            c2 = db2.cursor()
                            n_saved = 0
                            
                            c2_nv = db2.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                            if bp_v:
                                c2_nv.execute("""SELECT id, ma_nv FROM nhan_vien
                                                 WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND phong_ban_lam_viec = ANY(%s)
                                                 ORDER BY ma_nv ASC""", (bp_v,))
                            else:
                                c2_nv.execute("""SELECT id, ma_nv FROM nhan_vien
                                                 WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY ma_nv ASC""")
                            nv_map = {row['ma_nv']: row['id'] for row in c2_nv.fetchall()}
                            c2_nv.close()
                            
                            for idx, row in edited_df.iterrows():
                                nv_ma = row.get("Mã NV", "")
                                if not nv_ma or nv_ma not in nv_map:
                                    continue
                                nv_id = nv_map[nv_ma]
                                loai = row.get("Loại", "")
                                
                                for d, title in zip(day_list, col_titles):
                                    v_ngay = None
                                    v_dem = None
                                    v_tc = 0
                                    
                                    if loai == "Ca chính":
                                        val = str(row[title] or "").strip()
                                        if val:
                                            # Xác định ca ngày hay ca đêm
                                            if val.upper() == "N":
                                                v_ngay = "N"
                                            elif val.upper() == "D":
                                                v_dem = "D"
                                            elif val.upper() == "CN":
                                                v_ngay = "CN"
                                            else:
                                                v_ngay = cc_normalize_marker(val) if val else None
                                    elif loai == "Tăng ca":
                                        val = str(row[title] or "").strip()
                                        try:
                                            v_tc = float(val.replace(",", ".")) if val else 0
                                        except ValueError:
                                            v_tc = 0
                                    
                                    # Bỏ qua nếu tất cả đều trống và chưa có trong DB
                                    if v_ngay is None and v_dem is None and v_tc == 0 and (nv_id, d) not in existing:
                                        continue
                                    
                                    c2.execute("""
                                        INSERT INTO cham_cong (nhan_vien_id, ngay, ca_ngay, ca_dem, gio_tang_ca, nguon, created_by, updated_at)
                                        VALUES (%s, %s, %s, %s, %s, 'THU_CONG', %s, NOW())
                                        ON CONFLICT (nhan_vien_id, ngay) DO UPDATE SET
                                            ca_ngay = EXCLUDED.ca_ngay,
                                            ca_dem = EXCLUDED.ca_dem,
                                            gio_tang_ca = EXCLUDED.gio_tang_ca,
                                            updated_at = NOW()
                                    """, (nv_id, d, v_ngay, v_dem, v_tc, st.session_state.username))
                                    n_saved += 1
                            
                            db2.commit()
                            c2.close(); db2.close()
                            st.success(f"✅ Đã lưu {n_saved} lượt chấm công tháng {thang_v}/{nam_v}.")
                            st.session_state.cc_edit_mode = False
                            st.session_state.cc_force_save = False
                            st.session_state.cc_force_save_approved = False
                            st.rerun()

    # ========== 2. TRÍCH XUẤT TỪ MÁY CHẤM VÂN TAY ==========
    elif st.session_state.get('cc_method') == 'fingerprint':
        st.info("""
        ### 🚧 Tính năng đang phát triển
        
        Dự kiến hỗ trợ:
        - Upload file dữ liệu xuất từ máy chấm vân tay (.xls/.csv)
        - Ánh xạ mã nhân viên trên máy chấm công với Mã NV
        - Tự động quy đổi giờ vào/ra thành mã công
        """)

    # ========== 3. FACE ID ==========
    elif st.session_state.get('cc_method') == 'faceid':
        st.info("""
        ### 🚧 Tính năng đang phát triển
        
        Dự kiến hỗ trợ:
        - ✅ Đăng ký khuôn mặt cho nhân viên
        - ✅ Chấm công bằng camera
        - ✅ Lịch sử chấm công theo thời gian thực
        """)

# ========== TÍNH THU NHẬP ==========
elif menu == "💰 Tính thu nhập":
    st.title("💰 Tính thu nhập (Lương & Phụ cấp)")
    st.caption("Tính toán lương, thưởng và các khoản phụ cấp cho nhân viên")
    
    st.info("""
    ### 🚧 Tính năng đang hoàn thiện
    
    Nội dung đang được phát triển. Các tính năng sắp ra mắt:
    - ✅ Tính lương cơ bản theo hệ số
    - ✅ Tính các khoản phụ cấp (chức vụ, thâm niên, trách nhiệm...)
    - ✅ Tính thuế TNCN
    - ✅ Tính các khoản khấu trừ (BHXH, BHYT, BHTN, đoàn phí...)
    - ✅ Tổng hợp bảng lương tháng
    - ✅ Xuất bảng lương Excel/PDF
    - ✅ Gửi bảng lương qua email/Zalo cho nhân viên
    
    ⏳ **Dự kiến hoàn thành: Quý 4/2026**
    """)
    
    # Thêm form tính thử nghiệm demo
    with st.expander("🧪 Thử nghiệm tính lương (Demo)"):
        db_demo = st.session_state.db_engine.get_connection()
        c_demo = db_demo.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c_demo.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY ho_ten LIMIT 10")
        nv_list = c_demo.fetchall()
        db_demo.close()
        
        if nv_list:
            nv_options = {f"{nv['ma_nv']} - {nv['ho_ten']}": nv['id'] for nv in nv_list}
            selected_nv = st.selectbox("Chọn nhân viên để tính thử:", list(nv_options.keys()), help="💡 Gõ mã NV hoặc tên để tìm nhanh trong danh sách")
            
            col_luong1, col_luong2 = st.columns(2)
            with col_luong1:
                luong_co_ban = st.number_input("Lương cơ bản (VNĐ)", min_value=0, value=5000000, step=500000)
                phu_cap_chuc_vu = st.number_input("Phụ cấp chức vụ (VNĐ)", min_value=0, value=0, step=100000)
            with col_luong2:
                phu_cap_tnvk = st.number_input("Phụ cấp thâm niên VK (%)", min_value=0.0, value=0.0, step=0.5)
                phu_cap_tnn = st.number_input("Phụ cấp thâm niên nghề (%)", min_value=0.0, value=0.0, step=0.5)
            
            tong_luong = luong_co_ban + phu_cap_chuc_vu
            tong_luong += luong_co_ban * phu_cap_tnvk / 100
            tong_luong += luong_co_ban * phu_cap_tnn / 100
            
            st.markdown("---")
            st.subheader("📊 Kết quả tính thử:")
            
            col_kq1, col_kq2, col_kq3 = st.columns(3)
            col_kq1.metric("Lương cơ bản", f"{luong_co_ban:,.0f} VNĐ")
            col_kq2.metric("Phụ cấp", f"{(tong_luong - luong_co_ban):,.0f} VNĐ")
            col_kq3.metric("Tổng thu nhập", f"{tong_luong:,.0f} VNĐ")
            
            st.caption("⚠️ Đây chỉ là tính toán tham khảo. Tính năng chính thức sẽ tích hợp với dữ liệu nhân viên và chấm công.")
        else:
            st.warning("Chưa có nhân viên nào trong hệ thống để thử nghiệm!")

# ========== UPLOAD ==========
elif menu=="📁 Upload hồ sơ" and st.session_state.role=="admin":
    st.title("📁 Quản lý hồ sơ nhân viên")
    tab_upload, tab_list, tab_avatar = st.tabs(["📤 UPLOAD HỒ SƠ", "📋 DANH SÁCH HỒ SƠ", "📸 UPLOAD ẢNH HỒ SƠ"])
    
    with tab_avatar:
        st.subheader("📸 Upload ảnh hồ sơ cho nhân viên")
        st.caption("Chọn nhân viên và tải ảnh lên. Ảnh sẽ được lưu vào cột `anh_ho_so` trong bảng nhân viên.")
        
        # Lấy danh sách nhân viên chưa có ảnh hồ sơ
        db_avatar = st.session_state.db_engine.get_connection()
        c_avatar = db_avatar.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c_avatar.execute("""
            SELECT id, ma_nv, ho_ten FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') 
            AND (anh_ho_so IS NULL OR anh_ho_so = '')
            ORDER BY id DESC
        """)
        nv_chua_anh = c_avatar.fetchall()
        db_avatar.close()

        if not nv_chua_anh:
            st.success("🎉 Tất cả nhân viên đã có ảnh hồ sơ!")
        else:
            # Tạo dict chọn nhân viên
            nv_map = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in nv_chua_anh}
            selected_nv_label = st.selectbox("📌 Chọn nhân viên cần upload ảnh:", list(nv_map.keys()))
            selected_nv_id = nv_map[selected_nv_label]
            
            # Lấy thông tin nhân viên đã chọn
            db_detail = st.session_state.db_engine.get_connection()
            c_detail = db_detail.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c_detail.execute("SELECT ma_nv, ho_ten FROM nhan_vien WHERE id = %s", (selected_nv_id,))
            nv_info = c_detail.fetchone()
            db_detail.close()

            if nv_info:
                st.markdown(f"**Nhân viên:** {nv_info['ma_nv']} - {nv_info['ho_ten']}")
                
                # File uploader cho ảnh
                anh_upload = st.file_uploader("Chọn ảnh hồ sơ (png, jpg, jpeg)", type=["png", "jpg", "jpeg"], key="avatar_upload_single")
                
                if anh_upload is not None:
                    st.image(anh_upload, caption="Ảnh xem trước", width=200)
                
                if st.button("📤 UPLOAD ẢNH", type="primary", width='stretch'):
                    if anh_upload is None:
                        st.error("❌ Vui lòng chọn ảnh để upload!")
                    else:
                        # Upload ảnh lên Storage
                        storage_path = upload_anh_ho_so(nv_info['ma_nv'], nv_info['ho_ten'], anh_upload)
                        if storage_path:
                            # Cập nhật đường dẫn vào bảng nhan_vien
                            db_update = st.session_state.db_engine.get_connection()
                            c_update = db_update.cursor()
                            c_update.execute("UPDATE nhan_vien SET anh_ho_so = %s WHERE id = %s", (storage_path, selected_nv_id))
                            db_update.commit()
                            db_update.close()
                            st.success(f"✅ Đã upload ảnh hồ sơ thành công cho {nv_info['ho_ten']}!")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("❌ Upload ảnh thất bại. Vui lòng kiểm tra cấu hình Storage!")
    
    with tab_upload:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY id DESC")
        nvl = c.fetchall()
        db.close()
        
        if nvl:
            nd = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in nvl}
            id_to_hoten = {x['id']: x['ho_ten'] for x in nvl}
            cn = st.selectbox("📌 Chọn nhân viên:", list(nd.keys()))
            
            col1, col2 = st.columns(2)
            with col1:
                lh = st.selectbox("📂 Loại hồ sơ:", ["BANG_CAP", "CHUNG_CHI", "CCCD", "HOP_DONG", "SO_YEU_LY_LICH", "KHAC"])
            with col2:
                st.markdown("")
                st.caption("💡 **Hướng dẫn: Chọn loại giấy tờ, sau đó chọn file từ thư mục để Upload!**")
            
            fl = st.file_uploader("📎 Chọn file:", type=['pdf', 'jpg', 'png', 'jpeg', 'doc', 'docx'])
            
            if fl:
                st.info(f"📄 Tên file: {fl.name} | 📏 Kích thước: {fl.size/1024:.1f} KB")
            
            if fl and st.button("📤 UPLOAD", type="primary", width='stretch'):
                nid = nd[cn]
                ngay_upload_str = datetime.now().strftime('%Y%m%d')
                safe_name = sanitize_storage_filename(fl.name)
                ho_ten_folder = sanitize_storage_filename(id_to_hoten.get(nid, str(nid)))
                # Cấu trúc: {Họ tên nhân viên}/{Loại hồ sơ}_{ngày upload}_{tên file}
                base_path = f"{ho_ten_folder}/{lh}_{ngay_upload_str}_{safe_name}"

                sb = get_supabase_storage()
                if not sb:
                    st.error("❌ Chưa cấu hình Supabase Storage. Vui lòng khai báo `SUPABASE_URL` và `SUPABASE_KEY` trong secrets/.env.")
                else:
                    try:
                        storage_path = upload_to_storage_unique(
                            sb, SUPABASE_BUCKET, base_path,
                            fl.getvalue(), fl.type
                        )

                        db = st.session_state.db_engine.get_connection()
                        c = db.cursor()
                        c.execute("""
                            INSERT INTO ho_so_nhan_vien (nhan_vien_id, loai_ho_so, ten_file, duong_dan_file, ngay_upload) 
                            VALUES (%s, %s, %s, %s, CURRENT_DATE)
                        """, (nid, lh, fl.name, storage_path))
                        db.commit()
                        db.close()

                        st.success(f"✅ Đã upload thành công lên Supabase Storage!\n📁 Đường dẫn: {storage_path}")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Lỗi khi upload lên Supabase Storage: {e}")
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")
    
    with tab_list:
        st.subheader("📋 Danh sách hồ sơ đã upload")
        
        # --- HÀM CACHE CHO PRESIGNED URL ---
        @st.cache_data(ttl=3600, show_spinner=False)
        def get_presigned_url_cached(storage_path: str) -> str:
            """Tạo presigned URL có cache 1 giờ"""
            if not storage_path:
                return ""
            try:
                sb = get_supabase_storage()
                if not sb:
                    return ""
                # Tạo signed URL có hiệu lực 1 giờ
                return sb.storage.from_(SUPABASE_BUCKET).create_signed_url(storage_path, expires_in=3600)
            except Exception as e:
                print(f"Lỗi tạo presigned URL: {e}")
                return ""
        
        @st.cache_data(ttl=3600, show_spinner=False)
        def get_file_bytes_cached(storage_path: str) -> bytes:
            """Tải file bytes có cache 1 giờ"""
            if not storage_path:
                return None
            try:
                sb = get_supabase_storage()
                if not sb:
                    return None
                return sb.storage.from_(SUPABASE_BUCKET).download(storage_path)
            except Exception as e:
                print(f"Lỗi tải file: {e}")
                return None
        
        # --- Lấy danh sách nhân viên ---
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien ORDER BY id DESC")
        nvl = c.fetchall()
        db.close()
        
        if nvl:
            nd = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in nvl}
            selected_nv = st.selectbox("🔍 Chọn nhân viên để xem hồ sơ:", list(nd.keys()), key="view_hoso")
            nv_id = nd[selected_nv]
            
            # --- Lấy danh sách hồ sơ ---
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("""
                SELECT id, loai_ho_so, ten_file, duong_dan_file, ngay_upload 
                FROM ho_so_nhan_vien 
                WHERE nhan_vien_id = %s 
                ORDER BY ngay_upload DESC, id DESC
            """, (nv_id,))
            hs_list = c.fetchall()
            db.close()
            
            if hs_list:
                # --- Hiển thị danh sách (KHÔNG tải file) ---
                hs_data = []
                for i, hs in enumerate(hs_list, 1):
                    hs_data.append({
                        "STT": i,
                        "Loại hồ sơ": hs['loai_ho_so'],
                        "Tên file gốc": hs['ten_file'],
                        "Ngày upload": format_date(hs['ngay_upload']),
                        "ID": hs['id'],
                        "Đường dẫn": hs['duong_dan_file']
                    })
                df_hs = pd.DataFrame(hs_data)
                st.dataframe(df_hs[['STT', 'Loại hồ sơ', 'Tên file gốc', 'Ngày upload']], 
                            width='stretch', hide_index=True)
                
                st.divider()
                if st.button("❌ THOÁT", width='stretch', key="exit_hoso_list"):
                    st.session_state.pop('selected_nv', None)
                    st.rerun()
                
                # --- Chọn hồ sơ để xem (ON-DEMAND LOADING) ---
                hs_options = {f"{hs['loai_ho_so']} - {hs['ten_file']}": hs for hs in hs_list}
                selected_hs_name = st.selectbox("Chọn hồ sơ:", list(hs_options.keys()), key="select_hs_preview")
                selected_hs = hs_options[selected_hs_name]
                
                # Hiển thị thông tin hồ sơ
                st.markdown(f"""
                **📄 {selected_hs['loai_ho_so']}** - {selected_hs['ten_file']}  
                📅 {format_date(selected_hs['ngay_upload'])}
                """)
                
                # --- Xác định loại file ---
                file_ext = selected_hs['ten_file'].lower().split('.')[-1]
                is_image = file_ext in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp']
                is_pdf = file_ext == 'pdf'
                
                # --- State keys ---
                preview_state_key = f"preview_active_{selected_hs['id']}"
                download_url_key = f"download_url_{selected_hs['id']}"
                preview_data_key = f"preview_data_{selected_hs['id']}"
                
                # Khởi tạo state nếu chưa có
                if preview_state_key not in st.session_state:
                    st.session_state[preview_state_key] = False
                
                # --- Nút PREVIEW (On-demand) ---
                col_preview_btn, col_download_btn = st.columns(2)
                
                with col_preview_btn:
                    if is_image or is_pdf:
                        if st.button("👁️ PREVIEW", width='stretch', type="secondary", 
                                    key=f"preview_btn_{selected_hs['id']}"):
                            # CHỈ TẢI KHI BẤM NÚT
                            with st.spinner("⏳ Đang tải file preview..."):
                                file_bytes = get_file_bytes_cached(selected_hs['duong_dan_file'])
                                if file_bytes:
                                    st.session_state[preview_data_key] = file_bytes
                                    st.session_state[preview_state_key] = True
                                    st.rerun()
                                else:
                                    st.error("❌ Không thể tải file để preview")
                    else:
                        st.button("👁️ PREVIEW", disabled=True, width='stretch', 
                                 help="Không thể preview loại file này")
                
                with col_download_btn:
                    if st.button("📥 TẢI HỒ SƠ", width='stretch', 
                               key=f"download_btn_{selected_hs['id']}"):
                        # CHỈ TẠO URL KHI BẤM NÚT
                        with st.spinner("⏳ Đang tạo link tải..."):
                            url = get_presigned_url_cached(selected_hs['duong_dan_file'])
                            if url:
                                st.session_state[download_url_key] = url
                                st.rerun()
                            else:
                                st.error("❌ Không thể tạo link tải")
                
                # --- HIỂN THỊ PREVIEW (nếu có) ---
                if st.session_state.get(preview_state_key, False):
                    file_bytes = st.session_state.get(preview_data_key)
                    if file_bytes:
                        st.markdown("---")
                        st.subheader("📄 Xem trước")
                        
                        if is_image:
                            img_base64 = base64.b64encode(file_bytes).decode()
                            st.image(f"data:image/jpeg;base64,{img_base64}", width=400)
                        elif is_pdf:
                            pdf_base64 = base64.b64encode(file_bytes).decode()
                            st.markdown(f"""
                            <iframe src="data:application/pdf;base64,{pdf_base64}" 
                                    width="100%" height="600px" style="border:none;border-radius:8px;">
                            </iframe>
                            """, unsafe_allow_html=True)
                        
                        # Nút đóng preview
                        if st.button("❌ Đóng preview", key=f"close_preview_{selected_hs['id']}"):
                            st.session_state[preview_state_key] = False
                            if preview_data_key in st.session_state:
                                del st.session_state[preview_data_key]
                            st.rerun()
                    else:
                        st.warning("⚠️ Không có dữ liệu preview")
                
                # --- HIỂN THỊ LINK TẢI (nếu có) ---
                if st.session_state.get(download_url_key):
                    st.markdown("---")
                    st.success("✅ Link tải đã sẵn sàng (có hiệu lực 1 giờ)")
                    url = st.session_state[download_url_key]
                    st.markdown(f"""
                    <div style="background:#f0fdf4;padding:12px 16px;border-radius:8px;border:1px solid #bbf7d0;">
                        <a href="{url}" target="_blank" style="font-size:16px;font-weight:600;color:#166534;">
                            📥 Tải xuống: {selected_hs['ten_file']}
                        </a>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if st.button("❌ Đóng link tải", key=f"close_download_{selected_hs['id']}"):
                        if download_url_key in st.session_state:
                            del st.session_state[download_url_key]
                        st.rerun()
                
                # --- Nút xóa hồ sơ ---
                st.divider()
                col_del1, col_del2, col_del3 = st.columns([1, 2, 1])
                with col_del2:
                    if st.button("🗑️ XÓA HỒ SƠ NÀY", width='stretch', type="secondary"):
                        try:
                            sb = get_supabase_storage()
                            if sb:
                                try:
                                    sb.storage.from_(SUPABASE_BUCKET).remove([selected_hs['duong_dan_file']])
                                except Exception as e_storage:
                                    st.warning(f"⚠️ Không xóa được file trên Storage (vẫn xóa bản ghi): {e_storage}")
                            db = st.session_state.db_engine.get_connection()
                            c = db.cursor()
                            c.execute("DELETE FROM ho_so_nhan_vien WHERE id = %s", (selected_hs['id'],))
                            db.commit()
                            db.close()
                            # Xóa cache
                            get_presigned_url_cached.clear()
                            get_file_bytes_cached.clear()
                            st.success(f"✅ Đã xóa hồ sơ: {selected_hs['ten_file']}")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Lỗi khi xóa: {e}")
            else:
                st.info(f"📭 Nhân viên này chưa có hồ sơ nào được upload.")
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")

# ========== DANH MỤC CHỨC DANH ==========
elif menu == "⚙️ Danh mục" and st.session_state.role == "admin":
    st.title("⚙️ Danh mục cấu hình theo doanh nghiệp")
    st.caption("Mỗi khách hàng tự đặt tên Phòng ban, Chức danh, Loại hợp đồng, Trình độ học vấn phù hợp với cơ cấu công ty mình — không ảnh hưởng đến khách hàng khác.")

    def _quan_ly_danh_muc_don_gian(ten_bang, cot_ten, tieu_de, placeholder):
        """Hàm dùng chung để quản lý CRUD cho các bảng danh mục dạng đơn giản
        (id, cột tên, thu_tu, trang_thai) — tránh lặp code cho từng loại danh mục."""
        with st.expander(f"➕ Thêm {tieu_de.lower()} mới", expanded=False):
            ten_moi = st.text_input("Tên", key=f"add_{ten_bang}", placeholder=placeholder)
            if st.button("💾 Lưu", key=f"btn_add_{ten_bang}"):
                if ten_moi.strip():
                    try:
                        if ten_bang == "danh_muc_phong_ban":
                            # Phòng ban: dùng chuẩn hóa kiểu tiếng Việt (không viết hoa mọi từ)
                            ten_chuan_hoa = chuan_hoa_ten_phong_ban(ten_moi)
                        else:
                            ten_chuan_hoa = ten_moi.strip()[:1].upper() + ten_moi.strip()[1:]  # chỉ viết hoa chữ cái đầu
                        db = st.session_state.db_engine.get_connection(); c = db.cursor()
                        c.execute(f"INSERT INTO {ten_bang} ({cot_ten}) VALUES (%s) ON CONFLICT DO NOTHING",
                                  (ten_chuan_hoa,))
                        db.commit(); db.close()
                        st.success(f"✅ Đã thêm: {ten_chuan_hoa}"); st.cache_data.clear(); st.rerun()
                    except Exception as e:
                        st.error(f"❌ Lỗi: {e}")
                else:
                    st.error("Vui lòng nhập tên!")

        db = st.session_state.db_engine.get_connection(); c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute(f"SELECT id, {cot_ten}, trang_thai FROM {ten_bang} ORDER BY thu_tu, id")
        ds = c.fetchall(); db.close()
        if ds:
            df = pd.DataFrame(ds); df.columns = ['ID', tieu_de, 'Trạng thái']
            st.dataframe(df, width='stretch', hide_index=True)
            idx_xoa = st.number_input("Nhập ID cần xoá:", min_value=1, step=1, key=f"del_{ten_bang}")
            if st.button("🗑️ Xoá", key=f"btn_del_{ten_bang}"):
                if not can_delete():
                    st.error("❌ Bạn không có quyền xóa dữ liệu!")
                else:
                    db = st.session_state.db_engine.get_connection()
                    c = db.cursor()
                    c.execute(f"DELETE FROM {ten_bang} WHERE id=%s", (idx_xoa,))
                    db.commit()
                    db.close()
                    st.success("🗑️ Đã xoá!")
                    st.cache_data.clear()
                    st.rerun()
        else:
            st.info(f"Chưa có {tieu_de.lower()} nào.")

    tab_pb, tab_cd, tab_hd, tab_hv, tab_mau_hd, tab_cv, tab_cty = st.tabs([
        "🏢 Phòng ban", "💼 Chức danh", "📄 Loại hợp đồng", "🎓 Trình độ học vấn",
        "📃 Mẫu Hợp đồng", "🎖️ Chức vụ", "⚙️ Cấu hình Doanh nghiệp"
    ])

    with tab_pb:
        _quan_ly_danh_muc_don_gian("danh_muc_phong_ban", "ten_phong_ban", "Phòng ban", "VD: Kinh doanh")

    with tab_cd:
        # Chức danh tiếp tục dùng bảng vi_tri_cong_tac có sẵn để không phá vỡ dữ liệu cũ
        with st.expander("➕ Thêm chức danh mới", expanded=False):
            with st.form("add_chuc_danh"):
                ten_moi = st.text_input("Tên chức danh *"); mo_ta = st.text_area("Mô tả")
                if st.form_submit_button("💾 LƯU"):
                    if ten_moi:
                        db = st.session_state.db_engine.get_connection(); c = db.cursor()
                        c.execute("SELECT COALESCE(MIN(t1.id + 1), 1) FROM vi_tri_cong_tac t1 LEFT JOIN vi_tri_cong_tac t2 ON t1.id + 1 = t2.id WHERE t2.id IS NULL AND t1.id >= 1")
                        id_trong = c.fetchone()[0]
                        c.execute("SELECT COALESCE(MAX(id),0) FROM vi_tri_cong_tac")
                        id_max = c.fetchone()[0]
                        id_moi = id_trong if id_trong <= id_max + 1 else id_max + 1
                        c.execute("INSERT INTO vi_tri_cong_tac (id, ten_vi_tri, ghi_chu) VALUES (%s, %s, %s)", (id_moi, ten_moi, mo_ta))
                        db.commit(); db.close(); st.success(f"✅ Đã thêm: {ten_moi}"); st.cache_data.clear(); st.rerun()
                    else: st.error("Tên chức danh không được để trống!")
        st.subheader("📋 Danh sách chức danh")
        db = st.session_state.db_engine.get_connection(); c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ten_vi_tri, ghi_chu FROM vi_tri_cong_tac ORDER BY id")
        ds = c.fetchall(); db.close()
        if ds:
            df = pd.DataFrame(ds); df.columns = ['ID', 'Tên chức danh', 'Ghi chú']; st.dataframe(df, width='stretch', hide_index=True)
            st.divider(); cdx = st.number_input("Nhập ID cần xóa:", min_value=1, step=1)
            if st.button("🗑️ XÓA", key="del_cd"):
                db = st.session_state.db_engine.get_connection(); c = db.cursor()
                c.execute("DELETE FROM vi_tri_cong_tac WHERE id=%s", (cdx,)); db.commit(); db.close(); st.success("🗑️ Đã xóa!"); st.cache_data.clear(); st.rerun()
        else: st.info("Chưa có chức danh nào")

    with tab_hd:
        _quan_ly_danh_muc_don_gian("danh_muc_loai_hop_dong", "ten_loai_hd", "Loại hợp đồng", "VD: Hợp đồng thời vụ")

    with tab_hv:
        _quan_ly_danh_muc_don_gian("danh_muc_trinh_do_hoc_van", "ten_trinh_do", "Trình độ học vấn", "VD: Cử nhân")

    with tab_mau_hd:
        st.caption("Tuỳ chỉnh nội dung từng Điều trong Hợp đồng lao động (HĐLĐ) và Hợp đồng thử việc (HĐTV). "
                   "Điều nào chưa tuỳ chỉnh sẽ tự dùng nội dung mặc định. "
                   "Có thể dùng {vi_tri}, {ngay_hieu_luc} (HĐLĐ - Điều 1) hoặc {vi_tri}, {ngay_bat_dau}, {ngay_ket_thuc} (HĐTV - Điều 1) "
                   "— hệ thống sẽ tự thay bằng thông tin thực tế của từng nhân viên khi in. "
                   "Dòng bắt đầu bằng '## ' sẽ in đậm làm tiêu đề phụ (VD: '## 1. Nghĩa vụ:').")

        loai_hd_chon = st.radio("Chọn loại hợp đồng:", ["HĐLĐ (không xác định thời hạn)", "HĐTV (thử việc)"],
                                 horizontal=True, key="mau_hd_loai")
        loai_hd_ma = "HDLD" if loai_hd_chon.startswith("HĐLĐ") else "HDTV"
        mac_dinh = DEFAULT_DIEU_HDLD if loai_hd_ma == "HDLD" else DEFAULT_DIEU_HDTV

        tuy_chinh_hien_tai = get_all_dieu_hop_dong(loai_hd_ma)
        ds_ma_dieu_hien_thi = get_ds_ma_dieu(tuy_chinh_hien_tai)
        MAC_DINH_KEYS = ("dieu1", "dieu2", "dieu3", "dieu4", "dieu5")

        for ma_dieu in ds_ma_dieu_hien_thi:
            la_mac_dinh = ma_dieu in MAC_DINH_KEYS
            mac_dinh_tieu_de, mac_dinh_noi_dung = mac_dinh.get(ma_dieu, ("", ""))
            if la_mac_dinh and not mac_dinh_tieu_de and ma_dieu not in mac_dinh:
                continue  # loại HĐ này không có điều này (VD HĐTV không có dieu1 mặc định cũ)
            info_hien_tai = tuy_chinh_hien_tai.get(ma_dieu)
            hien_tai_tieu_de = info_hien_tai[0] if info_hien_tai else mac_dinh_tieu_de
            hien_tai_noi_dung = info_hien_tai[1] if info_hien_tai else mac_dinh_noi_dung
            hien_tai_thu_tu = info_hien_tai[2] if (info_hien_tai and len(info_hien_tai) > 2 and info_hien_tai[2]) else 0
            with st.expander(f"📝 {hien_tai_tieu_de or ma_dieu}" + ("" if la_mac_dinh else "  🆕"), expanded=False):
                da_tuy_chinh = info_hien_tai is not None
                if da_tuy_chinh:
                    tieu_de_moi = st.text_input("Tiêu đề Điều:", value=hien_tai_tieu_de, key=f"mau_hd_td_{loai_hd_ma}_{ma_dieu}")
                    noi_dung_moi = st.text_area("Nội dung:", value=hien_tai_noi_dung, height=220, key=f"mau_hd_nd_{loai_hd_ma}_{ma_dieu}")
                else:
                    st.caption("💡 Nội dung hiển thị mờ bên dưới chỉ là **gợi ý mặc định**, chưa phải nội dung đã lưu. "
                               "Nhập nội dung để tạo bản tuỳ chỉnh riêng, hoặc để trống & không bấm Lưu để tiếp tục dùng mặc định.")
                    tieu_de_moi = st.text_input("Tiêu đề Điều:", value="", placeholder=hien_tai_tieu_de, key=f"mau_hd_td_{loai_hd_ma}_{ma_dieu}")
                    noi_dung_moi = st.text_area("Nội dung:", value="", placeholder=hien_tai_noi_dung, height=220, key=f"mau_hd_nd_{loai_hd_ma}_{ma_dieu}")
                if not la_mac_dinh:
                    thu_tu_moi = st.number_input("Vị trí hiển thị (số nhỏ hơn đứng trước, chèn xen giữa các Điều mặc định):",
                                                  min_value=1, max_value=999, value=int(hien_tai_thu_tu) or 999,
                                                  key=f"mau_hd_tt_{loai_hd_ma}_{ma_dieu}")
                else:
                    thu_tu_moi = 0
                col_luu, col_reset, col_xoa = st.columns(3)
                with col_luu:
                    if st.button("💾 Lưu", key=f"mau_hd_save_{loai_hd_ma}_{ma_dieu}", width='stretch', type="primary"):
                        if not da_tuy_chinh and not tieu_de_moi.strip() and not noi_dung_moi.strip():
                            st.warning("⚠️ Bạn chưa nhập nội dung tuỳ chỉnh nào (nội dung mờ chỉ là gợi ý). "
                                       "Vẫn tiếp tục dùng nội dung mặc định, không có gì để lưu.")
                        else:
                          try:
                            db = st.session_state.db_engine.get_connection()
                            c = db.cursor()
                            c.execute("""
                                INSERT INTO mau_dieu_hop_dong (loai_hd, ma_dieu, tieu_de, noi_dung, thu_tu, updated_at)
                                VALUES (%s, %s, %s, %s, %s, NOW())
                                ON CONFLICT (loai_hd, ma_dieu) DO UPDATE
                                SET tieu_de = EXCLUDED.tieu_de, noi_dung = EXCLUDED.noi_dung,
                                    thu_tu = EXCLUDED.thu_tu, updated_at = NOW()
                            """, (loai_hd_ma, ma_dieu, tieu_de_moi, noi_dung_moi, thu_tu_moi))
                            db.commit(); db.close()
                            get_all_dieu_hop_dong.clear()
                            st.success(f"✅ Đã lưu {ma_dieu}")
                            st.rerun()
                          except Exception as e:
                            st.error(f"❌ Lỗi: {e}")
                with col_reset:
                    if la_mac_dinh:
                        if st.button("↩️ Khôi phục mặc định", key=f"mau_hd_reset_{loai_hd_ma}_{ma_dieu}", width='stretch'):
                            try:
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                c.execute("DELETE FROM mau_dieu_hop_dong WHERE loai_hd=%s AND ma_dieu=%s", (loai_hd_ma, ma_dieu))
                                db.commit(); db.close()
                                get_all_dieu_hop_dong.clear()
                                st.success("✅ Đã khôi phục mặc định")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")
                with col_xoa:
                    if not la_mac_dinh:
                        if st.button("🗑️ Xoá Điều này", key=f"mau_hd_xoa_{loai_hd_ma}_{ma_dieu}", width='stretch'):
                            try:
                                db = st.session_state.db_engine.get_connection()
                                c = db.cursor()
                                c.execute("DELETE FROM mau_dieu_hop_dong WHERE loai_hd=%s AND ma_dieu=%s", (loai_hd_ma, ma_dieu))
                                db.commit(); db.close()
                                get_all_dieu_hop_dong.clear()
                                st.success("✅ Đã xoá Điều")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi: {e}")

        st.divider()
        with st.expander("➕ Thêm Điều mới", expanded=False):
            tuy_chinh_hdld_all = get_all_dieu_hop_dong("HDLD")
            tuy_chinh_hdtv_all = get_all_dieu_hop_dong("HDTV")
            ma_dieu_de_xuat = sinh_ma_dieu_moi(tuy_chinh_hdld_all, tuy_chinh_hdtv_all)
            st.caption(f"Mã Điều mới sẽ được tạo tự động: **{ma_dieu_de_xuat}**")
            tieu_de_them = st.text_input("Tiêu đề Điều mới:", placeholder="VD: Điều 6. Bảo mật thông tin:", key=f"mau_hd_them_td_{loai_hd_ma}")
            noi_dung_them = st.text_area("Nội dung:", height=180, key=f"mau_hd_them_nd_{loai_hd_ma}",
                                          placeholder="-    Nội dung dòng 1;\n-    Nội dung dòng 2;\n## Tiêu đề phụ in đậm\n-    Nội dung...")
            thu_tu_them = st.number_input("Vị trí hiển thị (số nhỏ hơn đứng trước, VD: 6 = ngay sau Điều 5):",
                                           min_value=1, max_value=999, value=6, key=f"mau_hd_them_tt_{loai_hd_ma}")
            if st.button("➕ Thêm Điều này", key=f"mau_hd_them_btn_{loai_hd_ma}", type="primary"):
                if not tieu_de_them.strip():
                    st.error("⚠️ Vui lòng nhập tiêu đề Điều!")
                else:
                    try:
                        db = st.session_state.db_engine.get_connection()
                        c = db.cursor()
                        c.execute("""
                            INSERT INTO mau_dieu_hop_dong (loai_hd, ma_dieu, tieu_de, noi_dung, thu_tu, updated_at)
                            VALUES (%s, %s, %s, %s, %s, NOW())
                        """, (loai_hd_ma, ma_dieu_de_xuat, tieu_de_them, noi_dung_them, thu_tu_them))
                        db.commit(); db.close()
                        get_all_dieu_hop_dong.clear()
                        st.success(f"✅ Đã thêm {ma_dieu_de_xuat} vào {loai_hd_chon}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Lỗi: {e}")

    with tab_cv:
        st.subheader("🎖️ Danh mục Chức vụ")
        st.caption("Chức vụ (Trưởng phòng, Phó phòng, Tổ trưởng...) có thể gọi khác nhau tuỳ doanh nghiệp — quản lý danh mục riêng tại đây.")
        try:
            db_cv0 = st.session_state.db_engine.get_connection()
            c_cv0 = db_cv0.cursor()
            c_cv0.execute("""
                CREATE TABLE IF NOT EXISTS chuc_vu_danh_muc (
                    id SERIAL PRIMARY KEY,
                    ten_chuc_vu VARCHAR(150) UNIQUE NOT NULL,
                    thu_tu INT DEFAULT 0,
                    trang_thai VARCHAR(20) DEFAULT 'Hoạt động'
                )
            """)
            db_cv0.commit(); db_cv0.close()
        except Exception as e:
            st.error(f"❌ Lỗi khởi tạo danh mục Chức vụ: {e}")
        _quan_ly_danh_muc_don_gian('chuc_vu_danh_muc', 'ten_chuc_vu', 'Chức vụ', 'VD: Trưởng phòng, Tổ trưởng, Phó Tổng Giám đốc...')

    with tab_cty:
        st.subheader("⚙️ Cấu hình chung của Doanh nghiệp")
        st.caption("Toàn bộ thiết lập áp dụng cho doanh nghiệp bạn — tập trung quản lý tại đây thay vì rải rác nhiều trang.")

        tenant_info = st.session_state.get('tenant', {}) or {}
        st.text_input(
            "🏷️ Mã công ty (ma_cty)", value=tenant_info.get('ma_cty', 'CHL'), disabled=True,
            help="Mã công ty được cấp khi khởi tạo tài khoản. Không tự đổi tại đây vì sẽ làm sai lệch số văn bản/hợp đồng đã phát hành theo mã cũ — liên hệ nhà cung cấp app nếu thực sự cần đổi."
        )

        st.divider()
        st.markdown("**📄 Đánh số Công văn đi**")
        cv_option_hien_tai = get_cv_danh_so_option()
        cv_option_moi = st.radio(
            "Phương án đánh số công văn đi:",
            options=['CHUNG', 'RIENG'],
            index=0 if cv_option_hien_tai == 'CHUNG' else 1,
            format_func=lambda x: "📌 Số chung cho tất cả loại công văn" if x == 'CHUNG' else "📌 Mỗi loại công văn có số riêng",
            key="cty_cv_option_radio"
        )
        st.caption("💡 Muốn xem trạng thái số hiện tại theo từng loại hoặc đặt lại số, vào menu "
                   "**Quản lý Công văn & HĐ kinh tế → ⚙️ Cấu hình đánh số công văn**.")

        st.markdown("**📄 Đánh số Hợp đồng kinh tế (HĐKT)**")
        st.caption("Mẫu số: **stt/năm/Prefix-ma_cty** (VD: 04/2026/HĐKT-CHL)")
        hdkt_prefix_hien_tai = get_hdkt_prefix()
        hdkt_prefix_moi = st.text_input("Prefix đánh số HĐKT:", value=hdkt_prefix_hien_tai, key="cty_hdkt_prefix_input")

        st.divider()
        st.markdown("**📋 Hạn nộp Báo cáo Tăng/Giảm BHXH hàng tháng**")
        st.caption("Mỗi doanh nghiệp có 1 ngày chốt hạn riêng trong tháng (VD: CHL nộp trước ngày 20). "
                   "Dashboard sẽ tự cảnh báo tăng dần mức độ trong 5 ngày trước hạn.")
        han_bhxh_hien_tai = get_han_nop_bhxh()
        han_bhxh_moi = st.number_input(
            "Ngày trong tháng phải nộp (1-28):", min_value=1, max_value=28,
            value=han_bhxh_hien_tai, step=1, key="cty_han_bhxh_input"
        )

        st.divider()
        st.markdown("**🕒 Cấu hình Chấm công**")
        st.caption("⚠️ Khung sườn cấu hình — logic tính công/đi trễ sẽ được tích hợp và hoàn thiện dần theo nhu cầu phát sinh.")
        cc_hien_tai = get_cau_hinh_cham_cong()
        col_cc1, col_cc2, col_cc3 = st.columns(3)
        with col_cc1:
            cc_gio_vao_moi = st.time_input("Giờ vào chuẩn:", value=cc_hien_tai['gio_vao'], key="cty_cc_gio_vao")
        with col_cc2:
            cc_gio_ra_moi = st.time_input("Giờ ra chuẩn:", value=cc_hien_tai['gio_ra'], key="cty_cc_gio_ra")
        with col_cc3:
            cc_phut_tre_moi = st.number_input("Số phút cho phép trễ:", min_value=0, max_value=120,
                                               value=cc_hien_tai['phut_tre'], step=5, key="cty_cc_phut_tre")

        st.divider()
        st.markdown("**💰 Phần mềm tính lương**")
        st.caption("⚠️ [SẼ BUILD HOÀN THIỆN SAU] Mỗi công thức tính lương là 1 file riêng trong thư mục `salary/salary_{key}.py`. "
                   "Chọn công thức áp dụng cho doanh nghiệp tại đây; các công thức mới sẽ được bổ sung theo nhu cầu phát sinh.")
        luong_key_hien_tai = get_cau_hinh_luong_key()
        luong_key_moi = st.selectbox(
            "Công thức tính lương áp dụng:",
            options=["salary_1"],
            index=0,
            format_func=lambda x: f"{x} (mặc định — chưa có công thức khác)",
            key="cty_luong_key_select",
            help="Danh sách sẽ tự động mở rộng khi có thêm file salary/salary_{key}.py mới."
        )

        st.divider()
        if st.button("💾 SAVE CẤU HÌNH", type="primary", width='stretch', key="btn_save_cau_hinh_cty"):
            loi_luu = []
            if not update_cv_danh_so_option(cv_option_moi):
                loi_luu.append("Đánh số công văn")
            if not update_hdkt_prefix(hdkt_prefix_moi.strip()):
                loi_luu.append("Prefix HĐKT")
            if not update_han_nop_bhxh(int(han_bhxh_moi)):
                loi_luu.append("Hạn nộp BHXH")
            if not update_cau_hinh_cham_cong(cc_gio_vao_moi, cc_gio_ra_moi, int(cc_phut_tre_moi)):
                loi_luu.append("Cấu hình chấm công")
            if not update_cau_hinh_luong_key(luong_key_moi):
                loi_luu.append("Công thức lương")

            if loi_luu:
                st.error(f"❌ Lưu thất bại một số mục: {', '.join(loi_luu)}")
            else:
                st.success("✅ Đã lưu toàn bộ cấu hình doanh nghiệp!")
                st.cache_data.clear()
                st.rerun()

# ========== BHXH ==========
elif menu == "📋 BHXH":
    st.title("📋 Quản lý BHXH")
    
    t1, t2, t3 = st.tabs(["📊 Tổng quan", "📝 Báo cáo tăng/giảm D02-LT", "💰 Dự toán đóng BHXH"])
    
    with t1:
        st.subheader("📊 Tổng quan tình hình đóng BHXH")
        
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        # Thống kê chung
        c.execute("SELECT COUNT(*) as tong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')")
        tong_ld = c.fetchone()['tong']
        
        c.execute("""
            SELECT COUNT(*) as dang_dong 
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') 
            AND thang_bat_dau_bh IS NOT NULL  -- ĐÃ CÓ NGÀY BẮT ĐẦU = ĐANG THAM GIA
        """)
        dang_dong = c.fetchone()['dang_dong']

        c.execute("""
            SELECT COUNT(*) as chua_dong 
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') 
            AND thang_bat_dau_bh IS NULL  -- CHƯA CÓ NGÀY BẮT ĐẦU = CHƯA THAM GIA
        """)
        chua_dong = c.fetchone()['chua_dong']
        
        c.execute("SELECT COUNT(*) as da_nghi FROM nhan_vien WHERE trang_thai = 'NGHI_VIEC'")
        da_nghi = c.fetchone()['da_nghi']
        
        db.close()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("👥 Tổng lao động", tong_ld)
        col2.metric("✅ Đang đóng BHXH", dang_dong, delta=f"{dang_dong/tong_ld*100:.0f}%" if tong_ld > 0 else None)
        col3.metric("⏳ Chưa đóng BHXH", chua_dong, delta=f"-{chua_dong}" if chua_dong > 0 else None, delta_color="inverse")
        col4.metric("📋 Đã nghỉ việc", da_nghi)
        
        st.divider()
        
        # Danh sách lao động chưa đóng BHXH
        st.subheader("⚠️ Lao động chưa đóng BHXH")
        
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT ma_nv, ho_ten, chuc_danh_nghe, ngay_vao_lam, loai_hop_dong, 
                   thang_bat_dau_bh, trang_thai_bhxh
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') 
            AND thang_bat_dau_bh IS NULL
            ORDER BY ngay_vao_lam ASC
        """)
        chua_dong_list = c.fetchall()

        db.close()
        
        if chua_dong_list:
            df_chua_dong = pd.DataFrame(chua_dong_list)
            for col in df_chua_dong.columns:
                if 'ngay' in col.lower():
                    df_chua_dong[col] = df_chua_dong[col].apply(format_date)
            st.dataframe(df_chua_dong, width='stretch', hide_index=True)
            
            if st.session_state.role == "admin":
                st.warning("💡 Hướng dẫn: Vào menu '✅ Nhân viên' -> chọn nhân viên -> sửa thông tin -> cập nhật 'Bắt đầu BH' và chuyển trạng thái BHXH thành 'ĐANG ĐÓNG'")
        else:
            st.success("✅ Tất cả lao động đã được đăng ký đóng BHXH!")
    
    with t2:
        st.subheader("📝 Báo cáo tăng/giảm lao động tham gia BHXH (Mẫu D02-LT)")
        st.caption("Theo Quyết định 595/QĐ-BHXH và mẫu D02-LT - Dùng để kê khai tăng/giảm lao động tham gia BHXH, BHYT, BHTN")
        
        col_from, col_to = st.columns(2)
        with col_from:
            tu_ngay = st.date_input("📅 Từ ngày (theo tháng bắt đầu/kết thúc BHXH):", 
                                    value=date(date.today().year, 1, 1), 
                                    key="d02_tu")
        with col_to:
            den_ngay = st.date_input("📅 Đến ngày:", 
                                    value=date.today(), 
                                    key="d02_den")
        
        # Nút xuất báo cáo - ĐẶT NGAY PHÍA DƯỚI BỘ LỌC NGÀY
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            export_clicked = st.button("📥 XUẤT EXCEL D02-LT (Mẫu BHXH)", 
                                       type="primary", 
                                       width='stretch',
                                       use_container_width=True)
        
        st.divider()
        
        # Khởi tạo biến để lưu kết quả truy vấn
        tang_list = []
        giam_list = []
        
        # Chỉ truy vấn khi cần (khi người dùng click nút hoặc muốn xem trước)
        # Nhưng để hiển thị preview, chúng ta vẫn chạy truy vấn
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        # Lao động tăng trong kỳ (dựa vào thang_bat_dau_bh)
        c.execute("""
            SELECT 
                nv.id, nv.ma_nv, nv.ho_ten, nv.ma_so_bhxh, nv.ngay_sinh, nv.gioi_tinh, nv.so_cccd,
                nv.chuc_danh_nghe, nv.phong_ban_lam_viec, nv.luong_bao_hiem, nv.he_so_luong,
                nv.thang_bat_dau_bh as ngay_bat_dau,
                nv.loai_hop_dong, nv.so_hdld, nv.ngay_vao_lam, nv.thuong_tru,
                nv.phu_cap_chuc_vu, nv.phu_cap_tnvk, nv.phu_cap_tnn,
                nv.muc_huong_bhyt, nv.ty_le_dong, nv.muc_tien_dong, nv.phuong_thuc_dong,
                nv.quoc_tich, nv.dan_toc, nv.dien_thoai, nv.email_lien_he,
                nv.tinh_nhan_hs, nv.phuong_nhan_hs, nv.dia_chi_nhan_hs,
                nv.tinh_kcb, nv.noi_dang_ky_kcb, nv.dang_ky_nhan_so,
                nv.ngay_ky_hd, nv.ngay_ket_thuc, nv.ten_don_vi_thu_huong
            FROM nhan_vien nv
            WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND nv.thang_bat_dau_bh IS NOT NULL
            AND nv.thang_bat_dau_bh BETWEEN %s AND %s
            ORDER BY nv.thang_bat_dau_bh ASC
        """, (tu_ngay, den_ngay))
        tang_list = c.fetchall()
        
        # Lao động giảm trong kỳ (dựa vào thang_ket_thuc_bh)
        c.execute("""
            SELECT 
                nv.id, nv.ma_nv, nv.ho_ten, nv.ma_so_bhxh, nv.ngay_sinh, nv.gioi_tinh, nv.so_cccd,
                nv.chuc_danh_nghe, nv.phong_ban_lam_viec, nv.luong_bao_hiem, nv.he_so_luong,
                nv.thang_ket_thuc_bh as ngay_ket_thuc,
                nv.loai_hop_dong, nv.so_hdld, nv.ngay_vao_lam, nv.thuong_tru,
                nv.ly_do_nghi
            FROM nhan_vien nv
            WHERE nv.trang_thai = 'NGHI_VIEC'
            AND nv.thang_ket_thuc_bh BETWEEN %s AND %s
            ORDER BY nv.thang_ket_thuc_bh ASC
        """, (tu_ngay, den_ngay))
        giam_list = c.fetchall()
        db.close()
        
        # Hiển thị preview
        col_tang, col_giam = st.columns(2)
        with col_tang:
            st.markdown(f"### 🟢 LAO ĐỘNG TĂNG ({len(tang_list)})")
            if tang_list:
                df_tang = pd.DataFrame(tang_list)
                for col in df_tang.columns:
                    if 'ngay' in col.lower():
                        df_tang[col] = df_tang[col].apply(format_date)
                preview_cols = ['ma_nv', 'ho_ten', 'ma_so_bhxh', 'ngay_bat_dau']
                available_cols = [c for c in preview_cols if c in df_tang.columns]
                df_preview = df_tang[available_cols]
                df_preview.columns = ['Mã NV', 'Họ tên', 'Mã BHXH', 'Ngày bắt đầu']
                st.dataframe(df_preview, width='stretch', hide_index=True, height=300)
            else:
                st.info("📭 Không có lao động tăng trong kỳ")
        
        with col_giam:
            st.markdown(f"### 🔴 LAO ĐỘNG GIẢM ({len(giam_list)})")
            if giam_list:
                df_giam = pd.DataFrame(giam_list)
                for col in df_giam.columns:
                    if 'ngay' in col.lower():
                        df_giam[col] = df_giam[col].apply(format_date)
                preview_cols = ['ma_nv', 'ho_ten', 'ma_so_bhxh', 'ngay_ket_thuc']
                available_cols = [c for c in preview_cols if c in df_giam.columns]
                df_preview = df_giam[available_cols]
                df_preview.columns = ['Mã NV', 'Họ tên', 'Mã BHXH', 'Ngày kết thúc']
                st.dataframe(df_preview, width='stretch', hide_index=True, height=300)
            else:
                st.info("📭 Không có lao động giảm trong kỳ")
        
        st.divider()
        
        # ===== XỬ LÝ XUẤT EXCEL KHI NHẤN NÚT =====
        if export_clicked:
            if tang_list or giam_list:
                with st.spinner("Đang tạo báo cáo D02-LT theo mẫu BHXH... Vui lòng chờ..."):
                    try:
                        # Gọi hàm tạo báo cáo
                        filename = tao_bao_cao_bhxh_d02_lt(
                            tang_list, 
                            giam_list, 
                            tu_ngay, 
                            den_ngay, 
                            COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA"),
                            COMPANY_CONFIG.get("ma_don_vi_BHXH", "4400000000")
                        )
                        
                        # Đọc file và tải xuống
                        with open(filename, "rb") as f:
                            file_data = f.read()
                        
                        st.success(f"✅ Đã tạo báo cáo thành công! {len(tang_list)} lao động tăng, {len(giam_list)} lao động giảm.")
                        
                        st.download_button(
                            label="📥 TẢI FILE EXCEL D02-LT (Đúng mẫu BHXH)",
                            data=file_data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch',
                            key="download_d02_lt"
                        )
                        
                        # Xóa file tạm sau khi đã đọc
                        import os
                        if os.path.exists(filename):
                            os.remove(filename)
                            
                    except Exception as e:
                        st.error(f"❌ Lỗi khi tạo báo cáo: {str(e)}")
                        st.exception(e)
            else:
                st.warning("⚠️ Không có biến động lao động (tăng hoặc giảm) trong kỳ để xuất báo cáo!")
    
    with t3:
        st.subheader("💰 DỰ TOÁN ĐÓNG BHXH")
        st.caption("Tính toán các khoản phải nộp Bảo hiểm xã hội, Bảo hiểm y tế, Bảo hiểm thất nghiệp")
        
        st.info("""
        ### 🚧 Tính năng đang hoàn thiện
        
        Nội dung đang được phát triển. Các tính năng sắp ra mắt:
        - ✅ Tính toán mức đóng BHXH theo lương cơ sở
        - ✅ Tính toán các khoản phụ cấp tính đóng BHXH
        - ✅ Bảng kê chi tiết từng nhân viên
        - ✅ Xuất báo cáo kê khai BHXH theo mẫu quy định
        - ✅ Tổng hợp số tiền phải nộp theo tháng/quý/năm
        
        ⏳ **Dự kiến hoàn thành: Quý 3/2026**
        """)
        
        # Thêm một số thông tin tham khảo
        with st.expander("📌 Thông tin tham khảo về tỷ lệ đóng BHXH hiện hành"):
            st.markdown("""
            **Tỷ lệ trích BHXH, BHYT, BHTN theo quy định hiện hành:**
            
            | Loại | Doanh nghiệp | Người lao động | Tổng |
            |------|-------------|----------------|------|
            | BHXH | 17.5% | 8% | 25.5% |
            | BHYT | 3% | 1.5% | 4.5% |
            | BHTN | 1% | 1% | 2% |
            | BHTNLĐ-BNN | 0.5% | 0% | 0.5% |
            | **Tổng cộng** | **22%** | **10.5%** | **32.5%** |
            
            *Lưu ý: Tỷ lệ có thể thay đổi theo quy định mới của Nhà nước.*
            """)
        
        # Thêm form nhập thử nghiệm (có thể comment lại sau)
        with st.expander("🧪 Thử nghiệm tính toán (Demo)"):
            col_demo1, col_demo2 = st.columns(2)
            with col_demo1:
                luong_demo = st.number_input("Lương tháng (VNĐ)", min_value=0, value=5000000, step=500000)
            with col_demo2:
                chon_ty_le = st.selectbox("Áp dụng tỷ lệ", ["Theo quy định", "Tùy chỉnh"])
            
            if chon_ty_le == "Theo quy định":
                ty_le_nld = 10.5
                ty_le_nsdl = 22.0
            else:
                col_ty1, col_ty2 = st.columns(2)
                with col_ty1:
                    ty_le_nld = st.number_input("Tỷ lệ NLĐ (%)", min_value=0.0, max_value=50.0, value=10.5, step=0.5)
                with col_ty2:
                    ty_le_nsdl = st.number_input("Tỷ lệ NSDLĐ (%)", min_value=0.0, max_value=50.0, value=22.0, step=0.5)
            
            tien_nld = luong_demo * ty_le_nld / 100
            tien_nsdl = luong_demo * ty_le_nsdl / 100
            tong_tien = tien_nld + tien_nsdl
            
            st.markdown("---")
            col_kq1, col_kq2, col_kq3 = st.columns(3)
            col_kq1.metric("NLĐ đóng", f"{tien_nld:,.0f} VNĐ", f"({ty_le_nld}%)")
            col_kq2.metric("NSDLĐ đóng", f"{tien_nsdl:,.0f} VNĐ", f"({ty_le_nsdl}%)")
            col_kq3.metric("Tổng tiền", f"{tong_tien:,.0f} VNĐ", "cả 2 bên")
        
# ========== BÁO CÁO TÌNH HÌNH SỬ DỤNG LAO ĐỘNG MẪU 01/PLI (EXCEL) ==========
elif menu == "📋 Báo cáo định kỳ":
    st.title("📋 Báo cáo định kỳ")

    tab_bc_pli, tab_bc_tk, tab_bc_tanggiam, tab_bc_tinhhinh, tab_bc_yte, tab_bc_atvsld, \
    tab_bc_tnld, tab_bc_huanluyen, tab_bc_socapcuu, tab_bc_quantrac = st.tabs([
        "📋 Báo cáo 01/PLI", "📊 Báo cáo thống kê nhân sự", "📊 Báo cáo tăng/giảm nhân sự trong kỳ",
        "📈 Tình hình sử dụng lao động", "🏥 Y tế Lao động", "🦺 Công tác ATVSLĐ",
        "⚠️ Tai nạn lao động", "🎓 Huấn luyện ATVSLĐ", "🚑 Mạng lưới sơ cấp cứu", "🌡️ Quan trắc môi trường LĐ"
    ])

    def _bao_cao_dang_phat_trien(ten_bao_cao, mo_ta, icon, mau_sac="#f59e0b"):
        """Card hiện đại cho các báo cáo định kỳ chưa có logic - sẽ bổ sung sau."""
        st.markdown(f"""
        <div style="border:1px solid #e5e7eb;border-radius:16px;padding:28px 24px;
             background:linear-gradient(135deg,#fffbeb 0%,#ffffff 100%);
             box-shadow:0 2px 10px rgba(0,0,0,0.04);">
            <div style="display:flex;align-items:center;gap:14px;margin-bottom:10px;">
                <div style="font-size:34px;">{icon}</div>
                <div>
                    <div style="font-size:19px;font-weight:700;color:#111827;">{ten_bao_cao}</div>
                    <span style="display:inline-block;margin-top:4px;padding:2px 10px;border-radius:999px;
                          background:{mau_sac};color:white;font-size:11px;font-weight:600;letter-spacing:.3px;">
                          🚧 SẼ BỔ SUNG SAU
                    </span>
                </div>
            </div>
            <div style="color:#4b5563;font-size:14px;line-height:1.6;margin-top:8px;">{mo_ta}</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        with st.expander("ℹ️ Khi triển khai, báo cáo này sẽ cần"):
            st.markdown("""
            - Biểu mẫu/Thông tư pháp lý làm căn cứ (bạn cung cấp khi sẵn sàng)
            - Kỳ báo cáo (tháng/quý/năm) và bộ lọc theo phòng ban
            - Xuất file Word/Excel theo đúng mẫu quy định
            """)

    with tab_bc_tinhhinh:
        st.caption("💡 Lưu ý: tab **📋 Báo cáo 01/PLI** hiện tại cũng đang thể hiện nội dung "
                   "\"tình hình sử dụng lao động\" theo mẫu 01/PLI. Nếu đây là 1 báo cáo khác "
                   "(mẫu/kỳ báo cáo khác), bạn gửi mẫu cụ thể để tôi phân biệt rõ khi triển khai.")
        _bao_cao_dang_phat_trien(
            "Báo cáo tình hình sử dụng lao động",
            "Tổng hợp định kỳ tình hình sử dụng lao động của doanh nghiệp (số lượng, cơ cấu, biến động) "
            "theo quy định báo cáo lao động định kỳ.",
            "📈"
        )

    with tab_bc_yte:
        _bao_cao_dang_phat_trien(
            "Báo cáo Y tế Lao động",
            "Báo cáo công tác y tế lao động: khám sức khỏe định kỳ, bệnh nghề nghiệp, tình hình sức khỏe người lao động.",
            "🏥"
        )

    with tab_bc_atvsld:
        _bao_cao_dang_phat_trien(
            "Báo cáo công tác An toàn, Vệ sinh lao động (ATVSLĐ)",
            "Tổng hợp công tác an toàn vệ sinh lao động: tổ chức bộ máy ATVSLĐ, tự kiểm tra, cải thiện điều kiện làm việc.",
            "🦺"
        )

    with tab_bc_tnld:
        _bao_cao_dang_phat_trien(
            "Báo cáo Tai nạn lao động",
            "Thống kê, khai báo các vụ tai nạn lao động phát sinh trong kỳ báo cáo theo quy định.",
            "⚠️", mau_sac="#ef4444"
        )

    with tab_bc_huanluyen:
        _bao_cao_dang_phat_trien(
            "Báo cáo Huấn luyện ATVSLĐ",
            "Tổng hợp tình hình huấn luyện an toàn vệ sinh lao động theo nhóm đối tượng, thời hạn huấn luyện lại.",
            "🎓"
        )

    with tab_bc_socapcuu:
        _bao_cao_dang_phat_trien(
            "Báo cáo hoạt động mạng lưới sơ cấp cứu",
            "Tình hình tổ chức, hoạt động của mạng lưới sơ cấp cứu tại doanh nghiệp.",
            "🚑"
        )

    with tab_bc_quantrac:
        _bao_cao_dang_phat_trien(
            "Báo cáo Quan trắc môi trường lao động",
            "Kết quả quan trắc môi trường lao động định kỳ (các yếu tố có hại tại nơi làm việc).",
            "🌡️"
        )

    with tab_bc_pli:
        st.subheader("📋 Báo cáo tình hình sử dụng lao động")
        st.caption("Theo mẫu 01/PLI Phụ lục I - Nghị định 145/2020/NĐ-CP (sửa đổi bởi Nghị định 35/2022/NĐ-CP)")
    
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
        col1, col2 = st.columns(2)
        with col1:
            tu_ngay = st.date_input("📅 Từ ngày:", value=date(date.today().year, 1, 1), key="pli_tu")
        with col2:
            den_ngay = st.date_input("📅 Đến ngày:", value=date.today(), key="pli_den")
    
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT 
                nv.STT, nv.ma_nv, nv.ho_ten, nv.ma_so_bhxh, nv.ngay_sinh, nv.gioi_tinh,
                nv.so_cccd, nv.chuc_danh_nghe, nv.luong_bao_hiem, nv.he_so_luong,
                nv.phu_cap_chuc_vu, nv.phu_cap_tnvk, nv.phu_cap_tnn, nv.loai_hop_dong,
                nv.ngay_vao_lam, nv.ngay_ky_hd, nv.ngay_ket_thuc, nv.thang_bat_dau_bh,
                nv.thang_ket_thuc_bh, nv.so_hdld, nv.phong_ban_lam_viec, nv.noi_lam_viec,
                nv.ten_don_vi_thu_huong
            FROM nhan_vien nv
            WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC', 'NGHI_VIEC')
            AND nv.ngay_vao_lam <= %s
            AND (nv.ngay_ket_thuc IS NULL OR nv.ngay_ket_thuc >= %s)
            ORDER BY nv.STT ASC
        """, (den_ngay, tu_ngay))
        ds_lao_dong = c.fetchall()
        db.close()
    
        st.info(f"📊 Tổng số lao động đang làm việc: **{len(ds_lao_dong)}** người")
    
        # Hiển thị bảng dữ liệu trước khi xuất (cho cả admin và viewer)
        if ds_lao_dong:
            st.subheader("📋 Danh sách lao động")
            df_preview = pd.DataFrame(ds_lao_dong)
            for col in df_preview.columns:
                if 'ngay' in col.lower():
                    df_preview[col] = df_preview[col].apply(format_date)
        
            preview_cols = ['ma_nv', 'ho_ten', 'chuc_danh_nghe', 'loai_hop_dong', 'ngay_vao_lam', 'ma_so_bhxh', 'ten_don_vi_thu_huong']
            available_preview = [c for c in preview_cols if c in df_preview.columns]
            df_display = df_preview[available_preview]
            col_map_preview = {
                'ma_nv': 'Mã NV',
                'ho_ten': 'Họ tên',
                'chuc_danh_nghe': 'Chức danh',
                'loai_hop_dong': 'Loại HĐ',
                'ngay_vao_lam': 'Ngày vào làm',
                'ma_so_bhxh': 'Mã BHXH',
                'ten_don_vi_thu_huong': 'Tên đơn vị thụ hưởng',
            }
            df_display.rename(columns=col_map_preview, inplace=True)
            st.dataframe(df_display, width='stretch', hide_index=True, height=400)
        
            st.divider()
        
            # Chỉ admin mới được xuất Excel
            if st.session_state.role == "admin":
                if st.button("📥 XUẤT EXCEL MẪU 01/PLI", type="primary", width='stretch'):
                    if not can_export():
                        st.error("❌ Bạn không có quyền xuất báo cáo!")
                    else:
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "BC_Tinh_hinh_su_dung_LD"
                    
                        ten_cong_ty = COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA")
                        dia_chi = COMPANY_CONFIG.get("dia_chi", "")
                        ma_so_thue = COMPANY_CONFIG.get("ma_so_thue", "")
                        dien_thoai_cty = COMPANY_CONFIG.get("dien_thoai_cty", "")
                    
                        # Header
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
                        ws['A1'] = ten_cong_ty
                        ws['A1'].font = Font(bold=True, size=13, name='Times New Roman')
                        ws['A1'].alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells(start_row=1, start_column=20, end_row=1, end_column=26)
                        ws['T1'] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
                        ws['T1'].font = Font(bold=True, size=13, name='Times New Roman')
                        ws['T1'].alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
                        ws['A2'] = f"Số: 01/BC-01PLI-{datetime.now().year}/CHL"
                        ws['A2'].font = Font(size=12, name='Times New Roman')
                        ws['A2'].alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells(start_row=2, start_column=20, end_row=2, end_column=26)
                        ws['T2'] = "Độc lập - Tự do - Hạnh phúc"
                        ws['T2'].font = Font(italic=True, size=12, name='Times New Roman')
                        ws['T2'].alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells(start_row=3, start_column=20, end_row=3, end_column=26)
                        ws['T3'] = f"Quảng Trị, ngày {date.today().day} tháng {date.today().month} năm {date.today().year}"
                        ws['T3'].font = Font(italic=True, size=12, name='Times New Roman')
                        ws['T3'].alignment = Alignment(horizontal='right')
                    
                        ws.merge_cells('A5:AA5')
                        ws['A5'] = "BÁO CÁO TÌNH HÌNH SỬ DỤNG LAO ĐỘNG"
                        ws['A5'].font = Font(bold=True, size=13, name='Times New Roman')
                        ws['A5'].alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells('A6:AA6')
                        ws['A6'] = f"(Từ ngày {tu_ngay.strftime('%d/%m/%Y')} đến ngày {den_ngay.strftime('%d/%m/%Y')})"
                        ws['A6'].font = Font(size=12, name='Times New Roman')
                        ws['A6'].alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells('A8:AA8')
                        ws['A8'] = "Kính gửi: SỞ NỘI VỤ TỈNH QUẢNG TRỊ"
                        ws['A8'].font = Font(bold=True, size=11, name='Times New Roman')
                        ws['A8'].alignment = Alignment(horizontal='left')
                    
                        ws['A10'] = "1. Thông tin chung về doanh nghiệp:"
                        ws['A10'].font = Font(bold=True, size=11, name='Times New Roman')
                    
                        row_info = 11
                        for label in [f"- Tên doanh nghiệp: {ten_cong_ty}", f"- Địa chỉ: {dia_chi}", 
                                     f"- Mã số thuế: {ma_so_thue}", f"- Điện thoại: {dien_thoai_cty}"]:
                            ws[f'A{row_info}'] = label
                            ws[f'A{row_info}'].font = Font(size=11, name='Times New Roman')
                            row_info += 1
                    
                        ws[f'A{row_info + 1}'] = "2. Thông tin tình hình sử dụng lao động của đơn vị:"
                        ws[f'A{row_info + 1}'].font = Font(bold=True, size=11, name='Times New Roman')
                    
                        header_row = 18
                        col_widths = [5, 25, 18, 15, 8, 18, 25, 12, 18, 18, 12, 15, 
                                     12, 12, 12, 12, 15, 12, 12, 18, 18, 18, 18, 18, 18, 18, 20]
                        for i, w in enumerate(col_widths, 1):
                            ws.column_dimensions[get_column_letter(i)].width = w
                    
                        stt_row = header_row + 3
                        for col in range(1, 28):
                            ws.cell(row=stt_row, column=col, value=f"({col})")
                            ws.cell(row=stt_row, column=col).font = Font(size=9, name='Times New Roman')
                            ws.cell(row=stt_row, column=col).alignment = Alignment(horizontal='center')
                            ws.cell(row=stt_row, column=col).border = thin_border
                    
                        # Merge cells header
                        ws.merge_cells(start_row=header_row, start_column=1, end_row=header_row+2, end_column=1)
                        ws.cell(row=header_row, column=1, value="STT")
                    
                        ws.merge_cells(start_row=header_row, start_column=2, end_row=header_row+2, end_column=2)
                        ws.cell(row=header_row, column=2, value="Họ và tên")
                    
                        ws.merge_cells(start_row=header_row, start_column=3, end_row=header_row+2, end_column=3)
                        ws.cell(row=header_row, column=3, value="Mã số BHXH")
                    
                        ws.merge_cells(start_row=header_row, start_column=4, end_row=header_row+2, end_column=4)
                        ws.cell(row=header_row, column=4, value="Ngày sinh")
                    
                        ws.merge_cells(start_row=header_row, start_column=5, end_row=header_row+2, end_column=5)
                        ws.cell(row=header_row, column=5, value="Giới tính")
                    
                        ws.merge_cells(start_row=header_row, start_column=6, end_row=header_row+2, end_column=6)
                        ws.cell(row=header_row, column=6, value="Số CCCD/Hộ chiếu")
                    
                        ws.merge_cells(start_row=header_row, start_column=7, end_row=header_row+2, end_column=7)
                        ws.cell(row=header_row, column=7, value="Chức danh nghề, vị trí, công việc")
                    
                        ws.merge_cells(start_row=header_row, start_column=8, end_row=header_row, end_column=11)
                        ws.cell(row=header_row, column=8, value="Vị trí việc làm (2)")
                    
                        ws.merge_cells(start_row=header_row, start_column=12, end_row=header_row, end_column=17)
                        ws.cell(row=header_row, column=12, value="Tiền lương")
                    
                        ws.merge_cells(start_row=header_row, start_column=20, end_row=header_row, end_column=24)
                        ws.cell(row=header_row, column=20, value="Loại và hiệu lực hợp đồng")
                    
                        ws.merge_cells(start_row=header_row, start_column=18, end_row=header_row+1, end_column=19)
                        ws.cell(row=header_row, column=18, value="Ngành nghề nặng nhọc, độc hại")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=13, end_row=header_row+1, end_column=17)
                        ws.cell(row=header_row+1, column=13, value="Phụ cấp")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=21, end_row=header_row+1, end_column=22)
                        ws.cell(row=header_row+1, column=21, value="Hiệu lực HĐLĐ xác định thời hạn")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=23, end_row=header_row+1, end_column=24)
                        cell = ws.cell(row=header_row+1, column=23, value="Hiệu lực HĐLĐ khác (dưới 1 tháng, thử việc)")
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    
                        ws.merge_cells(start_row=header_row+1, start_column=8, end_row=header_row+2, end_column=8)
                        ws.cell(row=header_row+1, column=8, value="Nhà quản lý")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=9, end_row=header_row+2, end_column=9)
                        ws.cell(row=header_row+1, column=9, value="Chuyên môn kỹ thuật bậc cao")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=10, end_row=header_row+2, end_column=10)
                        ws.cell(row=header_row+1, column=10, value="Chuyên môn kỹ thuật bậc trung")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=11, end_row=header_row+2, end_column=11)
                        ws.cell(row=header_row+1, column=11, value="Khác")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=12, end_row=header_row+2, end_column=12)
                        ws.cell(row=header_row+1, column=12, value="Mức lương/Hệ số lương")
                    
                        ws.merge_cells(start_row=header_row+1, start_column=20, end_row=header_row+2, end_column=20)
                        ws.cell(row=header_row+1, column=20, value="Ngày bắt đầu HĐLĐ không xác định thời hạn")
                    
                        ws.merge_cells(start_row=header_row, start_column=25, end_row=header_row+2, end_column=25)
                        ws.cell(row=header_row, column=25, value="Thời điểm bắt đầu đóng BHXH")
                    
                        ws.merge_cells(start_row=header_row, start_column=26, end_row=header_row+2, end_column=26)
                        ws.cell(row=header_row, column=26, value="Thời điểm kết thúc đóng BHXH")
                    
                        ws.merge_cells(start_row=header_row, start_column=27, end_row=header_row+2, end_column=27)
                        ws.cell(row=header_row, column=27, value="Ghi chú")
                    
                        for row in range(header_row, header_row + 3):
                            for col in range(1, 28):
                                cell = ws.cell(row=row, column=col)
                                cell.border = thin_border
                    
                        ws.cell(row=header_row+2, column=13, value="Phụ cấp chức vụ")
                        ws.cell(row=header_row+2, column=14, value="Phụ cấp thâm niên VK(%)")
                        ws.cell(row=header_row+2, column=15, value="Phụ cấp thâm niên nghề (%)")
                        ws.cell(row=header_row+2, column=16, value="Phụ cấp thâm niên nghề (%)")
                        ws.cell(row=header_row+2, column=17, value="Các khoản bổ sung")
                        ws.cell(row=header_row+2, column=18, value="Ngày bắt đầu")
                        ws.cell(row=header_row+2, column=19, value="Ngày kết thúc")
                        ws.cell(row=header_row+2, column=21, value="Ngày bắt đầu")
                        ws.cell(row=header_row+2, column=22, value="Ngày kết thúc")
                        ws.cell(row=header_row+2, column=23, value="Ngày bắt đầu")
                        ws.cell(row=header_row+2, column=24, value="Ngày kết thúc")
                    
                        for row in range(header_row, header_row + 3):
                            for col in range(1, 28):
                                cell = ws.cell(row=row, column=col)
                                if cell.value:
                                    cell.font = Font(bold=True, size=10, name='Times New Roman')
                                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                    cell.border = thin_border
                    
                        for col in range(1, 28):
                            cell = ws.cell(row=stt_row, column=col)
                            cell.font = Font(size=9, name='Times New Roman')
                            cell.alignment = Alignment(horizontal='center')
                            cell.border = thin_border
                    
                        data_row = stt_row + 1
                        for idx, nv in enumerate(ds_lao_dong, 1):
                            row = data_row + idx - 1
                        
                            ws.cell(row=row, column=1, value=idx)
                            ws.cell(row=row, column=2, value=nv.get('ho_ten', ''))
                            ws.cell(row=row, column=3, value=nv.get('ma_so_bhxh', ''))
                            ws.cell(row=row, column=4, value=format_date(nv.get('ngay_sinh')))
                            gt = nv.get('gioi_tinh', '')
                            ws.cell(row=row, column=5, value='Nam' if gt == 'Nam' else 'Nữ' if gt == 'Nữ' else '')
                            ws.cell(row=row, column=6, value=nv.get('so_cccd', ''))
                            ws.cell(row=row, column=7, value=nv.get('chuc_danh_nghe', ''))
                        
                            cd = (nv.get('chuc_danh_nghe') or '').lower()
                            is_quan_ly = any(x in cd for x in ['giám đốc', 'trưởng phòng', 'phó', 'quản lý'])
                            ws.cell(row=row, column=8, value='x' if is_quan_ly else '')
                            is_chuyen_mon_cao = (any(x in cd for x in ['kỹ thuật', 'kĩ thuật']) and any(x in cd for x in ['cao', 'chính', 'kỹ sư']))
                            ws.cell(row=row, column=9, value='x' if is_chuyen_mon_cao else '')
                            is_khac = any(x in cd for x in ['phổ thông', 'lao động', 'tạp vụ', 'bảo vệ', 'tạp vụ', 'lái xe'])
                            ws.cell(row=row, column=11, value='x' if is_khac else '')
                            is_trung = (not is_quan_ly) and (not is_chuyen_mon_cao) and (not is_khac)
                            ws.cell(row=row, column=10, value='x' if is_trung else '')
                        
                            luong = nv.get('luong_bao_hiem', '')
                            heso = nv.get('he_so_luong', '')
                            ws.cell(row=row, column=12, value=f"Hệ số: {heso}" if heso and str(heso).strip() else str(luong) if luong else '')
                            ws.cell(row=row, column=13, value=str(nv.get('phu_cap_chuc_vu', '')) if nv.get('phu_cap_chuc_vu') else '')
                            ws.cell(row=row, column=14, value=f"{nv.get('phu_cap_tnvk', '')}%" if nv.get('phu_cap_tnvk') else '')
                            ws.cell(row=row, column=15, value=f"{nv.get('phu_cap_tnn', '')}%" if nv.get('phu_cap_tnn') else '')
                            ws.cell(row=row, column=16, value='')
                            ws.cell(row=row, column=17, value='')
                            ws.cell(row=row, column=18, value='')
                            ws.cell(row=row, column=19, value='')
                        
                            loai_hd = nv.get('loai_hop_dong', '')
                            ngay_bd = nv.get('ngay_ky_hd') or nv.get('ngay_vao_lam')
                            ngay_kt = nv.get('ngay_ket_thuc')
                            ws.cell(row=row, column=20, value=format_date(ngay_bd) if loai_hd == 'Không xác định thời hạn' else '')
                        
                            if loai_hd == 'Xác định thời hạn':
                                ws.cell(row=row, column=21, value=format_date(ngay_bd))
                                ws.cell(row=row, column=22, value=format_date(ngay_kt) if ngay_kt else '')
                            else:
                                ws.cell(row=row, column=21, value='')
                                ws.cell(row=row, column=22, value='')
                        
                            if loai_hd == 'Thử việc':
                                ws.cell(row=row, column=23, value=format_date(ngay_bd))
                                ws.cell(row=row, column=24, value=format_date(ngay_kt) if ngay_kt else '')
                            else:
                                ws.cell(row=row, column=23, value='')
                                ws.cell(row=row, column=24, value='')
                        
                            ws.cell(row=row, column=25, value=format_date(nv.get('thang_bat_dau_bh')))
                            ws.cell(row=row, column=26, value=format_date(nv.get('thang_ket_thuc_bh')))
                            ws.cell(row=row, column=27, value=nv.get('so_hdld', ''))
                        
                            for col in range(1, 28):
                                cell = ws.cell(row=row, column=col)
                                cell.border = thin_border
                                cell.font = Font(size=10, name='Times New Roman')
                                if col in [1, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]:
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                                else:
                                    cell.alignment = Alignment(horizontal='left', vertical='center')
                    
                        total_row = data_row + len(ds_lao_dong)
                        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
                        ws.cell(row=total_row, column=1, value=f"Tổng cộng: {len(ds_lao_dong)} người")
                        ws.cell(row=total_row, column=1).font = Font(bold=True, size=10, name='Times New Roman')
                        ws.cell(row=total_row, column=1).border = thin_border
                    
                        sign_row = total_row + 3
                        ws.merge_cells(start_row=sign_row, start_column=23, end_row=sign_row, end_column=27)
                        ws.cell(row=sign_row, column=23, value="ĐẠI DIỆN DOANH NGHIỆP")
                        ws.cell(row=sign_row, column=23).font = Font(bold=True, size=11, name='Times New Roman')
                        ws.cell(row=sign_row, column=23).alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells(start_row=sign_row+1, start_column=23, end_row=sign_row+1, end_column=27)
                        ws.cell(row=sign_row+1, column=23, value="(Ký, đóng dấu, ghi rõ họ tên)")
                        ws.cell(row=sign_row+1, column=23).font = Font(size=10, name='Times New Roman')
                        ws.cell(row=sign_row+1, column=23).alignment = Alignment(horizontal='center')
                    
                        ws.merge_cells(start_row=sign_row+2, start_column=23, end_row=sign_row+2, end_column=27)
                        ws.cell(row=sign_row+2, column=23, value=COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC').upper())
                        ws.cell(row=sign_row+2, column=23).font = Font(bold=True, size=11, name='Times New Roman')
                        ws.cell(row=sign_row+2, column=23).alignment = Alignment(horizontal='center')
                    
                        filename = f"Bao_cao_01_PLI_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
                        wb.save(filename)
                    
                        with open(filename, "rb") as f:
                            st.download_button(
                                label="📥 TẢI FILE EXCEL",
                                data=f,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                width='stretch'
                            )
                        st.success(f"✅ Đã xuất báo cáo với {len(ds_lao_dong)} lao động")
                        pass
                    
            else:
                st.info("🔒 Chỉ Admin mới có quyền xuất file Excel báo cáo 01/PLI. Bạn đang ở chế độ xem (Viewer).")
                st.caption("💡 Với quyền Viewer, bạn có thể xem danh sách lao động ở trên nhưng không thể tải file Excel.")
        else:
            st.warning("⚠️ Không có lao động nào đang làm việc trong kỳ báo cáo!")
        
    # ========== QUẢN LÝ CÔNG VĂN & HĐ KINH TẾ ==========

    with tab_bc_tk:
        st.caption("⚙️ Tùy chọn bộ lọc và xuất báo cáo thống kê nhân sự")

        # ── Tùy chọn bộ lọc ──
        with st.expander("⚙️ Tùy chọn xuất báo cáo thống kê nhân sự", expanded=True):
            col_tk_date1, col_tk_date2, col_tk_date3 = st.columns([1, 1, 1])
            with col_tk_date1:
                tk_tu_ngay = st.date_input("📅 Từ ngày:", value=date(date.today().year, 1, 1), key="tk_tu_ngay")
            with col_tk_date2:
                tk_den_ngay = st.date_input("📅 Đến ngày:", value=date.today(), key="tk_den_ngay")
            with col_tk_date3:
                loai_hd_filter = st.selectbox(
                    "Loại hợp đồng:",
                    ["Tất cả", "Không xác định thời hạn", "Thử việc"],
                    key="tk_loai_hd"
                )
            col_tk1, col_tk2 = st.columns([1, 2])
            with col_tk1:
                pass  # placeholder

            # Danh sách tất cả cột bảng nhan_vien (trừ id)
            ALL_COLUMNS_LABELS = {
                "ma_nv":              "Mã NV",
                "ho_ten":             "Họ tên",
                "ngay_sinh":          "Ngày sinh",
                "gioi_tinh":          "Giới tính",
                "chuc_danh_nghe":     "Chức danh",
                "phong_ban_lam_viec": "Phòng ban",
                "loai_hop_dong":      "Loại HĐ",
                "ngay_vao_lam":       "Ngày vào làm",
                "ngay_ky_hd":         "Ngày ký HĐ",
                "so_hdld":            "Số HĐLĐ",
                "so_cccd":            "Số CCCD",
                "thuong_tru":         "Thường trú",
                "dien_thoai":         "Điện thoại",
                "ma_so_bhxh":         "Mã BHXH",
                "thang_bat_dau_bh":   "BĐ đóng BH",
                "so_tai_khoan_nh":    "STK",
                "chi_nhanh_nh":       "Chi nhánh NH",
                "ho_so":              "Hồ sơ",
                "ten_don_vi_thu_huong": "Tên đơn vị thụ hưởng",
            }

            # Thứ tự ưu tiên mặc định (tất cả tích mặc định)
            DEFAULT_PRIORITY = [
                "ma_nv", "ho_ten", "ngay_sinh", "gioi_tinh",
                "chuc_danh_nghe", "phong_ban_lam_viec", "loai_hop_dong",
                "ngay_vao_lam", "ngay_ky_hd", "so_hdld",
                "so_cccd", "thuong_tru", "dien_thoai", "ma_so_bhxh",
                "thang_bat_dau_bh", "so_tai_khoan_nh", "chi_nhanh_nh", "ho_so",
                "ten_don_vi_thu_huong",
            ]
            DEFAULT_CHECKED = set(DEFAULT_PRIORITY)

            with col_tk2:
                st.caption("📋 Chọn các cột cần xuất:")
                col_chk = st.columns(4)
                selected_cols = []
                for idx, (col_key, col_label) in enumerate(ALL_COLUMNS_LABELS.items()):
                    default_val = col_key in DEFAULT_CHECKED
                    checked = col_chk[idx % 4].checkbox(col_label, value=default_val, key=f"tk_col_{col_key}")
                    if checked:
                        selected_cols.append(col_key)

            if st.button("📊 XUẤT THỐNG KÊ NHÂN SỰ (EXCEL)", type="primary", width='stretch', key="btn_tk_nhansu"):
                from openpyxl import Workbook
                from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
                from openpyxl.utils import get_column_letter

                if not selected_cols:
                    st.error("⚠️ Vui lòng chọn ít nhất 1 cột!")
                else:
                    db_tk = st.session_state.db_engine.get_connection()
                    c_tk = db_tk.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

                    # Sắp xếp selected_cols theo thứ tự ưu tiên
                    priority_order = DEFAULT_PRIORITY + [k for k in ALL_COLUMNS_LABELS if k not in DEFAULT_PRIORITY]
                    selected_cols_sorted = sorted(selected_cols, key=lambda x: priority_order.index(x) if x in priority_order else 999)

                    sql_cols = ", ".join(selected_cols_sorted)
                    sql_tk = f"SELECT {sql_cols} FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND ngay_vao_lam BETWEEN %s AND %s"
                    params_tk = [tk_tu_ngay, tk_den_ngay]
                    if loai_hd_filter == "Không xác định thời hạn":
                        sql_tk += " AND loai_hop_dong = %s"
                        params_tk.append("Không xác định thời hạn")
                    elif loai_hd_filter == "Thử việc":
                        sql_tk += " AND trang_thai = 'THU_VIEC'"
                    sql_tk += " ORDER BY STT ASC"
                    c_tk.execute(sql_tk, tuple(params_tk))
                    ds_tk = c_tk.fetchall()
                    db_tk.close()

                    if not ds_tk:
                        st.warning("⚠️ Không có nhân viên nào phù hợp với bộ lọc!")
                    else:
                        thin_border_tk = Border(
                            left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin')
                        )
                        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
                        stat_fill  = PatternFill(start_color="D9E8F5", end_color="D9E8F5", fill_type="solid")

                        wb_tk = Workbook()
                        ws_tk = wb_tk.active
                        ws_tk.title = "Thống kê nhân sự"

                        ten_cong_ty_tk = COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA")
                        dia_chi_tk     = COMPANY_CONFIG.get("dia_chi", "")
                        dien_thoai_tk  = COMPANY_CONFIG.get("dien_thoai_cty", "")
                        n_cols = len(selected_cols_sorted) + 1  # +1 cho cột STT

                        # ── Thông tin công ty ──
                        ws_tk.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
                        ws_tk['A1'] = ten_cong_ty_tk
                        ws_tk['A1'].font = Font(bold=True, size=13, name='Times New Roman')
                        ws_tk['A1'].alignment = Alignment(horizontal='center')

                        ws_tk.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
                        ws_tk['A2'] = f"Địa chỉ: {dia_chi_tk}  |  ĐT: {dien_thoai_tk}"
                        ws_tk['A2'].font = Font(size=10, name='Times New Roman', italic=True)
                        ws_tk['A2'].alignment = Alignment(horizontal='center')

                        # ── Tiêu đề báo cáo ──
                        loai_hd_label = f" - Loại HĐ: {loai_hd_filter}" if loai_hd_filter != "Tất cả" else ""
                        ws_tk.merge_cells(start_row=4, start_column=1, end_row=4, end_column=n_cols)
                        ws_tk['A4'] = "BÁO CÁO THỐNG KÊ NHÂN SỰ" + loai_hd_label
                        ws_tk['A4'].font = Font(bold=True, size=14, name='Times New Roman')
                        ws_tk['A4'].alignment = Alignment(horizontal='center')

                        ws_tk.merge_cells(start_row=5, start_column=1, end_row=5, end_column=n_cols)
                        ws_tk['A5'] = f"Từ ngày {tk_tu_ngay.strftime('%d/%m/%Y')} đến ngày {tk_den_ngay.strftime('%d/%m/%Y')}  |  Xuất lúc: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
                        ws_tk['A5'].font = Font(size=10, name='Times New Roman', italic=True)
                        ws_tk['A5'].alignment = Alignment(horizontal='center')

                        # ── Header bảng ──
                        header_row_tk = 7
                        ws_tk.cell(row=header_row_tk, column=1, value="STT").font = Font(bold=True, size=10, name='Times New Roman', color="FFFFFF")
                        ws_tk.cell(row=header_row_tk, column=1).fill = header_fill
                        ws_tk.cell(row=header_row_tk, column=1).alignment = Alignment(horizontal='center', vertical='center')
                        ws_tk.cell(row=header_row_tk, column=1).border = thin_border_tk

                        for col_idx, col_key in enumerate(selected_cols_sorted, 2):
                            cell = ws_tk.cell(row=header_row_tk, column=col_idx, value=ALL_COLUMNS_LABELS.get(col_key, col_key))
                            cell.font = Font(bold=True, size=10, name='Times New Roman', color="FFFFFF")
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            cell.border = thin_border_tk

                        # ── Dữ liệu ──
                        def fmt_val(key, val):
                            if val is None:
                                return ""
                            if 'ngay' in key.lower() or 'thang' in key.lower():
                                return format_date(val)
                            return val

                        date_cols = {k for k in selected_cols_sorted if 'ngay' in k or 'thang' in k}
                        center_cols = {k for k in selected_cols_sorted if k in (
                            "ma_nv","ngay_sinh","gioi_tinh","loai_hop_dong",
                            "ngay_vao_lam","ngay_ky_hd","ngay_ket_thuc","trang_thai",
                            "thang_bat_dau_bh","thang_ket_thuc_bh","he_so_luong",
                            "phu_cap_tnvk","phu_cap_tnn","muc_huong_bhyt"
                        )}

                        for stt_idx, nv in enumerate(ds_tk, 1):
                            row = header_row_tk + stt_idx
                            ws_tk.cell(row=row, column=1, value=stt_idx).border = thin_border_tk
                            ws_tk.cell(row=row, column=1).alignment = Alignment(horizontal='center', vertical='center')
                            ws_tk.cell(row=row, column=1).font = Font(size=10, name='Times New Roman')
                            for col_idx, col_key in enumerate(selected_cols_sorted, 2):
                                raw = nv.get(col_key)
                                val = fmt_val(col_key, raw)
                                cell = ws_tk.cell(row=row, column=col_idx, value=val)
                                cell.font = Font(size=10, name='Times New Roman')
                                cell.border = thin_border_tk
                                if col_key in center_cols:
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                                else:
                                    cell.alignment = Alignment(horizontal='left', vertical='center')

                        total_row_tk = header_row_tk + len(ds_tk) + 1
                        ws_tk.merge_cells(start_row=total_row_tk, start_column=1, end_row=total_row_tk, end_column=n_cols)
                        ws_tk.cell(row=total_row_tk, column=1, value=f"TỔNG CỘNG: {len(ds_tk)} nhân viên")
                        ws_tk.cell(row=total_row_tk, column=1).font = Font(bold=True, size=11, name='Times New Roman')
                        ws_tk.cell(row=total_row_tk, column=1).alignment = Alignment(horizontal='left')

                        # ── Thống kê theo giới tính ──
                        stat_start = total_row_tk + 2
                        ws_tk.merge_cells(start_row=stat_start, start_column=1, end_row=stat_start, end_column=n_cols)
                        ws_tk.cell(row=stat_start, column=1, value="THỐNG KÊ THEO GIỚI TÍNH").font = Font(bold=True, size=11, name='Times New Roman')

                        nam_count  = sum(1 for nv in ds_tk if (nv.get('gioi_tinh') or '') == 'Nam')
                        nu_count   = sum(1 for nv in ds_tk if (nv.get('gioi_tinh') or '') == 'Nữ')
                        khac_count = len(ds_tk) - nam_count - nu_count

                        for r_offset, (label, cnt) in enumerate([("Nam", nam_count), ("Nữ", nu_count), ("Khác/Chưa xác định", khac_count)], 1):
                            r = stat_start + r_offset
                            ws_tk.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
                            ws_tk.cell(row=r, column=1, value=f"  {label}:").font = Font(size=10, name='Times New Roman')
                            ws_tk.cell(row=r, column=4, value=cnt).font = Font(size=10, name='Times New Roman')
                            for cc in range(1, 5):
                                ws_tk.cell(row=r, column=cc).fill = stat_fill

                        # ── Thống kê theo loại hợp đồng ──
                        stat2_start = stat_start + 5
                        ws_tk.merge_cells(start_row=stat2_start, start_column=1, end_row=stat2_start, end_column=n_cols)
                        ws_tk.cell(row=stat2_start, column=1, value="THỐNG KÊ THEO LOẠI HỢP ĐỒNG").font = Font(bold=True, size=11, name='Times New Roman')

                        hd_types = {"Không xác định thời hạn": 0, "Xác định thời hạn": 0, "Thử việc": 0, "Khác": 0}
                        for nv in ds_tk:
                            loai = (nv.get('loai_hop_dong') or '').strip()
                            if loai in hd_types:
                                hd_types[loai] += 1
                            elif (nv.get('trang_thai') or '') == 'THU_VIEC':
                                hd_types["Thử việc"] += 1
                            else:
                                hd_types["Khác"] += 1

                        for r_offset, (label, cnt) in enumerate(hd_types.items(), 1):
                            r = stat2_start + r_offset
                            ws_tk.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
                            ws_tk.cell(row=r, column=1, value=f"  {label}:").font = Font(size=10, name='Times New Roman')
                            ws_tk.cell(row=r, column=4, value=cnt).font = Font(size=10, name='Times New Roman')
                            for cc in range(1, 5):
                                ws_tk.cell(row=r, column=cc).fill = stat_fill

                        # ── Footer người lập báo cáo ──
                        footer_row = stat2_start + len(hd_types) + 3
                        ws_tk.merge_cells(start_row=footer_row, start_column=1, end_row=footer_row, end_column=3)
                        ws_tk.cell(row=footer_row, column=1, value="NGƯỜI LẬP BÁO CÁO")
                        ws_tk.cell(row=footer_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
                        ws_tk.cell(row=footer_row, column=1).alignment = Alignment(horizontal='center')

                        ws_tk.merge_cells(start_row=footer_row+1, start_column=1, end_row=footer_row+1, end_column=3)
                        ws_tk.cell(row=footer_row+1, column=1, value="(Ký, ghi rõ họ tên)")
                        ws_tk.cell(row=footer_row+1, column=1).font = Font(size=10, name='Times New Roman', italic=True)
                        ws_tk.cell(row=footer_row+1, column=1).alignment = Alignment(horizontal='center')

                        # ── Độ rộng cột ──
                        ws_tk.column_dimensions['A'].width = 5
                        for col_idx, col_key in enumerate(selected_cols_sorted, 2):
                            if col_key in ('ho_ten', 'thuong_tru', 'nguyen_quan', 'noi_cap_cccd', 'chuc_danh_nghe', 'ten_don_vi_thu_huong'):
                                w = 28
                            elif col_key in ('ma_nv', 'gioi_tinh', 'he_so_luong', 'phu_cap_tnvk', 'phu_cap_tnn'):
                                w = 12
                            elif 'ngay' in col_key or 'thang' in col_key:
                                w = 16
                            else:
                                w = 20
                            ws_tk.column_dimensions[get_column_letter(col_idx)].width = w

                        ws_tk.row_dimensions[header_row_tk].height = 30

                        fname_tk = f"ThongKe_NhanSu_{tk_tu_ngay.strftime('%d%m%Y')}_{tk_den_ngay.strftime('%d%m%Y')}.xlsx"
                        wb_tk.save(fname_tk)
                        with open(fname_tk, "rb") as f:
                            st.download_button(
                                label="📥 TẢI FILE THỐNG KÊ NHÂN SỰ",
                                data=f,
                                file_name=fname_tk,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                width='stretch'
                            )
                        st.success(f"✅ Đã xuất thống kê {len(ds_tk)} nhân viên với {len(selected_cols_sorted)} cột.")


    with tab_bc_tanggiam:
        col_from, col_to, col_xuat_bc = st.columns(3)
        with col_from:
            tu_ngay_bc = st.date_input("Từ ngày:", value=date.today().replace(day=1), key="bc_tu")
        with col_to:
            den_ngay_bc = st.date_input("Đến ngày:", value=date.today(), key="bc_den")
        with col_xuat_bc:
            st.write("")  # căn chỉnh cho nút thẳng hàng với 2 ô ngày (bù khoảng trống label)
            xuat_bc = st.button("📄 XUẤT BÁO CÁO WORD", width='stretch')

        if xuat_bc:
            db = st.session_state.db_engine.get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("""
                SELECT ho_ten, chuc_danh_nghe, phong_ban_lam_viec, loai_hop_dong, ngay_vao_lam,
                       ngay_sinh, so_hdld, ngay_ky_hd
                FROM nhan_vien 
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                AND ngay_vao_lam BETWEEN %s AND %s
                ORDER BY ngay_vao_lam ASC
            """, (tu_ngay_bc, den_ngay_bc))
            tang_list = c.fetchall()
            c.execute("""
                SELECT ho_ten, chuc_danh_nghe, phong_ban_lam_viec, loai_hop_dong, ngay_vao_lam, ngay_ket_thuc,
                       ngay_sinh, so_hdld, ngay_ky_hd
                FROM nhan_vien 
                WHERE trang_thai = 'NGHI_VIEC'
                AND ngay_ket_thuc BETWEEN %s AND %s
                ORDER BY ngay_ket_thuc ASC
            """, (tu_ngay_bc, den_ngay_bc))
            giam_list = c.fetchall()
            db.close()
            if tang_list or giam_list:
                file_path = tao_bao_cao_tang_giam(tang_list, giam_list, tu_ngay_bc, den_ngay_bc)
                with open(file_path, "rb") as f:
                    st.download_button(
                        label="📥 TẢI FILE BÁO CÁO (Word)",
                        data=f,
                        file_name=f"Bao_cao_tang_giam_{tu_ngay_bc.strftime('%d%m%Y')}_{den_ngay_bc.strftime('%d%m%Y')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.info("Không có biến động nhân sự trong kỳ.")


elif menu == "📄 Quản lý Công văn & HĐ kinh tế":
    show_quan_ly_cong_van()

# ========== CHAT NỘI BỘ NÂNG CẤP ==========
elif menu == "💬 Chat nội bộ":
    st.title("💬 Chat nội bộ")
    
    # Khởi tạo bảng chat
    #chat_utils.init_chat_tables()
    
    # Kiểm tra đăng nhập
    if 'nhan_vien_id' not in st.session_state:
        st.warning("Vui lòng đăng nhập để sử dụng chat!")
        st.stop()
    
    user_id = st.session_state.nhan_vien_id
    current_user_name = st.session_state.get('ho_ten_dang_nhap', 'Bạn')
    
    # ====== CSS CUSTOM ======
    st.markdown("""
    <style>
    /* Reset chat container */
    .chat-container {
        display: flex;
        height: 600px;
        border-radius: 12px;
        overflow: hidden;
        background: #fff;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        border: 1px solid #e8ecf1;
        margin-top: 10px;
    }
    
    /* Sidebar rooms */
    .chat-sidebar {
        width: 320px;
        min-width: 260px;
        background: #f8f9fa;
        border-right: 1px solid #e8ecf1;
        overflow-y: auto;
        padding: 8px 0;
        flex-shrink: 0;
    }
    
    .chat-main {
        flex: 1;
        display: flex;
        flex-direction: column;
        background: #fff;
        overflow: hidden;
    }
    
    .chat-header {
        padding: 12px 20px;
        background: #fff;
        border-bottom: 1px solid #e8ecf1;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-shrink: 0;
    }
    
    .chat-header h4 {
        margin: 0;
        font-weight: 600;
        color: #1e293b;
        font-size: 16px;
    }
    
    .chat-header .member-count {
        font-size: 12px;
        color: #94a3b8;
    }
    
    .chat-messages {
        flex: 1;
        overflow-y: auto;
        padding: 16px 20px;
        background: #f0f2f5;
        display: flex;
        flex-direction: column;
    }
    
    /* Message bubbles */
    .msg-wrapper {
        display: flex;
        margin-bottom: 8px;
        animation: fadeIn 0.2s ease;
    }
    
    .msg-wrapper.self {
        justify-content: flex-end;
    }
    
    .msg-wrapper.other {
        justify-content: flex-start;
    }
    
    .msg-bubble {
        max-width: 70%;
        padding: 8px 14px;
        border-radius: 16px;
        word-wrap: break-word;
        font-size: 14px;
        line-height: 1.5;
        position: relative;
        box-shadow: 0 1px 2px rgba(0,0,0,0.06);
    }
    
    .msg-wrapper.self .msg-bubble {
        background: #0084ff;
        color: #fff;
        border-bottom-right-radius: 4px;
    }
    
    .msg-wrapper.other .msg-bubble {
        background: #fff;
        color: #1e293b;
        border-bottom-left-radius: 4px;
        border: 1px solid #e8ecf1;
    }
    
    .msg-sender {
        font-size: 12px;
        font-weight: 600;
        color: #64748b;
        margin-bottom: 2px;
        display: block;
    }
    
    .msg-wrapper.self .msg-sender {
        color: rgba(255,255,255,0.8);
        text-align: right;
    }
    
    .msg-time {
        font-size: 10px;
        color: #94a3b8;
        margin-top: 2px;
        display: block;
        text-align: right;
    }
    
    .msg-wrapper.self .msg-time {
        color: rgba(255,255,255,0.7);
    }
    
    /* File/Image bubbles */
    .msg-image {
        max-width: 300px;
        max-height: 300px;
        border-radius: 12px;
        cursor: pointer;
        margin: 2px 0;
        display: block;
    }
    
    .msg-file {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 8px 12px;
        background: rgba(0,0,0,0.05);
        border-radius: 8px;
        text-decoration: none;
        color: #1e293b;
        font-size: 13px;
    }
    
    .msg-file:hover {
        background: rgba(0,0,0,0.08);
    }
    
    .msg-file .file-icon {
        font-size: 24px;
    }
    
    .msg-file .file-name {
        flex: 1;
        word-break: break-all;
    }
    
    .msg-file .file-size {
        font-size: 11px;
        color: #94a3b8;
    }
    
    .msg-payslip {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 12px;
        padding: 12px 16px;
        font-family: monospace;
        font-size: 13px;
        line-height: 1.6;
        white-space: pre-wrap;
        min-width: 200px;
        max-width: 100%;
    }
    
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(8px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    /* Room item */
    .room-item {
        display: flex;
        align-items: center;
        padding: 8px 16px;
        cursor: pointer;
        transition: background 0.15s;
        border-radius: 8px;
        margin: 2px 8px;
        gap: 12px;
    }
    
    .room-item:hover {
        background: #e8ecf1;
    }
    
    .room-item.active {
        background: #dbeafe;
    }
    
    .room-avatar {
        width: 40px;
        height: 40px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 16px;
        color: #fff;
        flex-shrink: 0;
        background: #94a3b8;
    }
    
    .room-avatar.broadcast {
        background: #f59e0b;
    }
    
    .room-avatar.group {
        background: #8b5cf6;
    }
    
    .room-avatar.private {
        background: #0ea5e9;
    }
    
    .room-info {
        flex: 1;
        min-width: 0;
    }
    
    .room-name {
        font-weight: 500;
        font-size: 14px;
        color: #1e293b;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    
    .room-last-msg {
        font-size: 12px;
        color: #94a3b8;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    
    .room-badge {
        background: #ef4444;
        color: #fff;
        border-radius: 50%;
        padding: 2px 8px;
        font-size: 11px;
        font-weight: 700;
        min-width: 20px;
        text-align: center;
        flex-shrink: 0;
    }
    
    /* Chat input */
    .chat-input-area {
        display: flex;
        gap: 8px;
        padding: 12px 16px;
        background: #fff;
        border-top: 1px solid #e8ecf1;
        align-items: flex-end;
        flex-shrink: 0;
    }
    
    .chat-input-area textarea {
        flex: 1;
        border: 1px solid #d1d5db;
        border-radius: 20px;
        padding: 8px 16px;
        resize: none;
        font-size: 14px;
        max-height: 100px;
        min-height: 40px;
        font-family: inherit;
        transition: border-color 0.2s;
        background: #f8f9fa;
    }
    
    .chat-input-area textarea:focus {
        outline: none;
        border-color: #0084ff;
        background: #fff;
    }
    
    .chat-input-area .send-btn {
        background: #0084ff;
        color: #fff;
        border: none;
        border-radius: 50%;
        width: 40px;
        height: 40px;
        font-size: 18px;
        cursor: pointer;
        transition: background 0.2s;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
    }
    
    .chat-input-area .send-btn:hover {
        background: #0066cc;
    }
    
    .chat-input-area .send-btn:disabled {
        background: #94a3b8;
        cursor: not-allowed;
    }
    
    .chat-input-area .attach-btn {
        background: none;
        border: none;
        font-size: 22px;
        cursor: pointer;
        color: #94a3b8;
        padding: 8px;
        transition: color 0.2s;
        flex-shrink: 0;
    }
    
    .chat-input-area .attach-btn:hover {
        color: #1e293b;
    }
    
    /* Empty state */
    .empty-state {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        height: 100%;
        color: #94a3b8;
        text-align: center;
        padding: 40px;
    }
    
    .empty-state .icon {
        font-size: 48px;
        margin-bottom: 12px;
    }
    
    /* Scrollbar */
    .chat-messages::-webkit-scrollbar,
    .chat-sidebar::-webkit-scrollbar {
        width: 4px;
    }
    
    .chat-messages::-webkit-scrollbar-track,
    .chat-sidebar::-webkit-scrollbar-track {
        background: transparent;
    }
    
    .chat-messages::-webkit-scrollbar-thumb,
    .chat-sidebar::-webkit-scrollbar-thumb {
        background: #d1d5db;
        border-radius: 4px;
    }
    
    .chat-messages::-webkit-scrollbar-thumb:hover,
    .chat-sidebar::-webkit-scrollbar-thumb:hover {
        background: #9ca3af;
    }
    
    @media (max-width: 768px) {
        .chat-container {
            flex-direction: column;
            height: auto;
            min-height: 500px;
        }
        .chat-sidebar {
            width: 100%;
            max-height: 200px;
            border-right: none;
            border-bottom: 1px solid #e8ecf1;
        }
        .chat-main {
            height: 400px;
        }
    }
    </style>
    """, unsafe_allow_html=True)
    
    # ====== XỬ LÝ SESSION STATE ======
    if 'chat_room_id' not in st.session_state:
        st.session_state.chat_room_id = None
    if 'chat_search' not in st.session_state:
        st.session_state.chat_search = ''
    if 'chat_msg_input' not in st.session_state:
        st.session_state.chat_msg_input = ''
    
    # ====== LẤY DANH SÁCH PHÒNG ======
    rooms = chat_utils.get_user_chat_rooms(user_id)
    
    # Thêm thông tin bổ sung cho mỗi phòng
    enriched_rooms = []
    for room in rooms:
        if room.get('deleted_at'):
            continue
        room_dict = dict(room)
        room_dict['display_name'] = chat_utils.get_room_display_name(room, user_id)
        room_dict['unread_count'] = chat_utils.get_room_unread_count(room['id'], user_id)
        last_msg = chat_utils.get_room_last_message(room['id'])
        room_dict['last_msg'] = last_msg
        room_dict['participants'] = chat_utils.get_room_participants(room['id'])
        enriched_rooms.append(room_dict)
    
    # Sắp xếp: broadcast lên đầu, sau đó theo updated_at
    enriched_rooms.sort(key=lambda r: (
        0 if r['room_type'] == 'broadcast' else 1,
        r.get('updated_at') or datetime(1970, 1, 1),
    ), reverse=True)
    
    # ====== LAYOUT CHAT ======
    col_sidebar, col_main = st.columns([1, 3], gap="small")
    
    with col_sidebar:
        st.markdown("### 💬 Phòng chat")
        
        # Ô tìm kiếm
        search_val = st.text_input(
            "🔍 Tìm kiếm phòng",
            value=st.session_state.chat_search,
            placeholder="Nhập tên phòng...",
            label_visibility="collapsed",
            key="chat_search_input"
        )
        if search_val != st.session_state.chat_search:
            st.session_state.chat_search = search_val
            st.rerun()
        
        # Nút tạo phòng mới
        with st.popover("➕ Tạo phòng mới", use_container_width=True):
            tab_new_private, tab_new_group, tab_new_broadcast = st.tabs(["👤 1-1", "👥 Nhóm", "📢 Broadcast"])
            
            with tab_new_private:
                st.caption("Chọn người để bắt đầu chat riêng")
                
                # Tìm kiếm nhân viên
                search_emp = st.text_input("Tìm nhân viên", placeholder="Tên hoặc mã NV...", key="search_emp_private")
                employees = chat_utils.search_employees(search_emp) if search_emp else chat_utils.get_all_employees()
                
                # Lọc bỏ bản thân
                employees = [e for e in employees if e['id'] != user_id]
                
                if employees:
                    emp_options = {f"{e['ho_ten']} ({e['ma_nv']})": e['id'] for e in employees}
                    selected_emp = st.selectbox("Chọn người dùng", list(emp_options.keys()), key="select_emp_private")
                    
                    if st.button("💬 Bắt đầu chat", width='stretch', type="primary"):
                        target_id = emp_options[selected_emp]
                        room_id = chat_utils.create_private_room(user_id, target_id)
                        if room_id:
                            st.session_state.chat_room_id = room_id
                            st.success("✅ Đã tạo phòng chat!")
                            st.rerun()
                else:
                    st.info("Không tìm thấy nhân viên nào")
            
            with tab_new_group:
                group_name = st.text_input("Tên nhóm", placeholder="Nhập tên nhóm...", key="group_name_input")
                
                all_emps = chat_utils.get_all_employees()
                emp_options = {f"{e['ho_ten']} ({e['ma_nv']})": e['id'] for e in all_emps if e['id'] != user_id}
                
                selected_members = st.multiselect(
                    "Chọn thành viên",
                    options=list(emp_options.keys()),
                    placeholder="Chọn ít nhất 1 người...",
                    key="group_members_select"
                )
                
                if st.button("➕ Tạo nhóm", width='stretch', type="primary"):
                    if not group_name.strip():
                        st.error("Vui lòng nhập tên nhóm!")
                    elif not selected_members:
                        st.error("Vui lòng chọn ít nhất 1 thành viên!")
                    else:
                        member_ids = [emp_options[name] for name in selected_members]
                        room_id = chat_utils.create_group_room(group_name.strip(), user_id, member_ids)
                        if room_id:
                            st.session_state.chat_room_id = room_id
                            st.success(f"✅ Đã tạo nhóm '{group_name}'")
                            st.rerun()
            
            with tab_new_broadcast:
                st.caption("Gửi thông báo đến TẤT CẢ nhân viên")
                st.warning("⚠️ Chỉ Admin và HR mới có quyền gửi thông báo chung", icon="🔒")
                
                if st.session_state.role in ['admin', 'hr']:
                    broadcast_room_id = chat_utils.get_or_create_broadcast_room()
                    if broadcast_room_id:
                        if st.button("📢 Mở phòng thông báo chung", width='stretch', type="primary"):
                            st.session_state.chat_room_id = broadcast_room_id
                            st.rerun()
                else:
                    st.info("Bạn không có quyền gửi thông báo chung")
        
        st.divider()
        
        # ====== DANH SÁCH PHÒNG ======
        if enriched_rooms:
            filtered_rooms = enriched_rooms
            if st.session_state.chat_search:
                search_lower = st.session_state.chat_search.lower()
                filtered_rooms = [
                    r for r in enriched_rooms 
                    if search_lower in r['display_name'].lower()
                ]
            
            for room in filtered_rooms:
                is_active = st.session_state.chat_room_id == room['id']
                is_broadcast = room['room_type'] == 'broadcast'
                is_group = room['room_type'] == 'group'
                
                # Avatar
                if is_broadcast:
                    avatar_bg = '#f59e0b'
                    avatar_text = '📢'
                elif is_group:
                    avatar_bg = '#8b5cf6'
                    avatar_text = '👥'
                else:
                    # Private: lấy tên người kia
                    other = next((p for p in room.get('participants', []) if p['id'] != user_id), None)
                    if other:
                        name_parts = other['ho_ten'].split()
                        initials = ''.join([p[0].upper() for p in name_parts[:2]])
                    else:
                        initials = '??'
                    avatar_text = initials
                    avatar_bg = '#0ea5e9'
                
                # Badge số tin chưa đọc
                unread = room.get('unread_count', 0)
                badge_html = f'<span class="room-badge">{unread}</span>' if unread > 0 else ''
                
                # Tin nhắn cuối
                last_msg = room.get('last_msg')
                last_preview = ''
                if last_msg:
                    content = last_msg.get('content', '')
                    msg_type = last_msg.get('message_type', 'text')
                    if msg_type == 'image':
                        last_preview = '📷 [Ảnh]'
                    elif msg_type == 'file':
                        last_preview = '📎 [File]'
                    elif msg_type == 'payslip':
                        last_preview = '📄 [Phiếu lương]'
                    else:
                        last_preview = content[:50] + ('...' if len(content) > 50 else '')
                    last_preview = last_preview.replace('\n', ' ')
                
                # Tạo button HTML
                room_id = room['id']
                display_name = room['display_name']
                active_class = 'active' if is_active else ''
                
                # Dùng st.button để xử lý click
                button_label = f"{display_name}"
                if st.button(
                    button_label,
                    key=f"room_{room_id}",
                    use_container_width=True,
                    type="primary" if is_active else "secondary"
                ):
                    st.session_state.chat_room_id = room_id
                    # Đánh dấu đã đọc
                    chat_utils.mark_messages_as_read(room_id, user_id)
                    st.rerun()
                
                # Hiển thị badge và preview bên dưới button (dùng markdown)
                if unread > 0:
                    st.markdown(f"<span style='float:right;background:#ef4444;color:#fff;border-radius:50%;padding:0 8px;font-size:11px;'>{unread}</span>", unsafe_allow_html=True)
                if last_preview:
                    st.caption(f"💬 {last_preview}")
                
                st.divider()
        else:
            st.info("📭 Chưa có phòng chat nào")
            st.caption("👆 Bấm 'Tạo phòng mới' để bắt đầu")
    
    with col_main:
        # ====== KHUNG CHAT CHÍNH ======
        if st.session_state.chat_room_id:
            room_id = st.session_state.chat_room_id
            
            # Tìm thông tin phòng hiện tại
            current_room = next((r for r in enriched_rooms if r['id'] == room_id), None)
            if not current_room:
                # Phòng có thể mới tạo, load lại
                st.rerun()
            
            display_name = current_room['display_name']
            room_type = current_room['room_type']
            participants = current_room.get('participants', [])
            
            # Đánh dấu đã đọc khi vào phòng
            chat_utils.mark_messages_as_read(room_id, user_id)
            
            # ====== HEADER ======
            col_header1, col_header2 = st.columns([3, 1])
            with col_header1:
                st.markdown(f"### 💬 {display_name}")
                if room_type == 'group':
                    member_names = [p['ho_ten'] for p in participants if p['id'] != user_id]
                    st.caption(f"👥 {len(participants)} thành viên: {', '.join(member_names[:5])}{'...' if len(member_names) > 5 else ''}")
                elif room_type == 'broadcast':
                    st.caption("📢 Thông báo đến tất cả nhân viên")
                else:
                    other = next((p for p in participants if p['id'] != user_id), None)
                    if other:
                        st.caption(f"👤 {other.get('ho_ten', '')} - {other.get('phong_ban_lam_viec', '')}")
            
            with col_header2:
                if room_type == 'group' and st.session_state.role in ['admin', 'hr']:
                    with st.popover("⚙️ Quản lý nhóm"):
                        st.caption("Thêm thành viên")
                        all_emps = chat_utils.get_all_employees()
                        current_ids = [p['id'] for p in participants]
                        available = [e for e in all_emps if e['id'] not in current_ids and e['id'] != user_id]
                        
                        if available:
                            add_options = {f"{e['ho_ten']} ({e['ma_nv']})": e['id'] for e in available}
                            to_add = st.multiselect("Chọn thành viên thêm", list(add_options.keys()))
                            if st.button("➕ Thêm", width='stretch'):
                                ids = [add_options[name] for name in to_add]
                                if chat_utils.add_participants_to_room(room_id, ids):
                                    st.success("✅ Đã thêm thành viên!")
                                    st.rerun()
                        else:
                            st.info("Đã có tất cả nhân viên trong nhóm")
                        
                        st.divider()
                        if st.button("🚪 Rời nhóm", width='stretch', type="secondary"):
                            if chat_utils.remove_participant_from_room(room_id, user_id):
                                st.session_state.chat_room_id = None
                                st.success("Đã rời nhóm")
                                st.rerun()
            
            st.divider()
            
            # ====== TIN NHẮN ======
            messages = chat_utils.get_room_messages(room_id)
            
            # Container tin nhắn
            msg_container = st.container(height=400, border=True)
            
            with msg_container:
                if messages:
                    for msg in messages:
                        is_self = msg['sender_id'] == user_id
                        sender_name = msg.get('sender_name', 'Hệ thống')
                        msg_type = msg.get('message_type', 'text')
                        content = msg.get('content', '')
                        file_url = msg.get('file_url')
                        file_name = msg.get('file_name', '')
                        sent_at = msg.get('sent_at')
                        time_str = sent_at.strftime('%H:%M') if sent_at else ''
                        
                        # Xác định class
                        wrapper_class = "msg-wrapper self" if is_self else "msg-wrapper other"
                        
                        # Xây dựng nội dung bong bóng
                        if msg_type == 'image' and file_url:
                            # ===== CẢI TIẾN: Dùng cache, lazy loading =====
                            img_bytes = get_chat_image_bytes_cached(file_url)
                            if img_bytes:
                                img_b64 = base64.b64encode(img_bytes).decode()
                                # Thêm loading="lazy" để chỉ tải khi cuộn đến
                                bubble_content = f'<img src="data:image/jpeg;base64,{img_b64}" class="msg-image" loading="lazy" onclick="window.open(this.src)" />'
                            else:
                                bubble_content = f'📷 Ảnh: {file_name} (không thể tải)'
                        
                        elif msg_type == 'file' and file_url:
                            file_size = msg.get('file_size', 0)
                            size_str = f"{file_size/1024:.1f} KB" if file_size < 1024*1024 else f"{file_size/(1024*1024):.1f} MB"
                            bubble_content = f"""
                            <div class="msg-file">
                                <span class="file-icon">📎</span>
                                <span class="file-name">{file_name}</span>
                                <span class="file-size">{size_str}</span>
                            </div>
                            """
                            # Thêm nút tải
                            download_key = f"download_{msg['id']}"
                            if st.button("📥 Tải xuống", key=download_key):
                                file_bytes = chat_utils.get_chat_file_bytes(file_url)
                                if file_bytes:
                                    st.download_button(
                                        label="Tải file",
                                        data=file_bytes,
                                        file_name=file_name,
                                        key=f"download_btn_{msg['id']}"
                                    )
                        
                        elif msg_type == 'payslip':
                            bubble_content = f'<div class="msg-payslip">{content}</div>'
                        
                        else:  # text
                            bubble_content = content.replace('\n', '<br>')
                        
                        # Tên người gửi (cho nhóm và broadcast)
                        sender_display = ''
                        if room_type in ['group', 'broadcast'] and not is_self:
                            sender_display = f'<span class="msg-sender">{sender_name}</span>'
                        
                        # Render bong bóng
                        st.markdown(f"""
                        <div class="{wrapper_class}">
                            <div class="msg-bubble">
                                {sender_display}
                                <div>{bubble_content}</div>
                                <span class="msg-time">{time_str}</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown("""
                    <div class="empty-state">
                        <div class="icon">💬</div>
                        <h4>Chưa có tin nhắn</h4>
                        <p>Hãy bắt đầu trò chuyện!</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            # ====== Ô NHẬP LIỆU ======
            st.divider()
            
            # Kiểm tra quyền gửi broadcast
            can_send = True
            if current_room['room_type'] == 'broadcast':
                if st.session_state.role not in ['admin', 'hr']:
                    can_send = False
                    st.warning("🔒 Chỉ Admin và HR mới được gửi thông báo chung")
            
            if can_send:
                col_input, col_send, col_attach, col_payslip = st.columns([5, 0.8, 0.8, 0.8])
                
                with col_input:
                    msg_input = st.text_area(
                        "Nhập tin nhắn...",
                        value=st.session_state.get('chat_msg_temp', ''),
                        height=60,
                        max_chars=2000,
                        key="chat_msg_area",
                        label_visibility="collapsed",
                        placeholder="Nhập tin nhắn... (Enter để gửi, Shift+Enter để xuống dòng)"
                    )
                
                with col_send:
                    if st.button("📤", key="send_msg_btn", help="Gửi tin nhắn", use_container_width=True):
                        if msg_input.strip():
                            if chat_utils.send_message(room_id, user_id, msg_input.strip()):
                                st.session_state.chat_msg_temp = ''
                                st.rerun()
                
                with col_attach:
                    with st.popover("📎 Đính kèm", use_container_width=True):
                        st.caption("Chọn file hoặc ảnh")
                        uploaded_file = st.file_uploader(
                            "Chọn file",
                            type=["png", "jpg", "jpeg", "gif", "pdf", "doc", "docx", "xlsx", "txt", "zip"],
                            key="chat_file_upload",
                            label_visibility="collapsed"
                        )
                        
                        if uploaded_file is not None:
                            file_bytes = uploaded_file.getvalue()
                            file_name = uploaded_file.name
                            content_type = uploaded_file.type or mimetypes.guess_type(file_name)[0] or 'application/octet-stream'
                            file_size = len(file_bytes)
                            
                            # Xác định loại
                            is_image = content_type.startswith('image/')
                            is_video = content_type.startswith('video/')
                            
                            if is_image:
                                st.image(file_bytes, caption=file_name, width=200)
                            else:
                                st.info(f"📎 {file_name} ({file_size/1024:.1f} KB)")
                            
                            if st.button("📤 Gửi file", width='stretch', type="primary"):
                                # Upload lên Storage
                                file_path = chat_utils.upload_chat_file(file_bytes, file_name, content_type)
                                if file_path:
                                    if is_image:
                                        success = chat_utils.send_image_message(room_id, user_id, file_path, file_name, msg_input.strip() or '')
                                    else:
                                        success = chat_utils.send_file_message(room_id, user_id, file_path, file_name, file_size)
                                    
                                    if success:
                                        st.session_state.chat_msg_temp = ''
                                        st.success("✅ Đã gửi file!")
                                        st.rerun()
                                else:
                                    st.error("❌ Lỗi upload file!")
                
                with col_payslip:
                    if st.session_state.role in ['admin', 'kt_luong', 'hr']:
                        with st.popover("📄 Phiếu lương", use_container_width=True):
                            st.subheader("📄 Gửi phiếu lương")
                            st.caption("Chọn nhân viên và tháng để gửi phiếu lương")
                            
                            # Chọn nhân viên
                            all_emps = chat_utils.get_all_employees()
                            emp_options = {f"{e['ho_ten']} ({e['ma_nv']})": e['id'] for e in all_emps}
                            selected_emp_label = st.selectbox("Nhân viên", list(emp_options.keys()), key="payslip_emp")
                            target_emp_id = emp_options[selected_emp_label]
                            target_emp_name = selected_emp_label.split(' (')[0]
                            
                            # Chọn tháng
                            col_month, col_year = st.columns(2)
                            with col_month:
                                p_month = st.selectbox("Tháng", list(range(1, 13)), index=datetime.now().month-1, key="payslip_month")
                            with col_year:
                                p_year = st.number_input("Năm", min_value=2020, max_value=2100, value=datetime.now().year, key="payslip_year")
                            
                            # Demo: tính lương mẫu
                            st.caption("💡 Nhập các khoản thu nhập (demo)")
                            luong_cb = st.number_input("Lương cơ bản", min_value=0, value=5000000, step=100000)
                            pc_cv = st.number_input("Phụ cấp chức vụ", min_value=0, value=0, step=100000)
                            pc_tnvk = st.number_input("PC thâm niên VK (%)", min_value=0.0, value=0.0, step=0.5)
                            pc_tnn = st.number_input("PC thâm niên nghề (%)", min_value=0.0, value=0.0, step=0.5)
                            
                            if st.button("📤 Gửi phiếu lương", width='stretch', type="primary"):
                                # Tính các khoản
                                tong = luong_cb + pc_cv + (luong_cb * pc_tnvk / 100) + (luong_cb * pc_tnn / 100)
                                bhxh = luong_cb * 0.08
                                bhyt = luong_cb * 0.015
                                bhtn = luong_cb * 0.01
                                thuc_nhan = tong - bhxh - bhyt - bhtn
                                
                                salary_data = {
                                    'luong_co_ban': luong_cb,
                                    'phu_cap_chuc_vu': pc_cv,
                                    'phu_cap_tnvk': luong_cb * pc_tnvk / 100,
                                    'phu_cap_tnn': luong_cb * pc_tnn / 100,
                                    'tong': tong,
                                    'bhxh': bhxh,
                                    'bhyt': bhyt,
                                    'bhtn': bhtn,
                                    'thuc_nhan': thuc_nhan
                                }
                                
                                # Tìm hoặc tạo phòng private
                                room_id_target = chat_utils.create_private_room(user_id, target_emp_id)
                                if room_id_target:
                                    if chat_utils.send_payslip_message(
                                        room_id_target, user_id,
                                        target_emp_name, p_month, p_year, salary_data
                                    ):
                                        st.success(f"✅ Đã gửi phiếu lương cho {target_emp_name}!")
                                        st.balloons()
                                        st.rerun()
                                    else:
                                        st.error("❌ Lỗi gửi phiếu lương!")
                                else:
                                    st.error("❌ Không thể tạo phòng chat!")
                    else:
                        with st.popover("📄 Phiếu lương", use_container_width=True):
                            st.info("🔒 Chỉ Admin, HR và Kế toán lương mới có quyền gửi phiếu lương")
                
                # Xử lý Enter để gửi
                if msg_input and st.session_state.get('chat_msg_temp') != msg_input:
                    st.session_state.chat_msg_temp = msg_input
                
                # Tự động gửi khi nhấn Enter (không Shift)
                # Sử dụng JavaScript để bắt sự kiện
                st.markdown("""
                <script>
                document.addEventListener('DOMContentLoaded', function() {
                    const textarea = document.querySelector('[data-testid="stTextArea"] textarea');
                    if (textarea) {
                        textarea.addEventListener('keydown', function(e) {
                            if (e.key === 'Enter' && !e.shiftKey) {
                                e.preventDefault();
                                const sendBtn = document.querySelector('[data-testid="stButton"] button[kind="primary"]');
                                if (sendBtn) sendBtn.click();
                            }
                        });
                    }
                });
                </script>
                """, unsafe_allow_html=True)
        
        else:
            # Chưa chọn phòng
            st.markdown("""
            <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:400px;color:#94a3b8;text-align:center;">
                <div style="font-size:64px;margin-bottom:16px;">💬</div>
                <h3 style="color:#1e293b;">Chọn một phòng chat</h3>
                <p>Chọn phòng từ danh sách bên trái hoặc tạo phòng mới</p>
                <p style="font-size:12px;margin-top:8px;">💡 Bấm "Tạo phòng mới" để bắt đầu trò chuyện</p>
            </div>
            """, unsafe_allow_html=True)



# ========== CHATBOT GIẢI ĐÁP ==========
elif menu == "🤖 Chatbot Giải đáp":
    st.title("🤖 AI Tư vấn Hành chính Nhân sự")
    st.caption("BHXH · BHYT · Thuế TNCN · Lao động · Thai sản · Thất nghiệp — AI phân tích và trích dẫn điều luật cụ thể.")

    if "chatbot_history" not in st.session_state:
        st.session_state.chatbot_history = []
    if "chatbot_display" not in st.session_state:
        st.session_state.chatbot_display = []

    cau_hoi_bam = None
    if not st.session_state.chatbot_display:
        st.markdown("""
        <div style="text-align:center;padding-top:10px;">
            <div style="font-size:36px;margin-bottom:8px;">⚖️</div>
            <h3 style="color:#1e3a5f;">Chào mừng đến với AI Tư vấn HCNS</h3>
            <p style="font-size:13px;color:#6b7280;line-height:1.7;">Hỏi về quyền lợi BHXH, BHYT, thai sản, thất nghiệp,
            thuế TNCN, hợp đồng lao động — tôi sẽ phân tích và trích dẫn điều luật cụ thể.</p>
        </div>
        """, unsafe_allow_html=True)
        cau_hoi_mau = [
            ("🤰", "Chế độ thai sản", "Tôi đang mang thai, cần đóng BHXH bao lâu để hưởng thai sản? Nghỉ được mấy tháng và hưởng mức lương bao nhiêu?"),
            ("💰", "Tính thuế TNCN", "Lương gross 20 triệu, có 1 người phụ thuộc, đóng BHXH 8%, thuế TNCN phải nộp bao nhiêu?"),
            ("📋", "Trợ cấp thất nghiệp", "Tôi đã đóng BHTN được 3 năm liên tục, vừa mất việc. Được hưởng trợ cấp thất nghiệp bao nhiêu tháng và mức hưởng tính thế nào?"),
            ("🏥", "BHXH ốm đau", "Nhân viên đã đóng BHXH 10 năm, bị ốm nghỉ 45 ngày liên tiếp trong năm. Mức hưởng BHXH ốm đau tính thế nào?"),
            ("📝", "Hợp đồng lao động", "Công ty muốn chấm dứt hợp đồng với nhân viên đã làm 3 năm. Cần thực hiện đúng quy trình gì và có phải trả trợ cấp thôi việc không?"),
            ("👨‍👩‍👧", "Giảm trừ gia cảnh", "Điều kiện và thủ tục đăng ký người phụ thuộc để giảm trừ gia cảnh thuế TNCN là gì? Hồ sơ gồm những gì?"),
        ]
        cols_q = st.columns(3)
        for i, (icon, label, full_q) in enumerate(cau_hoi_mau):
            with cols_q[i % 3]:
                if st.button(f"{icon} {label}", key=f"chatbot_qb_{i}", width='stretch'):
                    cau_hoi_bam = full_q
    else:
        for msg in st.session_state.chatbot_display:
            if msg["role"] == "user":
                with st.chat_message("user"):
                    st.markdown(msg["text"])
            else:
                with st.chat_message("assistant", avatar="⚖️"):
                    st.markdown(_chatbot_render_answer_html(msg["data"]), unsafe_allow_html=True)

    cau_hoi_go = st.chat_input("Đặt câu hỏi về BHXH, thuế TNCN, thai sản, hợp đồng lao động...")
    cau_hoi_cuoi = cau_hoi_bam or cau_hoi_go

    if cau_hoi_cuoi:
        st.session_state.chatbot_display.append({"role": "user", "text": cau_hoi_cuoi})
        st.session_state.chatbot_history.append({"role": "user", "content": cau_hoi_cuoi})
        with st.spinner("⚖️ Đang phân tích điều luật liên quan..."):
            laws = _chatbot_detect_laws(cau_hoi_cuoi)
            system_prompt = _chatbot_system_prompt(laws)
            ket_qua = _chatbot_call_claude(system_prompt, st.session_state.chatbot_history)
        st.session_state.chatbot_history.append({"role": "assistant", "content": json.dumps(ket_qua, ensure_ascii=False)})
        st.session_state.chatbot_display.append({"role": "ai", "data": ket_qua})
        st.rerun()

    st.caption("ℹ️ Kết quả tư vấn mang tính tham khảo. Vui lòng xác nhận với chuyên gia pháp lý cho các quyết định quan trọng.")
    if st.session_state.chatbot_display:
        if st.button("🗑️ Xoá lịch sử trò chuyện"):
            st.session_state.chatbot_history = []
            st.session_state.chatbot_display = []
            st.rerun()

# ========== HƯỚNG DẪN SỬ DỤNG ==========
elif menu == "🔑 Quản lý MK":
    st.title("🔑 Quản lý mật khẩu")

    if st.session_state.role == "admin":
        tab_doi_mk, tab_admin_reset = st.tabs(["🔒 Đổi mật khẩu của tôi", "🛠️ Reset mật khẩu nhân viên (Admin)"])
    else:
        tab_doi_mk = st.container()

    with tab_doi_mk:
        st.subheader("🔒 Đổi mật khẩu của tôi")
        st.caption("Nếu nghi ngờ mật khẩu bị lộ, hãy chủ động đổi ngay tại đây.")
        mk_hien_tai = st.text_input("Mật khẩu hiện tại:", type="password", key="doimk_hientai")
        mk_moi_ts = st.text_input("Mật khẩu mới:", type="password", key="doimk_moi")
        mk_moi_ts2 = st.text_input("Nhập lại mật khẩu mới:", type="password", key="doimk_moi2")
        if st.button("✅ Xác nhận đổi mật khẩu", key="btn_doi_mk_tuchu", type="primary"):
            if not st.session_state.get('nhan_vien_id'):
                st.error("❌ Không xác định được tài khoản đang đăng nhập (tài khoản Admin hệ thống không đổi được ở đây).")
            else:
                db_dmk = st.session_state.db_engine.get_connection()
                c_dmk = db_dmk.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                c_dmk.execute("SELECT mat_khau_hash FROM nhan_vien WHERE id=%s", (st.session_state.nhan_vien_id,))
                row_dmk = c_dmk.fetchone()
                if not row_dmk or not bcrypt.checkpw(mk_hien_tai.encode(), row_dmk['mat_khau_hash'].encode()):
                    db_dmk.close()
                    st.error("❌ Mật khẩu hiện tại không đúng.")
                elif len(mk_moi_ts) < 6:
                    db_dmk.close()
                    st.error("Mật khẩu mới phải có ít nhất 6 ký tự.")
                elif mk_moi_ts != mk_moi_ts2:
                    db_dmk.close()
                    st.error("Hai mật khẩu nhập lại không khớp.")
                else:
                    c_dmk2 = db_dmk.cursor()
                    new_hash_ts = bcrypt.hashpw(mk_moi_ts.encode(), bcrypt.gensalt()).decode()
                    c_dmk2.execute("UPDATE nhan_vien SET mat_khau_hash=%s WHERE id=%s",
                                   (new_hash_ts, st.session_state.nhan_vien_id))
                    db_dmk.commit(); db_dmk.close()
                    st.success("✅ Đổi mật khẩu thành công!")

    if st.session_state.role == "admin":
        with tab_admin_reset:
            st.subheader("🛠️ Reset mật khẩu nhân viên (dành cho trường hợp quên mật khẩu & không có Email liên hệ)")
            st.caption("Mật khẩu sẽ được đặt lại về mặc định = **số điện thoại** của nhân viên, "
                       "và nhân viên sẽ bị buộc đổi mật khẩu ngay trong lần đăng nhập tiếp theo.")
            db_rst = st.session_state.db_engine.get_connection()
            c_rst = db_rst.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c_rst.execute("""SELECT id, ho_ten, ma_nv, dien_thoai FROM nhan_vien
                              WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY ho_ten""")
            ds_nv_rst = c_rst.fetchall()
            db_rst.close()
            tuy_chon_rst = {f"{r['ho_ten']} ({r['ma_nv']}) - SĐT: {r.get('dien_thoai') or 'chưa có'}": r for r in ds_nv_rst}
            chon_rst = st.selectbox("Chọn nhân viên:", ["-- Chọn --"] + list(tuy_chon_rst.keys()),
                                     key="chon_reset_mk", help="💡 Gõ tên/mã NV để tìm nhanh")
            if chon_rst != "-- Chọn --":
                nv_rst = tuy_chon_rst[chon_rst]
                if not nv_rst.get('dien_thoai'):
                    st.error("❌ Nhân viên chưa có số điện thoại trong hồ sơ nên không thể đặt mật khẩu mặc định. Vui lòng cập nhật SĐT trước.")
                else:
                    st.warning(f"Sẽ đặt lại mật khẩu của **{nv_rst['ho_ten']}** về **{nv_rst['dien_thoai']}** và buộc đổi mật khẩu ở lần đăng nhập tới.")
                    if st.button("🔄 Xác nhận Reset mật khẩu", key=f"btn_reset_mk_{nv_rst['id']}", type="primary"):
                        db_r2 = st.session_state.db_engine.get_connection()
                        c_r2 = db_r2.cursor()
                        new_hash_rst = bcrypt.hashpw(nv_rst['dien_thoai'].encode(), bcrypt.gensalt()).decode()
                        c_r2.execute("UPDATE nhan_vien SET mat_khau_hash=%s, phai_doi_mat_khau=TRUE WHERE id=%s",
                                     (new_hash_rst, nv_rst['id']))
                        db_r2.commit(); db_r2.close()
                        st.success(f"✅ Đã reset mật khẩu về SĐT ({nv_rst['dien_thoai']}). Thông báo cho nhân viên đăng nhập lại và đổi mật khẩu mới.")

elif menu == "🖼️ Tạo ảnh thẻ NV":
    photo_card_gender.render()

elif menu == "🔍 Audit Dashboard":
    st.title("🔍 Audit lệch số liệu Dashboard")
    st.caption(
        "Công cụ nội bộ (chỉ Admin) để kiểm tra các biểu đồ trong 📊 Dashboard có "
        "cùng Tổng với nhau không, và tìm nguyên nhân nếu bị lệch (LIMIT, filter thiếu "
        "đồng bộ, dữ liệu NULL/rỗng...). Dùng ngay kết nối DB hiện tại của công ty bạn "
        "đang đăng nhập — không cần cấu hình gì thêm."
    )
    if st.session_state.role != "admin":
        st.error("❌ Chỉ Admin mới được dùng công cụ này!")
    else:
        db_a = st.session_state.db_engine.get_connection()
        c_a = db_a.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        # Tiêu chuẩn lọc nhân sự — PHẢI giống hệt biến DK_CHUAN_NV dùng ở 📊 Dashboard.
        # Nếu sửa điều kiện lọc ở Dashboard, nhớ sửa lại y hệt ở đây.
        DK_CHUAN_NV_AUDIT = "trang_thai IN ('DANG_LAM', 'THU_VIEC') AND so_hdld IS NOT NULL AND so_hdld != ''"

        c_a.execute(f"SELECT COUNT(*) AS t FROM nhan_vien WHERE {DK_CHUAN_NV_AUDIT}")
        TONG_CHUAN = c_a.fetchone()['t']

        st.subheader("0️⃣ Tổng chuẩn (dùng để đối chiếu mọi biểu đồ)")
        st.success(f"**Tổng chuẩn = {TONG_CHUAN} nhân viên**")
        st.caption("Điều kiện: `" + DK_CHUAN_NV_AUDIT + "`")
        st.divider()

        # 1) So sánh Tổng từng biểu đồ với Tổng chuẩn
        st.subheader("1️⃣ So sánh Tổng từng biểu đồ với Tổng chuẩn")
        checks_a = []
        c_a.execute(f"SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV_AUDIT}")
        checks_a.append(("Cơ cấu theo Phòng ban", c_a.fetchone()['t'], "không filter thêm"))
        c_a.execute(f"SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV_AUDIT}")
        checks_a.append(("Cơ cấu theo Giới tính", c_a.fetchone()['t'], "không filter thêm"))
        c_a.execute(f"SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV_AUDIT}")
        checks_a.append(("Cơ cấu theo Trình độ học vấn", c_a.fetchone()['t'], "không filter thêm"))
        c_a.execute(f"""
            SELECT COUNT(*) t FROM nhan_vien
            WHERE {DK_CHUAN_NV_AUDIT} AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != ''
        """)
        checks_a.append(("Cơ cấu theo Chức danh (SQL đầy đủ)", c_a.fetchone()['t'], "filter thêm: chuc_danh_nghe IS NOT NULL AND != ''"))
        c_a.execute(f"SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV_AUDIT} AND ngay_sinh IS NOT NULL")
        checks_a.append(("Cơ cấu theo Độ tuổi", c_a.fetchone()['t'], "filter thêm: ngay_sinh IS NOT NULL (chủ ý)"))
        c_a.execute(f"""
            SELECT COUNT(*) t FROM nhan_vien
            WHERE ngay_vao_lam >= (CURRENT_DATE - INTERVAL '6 months') AND {DK_CHUAN_NV_AUDIT}
        """)
        checks_a.append(("Xu hướng tuyển dụng 6 tháng", c_a.fetchone()['t'], "filter thêm: chỉ 6 tháng gần nhất (chủ ý)"))

        rows_out_a = []
        for ten_bd, tong_bd, ghi_chu in checks_a:
            lech = tong_bd - TONG_CHUAN
            if lech == 0:
                trang_thai = "✅ Khớp"
            elif "chủ ý" in ghi_chu:
                trang_thai = f"ℹ️ Lệch {lech:+d} — chủ ý (xem ghi chú)"
            else:
                trang_thai = f"🚨 LỆCH {lech:+d} — CẦN KIỂM TRA"
            rows_out_a.append({
                "Biểu đồ": ten_bd, "Tổng": tong_bd, "Tổng chuẩn": TONG_CHUAN,
                "Chênh lệch": lech, "Trạng thái": trang_thai, "Ghi chú": ghi_chu,
            })
        st.dataframe(pd.DataFrame(rows_out_a), width='stretch', hide_index=True)
        st.divider()

        # 2) Chi tiết toàn bộ chức danh — lộ ra những chức danh sẽ bị LIMIT cắt mất
        st.subheader("2️⃣ Chi tiết toàn bộ Chức danh (không LIMIT)")
        c_a.execute(f"""
            SELECT chuc_danh_nghe, COUNT(*) as so_luong
            FROM nhan_vien WHERE {DK_CHUAN_NV_AUDIT}
            AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != ''
            GROUP BY chuc_danh_nghe
            ORDER BY so_luong DESC
        """)
        df_full_role_a = pd.DataFrame(c_a.fetchall())
        if not df_full_role_a.empty:
            df_full_role_a.index = range(1, len(df_full_role_a) + 1)
            df_full_role_a['Sẽ bị LIMIT 10 cắt mất?'] = [
                "🚨 CÓ" if i > 10 else "✅ Không" for i in df_full_role_a.index
            ]
            st.dataframe(df_full_role_a, width='stretch')
            so_bi_cat = (df_full_role_a.index > 10).sum()
            if so_bi_cat > 0:
                nguoi_bi_cat = df_full_role_a[df_full_role_a.index > 10]['so_luong'].sum()
                st.warning(f"⚠️ {so_bi_cat} chức danh ({nguoi_bi_cat} người) sẽ bị mất nếu SQL có LIMIT 10.")
            else:
                st.success("✅ Tổng số chức danh ≤ 10, không có nguy cơ bị LIMIT cắt mất dữ liệu.")
        else:
            st.info("Không có dữ liệu chức danh.")
        st.divider()

        # 3) Nhân viên bị loại khỏi Tổng chuẩn kèm lý do
        st.subheader("3️⃣ Nhân viên KHÔNG nằm trong Tổng chuẩn (kèm lý do)")
        c_a.execute(f"""
            SELECT ma_nv, ho_ten, trang_thai, so_hdld,
                CASE
                    WHEN trang_thai NOT IN ('DANG_LAM','THU_VIEC') THEN 'Trạng thái không phải Đang làm/Thử việc: ' || COALESCE(trang_thai,'(rỗng)')
                    WHEN so_hdld IS NULL OR so_hdld = '' THEN 'Chưa có số HĐLĐ (hồ sơ chưa hoàn thiện)'
                    ELSE 'Không rõ'
                END as ly_do
            FROM nhan_vien
            WHERE NOT ({DK_CHUAN_NV_AUDIT})
            ORDER BY trang_thai, ho_ten
        """)
        loai_tru_a = c_a.fetchall()
        if loai_tru_a:
            st.dataframe(pd.DataFrame(loai_tru_a), width='stretch', hide_index=True)
            st.caption(f"Tổng cộng {len(loai_tru_a)} nhân viên bị loại khỏi Tổng chuẩn vì lý do trên.")
        else:
            st.info("Không có nhân viên nào bị loại — mọi bản ghi đều đạt Tổng chuẩn.")
        st.divider()

        # 4) Cảnh báo NULL/rỗng ở các trường dùng để group biểu đồ
        st.subheader("4️⃣ Kiểm tra NULL/rỗng ở các trường dùng để nhóm biểu đồ")
        truong_can_kiem_tra_a = {
            "gioi_tinh": "Giới tính", "trinh_do": "Trình độ",
            "phong_ban_lam_viec": "Phòng ban", "ngay_sinh": "Ngày sinh",
            "chuc_danh_nghe": "Chức danh",
        }
        rows_null_a = []
        chi_tiet_null_a = {}
        for cot, nhan in truong_can_kiem_tra_a.items():
            c_a.execute(f"""
                SELECT ma_nv, ho_ten, trang_thai
                FROM nhan_vien
                WHERE {DK_CHUAN_NV_AUDIT} AND ({cot} IS NULL OR {cot}::text = '')
                ORDER BY ho_ten
            """)
            ds_nv_thieu = c_a.fetchall()
            so_luong_null = len(ds_nv_thieu)
            chi_tiet_null_a[nhan] = ds_nv_thieu
            rows_null_a.append({
                "Trường": nhan, "Số nhân viên NULL/rỗng": so_luong_null,
                "Trạng thái": "✅ Không có" if so_luong_null == 0 else f"⚠️ Có {so_luong_null} người thiếu dữ liệu"
            })
        st.dataframe(pd.DataFrame(rows_null_a), width='stretch', hide_index=True)

        # Liệt kê CỤ THỂ tên/mã nhân viên cho từng trường bị thiếu, để không phải tự viết SQL
        for nhan, ds_nv_thieu in chi_tiet_null_a.items():
            if ds_nv_thieu:
                with st.expander(f"👤 Danh sách nhân viên thiếu '{nhan}' ({len(ds_nv_thieu)} người)"):
                    st.dataframe(pd.DataFrame(ds_nv_thieu), width='stretch', hide_index=True)
        db_a.close()

elif menu == "📘 Hướng dẫn sử dụng":
    st.title("📘 Hướng dẫn sử dụng HRM-Port")
    st.caption("Tổng quan các chức năng chính của hệ thống — dành cho người dùng mới.")

    st.markdown("""
### 📊 Dashboard
Bức tranh tổng quan về nhân sự: tổng số nhân viên, cơ cấu theo phòng ban, độ tuổi, giới tính, xu hướng tuyển dụng...
giúp Ban điều hành nắm tình hình chỉ trong vài giây, không cần chờ báo cáo tổng hợp thủ công.

### ✅ Nhân viên
Quản lý toàn bộ hồ sơ nhân viên: thêm mới, cập nhật thông tin, tra cứu nhanh, in Hợp đồng lao động/Hợp đồng thử việc,
ra các Quyết định nhân sự (bổ nhiệm, điều chuyển, chấm dứt HĐLĐ...).

🎉 **Đặc biệt: Gửi lời chúc sinh nhật tự động** — hệ thống tự nhắc và hỗ trợ gửi lời chúc mừng sinh nhật đến từng
CBCNV. Đây là một chi tiết nhỏ nhưng có sức nặng lớn: nó giúp gắn kết giữa Ban điều hành với người lao động,
khiến nhân viên cảm thấy được quan tâm như một cá nhân chứ không chỉ là một con số trên bảng lương — góp phần
xây dựng văn hoá doanh nghiệp gắn bó, nhân văn.

### 📋 BHXH / 📋 Báo cáo định kỳ
Theo dõi tình hình đóng BHXH, tự tạo báo cáo tăng/giảm D02-LT, dự toán số tiền phải đóng theo kỳ — giảm tối đa
thao tác thủ công so với việc tự tổng hợp trên Excel.

### 🕒 Chấm công / 💰 Tính thu nhập
Quản lý chấm công theo ca, tự động tính lương, phụ cấp, các khoản khấu trừ theo đúng quy định hiện hành.

### 📄 Quản lý Công văn & HĐ kinh tế
Lưu trữ, tra cứu công văn đến/đi và hợp đồng kinh tế tập trung — tránh thất lạc, dễ dàng tìm lại khi cần đối chiếu.

### ⏰ Báo cáo tự động & Nhắc hạn — không lo bị "miss" deadline
Hệ thống có các loại **báo cáo tự động** (tăng/giảm nhân sự, BHXH, hợp đồng...) giúp tiết kiệm thời gian tổng hợp
thủ công, đồng thời có **thông báo nhắc nhở các mốc quan trọng sắp đến hạn** (hết hạn HĐLĐ, hết hạn thử việc...),
giúp bộ phận Nhân sự chủ động xử lý trước hạn, tránh bỏ sót ảnh hưởng đến quyền lợi người lao động và rủi ro pháp lý
cho doanh nghiệp.

### 📊 Menu Báo cáo
Nơi tập trung liệt kê và chạy tất cả các loại báo cáo nhân sự sẵn có, xuất trực tiếp ra file để gửi cho Ban giám đốc,
cơ quan BHXH, hoặc lưu trữ nội bộ mà không cần thao tác qua nhiều màn hình.

### 🤖 Chatbot Giải đáp
Trợ lý AI trả lời nhanh các câu hỏi về BHXH, BHYT, thuế TNCN, thai sản, thất nghiệp, hợp đồng lao động — có trích dẫn
điều luật cụ thể, giúp CBCNV và bộ phận Nhân sự tự tra cứu quyền lợi mà không cần chờ hỏi trực tiếp.

### 💬 Chat nội bộ
Kênh trao đổi nội bộ ngay trong app — không cần chuyển qua ứng dụng nhắn tin khác.
""")

    st.info("💡 Có thắc mắc trong quá trình sử dụng, hãy dùng ngay mục **🤖 Chatbot Giải đáp** hoặc liên hệ bộ phận Nhân sự / IT để được hỗ trợ.")

st.sidebar.divider()
st.sidebar.caption("© 2026 HRM Master | © copyright: Mr.Tuyen - 0961778150")


#===== Hàm xử lý chính ===== 
def main():
    """Giữ tương thích với `if __name__ == '__main__': main()` ở cuối file.
    Landing Page đã bị bỏ (mỗi tenant có domain riêng, vào thẳng màn hình đăng nhập
    ở luồng phía trên) — hàm này không còn logic để chạy, chỉ giữ lại cho an toàn."""
    pass

def reset_ui_and_cache():
    """Reset toàn bộ cache và session state để refresh UI"""
    st.cache_data.clear()
    st.cache_resource.clear()
    
    # Giữ lại các session state quan trọng
    keep_keys = ['logged_in', 'role', 'username', 'language', 'show_hrm']
    for key in list(st.session_state.keys()):
        if key not in keep_keys:
            del st.session_state[key]
    
    st.rerun()

# Khởi tạo bảng chat khi app chạy
chat_utils.init_chat_tables()

# Chạy ứng dụng
if __name__ == "__main__":
    main()