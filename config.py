DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "hrm_port",
}

COMPANY_CONFIG = {
    "ten_cong_ty": "CÔNG TY CỔ PHẦN CẢNG HÒN LA",
    "dai_dien": "Nguyễn Đình Thi ",
    "chuc_vu": "Tổng giám đốc",
    "ma_so_thue": "0108872052",
    "dien_thoai_cty": "0966369828",
    "ma_don_vi_BHXH": "................",
    "ma_vung_luong" : "....",
    "dia_chi": "Cảng tổng hợp Quốc tế Hòn La, X. Phú Trạch, T. Quảng Trị, Việt Nam",
    "loi_nhan_zalo": "Vui lòng kiểm tra và phản hồi nếu có sai sót trong ngày. Mọi thắc mắc liên hệ HCNS. Xin cảm ơn!",
    "zalo_group_link": "https://zalo.me/g/utwgzl694",  
    "zalo_group_name": "NLĐ Cảng QT Hòn La"  
}

#mst 0108872052; sdt: 0966369828; Nguyễn ĐÌnh Thi  #mqsb pgik ogxk xpdc

BHXH_CONFIG = {
    "ma_don_vi": "AB1234CD56",  # Mã đơn vị BHXH của công ty
    "ty_le_NLD": 8.0,    # Tỷ lệ đóng NLĐ (%)
    "ty_le_NSDLD": 17.5, # Tỷ lệ đóng NSDLĐ (%)
    "ty_le_BHTN_NLD": 1.0,
    "ty_le_BHTN_NSDLD": 1.0,
    "ty_le_BHYT_NLD": 1.5,
    "ty_le_BHYT_NSDLD": 3.0,
}

EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "email": "nguyen.support999@gmail.com",           # Email Gmail của bạn
    "password": "mqsb pgik ogxk xpdc",       # ← Dán 16 ký tự vừa copy
    "nguoi_nhan": "duhocanphuloc@gmail.com"     # Email nhận báo cáo
}

# ===== CẤU HÌNH TELEGRAM =====
TELEGRAM_CONFIG = {
    "bot_token": "123456:ABC-DEF1234ghIkl",  # ← THAY bằng token từ @BotFather
    "chat_id": "123456789"                   # ← THAY bằng chat_id của bạn
}
