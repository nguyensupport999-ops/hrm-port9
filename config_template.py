# config_template.py - COPY THIS FILE TO config.py AND FILL IN YOUR VALUES
# DO NOT COMMIT config.py TO GITHUB!

DB_CONFIG = {
    "host": "localhost",
    "port": 3306,
    "user": "your_db_user",
    "password": "your_db_password",
    "database": "your_db_name",
}

COMPANY_CONFIG = {
    "ten_cong_ty": "TÊN CÔNG TY CỦA BẠN",
    "dai_dien": "Tên người đại diện",
    "chuc_vu": "Chức vụ",
    "ma_so_thue": "Mã số thuế",
    "dien_thoai_cty": "Số điện thoại",
    "ma_don_vi_BHXH": "................",
    "ma_vung_luong": "....",
    "dia_chi": "Địa chỉ công ty",
    "loi_nhan_zalo": "Vui lòng kiểm tra và phản hồi nếu có sai sót...",
    "zalo_group_link": "https://zalo.me/g/...",
    "zalo_group_name": "Tên nhóm Zalo"
}

BHXH_CONFIG = {
    "ma_don_vi": "MÃ ĐƠN VỊ BHXH",
    "ty_le_NLD": 8.0,
    "ty_le_NSDLD": 17.5,
    "ty_le_BHTN_NLD": 1.0,
    "ty_le_BHTN_NSDLD": 1.0,
    "ty_le_BHYT_NLD": 1.5,
    "ty_le_BHYT_NSDLD": 3.0,
}

EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "email": "your_email@gmail.com",
    "password": "your_app_password",
    "nguoi_nhan": "recipient_email@gmail.com"
}

TELEGRAM_CONFIG = {
    "bot_token": "YOUR_BOT_TOKEN",
    "chat_id": "YOUR_CHAT_ID"
}

USERS = {
    "admin": {
        "password": "CHANGE_ME",
        "role": "admin"
    },
    "user": {
        "password": "CHANGE_ME",
        "role": "viewer"
    }
}