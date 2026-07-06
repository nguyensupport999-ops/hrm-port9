# -*- coding: utf-8 -*-
"""
control_plane.py
=================
Module quản lý "Control Plane" cho mô hình SaaS đa khách hàng (multi-tenant).

Ý TƯỞNG:
- Có 1 Supabase project RIÊNG (do đội ngũ vận hành app sở hữu, KHÔNG phải của khách hàng)
  chỉ chứa duy nhất 1 bảng `tenants`. Bảng này KHÔNG chứa bất kỳ dữ liệu nhân sự nào,
  chỉ chứa thông tin để "trỏ" App tới đúng Supabase (Postgres DB + Storage) của từng khách hàng.
- Mật khẩu DB và Supabase Key của khách hàng được MÃ HOÁ (Fernet/AES) trước khi lưu vào
  bảng `tenants`, khoá giải mã (FERNET_KEY) chỉ nằm trong st.secrets của App, không lưu
  trong git/code. Nhờ vậy dù ai đó xem được bảng `tenants` cũng không đọc được mật khẩu thật.
- Khi 1 khách hàng đăng nhập, họ chỉ cần nhập "Mã công ty" (vd: CHL) -> App tra bảng
  `tenants` -> giải mã -> kết nối đúng Postgres + Storage của khách đó cho toàn bộ phiên làm việc.

CÁCH THIẾT LẬP (làm 1 lần khi triển khai hệ thống):
1. Tạo 1 Supabase project mới, đặt tên gợi ý: "hrm-saas-control-plane".
2. Chạy đoạn SQL trong SCHEMA_SQL (bên dưới) trên project đó để tạo bảng `tenants`.
3. Sinh 1 FERNET_KEY (chạy: python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())")
4. Khai báo trong st.secrets (Streamlit Cloud > Settings > Secrets) của App:

    [control_plane]
    host = "xxxx.supabase.co"   # host Postgres của project control-plane
    port = 5432
    user = "postgres"
    password = "..."            # mật khẩu Postgres của project control-plane
    database = "postgres"
    fernet_key = "..."          # khoá sinh ở bước 3

    [super_admin]
    username = "your_admin"
    password = "mat_khau_manh_cua_ban"

Từ đó, mỗi khi thêm khách hàng mới, chỉ cần vào trang "⚙️ Quản trị hệ thống" (chỉ super-admin
thấy được) trong App để thêm 1 dòng tenant mới — KHÔNG cần sửa code, KHÔNG cần deploy lại.
"""

import os
import psycopg2
import psycopg2.extras
import streamlit as st
from cryptography.fernet import Fernet


SCHEMA_SQL = """
-- Chạy đoạn này trên Supabase project dùng làm CONTROL PLANE (chỉ 1 lần)
CREATE TABLE IF NOT EXISTS tenants (
    id              SERIAL PRIMARY KEY,
    ma_cty          TEXT UNIQUE NOT NULL,      -- Mã công ty, vd: 'CHL'
    ten_cty         TEXT NOT NULL,             -- Tên đầy đủ, vd: 'Công ty Cổ phần Cảng Hòn La'
    logo_url        TEXT,                      -- Link logo (tuỳ chọn)
    db_host         TEXT NOT NULL,
    db_port         TEXT NOT NULL DEFAULT '5432',
    db_user         TEXT NOT NULL,
    db_password_enc TEXT NOT NULL,             -- mật khẩu Postgres đã mã hoá
    db_name         TEXT NOT NULL DEFAULT 'postgres',
    supabase_url    TEXT NOT NULL,             -- dùng cho Storage (ảnh NV, hồ sơ, file chat)
    supabase_key_enc TEXT NOT NULL,            -- Supabase API key đã mã hoá
    goi_dich_vu     TEXT DEFAULT 'standard',
    trang_thai      TEXT DEFAULT 'active',     -- active / suspended
    created_at      TIMESTAMP DEFAULT NOW()
);
"""


# ---------- Kết nối tới Control Plane ----------
def get_control_plane_connection():
    """Kết nối Postgres tới project Control Plane (KHÁC với DB của từng khách hàng)."""
    cfg = st.secrets["control_plane"]
    return psycopg2.connect(
        host=cfg["host"],
        port=cfg.get("port", 5432),
        user=cfg["user"],
        password=cfg["password"],
        database=cfg.get("database", "postgres"),
    )


def _get_fernet() -> Fernet:
    key = st.secrets["control_plane"]["fernet_key"]
    if isinstance(key, str):
        key = key.encode()
    return Fernet(key)


def encrypt_text(plain_text: str) -> str:
    if not plain_text:
        return ""
    return _get_fernet().encrypt(plain_text.encode()).decode()


def decrypt_text(cipher_text: str) -> str:
    if not cipher_text:
        return ""
    return _get_fernet().decrypt(cipher_text.encode()).decode()


# ---------- Tra cứu tenant ----------
@st.cache_data(ttl=60, show_spinner=False)
def get_tenant_by_code(ma_cty: str):
    """Trả về dict thông tin kết nối (đã giải mã) của 1 khách hàng theo mã công ty.
    Cache 60 giây để giảm số lần gọi Control Plane khi người dùng thao tác liên tục."""
    if not ma_cty:
        return None
    conn = get_control_plane_connection()
    try:
        c = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT * FROM tenants WHERE UPPER(ma_cty) = UPPER(%s)", (ma_cty.strip(),))
        row = c.fetchone()
    finally:
        conn.close()

    if not row:
        return None
    if row["trang_thai"] != "active":
        return {"error": "SUSPENDED", "ten_cty": row["ten_cty"]}

    return {
        "ma_cty": row["ma_cty"],
        "ten_cty": row["ten_cty"],
        "logo_url": row["logo_url"],
        "db_host": row["db_host"],
        "db_port": row["db_port"],
        "db_user": row["db_user"],
        "db_password": decrypt_text(row["db_password_enc"]),
        "db_name": row["db_name"],
        "supabase_url": row["supabase_url"],
        "supabase_key": decrypt_text(row["supabase_key_enc"]),
        "goi_dich_vu": row["goi_dich_vu"],
    }


def list_tenants():
    """Danh sách tenant cho trang Quản trị hệ thống (KHÔNG trả về mật khẩu/khoá đã giải mã)."""
    conn = get_control_plane_connection()
    try:
        c = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""SELECT id, ma_cty, ten_cty, logo_url, db_host, supabase_url,
                            goi_dich_vu, trang_thai, created_at
                     FROM tenants ORDER BY created_at DESC""")
        return c.fetchall()
    finally:
        conn.close()


def add_tenant(ma_cty, ten_cty, db_host, db_port, db_user, db_password,
               db_name, supabase_url, supabase_key, logo_url=None, goi_dich_vu="standard"):
    """Thêm khách hàng mới. Mật khẩu/Key được mã hoá trước khi lưu."""
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        c.execute("""
            INSERT INTO tenants (ma_cty, ten_cty, logo_url, db_host, db_port, db_user,
                                  db_password_enc, db_name, supabase_url, supabase_key_enc, goi_dich_vu)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (ma_cty.strip().upper(), ten_cty, logo_url, db_host, str(db_port), db_user,
              encrypt_text(db_password), db_name, supabase_url, encrypt_text(supabase_key), goi_dich_vu))
        conn.commit()
    finally:
        conn.close()
    get_tenant_by_code.clear()  # xoá cache để đọc lại ngay


def update_tenant_status(ma_cty, trang_thai):
    """Bật/tắt (active / suspended) 1 khách hàng — dùng khi khách ngừng hợp đồng."""
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE tenants SET trang_thai=%s WHERE UPPER(ma_cty)=UPPER(%s)", (trang_thai, ma_cty))
        conn.commit()
    finally:
        conn.close()
    get_tenant_by_code.clear()


def delete_tenant(ma_cty):
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        c.execute("DELETE FROM tenants WHERE UPPER(ma_cty)=UPPER(%s)", (ma_cty,))
        conn.commit()
    finally:
        conn.close()
    get_tenant_by_code.clear()


def check_super_admin(username, password):
    """Xác thực tài khoản quản trị hệ thống (chỉ đội vận hành app dùng để thêm/sửa khách hàng).
    Hoàn toàn tách biệt với tài khoản của nhân viên trong từng công ty khách hàng."""
    try:
        sa = st.secrets["super_admin"]
        return username == sa["username"] and password == sa["password"]
    except Exception:
        return False
