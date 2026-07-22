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
    db_password_enc TEXT NOT NULL,             -- mật khẩu Postgres đã hoá
    db_name         TEXT NOT NULL DEFAULT 'postgres',
    supabase_url    TEXT NOT NULL,             -- dùng cho Storage (ảnh NV, hồ sơ, file chat)
    supabase_key_enc TEXT NOT NULL,            -- Supabase API key đã mã hoá
    goi_dich_vu     TEXT DEFAULT 'standard',
    trang_thai      TEXT DEFAULT 'active',     -- active / suspended
    
    -- Các cột phục vụ cấu hình động (Branding & Business Logic)
    dai_dien         TEXT,
    chuc_vu          TEXT,
    ma_so_thue       TEXT,
    dien_thoai_cty   TEXT,
    ma_don_vi_bhxh   TEXT,
    ma_vung_luong    TEXT,
    dia_chi          TEXT,
    loi_nhan_zalo    TEXT,
    zalo_group_link  TEXT,
    zalo_group_name  TEXT,

    -- Ngôn ngữ giao diện của tenant: 'VI' (chỉ Việt, mặc định) / 'VI_EN' / 'VI_ZH' / 'VI_KO'.
    -- Tiếng Việt luôn là ngôn ngữ chính; ngôn ngữ phụ (nếu có) chỉ hiển thị thêm trong
    -- ngoặc, cỡ chữ nhỏ hơn, phục vụ khách FDI. Xem module i18n.py.
    ngon_ngu        TEXT DEFAULT 'VI',

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


def ensure_control_plane_schema():
    """Đảm bảo bảng tenants có đầy đủ các cột phục vụ Branding và cấu hình động."""
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        # Chạy ALTER TABLE để thêm các cột nếu chưa tồn tại
        columns_to_add = [
            ("dai_dien", "TEXT"),
            ("chuc_vu", "TEXT"),
            ("ma_so_thue", "TEXT"),
            ("dien_thoai_cty", "TEXT"),
            ("ma_don_vi_bhxh", "TEXT"),
            ("ma_vung_luong", "TEXT"),
            ("dia_chi", "TEXT"),
            ("loi_nhan_zalo", "TEXT"),
            ("zalo_group_link", "TEXT"),
            ("zalo_group_name", "TEXT"),
            ("ngon_ngu", "TEXT DEFAULT 'VI'"),
        ]
        for col_name, col_type in columns_to_add:
            c.execute(f"ALTER TABLE tenants ADD COLUMN IF NOT EXISTS {col_name} {col_type}")
        conn.commit()
    except Exception as e:
        print(f"Error ensuring control plane schema: {e}")
    finally:
        conn.close()


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
    ensure_control_plane_schema()
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
        "ngon_ngu": row.get("ngon_ngu") or "VI",

        # Metadata cấu hình động
        "dai_dien": row.get("dai_dien") or "",
        "chuc_vu": row.get("chuc_vu") or "",
        "ma_so_thue": row.get("ma_so_thue") or "",
        "dien_thoai_cty": row.get("dien_thoai_cty") or "",
        "ma_don_vi_BHXH": row.get("ma_don_vi_bhxh") or "",
        "ma_vung_luong": row.get("ma_vung_luong") or "",
        "dia_chi": row.get("dia_chi") or "",
        "loi_nhan_zalo": row.get("loi_nhan_zalo") or "",
        "zalo_group_link": row.get("zalo_group_link") or "",
        "zalo_group_name": row.get("zalo_group_name") or "",
    }


def list_tenants():
    """Danh sách tenant cho trang Quản trị hệ thống (KHÔNG trả về mật khẩu/khoá đã giải mã)."""
    ensure_control_plane_schema()
    conn = get_control_plane_connection()
    try:
        c = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""SELECT id, ma_cty, ten_cty, logo_url, db_host, supabase_url,
                            goi_dich_vu, trang_thai, created_at,
                            dai_dien, chuc_vu, ma_so_thue, dien_thoai_cty,
                            ma_don_vi_bhxh as "ma_don_vi_BHXH", ma_vung_luong, dia_chi,
                            loi_nhan_zalo, zalo_group_link, zalo_group_name,
                            COALESCE(ngon_ngu, 'VI') as ngon_ngu
                     FROM tenants ORDER BY created_at DESC""")
        return c.fetchall()
    finally:
        conn.close()


def add_tenant(ma_cty, ten_cty, db_host, db_port, db_user, db_password,
               db_name, supabase_url, supabase_key, logo_url=None, goi_dich_vu="standard",
               dai_dien=None, chuc_vu=None, ma_so_thue=None, dien_thoai_cty=None,
               ma_don_vi_BHXH=None, ma_vung_luong=None, dia_chi=None, loi_nhan_zalo=None,
               zalo_group_link=None, zalo_group_name=None, migration_sql=None,
               ngon_ngu="VI"):
    """Thêm khách hàng mới. Mật khẩu/Key được mã hoá trước khi lưu.
    Nếu có migration_sql, tự động chạy script tạo bảng trên database của tenant mới."""
    ensure_control_plane_schema()
    
    # 1. Thêm tenant vào control plane
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        c.execute("""
            INSERT INTO tenants (
                ma_cty, ten_cty, logo_url, db_host, db_port, db_user,
                db_password_enc, db_name, supabase_url, supabase_key_enc, goi_dich_vu,
                dai_dien, chuc_vu, ma_so_thue, dien_thoai_cty, ma_don_vi_bhxh,
                ma_vung_luong, dia_chi, loi_nhan_zalo, zalo_group_link, zalo_group_name,
                ngon_ngu
            )
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            ma_cty.strip().upper(), ten_cty, logo_url, db_host, str(db_port), db_user,
            encrypt_text(db_password), db_name, supabase_url, encrypt_text(supabase_key), goi_dich_vu,
            dai_dien, chuc_vu, ma_so_thue, dien_thoai_cty, ma_don_vi_BHXH,
            ma_vung_luong, dia_chi, loi_nhan_zalo, zalo_group_link, zalo_group_name,
            (ngon_ngu or "VI").strip().upper()
        ))
        conn.commit()
    finally:
        conn.close()
        
    # 2. Chạy database migration trên database của tenant mới
    if migration_sql:
        tenant_conn = psycopg2.connect(
            host=db_host,
            port=str(db_port),
            user=db_user,
            password=db_password,
            database=db_name
        )
        try:
            tc = tenant_conn.cursor()
            tc.execute(migration_sql)
            tenant_conn.commit()
        except Exception as e:
            tenant_conn.rollback()
            # Xoá tenant vừa thêm ở control plane để tránh dữ liệu mồ côi nếu lỗi migration
            delete_tenant(ma_cty)
            raise e
        finally:
            tenant_conn.close()
            
    get_tenant_by_code.clear()  # xoá cache để đọc lại ngay


def update_tenant_language(ma_cty, ngon_ngu):
    """Đổi ngôn ngữ giao diện của 1 tenant đã tồn tại (VI / VI_EN / VI_ZH / VI_KO)."""
    ensure_control_plane_schema()
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE tenants SET ngon_ngu=%s WHERE UPPER(ma_cty)=UPPER(%s)",
                   ((ngon_ngu or "VI").strip().upper(), ma_cty))
        conn.commit()
    finally:
        conn.close()
    get_tenant_by_code.clear()


def update_tenant_logo(ma_cty, logo_url):
    """Cập nhật link logo (Storage public URL) cho 1 tenant đã tồn tại."""
    ensure_control_plane_schema()
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE tenants SET logo_url=%s WHERE UPPER(ma_cty)=UPPER(%s)",
                   (logo_url, ma_cty))
        conn.commit()
    finally:
        conn.close()
    get_tenant_by_code.clear()


def update_tenant_status(ma_cty, trang_thai):
    """Bật/tắt (active / suspended) 1 khách hàng — dùng khi khách ngừng hợp đồng."""
    ensure_control_plane_schema()
    conn = get_control_plane_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE tenants SET trang_thai=%s WHERE UPPER(ma_cty)=UPPER(%s)", (trang_thai, ma_cty))
        conn.commit()
    finally:
        conn.close()
    get_tenant_by_code.clear()


def delete_tenant(ma_cty):
    ensure_control_plane_schema()
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


# ============================================================================
# ĐA KHÁCH HÀNG QUA SUBDOMAIN (White-labeling) + CHẾ ĐỘ DEMO
# ============================================================================
#
# GHI CHÚ THIẾT KẾ VỀ DEMO MODE:
# Thay vì xây một "lớp giả lập SQL" chung để chặn/đánh tráo mọi câu lệnh
# SELECT/INSERT/UPDATE trong app.py (file có hơn 100 điểm gọi get_connection()
# với rất nhiều câu SQL Postgres thuần, đặc thù), giải pháp AN TOÀN và ÍT RỦI RO
# HỎNG HÓC hơn nhiều là: coi "DEMO" như MỘT TENANT BÌNH THƯỜNG trong bảng
# `tenants` (trỏ tới 1 Supabase/Postgres THẬT, dùng chung, đã được seed sẵn dữ
# liệu mẫu). Nhờ vậy toàn bộ ~7500 dòng SQL hiện có chạy đúng y hệt, không cần
# sửa từng câu truy vấn. Điểm khác biệt DUY NHẤT: mọi lệnh GHI (INSERT/UPDATE/
# DELETE/DROP/...) trên tenant DEMO sẽ bị chặn ở tầng cursor để bảo vệ dữ liệu
# dùng chung khỏi bị khách vãng lai chỉnh sửa/xoá.
#
# Cách kích hoạt Demo Mode thực tế:
#   1. Dùng đúng trang "Quản trị hệ thống" đã có sẵn trong app.py để add_tenant()
#      với ma_cty="DEMO", trỏ tới 1 Supabase project demo (miễn phí) đã được
#      seed dữ liệu mẫu (xem seed_demo_data.sql đi kèm).
#   2. Truy cập qua demo.kendu-ai.com (Option A) hoặc ?tenant=DEMO (Option B/
#      Streamlit Cloud) đều tự động nhận diện và nạp đúng tenant DEMO.


class _ReadOnlyGuardCursor:
    """Bọc 1 cursor psycopg2 thật, chặn các câu lệnh GHI để bảo vệ DB Demo dùng chung.
    Mọi câu SELECT vẫn chạy bình thường trên dữ liệu mẫu thật."""

    _BLOCKED_KEYWORDS = ("insert", "update", "delete", "drop", "truncate", "alter", "grant", "revoke")

    def __init__(self, real_cursor):
        self._cursor = real_cursor

    def execute(self, query, params=None):
        first_word = ""
        stripped = query.strip() if isinstance(query, str) else ""
        if stripped:
            first_word = stripped.split(None, 1)[0].lower()
        if first_word in self._BLOCKED_KEYWORDS:
            raise psycopg2.Error(
                "🔒 Đây là dữ liệu DEMO dùng chung cho mục đích trải nghiệm — "
                "không thể lưu/sửa/xoá. Vui lòng đăng ký dùng thử để sử dụng đầy đủ tính năng."
            )
        if params is not None:
            return self._cursor.execute(query, params)
        return self._cursor.execute(query)

    def executemany(self, query, params_list):
        raise psycopg2.Error(
            "🔒 Đây là dữ liệu DEMO dùng chung — không thể lưu/sửa/xoá hàng loạt."
        )

    def __getattr__(self, name):
        return getattr(self._cursor, name)

    def __iter__(self):
        return iter(self._cursor)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._cursor.close()


class _ReadOnlyGuardConnection:
    """Bọc 1 connection psycopg2 thật; chỉ cursor() trả về là cursor có chặn ghi."""

    def __init__(self, real_conn):
        self._conn = real_conn

    def cursor(self, *args, **kwargs):
        return _ReadOnlyGuardCursor(self._conn.cursor(*args, **kwargs))

    def __getattr__(self, name):
        return getattr(self._conn, name)


class DatabaseEngine:
    """Quản lý kết nối Postgres theo từng tenant (đa khách hàng).

    - tenant=None: chưa xác định công ty, get_connection() sẽ báo lỗi rõ ràng
      thay vì crash mơ hồ (buộc code gọi phải xử lý bước chọn công ty trước).
    - tenant thường: mở kết nối thật tới đúng Postgres của khách hàng đó.
    - tenant với ma_cty == 'DEMO': mở kết nối thật tới DB demo dùng chung,
      nhưng bọc cursor để chặn mọi lệnh ghi (xem _ReadOnlyGuardCursor).
    """

    def __init__(self, tenant: dict | None):
        self.tenant = tenant

    @property
    def is_demo(self) -> bool:
        return bool(self.tenant) and str(self.tenant.get("ma_cty", "")).upper() == "DEMO"

    def get_connection(self):
        if not self.tenant:
            raise RuntimeError(
                "Chưa xác định công ty (tenant) — không thể mở kết nối cơ sở dữ liệu. "
                "Hãy chọn công ty trước khi thao tác."
            )
        conn = psycopg2.connect(
            host=self.tenant["db_host"],
            port=str(self.tenant.get("db_port", "5432")),
            user=self.tenant["db_user"],
            password=self.tenant["db_password"],
            database=self.tenant.get("db_name", "postgres"),
        )
        if self.is_demo:
            return _ReadOnlyGuardConnection(conn)
        return conn


# ---------- Nhận diện tenant tự động qua Subdomain / query param ----------

# Các domain "gốc" (apex) — nếu hostname trùng khớp domain gốc (không có subdomain
# công ty), coi như đang ở Landing Page / cần chọn công ty thủ công như hiện tại.
_ROOT_DOMAINS = ("kendu-ai.com", "streamlit.app", "localhost", "127.0.0.1")


def _get_request_hostname() -> str:
    """Lấy hostname đầy đủ của request hiện tại, vd 'honla.kendu-ai.com'.
    Dùng st.context.headers (Streamlit >= 1.32) khi có; nếu môi trường không
    hỗ trợ (hoặc chạy sau proxy không forward đúng Host) trả về chuỗi rỗng."""
    try:
        host = st.context.headers.get("host", "") or ""
    except Exception:
        host = ""
    return host.split(":")[0].strip().lower()


def _extract_subdomain_code(hostname: str):
    """Tách mã công ty (ma_cty) từ subdomain. Trả về None nếu đang ở domain gốc
    (không có subdomain riêng) hoặc hostname không xác định được."""
    if not hostname:
        return None
    for root in _ROOT_DOMAINS:
        if hostname == root or hostname == f"www.{root}":
            return None
        suffix = f".{root}"
        if hostname.endswith(suffix):
            sub = hostname[: -len(suffix)]
            if sub and sub != "www":
                return sub.upper()
    return None


def resolve_tenant():
    """Tự động nhận diện tenant, theo thứ tự ưu tiên:
    1) st.secrets['tenant_code'] — dùng cho mô hình "1 app Streamlit / 1 khách hàng":
       mỗi khách có 1 deployment riêng trên Streamlit Cloud, Secret tenant_code khoá
       cứng app đó vào đúng 1 công ty, bỏ qua hẳn bước chọn công ty thủ công.
    2) Subdomain (honla.kendu-ai.com) hoặc query param ?tenant=HONLA — dùng cho app
       dùng chung nhiều tenant (như app "honla" hiện tại) hoặc phương án Iframe.

    Nếu KHÔNG xác định được tenant, hàm không làm gì cả — luồng chọn công ty thủ công
    đã có sẵn trong app.py (nhập Mã công ty ở sidebar) sẽ tự xử lý tiếp."""
    if st.session_state.get("tenant"):
        return

    # ---- Ưu tiên 1: Secret tenant_code (app riêng cho 1 khách) ----
    try:
        secret_tenant_code = st.secrets.get("tenant_code")
    except Exception:
        secret_tenant_code = None
    if secret_tenant_code:
        tenant = get_tenant_by_code(secret_tenant_code.strip())
        if tenant and tenant.get("error") != "SUSPENDED":
            st.session_state.tenant = tenant
            st.session_state.db_engine = DatabaseEngine(tenant)
            st.session_state["_tenant_locked"] = True  # app.py dùng cờ này để bỏ qua Landing Page
            return
        # Nếu Secret có nhưng mã sai/khoá tài khoản -> KHÔNG rơi về chọn công ty thủ công
        # (vì app này vốn được provision riêng cho 1 khách), mà báo lỗi rõ ràng.
        st.session_state["_tenant_locked_error"] = (
            "SUSPENDED" if tenant and tenant.get("error") == "SUSPENDED" else "NOT_FOUND"
        )
        return

    # ---- Ưu tiên 2: Subdomain / query param (app dùng chung nhiều tenant) ----
    ma_cty = None
    try:
        qp_tenant = st.query_params.get("tenant")
    except Exception:
        qp_tenant = None
    if qp_tenant:
        ma_cty = qp_tenant.strip().upper()
    else:
        ma_cty = _extract_subdomain_code(_get_request_hostname())

    if not ma_cty:
        return

    tenant = get_tenant_by_code(ma_cty)
    if not tenant or tenant.get("error") == "SUSPENDED":
        return

    st.session_state.tenant = tenant
    st.session_state.db_engine = DatabaseEngine(tenant)