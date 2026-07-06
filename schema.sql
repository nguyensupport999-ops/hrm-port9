-- ============================================================================
-- schema.sql — HRM-Port
-- ============================================================================
-- File này được app.py TỰ ĐỘNG CHẠY mỗi khi Super Admin "Thêm khách hàng mới"
-- ở trang ⚙️ Quản trị hệ thống (xem control_plane.add_tenant(migration_sql=...)).
-- Yêu cầu: đặt file này CÙNG THƯ MỤC với app.py trên server/Streamlit Cloud.
--
-- NGUỒN GỐC: được dựng lại (reverse-engineer) bằng cách quét toàn bộ câu
-- SELECT/INSERT/UPDATE trong app.py, vì repo hiện không có schema.sql gốc nào.
-- Đã cố gắng bao quát hết các cột đang được code sử dụng tại thời điểm viết,
-- nhưng vì suy ra từ cách dùng chứ không phải từ 1 schema gốc, RẤT NÊN kiểm
-- thử kỹ (tạo 1 tenant mới, thử đủ các chức năng) trước khi dùng cho khách
-- hàng thật.
--
-- AN TOÀN CHẠY LẠI: mọi CREATE TABLE đều có IF NOT EXISTS, và có thêm các
-- câu ALTER TABLE ... ADD COLUMN IF NOT EXISTS phía dưới mỗi bảng — nhờ vậy
-- có thể chạy file này lại trên 1 DB tenant ĐÃ TỒN TẠI (như DB "Cảng Hòn La"
-- hiện tại, đang thiếu cột mat_khau_hash/phai_doi_mat_khau) để tự bổ sung
-- cột còn thiếu MÀ KHÔNG mất dữ liệu hiện có.
-- ============================================================================


-- ============================================================================
-- 1. NHÂN VIÊN — bảng trung tâm: hồ sơ nhân sự + tài khoản đăng nhập
-- ============================================================================
CREATE TABLE IF NOT EXISTS nhan_vien (
    id                          SERIAL PRIMARY KEY,
    ho_ten                      TEXT NOT NULL,
    created_at                  TIMESTAMP DEFAULT NOW()
);

ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS stt                         INTEGER;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ma_nv                       VARCHAR(20);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS so_hdld                     VARCHAR(50);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS chuc_danh_nghe              TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ngay_sinh                   DATE;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS gioi_tinh                   VARCHAR(10);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS so_cccd                     VARCHAR(20);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ngay_cap_cccd               DATE;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS noi_cap_cccd                TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS nguyen_quan                 TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS thuong_tru                  TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS dien_thoai                  VARCHAR(20);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS email                       VARCHAR(100);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS email_lien_he               VARCHAR(100);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ho_so                       TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS luong_bao_hiem              NUMERIC(15,2);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ma_so_bhxh                  VARCHAR(20);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ngay_vao_lam                DATE;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS noi_lam_viec                TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS so_tai_khoan_nh             VARCHAR(30);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS chi_nhanh_nh                TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ngay_ky_hd                  DATE;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS loai_hop_dong               VARCHAR(50);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS nhom_bhxh                   VARCHAR(20);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS thang_bat_dau_bh            DATE;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS thang_ket_thuc_bh           DATE;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS trang_thai                  VARCHAR(20) DEFAULT 'THU_VIEC';
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS trang_thai_bhxh             VARCHAR(20) DEFAULT 'CHUA_DONG';
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phong_ban_lam_viec          TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ngay_ket_thuc               DATE;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS quoc_tich                   VARCHAR(50) DEFAULT 'Việt Nam';
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS dan_toc                     VARCHAR(50) DEFAULT 'Kinh';
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS he_so_luong                 NUMERIC(6,2);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phu_cap_chuc_vu             NUMERIC(15,2);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phu_cap_tnvk                NUMERIC(15,2);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phu_cap_tnn                 NUMERIC(15,2);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS muc_huong_bhyt              VARCHAR(10);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ty_le_dong                  NUMERIC(6,2);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS muc_tien_dong               NUMERIC(15,2);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phuong_thuc_dong            VARCHAR(50);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS tinh_nhan_hs                TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phuong_nhan_hs              TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS dia_chi_nhan_hs             TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS tinh_kcb                    TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS noi_dang_ky_kcb             TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS dang_ky_nhan_so             VARCHAR(10);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ten_don_vi_thu_huong        TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ngay_chinh_thuc             DATE;
-- Thông tin chủ hộ (dùng cho tờ khai BHXH hộ gia đình)
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ho_ten_chu_ho               TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS so_cccd_chu_ho              VARCHAR(20);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS tinh_thanh_pho_chu_ho       TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phuong_xa_chu_ho            TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS tinh_thanh_pho_thuong_tru   TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ma_tinh_thuong_tru          VARCHAR(10);
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phuong_xa_thuong_tru        TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS ma_phuong_xa_thuong_tru     VARCHAR(10);
-- Tài khoản đăng nhập (QUAN TRỌNG — đây chính là 2 cột bị thiếu gây lỗi đăng nhập)
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS vai_tro                     VARCHAR(20) DEFAULT 'nhan_vien';
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS mat_khau_hash               TEXT;
ALTER TABLE nhan_vien ADD COLUMN IF NOT EXISTS phai_doi_mat_khau           BOOLEAN DEFAULT FALSE;

-- Index/ràng buộc duy nhất (bọc trong DO block để không lỗi nếu đã tồn tại
-- hoặc nếu dữ liệu cũ có trùng lặp — khi đó chỉ bỏ qua, không chặn migration)
DO $$ BEGIN
    ALTER TABLE nhan_vien ADD CONSTRAINT nhan_vien_ma_nv_key UNIQUE (ma_nv);
EXCEPTION WHEN duplicate_table OR duplicate_object OR unique_violation THEN NULL;
END $$;
DO $$ BEGIN
    ALTER TABLE nhan_vien ADD CONSTRAINT nhan_vien_dien_thoai_key UNIQUE (dien_thoai);
EXCEPTION WHEN duplicate_table OR duplicate_object OR unique_violation THEN NULL;
END $$;


-- ============================================================================
-- 2. ỨNG VIÊN (tuyển dụng)
-- ============================================================================
CREATE TABLE IF NOT EXISTS ung_vien (
    id                  SERIAL PRIMARY KEY,
    ho_ten              TEXT NOT NULL,
    created_at          TIMESTAMP DEFAULT NOW()
);
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS ma_uv               VARCHAR(20);
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS vi_tri_du_tuyen     TEXT;
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS dien_thoai          VARCHAR(20);
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS ngay_sinh           DATE;
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS gioi_tinh           VARCHAR(10);
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS ngay_vao_lam        DATE;
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS luong_bao_hiem      NUMERIC(15,2);
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS trang_thai          VARCHAR(20) DEFAULT 'CHO_DUYET';  -- CHO_DUYET / TU_CHOI / DA_NHAN_VIEC
ALTER TABLE ung_vien ADD COLUMN IF NOT EXISTS ma_nv               VARCHAR(20);  -- gán khi ứng viên -> nhân viên chính thức


-- ============================================================================
-- 3. DANH MỤC VỊ TRÍ CÔNG TÁC (chức danh)
-- ============================================================================
CREATE TABLE IF NOT EXISTS vi_tri_cong_tac (
    id          SERIAL PRIMARY KEY,
    ten_vi_tri  TEXT NOT NULL,
    ghi_chu     TEXT
);


-- ============================================================================
-- 4. LỊCH SỬ CÔNG TÁC (lịch sử chức danh/phòng ban/hợp đồng theo thời gian)
-- ============================================================================
CREATE TABLE IF NOT EXISTS lich_su_cong_tac (
    id              SERIAL PRIMARY KEY,
    nhan_vien_id    INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE,
    created_at      TIMESTAMP DEFAULT NOW()
);
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS tu_ngay         DATE;
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS den_ngay        DATE;
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS chuc_danh       TEXT;
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS phong_ban       TEXT;
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS noi_lam_viec    TEXT;
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS loai_hop_dong   VARCHAR(50);
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS he_so_luong     NUMERIC(6,2);
ALTER TABLE lich_su_cong_tac ADD COLUMN IF NOT EXISTS so_hop_dong     VARCHAR(50);


-- ============================================================================
-- 5. QUYẾT ĐỊNH NHÂN SỰ (thử việc/chính thức/điều chuyển/bổ nhiệm/tăng lương/nghỉ việc)
-- ============================================================================
CREATE TABLE IF NOT EXISTS quyet_dinh_nhan_su (
    id                  SERIAL PRIMARY KEY,
    nhan_vien_id        INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE,
    created_at          TIMESTAMP DEFAULT NOW()
);
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS loai_quyet_dinh     VARCHAR(30);  -- THU_VIEC / CHINH_THUC / DIEU_CHUYEN / BO_NHIEM / TANG_LUONG / NGHI_VIEC
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS ngay_quyet_dinh     DATE;
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS ngay_hieu_luc       DATE;
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS noi_dung            TEXT;
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS so_quyet_dinh       VARCHAR(50);
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS loai_hop_dong_cu    VARCHAR(50);
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS loai_hop_dong_moi   VARCHAR(50);
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS he_so_luong_cu      NUMERIC(6,2);
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS he_so_luong_moi     NUMERIC(6,2);
ALTER TABLE quyet_dinh_nhan_su ADD COLUMN IF NOT EXISTS so_hd_cu            VARCHAR(50);


-- ============================================================================
-- 6. PHỤ LỤC GIA ĐÌNH (thành viên hộ gia đình — tờ khai BHXH)
-- ============================================================================
CREATE TABLE IF NOT EXISTS phu_luc_gia_dinh (
    id                  SERIAL PRIMARY KEY,
    nhan_vien_id        INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE
);
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS ho_ten              TEXT;
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS ngay_sinh           DATE;
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS gioi_tinh           VARCHAR(10);
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS quoc_tich           VARCHAR(50);
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS dan_toc             VARCHAR(50);
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS quan_he_voi_chu_ho  VARCHAR(50);
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS tinh_thanh_pho      TEXT;
ALTER TABLE phu_luc_gia_dinh ADD COLUMN IF NOT EXISTS phuong_xa           TEXT;


-- ============================================================================
-- 7. HỒ SƠ NHÂN VIÊN (file đính kèm upload lên Supabase Storage)
-- ============================================================================
CREATE TABLE IF NOT EXISTS ho_so_nhan_vien (
    id                  SERIAL PRIMARY KEY,
    nhan_vien_id        INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE
);
ALTER TABLE ho_so_nhan_vien ADD COLUMN IF NOT EXISTS loai_ho_so      VARCHAR(50);
ALTER TABLE ho_so_nhan_vien ADD COLUMN IF NOT EXISTS ten_file        TEXT;
ALTER TABLE ho_so_nhan_vien ADD COLUMN IF NOT EXISTS duong_dan_file  TEXT;
ALTER TABLE ho_so_nhan_vien ADD COLUMN IF NOT EXISTS ngay_upload     DATE DEFAULT CURRENT_DATE;


-- ============================================================================
-- 8. LỊCH SỬ GỬI LỜI CHÚC (sinh nhật qua Zalo/Email...)
-- ============================================================================
CREATE TABLE IF NOT EXISTS lich_su_gui_loi_chuc (
    id                  SERIAL PRIMARY KEY,
    nhan_vien_id        INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE,
    ngay_gui            TIMESTAMP DEFAULT NOW()
);
ALTER TABLE lich_su_gui_loi_chuc ADD COLUMN IF NOT EXISTS loai_chuc   VARCHAR(30);  -- SINH_NHAT, ...
ALTER TABLE lich_su_gui_loi_chuc ADD COLUMN IF NOT EXISTS noi_dung    TEXT;
ALTER TABLE lich_su_gui_loi_chuc ADD COLUMN IF NOT EXISTS kenh_gui    VARCHAR(20);  -- ZALO / EMAIL / ...
ALTER TABLE lich_su_gui_loi_chuc ADD COLUMN IF NOT EXISTS trang_thai  VARCHAR(20);


-- ============================================================================
-- 9. CHẤM CÔNG
-- (Bảng này app.py đã tự tạo lúc chạy qua ensure_cham_cong_table() — khai báo
--  lại ở đây để có sẵn ngay từ đầu cho khách hàng mới, tránh phải chờ lần
--  chạy tính năng Chấm công đầu tiên. Giữ CHÍNH XÁC cấu trúc app.py đang dùng.)
-- ============================================================================
CREATE TABLE IF NOT EXISTS cham_cong (
    id                  SERIAL PRIMARY KEY,
    nhan_vien_id        INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE,
    ngay                DATE NOT NULL,
    ma_cong             VARCHAR(10),
    ca_ngay             VARCHAR(10),
    ca_dem              VARCHAR(10),
    gio_tang_ca         NUMERIC(5,2) DEFAULT 0,
    gio_tang_ca_le      NUMERIC(5,2) DEFAULT 0,
    ghi_chu             TEXT,
    nguon               VARCHAR(20) DEFAULT 'THU_CONG',
    created_by          VARCHAR(100),
    created_at          TIMESTAMP DEFAULT NOW(),
    updated_at          TIMESTAMP DEFAULT NOW(),
    UNIQUE(nhan_vien_id, ngay)
);


-- ============================================================================
-- 10. DANH MỤC HÀNH CHÍNH (Tỉnh/Thành phố + Phường/Xã) — dùng cho tờ khai BHXH
-- ============================================================================
-- LƯU Ý: đây là bảng DÙNG CHUNG (danh mục hành chính không đổi theo tenant),
-- cần được NẠP ĐẦY ĐỦ DỮ LIỆU 63 tỉnh/thành + toàn bộ phường/xã 1 lần duy
-- nhất (ví dụ từ dữ liệu công bố của Tổng cục Thống kê/GSO), vì file này chỉ
-- tạo cấu trúc bảng chứ KHÔNG chứa sẵn dữ liệu địa danh đầy đủ. Nếu bạn cần,
-- mình có thể hỗ trợ viết script nạp dữ liệu riêng.
CREATE TABLE IF NOT EXISTS danh_muc_tinh (
    ma_tinh     VARCHAR(10) PRIMARY KEY,
    ten_tinh    TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS danh_muc_phuong_xa (
    ma_xa       VARCHAR(10) PRIMARY KEY,
    ten_xa      TEXT NOT NULL,
    ma_tinh     VARCHAR(10) REFERENCES danh_muc_tinh(ma_tinh)
);


-- ============================================================================
-- 11. CHAT NỘI BỘ (đang phát triển — bảng chuẩn bị sẵn để không phải chạy
--     thêm migration khi tính năng hoàn thiện)
-- ============================================================================
CREATE TABLE IF NOT EXISTS chat_rooms (
    id          SERIAL PRIMARY KEY,
    ten_phong   TEXT NOT NULL,
    loai_phong  VARCHAR(20) DEFAULT 'NHOM',  -- 1_1 / NHOM / TOAN_CONG_TY
    created_at  TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS chat_messages (
    id              SERIAL PRIMARY KEY,
    room_id         INTEGER REFERENCES chat_rooms(id) ON DELETE CASCADE,
    nhan_vien_id    INTEGER REFERENCES nhan_vien(id) ON DELETE SET NULL,
    noi_dung        TEXT,
    created_at      TIMESTAMP DEFAULT NOW()
);


-- ============================================================================
-- INDEX PHỤ TRỢ (tăng tốc các truy vấn phổ biến)
-- ============================================================================
CREATE INDEX IF NOT EXISTS idx_nhan_vien_trang_thai ON nhan_vien(trang_thai);
CREATE INDEX IF NOT EXISTS idx_ung_vien_trang_thai ON ung_vien(trang_thai);
CREATE INDEX IF NOT EXISTS idx_lich_su_cong_tac_nv ON lich_su_cong_tac(nhan_vien_id);
CREATE INDEX IF NOT EXISTS idx_quyet_dinh_nhan_su_nv ON quyet_dinh_nhan_su(nhan_vien_id);
CREATE INDEX IF NOT EXISTS idx_ho_so_nhan_vien_nv ON ho_so_nhan_vien(nhan_vien_id);
CREATE INDEX IF NOT EXISTS idx_danh_muc_phuong_xa_tinh ON danh_muc_phuong_xa(ma_tinh);
