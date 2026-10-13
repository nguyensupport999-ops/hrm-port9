# luong_bh_hieu_luc.py
import re
from datetime import date


def parse_luong(s):
    """'4.650.000' / '4,650,000' / '4650000.0' / 4650000 -> 4650000 (int). Rỗng -> 0."""
    if s is None:
        return 0
    if isinstance(s, (int, float)):
        return int(s)
    t = str(s).strip()
    t = re.sub(r'[.,]0{1,2}$', '', t)   # bỏ '.0' / '.00', không đụng '.000' ngăn nghìn
    t = re.sub(r'\D', '', t)
    return int(t) if t else 0


def ensure_table(conn):
    """Gọi SAU ensure_qdns_table() vì có khoá ngoại tới quyet_dinh_nhan_su."""
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS lich_su_luong_bh (
            id SERIAL PRIMARY KEY,
            nhan_vien_id INTEGER NOT NULL REFERENCES nhan_vien(id) ON DELETE CASCADE,
            qd_id INTEGER REFERENCES quyet_dinh_nhan_su(id) ON DELETE CASCADE,
            so_qd VARCHAR(50),
            ngay_qd DATE,
            ngay_hieu_luc DATE NOT NULL,
            thang_ap_dung DATE NOT NULL,      -- ngày 1 của tháng bắt đầu kê khai mức mới
            luong_cu BIGINT,
            luong_moi BIGINT NOT NULL,
            ly_do TEXT,
            da_ap_dung BOOLEAN DEFAULT FALSE, -- đã ghi vào nhan_vien.luong_bao_hiem chưa
            nguoi_tao VARCHAR(100),
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    c.execute("CREATE INDEX IF NOT EXISTS idx_lsl_nv_thang ON lich_su_luong_bh (nhan_vien_id, thang_ap_dung)")
    conn.commit()


def dong_bo_luong_vao_ho_so(conn, hom_nay=None, commit=True):
    """Ghi các mức lương đã ĐẾN ngày hiệu lực vào nhan_vien.luong_bao_hiem."""
    hom_nay = hom_nay or date.today()
    c = conn.cursor()
    c.execute("""
        SELECT id, nhan_vien_id, luong_moi FROM lich_su_luong_bh
        WHERE COALESCE(da_ap_dung, FALSE) = FALSE AND ngay_hieu_luc <= %s
        ORDER BY ngay_hieu_luc, id
    """, (hom_nay,))
    for _id, nv_id, luong in c.fetchall():
        c.execute("UPDATE nhan_vien SET luong_bao_hiem = %s WHERE id = %s", (str(luong), nv_id))
        c.execute("UPDATE lich_su_luong_bh SET da_ap_dung = TRUE WHERE id = %s", (_id,))
    if commit:
        conn.commit()


def hoan_tac_khi_xoa_qd(conn, qd_id):
    """Gọi TRƯỚC khi xoá QĐ: trả nhan_vien.luong_bao_hiem về mức trước đó."""
    c = conn.cursor()
    c.execute("SELECT nhan_vien_id, luong_cu FROM lich_su_luong_bh WHERE qd_id = %s AND da_ap_dung", (qd_id,))
    for nv_id, luong_cu in c.fetchall():
        c.execute("""
            SELECT luong_moi FROM lich_su_luong_bh
            WHERE nhan_vien_id = %s AND da_ap_dung AND qd_id <> %s
            ORDER BY thang_ap_dung DESC, ngay_hieu_luc DESC, id DESC LIMIT 1
        """, (nv_id, qd_id))
        r = c.fetchone()
        c.execute("UPDATE nhan_vien SET luong_bao_hiem = %s WHERE id = %s",
                  (str(r[0] if r else luong_cu), nv_id))


def lay_luong_hieu_luc(conn, den_ngay, nv_ids=None, strict=False):
    """{nhan_vien_id: lương BH (int)} có hiệu lực TRONG THÁNG chứa `den_ngay`.
    - Có điều chỉnh với thang_ap_dung <= tháng đó -> lấy mức mới nhất.
    - Có điều chỉnh nhưng đều SAU tháng đó        -> lấy luong_cu của điều chỉnh sớm nhất.
    - Chưa từng điều chỉnh                        -> lấy nhan_vien.luong_bao_hiem.
    strict=True: không nuốt lỗi (bảng lịch sử chưa có...) mà raise để nơi gọi báo cảnh báo.
    """
    thang = date(den_ngay.year, den_ngay.month, 1)
    loc, params = "", {"thang": thang}
    if nv_ids:
        loc = "AND nv.id = ANY(%(ids)s)"
        params["ids"] = list(nv_ids)
    c = conn.cursor()
    try:
        c.execute(f"""
            SELECT nv.id, nv.luong_bao_hiem,
                   (SELECT l.luong_moi FROM lich_su_luong_bh l
                     WHERE l.nhan_vien_id = nv.id AND l.thang_ap_dung <= %(thang)s
                     ORDER BY l.thang_ap_dung DESC, l.ngay_hieu_luc DESC, l.id DESC LIMIT 1),
                   (SELECT l.luong_cu FROM lich_su_luong_bh l
                     WHERE l.nhan_vien_id = nv.id
                     ORDER BY l.thang_ap_dung ASC, l.ngay_hieu_luc ASC, l.id ASC LIMIT 1)
            FROM nhan_vien nv WHERE 1=1 {loc}
        """, params)
        rows = c.fetchall()
    except Exception:
        conn.rollback()
        if strict:
            raise
        # bảng lịch sử chưa có -> dùng lương hiện tại trong hồ sơ
        c = conn.cursor()
        c.execute("SELECT id, luong_bao_hiem, NULL, NULL FROM nhan_vien nv WHERE 1=1 " + loc, params)
        rows = c.fetchall()
    kq = {}
    for nv_id, lbh, moi, cu in rows:
        kq[nv_id] = int(moi) if moi is not None else (int(cu) if cu is not None else parse_luong(lbh))
    return kq


def lay_dieu_chinh_trong_ky(conn, tu_ngay, den_ngay):
    """Các điều chỉnh có NGÀY HIỆU LỰC nằm trong kỳ (để ghi chú trong báo cáo)."""
    c = conn.cursor()
    c.execute("""
        SELECT nv.ma_nv, nv.ho_ten, l.so_qd, l.ngay_hieu_luc, l.thang_ap_dung, l.luong_cu, l.luong_moi
        FROM lich_su_luong_bh l JOIN nhan_vien nv ON nv.id = l.nhan_vien_id
        WHERE l.ngay_hieu_luc BETWEEN %s AND %s
        ORDER BY l.ngay_hieu_luc, nv.ma_nv
    """, (tu_ngay, den_ngay))
    return c.fetchall()
    
def lay_dieu_chinh_thang(conn, thang):
    """{nhan_vien_id: (so_qd, ngay_hieu_luc, luong_cu, luong_moi)} của các điều chỉnh BẮT ĐẦU áp dụng
    đúng tháng `thang` (nếu có nhiều điều chỉnh trong tháng thì lấy cái hiệu lực muộn nhất)."""
    thang = date(thang.year, thang.month, 1)
    c = conn.cursor()
    try:
        c.execute("""
            SELECT nhan_vien_id, so_qd, ngay_hieu_luc, luong_cu, luong_moi
            FROM lich_su_luong_bh WHERE thang_ap_dung = %s
            ORDER BY ngay_hieu_luc, id
        """, (thang,))
        return {r[0]: (r[1], r[2], r[3], r[4]) for r in c.fetchall()}
    except Exception:
        conn.rollback()
        return {}


def dem_nv_co_lich_su(conn):
    """Số nhân viên đã có ít nhất 1 điều chỉnh trong lich_su_luong_bh (raise nếu bảng chưa có)."""
    c = conn.cursor()
    c.execute("SELECT COUNT(DISTINCT nhan_vien_id) FROM lich_su_luong_bh")
    return c.fetchone()[0]