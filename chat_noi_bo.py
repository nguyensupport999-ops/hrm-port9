"""
============================================================================
 MODULE "CHAT NỘI BỘ" — tách riêng khỏi app.py (thay thế chat_utils.py cũ)
============================================================================
Cách tích hợp vào app.py:

    import chat_noi_bo                       # thay cho `import chat_utils`

    elif menu == "💬 Chat nội bộ":
        chat_noi_bo.render()

Toàn bộ schema, logic phân quyền và giao diện đều nằm trong 1 file này.

--------------------------------------------------------------------------
QUY TẮC NGHIỆP VỤ (theo đúng yêu cầu)
--------------------------------------------------------------------------
1) Mỗi PHÒNG BAN có đúng 1 "group mặc định" (room_type='department'), tự
   động tạo & tự động đồng bộ thành viên theo trường nhan_vien.phong_ban_lam_viec
   hiện tại (mỗi khi nhân viên được điều chuyển, họ tự rời group phòng cũ và
   vào group phòng mới — KHÔNG cần thao tác thủ công).
2) Nhân viên KHÔNG có quyền tự thêm mình / rời khỏi group phòng ban (is_locked)
   và cũng không rời được kênh "📢 Thông báo chung" (broadcast, is_locked).
3) Thành viên thuộc phòng ban 'Hội Đồng Quản Trị' / 'Ban Tổng Giám Đốc'
   (xem PHONG_BAN_LANH_DAO_CAO_CAP) được tạo group tùy ý tới toàn thể nhân
   viên hoặc thành viên từ nhiều phòng ban khác nhau.
4) Chỉ "Trưởng phòng" (chuc_vu chứa từ khóa trưởng — KHÔNG bao gồm cấp phó)
   mới được tạo group gồm thành viên từ phòng ban khác.
5) Nhân viên thường: chỉ chat 1:1 hoặc chat trong group mình đã được thêm sẵn,
   KHÔNG thấy tùy chọn "Tạo nhóm mới".

Điểm mở rộng cho tương lai (CHƯA triển khai ở bản này, chỉ chừa chỗ):
   - message_type đã có sẵn 'payslip' (kt_luong gửi phiếu lương theo ma_nv).
   - Có thể thêm message_type='leave_request' cho nghiệp vụ gửi/duyệt đơn
     nghỉ phép ngay trong khung chat của group phòng ban tương ứng.
============================================================================
"""

import re
import hashlib
import unicodedata
from contextlib import contextmanager
from datetime import datetime

import psycopg2
import psycopg2.extras
import streamlit as st
import streamlit.components.v1 as components


# ============================================================================
# 1) HẰNG SỐ — PHẢI ĐỒNG BỘ với PHONG_BAN_THU_TU / PHONG_BAN_LANH_DAO_CAO_CAP
#    đang khai báo trong app.py.  Nếu sửa danh mục phòng ban bên app.py,
#    nhớ sửa lại đúng y hệt ở đây (copy nguyên văn từng dòng).
# ============================================================================
PHONG_BAN_THU_TU = [
    "Hội Đồng Quản Trị",
    "Ban Tổng Giám Đốc",
    "Phòng Hành Chính Nhân Sự",
    "Phòng Kinh Doanh",
    "Phòng Tài Chính",
    "Phòng Điều Độ",
    "Tổ Cơ Giới",
    "Đội Bốc Xếp",
    "Phòng KT - Cơ Điện",
    "Đội Bảo Vệ",
]

PHONG_BAN_LANH_DAO_CAO_CAP = ("Hội Đồng Quản Trị", "Ban Tổng Giám Đốc")

# Từ khóa nhận diện "Trưởng phòng / Tổ trưởng / Đội trưởng..." — so khớp theo
# TỪ KHÓA (không phân biệt hoa/thường), giống hệt cách app.py đang làm ở phần
# sơ đồ tổ chức, để không lặp lại lỗi "so khớp chính xác" từng khiến nhiều
# Trưởng phòng / Tổ trưởng / Đội trưởng không được nhận diện đúng.
TU_KHOA_TRUONG_PHONG = [
    "trưởng phòng", "tổ trưởng", "đội trưởng", "trưởng ban",
    "trưởng bộ phận", "quản đốc", "phụ trách",
]

MAU_AVATAR = [
    "#0084ff", "#f59e0b", "#10b981", "#8b5cf6", "#ef4444",
    "#06b6d4", "#ec4899", "#84cc16", "#f97316", "#6366f1",
]

# Khoảng thời gian tự làm mới tin nhắn (giây) — dùng cho st.fragment(run_every=...)
AUTO_REFRESH_INTERVAL = "5s"


# ============================================================================
# 2) PHÂN QUYỀN
# ============================================================================
def _chuan_hoa(s):
    """Chuẩn hóa chuỗi: bỏ khoảng trắng thừa, viết thường."""
    return " ".join((s or "").strip().split()).lower()


def la_lanh_dao_cao_cap(phong_ban_lam_viec):
    """True nếu nhân viên thuộc HĐQT / BTGĐ — được tạo group toàn công ty / liên phòng ban."""
    pb = _chuan_hoa(phong_ban_lam_viec)
    return any(pb == _chuan_hoa(x) for x in PHONG_BAN_LANH_DAO_CAO_CAP)


def la_truong_phong(chuc_vu):
    """True nếu là trưởng đơn vị (KHÔNG tính cấp phó) — được tạo group liên phòng ban."""
    cv = _chuan_hoa(chuc_vu)
    if not cv or cv.startswith("phó"):
        return False
    return any(tk in cv for tk in TU_KHOA_TRUONG_PHONG)


def co_quyen_tao_nhom(nv):
    """Chỉ lãnh đạo cấp cao hoặc trưởng phòng mới được tạo group tùy chỉnh.
    Nhân viên thường chỉ được chat 1:1."""
    if not nv:
        return False
    return (la_lanh_dao_cao_cap(nv.get("phong_ban_lam_viec"))
            or la_truong_phong(nv.get("chuc_vu")))


def co_quyen_quan_ly_nhom(nv, room):
    """Ai được thêm / xóa thành viên, đổi tên, giải tán 1 group tùy chỉnh
    (room_type='group'): người tạo ra group đó, hoặc lãnh đạo cấp cao."""
    if not nv or not room:
        return False
    if room.get("room_type") != "group":
        return False
    return (room.get("created_by") == nv.get("id")
            or la_lanh_dao_cao_cap(nv.get("phong_ban_lam_viec")))


# ============================================================================
# 3) KẾT NỐI DB — context manager tránh rò rỉ connection
# ============================================================================
@contextmanager
def _db_cursor(dict_cursor=True):
    """Context manager đảm bảo connection luôn được đóng, kể cả khi exception."""
    db = st.session_state.db_engine.get_connection()
    try:
        if dict_cursor:
            cur = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        else:
            cur = db.cursor()
        yield db, cur
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


# ============================================================================
# 4) SCHEMA — tạo bảng & đồng bộ group phòng ban mặc định
# ============================================================================
def init_chat_tables():
    """Tạo bảng chat nếu chưa có (an toàn khi gọi lại nhiều lần)."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            # ---------- chat_rooms ----------
            c.execute("""
                CREATE TABLE IF NOT EXISTS chat_rooms (
                    id          SERIAL PRIMARY KEY,
                    room_name   VARCHAR(150),
                    room_type   VARCHAR(20) NOT NULL DEFAULT 'private',
                                -- private | department | group | broadcast
                    phong_ban   VARCHAR(100),
                    is_locked   BOOLEAN DEFAULT FALSE,
                    created_by  INTEGER,
                    created_at  TIMESTAMP DEFAULT NOW(),
                    updated_at  TIMESTAMP DEFAULT NOW(),
                    deleted_at  TIMESTAMP
                )
            """)
            # Đảm bảo các cột bổ sung (cho DB cũ đã có bảng nhưng thiếu cột)
            for stmt in [
                "ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS phong_ban  VARCHAR(100)",
                "ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS is_locked  BOOLEAN DEFAULT FALSE",
                "ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP DEFAULT NOW()",
                "ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMP",
            ]:
                c.execute(stmt)

            # ---------- chat_participants ----------
            c.execute("""
                CREATE TABLE IF NOT EXISTS chat_participants (
                    room_id      INTEGER NOT NULL,
                    user_id      INTEGER NOT NULL,
                    joined_at    TIMESTAMP DEFAULT NOW(),
                    last_read_at TIMESTAMP DEFAULT NOW(),
                    PRIMARY KEY (room_id, user_id)
                )
            """)
            c.execute("ALTER TABLE chat_participants "
                      "ADD COLUMN IF NOT EXISTS last_read_at TIMESTAMP DEFAULT NOW()")

            # ---------- chat_messages ----------
            c.execute("""
                CREATE TABLE IF NOT EXISTS chat_messages (
                    id            SERIAL PRIMARY KEY,
                    room_id       INTEGER NOT NULL,
                    sender_id     INTEGER NOT NULL,
                    content       TEXT,
                    message_type  VARCHAR(20) DEFAULT 'text',
                                  -- text | image | file | payslip | system
                    file_url      TEXT,
                    file_name     TEXT,
                    file_size     INTEGER,
                    is_read       BOOLEAN DEFAULT FALSE,
                    sent_at       TIMESTAMP DEFAULT NOW()
                )
            """)
            for stmt in [
                "ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS message_type VARCHAR(20) DEFAULT 'text'",
                "ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS file_url    TEXT",
                "ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS file_name   TEXT",
                "ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS file_size   INTEGER",
            ]:
                c.execute(stmt)

            # ---------- Indexes cho hiệu năng ----------
            c.execute("""
                CREATE INDEX IF NOT EXISTS idx_chat_messages_room_sent
                ON chat_messages (room_id, sent_at)
            """)
            c.execute("""
                CREATE INDEX IF NOT EXISTS idx_chat_participants_user
                ON chat_participants (user_id)
            """)

            # ---------- Kênh "📢 Thông báo chung" (broadcast, khóa) ----------
            c.execute("""
                INSERT INTO chat_rooms
                    (room_name, room_type, is_locked, created_by, created_at, updated_at)
                SELECT '📢 Thông báo chung', 'broadcast', TRUE, 0, NOW(), NOW()
                WHERE NOT EXISTS (
                    SELECT 1 FROM chat_rooms
                    WHERE room_type = 'broadcast' AND deleted_at IS NULL
                )
            """)
            c.execute("""
                SELECT id FROM chat_rooms
                WHERE room_type = 'broadcast' AND deleted_at IS NULL
                LIMIT 1
            """)
            row = c.fetchone()
            broadcast_id = row[0] if row else None
            if broadcast_id:
                c.execute("""
                    INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
                    SELECT %s, nv.id, NOW(), NOW()
                    FROM nhan_vien nv
                    WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
                      AND NOT EXISTS (
                        SELECT 1 FROM chat_participants cp
                        WHERE cp.room_id = %s AND cp.user_id = nv.id
                    )
                """, (broadcast_id, broadcast_id))

            db.commit()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi khởi tạo bảng chat: {e}")
        return False


def sync_department_rooms():
    """Đảm bảo mỗi phòng ban trong PHONG_BAN_THU_TU có đúng 1 group mặc định,
    và thành viên của group đó luôn khớp với nhan_vien.phong_ban_lam_viec hiện tại.

    An toàn khi gọi nhiều lần (idempotent).  Nên gọi lại ngay sau khi 1 Quyết định
    điều chuyển phòng ban được lưu, để group chat cập nhật tức thì."""
    try:
        with _db_cursor() as (db, c):
            for ten_pb in PHONG_BAN_THU_TU:
                # Tìm hoặc tạo room cho phòng ban
                c.execute("""
                    SELECT id FROM chat_rooms
                    WHERE room_type = 'department' AND phong_ban = %s AND deleted_at IS NULL
                """, (ten_pb,))
                row = c.fetchone()
                if row:
                    room_id = row["id"]
                else:
                    c.execute("""
                        INSERT INTO chat_rooms
                            (room_name, room_type, phong_ban, is_locked, created_by, updated_at)
                        VALUES (%s, 'department', %s, TRUE, 0, NOW())
                        RETURNING id
                    """, (ten_pb, ten_pb))
                    room_id = c.fetchone()["id"]

                # Nhân viên ĐÚNG RA phải ở trong room này
                c.execute("""
                    SELECT id FROM nhan_vien
                    WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                      AND phong_ban_lam_viec = %s
                """, (ten_pb,))
                dung_id = {r["id"] for r in c.fetchall()}

                # Nhân viên HIỆN ĐANG ở trong room
                c.execute("SELECT user_id FROM chat_participants WHERE room_id = %s",
                          (room_id,))
                hien_co_id = {r["user_id"] for r in c.fetchall()}

                # Thêm mới
                for uid in (dung_id - hien_co_id):
                    c.execute("""
                        INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
                        VALUES (%s, %s, NOW(), NOW())
                        ON CONFLICT (room_id, user_id) DO NOTHING
                    """, (room_id, uid))

                # Xóa người không còn thuộc phòng ban
                for uid in (hien_co_id - dung_id):
                    c.execute("DELETE FROM chat_participants "
                              "WHERE room_id = %s AND user_id = %s",
                              (room_id, uid))

            db.commit()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đồng bộ group phòng ban: {e}")
        return False


# ============================================================================
# 5) DATA LAYER
# ============================================================================

# ---------- Nhân viên ----------

def get_current_nv(user_id):
    """Lấy thông tin nhân viên hiện tại (chuc_vu, phong_ban_lam_viec...) để tính quyền.
    Luôn đọc trực tiếp từ DB để phản ánh đúng ngay sau khi bị điều chuyển."""
    try:
        with _db_cursor() as (_, c):
            c.execute("""
                SELECT id, ma_nv, ho_ten, chuc_vu, phong_ban_lam_viec
                FROM nhan_vien WHERE id = %s
            """, (user_id,))
            row = c.fetchone()
            return dict(row) if row else None
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy thông tin nhân viên: {e}")
        return None


def get_all_employees(exclude_id=None):
    """Lấy danh sách toàn bộ nhân viên đang làm / thử việc."""
    try:
        with _db_cursor() as (_, c):
            c.execute("""
                SELECT id, ma_nv, ho_ten, dien_thoai, phong_ban_lam_viec, chuc_vu
                FROM nhan_vien
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                ORDER BY ho_ten ASC
            """)
            rows = c.fetchall()
        return [r for r in rows if r["id"] != exclude_id] if exclude_id else rows
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy danh sách nhân viên: {e}")
        return []


def search_employees(keyword, exclude_id=None):
    """Tìm nhân viên theo tên hoặc mã NV (ILIKE)."""
    try:
        with _db_cursor() as (_, c):
            kw = f"%{keyword}%"
            c.execute("""
                SELECT id, ma_nv, ho_ten, dien_thoai, phong_ban_lam_viec, chuc_vu
                FROM nhan_vien
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
                  AND (ho_ten ILIKE %s OR ma_nv ILIKE %s)
                ORDER BY ho_ten ASC LIMIT 30
            """, (kw, kw))
            rows = c.fetchall()
        return [r for r in rows if r["id"] != exclude_id] if exclude_id else rows
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi tìm nhân viên: {e}")
        return []


# ---------- Phòng chat ----------

def get_all_rooms_admin():
    """Lấy TẤT CẢ phòng chat — chỉ dùng cho chế độ giám sát admin/HR."""
    try:
        with _db_cursor() as (_, c):
            c.execute("""
                SELECT * FROM chat_rooms
                WHERE deleted_at IS NULL
                ORDER BY updated_at DESC
            """)
            return c.fetchall()
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy toàn bộ phòng (giám sát): {e}")
        return []


def get_user_chat_rooms(user_id):
    """Lấy danh sách phòng chat mà nhân viên đang tham gia."""
    try:
        with _db_cursor() as (_, c):
            c.execute("""
                SELECT r.* FROM chat_rooms r
                JOIN chat_participants p ON r.id = p.room_id
                WHERE p.user_id = %s AND r.deleted_at IS NULL
                ORDER BY r.updated_at DESC
            """, (user_id,))
            return c.fetchall()
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy danh sách phòng: {e}")
        return []


def get_room_by_id(room_id):
    """Lấy thông tin 1 phòng chat theo ID."""
    try:
        with _db_cursor() as (_, c):
            c.execute("SELECT * FROM chat_rooms WHERE id = %s AND deleted_at IS NULL",
                      (room_id,))
            return c.fetchone()
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy phòng #{room_id}: {e}")
        return None


def get_room_participants(room_id):
    """Lấy danh sách thành viên của 1 phòng chat."""
    try:
        with _db_cursor() as (_, c):
            c.execute("""
                SELECT nv.id, nv.ma_nv, nv.ho_ten, nv.dien_thoai,
                       nv.phong_ban_lam_viec, nv.chuc_vu
                FROM chat_participants p
                JOIN nhan_vien nv ON p.user_id = nv.id
                WHERE p.room_id = %s
                ORDER BY nv.ho_ten
            """, (room_id,))
            return c.fetchall()
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy thành viên phòng: {e}")
        return []


def get_room_display_name(room, user_id):
    """Tên hiển thị phòng chat — tùy loại phòng."""
    rt = room.get("room_type", "private")
    if rt == "broadcast":
        return "📢 Thông báo chung"
    if rt == "department":
        return f"🏢 {room.get('phong_ban') or room.get('room_name')}"
    if rt == "group":
        return f"👥 {room.get('room_name') or 'Nhóm chat'}"
    # private — hiển thị tên người kia
    participants = get_room_participants(room["id"])
    for p in participants:
        if p["id"] != user_id:
            return p["ho_ten"]
    return "Phòng chat"


def get_room_last_message(room_id):
    """Tin nhắn mới nhất trong 1 phòng."""
    try:
        with _db_cursor() as (_, c):
            c.execute("""
                SELECT m.*, nv.ho_ten AS sender_name
                FROM chat_messages m
                LEFT JOIN nhan_vien nv ON m.sender_id = nv.id
                WHERE m.room_id = %s
                ORDER BY m.sent_at DESC LIMIT 1
            """, (room_id,))
            return c.fetchone()
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy tin nhắn cuối: {e}")
        return None


def get_room_unread_count(room_id, user_id):
    """Đếm số tin nhắn chưa đọc trong phòng (so với last_read_at)."""
    try:
        with _db_cursor(dict_cursor=False) as (_, c):
            c.execute("""
                SELECT COUNT(*) FROM chat_messages m
                WHERE m.room_id = %s AND m.sender_id != %s
                  AND m.sent_at > COALESCE(
                    (SELECT last_read_at FROM chat_participants
                     WHERE room_id = %s AND user_id = %s),
                    '1970-01-01'
                  )
            """, (room_id, user_id, room_id, user_id))
            return c.fetchone()[0]
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đếm tin chưa đọc: {e}")
        return 0


def get_total_unread_count(user_id):
    """Tổng tin nhắn chưa đọc trên TẤT CẢ các phòng — dùng để hiển thị badge tổng."""
    try:
        with _db_cursor(dict_cursor=False) as (_, c):
            c.execute("""
                SELECT COUNT(*) FROM chat_messages m
                JOIN chat_participants p
                  ON p.room_id = m.room_id AND p.user_id = %s
                JOIN chat_rooms r
                  ON r.id = m.room_id AND r.deleted_at IS NULL
                WHERE m.sender_id != %s
                  AND m.sent_at > COALESCE(p.last_read_at, '1970-01-01')
            """, (user_id, user_id))
            return c.fetchone()[0]
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đếm tổng chưa đọc: {e}")
        return 0


# ---------- Tin nhắn ----------

def get_room_messages(room_id, limit=300):
    """Lấy lịch sử tin nhắn của 1 phòng (ASC theo thời gian)."""
    try:
        with _db_cursor() as (_, c):
            c.execute("""
                SELECT m.*, nv.ho_ten AS sender_name
                FROM chat_messages m
                LEFT JOIN nhan_vien nv ON m.sender_id = nv.id
                WHERE m.room_id = %s
                ORDER BY m.sent_at ASC LIMIT %s
            """, (room_id, limit))
            return c.fetchall()
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy tin nhắn: {e}")
        return []


def mark_messages_as_read(room_id, user_id):
    """Đánh dấu tất cả tin nhắn trong phòng là đã đọc (cho user_id)."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            c.execute("""
                UPDATE chat_messages SET is_read = TRUE
                WHERE room_id = %s AND sender_id != %s AND is_read = FALSE
            """, (room_id, user_id))
            c.execute("""
                UPDATE chat_participants SET last_read_at = NOW()
                WHERE room_id = %s AND user_id = %s
            """, (room_id, user_id))
            db.commit()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đánh dấu đã đọc: {e}")
        return False


def send_message(room_id, sender_id, content, message_type="text",
                 file_url=None, file_name=None, file_size=None):
    """Gửi 1 tin nhắn vào phòng (và cập nhật updated_at của phòng)."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            c.execute("""
                INSERT INTO chat_messages
                    (room_id, sender_id, content, message_type,
                     file_url, file_name, file_size)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (room_id, sender_id, content, message_type,
                  file_url, file_name, file_size))
            c.execute("UPDATE chat_rooms SET updated_at = NOW() WHERE id = %s",
                      (room_id,))
            db.commit()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi gửi tin nhắn: {e}")
        return False


def send_system_message(room_id, content):
    """Gửi tin nhắn hệ thống (ví dụ: 'A đã thêm B vào nhóm')."""
    return send_message(room_id, 0, content, message_type="system")


def send_image_message(room_id, sender_id, file_url, file_name, caption=""):
    return send_message(room_id, sender_id, caption, "image", file_url, file_name)


def send_file_message(room_id, sender_id, file_url, file_name, file_size):
    return send_message(room_id, sender_id, file_name, "file", file_url, file_name, file_size)


def send_payslip_message(room_id, sender_id, employee_name, month, year, salary_data):
    """Gửi phiếu lương dạng tin nhắn đặc biệt (message_type='payslip')."""
    content = f"""📄 PHIẾU LƯƠNG THÁNG {month}/{year}

👤 Nhân viên: {employee_name}
{'-' * 30}
Lương cơ bản:       {salary_data.get('luong_co_ban', 0):>13,.0f} VNĐ
Phụ cấp chức vụ:    {salary_data.get('phu_cap_chuc_vu', 0):>13,.0f} VNĐ
Phụ cấp thâm niên:  {salary_data.get('phu_cap_tnvk', 0):>13,.0f} VNĐ
Phụ cấp trách nhiệm:{salary_data.get('phu_cap_tnn', 0):>13,.0f} VNĐ
{'-' * 30}
Tổng thu nhập:      {salary_data.get('tong', 0):>13,.0f} VNĐ
{'-' * 30}
BHXH (8%):           {salary_data.get('bhxh', 0):>12,.0f} VNĐ
BHYT (1.5%):         {salary_data.get('bhyt', 0):>12,.0f} VNĐ
BHTN (1%):           {salary_data.get('bhtn', 0):>12,.0f} VNĐ
{'-' * 30}
Thực nhận:          {salary_data.get('thuc_nhan', 0):>13,.0f} VNĐ

📅 Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}"""
    return send_message(room_id, sender_id, content, "payslip")


# ---------- Phòng — tạo / sửa / xóa ----------

def create_private_room(user1_id, user2_id):
    """Tạo (hoặc trả về nếu đã tồn tại) phòng chat 1:1."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            # Kiểm tra đã có phòng 1:1 giữa 2 người chưa
            c.execute("""
                SELECT r.id FROM chat_rooms r
                JOIN chat_participants p1 ON r.id = p1.room_id
                JOIN chat_participants p2 ON r.id = p2.room_id
                WHERE r.room_type = 'private'
                  AND p1.user_id = %s AND p2.user_id = %s
                  AND r.deleted_at IS NULL
            """, (user1_id, user2_id))
            existing = c.fetchone()
            if existing:
                return existing[0]

            c.execute("""
                INSERT INTO chat_rooms (room_name, room_type, created_by, updated_at)
                VALUES (%s, 'private', %s, NOW()) RETURNING id
            """, (f"Chat {user1_id}-{user2_id}", user1_id))
            room_id = c.fetchone()[0]
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id)
                VALUES (%s, %s), (%s, %s)
            """, (room_id, user1_id, room_id, user2_id))
            db.commit()
            return room_id
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi tạo phòng 1:1: {e}")
        return None


def create_group_room(room_name, creator_id, member_ids):
    """Tạo group tùy chỉnh (room_type='group').
    ⚠️ KHÔNG gọi trực tiếp từ UI mà không kiểm tra co_quyen_tao_nhom() trước."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            c.execute("""
                INSERT INTO chat_rooms (room_name, room_type, created_by, updated_at)
                VALUES (%s, 'group', %s, NOW()) RETURNING id
            """, (room_name, creator_id))
            room_id = c.fetchone()[0]
            all_ids = {creator_id, *member_ids}
            for uid in all_ids:
                c.execute("""
                    INSERT INTO chat_participants (room_id, user_id)
                    VALUES (%s, %s) ON CONFLICT (room_id, user_id) DO NOTHING
                """, (room_id, uid))
            db.commit()

            # System message
            send_system_message(room_id, f"🎉 Nhóm \"{room_name}\" đã được tạo "
                                f"với {len(all_ids)} thành viên.")
            return room_id
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi tạo nhóm: {e}")
        return None


def rename_room(room_id, new_name):
    """Đổi tên phòng chat (chỉ áp dụng cho room_type='group')."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            c.execute("UPDATE chat_rooms SET room_name = %s, updated_at = NOW() "
                      "WHERE id = %s AND room_type = 'group'",
                      (new_name, room_id))
            db.commit()
        send_system_message(room_id, f"✏️ Nhóm đã được đổi tên thành \"{new_name}\".")
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đổi tên phòng: {e}")
        return False


def add_participants_to_room(room_id, user_ids, added_by_name=""):
    """Thêm thành viên vào phòng chat (với system message)."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            added_names = []
            for uid in user_ids:
                c.execute("""
                    INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
                    VALUES (%s, %s, NOW(), NOW())
                    ON CONFLICT (room_id, user_id) DO NOTHING
                """, (room_id, uid))
                # Lấy tên nhân viên
                c.execute("SELECT ho_ten FROM nhan_vien WHERE id = %s", (uid,))
                r = c.fetchone()
                if r:
                    added_names.append(r[0])
            db.commit()

        if added_names:
            ten_ds = ", ".join(added_names)
            prefix = f"{added_by_name} đã thêm" if added_by_name else "Đã thêm"
            send_system_message(room_id, f"➕ {prefix} {ten_ds} vào nhóm.")
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi thêm thành viên: {e}")
        return False


def remove_participant_from_room(room_id, user_id, removed_by_name="", removed_name=""):
    """Xóa 1 thành viên khỏi phòng chat (với system message)."""
    try:
        with _db_cursor(dict_cursor=False) as (db, c):
            c.execute("DELETE FROM chat_participants WHERE room_id = %s AND user_id = %s",
                      (room_id, user_id))
            db.commit()
        if removed_name:
            if removed_by_name and removed_by_name != removed_name:
                send_system_message(room_id,
                                    f"➖ {removed_by_name} đã xóa {removed_name} khỏi nhóm.")
            else:
                send_system_message(room_id, f"🚪 {removed_name} đã rời nhóm.")
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi xóa thành viên: {e}")
        return False


def roi_nhom(room, user_id, user_name=""):
    """Rời group — CHỈ áp dụng cho room_type='group' (nhóm tùy chỉnh).
    Không cho rời 'department' / 'broadcast' (is_locked) hay 'private'."""
    if room.get("room_type") != "group" or room.get("is_locked"):
        return False
    return remove_participant_from_room(room["id"], user_id,
                                        removed_name=user_name)


def delete_room(room_id, deleted_by_name=""):
    """Soft-delete phòng chat (giải tán nhóm)."""
    try:
        if deleted_by_name:
            send_system_message(room_id,
                                f"🗑️ {deleted_by_name} đã giải tán nhóm này.")
        with _db_cursor(dict_cursor=False) as (db, c):
            c.execute("UPDATE chat_rooms SET deleted_at = NOW() WHERE id = %s",
                      (room_id,))
            db.commit()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi xóa phòng: {e}")
        return False


# ---------- Xuất lịch sử ----------

def export_room_transcript(room_id, room_display_name):
    """Xuất toàn bộ lịch sử chat ra văn bản thuần (.txt)."""
    msgs = get_room_messages(room_id, limit=5000)
    lines = [
        f"LỊCH SỬ CHAT — {room_display_name}",
        f"Xuất lúc: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        "=" * 60, "",
    ]
    for m in msgs:
        ts = m["sent_at"].strftime("%d/%m/%Y %H:%M") if m.get("sent_at") else ""
        ten = m.get("sender_name") or "Hệ thống"
        mt = m.get("message_type", "text")
        noi_dung = m.get("content") or ""

        if mt == "system":
            lines.append(f"[{ts}] *** {noi_dung} ***")
        elif mt == "image":
            lines.append(f"[{ts}] {ten}: [Hình ảnh: {m.get('file_name', '')}] {noi_dung}")
        elif mt == "file":
            lines.append(f"[{ts}] {ten}: [Tệp đính kèm: {m.get('file_name', '')}]")
        elif mt == "payslip":
            lines.append(f"[{ts}] {ten}:\n{noi_dung}")
        else:
            lines.append(f"[{ts}] {ten}: {noi_dung}")
    return "\n".join(lines)


# ============================================================================
# 6) UPLOAD FILE (Supabase Storage) — giữ nguyên logic gốc
# ============================================================================
def sanitize_filename(filename):
    """Loại bỏ dấu tiếng Việt & ký tự đặc biệt khỏi tên file."""
    normalized = unicodedata.normalize("NFD", filename)
    no_accent = "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")
    no_accent = no_accent.replace("đ", "d").replace("Đ", "D")
    no_accent = re.sub(r"\s+", "_", no_accent)
    safe = re.sub(r"[^A-Za-z0-9_.\-]", "", no_accent)
    return safe or "file"


def _get_supabase_client():
    from supabase import create_client
    tenant = st.session_state.get("tenant")
    if tenant:
        url = tenant.get("supabase_url")
        key = tenant.get("supabase_key")
    else:
        import os
        from dotenv import load_dotenv
        load_dotenv()
        url = os.getenv("SUPABASE_URL")
        key = os.getenv("SUPABASE_KEY")
    if not url or not key:
        return None
    return create_client(url, key)


def upload_chat_file(file_bytes, filename, content_type):
    """Upload file lên Supabase Storage, trả về đường dẫn (path)."""
    try:
        sb = _get_supabase_client()
        if not sb:
            return None
        bucket_name = "ho-so-nhan-vien"
        safe_name = sanitize_filename(filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = f"chat_files/{timestamp}_{safe_name}"
        sb.storage.from_(bucket_name).upload(
            path=path, file=file_bytes,
            file_options={"content-type": content_type or "application/octet-stream"},
        )
        return path
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi upload file: {e}")
        return None


def get_chat_file_bytes(file_url):
    """Tải file từ Supabase Storage."""
    try:
        sb = _get_supabase_client()
        if not sb:
            return None
        return sb.storage.from_("ho-so-nhan-vien").download(file_url)
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi tải file: {e}")
        return None


# ============================================================================
# 7) GIAO DIỆN (Streamlit UI)
# ============================================================================

# ---------- Tiện ích hiển thị ----------

def _mau_avatar(ten):
    h = int(hashlib.md5((ten or "?").encode("utf-8")).hexdigest(), 16)
    return MAU_AVATAR[h % len(MAU_AVATAR)]


def _initials(ten):
    parts = (ten or "?").strip().split()
    return (parts[-1][0] if parts else "?").upper()


def _avatar_html(ten, size=38):
    return (
        f'<div style="width:{size}px;height:{size}px;border-radius:50%;'
        f'background:{_mau_avatar(ten)};color:#fff;display:inline-flex;'
        f'align-items:center;justify-content:center;font-weight:700;'
        f'font-size:{size * 0.42:.0f}px;flex-shrink:0;">'
        f'{_initials(ten)}</div>'
    )


def _thoi_gian(dt):
    """Định dạng thời gian thân thiện (Hôm nay → HH:MM, tuần này → Thứ HH:MM, ...)."""
    if not dt:
        return ""
    now = datetime.now()
    if dt.date() == now.date():
        return dt.strftime("%H:%M")
    if (now.date() - dt.date()).days < 7:
        weekdays_vi = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]
        return f"{weekdays_vi[dt.weekday()]} {dt.strftime('%H:%M')}"
    if dt.year == now.year:
        return dt.strftime("%d/%m %H:%M")
    return dt.strftime("%d/%m/%Y %H:%M")


def _last_msg_preview(last, my_id):
    """Tạo chuỗi xem trước tin nhắn cuối cho sidebar."""
    if not last:
        return "Chưa có tin nhắn nào"
    mt = last.get("message_type", "text")
    if mt == "system":
        return (last.get("content") or "")[:42]

    if last["sender_id"] == my_id:
        prefix = "Bạn: "
    elif last["sender_id"] == 0:
        prefix = ""
    else:
        name_parts = (last.get("sender_name") or "").split()
        prefix = f"{name_parts[-1]}: " if name_parts else ""

    if mt == "image":
        body = "📷 Hình ảnh"
    elif mt == "file":
        body = f"📎 {last.get('file_name', 'Tệp')}"
    elif mt == "payslip":
        body = "📄 Phiếu lương"
    else:
        body = last.get("content") or ""

    return (prefix + body)[:48]


def _inject_auto_scroll():
    """Inject JavaScript để tự cuộn xuống cuối khung tin nhắn."""
    components.html("""
    <script>
    (function() {
        const containers = window.parent.document.querySelectorAll(
            '[data-testid="stVerticalBlock"] [data-testid="stVerticalBlockBorderWrapper"]'
        );
        // Cuộn container cuối cùng có chiều cao cố định (khung tin nhắn)
        for (let i = containers.length - 1; i >= 0; i--) {
            const el = containers[i].querySelector('[data-testid="ScrollToBottomContainer"]')
                       || containers[i];
            if (el.scrollHeight > el.clientHeight) {
                el.scrollTop = el.scrollHeight;
                break;
            }
        }
    })();
    </script>
    """, height=0)


# ---------- CSS toàn cục ----------

def _css():
    st.markdown("""
    <style>
    /* ---- Badge tin chưa đọc ---- */
    .cnb-badge {
        background: #ef4444; color: #fff; font-size: 10.5px; font-weight: 700;
        border-radius: 10px; padding: 1px 7px; min-width: 18px; text-align: center;
        display: inline-block; line-height: 1.5;
    }
    /* ---- Phiếu lương ---- */
    .cnb-payslip {
        background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 14px;
        padding: 12px 16px; font-size: 13px; white-space: pre-wrap;
        font-family: ui-monospace, 'Cascadia Code', 'Fira Code', monospace;
    }
    /* ---- File đính kèm ---- */
    .cnb-file {
        display: inline-flex; align-items: center; gap: 8px;
        padding: 8px 14px; background: rgba(0,0,0,0.04); border-radius: 12px;
        font-size: 13px;
    }
    /* ---- System message ---- */
    .cnb-system-msg {
        text-align: center; font-size: 12px; color: #94a3b8;
        padding: 6px 0; font-style: italic;
    }
    /* ---- Sidebar room item ---- */
    .cnb-room-card {
        display: flex; align-items: center; gap: 10px;
        padding: 8px 10px; border-radius: 10px; margin-bottom: 2px;
    }
    .cnb-room-card:hover { background: #f1f5f9; }
    /* ---- Ẩn label mặc định của một số widget ---- */
    div[data-testid="stFileUploader"] > label { display: none; }
    </style>
    """, unsafe_allow_html=True)


# ---------- Dialog: Tạo phòng mới ----------

@st.dialog("➕ Tạo phòng mới", width="large")
def _dialog_tao_phong(nv):
    """Dialog tạo phòng chat 1:1 hoặc nhóm mới."""
    co_the_tao_nhom = co_quyen_tao_nhom(nv)
    tab_labels = ["💬 Chat 1:1"]
    if co_the_tao_nhom:
        tab_labels.append("👥 Tạo nhóm mới")
    tabs = st.tabs(tab_labels)

    # ─── Tab 1: Chat 1:1 ───
    with tabs[0]:
        kw = st.text_input("🔍 Tìm theo tên hoặc mã NV", key="cnb_search_private",
                           placeholder="Nhập tên hoặc mã nhân viên...")
        emps = (search_employees(kw, exclude_id=nv["id"]) if kw
                else get_all_employees(exclude_id=nv["id"]))

        if not emps:
            st.info("Không tìm thấy nhân viên nào.")
        else:
            st.caption(f"Hiển thị {min(len(emps), 25)}/{len(emps)} kết quả")
            for e in emps[:25]:
                c1, c2 = st.columns([0.75, 0.25])
                with c1:
                    pb = e.get("phong_ban_lam_viec") or ""
                    cv = e.get("chuc_vu") or ""
                    sub = f"{pb}" + (f" · {cv}" if cv else "")
                    st.markdown(f"**{e['ho_ten']}**  \n{sub}")
                with c2:
                    if st.button("💬 Nhắn tin", key=f"priv_{e['id']}",
                                 use_container_width=True):
                        rid = create_private_room(nv["id"], e["id"])
                        if rid:
                            st.session_state["cnb_room_id"] = rid
                            st.rerun()

    # ─── Tab 2: Tạo nhóm (chỉ hiện nếu có quyền) ───
    if co_the_tao_nhom:
        with tabs[1]:
            pham_vi = ("toàn công ty"
                       if la_lanh_dao_cao_cap(nv.get("phong_ban_lam_viec"))
                       else "liên phòng ban")
            st.caption(f"✅ Bạn có quyền tạo nhóm với thành viên {pham_vi}.")

            ten_nhom = st.text_input("Tên nhóm", key="cnb_group_name",
                                     placeholder="VD: Dự án cảng Hòn La 2026")

            # Chọn theo phòng ban trước (tiện hơn khi nhiều nhân viên)
            loc_pb = st.multiselect(
                "Lọc theo phòng ban", PHONG_BAN_THU_TU,
                key="cnb_group_filter_pb",
                placeholder="Chọn phòng ban để lọc (hoặc bỏ qua)..."
            )

            kw2 = st.text_input("🔍 Tìm thành viên", key="cnb_search_group",
                                placeholder="Nhập tên để tìm nhanh...")
            emps2 = (search_employees(kw2, exclude_id=nv["id"]) if kw2
                     else get_all_employees(exclude_id=nv["id"]))

            # Lọc theo phòng ban nếu có chọn
            if loc_pb:
                emps2 = [e for e in emps2
                         if e.get("phong_ban_lam_viec") in loc_pb]

            options = {
                f"{e['ho_ten']} — {e.get('phong_ban_lam_viec') or ''}"
                + (f" ({e.get('chuc_vu', '')})" if e.get('chuc_vu') else ""):
                e["id"]
                for e in emps2
            }
            chon = st.multiselect("Thành viên", list(options.keys()),
                                  key="cnb_group_members")

            c1, c2 = st.columns(2)
            with c1:
                st.metric("Đã chọn", f"{len(chon)} người")
            with c2:
                if st.button("🚀 Tạo nhóm", type="primary",
                             disabled=not (ten_nhom.strip() and chon),
                             use_container_width=True):
                    member_ids = [options[k] for k in chon]
                    rid = create_group_room(ten_nhom.strip(), nv["id"], member_ids)
                    if rid:
                        st.session_state["cnb_room_id"] = rid
                        st.success("✅ Đã tạo nhóm thành công!")
                        st.rerun()


# ---------- Dialog: Quản lý nhóm ----------

@st.dialog("⚙️ Quản lý nhóm", width="large")
def _dialog_quan_ly_nhom(nv, room):
    """Dialog quản lý thành viên, đổi tên, giải tán nhóm."""
    participants = get_room_participants(room["id"])
    is_manager = co_quyen_quan_ly_nhom(nv, room)
    rt = room.get("room_type", "private")

    # ─── Thông tin nhóm ───
    display = get_room_display_name(room, nv["id"])
    st.markdown(f"### {display}")

    cols_info = st.columns(3)
    with cols_info[0]:
        st.metric("Thành viên", len(participants))
    with cols_info[1]:
        type_label = {
            "private": "Chat 1:1", "department": "Phòng ban",
            "group": "Nhóm tùy chỉnh", "broadcast": "Kênh thông báo",
        }
        st.metric("Loại phòng", type_label.get(rt, rt))
    with cols_info[2]:
        if room.get("is_locked"):
            st.metric("Trạng thái", "🔒 Khóa")
        else:
            st.metric("Trạng thái", "🔓 Mở")

    st.divider()

    # ─── Tab quản lý ───
    tab_labels = ["👥 Thành viên"]
    if is_manager:
        tab_labels.extend(["➕ Thêm thành viên", "✏️ Cài đặt nhóm"])
    mgmt_tabs = st.tabs(tab_labels)

    # Tab: Danh sách thành viên
    with mgmt_tabs[0]:
        for p in participants:
            c1, c2, c3 = st.columns([0.1, 0.6, 0.3])
            with c1:
                st.markdown(_avatar_html(p["ho_ten"], 32), unsafe_allow_html=True)
            with c2:
                is_creator = (p["id"] == room.get("created_by"))
                badge = " 👑" if is_creator else ""
                st.markdown(f"**{p['ho_ten']}**{badge}")
                st.caption(f"{p.get('phong_ban_lam_viec', '')} · {p.get('chuc_vu', '')}")
            with c3:
                # Cho phép quản lý nhóm xóa thành viên (trừ chính mình và người tạo)
                if (is_manager and rt == "group"
                        and p["id"] != nv["id"]
                        and p["id"] != room.get("created_by")):
                    if st.button("❌ Xóa", key=f"rm_{room['id']}_{p['id']}",
                                 use_container_width=True):
                        remove_participant_from_room(
                            room["id"], p["id"],
                            removed_by_name=nv["ho_ten"],
                            removed_name=p["ho_ten"])
                        st.rerun()

    # Tab: Thêm thành viên (chỉ cho quản lý nhóm group)
    if is_manager:
        with mgmt_tabs[1]:
            if rt != "group":
                st.info("ℹ️ Không thể thêm thành viên vào phòng loại này."
                        " Thành viên phòng ban được đồng bộ tự động.")
            else:
                current_ids = {p["id"] for p in participants}
                candidates = get_all_employees()
                candidates = [e for e in candidates if e["id"] not in current_ids]

                kw_add = st.text_input("🔍 Tìm nhân viên", key=f"add_search_{room['id']}",
                                       placeholder="Nhập tên...")
                if kw_add:
                    kw_lower = kw_add.lower()
                    candidates = [e for e in candidates
                                  if kw_lower in (e.get("ho_ten") or "").lower()
                                  or kw_lower in (e.get("ma_nv") or "").lower()]

                options_add = {
                    f"{e['ho_ten']} — {e.get('phong_ban_lam_viec', '')}": e["id"]
                    for e in candidates
                }
                chon_them = st.multiselect("Chọn thành viên mới",
                                           list(options_add.keys()),
                                           key=f"add_mem_{room['id']}")
                if st.button("➕ Thêm vào nhóm", key=f"btn_add_mem_{room['id']}",
                             type="primary", disabled=not chon_them):
                    ids = [options_add[k] for k in chon_them]
                    add_participants_to_room(room["id"], ids,
                                            added_by_name=nv["ho_ten"])
                    st.success(f"✅ Đã thêm {len(ids)} thành viên!")
                    st.rerun()

        # Tab: Cài đặt nhóm
        with mgmt_tabs[2]:
            if rt == "group":
                st.subheader("✏️ Đổi tên nhóm")
                new_name = st.text_input("Tên mới", value=room.get("room_name", ""),
                                         key=f"rename_{room['id']}")
                if st.button("💾 Lưu tên mới", key=f"btn_rename_{room['id']}"):
                    if new_name.strip() and new_name.strip() != room.get("room_name"):
                        rename_room(room["id"], new_name.strip())
                        st.success("✅ Đã đổi tên nhóm!")
                        st.rerun()

                st.divider()
                st.subheader("🗑️ Giải tán nhóm")
                st.warning("⚠️ Thao tác này không thể hoàn tác. "
                           "Tất cả thành viên sẽ bị xóa khỏi nhóm.")
                confirm = st.text_input("Nhập tên nhóm để xác nhận giải tán:",
                                        key=f"confirm_del_{room['id']}",
                                        placeholder=room.get("room_name", ""))
                if st.button("🗑️ Giải tán nhóm", key=f"btn_del_{room['id']}",
                             type="primary",
                             disabled=(confirm != room.get("room_name"))):
                    delete_room(room["id"], deleted_by_name=nv["ho_ten"])
                    st.session_state["cnb_room_id"] = None
                    st.success("✅ Nhóm đã bị giải tán.")
                    st.rerun()
            else:
                st.info("ℹ️ Phòng hệ thống (phòng ban / broadcast) không thể "
                        "đổi tên hoặc giải tán.")


# ---------- Sidebar: Danh sách phòng chat ----------

def _render_sidebar(nv, rooms, active_room_id):
    """Cột bên trái: nút tạo phòng, ô tìm kiếm, danh sách phòng."""
    # Nút hành động trên cùng
    top1, top2 = st.columns([0.75, 0.25])
    with top1:
        if st.button("➕ Tạo phòng mới", use_container_width=True, type="primary"):
            _dialog_tao_phong(nv)
    with top2:
        if st.button("🔄", use_container_width=True, help="Làm mới"):
            # Đồng bộ lại group phòng ban khi bấm refresh
            sync_department_rooms()
            st.rerun()

    # Ô tìm kiếm
    tim = st.text_input("🔎 Tìm phòng chat...", key="cnb_search_room",
                        label_visibility="collapsed",
                        placeholder="🔎 Tìm phòng chat...")

    # Tổng tin chưa đọc
    total_unread = get_total_unread_count(nv["id"])
    if total_unread:
        st.markdown(f'<div style="text-align:center;padding:4px 0;">'
                    f'<span class="cnb-badge" style="font-size:12px;padding:2px 10px;">'
                    f'📬 {total_unread} tin nhắn chưa đọc</span></div>',
                    unsafe_allow_html=True)

    st.divider()

    if not rooms:
        st.caption("Chưa có phòng chat nào.")
        return

    # Phân nhóm phòng chat theo loại
    broadcast_rooms = [r for r in rooms if r.get("room_type") == "broadcast"]
    dept_rooms = [r for r in rooms if r.get("room_type") == "department"]
    group_rooms = [r for r in rooms if r.get("room_type") == "group"]
    private_rooms = [r for r in rooms if r.get("room_type") == "private"]

    def _render_room_btn(room):
        """Render 1 dòng phòng chat trong sidebar."""
        display = get_room_display_name(room, nv["id"])
        # Lọc theo từ khóa tìm kiếm
        if tim and tim.strip().lower() not in display.lower():
            return False  # bị lọc, không hiển thị

        last = get_room_last_message(room["id"])
        preview = _last_msg_preview(last, nv["id"])
        unread = get_room_unread_count(room["id"], nv["id"])
        ts = _thoi_gian(last["sent_at"]) if last else ""
        is_active = (room["id"] == active_room_id)

        # Layout: avatar | tên + preview | thời gian + badge
        cols = st.columns([0.12, 0.66, 0.22])
        with cols[0]:
            name_for_avatar = display.lstrip("📢🏢👥 ")
            st.markdown(_avatar_html(name_for_avatar, 34), unsafe_allow_html=True)
        with cols[1]:
            btn_type = "primary" if is_active else "secondary"
            if st.button(display, key=f"room_btn_{room['id']}",
                         use_container_width=True, type=btn_type):
                st.session_state["cnb_room_id"] = room["id"]
                mark_messages_as_read(room["id"], nv["id"])
                st.rerun()
            lock_icon = "🔒 " if room.get("is_locked") else ""
            st.caption(f"{lock_icon}{preview}")
        with cols[2]:
            if ts:
                st.caption(ts)
            if unread:
                st.markdown(f'<span class="cnb-badge">{unread}</span>',
                            unsafe_allow_html=True)
        return True  # đã hiển thị

    # Hiển thị theo nhóm
    shown_any = False

    if broadcast_rooms:
        for r in broadcast_rooms:
            if _render_room_btn(r):
                shown_any = True

    if dept_rooms:
        if not tim:
            st.caption("🏢 PHÒNG BAN")
        for r in dept_rooms:
            if _render_room_btn(r):
                shown_any = True

    if group_rooms:
        if not tim:
            st.caption("👥 NHÓM")
        for r in group_rooms:
            if _render_room_btn(r):
                shown_any = True

    if private_rooms:
        if not tim:
            st.caption("💬 TIN NHẮN RIÊNG")
        for r in private_rooms:
            if _render_room_btn(r):
                shown_any = True

    if not shown_any and tim:
        st.caption(f"Không tìm thấy phòng nào khớp '{tim}'.")


# ---------- Khung tin nhắn (với auto-refresh) ----------

def _render_message_bubble(msg, is_self, nv_id):
    """Render 1 tin nhắn duy nhất."""
    mtype = msg.get("message_type", "text")
    sender = msg.get("sender_name") or "Hệ thống"
    time_str = _thoi_gian(msg["sent_at"])

    # System message — căn giữa, không avatar
    if mtype == "system":
        st.markdown(f'<div class="cnb-system-msg">{msg.get("content", "")}</div>',
                    unsafe_allow_html=True)
        return

    # Tin nhắn thường — dùng st.chat_message
    avatar_char = "👤" if is_self else _initials(sender)
    role = "user" if is_self else "assistant"

    with st.chat_message(role, avatar=avatar_char):
        if not is_self:
            st.caption(f"**{sender}**")

        if mtype == "image" and msg.get("file_url"):
            file_bytes = get_chat_file_bytes(msg["file_url"])
            if file_bytes:
                st.image(file_bytes, width=280)
            if msg.get("content"):
                st.write(msg["content"])
        elif mtype == "file" and msg.get("file_url"):
            st.markdown(f'<div class="cnb-file">📎 <b>{msg.get("file_name", "")}</b></div>',
                        unsafe_allow_html=True)
            file_bytes = get_chat_file_bytes(msg["file_url"])
            if file_bytes:
                st.download_button("⬇️ Tải xuống", file_bytes,
                                   file_name=msg.get("file_name", "file"),
                                   key=f"dl_{msg['id']}")
        elif mtype == "payslip":
            st.markdown(f'<div class="cnb-payslip">{msg.get("content", "")}</div>',
                        unsafe_allow_html=True)
        else:
            st.write(msg.get("content", ""))

        st.caption(time_str)


def _render_messages_fragment(room_id, nv_id):
    """Hiển thị danh sách tin nhắn trong container cuộn, tự làm mới mỗi 5 giây."""
    msgs = get_room_messages(room_id)
    if not msgs:
        st.info("💬 Chưa có tin nhắn nào — hãy bắt đầu cuộc trò chuyện!")
    else:
        for m in msgs:
            _render_message_bubble(m, is_self=(m["sender_id"] == nv_id), nv_id=nv_id)

    mark_messages_as_read(room_id, nv_id)


# ---------- Cửa sổ chat chính ----------

def _render_chat_window(nv, room):
    """Cột bên phải: header, khung tin nhắn, ô nhập."""
    participants = get_room_participants(room["id"])
    display = get_room_display_name(room, nv["id"])
    rt = room.get("room_type", "private")

    # ─── Header ───
    h1, h2, h3 = st.columns([0.55, 0.22, 0.23])
    with h1:
        st.markdown(f"### {display}")
        lock_txt = "  ·  🔒 Không thể rời" if room.get("is_locked") else ""
        st.caption(f"{len(participants)} thành viên{lock_txt}")
    with h2:
        transcript = export_room_transcript(room["id"], display)
        st.download_button(
            "💾 Lưu hội thoại", transcript.encode("utf-8"),
            file_name=f"chat_{datetime.now().strftime('%Y%m%d')}_{room['id']}.txt",
            key=f"export_{room['id']}", use_container_width=True,
        )
    with h3:
        if st.button("⚙️ Tùy chọn", key=f"opts_{room['id']}",
                     use_container_width=True):
            _dialog_quan_ly_nhom(nv, room)

    st.divider()

    # ─── Khung tin nhắn (container cuộn cố định) ───
    with st.container(height=420):
        _render_messages_fragment(room["id"], nv["id"])

    # ─── Nút rời nhóm (cho group mở) ───
    if rt == "group" and not room.get("is_locked"):
        with st.expander("🚪 Rời nhóm này"):
            st.warning("Bạn sẽ không còn nhận tin nhắn từ nhóm này nữa.")
            if st.button("Xác nhận rời nhóm", key=f"leave_{room['id']}",
                         type="primary"):
                roi_nhom(room, nv["id"], user_name=nv["ho_ten"])
                st.session_state["cnb_room_id"] = None
                st.rerun()

    # ─── Ô nhập tin nhắn + đính kèm file ───
    st.divider()

    # File uploader ĐẶT NGOÀI form (tránh bug Streamlit file_uploader trong form)
    file_dinh_kem = st.file_uploader(
        "📎 Đính kèm file", key=f"file_{room['id']}",
        label_visibility="collapsed",
        help="Chọn file ảnh hoặc tài liệu để gửi kèm",
    )
    if file_dinh_kem:
        st.caption(f"📎 Đã chọn: {file_dinh_kem.name} "
                   f"({file_dinh_kem.size / 1024:.0f} KB)")

    # Chat input (pinned to bottom)
    noi_dung = st.chat_input("Nhập tin nhắn...", key=f"chat_input_{room['id']}")

    if noi_dung or file_dinh_kem:
        # Xử lý gửi file (nếu có)
        if file_dinh_kem:
            file_bytes = file_dinh_kem.read()
            content_type = file_dinh_kem.type or "application/octet-stream"
            path = upload_chat_file(file_bytes, file_dinh_kem.name, content_type)
            if path:
                if content_type.startswith("image/"):
                    send_image_message(room["id"], nv["id"], path,
                                       file_dinh_kem.name,
                                       noi_dung or "")
                else:
                    send_file_message(room["id"], nv["id"], path,
                                      file_dinh_kem.name, len(file_bytes))
                # Nếu có cả text lẫn file (không phải ảnh), gửi text riêng
                if noi_dung and not content_type.startswith("image/"):
                    send_message(room["id"], nv["id"], noi_dung)
            else:
                st.error("❌ Tải file lên thất bại. Vui lòng thử lại.")
        elif noi_dung:
            send_message(room["id"], nv["id"], noi_dung)

        st.rerun()


# ---------- Trang trống (chưa chọn phòng) ----------

def _render_empty_state():
    st.markdown("---")
    st.markdown("""
    <div style="display:flex;flex-direction:column;align-items:center;
                justify-content:center;padding:80px 20px;color:#94a3b8;">
        <div style="font-size:64px;margin-bottom:16px;">💬</div>
        <h2 style="color:#334155;margin:0 0 8px;">Chọn một phòng chat</h2>
        <p>Chọn phòng từ danh sách bên trái hoặc tạo phòng mới để bắt đầu</p>
        <p style="font-size:13px;">💡 Bấm <b>"➕ Tạo phòng mới"</b> để bắt đầu trò chuyện</p>
    </div>
    """, unsafe_allow_html=True)


# ---------- Chế độ giám sát cho tài khoản hệ thống ----------

def _ten_phong_giam_sat(room):
    """Tên hiển thị phòng — dùng riêng cho chế độ giám sát (không có user_id)."""
    if room.get("room_type") != "private":
        return get_room_display_name(room, None)
    ps = get_room_participants(room["id"])
    return " ↔ ".join(p["ho_ten"] for p in ps) if ps else "Chat 1:1"


def _render_tai_khoan_he_thong():
    """Giao diện khi tài khoản hệ thống (admin/HR) mở Chat nội bộ.
    Không gắn với nhân viên nên chỉ xem (giám sát), không gửi được."""
    st.title("💬 Chat nội bộ")
    role = st.session_state.get("role")
    uname = st.session_state.get("username")
    st.info(
        f"ℹ️ Tài khoản **{uname}** (vai trò: `{role}`) không gắn với hồ sơ nhân viên "
        "nên không thể **gửi/nhận** tin nhắn cá nhân.\n\n"
        "👉 Để chat như nhân viên, đăng xuất rồi đăng nhập bằng **số điện thoại "
        "của nhân viên thực tế** trong hệ thống."
    )

    if role not in ("admin", "hr"):
        return

    st.divider()
    st.subheader("🕵️ Chế độ giám sát (chỉ xem)")
    st.caption("Dành riêng cho admin/HR xem lại nội dung các phòng chat. "
               "Không thể gửi tin nhắn ở chế độ này.")

    if not st.session_state.get("_cnb_da_khoi_tao"):
        init_chat_tables()
        sync_department_rooms()
        st.session_state["_cnb_da_khoi_tao"] = True

    _css()
    rooms = get_all_rooms_admin()
    if not rooms:
        st.caption("Chưa có phòng chat nào trong hệ thống.")
        return

    # Phân loại để dễ tìm
    col_filter1, col_filter2 = st.columns(2)
    with col_filter1:
        room_types = {
            "Tất cả": None, "📢 Broadcast": "broadcast",
            "🏢 Phòng ban": "department", "👥 Nhóm": "group",
            "💬 1:1": "private",
        }
        loc_loai = st.selectbox("Lọc theo loại", list(room_types.keys()),
                                key="cnb_admin_filter_type")
    with col_filter2:
        tim_gs = st.text_input("Tìm phòng", key="cnb_admin_search",
                               placeholder="Nhập tên phòng...")

    filtered = rooms
    if room_types[loc_loai]:
        filtered = [r for r in filtered if r.get("room_type") == room_types[loc_loai]]
    if tim_gs:
        kw = tim_gs.strip().lower()
        filtered = [r for r in filtered if kw in _ten_phong_giam_sat(r).lower()]

    if not filtered:
        st.caption("Không tìm thấy phòng nào phù hợp.")
        return

    st.caption(f"Tìm thấy {len(filtered)} phòng chat")

    options = {
        f"{_ten_phong_giam_sat(r)}  ·  #{r['id']}  ·  "
        f"{r.get('room_type', '?')}": r["id"]
        for r in filtered
    }
    chon = st.selectbox("Chọn phòng để xem", list(options.keys()),
                        key="cnb_admin_view_room")
    if not chon:
        return

    room_id = options[chon]
    room = get_room_by_id(room_id)

    # Thông tin phòng
    if room:
        ps = get_room_participants(room_id)
        cols_meta = st.columns(4)
        with cols_meta[0]:
            st.metric("Thành viên", len(ps))
        with cols_meta[1]:
            total_msgs = len(get_room_messages(room_id, limit=9999))
            st.metric("Tổng tin nhắn", total_msgs)
        with cols_meta[2]:
            st.metric("Loại", room.get("room_type", "?"))
        with cols_meta[3]:
            created = room.get("created_at")
            st.metric("Tạo lúc",
                      created.strftime("%d/%m/%Y") if created else "?")

        # Danh sách thành viên (có thể mở rộng)
        with st.expander(f"👥 Danh sách thành viên ({len(ps)})"):
            for p in ps:
                st.caption(f"👤 {p['ho_ten']} — "
                           f"{p.get('phong_ban_lam_viec', '')} · "
                           f"{p.get('chuc_vu', '')}")

    st.divider()

    # Xuất lịch sử
    display_name = _ten_phong_giam_sat(room) if room else f"Phòng #{room_id}"
    transcript = export_room_transcript(room_id, display_name)
    st.download_button("💾 Xuất lịch sử chat", transcript.encode("utf-8"),
                       file_name=f"admin_chat_{room_id}_{datetime.now().strftime('%Y%m%d')}.txt",
                       key=f"admin_export_{room_id}")

    # Hiển thị tin nhắn
    with st.container(height=480):
        msgs = get_room_messages(room_id)
        if not msgs:
            st.caption("Chưa có tin nhắn nào trong phòng này.")
        else:
            for m in msgs:
                mtype = m.get("message_type", "text")
                sender = m.get("sender_name") or "Hệ thống"
                time_str = _thoi_gian(m["sent_at"])

                if mtype == "system":
                    st.markdown(f'<div class="cnb-system-msg">{m.get("content", "")}</div>',
                                unsafe_allow_html=True)
                    continue

                with st.chat_message("assistant", avatar=_initials(sender)):
                    st.caption(f"**{sender}** · {time_str}")
                    if mtype == "payslip":
                        st.markdown(f'<div class="cnb-payslip">{m.get("content", "")}</div>',
                                    unsafe_allow_html=True)
                    elif mtype == "image":
                        if m.get("file_url"):
                            fb = get_chat_file_bytes(m["file_url"])
                            if fb:
                                st.image(fb, width=280)
                        if m.get("content"):
                            st.write(m["content"])
                    elif mtype == "file":
                        st.markdown(f'<div class="cnb-file">'
                                    f'📎 <b>{m.get("file_name", "")}</b></div>',
                                    unsafe_allow_html=True)
                    else:
                        st.write(m.get("content", ""))


# ============================================================================
# 8) ENTRY POINT
# ============================================================================

def render():
    """Entry point — gọi từ app.py:

        import chat_noi_bo

        elif menu == "💬 Chat nội bộ":
            chat_noi_bo.render()
    """
    if not st.session_state.get("logged_in"):
        st.warning("⚠️ Vui lòng đăng nhập để sử dụng Chat nội bộ.")
        return

    user_id = st.session_state.get("nhan_vien_id")
    if not user_id:
        # Tài khoản hệ thống — không gắn nhân viên → chế độ giám sát
        _render_tai_khoan_he_thong()
        return

    # Khởi tạo & đồng bộ 1 LẦN / phiên (tránh query lặp mỗi rerun)
    if not st.session_state.get("_cnb_da_khoi_tao"):
        ok1 = init_chat_tables()
        ok2 = sync_department_rooms()
        if not ok1 or not ok2:
            st.error("❌ Không thể khởi tạo hệ thống chat. Kiểm tra kết nối DB.")
            return
        st.session_state["_cnb_da_khoi_tao"] = True

    nv = get_current_nv(user_id)
    if not nv:
        st.error("❌ Không tìm thấy thông tin nhân viên. Vui lòng liên hệ admin.")
        return

    # ─── Giao diện chính ───
    st.title("💬 Chat nội bộ")
    st.caption(f"Xin chào **{nv['ho_ten']}** · {nv.get('phong_ban_lam_viec', '')} · "
               f"{nv.get('chuc_vu', '')}")
    _css()

    rooms = get_user_chat_rooms(user_id)
    active_room_id = st.session_state.get("cnb_room_id")

    # Kiểm tra phòng đang mở vẫn còn tồn tại
    active_room = None
    if active_room_id:
        active_room = next((r for r in rooms if r["id"] == active_room_id), None)
        if not active_room:
            # Phòng đã bị xóa hoặc user đã rời → reset
            st.session_state["cnb_room_id"] = None
            active_room_id = None

    # ─── Layout 2 cột ───
    col_side, col_main = st.columns([0.32, 0.68], gap="medium")

    with col_side:
        _render_sidebar(nv, rooms, active_room_id)

    with col_main:
        if active_room:
            _render_chat_window(nv, active_room)
        else:
            _render_empty_state()
