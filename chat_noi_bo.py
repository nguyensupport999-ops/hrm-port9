# chat_noi_bo.py
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
from datetime import datetime

import psycopg2
import psycopg2.extras
import streamlit as st


# ============================================================================
# 1) HẰNG SỐ — PHẢI ĐỒNG BỘ với PHONG_BAN_THU_TU / PHONG_BAN_LANH_DAO_CAO_CAP
#    đang khai báo trong app.py. Nếu sửa danh mục phòng ban bên app.py,
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
# Trưởng phòng/Tổ trưởng/Đội trưởng không được nhận diện đúng.
TU_KHOA_TRUONG_PHONG = [
    "trưởng phòng", "tổ trưởng", "đội trưởng", "trưởng ban",
    "trưởng bộ phận", "quản đốc", "phụ trách",
]

MAU_AVATAR = ["#0084ff", "#f59e0b", "#10b981", "#8b5cf6", "#ef4444",
              "#06b6d4", "#ec4899", "#84cc16", "#f97316", "#6366f1"]


# ============================================================================
# 2) PHÂN QUYỀN
# ============================================================================
def _chuan_hoa(s):
    return " ".join((s or "").strip().split()).lower()


def la_lanh_dao_cao_cap(phong_ban_lam_viec):
    """True nếu nhân viên thuộc HĐQT/BTGĐ — được tạo group toàn công ty / liên phòng ban."""
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
    return la_lanh_dao_cao_cap(nv.get("phong_ban_lam_viec")) or la_truong_phong(nv.get("chuc_vu"))


def co_quyen_quan_ly_nhom(nv, room):
    """Ai được thêm/xóa thành viên, đổi tên, giải tán 1 group tùy chỉnh (room_type='group'):
    người tạo ra group đó, hoặc lãnh đạo cấp cao (giám sát toàn công ty)."""
    if not nv or not room:
        return False
    if room.get("room_type") != "group":
        return False
    return room.get("created_by") == nv.get("id") or la_lanh_dao_cao_cap(nv.get("phong_ban_lam_viec"))


# ============================================================================
# 3) KẾT NỐI DB — helper dùng chung
# ============================================================================
def _conn(dict_cursor=True):
    db = st.session_state.db_engine.get_connection()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor) if dict_cursor else db.cursor()
    return db, c


# ============================================================================
# 4) SCHEMA — tạo bảng & đồng bộ group phòng ban mặc định
# ============================================================================
def init_chat_tables():
    """Tạo bảng chat nếu chưa có (an toàn khi gọi lại nhiều lần)."""
    try:
        db, c = _conn(dict_cursor=False)

        c.execute("""
            CREATE TABLE IF NOT EXISTS chat_rooms (
                id SERIAL PRIMARY KEY,
                room_name VARCHAR(150),
                room_type VARCHAR(20) NOT NULL DEFAULT 'private',  -- private | department | group | broadcast
                phong_ban VARCHAR(100),        -- chỉ có giá trị khi room_type = 'department'
                is_locked BOOLEAN DEFAULT FALSE,  -- TRUE = không cho tự thêm/rời (department, broadcast)
                created_by INTEGER,
                created_at TIMESTAMP DEFAULT NOW(),
                updated_at TIMESTAMP DEFAULT NOW(),
                deleted_at TIMESTAMP
            )
        """)
        c.execute("ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS phong_ban VARCHAR(100)")
        c.execute("ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS is_locked BOOLEAN DEFAULT FALSE")
        c.execute("ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP DEFAULT NOW()")
        c.execute("ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMP")

        c.execute("""
            CREATE TABLE IF NOT EXISTS chat_participants (
                room_id INTEGER NOT NULL,
                user_id INTEGER NOT NULL,
                joined_at TIMESTAMP DEFAULT NOW(),
                last_read_at TIMESTAMP DEFAULT NOW(),
                PRIMARY KEY (room_id, user_id)
            )
        """)
        c.execute("ALTER TABLE chat_participants ADD COLUMN IF NOT EXISTS last_read_at TIMESTAMP DEFAULT NOW()")

        c.execute("""
            CREATE TABLE IF NOT EXISTS chat_messages (
                id SERIAL PRIMARY KEY,
                room_id INTEGER NOT NULL,
                sender_id INTEGER NOT NULL,
                content TEXT,
                message_type VARCHAR(20) DEFAULT 'text',  -- text | image | file | payslip | system
                file_url TEXT,
                file_name TEXT,
                file_size INTEGER,
                is_read BOOLEAN DEFAULT FALSE,
                sent_at TIMESTAMP DEFAULT NOW()
            )
        """)
        c.execute("ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS message_type VARCHAR(20) DEFAULT 'text'")
        c.execute("ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS file_url TEXT")
        c.execute("ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS file_name TEXT")
        c.execute("ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS file_size INTEGER")

        # Kênh "Thông báo chung" — toàn công ty, khóa (không rời được)
        c.execute("""
            INSERT INTO chat_rooms (room_name, room_type, is_locked, created_by, created_at, updated_at)
            SELECT '📢 Thông báo chung', 'broadcast', TRUE, 0, NOW(), NOW()
            WHERE NOT EXISTS (SELECT 1 FROM chat_rooms WHERE room_type = 'broadcast' AND deleted_at IS NULL)
        """)
        c.execute("SELECT id FROM chat_rooms WHERE room_type = 'broadcast' AND deleted_at IS NULL LIMIT 1")
        row = c.fetchone()
        broadcast_id = row[0] if row else None
        if broadcast_id:
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
                SELECT %s, nv.id, NOW(), NOW()
                FROM nhan_vien nv
                WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
                AND NOT EXISTS (
                    SELECT 1 FROM chat_participants cp WHERE cp.room_id = %s AND cp.user_id = nv.id
                )
            """, (broadcast_id, broadcast_id))

        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi khởi tạo bảng chat: {e}")
        return False


def sync_department_rooms():
    """Đảm bảo mỗi phòng ban trong PHONG_BAN_THU_TU có đúng 1 group mặc định,
    và thành viên của group đó luôn khớp với nhan_vien.phong_ban_lam_viec hiện tại.
    An toàn khi gọi nhiều lần (idempotent). Nên gọi lại ngay sau khi 1 Quyết định
    điều chuyển phòng ban được lưu, để group chat cập nhật tức thì thay vì chờ
    lượt render kế tiếp."""
    try:
        db, c = _conn()
        for ten_pb in PHONG_BAN_THU_TU:
            c.execute("""
                SELECT id FROM chat_rooms
                WHERE room_type = 'department' AND phong_ban = %s AND deleted_at IS NULL
            """, (ten_pb,))
            row = c.fetchone()
            if row:
                room_id = row["id"]
            else:
                c.execute("""
                    INSERT INTO chat_rooms (room_name, room_type, phong_ban, is_locked, created_by, updated_at)
                    VALUES (%s, 'department', %s, TRUE, 0, NOW()) RETURNING id
                """, (ten_pb, ten_pb))
                room_id = c.fetchone()["id"]

            # Nhân viên hiện đang thuộc phòng ban này (đang làm/thử việc)
            c.execute("""
                SELECT id FROM nhan_vien
                WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND phong_ban_lam_viec = %s
            """, (ten_pb,))
            dung_id = {r["id"] for r in c.fetchall()}

            c.execute("SELECT user_id FROM chat_participants WHERE room_id = %s", (room_id,))
            hien_co_id = {r["user_id"] for r in c.fetchall()}

            them_moi = dung_id - hien_co_id
            for uid in them_moi:
                c.execute("""
                    INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
                    VALUES (%s, %s, NOW(), NOW()) ON CONFLICT (room_id, user_id) DO NOTHING
                """, (room_id, uid))

            can_xoa = hien_co_id - dung_id
            for uid in can_xoa:
                c.execute("DELETE FROM chat_participants WHERE room_id = %s AND user_id = %s", (room_id, uid))

        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đồng bộ group phòng ban: {e}")
        return False


# ============================================================================
# 5) DATA LAYER
# ============================================================================
def get_current_nv(user_id):
    """Lấy thông tin nhân viên hiện tại (chuc_vu, phong_ban_lam_viec...) để tính quyền.
    Luôn đọc trực tiếp từ DB (không cache lâu) để phản ánh đúng ngay sau khi bị điều chuyển."""
    try:
        db, c = _conn()
        c.execute("""
            SELECT id, ma_nv, ho_ten, chuc_vu, phong_ban_lam_viec
            FROM nhan_vien WHERE id = %s
        """, (user_id,))
        row = c.fetchone()
        db.close()
        return dict(row) if row else None
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy thông tin nhân viên: {e}")
        return None


def get_all_employees(exclude_id=None):
    try:
        db, c = _conn()
        c.execute("""
            SELECT id, ma_nv, ho_ten, dien_thoai, phong_ban_lam_viec, chuc_vu
            FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            ORDER BY ho_ten ASC
        """)
        rows = c.fetchall()
        db.close()
        return [r for r in rows if r["id"] != exclude_id] if exclude_id else rows
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy danh sách nhân viên: {e}")
        return []


def search_employees(keyword, exclude_id=None):
    try:
        db, c = _conn()
        kw = f"%{keyword}%"
        c.execute("""
            SELECT id, ma_nv, ho_ten, dien_thoai, phong_ban_lam_viec, chuc_vu
            FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND (ho_ten ILIKE %s OR ma_nv ILIKE %s) ORDER BY ho_ten ASC LIMIT 30
        """, (kw, kw))
        rows = c.fetchall()
        db.close()
        return [r for r in rows if r["id"] != exclude_id] if exclude_id else rows
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi tìm nhân viên: {e}")
        return []


def get_user_chat_rooms(user_id):
    try:
        db, c = _conn()
        c.execute("""
            SELECT r.* FROM chat_rooms r
            JOIN chat_participants p ON r.id = p.room_id
            WHERE p.user_id = %s AND r.deleted_at IS NULL
            ORDER BY r.updated_at DESC
        """, (user_id,))
        rows = c.fetchall()
        db.close()
        return rows
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy danh sách phòng: {e}")
        return []


def get_room_participants(room_id):
    try:
        db, c = _conn()
        c.execute("""
            SELECT nv.id, nv.ma_nv, nv.ho_ten, nv.dien_thoai, nv.phong_ban_lam_viec, nv.chuc_vu
            FROM chat_participants p JOIN nhan_vien nv ON p.user_id = nv.id
            WHERE p.room_id = %s ORDER BY nv.ho_ten
        """, (room_id,))
        rows = c.fetchall()
        db.close()
        return rows
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy thành viên phòng: {e}")
        return []


def get_room_display_name(room, user_id):
    rt = room.get("room_type", "private")
    if rt == "broadcast":
        return "📢 Thông báo chung"
    if rt == "department":
        return f"🏢 {room.get('phong_ban') or room.get('room_name')}"
    if rt == "group":
        return f"👥 {room.get('room_name') or 'Nhóm chat'}"
    participants = get_room_participants(room["id"])
    for p in participants:
        if p["id"] != user_id:
            return p["ho_ten"]
    return "Phòng chat"


def get_room_last_message(room_id):
    try:
        db, c = _conn()
        c.execute("""
            SELECT m.*, nv.ho_ten as sender_name FROM chat_messages m
            LEFT JOIN nhan_vien nv ON m.sender_id = nv.id
            WHERE m.room_id = %s ORDER BY m.sent_at DESC LIMIT 1
        """, (room_id,))
        row = c.fetchone()
        db.close()
        return row
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy tin nhắn cuối: {e}")
        return None


def get_room_unread_count(room_id, user_id):
    try:
        db, c = _conn(dict_cursor=False)
        c.execute("""
            SELECT COUNT(*) FROM chat_messages m
            WHERE m.room_id = %s AND m.sender_id != %s
            AND m.sent_at > COALESCE(
                (SELECT last_read_at FROM chat_participants WHERE room_id = %s AND user_id = %s), '1970-01-01'
            )
        """, (room_id, user_id, room_id, user_id))
        n = c.fetchone()[0]
        db.close()
        return n
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đếm tin chưa đọc: {e}")
        return 0


def get_room_messages(room_id, limit=300):
    try:
        db, c = _conn()
        c.execute("""
            SELECT m.*, nv.ho_ten as sender_name FROM chat_messages m
            LEFT JOIN nhan_vien nv ON m.sender_id = nv.id
            WHERE m.room_id = %s ORDER BY m.sent_at ASC LIMIT %s
        """, (room_id, limit))
        rows = c.fetchall()
        db.close()
        return rows
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi lấy tin nhắn: {e}")
        return []


def mark_messages_as_read(room_id, user_id):
    try:
        db, c = _conn(dict_cursor=False)
        c.execute("""
            UPDATE chat_messages SET is_read = TRUE
            WHERE room_id = %s AND sender_id != %s AND is_read = FALSE
        """, (room_id, user_id))
        c.execute("""
            UPDATE chat_participants SET last_read_at = NOW() WHERE room_id = %s AND user_id = %s
        """, (room_id, user_id))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi đánh dấu đã đọc: {e}")
        return False


def create_private_room(user1_id, user2_id):
    try:
        db, c = _conn(dict_cursor=False)
        c.execute("""
            SELECT r.id FROM chat_rooms r
            JOIN chat_participants p1 ON r.id = p1.room_id
            JOIN chat_participants p2 ON r.id = p2.room_id
            WHERE r.room_type = 'private' AND p1.user_id = %s AND p2.user_id = %s AND r.deleted_at IS NULL
        """, (user1_id, user2_id))
        existing = c.fetchone()
        if existing:
            db.close()
            return existing[0]

        c.execute("""
            INSERT INTO chat_rooms (room_name, room_type, created_by, updated_at)
            VALUES (%s, 'private', %s, NOW()) RETURNING id
        """, (f"Chat {user1_id}-{user2_id}", user1_id))
        room_id = c.fetchone()[0]
        c.execute("""
            INSERT INTO chat_participants (room_id, user_id) VALUES (%s, %s), (%s, %s)
        """, (room_id, user1_id, room_id, user2_id))
        db.commit()
        db.close()
        return room_id
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi tạo phòng 1:1: {e}")
        return None


def create_group_room(room_name, creator_id, member_ids):
    """Tạo group tùy chỉnh (room_type='group'). KHÔNG gọi hàm này trực tiếp từ UI
    mà không kiểm tra co_quyen_tao_nhom() trước — hàm này không tự kiểm tra quyền."""
    try:
        db, c = _conn(dict_cursor=False)
        c.execute("""
            INSERT INTO chat_rooms (room_name, room_type, created_by, updated_at)
            VALUES (%s, 'group', %s, NOW()) RETURNING id
        """, (room_name, creator_id))
        room_id = c.fetchone()[0]
        for uid in {creator_id, *member_ids}:
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id) VALUES (%s, %s)
                ON CONFLICT (room_id, user_id) DO NOTHING
            """, (room_id, uid))
        db.commit()
        db.close()
        return room_id
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi tạo nhóm: {e}")
        return None


def add_participants_to_room(room_id, user_ids):
    try:
        db, c = _conn(dict_cursor=False)
        for uid in user_ids:
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
                VALUES (%s, %s, NOW(), NOW()) ON CONFLICT (room_id, user_id) DO NOTHING
            """, (room_id, uid))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi thêm thành viên: {e}")
        return False


def remove_participant_from_room(room_id, user_id):
    try:
        db, c = _conn(dict_cursor=False)
        c.execute("DELETE FROM chat_participants WHERE room_id = %s AND user_id = %s", (room_id, user_id))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi xóa thành viên: {e}")
        return False


def roi_nhom(room, user_id):
    """Rời group — CHỈ áp dụng cho room_type='group' (nhóm tùy chỉnh).
    Không cho rời 'department'/'broadcast' (is_locked) hay 'private'."""
    if room.get("room_type") != "group" or room.get("is_locked"):
        return False
    return remove_participant_from_room(room["id"], user_id)


def delete_room(room_id):
    try:
        db, c = _conn(dict_cursor=False)
        c.execute("UPDATE chat_rooms SET deleted_at = NOW() WHERE id = %s", (room_id,))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi xóa phòng: {e}")
        return False


def send_message(room_id, sender_id, content, message_type="text", file_url=None, file_name=None, file_size=None):
    try:
        db, c = _conn(dict_cursor=False)
        c.execute("""
            INSERT INTO chat_messages (room_id, sender_id, content, message_type, file_url, file_name, file_size)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (room_id, sender_id, content, message_type, file_url, file_name, file_size))
        c.execute("UPDATE chat_rooms SET updated_at = NOW() WHERE id = %s", (room_id,))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"[chat_noi_bo] Lỗi gửi tin nhắn: {e}")
        return False


def send_image_message(room_id, sender_id, file_url, file_name, caption=""):
    return send_message(room_id, sender_id, caption, "image", file_url, file_name)


def send_file_message(room_id, sender_id, file_url, file_name, file_size):
    return send_message(room_id, sender_id, file_name, "file", file_url, file_name, file_size)


def send_payslip_message(room_id, sender_id, employee_name, month, year, salary_data):
    content = f"""📄 PHIẾU LƯƠNG THÁNG {month}/{year}

👤 Nhân viên: {employee_name}
{'-'*30}
Lương cơ bản: {salary_data.get('luong_co_ban', 0):,.0f} VNĐ
Phụ cấp chức vụ: {salary_data.get('phu_cap_chuc_vu', 0):,.0f} VNĐ
Phụ cấp thâm niên: {salary_data.get('phu_cap_tnvk', 0):,.0f} VNĐ
Phụ cấp trách nhiệm: {salary_data.get('phu_cap_tnn', 0):,.0f} VNĐ
{'-'*30}
Tổng thu nhập: {salary_data.get('tong', 0):,.0f} VNĐ
{'-'*30}
BHXH (8%): {salary_data.get('bhxh', 0):,.0f} VNĐ
BHYT (1.5%): {salary_data.get('bhyt', 0):,.0f} VNĐ
BHTN (1%): {salary_data.get('bhtn', 0):,.0f} VNĐ
{'-'*30}
Thực nhận: {salary_data.get('thuc_nhan', 0):,.0f} VNĐ

📅 Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}"""
    return send_message(room_id, sender_id, content, "payslip")


def export_room_transcript(room_id, room_display_name):
    """Xuất toàn bộ lịch sử chat của 1 phòng ra văn bản thuần (.txt) để tải về —
    phục vụ nút 'Lưu hộp thoại chat'."""
    msgs = get_room_messages(room_id, limit=5000)
    lines = [f"LỊCH SỬ CHAT — {room_display_name}",
             f"Xuất lúc: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "=" * 50, ""]
    for m in msgs:
        ts = m["sent_at"].strftime("%d/%m/%Y %H:%M") if m.get("sent_at") else ""
        ten = m.get("sender_name") or "Người dùng"
        noi_dung = m.get("content") or ""
        if m.get("message_type") == "image":
            noi_dung = f"[Hình ảnh: {m.get('file_name', '')}] {noi_dung}"
        elif m.get("message_type") == "file":
            noi_dung = f"[Tệp đính kèm: {m.get('file_name', '')}]"
        lines.append(f"[{ts}] {ten}: {noi_dung}")
    return "\n".join(lines)


# ============================================================================
# 6) UPLOAD FILE (Supabase Storage) — giữ nguyên logic gốc từ chat_utils.py
# ============================================================================
def sanitize_filename(filename):
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
        url, key = tenant.get("supabase_url"), tenant.get("supabase_key")
    else:
        import os
        from dotenv import load_dotenv
        load_dotenv()
        url, key = os.getenv("SUPABASE_URL"), os.getenv("SUPABASE_KEY")
    if not url or not key:
        return None
    return create_client(url, key)


def upload_chat_file(file_bytes, filename, content_type):
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
def _mau_avatar(ten):
    h = int(hashlib.md5((ten or "?").encode("utf-8")).hexdigest(), 16)
    return MAU_AVATAR[h % len(MAU_AVATAR)]


def _initials(ten):
    parts = (ten or "?").strip().split()
    return (parts[-1][0] if parts else "?").upper()


def _avatar_html(ten, size=38):
    return (f'<div class="cnb-avatar" style="width:{size}px;height:{size}px;'
            f'background:{_mau_avatar(ten)};font-size:{size*0.42:.0f}px;">{_initials(ten)}</div>')


def _thoi_gian(dt):
    if not dt:
        return ""
    now = datetime.now()
    if dt.date() == now.date():
        return dt.strftime("%H:%M")
    if (now.date() - dt.date()).days < 7:
        return dt.strftime("%a %H:%M")
    return dt.strftime("%d/%m %H:%M")


def _css():
    st.markdown("""
    <style>
    .cnb-wrap { display:flex; height:640px; border-radius:16px; overflow:hidden;
        background:#fff; box-shadow:0 4px 24px rgba(15,23,42,0.08); border:1px solid #e8ecf1; }
    .cnb-side { width:320px; min-width:280px; background:#fafbfc; border-right:1px solid #e8ecf1;
        display:flex; flex-direction:column; flex-shrink:0; }
    .cnb-side-scroll { flex:1; overflow-y:auto; padding:6px; }
    .cnb-main { flex:1; display:flex; flex-direction:column; background:#f0f2f5; overflow:hidden; }
    .cnb-header { padding:14px 20px; background:#fff; border-bottom:1px solid #e8ecf1;
        display:flex; align-items:center; gap:12px; flex-shrink:0; }
    .cnb-header .cnb-title { font-weight:700; color:#0f172a; font-size:16px; margin:0; }
    .cnb-header .cnb-sub { font-size:12px; color:#94a3b8; margin:0; }
    .cnb-avatar { border-radius:50%; color:#fff; display:flex; align-items:center;
        justify-content:center; font-weight:700; flex-shrink:0; }
    .cnb-room-item { display:flex; gap:10px; align-items:center; padding:10px 12px; border-radius:12px;
        cursor:pointer; margin-bottom:2px; }
    .cnb-room-item:hover { background:#eef1f5; }
    .cnb-room-item.active { background:#e7f0ff; }
    .cnb-room-body { flex:1; min-width:0; }
    .cnb-room-name { font-size:13.5px; font-weight:600; color:#0f172a; white-space:nowrap;
        overflow:hidden; text-overflow:ellipsis; display:flex; align-items:center; gap:4px; }
    .cnb-room-last { font-size:12px; color:#94a3b8; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .cnb-room-meta { display:flex; flex-direction:column; align-items:flex-end; gap:4px; }
    .cnb-room-time { font-size:10.5px; color:#94a3b8; }
    .cnb-badge { background:#ef4444; color:#fff; font-size:10.5px; font-weight:700; border-radius:10px;
        padding:1px 6px; min-width:16px; text-align:center; }
    .cnb-lock { font-size:11px; color:#94a3b8; }
    .cnb-messages { flex:1; overflow-y:auto; padding:18px 22px; display:flex; flex-direction:column; gap:2px; }
    .cnb-msg-row { display:flex; gap:8px; margin-bottom:10px; align-items:flex-end; }
    .cnb-msg-row.self { flex-direction:row-reverse; }
    .cnb-msg-col { display:flex; flex-direction:column; max-width:62%; }
    .cnb-msg-row.self .cnb-msg-col { align-items:flex-end; }
    .cnb-msg-sender { font-size:11.5px; font-weight:600; color:#64748b; margin:0 2px 3px; }
    .cnb-bubble { padding:9px 14px; border-radius:16px; font-size:14px; line-height:1.5;
        box-shadow:0 1px 2px rgba(0,0,0,0.05); word-wrap:break-word; white-space:pre-wrap; }
    .cnb-msg-row.self .cnb-bubble { background:linear-gradient(135deg,#0084ff,#0066cc); color:#fff;
        border-bottom-right-radius:4px; }
    .cnb-msg-row.other .cnb-bubble { background:#fff; color:#0f172a; border:1px solid #e8ecf1;
        border-bottom-left-radius:4px; }
    .cnb-msg-time { font-size:10px; color:#a3adba; margin:3px 2px 0; }
    .cnb-payslip { background:#f0fdf4; border:1px solid #bbf7d0; border-radius:14px; padding:12px 16px;
        font-size:13px; white-space:pre-wrap; font-family:ui-monospace,monospace; }
    .cnb-file { display:flex; align-items:center; gap:10px; padding:10px 14px; background:rgba(0,0,0,0.04);
        border-radius:12px; text-decoration:none; color:#0f172a; }
    .cnb-empty { flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center;
        color:#94a3b8; gap:8px; }
    .cnb-empty .icon { font-size:52px; }
    </style>
    """, unsafe_allow_html=True)


def _render_room_item(room, nv, is_active):
    display = get_room_display_name(room, nv["id"])
    last = get_room_last_message(room["id"])
    last_txt = "Chưa có tin nhắn nào"
    if last:
        prefix = "Bạn: " if last["sender_id"] == nv["id"] else f"{(last.get('sender_name') or '').split(' ')[-1]}: "
        body = last.get("content") or ""
        if last.get("message_type") == "image":
            body = "[Hình ảnh]"
        elif last.get("message_type") == "file":
            body = f"[Tệp] {last.get('file_name','')}"
        elif last.get("message_type") == "payslip":
            body = "[Phiếu lương]"
        last_txt = (prefix + body)[:42]
    unread = get_room_unread_count(room["id"], nv["id"])
    lock_icon = "🔒 " if room.get("is_locked") else ""
    ts = _thoi_gian(last["sent_at"]) if last else ""

    cols = st.columns([0.16, 0.68, 0.16])
    with cols[0]:
        st.markdown(_avatar_html(display.split(" ", 1)[-1], 34), unsafe_allow_html=True)
    with cols[1]:
        if st.button(display, key=f"room_btn_{room['id']}", use_container_width=True,
                     type="primary" if is_active else "secondary"):
            st.session_state["cnb_room_id"] = room["id"]
            mark_messages_as_read(room["id"], nv["id"])
            st.rerun()
        st.caption(f"{lock_icon}{last_txt}")
    with cols[2]:
        st.caption(ts)
        if unread:
            st.markdown(f'<span class="cnb-badge">{unread}</span>', unsafe_allow_html=True)


@st.dialog("➕ Tạo phòng mới")
def _dialog_tao_phong(nv):
    co_the_tao_nhom = co_quyen_tao_nhom(nv)
    tabs = st.tabs(["💬 Chat 1:1"] + (["👥 Tạo nhóm mới"] if co_the_tao_nhom else []))

    with tabs[0]:
        kw = st.text_input("🔍 Tìm theo tên hoặc mã NV", key="cnb_search_private")
        emps = search_employees(kw, exclude_id=nv["id"]) if kw else get_all_employees(exclude_id=nv["id"])
        for e in emps[:25]:
            c1, c2 = st.columns([0.75, 0.25])
            with c1:
                st.markdown(f"**{e['ho_ten']}** · {e.get('phong_ban_lam_viec') or ''}")
            with c2:
                if st.button("Nhắn tin", key=f"priv_{e['id']}", use_container_width=True):
                    rid = create_private_room(nv["id"], e["id"])
                    if rid:
                        st.session_state["cnb_room_id"] = rid
                        st.rerun()

    if co_the_tao_nhom:
        with tabs[1]:
            pham_vi = ("toàn công ty" if la_lanh_dao_cao_cap(nv.get("phong_ban_lam_viec"))
                       else "liên phòng ban")
            st.caption(f"✅ Bạn có quyền tạo nhóm với thành viên {pham_vi}.")
            ten_nhom = st.text_input("Tên nhóm", key="cnb_group_name", placeholder="VD: Dự án cảng Hòn La 2026")
            kw2 = st.text_input("🔍 Tìm & chọn thành viên", key="cnb_search_group")
            emps2 = search_employees(kw2, exclude_id=nv["id"]) if kw2 else get_all_employees(exclude_id=nv["id"])
            options = {f"{e['ho_ten']} — {e.get('phong_ban_lam_viec') or ''}": e["id"] for e in emps2}
            chon = st.multiselect("Thành viên", list(options.keys()), key="cnb_group_members")
            if st.button("Tạo nhóm", type="primary", disabled=not (ten_nhom.strip() and chon)):
                member_ids = [options[k] for k in chon]
                rid = create_group_room(ten_nhom.strip(), nv["id"], member_ids)
                if rid:
                    st.session_state["cnb_room_id"] = rid
                    st.success("Đã tạo nhóm!")
                    st.rerun()
    else:
        st.info("ℹ️ Chỉ Trưởng phòng trở lên mới được tạo nhóm liên phòng ban. "
                "Bạn vẫn có thể chat 1:1 hoặc dùng các group phòng ban đã có sẵn.")


def _render_sidebar(nv, rooms, active_room_id):
    top1, top2 = st.columns([0.75, 0.25])
    with top1:
        if st.button("➕ Tạo phòng mới", use_container_width=True, type="primary"):
            _dialog_tao_phong(nv)
    with top2:
        if st.button("🔄", use_container_width=True, help="Làm mới danh sách phòng"):
            st.rerun()

    tim = st.text_input("🔎 Tìm phòng chat...", key="cnb_search_room", label_visibility="collapsed",
                         placeholder="🔎 Tìm phòng chat...")

    st.markdown('<div class="cnb-side-scroll">', unsafe_allow_html=True)
    if not rooms:
        st.caption("Chưa có phòng chat nào.")
    else:
        # Sắp xếp: broadcast/department lên trước theo hoạt động gần nhất, giữ nguyên order từ query
        for room in rooms:
            display = get_room_display_name(room, nv["id"])
            if tim and tim.strip().lower() not in display.lower():
                continue
            _render_room_item(room, nv, is_active=(room["id"] == active_room_id))
    st.markdown("</div>", unsafe_allow_html=True)


def _render_message(msg, is_self):
    wrapper_cls = "self" if is_self else "other"
    sender = msg.get("sender_name") or "Người dùng"
    time_str = _thoi_gian(msg["sent_at"])

    st.markdown(f'<div class="cnb-msg-row {wrapper_cls}">', unsafe_allow_html=True)
    cols = st.columns([0.08, 0.92]) if not is_self else st.columns([0.92, 0.08])
    body_col = cols[1] if not is_self else cols[0]
    avatar_col = cols[0] if not is_self else cols[1]

    with avatar_col:
        st.markdown(_avatar_html(sender, 30), unsafe_allow_html=True)

    with body_col:
        if not is_self:
            st.markdown(f'<div class="cnb-msg-sender">{sender}</div>', unsafe_allow_html=True)

        mtype = msg.get("message_type", "text")
        if mtype == "image" and msg.get("file_url"):
            file_bytes = get_chat_file_bytes(msg["file_url"])
            if file_bytes:
                st.image(file_bytes, width=280)
            if msg.get("content"):
                st.markdown(f'<div class="cnb-bubble">{msg["content"]}</div>', unsafe_allow_html=True)
        elif mtype == "file" and msg.get("file_url"):
            st.markdown(f'<div class="cnb-file">📎 <b>{msg.get("file_name","")}</b></div>', unsafe_allow_html=True)
            file_bytes = get_chat_file_bytes(msg["file_url"])
            if file_bytes:
                st.download_button("⬇️ Tải xuống", file_bytes, file_name=msg.get("file_name", "file"),
                                    key=f"dl_{msg['id']}")
        elif mtype == "payslip":
            st.markdown(f'<div class="cnb-payslip">{msg.get("content","")}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="cnb-bubble">{msg.get("content","")}</div>', unsafe_allow_html=True)

        st.markdown(f'<div class="cnb-msg-time">{time_str}</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def _render_empty_state():
    st.markdown("""
    <div class="cnb-empty">
        <div class="icon">💬</div>
        <h3 style="color:#334155;margin:0;">Chọn một phòng chat</h3>
        <div>Chọn phòng từ danh sách bên trái hoặc tạo phòng mới</div>
        <div style="font-size:13px;">💡 Bấm "➕ Tạo phòng mới" để bắt đầu trò chuyện</div>
    </div>
    """, unsafe_allow_html=True)


def _render_chat_window(nv, room):
    participants = get_room_participants(room["id"])
    display = get_room_display_name(room, nv["id"])

    h1, h2, h3 = st.columns([0.62, 0.2, 0.18])
    with h1:
        st.markdown(f"### {display}")
        st.caption(f"{len(participants)} thành viên" + ("  ·  🔒 Không thể rời nhóm này" if room.get("is_locked") else ""))
    with h2:
        transcript = export_room_transcript(room["id"], display)
        st.download_button("💾 Lưu hộp thoại", transcript.encode("utf-8"),
                            file_name=f"chat_{display}_{datetime.now().strftime('%Y%m%d')}.txt",
                            key=f"export_{room['id']}", use_container_width=True)
    with h3:
        with st.popover("⚙️ Tùy chọn", use_container_width=True):
            st.markdown(f"**👥 Thành viên ({len(participants)})**")
            for p in participants:
                st.caption(f"👤 {p['ho_ten']} — {p.get('phong_ban_lam_viec','') or ''}")

            if co_quyen_quan_ly_nhom(nv, room):
                st.divider()
                st.caption("Thêm thành viên (chỉ người quản lý nhóm)")
                current_ids = {p["id"] for p in participants}
                candidates = {f"{e['ho_ten']} — {e.get('phong_ban_lam_viec','')}": e["id"]
                              for e in get_all_employees() if e["id"] not in current_ids}
                chon_them = st.multiselect("Chọn thành viên mới", list(candidates.keys()), key=f"add_mem_{room['id']}")
                if st.button("➕ Thêm vào nhóm", key=f"btn_add_mem_{room['id']}"):
                    add_participants_to_room(room["id"], [candidates[k] for k in chon_them])
                    st.rerun()

            if room.get("room_type") == "group" and not room.get("is_locked"):
                st.divider()
                if st.button("🚪 Rời nhóm", key=f"leave_{room['id']}"):
                    if roi_nhom(room, nv["id"]):
                        st.session_state["cnb_room_id"] = None
                        st.rerun()

    st.markdown('<div class="cnb-messages">', unsafe_allow_html=True)
    msgs = get_room_messages(room["id"])
    if not msgs:
        st.caption("Chưa có tin nhắn nào — hãy bắt đầu cuộc trò chuyện!")
    else:
        for m in msgs:
            _render_message(m, is_self=(m["sender_id"] == nv["id"]))
    st.markdown("</div>", unsafe_allow_html=True)

    mark_messages_as_read(room["id"], nv["id"])

    with st.form(key=f"send_form_{room['id']}", clear_on_submit=True):
        c1, c2, c3 = st.columns([0.72, 0.14, 0.14])
        with c1:
            noi_dung = st.text_input("Tin nhắn", key=f"msg_input_{room['id']}",
                                      label_visibility="collapsed", placeholder="Nhập tin nhắn...")
        with c2:
            file_dinh_kem = st.file_uploader("📎", key=f"file_{room['id']}", label_visibility="collapsed")
        with c3:
            gui = st.form_submit_button("Gửi ➤", use_container_width=True, type="primary")

        if gui and (noi_dung.strip() or file_dinh_kem):
            if file_dinh_kem:
                file_bytes = file_dinh_kem.read()
                content_type = file_dinh_kem.type or "application/octet-stream"
                path = upload_chat_file(file_bytes, file_dinh_kem.name, content_type)
                if path:
                    if content_type.startswith("image/"):
                        send_image_message(room["id"], nv["id"], path, file_dinh_kem.name, noi_dung.strip())
                    else:
                        send_file_message(room["id"], nv["id"], path, file_dinh_kem.name, len(file_bytes))
                else:
                    st.error("❌ Tải file lên thất bại, thử lại sau.")
            elif noi_dung.strip():
                send_message(room["id"], nv["id"], noi_dung.strip())
            st.rerun()


def render():
    """Entry point — gọi từ app.py: `chat_noi_bo.render()`"""
    if "nhan_vien_id" not in st.session_state or not st.session_state.nhan_vien_id:
        st.warning("⚠️ Vui lòng đăng nhập để sử dụng Chat nội bộ.")
        return

    # Khởi tạo & đồng bộ 1 LẦN / phiên làm việc (tránh query lặp lại mỗi lần rerun)
    if not st.session_state.get("_cnb_da_khoi_tao"):
        init_chat_tables()
        sync_department_rooms()
        st.session_state["_cnb_da_khoi_tao"] = True

    user_id = st.session_state.nhan_vien_id
    nv = get_current_nv(user_id)
    if not nv:
        st.error("Không tìm thấy thông tin nhân viên hiện tại.")
        return

    st.title("💬 Chat nội bộ")
    _css()

    rooms = get_user_chat_rooms(user_id)
    active_room_id = st.session_state.get("cnb_room_id")
    active_room = next((r for r in rooms if r["id"] == active_room_id), None)

    st.markdown('<div class="cnb-wrap">', unsafe_allow_html=True)
    col_side, col_main = st.columns([0.32, 0.68])
    with col_side:
        st.markdown('<div class="cnb-side">', unsafe_allow_html=True)
        _render_sidebar(nv, rooms, active_room_id)
        st.markdown("</div>", unsafe_allow_html=True)
    with col_main:
        st.markdown('<div class="cnb-main">', unsafe_allow_html=True)
        if active_room:
            _render_chat_window(nv, active_room)
        else:
            _render_empty_state()
        st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)