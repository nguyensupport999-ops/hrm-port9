# chat_utils.py
"""
Module xử lý logic cho Chat nội bộ - Phiên bản nâng cấp
"""
import psycopg2
import psycopg2.extras
import streamlit as st
from datetime import datetime
import re
import unicodedata

# ========== HÀM HIỆN CÓ (GIỮ NGUYÊN) ==========

def get_user_chat_rooms(user_id):
    """Lấy danh sách phòng chat của một người dùng"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT r.*, 
                   (SELECT COUNT(*) FROM chat_messages 
                    WHERE room_id = r.id AND is_read = FALSE AND sender_id != %s) as unread_count
            FROM chat_rooms r
            JOIN chat_participants p ON r.id = p.room_id
            WHERE p.user_id = %s
            ORDER BY r.updated_at DESC
        """, (user_id, user_id))
        rooms = c.fetchall()
        db.close()
        return rooms
    except Exception as e:
        print(f"Lỗi lấy danh sách phòng: {e}")
        return []

def get_room_messages(room_id):
    """Lấy tin nhắn của một phòng"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT m.*, nv.ho_ten as sender_name
            FROM chat_messages m
            LEFT JOIN nhan_vien nv ON m.sender_id = nv.id
            WHERE m.room_id = %s
            ORDER BY m.sent_at ASC
        """, (room_id,))
        messages = c.fetchall()
        db.close()
        return messages
    except Exception as e:
        print(f"Lỗi lấy tin nhắn: {e}")
        return []

def send_message(room_id, sender_id, content, message_type='text', file_url=None, file_name=None, file_size=None):
    """Gửi tin nhắn mới (mở rộng hỗ trợ file)"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            INSERT INTO chat_messages (room_id, sender_id, content, message_type, file_url, file_name, file_size)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (room_id, sender_id, content, message_type, file_url, file_name, file_size))
        # Cập nhật thời gian updated_at cho phòng
        c.execute("UPDATE chat_rooms SET updated_at = NOW() WHERE id = %s", (room_id,))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi gửi tin nhắn: {e}")
        return False

def create_private_room(user1_id, user2_id):
    """Tạo phòng chat 1-1 giữa hai người"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        # Kiểm tra xem đã có phòng chưa
        c.execute("""
            SELECT r.id FROM chat_rooms r
            JOIN chat_participants p1 ON r.id = p1.room_id
            JOIN chat_participants p2 ON r.id = p2.room_id
            WHERE r.room_type = 'private' AND p1.user_id = %s AND p2.user_id = %s
            AND r.deleted_at IS NULL
        """, (user1_id, user2_id))
        existing = c.fetchone()
        if existing:
            return existing[0]
        
        # Tạo phòng mới
        c.execute("""
            INSERT INTO chat_rooms (room_name, room_type, created_by, updated_at) 
            VALUES (%s, 'private', %s, NOW()) RETURNING id
        """, (f"Chat {user1_id}-{user2_id}", user1_id))
        room_id = c.fetchone()[0]
        
        # Thêm 2 thành viên
        c.execute("""
            INSERT INTO chat_participants (room_id, user_id) 
            VALUES (%s, %s), (%s, %s)
        """, (room_id, user1_id, room_id, user2_id))
        
        db.commit()
        db.close()
        return room_id
    except Exception as e:
        print(f"Lỗi tạo phòng chat: {e}")
        return None

def create_group_room(room_name, creator_id, member_ids):
    """Tạo phòng chat nhóm"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        
        # Tạo phòng nhóm
        c.execute("""
            INSERT INTO chat_rooms (room_name, room_type, created_by, updated_at) 
            VALUES (%s, 'group', %s, NOW()) RETURNING id
        """, (room_name, creator_id))
        room_id = c.fetchone()[0]
        
        # Thêm các thành viên (bao gồm cả người tạo)
        for user_id in [creator_id] + member_ids:
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id) 
                VALUES (%s, %s)
                ON CONFLICT (room_id, user_id) DO NOTHING
            """, (room_id, user_id))
        
        db.commit()
        db.close()
        return room_id
    except Exception as e:
        print(f"Lỗi tạo phòng nhóm: {e}")
        return None

def get_room_participants(room_id):
    """Lấy danh sách thành viên trong phòng"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT nv.id, nv.ma_nv, nv.ho_ten, nv.dien_thoai, nv.phong_ban_lam_viec
            FROM chat_participants p
            JOIN nhan_vien nv ON p.user_id = nv.id
            WHERE p.room_id = %s
        """, (room_id,))
        members = c.fetchall()
        db.close()
        return members
    except Exception as e:
        print(f"Lỗi lấy danh sách thành viên: {e}")
        return []

def mark_messages_as_read(room_id, user_id):
    """Đánh dấu tin nhắn đã đọc"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        # Cập nhật is_read
        c.execute("""
            UPDATE chat_messages 
            SET is_read = TRUE 
            WHERE room_id = %s AND sender_id != %s AND is_read = FALSE
        """, (room_id, user_id))
        # Cập nhật last_read_at
        c.execute("""
            UPDATE chat_participants 
            SET last_read_at = NOW()
            WHERE room_id = %s AND user_id = %s
        """, (room_id, user_id))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi đánh dấu đã đọc: {e}")
        return False

def get_all_employees():
    """Lấy danh sách tất cả nhân viên đang làm việc"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT id, ma_nv, ho_ten, dien_thoai, phong_ban_lam_viec
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            ORDER BY ho_ten ASC
        """)
        employees = c.fetchall()
        db.close()
        return employees
    except Exception as e:
        print(f"Lỗi lấy danh sách nhân viên: {e}")
        return []


# ========== HÀM MỚI (NÂNG CẤP) ==========

def init_chat_tables():
    """Khởi tạo các bảng chat nếu chưa có (nâng cấp thêm cột mới)"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        
        # Bảng chat_rooms
        c.execute("""
            CREATE TABLE IF NOT EXISTS chat_rooms (
                id SERIAL PRIMARY KEY,
                room_name VARCHAR(100),
                room_type VARCHAR(20) NOT NULL DEFAULT 'private',
                created_by INTEGER,
                created_at TIMESTAMP DEFAULT NOW(),
                updated_at TIMESTAMP DEFAULT NOW(),
                deleted_at TIMESTAMP
            )
        """)
        # Thêm cột mới nếu chưa có
        c.execute("ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP DEFAULT NOW()")
        c.execute("ALTER TABLE chat_rooms ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMP")
        
        # Bảng chat_participants
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
        
        # Bảng chat_messages
        c.execute("""
            CREATE TABLE IF NOT EXISTS chat_messages (
                id SERIAL PRIMARY KEY,
                room_id INTEGER NOT NULL,
                sender_id INTEGER NOT NULL,
                content TEXT,
                message_type VARCHAR(20) DEFAULT 'text',
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
        
        # Tạo phòng BROADCAST nếu chưa có
        c.execute("""
            INSERT INTO chat_rooms (room_name, room_type, created_by, created_at, updated_at)
            SELECT '📢 Thông báo chung', 'broadcast', 0, NOW(), NOW()
            WHERE NOT EXISTS (
                SELECT 1 FROM chat_rooms WHERE room_type = 'broadcast' AND deleted_at IS NULL
            )
        """)
        
        # Thêm tất cả nhân viên vào phòng broadcast
        c.execute("""
            INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
            SELECT 
                (SELECT id FROM chat_rooms WHERE room_type = 'broadcast' AND deleted_at IS NULL),
                nv.id,
                NOW(),
                NOW()
            FROM nhan_vien nv
            WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND NOT EXISTS (
                SELECT 1 FROM chat_participants cp
                WHERE cp.room_id = (SELECT id FROM chat_rooms WHERE room_type = 'broadcast' AND deleted_at IS NULL)
                AND cp.user_id = nv.id
            )
        """)
        
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi khởi tạo bảng chat: {e}")
        return False

def get_or_create_broadcast_room():
    """Lấy hoặc tạo phòng broadcast"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT id FROM chat_rooms WHERE room_type = 'broadcast' AND deleted_at IS NULL")
        result = c.fetchone()
        db.close()
        return result[0] if result else None
    except Exception as e:
        print(f"Lỗi lấy phòng broadcast: {e}")
        return None

def get_room_unread_count(room_id, user_id):
    """Lấy số tin nhắn chưa đọc trong phòng"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            SELECT COUNT(*) FROM chat_messages m
            WHERE m.room_id = %s 
            AND m.sender_id != %s 
            AND m.is_read = FALSE
            AND m.sent_at > COALESCE(
                (SELECT last_read_at FROM chat_participants 
                 WHERE room_id = %s AND user_id = %s),
                '1970-01-01'
            )
        """, (room_id, user_id, room_id, user_id))
        count = c.fetchone()[0]
        db.close()
        return count
    except Exception as e:
        print(f"Lỗi lấy số tin chưa đọc: {e}")
        return 0

def get_room_last_message(room_id):
    """Lấy tin nhắn cuối cùng của phòng"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT m.*, nv.ho_ten as sender_name
            FROM chat_messages m
            LEFT JOIN nhan_vien nv ON m.sender_id = nv.id
            WHERE m.room_id = %s
            ORDER BY m.sent_at DESC
            LIMIT 1
        """, (room_id,))
        result = c.fetchone()
        db.close()
        return result
    except Exception as e:
        print(f"Lỗi lấy tin nhắn cuối: {e}")
        return None

def get_room_display_name(room, user_id):
    """Lấy tên hiển thị của phòng chat"""
    room_type = room.get('room_type', 'private')
    if room_type == 'broadcast':
        return '📢 Thông báo chung'
    elif room_type == 'group':
        return room.get('room_name', 'Nhóm chat')
    else:  # private
        participants = get_room_participants(room['id'])
        for p in participants:
            if p['id'] != user_id:
                return p['ho_ten']
        return 'Phòng chat'

def send_image_message(room_id, sender_id, file_url, file_name, caption=''):
    """Gửi tin nhắn ảnh"""
    return send_message(room_id, sender_id, caption, 'image', file_url, file_name)

def send_file_message(room_id, sender_id, file_url, file_name, file_size):
    """Gửi tin nhắn file"""
    return send_message(room_id, sender_id, file_name, 'file', file_url, file_name, file_size)

def send_payslip_message(room_id, sender_id, employee_name, month, year, salary_data):
    """Gửi tin nhắn phiếu lương"""
    content = f"""📄 **PHIẾU LƯƠNG THÁNG {month}/{year}**
    
👤 Nhân viên: {employee_name}

📊 Chi tiết thu nhập:
{'-' * 30}
Lương cơ bản: {salary_data.get('luong_co_ban', 0):,.0f} VNĐ
Phụ cấp chức vụ: {salary_data.get('phu_cap_chuc_vu', 0):,.0f} VNĐ
Phụ cấp thâm niên: {salary_data.get('phu_cap_tnvk', 0):,.0f} VNĐ
Phụ cấp trách nhiệm: {salary_data.get('phu_cap_tnn', 0):,.0f} VNĐ
{'-' * 30}
Tổng thu nhập: {salary_data.get('tong', 0):,.0f} VNĐ

{'-' * 30}
📌 Các khoản khấu trừ:
BHXH (8%): {salary_data.get('bhxh', 0):,.0f} VNĐ
BHYT (1.5%): {salary_data.get('bhyt', 0):,.0f} VNĐ
BHTN (1%): {salary_data.get('bhtn', 0):,.0f} VNĐ
{'-' * 30}
Thực nhận: {salary_data.get('thuc_nhan', 0):,.0f} VNĐ

📅 Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}
"""
    return send_message(room_id, sender_id, content, 'payslip')

def search_employees(keyword):
    """Tìm kiếm nhân viên theo tên hoặc mã"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        keyword = f"%{keyword}%"
        c.execute("""
            SELECT id, ma_nv, ho_ten, dien_thoai, phong_ban_lam_viec
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND (ho_ten ILIKE %s OR ma_nv ILIKE %s)
            ORDER BY ho_ten ASC
            LIMIT 20
        """, (keyword, keyword))
        results = c.fetchall()
        db.close()
        return results
    except Exception as e:
        print(f"Lỗi tìm kiếm nhân viên: {e}")
        return []

def add_participants_to_room(room_id, user_ids):
    """Thêm thành viên vào phòng nhóm"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        for user_id in user_ids:
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
                VALUES (%s, %s, NOW(), NOW())
                ON CONFLICT (room_id, user_id) DO NOTHING
            """, (room_id, user_id))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi thêm thành viên: {e}")
        return False

def remove_participant_from_room(room_id, user_id):
    """Xóa thành viên khỏi phòng nhóm"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("DELETE FROM chat_participants WHERE room_id = %s AND user_id = %s", (room_id, user_id))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi xóa thành viên: {e}")
        return False

def delete_room(room_id):
    """Xóa mềm phòng chat"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("UPDATE chat_rooms SET deleted_at = NOW() WHERE id = %s", (room_id,))
        db.commit()
        db.close()
        return True
    except Exception as e:
        print(f"Lỗi xóa phòng: {e}")
        return False


# ========== HÀM UPLOAD HỖ TRỢ ==========

def sanitize_filename(filename):
    """Chuẩn hóa tên file"""
    # Bỏ dấu tiếng Việt
    normalized = unicodedata.normalize('NFD', filename)
    no_accent = ''.join(ch for ch in normalized if unicodedata.category(ch) != 'Mn')
    no_accent = no_accent.replace('đ', 'd').replace('Đ', 'D')
    # Thay khoảng trắng bằng '_'
    no_accent = re.sub(r'\s+', '_', no_accent)
    # Chỉ giữ ký tự an toàn
    safe = re.sub(r'[^A-Za-z0-9_.\-]', '', no_accent)
    return safe or 'file'

def upload_chat_file(file_bytes, filename, content_type):
    """Upload file lên Supabase Storage cho chat"""
    try:
        # Import từ app (cần đảm bảo các hàm này tồn tại)
        # Sử dụng st.session_state để lấy storage client
        from supabase import create_client
        
        tenant = st.session_state.get('tenant')
        if tenant:
            url, key = tenant.get('supabase_url'), tenant.get('supabase_key')
        else:
            # Fallback
            import os
            from dotenv import load_dotenv
            load_dotenv()
            url = os.getenv('SUPABASE_URL')
            key = os.getenv('SUPABASE_KEY')
        
        if not url or not key:
            return None
        
        sb = create_client(url, key)
        bucket_name = "ho-so-nhan-vien"  # Dùng chung bucket
        
        safe_name = sanitize_filename(filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = f"chat_files/{timestamp}_{safe_name}"
        
        sb.storage.from_(bucket_name).upload(
            path=path,
            file=file_bytes,
            file_options={"content-type": content_type or "application/octet-stream"}
        )
        return path
    except Exception as e:
        print(f"Lỗi upload file: {e}")
        return None

def get_chat_file_bytes(file_url):
    """Tải file từ Supabase Storage"""
    try:
        from supabase import create_client
        import os
        from dotenv import load_dotenv
        
        tenant = st.session_state.get('tenant')
        if tenant:
            url, key = tenant.get('supabase_url'), tenant.get('supabase_key')
        else:
            load_dotenv()
            url = os.getenv('SUPABASE_URL')
            key = os.getenv('SUPABASE_KEY')
        
        if not url or not key:
            return None
        
        sb = create_client(url, key)
        bucket_name = "ho-so-nhan-vien"
        return sb.storage.from_(bucket_name).download(file_url)
    except Exception as e:
        print(f"Lỗi tải file: {e}")
        return None