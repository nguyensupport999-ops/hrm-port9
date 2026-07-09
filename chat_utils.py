# chat_utils.py (mở rộng)
"""
Module xử lý logic cho Chat nội bộ - Phiên bản nâng cấp
"""
import psycopg2
import psycopg2.extras
import streamlit as st
from datetime import datetime
import os
import re
import unicodedata

# ========== HÀM HIỆN CÓ (GIỮ NGUYÊN) ==========
# ... (giữ nguyên tất cả hàm cũ: get_user_chat_rooms, get_room_messages, 
#     send_message, create_private_room, create_group_room, 
#     get_room_participants, mark_messages_as_read, get_all_employees)

# ========== HÀM MỚI ==========

def init_chat_tables():
    """Khởi tạo các bảng chat nếu chưa có (nâng cấp thêm cột mới)"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        
        # Bảng chat_rooms - thêm cột updated_at và deleted_at
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
        
        # Bảng chat_participants - thêm last_read_at
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
        
        # Bảng chat_messages - thêm message_type, file_name, file_size
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
                SELECT 1 FROM chat_rooms WHERE room_type = 'broadcast'
            )
        """)
        
        # Thêm tất cả nhân viên vào phòng broadcast
        c.execute("""
            INSERT INTO chat_participants (room_id, user_id, joined_at, last_read_at)
            SELECT 
                (SELECT id FROM chat_rooms WHERE room_type = 'broadcast'),
                nv.id,
                NOW(),
                NOW()
            FROM nhan_vien nv
            WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND NOT EXISTS (
                SELECT 1 FROM chat_participants cp
                WHERE cp.room_id = (SELECT id FROM chat_rooms WHERE room_type = 'broadcast')
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
        c.execute("SELECT id FROM chat_rooms WHERE room_type = 'broadcast'")
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
    if room['room_type'] == 'broadcast':
        return '📢 Thông báo chung'
    elif room['room_type'] == 'group':
        return room['room_name'] or 'Nhóm chat'
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

def get_participants_for_room(room_id):
    """Lấy danh sách ID thành viên trong phòng"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("SELECT user_id FROM chat_participants WHERE room_id = %s", (room_id,))
        results = [row[0] for row in c.fetchall()]
        db.close()
        return results
    except Exception as e:
        print(f"Lỗi lấy thành viên: {e}")
        return []

def add_participants_to_room(room_id, user_ids):
    """Thêm thành viên vào phòng nhóm"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        for user_id in user_ids:
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id)
                VALUES (%s, %s)
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
        from app import get_supabase_storage, SUPABASE_BUCKET
        
        sb = get_supabase_storage()
        if not sb:
            return None
        
        safe_name = sanitize_filename(filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = f"chat_files/{timestamp}_{safe_name}"
        
        sb.storage.from_(SUPABASE_BUCKET).upload(
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
        from app import get_supabase_storage, SUPABASE_BUCKET
        sb = get_supabase_storage()
        if not sb:
            return None
        return sb.storage.from_(SUPABASE_BUCKET).download(file_url)
    except Exception as e:
        print(f"Lỗi tải file: {e}")
        return None