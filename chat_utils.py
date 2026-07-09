# chat_utils.py
"""
Module xử lý logic cho Chat nội bộ
"""
import psycopg2
import psycopg2.extras
import streamlit as st
from datetime import datetime

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

def send_message(room_id, sender_id, content, message_type='text', file_url=None):
    """Gửi tin nhắn mới"""
    try:
        db = st.session_state.db_engine.get_connection()
        c = db.cursor()
        c.execute("""
            INSERT INTO chat_messages (room_id, sender_id, content, message_type, file_url)
            VALUES (%s, %s, %s, %s, %s)
        """, (room_id, sender_id, content, message_type, file_url))
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
        """, (user1_id, user2_id))
        existing = c.fetchone()
        if existing:
            return existing[0]
        
        # Tạo phòng mới
        c.execute("""
            INSERT INTO chat_rooms (room_name, room_type, created_by) 
            VALUES (%s, 'private', %s) RETURNING id
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
            INSERT INTO chat_rooms (room_name, room_type, created_by) 
            VALUES (%s, 'group', %s) RETURNING id
        """, (room_name, creator_id))
        room_id = c.fetchone()[0]
        
        # Thêm các thành viên (bao gồm cả người tạo)
        for user_id in [creator_id] + member_ids:
            c.execute("""
                INSERT INTO chat_participants (room_id, user_id) 
                VALUES (%s, %s)
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
            SELECT nv.id, nv.ma_nv, nv.ho_ten, nv.dien_thoai
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
        c.execute("""
            UPDATE chat_messages 
            SET is_read = TRUE 
            WHERE room_id = %s AND sender_id != %s
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