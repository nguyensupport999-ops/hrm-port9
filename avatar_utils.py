# avatar_utils.py
import os
import base64
import streamlit as st

def get_avatar_base64(gioi_tinh, ho_ten=None):
    """
    Lấy ảnh avatar mặc định theo giới tính
    Trả về base64 của ảnh, hoặc None nếu không tìm thấy file
    """
    # Xác định file avatar theo giới tính
    avatar_file = "avatar_male.png" if gioi_tinh == "Nam" else "avatar_female.png"
    avatar_path = os.path.join(os.path.dirname(__file__), "static", avatar_file)
    
    # Thử đọc file avatar
    if os.path.exists(avatar_path):
        try:
            with open(avatar_path, "rb") as f:
                img_data = base64.b64encode(f.read()).decode()
                # Xác định loại file
                ext = avatar_file.split('.')[-1].lower()
                mime = f"image/{ext}" if ext in ['png', 'jpg', 'jpeg', 'gif'] else "image/png"
                return f"data:{mime};base64,{img_data}"
        except Exception as e:
            print(f"Lỗi đọc file avatar: {e}")
    
    # Fallback: tạo avatar từ UI Avatars
    if ho_ten:
        return f"https://ui-avatars.com/api/?name={ho_ten.replace(' ', '+')}&size=200&background=f59e0b&color=fff"
    return None

def get_avatar_html(gioi_tinh, ho_ten, size=200):
    """
    Trả về HTML cho avatar với kích thước chỉ định
    """
    avatar_url = get_avatar_base64(gioi_tinh, ho_ten)
    if avatar_url:
        return f"""
        <div class="avatar-wrapper">
            <img src="{avatar_url}" class="avatar-img" style="width:{size}px; height:{size}px;">
        </div>
        """
    return ""