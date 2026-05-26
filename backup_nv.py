import pandas as pd
import mysql.connector
from config import DB_CONFIG
from datetime import datetime, date
import os
import glob

def backup_all():
    """Backup toàn bộ bảng Nhan_vien và Ung_vien ra Excel (2 sheets)"""
    try:
        db = mysql.connector.connect(**DB_CONFIG)
        cursor = db.cursor(dictionary=True)
        
        # Lấy dữ liệu Nhan_vien
        cursor.execute("SELECT * FROM Nhan_vien ORDER BY STT ASC")
        data_nv = cursor.fetchall()
        
        # Lấy dữ liệu Ung_vien
        cursor.execute("SELECT * FROM Ung_vien ORDER BY Id ASC")
        data_uv = cursor.fetchall()
        
        db.close()
        
        # Tạo DataFrame
        df_nv = pd.DataFrame(data_nv) if data_nv else pd.DataFrame()
        df_uv = pd.DataFrame(data_uv) if data_uv else pd.DataFrame()
        
        # Định dạng ngày tháng cho Nhan_vien
        for col in df_nv.columns:
            if 'Ngay' in col or 'Thang' in col:
                try:
                    df_nv[col] = pd.to_datetime(df_nv[col]).dt.strftime('%d/%m/%Y')
                except:
                    pass
        
        # Định dạng ngày tháng cho Ung_vien
        for col in df_uv.columns:
            if 'Ngay' in col or 'Thang' in col:
                try:
                    df_uv[col] = pd.to_datetime(df_uv[col]).dt.strftime('%d/%m/%Y')
                except:
                    pass
        
        # Tạo thư mục backup
        backup_path = r"D:\HRM_Port\backup"
        os.makedirs(backup_path, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Backup_All_{timestamp}.xlsx"
        file_path = os.path.join(backup_path, file_name)
        
        # File cố định (ghi đè)
        file_fixed = os.path.join(backup_path, "Backup_All_Latest.xlsx")
        
        # Xuất ra Excel với 2 sheets
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df_nv.to_excel(writer, sheet_name='NhanVien', index=False)
            df_uv.to_excel(writer, sheet_name='UngVien', index=False)
        
        with pd.ExcelWriter(file_fixed, engine='openpyxl') as writer:
            df_nv.to_excel(writer, sheet_name='NhanVien', index=False)
            df_uv.to_excel(writer, sheet_name='UngVien', index=False)
        
        # Xóa file timestamp cũ, chỉ giữ 3 file gần nhất
        timestamp_files = sorted(glob.glob(os.path.join(backup_path, "Backup_All_*.xlsx")))
        timestamp_files = [f for f in timestamp_files if "Latest" not in f]
        while len(timestamp_files) > 3:
            os.remove(timestamp_files[0])
            timestamp_files.pop(0)
        
        print(f"✅ Backup thành công!")
        print(f"   Nhân viên: {len(data_nv)} records")
        print(f"   Ứng viên: {len(data_uv)} records")
        print(f"   File timestamp: {file_name}")
        print(f"   File latest: Backup_All_Latest.xlsx")
        
    except Exception as e:
        print(f"❌ Lỗi backup: {e}")

def backup_nhan_vien():
    """Backup chỉ bảng Nhan_vien (giữ nguyên để tương thích ngược)"""
    try:
        db = mysql.connector.connect(**DB_CONFIG)
        cursor = db.cursor(dictionary=True)
        
        cursor.execute("SELECT * FROM Nhan_vien ORDER BY STT ASC")
        data = cursor.fetchall()
        db.close()
        
        if not data:
            print("Không có dữ liệu để backup!")
            return
        
        df = pd.DataFrame(data)
        
        # Định dạng ngày tháng
        for col in df.columns:
            if 'Ngay' in col or 'Thang' in col:
                try:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%d/%m/%Y')
                except:
                    pass
        
        backup_path = r"D:\HRM_Port\backup"
        os.makedirs(backup_path, exist_ok=True)
        
        file_name = f"Backup_NhanVien_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = os.path.join(backup_path, file_name)
        file_fixed = os.path.join(backup_path, "Backup_NhanVien_Latest.xlsx")
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='NhanVien', index=False)
        
        with pd.ExcelWriter(file_fixed, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='NhanVien', index=False)
        
        # Xóa file cũ, giữ 3 file gần nhất
        import glob
        timestamp_files = sorted(glob.glob(os.path.join(backup_path, "Backup_NhanVien_*.xlsx")))
        timestamp_files = [f for f in timestamp_files if "Latest" not in f]
        while len(timestamp_files) > 3:
            os.remove(timestamp_files[0])
            timestamp_files.pop(0)
        
        print(f"✅ Backup thành công! {len(data)} nhân viên")
        print(f"   File timestamp: {file_name}")
        print(f"   File latest: Backup_NhanVien_Latest.xlsx")
        
    except Exception as e:
        print(f"❌ Lỗi backup: {e}")

if __name__ == "__main__":
    backup_all()  # Backup cả 2 bảng