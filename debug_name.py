import pandas as pd
import mysql.connector
from config import DB_CONFIG

file_path = r"D:\HRM_Port\Copy of STK NLD (2).xlsx"
df = pd.read_excel(file_path, sheet_name='STK', header=1, dtype=str)

db = mysql.connector.connect(**DB_CONFIG)
cursor = db.cursor(dictionary=True)

# Lấy danh sách tên từ DB
cursor.execute("SELECT Ho_ten FROM Nhan_vien WHERE Trang_thai = 'DANG_LAM'")
db_names = [row['Ho_ten'] for row in cursor.fetchall()]
db.close()

print("=== TÊN TRONG EXCEL ===  |  === TÊN TRONG DB ===")
for index, row in df.iterrows():
    ho_ten_excel = str(row.iloc[1]).strip() if not pd.isna(row.iloc[1]) else ''
    if ho_ten_excel == '' or ho_ten_excel == 'nan':
        continue
    
    # Tìm tên gần giống trong DB
    found = False
    for db_name in db_names:
        if ho_ten_excel.lower() == db_name.lower():
            found = True
            break
    
    if not found:
        print(f"❌ Excel: '{ho_ten_excel}' → KHÔNG TÌM THẤY TRONG DB")
        # Gợi ý tên gần đúng
        for db_name in db_names:
            if ho_ten_excel.lower()[:5] in db_name.lower() or db_name.lower()[:5] in ho_ten_excel.lower():
                print(f"   Gợi ý DB: '{db_name}'")