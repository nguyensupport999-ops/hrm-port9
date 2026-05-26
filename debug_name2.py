import pandas as pd
import mysql.connector
from config import DB_CONFIG

file_path = r"D:\HRM_Port\Copy of STK NLD (2).xlsx"
df = pd.read_excel(file_path, sheet_name='STK', header=1, dtype=str)

db = mysql.connector.connect(**DB_CONFIG)
cursor = db.cursor(dictionary=True)
cursor.execute("SELECT Ho_ten FROM Nhan_vien WHERE Trang_thai = 'DANG_LAM'")
db_names = [row['Ho_ten'].strip() for row in cursor.fetchall()]
db.close()

print("=== 3 TÊN ĐẦU TIÊN TRONG EXCEL ===")
for i in range(min(3, len(df))):
    ho_ten = str(df.iloc[i, 1]).strip()
    print(f"  Excel '{i}': |{ho_ten}| (len={len(ho_ten)})")

print("\n=== 3 TÊN ĐẦU TIÊN TRONG DB ===")
for i, name in enumerate(db_names[:3]):
    print(f"  DB '{i}': |{name}| (len={len(name)})")

print("\n=== SO SÁNH TÊN ĐẦU TIÊN ===")
excel_name = str(df.iloc[1, 1]).strip()
print(f"  Excel: |{excel_name}|")
print(f"  DB   : |{db_names[0]}|")
print(f"  Giống? {excel_name.lower() == db_names[0].lower()}")