import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

print("=== THÔNG TIN KẾT NỐI ===")
print(f"Host: {os.getenv('DB_HOST')}")
print(f"Port: {os.getenv('DB_PORT')}")
print(f"User: {os.getenv('DB_USER')}")
print(f"DB Name: {os.getenv('DB_NAME')}")
print("=========================")

try:
    conn = psycopg2.connect(
        host=os.getenv('DB_HOST'),
        port=os.getenv('DB_PORT'),
        user=os.getenv('DB_USER'),
        password=os.getenv('DB_PASSWORD'),
        database=os.getenv('DB_NAME')
    )
    print("✅ Kết nối thành công!")
    conn.close()
except Exception as e:
    print(f"❌ Lỗi: {e}")