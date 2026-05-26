import psycopg2
import os

# Đọc trực tiếp từ biến môi trường
host = "aws-1-ap-northeast-1.pooler.supabase.com"
port = "5432"
user = "postgres.ioesyihbsdxmxrotdetx"
password = "Xbr2w6bo1s5JY4Vq"
database = "postgres"

print(f"Connecting to: {host}:{port} as {user}")

try:
    conn = psycopg2.connect(
        host=host,
        port=port,
        user=user,
        password=password,
        database=database
    )
    print("✅ Kết nối thành công!")
    conn.close()
except Exception as e:
    print(f"❌ Lỗi: {e}")