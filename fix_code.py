import re

# Đọc file app.py
with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Thay thế import
content = content.replace('import mysql.connector', 'import psycopg2\nimport psycopg2.extras')

# Thay thế hàm get_connection
old_func = r'def get_connection\(\):.*?return mysql\.connector\.connect\([^)]+\)'
new_func = '''def get_connection():
    return psycopg2.connect(
        host=os.getenv('DB_HOST'),
        port=os.getenv('DB_PORT'),
        user=os.getenv('DB_USER'),
        password=os.getenv('DB_PASSWORD'),
        database=os.getenv('DB_NAME')
    )'''
content = re.sub(old_func, new_func, content, flags=re.DOTALL)

# Thay thế cursor
content = content.replace('db.cursor(dictionary=True)', 'db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)')

# Ghi lại file
with open('app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("✅ Đã sửa code xong!")