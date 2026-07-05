"""
backup_data.py
================
Module thực hiện backup dữ liệu HRM-Port:

  1. Xuất 2 bảng 'ung_vien' và 'nhan_vien' từ Supabase (Postgres) ra file Excel.
  2. Tải toàn bộ file hồ sơ (ảnh, PDF, docx...) đang lưu trên Supabase Storage
     (bucket SUPABASE_BUCKET) về thư mục local.

Kết quả được lưu vào:  D:\\hrm-port9\\backup\\<YYYY-MM-DD_HHMMSS>\\
    ├── DB\\ung_vien.xlsx
    ├── DB\\nhan_vien.xlsx
    ├── HoSo\\<id_nhan_vien>\\<Họ tên nhân viên>\\<Loại hồ sơ>_<ngày upload>_<tên file>
    └── backup_log.txt

Cách dùng:
  - Từ trong app Streamlit (nút "BACKUP DỮ LIỆU NGAY"):
        from backup_data import backup_all
        result = backup_all()

  - Chạy độc lập (vd. qua Windows Task Scheduler để backup tự động hàng tuần):
        python backup_data.py

Yêu cầu cấu hình (đặt trong file .env cạnh app.py, hoặc trong secrets.toml
của Streamlit nếu chạy trong app):
    DB_HOST=...
    DB_PORT=...
    DB_USER=...
    DB_PASSWORD=...
    DB_NAME=...
    SUPABASE_URL=...
    SUPABASE_KEY=...
"""
import os
from datetime import datetime

import psycopg2
import pandas as pd

# Phải trùng với SUPABASE_BUCKET khai báo trong app.py
SUPABASE_BUCKET = "ho-so-nhan-vien"

# Thư mục gốc lưu backup trên máy local (theo yêu cầu)
BACKUP_ROOT = r"D:\hrm-port9\backup"

# Danh sách các bảng cần backup ra Excel
TABLES_TO_BACKUP = ["ung_vien", "nhan_vien"]


def _load_config():
    """Đọc cấu hình DB & Supabase Storage.
    Ưu tiên st.secrets nếu đang chạy trong Streamlit, fallback sang .env
    để script này vẫn chạy độc lập được (vd. khi Windows Task Scheduler gọi)."""
    cfg = {}
    try:
        import streamlit as st
        if 'connections' in st.secrets and 'supabase' in st.secrets.connections:
            cfg['DB_HOST'] = st.secrets.connections.supabase.host
            cfg['DB_PORT'] = st.secrets.connections.supabase.port
            cfg['DB_USER'] = st.secrets.connections.supabase.user
            cfg['DB_PASSWORD'] = st.secrets.connections.supabase.password
            cfg['DB_NAME'] = st.secrets.connections.supabase.database
        if 'supabase' in st.secrets:
            cfg['SUPABASE_URL'] = st.secrets.supabase.get('url')
            cfg['SUPABASE_KEY'] = st.secrets.supabase.get('key')
    except Exception:
        pass

    from dotenv import load_dotenv
    load_dotenv()
    cfg.setdefault('DB_HOST', os.getenv('DB_HOST'))
    cfg.setdefault('DB_PORT', os.getenv('DB_PORT'))
    cfg.setdefault('DB_USER', os.getenv('DB_USER'))
    cfg.setdefault('DB_PASSWORD', os.getenv('DB_PASSWORD'))
    cfg.setdefault('DB_NAME', os.getenv('DB_NAME'))
    cfg.setdefault('SUPABASE_URL', os.getenv('SUPABASE_URL'))
    cfg.setdefault('SUPABASE_KEY', os.getenv('SUPABASE_KEY'))
    return cfg


def _get_db_connection(cfg):
    return psycopg2.connect(
        host=cfg.get('DB_HOST'),
        port=cfg.get('DB_PORT'),
        user=cfg.get('DB_USER'),
        password=cfg.get('DB_PASSWORD'),
        database=cfg.get('DB_NAME'),
    )


def backup_database_tables(cfg, dest_folder):
    """Xuất các bảng trong TABLES_TO_BACKUP ra file Excel trong dest_folder.
    Trả về dict: {ten_bang: (thanh_cong: bool, so_dong_hoac_loi, duong_dan_file)}"""
    os.makedirs(dest_folder, exist_ok=True)
    results = {}
    try:
        conn = _get_db_connection(cfg)
    except Exception as e:
        for table in TABLES_TO_BACKUP:
            results[table] = (False, f"Không kết nối được database: {e}", None)
        return results

    for table in TABLES_TO_BACKUP:
        try:
            df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
            out_path = os.path.join(dest_folder, f"{table}.xlsx")
            df.to_excel(out_path, index=False)
            results[table] = (True, len(df), out_path)
        except Exception as e:
            results[table] = (False, str(e), None)
    conn.close()
    return results


def backup_storage_files(cfg, dest_folder):
    """Tải toàn bộ file trong bucket Supabase Storage về dest_folder,
    giữ nguyên cấu trúc thư mục trên Storage:
        {id_nhan_vien}/{Họ tên nhân viên}/{Loại hồ sơ}_{ngày upload}_{tên file}
    Trả về dict: {ok, error, count}"""
    os.makedirs(dest_folder, exist_ok=True)

    if not cfg.get('SUPABASE_URL') or not cfg.get('SUPABASE_KEY'):
        return {"ok": False, "error": "Thiếu SUPABASE_URL / SUPABASE_KEY trong cấu hình", "count": 0}

    try:
        from supabase import create_client
    except ImportError:
        return {"ok": False, "error": "Chưa cài thư viện supabase (pip install supabase)", "count": 0}

    try:
        sb = create_client(cfg['SUPABASE_URL'], cfg['SUPABASE_KEY'])
    except Exception as e:
        return {"ok": False, "error": f"Không kết nối được Supabase Storage: {e}", "count": 0}

    count = 0
    errors = []

    def _walk(prefix=""):
        nonlocal count
        try:
            items = sb.storage.from_(SUPABASE_BUCKET).list(prefix)
        except Exception as e:
            errors.append(f"{prefix or '/'}: {e}")
            return
        for item in items or []:
            name = item.get("name")
            if not name:
                continue
            full_path = f"{prefix}/{name}" if prefix else name
            # "Thư mục ảo" trên Supabase Storage không có metadata/id -> đệ quy vào trong
            is_folder = item.get("id") is None and item.get("metadata") is None
            if is_folder:
                _walk(full_path)
            else:
                try:
                    data = sb.storage.from_(SUPABASE_BUCKET).download(full_path)
                    local_path = os.path.join(dest_folder, full_path.replace("/", os.sep))
                    os.makedirs(os.path.dirname(local_path), exist_ok=True)
                    with open(local_path, "wb") as f:
                        f.write(data)
                    count += 1
                except Exception as e:
                    errors.append(f"{full_path}: {e}")

    _walk("")
    return {"ok": len(errors) == 0, "error": "; ".join(errors) if errors else None, "count": count}


def backup_all(backup_root=BACKUP_ROOT):
    """Chạy backup toàn bộ: DB (Excel) + Storage (file hồ sơ).

    Tự nhận diện môi trường đang chạy:
      - Windows (local, kể cả khi mở Streamlit trên chính máy đó): ghi thẳng
        kết quả vào ổ đĩa `backup_root` (mặc định D:\\hrm-port9\\backup).
        Trả về {"mode": "local", "dest_folder": ..., "db": ..., "storage": ...}
      - Không phải Windows (vd. server Streamlit Cloud - Linux, không có ổ D:
        và không phải máy của người dùng): build backup vào thư mục tạm rồi
        nén thành 1 file .zip duy nhất để người dùng bấm tải xuống ngay trên
        trình duyệt. Trả về {"mode": "cloud", "zip_path": ..., "zip_bytes": ...,
        "db": ..., "storage": ...}
    """
    cfg = _load_config()
    timestamp = datetime.now().strftime('%Y-%m-%d_%H%M%S')
    is_windows = (os.name == 'nt')

    if is_windows:
        dest_folder = os.path.join(backup_root, timestamp)
        os.makedirs(dest_folder, exist_ok=True)
        work_folder = dest_folder
    else:
        import tempfile
        work_folder = os.path.join(tempfile.mkdtemp(prefix="hrm_backup_"), timestamp)
        os.makedirs(work_folder, exist_ok=True)

    db_result = backup_database_tables(cfg, os.path.join(work_folder, "DB"))
    storage_result = backup_storage_files(cfg, os.path.join(work_folder, "HoSo"))

    log_path = os.path.join(work_folder, "backup_log.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"BACKUP HRM-PORT - {timestamp}\n")
        f.write("=" * 60 + "\n")
        for table, res in db_result.items():
            if res[0]:
                f.write(f"[OK]  Bảng {table}: {res[1]} dòng -> {res[2]}\n")
            else:
                f.write(f"[LỖI] Bảng {table}: {res[1]}\n")
        if storage_result["ok"]:
            f.write(f"[OK]  Storage: đã tải {storage_result['count']} file\n")
        else:
            f.write(f"[LỖI] Storage: {storage_result['error']} (đã tải được {storage_result['count']} file trước khi lỗi)\n")

    if is_windows:
        return {
            "mode": "local",
            "dest_folder": work_folder,
            "db": db_result,
            "storage": storage_result,
        }

    # Môi trường Cloud/Linux: nén toàn bộ work_folder thành 1 file .zip
    import shutil
    zip_base = work_folder  # shutil sẽ tự thêm đuôi .zip
    zip_path = shutil.make_archive(zip_base, 'zip', work_folder)
    with open(zip_path, "rb") as f:
        zip_bytes = f.read()

    return {
        "mode": "cloud",
        "zip_path": zip_path,
        "zip_bytes": zip_bytes,
        "zip_filename": f"HRM_Port_Backup_{timestamp}.zip",
        "db": db_result,
        "storage": storage_result,
    }


if __name__ == "__main__":
    # Cho phép chạy độc lập, ví dụ qua Windows Task Scheduler:
    #   python backup_data.py
    # LƯU Ý: khi chạy qua Task Scheduler trên máy Windows của bạn, script sẽ
    # luôn ở "mode": "local" (ghi thẳng vào D:\hrm-port9\backup) vì os.name=='nt'
    # trên chính máy đó — không liên quan đến việc app Streamlit đang host ở đâu.
    result = backup_all()
    if result["mode"] == "local":
        print(f"Đã backup xong tại: {result['dest_folder']}")
    else:
        print(f"Đã backup xong (môi trường Cloud), file zip tại: {result['zip_path']}")
    for table, res in result['db'].items():
        status = "OK" if res[0] else "LỖI"
        print(f"  [{status}] {table}: {res[1]}")
    if result['storage']['ok']:
        print(f"  [OK] Storage: {result['storage']['count']} file")
    else:
        print(f"  [LỖI] Storage: {result['storage']['error']}")
