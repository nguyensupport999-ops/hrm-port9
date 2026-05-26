import pandas as pd

# Đọc file Excel mới (V7)
file_moi = r"D:\HRM_Port\26.05.09 -Danh sách phổng vấn -V7.xlsx"
df_moi = pd.read_excel(file_moi, sheet_name='Tong the', header=2)

# Copy file backup vào thư mục HRM_Port trước khi chạy
file_backup = r"D:\HRM_Port\backup\Backup_All_20260512_100848.xlsx"
df_backup = pd.read_excel(file_backup, sheet_name='UngVien')

# Đặt tên cột cho file mới
df_moi.columns = ['TT', 'Tham_chieu', 'Ho_ten', 'Vi_tri', 'Chap_thuan', 'Xem_xet',
                   'Thoi_gian_di_lam', 'Luong', 'Trao_doi_sep', 'Nam_sinh',
                   'Tinh_trang_hon_nhan', 'Gioi_tinh', 'Extra1', 'Dien_thoai',
                   'Vi_tri_2', 'Can_tuyen', 'Da_tuyen', 'Ghi_chu', 'Extra2',
                   'Vi_tri_3', 'Extra3', 'Da_tuyen_Huy', 'Extra4']

# Lấy danh sách tên + SĐT từ backup (giữ nguyên)
backup_names = set()
backup_phones = set()
for index, row in df_backup.iterrows():
    ho_ten = str(row['Ho_ten']).strip() if not pd.isna(row['Ho_ten']) else ''
    dien_thoai = str(row['Dien_thoai']).strip() if not pd.isna(row['Dien_thoai']) else ''
    if ho_ten:
        backup_names.add(ho_ten)
    if dien_thoai and dien_thoai not in ['nan', 'None', '']:
        backup_phones.add(dien_thoai)

# Lọc ứng viên mới nhưng chưa có trong backup
thieu = []
for index, row in df_moi.iterrows():
    ho_ten = str(row['Ho_ten']).strip() if not pd.isna(row['Ho_ten']) else ''
    dien_thoai = str(row['Dien_thoai']).strip() if not pd.isna(row['Dien_thoai']) else ''
    
    if ho_ten == '' or ho_ten == 'nan':
        continue
    
    # Chuẩn hóa SĐT
    if dien_thoai in ['nan', 'None', '', 'xz']:
        dien_thoai = ''
    else:
        dien_thoai = dien_thoai.replace(';', ',').split(',')[0].strip()
    
    # Kiểm tra có trong backup không (theo tên HOẶC SĐT)
    found = False
    if ho_ten in backup_names:
        found = True
    if dien_thoai and dien_thoai in backup_phones:
        found = True
    
    if not found:
        thieu.append({
            'Ho_ten': ho_ten,
            'Dien_thoai': dien_thoai,
            'Vi_tri': str(row['Vi_tri']).strip() if not pd.isna(row['Vi_tri']) else ''
        })

# Xuất kết quả
if thieu:
    df_thieu = pd.DataFrame(thieu)
    output = r"D:\HRM_Port\Danh_sach_ung_vien_thieu.xlsx"
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_thieu.to_excel(writer, sheet_name='Thieu', index=False)
    print(f"✅ Đã lưu {len(thieu)} ứng viên thiếu vào: {output}")
else:
    print("✅ Không có ứng viên nào thiếu!")

print(f"Tổng file mới: {len(df_moi)}")
print(f"Tổng backup: {len(backup_names)}")
print(f"Thiếu: {len(thieu)}")

if thieu:
    print("\nDanh sách thiếu:")
    for nv in thieu[:10]:
        print(f"  - {nv['Ho_ten']} | {nv['Dien_thoai']}")