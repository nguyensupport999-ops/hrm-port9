"""
salary_1.py — Công thức tính lương mặc định (Plugin #1)

QUY ƯỚC PLUGIN LƯƠNG:
- Mỗi công thức tính lương là 1 file riêng: salary/salary_{key}.py (key = số nguyên hoặc mã ngắn, VD: salary_1.py, salary_2.py...)
- Mỗi file PHẢI có hàm `tinh_luong(nhan_vien: dict, ky_luong: dict) -> dict` với input/output thống nhất bên dưới,
  để app.py có thể gọi bất kỳ plugin nào theo cùng 1 cách (chọn qua cấu hình 'luong_plugin_key' trong
  ⚙️ Danh mục → ⚙️ Cấu hình Doanh nghiệp).

INPUT:
    nhan_vien: dict thông tin nhân viên (toàn bộ row bảng nhan_vien: ho_ten, ma_nv, chuc_vu,
               loai_hop_dong, he_so_luong, phu_cap..., ngay_vao_lam, ...)
    ky_luong: dict thông tin kỳ lương, ví dụ:
        {
            'thang': 7, 'nam': 2026,
            'ngay_cong_chuan': 26,
            'ngay_cong_thuc_te': 24,
            'so_gio_tang_ca': 0,
            ...
        }

OUTPUT (dict) - TỐI THIỂU cần các khoá sau để phần còn lại của app hiển thị/xuất báo cáo được:
    {
        'luong_co_ban': 0,          # lương cơ bản theo hợp đồng/hệ số
        'phu_cap': 0,               # tổng phụ cấp
        'luong_tang_ca': 0,
        'khau_tru_bhxh_nld': 0,     # phần NLĐ đóng BHXH/BHYT/BHTN
        'thue_tncn': 0,
        'thuc_linh': 0,             # lương thực nhận cuối cùng
        'chi_tiet': {...}           # (tuỳ chọn) breakdown chi tiết để hiển thị/audit
    }

TRẠNG THÁI HIỆN TẠI: đây là bản khung (scaffold) — sẽ hoàn thiện công thức thực tế dần theo nhu cầu phát sinh.
"""


def tinh_luong(nhan_vien: dict, ky_luong: dict) -> dict:
    """Công thức tính lương mặc định #1 — CHƯA CÓ LOGIC THỰC TẾ, chỉ trả khung dữ liệu rỗng.
    TODO: bổ sung công thức thực tế khi có yêu cầu cụ thể từ doanh nghiệp.
    """
    return {
        'luong_co_ban': 0,
        'phu_cap': 0,
        'luong_tang_ca': 0,
        'khau_tru_bhxh_nld': 0,
        'thue_tncn': 0,
        'thuc_linh': 0,
        'chi_tiet': {
            'ghi_chu': 'salary_1.py chưa được cấu hình công thức thực tế — vui lòng bổ sung.'
        }
    }
