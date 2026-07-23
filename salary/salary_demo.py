# -*- coding: utf-8 -*-
"""
salary/salary_demo.py
======================
Công thức tính lương MẶC ĐỊNH (dùng chung cho mọi tenant CHƯA có công thức riêng).

CƠ CHẾ "MỖI TENANT 1 CÔNG THỨC LƯƠNG RIÊNG":
- app.py (hàm _load_tenant_module_or_demo) sẽ tìm file `salary/salary_{ma_so_thue}.py`
  của tenant đang đăng nhập (ma_so_thue lấy từ Control Plane — khoá quản lý tenant
  hiện tại, xem control_plane.py). Nếu tìm thấy, dùng công thức riêng của tenant đó.
- Nếu tenant CHƯA có file riêng (chưa tuỳ biến công thức lương), tự động dùng file
  này (salary_demo.py) làm mặc định.

MUỐN TẠO CÔNG THỨC LƯƠNG RIÊNG CHO 1 TENANT?
1. Copy file này thành `salary/salary_{ma_so_thue}.py` (đúng mã số thuế của tenant đó,
   VD: `salary/salary_0304577099.py`).
2. Viết lại nội dung hàm `tinh_luong()` bên dưới theo đúng công thức lương riêng của
   doanh nghiệp đó (bậc thang thuế TNCN riêng, phụ cấp riêng, cách tính OT riêng...).
   Giữ nguyên chữ ký hàm (tên hàm + tham số) để phần gọi trong app không cần sửa gì.
3. Không cần sửa gì ở app.py — lần tính lương tiếp theo của tenant đó sẽ tự dùng
   đúng công thức mới.

LƯU Ý: đây là bộ khung (interface) cho cơ chế "mỗi tenant 1 công thức lương riêng".
Module `tinh_thu_nhap.py` (màn "💰 Tính thu nhập") hiện có tự triển khai logic tính
lương riêng của mình, CHƯA gọi qua plugin này. Muốn tinh_thu_nhap.py áp dụng cơ chế
đa-tenant, hãy gọi `_load_tenant_module_or_demo('salary', 'salary', ma_so_thue)` từ
trong file đó rồi dùng module trả về (xem ví dụ cuối file này) thay vì hard-code công
thức tính lương ngay trong tinh_thu_nhap.py.
"""


def tinh_luong(nhan_vien: dict, thang: int, nam: int, du_lieu_cham_cong: dict, cau_hinh: dict = None) -> dict:
    """Tính lương 1 nhân viên trong 1 tháng. Đây là CHỮ KÝ HÀM CHUẨN mà mọi file
    salary_{ma_so_thue}.py cần tuân theo để app có thể gọi thay thế cho nhau.

    Tham số:
        nhan_vien: dict thông tin nhân viên (ma_nv, ho_ten, luong_co_ban, chuc_vu,
                   phong_ban_lam_viec, he_so_luong, ...).
        thang, nam: tháng/năm cần tính lương.
        du_lieu_cham_cong: dict dữ liệu chấm công trong tháng (số công chuẩn, số công
                   thực tế, số giờ OT, số ngày nghỉ phép/không lương...).
        cau_hinh: dict cấu hình lương của tenant (lấy từ get_cau_hinh(...) trong app.py,
                   VD mức đóng BHXH, mức giảm trừ gia cảnh...). Có thể để trống (None)
                   nếu công thức không cần thêm cấu hình ngoài các tham số trên.

    Trả về: dict breakdown lương, tối thiểu gồm các khoá:
        {
            'luong_co_ban': ...,      # Lương cơ bản theo hợp đồng
            'phu_cap': ...,           # Tổng phụ cấp (nếu có)
            'luong_ot': ...,          # Lương tăng ca (nếu có)
            'khau_tru_bhxh': ...,     # Khấu trừ BHXH/BHYT/BHTN (phần NLĐ đóng)
            'thue_tncn': ...,         # Thuế TNCN tạm khấu trừ (nếu có)
            'thuc_linh': ...,         # Số tiền thực lĩnh cuối cùng
        }

    ĐÂY LÀ CÔNG THỨC KHUNG/MẶC ĐỊNH — chỉ tính đơn giản theo tỷ lệ ngày công thực tế
    trên ngày công chuẩn, CHƯA có logic thuế TNCN/BHXH chi tiết. Áp dụng cho tenant nào
    chưa có công thức riêng; khi cần chính xác theo luật hiện hành, hãy tạo file
    salary_{ma_so_thue}.py riêng theo hướng dẫn ở đầu file này.
    """
    luong_co_ban = float(nhan_vien.get('luong_co_ban') or 0)
    cong_chuan = float((du_lieu_cham_cong or {}).get('cong_chuan') or 26)
    cong_thuc_te = float((du_lieu_cham_cong or {}).get('cong_thuc_te') or cong_chuan)

    ty_le_cong = (cong_thuc_te / cong_chuan) if cong_chuan else 0
    luong_theo_cong = round(luong_co_ban * ty_le_cong)

    phu_cap = float(nhan_vien.get('phu_cap') or 0)
    luong_ot = 0  # Chưa có logic OT trong bản khung mặc định

    khau_tru_bhxh = 0
    thue_tncn = 0
    thuc_linh = luong_theo_cong + phu_cap + luong_ot - khau_tru_bhxh - thue_tncn

    return {
        'luong_co_ban': luong_theo_cong,
        'phu_cap': phu_cap,
        'luong_ot': luong_ot,
        'khau_tru_bhxh': khau_tru_bhxh,
        'thue_tncn': thue_tncn,
        'thuc_linh': thuc_linh,
    }


# ----------------------------------------------------------------------------------
# Ví dụ cách gọi plugin này (hoặc plugin riêng của tenant) từ trong tinh_thu_nhap.py:
#
#   from app import _load_tenant_module_or_demo
#   ma_so_thue = (st.session_state.get('tenant') or {}).get('ma_so_thue', '')
#   module_luong = _load_tenant_module_or_demo('salary', 'salary', ma_so_thue)
#   ket_qua = module_luong.tinh_luong(nhan_vien, thang, nam, du_lieu_cham_cong)
# ----------------------------------------------------------------------------------
