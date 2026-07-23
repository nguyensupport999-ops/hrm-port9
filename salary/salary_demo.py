# -*- coding: utf-8 -*-
"""
salary/salary_demo.py
======================
Công thức tính lương MẶC ĐỊNH (dùng cho MỌI tenant chưa có công thức lương riêng).

CƠ CHẾ ĐA-TENANT:
- Màn "💰 Tính thu nhập" (tinh_thu_nhap.py) tự động nạp module này qua
  _load_salary_module() (bản sao độc lập của _load_tenant_module_or_demo() trong
  app.py, đặt ngay trong tinh_thu_nhap.py để tránh import vòng) mỗi khi cần tính
  lương cho 1 nhân viên.
- Nếu tenant đang đăng nhập có file riêng 'salary/salary_{ma_so_thue}.py' (đặt cùng
  cấp với app.py, dùng đúng Mã số thuế của khách hàng đó) và file đó có hàm
  tinh_luong() hợp lệ -> dùng công thức RIÊNG của khách hàng.
- Nếu KHÔNG có (hoặc chưa tạo) -> tự động rơi về file NÀY (salary_demo.py) làm mặc
  định.

HỢP ĐỒNG CHỮ KÝ (BẮT BUỘC mọi file salary_{ma_so_thue}.py phải tuân theo):

    def tinh_luong(nv: dict, thang: int, nam: int, cfg: dict = None, **kwargs) -> dict:

    Tham số:
    - nv    : dict thông tin 1 nhân viên (1 dòng bảng `nhan_vien`, các cột như
              'id', 'ma_nv', 'ho_ten', 'phong_ban_lam_viec', 'chuc_vu',
              'luong_bao_hiem', 'phu_cap_chuc_vu', 'ngay_vao_lam'...).
    - thang, nam : kỳ lương cần tính (tháng 1-12, năm).
    - cfg   : dict cấu hình lương/thuế hiện tại (lấy từ lay_cau_hinh_luong() trong
              tinh_thu_nhap.py) — chứa các hệ số như luong_co_so, bhxh_nld,
              giam_tru_ban_than... File riêng có thể bỏ qua cfg và tự định nghĩa
              công thức hoàn toàn khác nếu muốn.
    - **kwargs : các tham số phụ tuỳ chọn (VD phu_cap_trach_nhiem, phu_cap_khac,
              ghi_chu) — file riêng có thể nhận hoặc bỏ qua tuỳ nhu cầu.

    Trả về: dict chi tiết bảng lương của 1 nhân viên, TỐI THIỂU phải có các khoá
    sau (những màn hình xuất Excel/PDF/gửi chat trong tinh_thu_nhap.py đọc đúng các
    khoá này — thiếu khoá nào sẽ lỗi/hiện trống ở màn hình tương ứng):

        nhan_vien_id, ma_nv, ho_ten, phong_ban, chuc_vu, thang, nam,
        p1_luong_vi_tri, p2_luong_nang_luc, p3_luong_hieu_qua,
        phu_cap_chuc_vu, phu_cap_tham_nien, phu_cap_trach_nhiem, phu_cap_khac,
        tong_thu_nhap, luong_dong_bh,
        bhxh_nld, bhyt_nld, bhtn_nld, doan_phi, tong_khau_tru_bh,
        so_nguoi_phu_thuoc, giam_tru_gia_canh, thu_nhap_tinh_thue, thue_tncn,
        chi_tiet_thue, tong_khau_tru, thuc_nhan, ghi_chu

QUYẾT ĐỊNH: file mặc định NÀY không viết lại công thức 3P (P1/P2/P3 + BHXH + thuế
TNCN luỹ tiến) — công thức đó đã có sẵn, đầy đủ và đang chạy thật trong
tinh_thu_nhap.py (hàm tinh_luong_nhan_vien()). Vì vậy tinh_luong() ở đây CHỈ NỐI
(delegate) thẳng sang engine thật đó, để:
    1) Mọi tenant hiện tại (chưa có file riêng) vẫn tính lương ĐÚNG Y HỆT như trước
       khi có cơ chế đa-tenant này — không phá vỡ dữ liệu/luồng đang chạy.
    2) Chỉ những tenant THỰC SỰ cần công thức khác biệt mới phải viết
       salary_{ma_so_thue}.py riêng (xem khung sườn ví dụ ở cuối file này).

Import tinh_thu_nhap ở TRONG hàm (không phải ở đầu file) để tránh lỗi import vòng:
tinh_thu_nhap.py là nơi GỌI module này (qua _load_salary_module()), nên tới lúc
tinh_luong() thực sự chạy thì tinh_thu_nhap đã nạp xong hoàn toàn — import lúc đó
là an toàn.
"""


def tinh_luong(nv: dict, thang: int, nam: int, cfg: dict = None, **kwargs) -> dict:
    """Công thức lương mặc định = dùng thẳng engine 3P thật đang chạy trong
    tinh_thu_nhap.py (tinh_luong_nhan_vien). Xem docstring đầu file để biết hợp đồng
    chữ ký đầy đủ mà 1 file salary_{ma_so_thue}.py riêng cần tuân theo nếu muốn thay
    thế công thức này."""
    from tinh_thu_nhap import tinh_luong_nhan_vien
    return tinh_luong_nhan_vien(nv, thang, nam, cfg, **kwargs)


# ============================================================================
# KHUNG SƯỜN VÍ DỤ — cách viết 1 công thức lương HOÀN TOÀN RIÊNG cho 1 khách hàng
# ============================================================================
# Muốn có công thức lương RIÊNG (khác hẳn 3P mặc định) cho 1 khách hàng: tạo file
# 'salary/salary_{ma_so_thue}.py' (đặt cùng cấp với app.py, đúng Mã số thuế của
# khách hàng đó) rồi viết hàm tinh_luong() theo đúng hợp đồng chữ ký ở trên. Ví dụ
# công thức "lương khoán đơn giản" (không dùng 3P, không dùng cfg):
#
#   def tinh_luong(nv, thang, nam, cfg=None, **kwargs):
#       luong_co_ban = float(nv.get("luong_bao_hiem") or 0)
#       phu_cap = float(nv.get("phu_cap_chuc_vu") or 0)
#       tong_thu_nhap = luong_co_ban + phu_cap
#       bhxh_nld = round(luong_co_ban * 0.08)
#       thuc_nhan = tong_thu_nhap - bhxh_nld
#       return {
#           "nhan_vien_id": nv["id"], "ma_nv": nv.get("ma_nv"), "ho_ten": nv.get("ho_ten"),
#           "phong_ban": nv.get("phong_ban_lam_viec"), "chuc_vu": nv.get("chuc_vu"),
#           "thang": thang, "nam": nam,
#           "p1_luong_vi_tri": round(luong_co_ban), "p2_luong_nang_luc": 0, "p3_luong_hieu_qua": 0,
#           "phu_cap_chuc_vu": round(phu_cap), "phu_cap_tham_nien": 0,
#           "phu_cap_trach_nhiem": 0, "phu_cap_khac": 0,
#           "tong_thu_nhap": round(tong_thu_nhap), "luong_dong_bh": round(luong_co_ban),
#           "bhxh_nld": bhxh_nld, "bhyt_nld": 0, "bhtn_nld": 0, "doan_phi": 0,
#           "tong_khau_tru_bh": bhxh_nld,
#           "so_nguoi_phu_thuoc": 0, "giam_tru_gia_canh": 0,
#           "thu_nhap_tinh_thue": 0, "thue_tncn": 0, "chi_tiet_thue": [],
#           "tong_khau_tru": bhxh_nld, "thuc_nhan": round(thuc_nhan), "ghi_chu": "",
#       }
#
# Sau khi tạo file, KHÔNG cần đăng ký/khai báo gì thêm — màn "💰 Cấu hình công ty >
# Phần mềm tính lương" sẽ tự nhận diện và báo cho admin biết đang dùng đúng file
# riêng này (thay vì salary_demo.py).
