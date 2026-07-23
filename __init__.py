# -*- coding: utf-8 -*-
"""Package chứa công thức tính lương riêng theo từng tenant.

- salary_demo.py       : Công thức lương mặc định (nối vào engine 3P thật trong
                          tinh_thu_nhap.py) — dùng cho tenant chưa tuỳ biến.
- salary_{ma_so_thue}.py : Công thức lương riêng của 1 khách hàng cụ thể (tuỳ chọn).

Được tinh_thu_nhap.py nạp động qua _load_salary_module() mỗi khi tính lương.
"""
