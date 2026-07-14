"""
audit_dashboard_data.py
------------------------------------------------------------------------------
SCRIPT CHẨN ĐOÁN LỆCH SỐ LIỆU GIỮA CÁC BIỂU ĐỒ TRONG DASHBOARD (HRM-Port).

MỤC ĐÍCH
  Dashboard có 6 biểu đồ đều PHẢI dùng chung 1 tiêu chuẩn lọc nhân sự
  (DK_CHUAN_NV = đang làm/thử việc + đã có số HĐLĐ). Nếu một biểu đồ nào đó
  cho ra Tổng khác với các biểu đồ còn lại (như vụ "Cơ cấu theo Chức danh"
  = 43 trong khi các biểu đồ khác = 53), gần như chắc chắn là do:
    (a) câu SQL có thêm LIMIT / điều kiện lọc phụ mà các câu khác không có, hoặc
    (b) trường dùng để GROUP BY có giá trị NULL/rỗng bị PostgreSQL âm thầm bỏ qua
        (KHÔNG đúng — NULL vẫn được PostgreSQL gộp thành 1 nhóm, nhưng vẫn liệt
        kê ra đây để chắc chắn), hoặc
    (c) điều kiện WHERE phụ chỉ có ở 1 biểu đồ (vd chỉ lấy ngay_vao_lam trong
        6 tháng gần nhất cho biểu đồ xu hướng tuyển dụng — cái này là CHỦ Ý,
        không phải lỗi).

CÁCH DÙNG
  Đặt file này CÙNG THƯ MỤC với app.py và control_plane.py trong repo, sau đó
  chạy bằng Streamlit (để tái sử dụng đúng cơ chế multi-tenant + st.secrets
  mà app chính đang dùng):

      streamlit run audit_dashboard_data.py

  Nếu chỉ muốn chạy audit cho 1 khách hàng cụ thể (multi-tenant theo domain),
  hãy mở đúng URL của khách hàng đó rồi chạy script trỏ vào cùng domain, hoặc
  sửa phần "XÁC ĐỊNH TENANT" bên dưới cho phù hợp với cách resolve_tenant()
  của bạn hoạt động (theo domain / theo query param / theo secrets...).

KẾT QUẢ IN RA
  1. Bảng so sánh Tổng của từng biểu đồ so với "Tổng chuẩn" (DK_CHUAN_NV).
  2. Với biểu đồ Chức danh: liệt kê ĐẦY ĐỦ mọi chức danh + số lượng, đánh dấu
     rõ chức danh nào sẽ bị LIMIT cắt mất nếu chẳng may có LIMIT trong SQL.
  3. Danh sách nhân viên bị loại khỏi "Tổng chuẩn" kèm lý do (nghỉ việc /
     chưa có số HĐLĐ...) — để biết chắc 53 hay 43 mới là con số đúng.
  4. Cảnh báo nếu phát hiện NULL/rỗng ở các trường dùng để group (giới tính,
     trình độ, phòng ban, ngày sinh) trong nhóm nhân sự chuẩn.
------------------------------------------------------------------------------
"""

import streamlit as st
import psycopg2.extras

st.set_page_config(page_title="Audit Dashboard Data", layout="wide")
st.title("🔍 Audit lệch số liệu Dashboard")

# ──────────────────────────────────────────────────────────────────────────
# XÁC ĐỊNH TENANT & KẾT NỐI DB — dùng lại đúng cơ chế của app.py
# ──────────────────────────────────────────────────────────────────────────
try:
    from control_plane import DatabaseEngine, resolve_tenant
except ImportError:
    st.error(
        "❌ Không import được `control_plane`. Hãy đặt file audit_dashboard_data.py "
        "CÙNG THƯ MỤC với app.py và control_plane.py trong repo, rồi chạy lại "
        "bằng `streamlit run audit_dashboard_data.py`."
    )
    st.stop()

resolve_tenant()
if 'db_engine' not in st.session_state:
    st.session_state.db_engine = DatabaseEngine(st.session_state.get('tenant'))

db = st.session_state.db_engine.get_connection()
c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

if st.session_state.get('tenant'):
    st.caption(f"🏢 Tenant: {st.session_state.tenant.get('ma_cty', '(không rõ)')}")

# ──────────────────────────────────────────────────────────────────────────
# TIÊU CHUẨN LỌC NHÂN SỰ — PHẢI GIỐNG HỆT app.py (biến DK_CHUAN_NV)
# Nếu bạn sửa điều kiện này trong app.py, nhớ sửa lại y hệt ở đây.
# ──────────────────────────────────────────────────────────────────────────
DK_CHUAN_NV = "trang_thai IN ('DANG_LAM', 'THU_VIEC') AND so_hdld IS NOT NULL AND so_hdld != ''"

# ──────────────────────────────────────────────────────────────────────────
# 0. TỔNG CHUẨN
# ──────────────────────────────────────────────────────────────────────────
c.execute(f"SELECT COUNT(*) AS t FROM nhan_vien WHERE {DK_CHUAN_NV}")
TONG_CHUAN = c.fetchone()['t']

st.header("0️⃣ Tổng chuẩn (dùng để đối chiếu mọi biểu đồ)")
st.success(f"**Tổng chuẩn (DK_CHUAN_NV) = {TONG_CHUAN} nhân viên**")
st.caption("Điều kiện: `" + DK_CHUAN_NV + "`")

st.divider()

# ──────────────────────────────────────────────────────────────────────────
# 1. TỔNG CỦA TỪNG BIỂU ĐỒ — so sánh với Tổng chuẩn
# ──────────────────────────────────────────────────────────────────────────
st.header("1️⃣ So sánh Tổng từng biểu đồ với Tổng chuẩn")

checks = []

# a. Phòng ban
c.execute(f"""SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV}""")
checks.append(("Cơ cấu theo Phòng ban", c.fetchone()['t'], "không filter thêm"))

# b. Giới tính
c.execute(f"""SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV}""")
checks.append(("Cơ cấu theo Giới tính", c.fetchone()['t'], "không filter thêm"))

# c. Trình độ học vấn
c.execute(f"""SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV}""")
checks.append(("Cơ cấu theo Trình độ học vấn", c.fetchone()['t'], "không filter thêm"))

# d. Chức danh — ĐÂY LÀ BIỂU ĐỒ TỪNG BỊ LIMIT 10 GÂY LỆCH SỐ
c.execute(f"""
    SELECT COUNT(*) t FROM nhan_vien
    WHERE {DK_CHUAN_NV} AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != ''
""")
tong_co_chuc_danh = c.fetchone()['t']
checks.append((
    "Cơ cấu theo Chức danh (SQL đầy đủ, chưa LIMIT)",
    tong_co_chuc_danh,
    "filter thêm: chuc_danh_nghe IS NOT NULL AND != ''"
))

# e. Độ tuổi — chỉ tính người có ngày sinh (CHỦ Ý, không phải lỗi)
c.execute(f"""SELECT COUNT(*) t FROM nhan_vien WHERE {DK_CHUAN_NV} AND ngay_sinh IS NOT NULL""")
checks.append(("Cơ cấu theo Độ tuổi", c.fetchone()['t'], "filter thêm: ngay_sinh IS NOT NULL (chủ ý)"))

# f. Xu hướng tuyển dụng 6 tháng — CHỦ Ý chỉ lấy 6 tháng gần nhất, không phải lỗi
c.execute(f"""
    SELECT COUNT(*) t FROM nhan_vien
    WHERE ngay_vao_lam >= (CURRENT_DATE - INTERVAL '6 months') AND {DK_CHUAN_NV}
""")
checks.append(("Xu hướng tuyển dụng 6 tháng", c.fetchone()['t'], "filter thêm: chỉ 6 tháng gần nhất (chủ ý)"))

rows_out = []
for ten_bd, tong_bd, ghi_chu in checks:
    lech = tong_bd - TONG_CHUAN
    if lech == 0:
        trang_thai = "✅ Khớp"
    elif "chủ ý" in ghi_chu:
        trang_thai = f"ℹ️ Lệch {lech:+d} — chủ ý (xem ghi chú)"
    else:
        trang_thai = f"🚨 LỆCH {lech:+d} — CẦN KIỂM TRA"
    rows_out.append({
        "Biểu đồ": ten_bd,
        "Tổng": tong_bd,
        "Tổng chuẩn": TONG_CHUAN,
        "Chênh lệch": lech,
        "Trạng thái": trang_thai,
        "Ghi chú": ghi_chu,
    })

import pandas as pd
st.dataframe(pd.DataFrame(rows_out), width='stretch', hide_index=True)

st.divider()

# ──────────────────────────────────────────────────────────────────────────
# 2. CHI TIẾT: TOÀN BỘ CHỨC DANH — để lộ những chức danh sẽ bị LIMIT cắt mất
# ──────────────────────────────────────────────────────────────────────────
st.header("2️⃣ Chi tiết toàn bộ Chức danh (không LIMIT)")

c.execute(f"""
    SELECT chuc_danh_nghe, COUNT(*) as so_luong
    FROM nhan_vien WHERE {DK_CHUAN_NV}
    AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != ''
    GROUP BY chuc_danh_nghe
    ORDER BY so_luong DESC
""")
full_role_data = c.fetchall()
df_full_role = pd.DataFrame(full_role_data)

if not df_full_role.empty:
    df_full_role.index = range(1, len(df_full_role) + 1)
    df_full_role['Sẽ bị LIMIT 10 cắt mất?'] = [
        "🚨 CÓ — mất khỏi biểu đồ/tổng nếu SQL có LIMIT 10" if i > 10 else "✅ Không"
        for i in df_full_role.index
    ]
    st.dataframe(df_full_role, width='stretch')

    so_chuc_danh_bi_cat = (df_full_role.index > 10).sum()
    nguoi_bi_cat = df_full_role[df_full_role.index > 10]['so_luong'].sum() if 'so_luong' in df_full_role else 0
    if so_chuc_danh_bi_cat > 0:
        st.warning(
            f"⚠️ Nếu SQL có `LIMIT 10`, sẽ có **{so_chuc_danh_bi_cat} chức danh** "
            f"(tổng cộng **{nguoi_bi_cat} người**) bị loại hoàn toàn khỏi biểu đồ "
            f"'Cơ cấu theo Chức danh' — đây chính xác là nguyên nhân Tổng = 43 thay vì 53."
        )
    else:
        st.success("✅ Tổng số chức danh ≤ 10, không có nguy cơ bị LIMIT cắt mất dữ liệu.")
else:
    st.info("Không có dữ liệu chức danh.")

st.divider()

# ──────────────────────────────────────────────────────────────────────────
# 3. NHÂN VIÊN BỊ LOẠI KHỎI "TỔNG CHUẨN" — kèm lý do
# ──────────────────────────────────────────────────────────────────────────
st.header("3️⃣ Nhân viên KHÔNG nằm trong Tổng chuẩn (kèm lý do)")

c.execute(f"""
    SELECT ma_nv, ho_ten, trang_thai, so_hdld,
        CASE
            WHEN trang_thai NOT IN ('DANG_LAM','THU_VIEC') THEN 'Trạng thái không phải Đang làm/Thử việc: ' || COALESCE(trang_thai,'(rỗng)')
            WHEN so_hdld IS NULL OR so_hdld = '' THEN 'Chưa có số HĐLĐ (hồ sơ chưa hoàn thiện)'
            ELSE 'Không rõ'
        END as ly_do
    FROM nhan_vien
    WHERE NOT ({DK_CHUAN_NV})
    ORDER BY trang_thai, ho_ten
""")
loai_tru = c.fetchall()
if loai_tru:
    st.dataframe(pd.DataFrame(loai_tru), width='stretch', hide_index=True)
    st.caption(f"Tổng cộng {len(loai_tru)} nhân viên bị loại khỏi Tổng chuẩn vì lý do trên.")
else:
    st.info("Không có nhân viên nào bị loại — mọi bản ghi trong nhan_vien đều đạt Tổng chuẩn.")

st.divider()

# ──────────────────────────────────────────────────────────────────────────
# 4. CẢNH BÁO NULL/RỖNG Ở CÁC TRƯỜNG DÙNG GROUP BY (trong nhóm Tổng chuẩn)
# ──────────────────────────────────────────────────────────────────────────
st.header("4️⃣ Kiểm tra NULL/rỗng ở các trường dùng để nhóm biểu đồ")
st.caption(
    "Lưu ý: NULL vẫn được PostgreSQL GROUP BY thành 1 nhóm riêng (không bị mất "
    "khỏi Tổng), nhưng sẽ hiện thành nhãn trống/'None' trên biểu đồ — nên vẫn "
    "liệt kê ra đây để tiện dọn dữ liệu cho đẹp."
)

truong_can_kiem_tra = {
    "gioi_tinh": "Giới tính",
    "trinh_do": "Trình độ",
    "phong_ban_lam_viec": "Phòng ban",
    "ngay_sinh": "Ngày sinh",
    "chuc_danh_nghe": "Chức danh",
}

rows_null = []
for cot, nhan in truong_can_kiem_tra.items():
    c.execute(f"""
        SELECT COUNT(*) t FROM nhan_vien
        WHERE {DK_CHUAN_NV} AND ({cot} IS NULL OR {cot}::text = '')
    """)
    so_luong_null = c.fetchone()['t']
    rows_null.append({
        "Trường": nhan,
        "Số nhân viên NULL/rỗng": so_luong_null,
        "Trạng thái": "✅ Không có" if so_luong_null == 0 else f"⚠️ Có {so_luong_null} người thiếu dữ liệu"
    })

st.dataframe(pd.DataFrame(rows_null), width='stretch', hide_index=True)

db.close()

st.divider()
st.info(
    "✅ Đã audit xong. Nếu mục 1️⃣ có dòng '🚨 LỆCH — CẦN KIỂM TRA', hãy đối chiếu "
    "câu SQL tương ứng trong app.py (khu vực `# 2. Dữ liệu cho các biểu đồ` trong "
    "menu 📊 Dashboard) để tìm điều kiện lọc/LIMIT khác biệt, giống hệt cách phát "
    "hiện lỗi LIMIT 10 ở biểu đồ Chức danh."
)
