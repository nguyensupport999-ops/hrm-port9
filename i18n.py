# -*- coding: utf-8 -*-
"""
i18n.py
------------------------------------------------------------------------------
Module đa ngôn ngữ CHO GIAO DIỆN (menu, nhãn, form nhập) của HRM-Port.

NGUYÊN TẮC HIỂN THỊ (theo yêu cầu):
- Tiếng Việt LUÔN là ngôn ngữ chính — giữ nguyên, cỡ chữ bình thường.
- Ngôn ngữ phụ (nếu tenant chọn) hiển thị THÊM trong dấu ngoặc (), cỡ chữ NHỎ hơn.
- Không dịch/thay đổi bất kỳ giá trị nào dùng để so sánh logic trong code (vd. các
  chuỗi so sánh `if menu == "..."`, tên cột DB, mã trạng thái...). Chỉ đổi PHẦN HIỂN
  THỊ, luôn thông qua format_func / hàm t() / tm() — không sửa chuỗi gốc.

CÁCH DÙNG:
    import i18n
    i18n.set_active_language(tenant.get('ngon_ngu'))   # gọi 1 lần sau khi resolve tenant

    # Với label của widget Streamlit (button, selectbox, radio, text_input...):
    #   Widget label chỉ hỗ trợ Markdown cơ bản (KHÔNG hỗ trợ HTML/font-size), nên
    #   ngôn ngữ phụ được đặt trong ngoặc + in nghiêng để tạo cảm giác "phụ, nhỏ hơn".
    st.button(i18n.t("Lưu"))                 # -> "Lưu *(Save)*" nếu tenant chọn Việt-Anh

    # Với tiêu đề/văn bản qua st.markdown (HTML thật, cỡ chữ nhỏ hơn thật sự):
    st.markdown(i18n.tm("📊 Dashboard"), unsafe_allow_html=True)
    # -> "📊 Dashboard <span style='font-size:0.6em;color:#888;font-weight:400;'>(Bảng điều khiển)</span>"

Nếu tenant chọn "VI" (chỉ Tiếng Việt) hoặc chưa cấu hình, t()/tm() trả nguyên văn — không đổi gì.
Nếu 1 chuỗi chưa có trong từ điển TRANSLATIONS, t()/tm() cũng trả nguyên văn (không lỗi, không cảnh báo ồn ào)
để dễ mở rộng dần từng phần mà không sợ vỡ giao diện những phần chưa kịp dịch.
------------------------------------------------------------------------------
"""

import streamlit as st

# Các gói ngôn ngữ hỗ trợ khi Super Admin tạo/sửa tenant.
LANGUAGE_OPTIONS = {
    "VI":    "Tiếng Việt (chỉ 1 ngôn ngữ)",
    "VI_EN": "Việt - Anh  (Vietnamese - English)",
    "VI_ZH": "Việt - Trung (Vietnamese - Chinese 中文)",
    "VI_KO": "Việt - Hàn  (Vietnamese - Korean 한국어)",
}
DEFAULT_LANGUAGE = "VI"

# Mã ngôn ngữ phụ tương ứng mỗi gói (dùng để tra trong TRANSLATIONS bên dưới).
_SECONDARY_CODE = {
    "VI_EN": "EN",
    "VI_ZH": "ZH",
    "VI_KO": "KO",
}

# ──────────────────────────────────────────────────────────────────────────
# TỪ ĐIỂN — mở rộng dần theo thời gian. Key = chuỗi tiếng Việt GỐC dùng trong app.py
# (giữ đúng nguyên văn, kể cả icon, để tra cứu khớp 100%).
# ──────────────────────────────────────────────────────────────────────────
TRANSLATIONS = {
    # ---- Menu chính (sidebar) ----
    "📊 Dashboard":                       {"EN": "Dashboard",              "ZH": "仪表盘",         "KO": "대시보드"},
    "👤 Ứng viên":                        {"EN": "Candidates",             "ZH": "候选人",         "KO": "지원자"},
    "✅ Nhân viên":                       {"EN": "Employees",              "ZH": "员工",           "KO": "직원"},
    "📁 Upload hồ sơ":                    {"EN": "Upload Records",         "ZH": "上传档案",       "KO": "서류 업로드"},
    "⚙️ Danh mục":                        {"EN": "Categories",             "ZH": "类别管理",       "KO": "카테고리"},
    "📋 BHXH":                            {"EN": "Social Insurance",       "ZH": "社会保险",       "KO": "사회보험"},
    "📋 Báo cáo định kỳ":                 {"EN": "Periodic Reports",       "ZH": "定期报告",       "KO": "정기 보고서"},
    "🕒 Chấm công":                       {"EN": "Timekeeping",            "ZH": "考勤",           "KO": "근태관리"},
    "💰 Tính thu nhập":                   {"EN": "Payroll",                "ZH": "薪资计算",       "KO": "급여계산"},
    "📄 Quản lý Công văn & HĐ kinh tế":   {"EN": "Documents & Contracts",  "ZH": "公文与合同管理", "KO": "공문 및 계약 관리"},
    "💬 Chat nội bộ":                     {"EN": "Internal Chat",          "ZH": "内部聊天",       "KO": "사내 채팅"},
    "🤖 Chatbot Giải đáp":                {"EN": "Chatbot Assistant",      "ZH": "问答机器人",     "KO": "챗봇 도우미"},
    "🔑 Quản lý MK":                      {"EN": "Password/Access Mgmt",   "ZH": "密码/权限管理",  "KO": "비밀번호/권한 관리"},
    "🖼️ Tạo ảnh thẻ NV":                  {"EN": "Employee ID Photos",     "ZH": "员工证件照",     "KO": "직원 사진 생성"},
    "🔍 Audit Dashboard":                 {"EN": "Data Audit",             "ZH": "数据审计",       "KO": "데이터 감사"},
    "📘 Hướng dẫn sử dụng":               {"EN": "User Guide",             "ZH": "使用指南",       "KO": "사용 안내서"},

    # ---- Đăng nhập / công ty ----
    "🔐 Đăng nhập":                       {"EN": "Login",                  "ZH": "登录",           "KO": "로그인"},
    "Số điện thoại hoặc Tên đăng nhập":   {"EN": "Phone number or Username", "ZH": "电话号码或用户名", "KO": "전화번호 또는 아이디"},
    "Mật khẩu":                           {"EN": "Password",               "ZH": "密码",           "KO": "비밀번호"},
    "Đăng nhập":                          {"EN": "Log in",                 "ZH": "登录",           "KO": "로그인"},
    "🚪 Đăng xuất":                       {"EN": "Log out",                "ZH": "退出登录",       "KO": "로그아웃"},

    # ---- Nút thao tác dùng lặp lại nhiều nơi ----
    "💾 Lưu":                             {"EN": "Save",                   "ZH": "保存",           "KO": "저장"},
    "Lưu":                                {"EN": "Save",                   "ZH": "保存",           "KO": "저장"},
    "✏️ Sửa":                             {"EN": "Edit",                   "ZH": "编辑",           "KO": "수정"},
    "🗑️ Xóa":                             {"EN": "Delete",                 "ZH": "删除",           "KO": "삭제"},
    "❌ HỦY":                             {"EN": "Cancel",                 "ZH": "取消",           "KO": "취소"},
    "🔍 Tìm kiếm":                        {"EN": "Search",                 "ZH": "搜索",           "KO": "검색"},

    # ---- Nhãn form nhân viên hay dùng nhất ----
    "Họ và tên *":                        {"EN": "Full name *",            "ZH": "姓名 *",         "KO": "성명 *"},
    "Ngày sinh (dd/mm/yyyy)":             {"EN": "Date of birth (dd/mm/yyyy)", "ZH": "出生日期 (日/月/年)", "KO": "생년월일 (일/월/년)"},
    "Giới tính":                          {"EN": "Gender",                 "ZH": "性别",           "KO": "성별"},
    "CCCD":                               {"EN": "National ID",            "ZH": "身份证号",       "KO": "신분증 번호"},
    "SĐT":                                {"EN": "Phone",                  "ZH": "电话",           "KO": "전화번호"},
    "Chức danh":                          {"EN": "Job title",              "ZH": "职称",           "KO": "직책"},
    "Phòng ban":                          {"EN": "Department",             "ZH": "部门",           "KO": "부서"},
    "Trình độ":                           {"EN": "Education level",        "ZH": "学历",           "KO": "학력"},

    # ---- Tiêu đề trang (st.title đầu mỗi mục menu) ----
    "✅ Quản lý nhân viên":               {"EN": "Employee Management",    "ZH": "员工管理",       "KO": "직원 관리"},
    "💰 Tính thu nhập (Lương & Phụ cấp)": {"EN": "Payroll (Salary & Allowances)", "ZH": "薪资计算（工资与津贴）", "KO": "급여 계산 (급여 및 수당)"},
    "⚙️ Danh mục cấu hình theo doanh nghiệp": {"EN": "Company Configuration Categories", "ZH": "企业配置类别", "KO": "회사 설정 카테고리"},
    "📋 Quản lý BHXH":                    {"EN": "Social Insurance Management", "ZH": "社会保险管理", "KO": "사회보험 관리"},
}


def get_language_label(code: str) -> str:
    return LANGUAGE_OPTIONS.get(code or DEFAULT_LANGUAGE, LANGUAGE_OPTIONS[DEFAULT_LANGUAGE])


def set_active_language(ngon_ngu: str | None):
    """Gọi 1 lần (thường ngay sau khi xác định tenant) để nạp ngôn ngữ đang dùng
    vào session_state. Nếu không truyền / giá trị lạ -> mặc định 'VI' (chỉ Việt)."""
    code = (ngon_ngu or DEFAULT_LANGUAGE).strip().upper()
    if code not in LANGUAGE_OPTIONS:
        code = DEFAULT_LANGUAGE
    st.session_state["_ngon_ngu_active"] = code


def _secondary_text(vi_text: str) -> str | None:
    """Trả về bản dịch phụ (nếu có cấu hình ngôn ngữ phụ VÀ chuỗi này có trong từ điển)."""
    active = st.session_state.get("_ngon_ngu_active", DEFAULT_LANGUAGE)
    lang_code = _SECONDARY_CODE.get(active)
    if not lang_code:
        return None
    entry = TRANSLATIONS.get(vi_text)
    if not entry:
        return None
    return entry.get(lang_code)


def t(vi_text: str) -> str:
    """Dùng cho LABEL của widget Streamlit (button/selectbox/radio/text_input/tab...).
    Widget label của Streamlit chỉ hỗ trợ Markdown cơ bản (không hỗ trợ HTML/font-size),
    nên ngôn ngữ phụ được thêm trong ngoặc + in nghiêng để tạo cảm giác 'phụ, nhỏ hơn'."""
    phu = _secondary_text(vi_text)
    if not phu:
        return vi_text
    return f"{vi_text} *({phu})*"


def tm(vi_text: str) -> str:
    """Dùng cho nội dung hiển thị qua st.markdown(..., unsafe_allow_html=True)
    (tiêu đề trang, header...). Cho phép cỡ chữ THẬT SỰ nhỏ hơn cho ngôn ngữ phụ."""
    phu = _secondary_text(vi_text)
    if not phu:
        return vi_text
    return (
        f"{vi_text} "
        f"<span style='font-size:0.62em;font-weight:400;color:#888;'>({phu})</span>"
    )
