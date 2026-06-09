# job_scraper.py - TÌM ỨNG VIÊN (CV) từ các trang tuyển dụng
import requests
import pandas as pd
import streamlit as st
from typing import Dict, List, Optional
from datetime import datetime
import json

# ========== CẤU HÌNH ==========

class CandidateScraper:
    """Tìm kiếm ứng viên (CV) từ các nguồn"""
    
    def __init__(self):
        pass
    
    def search_candidates(self, keyword: str, location: str = "", experience: str = "") -> List[Dict]:
        """
        Tìm kiếm ứng viên theo từ khóa
        Trả về danh sách ứng viên tiềm năng
        """
        # Demo data - trong thực tế sẽ gọi API từ các trang tuyển dụng
        mock_candidates = self._get_mock_candidates(keyword, location, experience)
        return mock_candidates
    
    def _get_mock_candidates(self, keyword: str, location: str, experience: str) -> List[Dict]:
        """Dữ liệu mẫu ứng viên - thay thế bằng API thật sau"""
        
        candidates = [
            {
                'source': 'TopCV',
                'ho_ten': 'Nguyễn Văn An',
                'vi_tri_ung_tuyen': f'{keyword} - Chuyên viên cao cấp',
                'kinh_nghiem': '5 năm',
                'location': location or 'Hà Nội',
                'ky_nang': 'Python, SQL, Data Analysis, Team Management',
                'hoc_van': 'Đại học Bách Khoa Hà Nội - CNTT',
                'muc_luong_mong_muon': '25-30 triệu',
                'trang_thai': 'Đang tìm việc',
                'url_cv': 'https://example.com/cv/1',
                'last_active': datetime.now().strftime('%d/%m/%Y'),
                'dien_thoai': '0987xxx123',
                'email': 'an.nguyen@email.com'
            },
            {
                'source': 'TopCV',
                'ho_ten': 'Trần Thị Bình',
                'vi_tri_ung_tuyen': f'{keyword} - Trưởng phòng',
                'kinh_nghiem': '8 năm',
                'location': location or 'Hồ Chí Minh',
                'ky_nang': 'Leadership, Strategy, Project Management, Agile',
                'hoc_van': 'Đại học Kinh tế TP.HCM - Quản trị Kinh doanh',
                'muc_luong_mong_muon': '40-50 triệu',
                'trang_thai': 'Đang tìm việc',
                'url_cv': 'https://example.com/cv/2',
                'last_active': datetime.now().strftime('%d/%m/%Y'),
                'dien_thoai': '0912xxx456',
                'email': 'binh.tran@email.com'
            },
            {
                'source': 'CareerBuilder',
                'ho_ten': 'Lê Văn Cường',
                'vi_tri_ung_tuyen': f'{keyword} - Nhân viên',
                'kinh_nghiem': '2 năm',
                'location': location or 'Đà Nẵng',
                'ky_nang': 'HTML, CSS, JavaScript, React, Teamwork',
                'hoc_van': 'Đại học Duy Tân - Công nghệ thông tin',
                'muc_luong_mong_muon': '12-15 triệu',
                'trang_thai': 'Có thể bắt đầu ngay',
                'url_cv': 'https://example.com/cv/3',
                'last_active': datetime.now().strftime('%d/%m/%Y'),
                'dien_thoai': '0905xxx789',
                'email': 'cuong.le@email.com'
            },
            {
                'source': 'VietnamWorks',
                'ho_ten': 'Phạm Thị Dung',
                'vi_tri_ung_tuyen': f'{keyword} - Chuyên viên',
                'kinh_nghiem': '3 năm',
                'location': location or 'Hà Nội',
                'ky_nang': 'Marketing, Content, SEO, Social Media',
                'hoc_van': 'Đại học Ngoại thương - Marketing',
                'muc_luong_mong_muon': '18-22 triệu',
                'trang_thai': 'Đang tìm việc',
                'url_cv': 'https://example.com/cv/4',
                'last_active': datetime.now().strftime('%d/%m/%Y'),
                'dien_thoai': '0976xxx234',
                'email': 'dung.pham@email.com'
            },
            {
                'source': 'VietnamWorks',
                'ho_ten': 'Hoàng Văn Em',
                'vi_tri_ung_tuyen': f'{keyword} - Kỹ sư',
                'kinh_nghiem': '4 năm',
                'location': location or 'Hồ Chí Minh',
                'ky_nang': 'AutoCAD, SolidWorks, Mechanical Design',
                'hoc_van': 'Đại học Bách Khoa - Cơ khí',
                'muc_luong_mong_muon': '20-25 triệu',
                'trang_thai': 'Sẵn sàng phỏng vấn',
                'url_cv': 'https://example.com/cv/5',
                'last_active': datetime.now().strftime('%d/%m/%Y'),
                'dien_thoai': '0933xxx567',
                'email': 'em.hoang@email.com'
            }
        ]
        
        # Lọc theo từ khóa nếu có
        if keyword:
            keyword_lower = keyword.lower()
            candidates = [c for c in candidates if 
                         keyword_lower in c['vi_tri_ung_tuyen'].lower() or 
                         keyword_lower in c['ky_nang'].lower()]
        
        # Lọc theo kinh nghiệm nếu có
        if experience and experience != "Tất cả":
            exp_filter = {
                "Dưới 1 năm": (0, 1),
                "1-3 năm": (1, 3),
                "3-5 năm": (3, 5),
                "5-7 năm": (5, 7),
                "Trên 7 năm": (7, 100)
            }
            if experience in exp_filter:
                min_exp, max_exp = exp_filter[experience]
                candidates = [c for c in candidates if 
                             min_exp <= int(c['kinh_nghiem'].split()[0]) <= max_exp]
        
        return candidates


class CandidateSearchManager:
    """Quản lý tìm kiếm ứng viên"""
    
    def __init__(self):
        self.scraper = CandidateScraper()
    
    def search(self, criteria: Dict) -> pd.DataFrame:
        """
        Tìm kiếm ứng viên theo tiêu chí
        criteria = {
            'keyword': str,           # Từ khóa vị trí/kỹ năng
            'location': str,          # Địa điểm
            'experience': str,        # Kinh nghiệm
            'sources': List[str],     # Nguồn tìm kiếm
            'max_results': int        # Số lượng tối đa
        }
        """
        candidates = self.scraper.search_candidates(
            keyword=criteria.get('keyword', ''),
            location=criteria.get('location', ''),
            experience=criteria.get('experience', '')
        )
        
        max_results = criteria.get('max_results', 50)
        candidates = candidates[:max_results]
        
        return pd.DataFrame(candidates)


# ========== HÀM HIỂN THỊ UI TRONG STREAMLIT ==========
def show_job_search_interface(job_manager):
    """Hiển thị giao diện tìm kiếm ứng viên"""
    
    st.subheader("🔍 TÌM KIẾM ỨNG VIÊN TIỀM NĂNG")
    st.caption("Tìm kiếm CV phù hợp với nhu cầu tuyển dụng của bạn")
    
    # Bộ lọc tìm kiếm
    col1, col2 = st.columns(2)
    
    with col1:
        keyword = st.text_input(
            "📌 Vị trí / Kỹ năng cần tuyển", 
            placeholder="VD: Kế toán, IT, Kỹ sư cầu cảng, Nhân viên kinh doanh...",
            help="Nhập chức danh hoặc kỹ năng cần tìm"
        )
        
        sources = st.multiselect(
            "🌐 Nguồn dữ liệu ứng viên",
            options=["TopCV", "VietnamWorks", "CareerBuilder", "LinkedIn"],
            default=["TopCV", "VietnamWorks"],
            help="Chọn nguồn trang tuyển dụng để tìm kiếm CV"
        )
    
    with col2:
        location = st.text_input(
            "📍 Địa điểm làm việc", 
            placeholder="VD: Hà Nội, Hồ Chí Minh, Đà Nẵng (để trống nếu tất cả)",
            help="Để trống để tìm kiếm toàn quốc"
        )
        
        experience = st.selectbox(
            "⏳ Kinh nghiệm",
            options=["Tất cả", "Dưới 1 năm", "1-3 năm", "3-5 năm", "5-7 năm", "Trên 7 năm"],
            index=0
        )
        
        max_results = st.slider(
            "📊 Số lượng kết quả tối đa",
            min_value=10, max_value=100, value=30, step=10
        )
    
    st.divider()
    
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        search_clicked = st.button(
            "🚀 TÌM ỨNG VIÊN", 
            type="primary", 
            use_container_width=True
        )
    
    if search_clicked:
        if not keyword:
            st.warning("⚠️ Vui lòng nhập vị trí hoặc kỹ năng cần tuyển")
        else:
            criteria = {
                'keyword': keyword,
                'location': location,
                'experience': experience,
                'sources': sources,
                'max_results': max_results
            }
            
            with st.spinner("Đang tìm kiếm ứng viên phù hợp..."):
                df_results = job_manager.search(criteria)
            
            if df_results.empty:
                st.info("📭 Không tìm thấy ứng viên nào phù hợp. Vui lòng thử từ khóa khác.")
            else:
                st.success(f"✅ Tìm thấy {len(df_results)} ứng viên phù hợp")
                
                # Hiển thị từng ứng viên
                for idx, row in df_results.iterrows():
                    with st.container():
                        # Xác định màu nền theo nguồn
                        bg_color = "#FFF8E1" if row['source'] == "TopCV" else "#E3F2FD" if row['source'] == "VietnamWorks" else "#F3E5F5"
                        
                        st.markdown(f"""
                        <div style='border:1px solid #ddd;border-radius:12px;padding:16px;margin:12px 0;background:{bg_color};box-shadow:0 2px 4px rgba(0,0,0,0.05)'>
                            <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
                                <div>
                                    <span style='background:#2E7D32;color:white;padding:2px 8px;border-radius:15px;font-size:11px'>
                                        📌 {row['source']}
                                    </span>
                                    <span style='background:#FF9800;color:white;padding:2px 8px;border-radius:15px;font-size:11px;margin-left:5px'>
                                        🟢 {row['trang_thai']}
                                    </span>
                                </div>
                                <span style='color:#888;font-size:12px'>📅 Cập nhật: {row['last_active']}</span>
                            </div>
                            <h3 style='margin:8px 0 4px 0;color:#0f3b5c'>👤 {row['ho_ten']}</h3>
                            <p style='margin:4px 0;color:#f59e0b;font-weight:bold'>🎯 {row['vi_tri_ung_tuyen']}</p>
                            <div style='display:flex;flex-wrap:wrap;gap:16px;margin:10px 0;padding:8px 0;border-top:1px dashed #ccc;border-bottom:1px dashed #ccc'>
                                <span>📍 {row['location']}</span>
                                <span>⏳ {row['kinh_nghiem']} kinh nghiệm</span>
                                <span>💰 {row['muc_luong_mong_muon']}</span>
                            </div>
                            <div style='margin:8px 0'>
                                <span style='font-weight:bold'>🎓 Học vấn:</span> {row['hoc_van']}<br>
                                <span style='font-weight:bold'>⚡ Kỹ năng:</span> {row['ky_nang']}
                            </div>
                            <div style='margin:8px 0;background:white;padding:8px;border-radius:8px'>
                                <span style='font-weight:bold'>📞 Liên hệ:</span> {row['dien_thoai']} &nbsp;|&nbsp;
                                <span style='font-weight:bold'>✉️ Email:</span> {row['email']}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
                        with col_btn1:
                            if st.button(f"📥 Lưu ứng viên", key=f"save_{idx}"):
                                try:
                                    from app import get_connection
                                    db = get_connection()
                                    c = db.cursor()
                                    c.execute("""
                                        INSERT INTO ung_vien (ho_ten, vi_tri_du_tuyen, dien_thoai, email, luong_bao_hiem, trang_thai, ghi_chu)
                                        VALUES (%s, %s, %s, %s, %s, 'CHO_DUYET', %s)
                                    """, (
                                        row['ho_ten'], row['vi_tri_ung_tuyen'], row['dien_thoai'], 
                                        row['email'], row['muc_luong_mong_muon'], f"Từ {row['source']} - {row['ky_nang']}"
                                    ))
                                    db.commit()
                                    db.close()
                                    st.success(f"✅ Đã lưu {row['ho_ten']} vào danh sách ứng viên!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Lỗi: {e}")
                        with col_btn2:
                            st.markdown(f'<a href="{row["url_cv"]}" target="_blank"><button style="background:#2196F3;color:white;border:none;padding:6px 12px;border-radius:20px;cursor:pointer">👁️ Xem CV</button></a>', unsafe_allow_html=True)
                        
                        st.divider()
                
                # Xuất Excel
                with st.expander("📥 Xuất danh sách ứng viên", expanded=False):
                    st.caption("Xuất danh sách ứng viên tìm được ra file Excel")
                    if st.button("📊 XUẤT FILE EXCEL", use_container_width=True):
                        export_df = df_results.drop(columns=['url_cv'], errors='ignore')
                        filename = f"danh_sach_ung_vien_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        export_df.to_excel(filename, index=False)
                        with open(filename, "rb") as f:
                            st.download_button(
                                label="📥 TẢI FILE",
                                data=f,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )


def show_api_settings(job_manager):
    """Hiển thị giao diện cài đặt API (nếu cần)"""
    with st.expander("⚙️ CẤU HÌNH NGUỒN DỮ LIỆU", expanded=False):
        st.info("""
        **📌 Thông báo:**
        
        Hiện tại tính năng đang ở chế độ Demo với dữ liệu mẫu.
        
        **Trong tương lai sẽ tích hợp:**
        - 🔗 Kết nối API TopCV
        - 🔗 Kết nối API VietnamWorks  
        - 🔗 Import CV từ file PDF/Word
        - 🔗 Quét CV từ email
        
        Vui lòng liên hệ IT để được cấu hình API chính thức.
        """)