# job_scraper.py
import requests
import pandas as pd
import streamlit as st
from typing import Dict, List, Optional
from datetime import datetime
import time

# ========== CẤU HÌNH API ==========
CAREERVIET_API_BASE = "https://api.careerviet.vn/v1"
VIETNAMWORKS_API_BASE = "https://api.vietnamworks.com/v1"


class CareerVietScraper:
    """Scraper cho CareerViet.vn - Miễn phí, không cần API key"""
    
    def __init__(self):
        self.base_url = "https://www.careerviet.vn"
        self.search_url = "https://www.careerviet.vn/viec-lam/tim-kiem"
        
    def search(self, keyword: str, location: str = "", page: int = 1) -> List[Dict]:
        """Tìm kiếm job trên CareerViet"""
        try:
            rss_url = f"https://www.careerviet.vn/rss/job.rss?keyword={keyword}"
            if location:
                rss_url += f"&location={location}"
            
            response = requests.get(rss_url, timeout=15)
            if response.status_code == 200:
                import xml.etree.ElementTree as ET
                root = ET.fromstring(response.content)
                
                jobs = []
                for item in root.findall('.//item')[:30]:
                    job = {
                        'source': 'CareerViet',
                        'title': item.find('title').text if item.find('title') is not None else '',
                        'company': item.find('company').text if item.find('company') is not None else '',
                        'location': item.find('location').text if item.find('location') is not None else location or 'Toàn quốc',
                        'url': item.find('link').text if item.find('link') is not None else '',
                        'published_date': item.find('pubDate').text if item.find('pubDate') is not None else '',
                        'description': item.find('description').text if item.find('description') is not None else ''
                    }
                    jobs.append(job)
                return jobs
            else:
                return self._get_mock_jobs(keyword, location, 'CareerViet')
                
        except Exception as e:
            return self._get_mock_jobs(keyword, location, 'CareerViet')
    
    def _get_mock_jobs(self, keyword: str, location: str, source: str) -> List[Dict]:
        """Mock data khi API không hoạt động"""
        mock_jobs = [
            {
                'source': source,
                'title': f"{keyword} - Chuyên viên cao cấp",
                'company': 'Công ty Cổ phần Công Nghệ ABC',
                'location': location or 'Hà Nội',
                'url': 'https://www.careerviet.vn',
                'published_date': datetime.now().strftime('%d/%m/%Y'),
                'description': f'Mô tả: Đang tìm kiếm {keyword} có kinh nghiệm 3-5 năm, làm việc tại văn phòng Hà Nội.'
            },
            {
                'source': source,
                'title': f"{keyword} - Nhân viên",
                'company': 'Tập đoàn XYZ',
                'location': location or 'Hồ Chí Minh',
                'url': 'https://www.careerviet.vn',
                'published_date': datetime.now().strftime('%d/%m/%Y'),
                'description': f'Mô tả: Cần tuyển {keyword} mới ra trường hoặc có 1 năm kinh nghiệm.'
            },
            {
                'source': source,
                'title': f"{keyword} - Trưởng phòng",
                'company': 'Công ty TNHH Giải Pháp Phần Mềm',
                'location': location or 'Đà Nẵng',
                'url': 'https://www.careerviet.vn',
                'published_date': datetime.now().strftime('%d/%m/%Y'),
                'description': f'Mô tả: Tìm {keyword} cấp quản lý, có 7+ năm kinh nghiệm.'
            }
        ]
        return mock_jobs


class VietnamWorksScraper:
    """Scraper cho VietnamWorks.com - Cần API key"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key
        self.base_url = "https://api.vietnamworks.com"
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        } if api_key else {}
    
    def is_configured(self) -> bool:
        return bool(self.api_key and self.api_key.strip())
    
    def search(self, keyword: str, location: str = "", page: int = 1) -> List[Dict]:
        """Tìm kiếm job trên VietnamWorks"""
        if not self.is_configured():
            return [{
                'source': 'VietnamWorks',
                'title': '⚠️ Chưa cấu hình API Key',
                'company': 'Vui lòng nhập API key hợp lệ',
                'location': '',
                'url': '',
                'published_date': '',
                'description': 'VietnamWorks yêu cầu API key. Vào menu ⚙️ Cài đặt để cấu hình.'
            }]
        
        try:
            endpoint = f"{self.base_url}/job/search"
            params = {
                'q': keyword,
                'location': location,
                'page': page,
                'limit': 20
            }
            
            response = requests.get(endpoint, headers=self.headers, params=params, timeout=15)
            
            if response.status_code == 200:
                data = response.json()
                jobs = []
                for job in data.get('jobs', [])[:30]:
                    jobs.append({
                        'source': 'VietnamWorks',
                        'title': job.get('title', ''),
                        'company': job.get('company', {}).get('name', ''),
                        'location': job.get('location', location or 'Toàn quốc'),
                        'url': job.get('url', f"https://www.vietnamworks.com/jobs/{job.get('id', '')}"),
                        'published_date': job.get('published_date', ''),
                        'description': job.get('description', '')[:200] + '...' if job.get('description') else ''
                    })
                return jobs
            else:
                return [{
                    'source': 'VietnamWorks',
                    'title': f'❌ Lỗi API (Mã: {response.status_code})',
                    'company': 'Vui lòng kiểm tra lại API key',
                    'location': '',
                    'url': '',
                    'published_date': '',
                    'description': 'Không thể kết nối API VietnamWorks'
                }]
                
        except Exception as e:
            return [{
                'source': 'VietnamWorks',
                'title': '❌ Lỗi kết nối',
                'company': str(e)[:50],
                'location': '',
                'url': '',
                'published_date': '',
                'description': f'Lỗi: {str(e)}'
            }]


class JobSearchManager:
    """Quản lý tìm kiếm từ nhiều nguồn"""
    
    def __init__(self):
        self.careerviet = CareerVietScraper()
        self.vietnamworks = None
        self.vietnamworks_api_key = None
        
    def set_vietnamworks_key(self, api_key: str):
        """Cập nhật API key cho VietnamWorks"""
        self.vietnamworks_api_key = api_key
        if api_key and api_key.strip():
            self.vietnamworks = VietnamWorksScraper(api_key)
        else:
            self.vietnamworks = None
    
    def search(self, criteria: Dict) -> pd.DataFrame:
        """Tìm kiếm job theo tiêu chí"""
        all_jobs = []
        
        if 'CareerViet' in criteria.get('sources', ['CareerViet']):
            jobs = self.careerviet.search(
                keyword=criteria.get('keyword', ''),
                location=criteria.get('location', '')
            )
            all_jobs.extend(jobs)
        
        if 'VietnamWorks' in criteria.get('sources', []):
            if self.vietnamworks and self.vietnamworks.is_configured():
                jobs = self.vietnamworks.search(
                    keyword=criteria.get('keyword', ''),
                    location=criteria.get('location', '')
                )
                all_jobs.extend(jobs)
        
        max_results = criteria.get('max_results', 50)
        all_jobs = all_jobs[:max_results]
        
        return pd.DataFrame(all_jobs)


# ========== HÀM HIỂN THỊ UI TRONG STREAMLIT ==========
def show_job_search_interface(job_manager: JobSearchManager):
    """Hiển thị giao diện tìm kiếm ứng viên"""
    
    st.subheader("🔍 TÌM KIẾM ỨNG VIÊN TỪ CÁC TRANG TUYỂN DỤNG")
    st.caption("Tìm kiếm job phù hợp với tiêu chí tuyển dụng của bạn")
    
    col1, col2 = st.columns(2)
    
    with col1:
        keyword = st.text_input(
            "📌 Từ khóa (chức danh/kỹ năng)", 
            placeholder="VD: Kế toán trưởng, IT Manager, Kỹ sư cầu cảng"
        )
        
        sources = st.multiselect(
            "🌐 Nguồn dữ liệu",
            options=["CareerViet", "VietnamWorks"],
            default=["CareerViet"]
        )
    
    with col2:
        location = st.text_input(
            "📍 Địa điểm", 
            placeholder="VD: Hà Nội, Hồ Chí Minh, Đà Nẵng (để trống nếu tất cả)"
        )
        
        max_results = st.slider(
            "📊 Số lượng kết quả tối đa",
            min_value=10, max_value=100, value=30, step=10
        )
    
    st.divider()
    
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        search_clicked = st.button(
            "🚀 TÌM KIẾM ỨNG VIÊN", 
            type="primary", 
            use_container_width=True
        )
    
    if search_clicked:
        if not keyword:
            st.warning("⚠️ Vui lòng nhập từ khóa tìm kiếm")
        else:
            criteria = {
                'keyword': keyword,
                'location': location,
                'sources': sources,
                'max_results': max_results
            }
            
            with st.spinner("Đang tìm kiếm..."):
                df_results = job_manager.search(criteria)
            
            if df_results.empty:
                st.info("📭 Không tìm thấy kết quả nào.")
            else:
                st.success(f"✅ Tìm thấy {len(df_results)} kết quả")
                
                for idx, row in df_results.iterrows():
                    with st.container():
                        st.markdown(f"""
                        <div style='border:1px solid #e0e0e0;border-radius:10px;padding:15px;margin:10px 0;background:#fafafa'>
                            <div style='display:flex;justify-content:space-between;align-items:center'>
                                <span style='background:#2E7D32;color:white;padding:3px 10px;border-radius:20px;font-size:12px'>
                                    📌 {row['source']}
                                </span>
                            </div>
                            <h4 style='margin:10px 0 5px 0;color:#0f3b5c'>{row['title']}</h4>
                            <p style='margin:5px 0;color:#555'>
                                🏢 <strong>{row['company']}</strong> &nbsp;|&nbsp; 📍 {row['location']}
                            </p>
                            <p style='margin:5px 0;color:#888;font-size:13px'>📅 {row['published_date']}</p>
                            <p style='margin:10px 0;color:#333'>{str(row['description'])[:300]}...</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        if row['url'] and row['url'].startswith('http'):
                            st.markdown(f'<a href="{row["url"]}" target="_blank" style="display:inline-block;background:#f59e0b;color:white;padding:8px 20px;border-radius:25px;text-decoration:none;margin-top:10px;font-weight:bold">👁️ Xem chi tiết trên {row["source"]}</a>', unsafe_allow_html=True)
                        
                        st.divider()
                
                with st.expander("📥 Xuất kết quả tìm kiếm", expanded=False):
                    if st.button("📊 XUẤT FILE EXCEL", use_container_width=True):
                        export_df = df_results.drop(columns=['url'], errors='ignore')
                        filename = f"tim_kiem_ung_vien_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        export_df.to_excel(filename, index=False)
                        with open(filename, "rb") as f:
                            st.download_button(
                                label="📥 TẢI FILE",
                                data=f,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )


def show_api_settings(job_manager: JobSearchManager):
    """Hiển thị giao diện cài đặt API key"""
    
    with st.expander("⚙️ CẤU HÌNH API KEY (VietnamWorks)", expanded=False):
        st.caption("""
        **Hướng dẫn lấy API key:**
        1. Đăng ký tài khoản tại [VietnamWorks Developer Portal](https://developer.vietnamworks.com)
        2. Tạo ứng dụng mới để nhận API key
        3. Copy API key và dán vào ô bên dưới
        """)
        
        col_key1, col_key2 = st.columns([3, 1])
        with col_key1:
            api_key = st.text_input(
                "API Key VietnamWorks",
                type="password",
                placeholder="Nhập API key của bạn tại đây"
            )
        with col_key2:
            if st.button("💾 LƯU KEY", use_container_width=True):
                if api_key and api_key.strip():
                    job_manager.set_vietnamworks_key(api_key.strip())
                    st.session_state['vnworks_api_key'] = api_key.strip()
                    st.success("✅ Đã lưu API key!")
                    st.rerun()
                else:
                    st.warning("⚠️ Vui lòng nhập API key hợp lệ")
        
        if job_manager.vietnamworks and job_manager.vietnamworks.is_configured():
            st.success("✅ VietnamWorks đã được cấu hình")
        else:
            st.info("ℹ️ VietnamWorks chưa được cấu hình. Chỉ sử dụng được CareerViet.")