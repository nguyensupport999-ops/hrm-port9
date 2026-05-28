import streamlit as st
import psycopg2
import psycopg2.extras
from config import COMPANY_CONFIG
from datetime import datetime, date, timedelta
import os
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import tempfile
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import requests
from openpyxl.styles import Font, Alignment
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import qrcode
from io import BytesIO
from config_template import COMPANY_CONFIG, BHXH_CONFIG, EMAIL_CONFIG, TELEGRAM_CONFIG, USERS

# ========== DATABASE CONNECTION (SUPABASE) ==========
# ========== DATABASE CONNECTION (SUPABASE) ==========
def get_connection():
    import os
    import psycopg2
    
    # Thử đọc từ st.secrets trước (Streamlit Cloud)
    try:
        import streamlit as st
        if 'connections' in st.secrets and 'supabase' in st.secrets.connections:
            return psycopg2.connect(
                host=st.secrets.connections.supabase.host,
                port=st.secrets.connections.supabase.port,
                user=st.secrets.connections.supabase.user,
                password=st.secrets.connections.supabase.password,
                database=st.secrets.connections.supabase.database
            )
    except:
        pass
    
    # Fallback: đọc từ file .env (local)
    from dotenv import load_dotenv
    load_dotenv()
    
    return psycopg2.connect(
        host=os.getenv('DB_HOST'),
        port=os.getenv('DB_PORT'),
        user=os.getenv('DB_USER'),
        password=os.getenv('DB_PASSWORD'),
        database=os.getenv('DB_NAME')
    )

# Chèn logo vào sidebar
logo_path = "logo_cty.png"
if os.path.exists(logo_path):
    with st.sidebar:
        st.image(logo_path, use_container_width=True)
        st.divider()

st.set_page_config(page_title="HRM-Port", page_icon="🏗️", layout="wide")

def force_center(p):
    pPr = p._p.get_or_add_pPr()
    for jc in pPr.findall(qn('w:jc')):
        pPr.remove(jc)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)

st.markdown("""
<style>
    [data-testid="stDataFrame"] > div {
        overflow-x: auto !important;
    }
    [data-testid="stDataFrame"] table {
        min-width: 2000px !important;
        width: max-content !important;
    }
</style>
""", unsafe_allow_html=True)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def format_date(d):
    if d is None or pd.isna(d): return ''
    try: return d.strftime('%d/%m/%Y') if hasattr(d,'strftime') else str(d)[:10]
    except: return str(d)

def parse_date(s):
    if not s or s.strip()=='': return None
    try: return datetime.strptime(s.strip(),'%d/%m/%Y').date()
    except: return None

def tao_noi_dung_zalo(nv):
    ZC = COMPANY_CONFIG
    return f"""Gửi anh/chị: {nv.get('ho_ten','')},

Thông tin đã cập nhật:
- Họ tên: {nv.get('ho_ten','')}
- Ngày sinh: {format_date(nv.get('ngay_sinh'))}
- CCCD: {nv.get('so_cccd','')}
- Ngày cấp: {format_date(nv.get('ngay_cap_cccd'))}
- Thường trú: {nv.get('thuong_tru','')}
- Số BHXH: {nv.get('ma_so_bhxh','')}
- TK NH: {nv.get('so_tai_khoan_nh','')}
- CN NH: {nv.get('chi_nhanh_nh','')}

{ZC.get('loi_nhan_zalo','Vui lòng kiểm tra và phản hồi nếu có sai sót. Xin Cảm ơn!')}"""

def remove_table_border(tbl):
    for row in tbl.rows:
        for cell in row.cells:
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            b = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcBorders')
            if b is not None: tcPr.remove(b)

# ========== CÁC HÀM TẠO HỢP ĐỒNG (GIỮ NGUYÊN) ==========
def tao_hop_dong(nv):
    # ... giữ nguyên code cũ (quá dài, tôi giữ lại)
    CC = COMPANY_CONFIG; doc = Document()
    s = doc.styles['Normal']; s.font.name='Times New Roman'; s.font.size=Pt(13)
    s.paragraph_format.space_after=Pt(0); s.paragraph_format.space_before=Pt(0)
    sec = doc.sections[0]; sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(3.5); sec.right_margin=Cm(2)
    def add_p(text='', bold=False, size=Pt(13)):
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        if bold and p.runs:
            p.runs[0].bold = True
        if p.runs:
            p.runs[0].font.size = size
        return p
    def al(label,value):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(5))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r=p.add_run(f'{label}'); r.font.size=Pt(13)
        r=p.add_run('\t: '); r.font.size=Pt(13)
        r=p.add_run(f'{value}'); r.font.size=Pt(13)
    ht=doc.add_table(rows=4,cols=2); ht.alignment=WD_TABLE_ALIGNMENT.CENTER; ht.autofit=False; remove_table_border(ht)
    for row in ht.rows: row.cells[0].width=Cm(6); row.cells[1].width=Cm(11)
    c=ht.rows[0].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CÔNG TY CỔ PHẦN'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[0].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CẢNG HÒN LA'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Độc lập - Tự do - Hạnh phúc'); r.bold=True; r.italic=True; r.font.size=Pt(13)
    c=ht.rows[2].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*12); r.font.size=Pt(9)
    c=ht.rows[2].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*20); r.font.size=Pt(9)
    c=ht.rows[3].cells[1]; p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(20)
    ngay_ky = nv.get("ngay_ky_hd")
    ngay_vao = nv.get("ngay_vao_lam")
    nk = ngay_ky if ngay_ky else ngay_vao
    ns = 'Quảng Trị, ngày ... tháng ... năm ......'
    if nk and hasattr(nk, 'day'):
        ns = f'Quảng Trị, ngày {nk.day} tháng {nk.month:02d} năm {nk.year}'
    run = p.add_run(ns)
    run.font.size = Pt(13)
    run.italic = True
    c=ht.rows[3].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f'Số: {nv.get("so_hdld","...")}'); r.italic=True; r.font.size=Pt(12)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run('HỢP ĐỒNG LAO ĐỘNG')
    r.bold = True
    r.font.size = Pt(18)
    force_center(p)
    p2 = doc.add_paragraph('- Căn cứ thông tư 10/2020/TT-LĐTBXH ngày 12/11/2020 hướng dẫn thi hành một số điều của Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019 về nội dung của hợp đồng lao động;')
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p2 = doc.add_paragraph('- Căn cứ nhu cầu sử dụng lao động trong đơn vị.')
    doc.add_paragraph('Chúng tôi gồm:')
    p=doc.add_paragraph(); r=p.add_run(f'BÊN A: {CC["ten_cong_ty"]} (Người sử dụng LĐ)'); r.bold=True
    al('Đại diện',f"Ông {CC['dai_dien']}"); al('Chức vụ',CC['chuc_vu']); al('Mã số thuế',CC['ma_so_thue'])
    al('Điện thoại',CC['dien_thoai_cty']); al('Địa chỉ',CC['dia_chi']); doc.add_paragraph()
    p=doc.add_paragraph(); r=p.add_run('BÊN B: (Người lao động)'); r.bold=True
    sk=nv.get('so_tai_khoan_nh','')
    if nv.get('chi_nhanh_nh'): sk+=f' - {nv.get("chi_nhanh_nh")}'
    gt = nv.get('gioi_tinh','')
    xung_ho = 'Ông' if gt == 'Nam' else ('Bà' if gt == 'Nữ' else 'Ông/Bà')
    al(xung_ho, nv.get('ho_ten',''))
    al('Ngày sinh',format_date(nv.get('ngay_sinh')))
    al('Số CMND/CCCD',nv.get('so_cccd','')); al('Ngày cấp',format_date(nv.get('ngay_cap_cccd')))
    al('Nơi cấp',nv.get('noi_cap_cccd','')); al('Số TKNH',sk)
    al('Điện thoại',nv.get('dien_thoai','')); al('Thường trú',nv.get('thuong_tru',''))
    doc.add_paragraph('Thoả thuận ký kết Hợp đồng lao động với những điều khoản dưới đây:')
    p=doc.add_paragraph(); r=p.add_run('Điều 1. Thời hạn và công việc hợp đồng:'); r.bold=True
    ngay_hieu_luc = nv.get("ngay_ky_hd") or nv.get("ngay_vao_lam")
    ns2 = '.../.../..........'
    if ngay_hieu_luc and hasattr(ngay_hieu_luc, 'day'):
        ns2 = f'{ngay_hieu_luc.day} tháng {ngay_hieu_luc.month:02d} năm {ngay_hieu_luc.year}'
    elif ngay_hieu_luc:
        ns2 = str(ngay_hieu_luc)
    add_p(f'-    Bên B làm việc theo chế độ hợp đồng lao động không xác định thời hạn;')
    add_p(f'-    Thời gian: Từ ngày {ns2};')
    add_p('-    Địa điểm làm việc: Tại Cảng tổng hợp quốc tế Hòn La và các địa điểm khác theo sự sắp xếp của Công ty;')
    add_p(f'-    Vị trí: {nv.get("chuc_danh_nghe","")};')
    add_p('-    Công việc phải làm: Thực hiện công việc theo đúng chuyên môn dưới sự quản lý, điều hành của cấp trên;')
    add_p('-    Mức lương và phụ cấp: Theo thỏa thuận;')
    add_p('-    Hình thức trả lương: Tiền mặt hoặc chuyển khoản, theo lần chi trả;')
    add_p('-    Kỳ hạn trả lương: Theo quy định Công ty;')
    add_p('-    Chế độ nâng lương: Theo thỏa thuận.')
    p=doc.add_paragraph(); r=p.add_run('Điều 2. Chế độ làm việc:'); r.bold=True
    add_p('-    Thời gian làm việc: Theo tính chất công việc, do nhu cầu kinh doanh của Công ty nên thời gian làm việc của bên B là linh hoạt nhưng phải đảm bảo hoàn thành công việc được giao;')
    add_p('-    Thời gian nghỉ ngơi của người lao động: Theo thỏa thuận và phù hợp với quy định của pháp luật;')
    add_p('-    Ngoài giờ làm việc: Người lao động phải tự chịu trách nhiệm về các hoạt động cá nhân của mình.')
    p=doc.add_paragraph(); r=p.add_run('Điều 3. Nghĩa vụ, quyền lợi NLĐ:'); r.bold=True
    p=doc.add_paragraph(); r=p.add_run('1. Nghĩa vụ:'); r.bold=True
    add_p('-    Hoàn thành những công việc được giao và sẵn sàng chấp nhận mọi sự điều động khi có yêu cầu;')
    add_p('-    Chấp hành nghiêm túc nội quy, kỷ luật lao động, an toàn lao động và các quy định của Công ty và pháp luật của Nhà nước;')
    add_p('-    Người lao động có trách nhiệm tuân thủ đầy đủ quy định về an toàn lao động, quy trình vận hành thiết bị và hướng dẫn của Công ty. Trường hợp NLĐ cố ý vi phạm hoặc vi phạm nghiêm trọng quy định an toàn lao động gây thiệt hại thì phải chịu trách nhiệm theo quy định pháp luật và nội quy Công ty;')
    add_p('-    Bồi thường vi phạm vật chất : Phải bồi thường vật chất do cá nhân vi phạm quy định của Công ty về bảo quản trang thiết bị được giao.')
    p=doc.add_paragraph(); r=p.add_run('2. Quyền Lợi:'); r.bold=True
    add_p('-    Phương tiện đi lại: Tự túc;')
    add_p('-    Được Công ty đóng Bảo hiểm xã hội, bảo hiểm y tế, BHTN: theo chế độ hiện hành của Nhà nước và Quy định của Công ty;')
    add_p('-    Được Công ty cấp đầy đủ bảo hộ lao động theo đúng vị trí làm việc;')
    add_p('-    Được phân công công việc theo yêu cầu của Công ty phù hợp với khả năng và trình độ chuyên môn mà người lao động đáp ứng;')
    add_p('-    Các quyền lợi khác thực hiện theo quy định của Pháp luật Lao động như tạm dừng, chấm dứt hợp đồng.')
    p=doc.add_paragraph(); r=p.add_run('Điều 4. Nghĩa vụ, quyền hạn NSDLĐ:'); r.bold=True
    add_p('-    Bảo đảm việc làm và thực hiện đầy đủ những điều đã cam kết trong hợp đồng;')
    add_p('-    Thanh toán đầy đủ, đúng hạn các chế độ và quyền lợi cho người lao động theo hợp đồng;')
    add_p('-    Điều hành người lao động hoàn thành công việc theo hợp đồng;')
    add_p('-    Tạm hoãn, chấm dứt hợp đồng, kỷ luật người lao động theo quy định của pháp luật, và nội quy lao động của Công ty.')
    p=doc.add_paragraph(); r=p.add_run('Điều 5. Điều khoản chung:'); r.bold=True
    add_p('-    Những nội dung về quan hệ lao động không ghi trong hợp đồng này thì được áp dụng theo pháp luật lao động;')
    add_p('-    Những thoả thuận khác (nếu có): không;')
    add_p('-    Hợp đồng này có hiệu lực từ ngày ký và được làm thành 02 bản, Bên A giữ 01 bản, Bên B giữ 01 có giá trị pháp lý như nhau, để làm căn cứ thực hiện.')
    add_p('Bản HĐ này lập tại văn phòng Công ty CP Cảng Hòn La.'); doc.add_paragraph()
    ts=doc.add_table(rows=3,cols=2); ts.alignment=WD_TABLE_ALIGNMENT.CENTER; remove_table_border(ts)
    c=ts.rows[0].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[0].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI SỬ DỤNG LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[1].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[1].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[2].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(nv.get('ho_ten','').upper()); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[2].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(CC['dai_dien'].upper()); r.bold=True; r.font.size=Pt(13)
    tf=tempfile.NamedTemporaryFile(delete=False,suffix='.docx'); doc.save(tf.name); return tf.name

def tao_hop_dong_thu_viec(nv):
    # ... giữ nguyên code cũ (tương tự)
    CC = COMPANY_CONFIG; doc = Document()
    s = doc.styles['Normal']; s.font.name='Times New Roman'; s.font.size=Pt(13)
    s.paragraph_format.space_after=Pt(0); s.paragraph_format.space_before=Pt(0)
    s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sec = doc.sections[0]; sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(3.5); sec.right_margin=Cm(2)
    def al(label,value):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(5))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r=p.add_run(f'{label}'); r.font.size=Pt(13)
        r=p.add_run('\t: '); r.font.size=Pt(13)
        r=p.add_run(f'{value}'); r.font.size=Pt(13)
    ht=doc.add_table(rows=4,cols=2); ht.alignment=WD_TABLE_ALIGNMENT.CENTER; ht.autofit=False; remove_table_border(ht)
    for row in ht.rows: row.cells[0].width=Cm(6); row.cells[1].width=Cm(11)
    c=ht.rows[0].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CÔNG TY CỔ PHẦN'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[0].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('CẢNG HÒN LA'); r.bold=True; r.font.size=Pt(13)
    c=ht.rows[1].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('Độc lập - Tự do - Hạnh phúc'); r.bold=True; r.italic=True; r.font.size=Pt(13)
    c=ht.rows[2].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*12); r.font.size=Pt(9)
    c=ht.rows[2].cells[1]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run('─'*20); r.font.size=Pt(9)
    c=ht.rows[3].cells[0]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(f'Số: {nv.get("so_hdld",".../.../HĐTV-CHL")}'); r.italic=True; r.font.size=Pt(12)
    c=ht.rows[3].cells[1]; p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; p.paragraph_format.space_after=Pt(20)
    nk=nv.get("ngay_vao_lam") or nv.get("ngay_ky_hd")
    ns='Quảng Trị, ngày ... tháng ... năm ......'
    if nk and hasattr(nk,'day'): ns=f'Quảng Trị, ngày {nk.day} tháng {nk.month:02d} năm {nk.year}'
    run = p.add_run(ns)
    run.font.size = Pt(13)
    run.italic = True
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(20)
    r = p.add_run('HỢP ĐỒNG THỬ VIỆC')
    r.bold = True
    r.font.size = Pt(18)
    force_center(p)
    p2 = doc.add_paragraph('- Căn cứ thông tư 10/2020/TT-LĐTBXH ngày 12/11/2020 hướng dẫn thi hành một số điều của Bộ luật Lao động số 45/2019/QH14 ngày 20/11/2019 về nội dung của hợp đồng lao động;')
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph('- Căn cứ nhu cầu sử dụng lao động trong đơn vị.')
    doc.add_paragraph('Chúng tôi gồm:')
    p=doc.add_paragraph(); r=p.add_run(f'BÊN A: {CC["ten_cong_ty"]} (Người sử dụng LĐ)'); r.bold=True
    al('Đại diện',f"Ông {CC['dai_dien']}"); al('Chức vụ',CC['chuc_vu']); al('Mã số thuế',CC['ma_so_thue'])
    al('Điện thoại',CC['dien_thoai_cty']); al('Địa chỉ',CC['dia_chi'])
    p=doc.add_paragraph(); r=p.add_run('BÊN B: (Người lao động)'); r.bold=True
    sk=nv.get('so_tai_khoan_nh','')
    if nv.get('chi_nhanh_nh'): sk+=f' - {nv.get("chi_nhanh_nh")}'
    gt = nv.get('gioi_tinh','')
    xung_ho = 'Ông' if gt == 'Nam' else ('Bà' if gt == 'Nữ' else 'Ông/Bà')
    al(xung_ho, nv.get('ho_ten',''))
    al('Ngày sinh',format_date(nv.get('ngay_sinh')))
    al('Số CMND/CCCD',nv.get('so_cccd','')); al('Ngày cấp',format_date(nv.get('ngay_cap_cccd')))
    al('Nơi cấp',nv.get('noi_cap_cccd','')); al('Số TKNH',sk)
    al('Điện thoại',nv.get('dien_thoai','')); al('Thường trú',nv.get('thuong_tru',''))
    doc.add_paragraph('Thoả thuận ký kết Hợp đồng Thử việc với những điều khoản dưới đây:')
    p=doc.add_paragraph(); r=p.add_run('Điều 1. Thời hạn và công việc hợp đồng:'); r.bold=True
    ns2='.../.../..........'
    if nk and hasattr(nk,'day'): ns2=f'{nk.day} tháng {nk.month} năm {nk.year}'
    elif nk: ns2=str(nk)
    p=doc.add_paragraph(f'-    Bên B làm việc theo chế độ hợp đồng thử việc, có thời hạn 01 tháng;')
    if nk and hasattr(nk,'day'):
        nkt=nk+timedelta(days=30)
        p=doc.add_paragraph(f'     + Bắt đầu: {nk.day:02d}/{nk.month:02d}/{nk.year}')
        p=doc.add_paragraph(f'     + Kết thúc: {nkt.day:02d}/{nkt.month:02d}/{nkt.year}')
    else: doc.add_paragraph('     + Bắt đầu: .../.../......'); doc.add_paragraph('     + Kết thúc: .../.../......')
    p=doc.add_paragraph('-    Địa điểm làm việc: Tại Cảng tổng hợp quốc tế Hòn La và các địa điểm khác theo sự sắp xếp của Công ty;')
    p=doc.add_paragraph(f'-    Vị trí: {nv.get("chuc_danh_nghe","")};')
    p=doc.add_paragraph('-    Công việc phải làm: Thực hiện công việc theo đúng chuyên môn dưới sự quản lý, điều hành của cấp trên;')
    p=doc.add_paragraph('-    Mức lương và phụ cấp: Theo thỏa thuận;')
    p=doc.add_paragraph('-    Hình thức trả lương: Tiền mặt hoặc chuyển khoản, theo lần chi trả;')
    p=doc.add_paragraph('-    Kỳ hạn trả lương: Theo quy định Công ty;')
    p=doc.add_paragraph(); r=p.add_run('Điều 2. Chế độ làm việc:'); r.bold=True
    p=doc.add_paragraph('-    Thời gian làm việc: Theo tính chất công việc, do nhu cầu kinh doanh của Công ty nên thời gian làm việc của bên B là linh hoạt nhưng phải đảm bảo hoàn thành công việc được giao;')
    p=doc.add_paragraph('-    Thời gian nghỉ ngơi của người lao động: Theo thỏa thuận và phù hợp với quy định của pháp luật;')
    p=doc.add_paragraph('-    Ngoài giờ làm việc: Người lao động phải tự chịu trách nhiệm về các hoạt động cá nhân của mình.')
    p=doc.add_paragraph(); r=p.add_run('Điều 3. Nghĩa vụ, quyền lợi NLĐ:'); r.bold=True
    p=doc.add_paragraph(); r=p.add_run('1. Nghĩa vụ:'); r.bold=True
    p=doc.add_paragraph('-    Hoàn thành những công việc được giao và sẵn sàng chấp nhận mọi sự điều động khi có yêu cầu;')
    p=doc.add_paragraph('-    Chấp hành nghiêm túc nội quy, kỷ luật lao động, an toàn lao động và các quy định của Công ty và pháp luật của Nhà nước;')
    p=doc.add_paragraph('-    Người lao động có trách nhiệm tuân thủ đầy đủ quy định về an toàn lao động, quy trình vận hành thiết bị và hướng dẫn của Công ty. Trường hợp NLĐ cố ý vi phạm hoặc vi phạm nghiêm trọng quy định an toàn lao động gây thiệt hại thì phải chịu trách nhiệm theo quy định pháp luật và nội quy Công ty;')
    p=doc.add_paragraph('-    Bồi thường vi phạm vật chất : Phải bồi thường vật chất do cá nhân vi phạm quy định của Công ty về bảo quản trang thiết bị được giao.')
    p=doc.add_paragraph(); r=p.add_run('2. Quyền Lợi:'); r.bold=True
    p=doc.add_paragraph('-    Phương tiện đi lại: Tự túc;')
    p=doc.add_paragraph('-    Được Công ty cấp đầy đủ bảo hộ lao động theo đúng vị trí làm việc;')
    p=doc.add_paragraph('-    Được phân công công việc theo yêu cầu của Công ty phù hợp với khả năng và trình độ chuyên môn mà người lao động đáp ứng;')
    p=doc.add_paragraph('-    Các quyền lợi khác thực hiện theo quy định của Pháp luật Lao động như tạm dừng, chấm dứt hợp đồng.')
    p=doc.add_paragraph(); r=p.add_run('Điều 4. Nghĩa vụ, quyền hạn NSDLĐ:'); r.bold=True
    p=doc.add_paragraph('-    Bảo đảm việc làm và thực hiện đầy đủ những điều đã cam kết trong hợp đồng;')
    p=doc.add_paragraph('-    Thanh toán đầy đủ, đúng hạn các chế độ và quyền lợi cho người lao động theo hợp đồng;')
    p=doc.add_paragraph('-    Điều hành người lao động hoàn thành công việc theo hợp đồng;')
    p=doc.add_paragraph('-    Tạm hoãn, chấm dứt hợp đồng theo quy định của pháp luật, và nội quy lao động của Công ty;')
    p=doc.add_paragraph(); r=p.add_run('Điều 5. Điều khoản chung:'); r.bold=True
    p=doc.add_paragraph('-    Những nội dung về quan hệ lao động không ghi trong hợp đồng này thì được áp dụng theo pháp luật lao động;')
    p=doc.add_paragraph('-    Những thoả thuận khác (nếu có): không;')
    p=doc.add_paragraph('-    Hợp đồng này có hiệu lực từ ngày ký và được làm thành 02 bản, Bên A giữ 01 bản, Bên B giữ 01 có giá trị pháp lý như nhau, để làm căn cứ thực hiện.'); doc.add_paragraph()
    p=doc.add_paragraph('Bản HĐ này lập tại văn phòng Công ty CP Cảng Hòn La.'); doc.add_paragraph()
    ts=doc.add_table(rows=3,cols=2); ts.alignment=WD_TABLE_ALIGNMENT.CENTER; remove_table_border(ts)
    c=ts.rows[0].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[0].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run('NGƯỜI SỬ DỤNG LAO ĐỘNG'); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[1].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[1].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].add_run('').font.size=Pt(12); sp=c.add_paragraph(); sp.paragraph_format.space_after=Pt(60)
    c=ts.rows[2].cells[0]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(nv.get('ho_ten','').upper()); r.bold=True; r.font.size=Pt(13)
    c=ts.rows[2].cells[1]; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=c.paragraphs[0].add_run(CC['dai_dien'].upper()); r.bold=True; r.font.size=Pt(13)
    tf = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    doc.save(tf.name)
    return tf.name

def gui_email(loai, ds, file=None):
    from config import EMAIL_CONFIG as EC
    try:
        msg = MIMEMultipart()
        msg['From'] = EC['email']
        msg['To'] = EC['nguoi_nhan']
        tn = f"{datetime.now().month:02d}/{datetime.now().year}"
        msg['Subject'] = f"[HRM-Port] Báo cáo {loai} lao động tháng {tn}"
        nd = f"<h3>BÁO CÁO {loai.upper()} LĐ</h3><p>Tháng: <b>{tn}</b></p><p>SL: <b>{len(ds)}</b></p><hr><ul>"
        for nv in ds[:10]:
            nd += f"<li>{nv.get('ho_ten','')} - {nv.get('chuc_danh_nghe','')}</li>"
        if len(ds) > 10:
            nd += f"<li>... và {len(ds)-10} người khác</li>"
        nd += "</ul><p><b>File Excel đính kèm.</b></p>"
        msg.attach(MIMEText(nd, 'html', 'utf-8'))
        if file and os.path.exists(file):
            with open(file, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file)}"')
                msg.attach(part)
        srv = smtplib.SMTP(EC['smtp_server'], EC['smtp_port'])
        srv.starttls()
        srv.login(EC['email'], EC['password'])
        srv.send_message(msg)
        srv.quit()
        return True
    except Exception as e:
        st.error(f"Lỗi email: {e}")
        return False

def gui_telegram(msg):
    from config import TELEGRAM_CONFIG as TC
    try:
        url=f"https://api.telegram.org/bot{TC['bot_token']}/sendMessage"
        r=requests.post(url,data={"chat_id":TC['chat_id'],"text":msg,"parse_mode":"HTML"},timeout=10)
        return r.status_code==200
    except: return False

# ========== SIDEBAR + LOGIN ==========
st.sidebar.title("🏗️ HRM-Port")
st.sidebar.caption("Quản lý nhân sự cảng biển")

# Hàm kiểm tra đăng nhập từ secrets
def check_login(username, password):
    from config import USERS
    
    # Kiểm tra từ file config.py trước
    if username in USERS:
        return USERS[username]['password'] == password, USERS[username]['role']
    
    # Fallback: kiểm tra từ st.secrets (nếu có)
    try:
        if 'users' in st.secrets and username in st.secrets.users:
            return st.secrets.users[username]['password'] == password, st.secrets.users[username]['role']
    except:
        pass
    
    return False, None

if 'logged_in' not in st.session_state: 
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None

if not st.session_state.logged_in:
    st.sidebar.subheader("🔐 Đăng nhập")
    u = st.sidebar.text_input("Tài khoản")
    p = st.sidebar.text_input("Mật khẩu", type="password")
    c1, c2 = st.sidebar.columns(2)
    with c1:
        if st.button("Đăng nhập", use_container_width=True):
            success, role = check_login(u, p)
            if success:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.username = u
                st.rerun()
            else:
                st.sidebar.error("❌ Sai tài khoản hoặc mật khẩu!")
    with c2:
        if st.button("👁️ Xem thử", use_container_width=True):
            st.session_state.logged_in = True
            st.session_state.role = "viewer"
            st.session_state.username = "guest"
            st.rerun()
    st.stop()

# Menu theo role
if st.session_state.role == "admin":
    menu_options = ["📊 Dashboard","👤 Ứng viên","✅ Nhân viên","📁 Upload hồ sơ","⚙️ Danh mục","📋 BHXH","📋 Báo cáo 01/PLI"]
else:  # viewer
    menu_options = ["📊 Dashboard","👤 Ứng viên","✅ Nhân viên","📋 BHXH","📋 Báo cáo 01/PLI"]
menu = st.sidebar.radio("📋 Menu", menu_options)
st.sidebar.divider()
st.sidebar.caption(f"👤 {st.session_state.username} ({st.session_state.role})")
if st.sidebar.button("🚪 Đăng xuất", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None
    st.rerun()

# ========== DASHBOARD ==========
if menu == "📊 Dashboard":
    st.title("📊 Dashboard")
    db = get_connection()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute("SELECT COUNT(*) t FROM ung_vien")
    tuv = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC')")
    tnv = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM ung_vien WHERE trang_thai='CHO_DUYET'")
    cd = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM ung_vien WHERE trang_thai='TU_CHOI'")
    tc = c.fetchone()['t']
    c.execute("SELECT COUNT(*) t FROM ung_vien WHERE trang_thai='DA_NHAN_VIEC'")
    dn = c.fetchone()['t']
    cl1, cl2, cl3, cl4, cl5 = st.columns(5)
    cl1.metric("Tổng UV", tuv)
    cl2.metric("Nhân viên", tnv)
    cl3.metric("Chờ duyệt", cd)
    cl4.metric("Đã nhận", dn)
    cl5.metric("Từ chối", tc)
    st.divider()
    st.subheader("📌 Thông báo")
    c.execute("SELECT ho_ten FROM nhan_vien WHERE DATE(ngay_vao_lam)=CURRENT_DATE")
    hn = c.fetchall()
    c.execute("SELECT ho_ten FROM nhan_vien WHERE DATE(ngay_vao_lam)=CURRENT_DATE - INTERVAL '1 day'")
    hq = c.fetchall()
    if hn:
        st.success(f"🟢 Hôm nay có thêm: **{', '.join([x['ho_ten'] for x in hn])}**")
    if hq:
        st.info(f"🔵 Hôm qua có thêm: **{', '.join([x['ho_ten'] for x in hq])}**")
    if st.session_state.role == "admin":
        c.execute("""SELECT STT, ma_nv, ho_ten, ngay_vao_lam, 
            EXTRACT(DAY FROM ((ngay_vao_lam + INTERVAL '30 days') - CURRENT_DATE))::INTEGER as ngay_con_lai
            FROM nhan_vien 
            WHERE trang_thai = 'THU_VIEC' 
            AND EXTRACT(DAY FROM ((ngay_vao_lam + INTERVAL '30 days') - CURRENT_DATE))::INTEGER <= 5
            AND EXTRACT(DAY FROM ((ngay_vao_lam + INTERVAL '30 days') - CURRENT_DATE))::INTEGER >= 0 
            AND EXTRACT(DAY FROM (CURRENT_DATE - ngay_vao_lam))::INTEGER >= 25 
            ORDER BY ngay_con_lai ASC""")
        tv_sap_het = c.fetchall()
        for x in tv_sap_het:
            if x['ngay_con_lai'] == 0:
                st.error(f"⚠️ **{x.get('ma_nv','')} {x['ho_ten']}** - HÔM NAY LÀ NGÀY CUỐI HỢP ĐỒNG THỬ VIỆC!")
            else:
                st.warning(f"⚠️ **{x.get('ma_nv','')} {x['ho_ten']}** còn **{x['ngay_con_lai']}** ngày sẽ kết thúc hợp đồng thử việc!")
    
    # Phần sinh nhật (bỏ qua)
    sn_list = []
    
    c.execute("SELECT chuc_danh_nghe, COUNT(*) t FROM nhan_vien WHERE trang_thai='DANG_LAM' GROUP BY chuc_danh_nghe ORDER BY t DESC")
    data = c.fetchall()
    db.close()
    if data:
        st.divider()
        st.subheader("📈 Phân bố")
        df = pd.DataFrame(data)
        df.columns = ['Chức vụ', 'SL']
        st.bar_chart(df.set_index('Chức vụ'))
    st.divider()
    if st.session_state.role == "admin":
        if st.button("💾 BACKUP DỮ LIỆU NGAY", use_container_width=True):
            from backup_nv import backup_nhan_vien
            backup_nhan_vien()
            st.success("✅ Đã backup! Kiểm tra thư mục D:\\HRM_Port\\backup")
            
# ========== ỨNG VIÊN ==========
elif menu == "👤 Ứng viên":
    st.title("👤 Ứng viên")
    su = st.text_input("🔍 Tìm kiếm", key="suv")
    
    db_f = get_connection()
    c_f = db_f.cursor()
    c_f.execute("SELECT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
    ds_vi_tri = [row[0] for row in c_f.fetchall()]
    c_f.execute("SELECT DISTINCT vi_tri_du_tuyen FROM ung_vien WHERE vi_tri_du_tuyen IS NOT NULL AND vi_tri_du_tuyen != '' ORDER BY vi_tri_du_tuyen")
    for row in c_f.fetchall():
        if row[0] not in ds_vi_tri:
            ds_vi_tri.append(row[0])
    db_f.close()
    
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        filter_vi_tri = st.selectbox("🔍 Lọc Vị trí dự tuyển:", ["Tất cả"] + ds_vi_tri)
    
    # Chỉ admin mới thấy nút thêm ứng viên
    if st.session_state.role == "admin":
        with st.expander("➕ THÊM ỨNG VIÊN MỚI", expanded=False):
            with st.form("add_uv_form"):
                db_f = get_connection()
                c_f = db_f.cursor()
                c_f.execute("SELECT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
                ds_vt_uv = [row[0] for row in c_f.fetchall()]
                db_f.close()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    ho_ten_uv = st.text_input("Họ và tên *")
                    vi_tri_uv = st.selectbox("Vị trí dự tuyển", [""] + ds_vt_uv)
                    dien_thoai_uv = st.text_input("SĐT")
                with col2:
                    ngay_sinh_uv = st.text_input("Ngày sinh (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    gioi_tinh_uv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"])
                with col3:
                    ngay_vao_lam_uv = st.text_input("Ngày vào làm (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    ghi_chu_uv = st.text_area("Ghi chú")
                
                if st.form_submit_button("💾 LƯU", use_container_width=True):
                    if ho_ten_uv:
                        ngay_loi = []
                        if ngay_sinh_uv and not parse_date(ngay_sinh_uv): 
                            ngay_loi.append("Ngày sinh")
                        if ngay_vao_lam_uv and not parse_date(ngay_vao_lam_uv): 
                            ngay_loi.append("Ngày vào làm")
                        if ngay_loi:
                            st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                        else:
                            try:
                                db = get_connection()
                                c = db.cursor()
                                c.execute("""INSERT INTO ung_vien (ho_ten, vi_tri_du_tuyen, dien_thoai, 
                                    ngay_sinh, gioi_tinh, ngay_vao_lam, luong_bao_hiem, trang_thai)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, 'CHO_DUYET')""",
                                    (ho_ten_uv, vi_tri_uv, dien_thoai_uv, parse_date(ngay_sinh_uv),
                                     gioi_tinh_uv, parse_date(ngay_vao_lam_uv), ghi_chu_uv))
                                new_id = c.lastrowid
                                ma_uv = f"UV{new_id:04d}"
                                c.execute("UPDATE ung_vien SET ma_uv = %s WHERE id = %s", (ma_uv, new_id))
                                db.commit()
                                db.close()
                                st.success(f"✅ Đã thêm ứng viên: {ho_ten_uv} (Mã: {ma_uv})")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi thêm ứng viên: {e}")
                    else:
                        st.error("Họ tên không được để trống!")
    
    st.divider()
    
    t1, t2, t3, t4 = st.tabs(["📋 Tất cả", "⏳ Chờ duyệt", "✅ Đã nhận", "❌ Từ chối"])
    tm = {"📋 Tất cả": "", "⏳ Chờ duyệt": "CHO_DUYET", "✅ Đã nhận": "DA_NHAN_VIEC", "❌ Từ chối": "TU_CHOI"}
    
    for tn, tab in zip(tm.keys(), [t1, t2, t3, t4]):
        with tab:
            tt = tm[tn]
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            sql = "SELECT id, ma_uv, ho_ten, vi_tri_du_tuyen, dien_thoai, ngay_vao_lam, luong_bao_hiem, ngay_sinh, trang_thai FROM ung_vien WHERE 1=1"
            params = []
            if tt:
                sql += " AND trang_thai = %s"
                params.append(tt)
            if su:
                sql += " AND (ho_ten LIKE %s OR dien_thoai LIKE %s OR ma_uv LIKE %s)"
                params.extend([f'%{su}%', f'%{su}%', f'%{su}%'])
            if filter_vi_tri != "Tất cả":
                sql += " AND vi_tri_du_tuyen = %s"
                params.append(filter_vi_tri)
            sql += " ORDER BY id ASC"
            c.execute(sql, tuple(params))
            ds = c.fetchall()
            db.close()
            
            if ds:
                df = pd.DataFrame(ds)
                for col in df.columns:
                    if 'ngay' in col.lower():
                        df[col] = df[col].apply(format_date)
                
                display_cols = ['ma_uv', 'ho_ten', 'vi_tri_du_tuyen', 'dien_thoai', 'ngay_vao_lam', 'luong_bao_hiem', 'ngay_sinh', 'trang_thai']
                available_cols = [c for c in display_cols if c in df.columns]
                df_show = df[available_cols]
                
                col_map = {
                    'ma_uv': 'Mã UV',
                    'ho_ten': 'Họ tên',
                    'vi_tri_du_tuyen': 'Vị trí dự tuyển',
                    'dien_thoai': 'SĐT',
                    'ngay_vao_lam': 'Ngày vào làm',
                    'luong_bao_hiem': 'Ghi chú',
                    'ngay_sinh': 'Ngày sinh',
                    'trang_thai': 'Trạng thái',
                }
                df_show.rename(columns=col_map, inplace=True)
                
                st.caption(f"📌 {len(ds)} kết quả.")
                
                if st.session_state.role == "admin":
                    # Admin: hiển thị bảng có checkbox và nút chức năng
                    if 'selected' not in df.columns:
                        df.insert(0, 'selected', False)
                    df_show_with_checkbox = df[['selected'] + [c for c in df.columns if c in display_cols]]
                    df_show_with_checkbox.rename(columns={'selected': 'Chọn'}, inplace=True)
                    
                    edited_df = st.data_editor(
                        df_show_with_checkbox,
                        column_config={"Chọn": st.column_config.CheckboxColumn("Chọn", default=False)},
                        disabled=[col for col in df_show_with_checkbox.columns if col != 'Chọn'],
                        hide_index=True,
                        height=400,
                        key=f"uv_editor_{tn}"
                    )
                    
                    if edited_df is not None:
                        selected_rows = edited_df[edited_df['Chọn'] == True]
                        if len(selected_rows) > 1:
                            st.error("⚠️ Chỉ được chọn 1 ứng viên!")
                        elif len(selected_rows) == 1:
                            selected_idx = selected_rows.index[0]
                            selected_nv = df.iloc[selected_idx]
                            col_btn1, col_btn2 = st.columns(2)
                            with col_btn1:
                                if st.button(f"✏️ SỬA", key=f"edit_sel_{tn}"):
                                    st.session_state['edit_uv_id'] = int(selected_nv['id'])
                                    st.rerun()
                            if tn == "⏳ Chờ duyệt":
                                with col_btn2:
                                    if st.button(f"✅ CHUYỂN SANG THỬ VIỆC", type="primary", key=f"chuyen_uv_{tn}"):
                                        try:
                                            db = get_connection()
                                            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                            uv_id = int(selected_nv['id'])
                                            c.execute("SELECT * FROM ung_vien WHERE id = %s", (uv_id,))
                                            uv = c.fetchone()
                                            if uv:
                                                c.execute("SELECT COALESCE(MAX(STT),0)+1 as next_stt FROM nhan_vien")
                                                result = c.fetchone()
                                                stt_moi = int(result['next_stt']) if result else 1
                                                ma_nv = f"NV{stt_moi:03d}"
                                                nhl = uv.get('ngay_vao_lam') or date.today()
                                                if isinstance(nhl, str):
                                                    nhl = datetime.strptime(nhl, '%Y-%m-%d').date()
                                                c.execute("SELECT COUNT(*) as tv_count FROM nhan_vien WHERE so_hdld LIKE '%/HĐTV-CHL'")
                                                tv_result = c.fetchone()
                                                tv_cnt = int(tv_result['tv_count']) + 1
                                                so_hdtv = f"{tv_cnt:02d}/{nhl.year}/HĐTV-CHL"
                                                c.execute("""INSERT INTO nhan_vien (STT, ma_nv, so_hdld, ho_ten, chuc_danh_nghe, dien_thoai,
                                                    ngay_sinh, gioi_tinh, ngay_vao_lam, noi_lam_viec, loai_hop_dong, trang_thai, trang_thai_bhxh, ngay_ky_hd)
                                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'Thử việc', 'THU_VIEC', 'CHUA_DONG', %s)""",
                                                    (stt_moi, ma_nv, so_hdtv, uv['ho_ten'], uv['vi_tri_du_tuyen'], uv['dien_thoai'],
                                                     uv['ngay_sinh'], uv['gioi_tinh'], nhl, 'Cảng THQT Hòn La', nhl))
                                                c.execute("UPDATE ung_vien SET trang_thai='DA_NHAN_VIEC', ma_nv=%s WHERE id=%s", (ma_nv, uv_id))
                                                db.commit()
                                                st.success(f"✅ Đã chuyển {uv['ho_ten']} → {ma_nv}")
                                                st.rerun()
                                            db.close()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi: {e}")
                else:
                    # Viewer: chỉ hiển thị bảng
                    st.dataframe(df_show, use_container_width=True, hide_index=True, height=400)
            else:
                st.info("Không có dữ liệu")
    
    # Form sửa ứng viên (chỉ admin)
    if 'edit_uv_id' in st.session_state and st.session_state.role == "admin":
        st.divider()
        st.subheader(f"✏️ Sửa ứng viên")
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT * FROM ung_vien WHERE id = %s", (st.session_state['edit_uv_id'],))
        uv_data = c.fetchone()
        db.close()
        if uv_data:
            with st.form("edit_uv_direct"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    ho_ten_e = st.text_input("Họ và tên *", value=uv_data['ho_ten'] or '')
                    vi_tri_e = st.selectbox("Vị trí dự tuyển", [""] + ds_vi_tri,
                        index=([""] + ds_vi_tri).index(uv_data['vi_tri_du_tuyen']) if uv_data['vi_tri_du_tuyen'] in ds_vi_tri else 0)
                    dien_thoai_e = st.text_input("SĐT", value=uv_data['dien_thoai'] or '')
                with col2:
                    ngay_sinh_e = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(uv_data['ngay_sinh']))
                    gioi_tinh_e = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"],
                        index=["", "Nam", "Nữ", "Khác"].index(uv_data['gioi_tinh']) if uv_data['gioi_tinh'] in ["Nam", "Nữ", "Khác"] else 0)
                with col3:
                    ngay_vao_lam_e = st.text_input("Ngày vào làm (dd/mm/yyyy)", value=format_date(uv_data['ngay_vao_lam']))
                    ghi_chu_e = st.text_area("Ghi chú", value=uv_data['luong_bao_hiem'] or '')
                    trang_thai_e = st.selectbox("Trạng thái", ["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"],
                        index=["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"].index(uv_data['trang_thai']) if uv_data['trang_thai'] in ["CHO_DUYET", "TU_CHOI", "DA_NHAN_VIEC"] else 0)
                
                col_save, col_del, col_cancel = st.columns(3)
                with col_save:
                    if st.form_submit_button("💾 CẬP NHẬT"):
                        ngay_loi = []
                        if ngay_sinh_e and not parse_date(ngay_sinh_e): 
                            ngay_loi.append("Ngày sinh")
                        if ngay_vao_lam_e and not parse_date(ngay_vao_lam_e): 
                            ngay_loi.append("Ngày vào làm")
                        if ngay_loi:
                            st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                        else:
                            db = get_connection()
                            c = db.cursor()
                            c.execute("""UPDATE ung_vien SET ho_ten=%s, vi_tri_du_tuyen=%s, dien_thoai=%s,
                                ngay_sinh=%s, gioi_tinh=%s, ngay_vao_lam=%s, luong_bao_hiem=%s, trang_thai=%s
                                WHERE id=%s""",
                                (ho_ten_e, vi_tri_e, dien_thoai_e, parse_date(ngay_sinh_e), gioi_tinh_e,
                                 parse_date(ngay_vao_lam_e), ghi_chu_e, trang_thai_e, uv_data['id']))
                            db.commit()
                            db.close()
                            st.success("✅ Đã cập nhật!")
                            del st.session_state['edit_uv_id']
                            st.rerun()
                with col_del:
                    if st.form_submit_button("🗑️ XÓA"):
                        db = get_connection()
                        c = db.cursor()
                        c.execute("DELETE FROM ung_vien WHERE id = %s", (uv_data['id'],))
                        db.commit()
                        db.close()
                        st.success("🗑️ Đã xóa!")
                        del st.session_state['edit_uv_id']
                        st.rerun()
                with col_cancel:
                    if st.form_submit_button("❌ HỦY"):
                        del st.session_state['edit_uv_id']
                        st.rerun()
    
    # Quản lý danh mục vị trí dự tuyển (chỉ admin)
    if st.session_state.role == "admin":
        st.divider()
        with st.expander("⚙️ Quản lý danh mục Vị trí dự tuyển", expanded=False):
            with st.form("add_vi_tri_uv"):
                ten_vt_moi = st.text_input("Tên vị trí dự tuyển mới *")
                if st.form_submit_button("➕ Thêm"):
                    if ten_vt_moi:
                        db = get_connection()
                        c = db.cursor()
                        c.execute("SELECT COUNT(*) FROM vi_tri_cong_tac WHERE ten_vi_tri = %s", (ten_vt_moi,))
                        if c.fetchone()[0] == 0:
                            c.execute("INSERT INTO vi_tri_cong_tac (ten_vi_tri) VALUES (%s)", (ten_vt_moi,))
                            db.commit()
                            st.success(f"✅ Đã thêm: {ten_vt_moi}")
                            st.rerun()
                        else:
                            st.warning("Vị trí này đã tồn tại!")
                        db.close()
                    else:
                        st.error("Tên không được để trống!")
            
            db = get_connection()
            c = db.cursor()
            c.execute("SELECT id, ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
            ds_vt = c.fetchall()
            db.close()
            if ds_vt:
                st.caption("📋 Danh sách vị trí dự tuyển:")
                for row in ds_vt:
                    st.write(f"- {row[1]}")

# ========== NHÂN VIÊN ==========
elif menu == "✅ Nhân viên":
    st.title("✅ Quản lý nhân viên")
    
    tab_dang_lam, tab_da_nghi, tab_qtct = st.tabs(["📌 ĐANG LÀM VIỆC", "📋 ĐÃ NGHỈ VIỆC", "📜 LỊCH SỬ CÔNG TÁC"])
    
    with tab_dang_lam:
        st.caption("👥 Danh sách nhân viên đang làm việc (bao gồm thử việc)")
        sn = st.text_input("🔍 Tìm kiếm", key="snv_dang_lam")
        
        if st.session_state.role == "admin":
            with st.expander("➕ THÊM NHÂN VIÊN MỚI", expanded=False):
                with st.form("add_nv"):
                    db = get_connection()
                    c = db.cursor()
                    c.execute("SELECT DISTINCT ten_vi_tri FROM vi_tri_cong_tac ORDER BY ten_vi_tri")
                    dcv = [row[0] for row in c.fetchall()]
                    db.close()
                    st.caption("📝 Thông tin cá nhân")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        htn = st.text_input("Họ và tên *")
                        nsn = st.text_input("Ngày sinh (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                        gtn = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"])
                        qtn = st.text_input("Quốc tịch", value="Việt Nam")
                        dtn = st.text_input("Dân tộc", value="Kinh")
                    with c2:
                        scc = st.text_input("CCCD")
                        ncc = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                        ncc2 = st.text_input("Nơi cấp CCCD")
                        nqn = st.text_input("Nguyên quán")
                        ttn = st.text_input("Thường trú")
                    with c3:
                        dtn2 = st.text_input("SĐT")
                        emn = st.text_input("Email")
                        cdn = st.selectbox("Chức danh", [""] + dcv)
                        pbn = st.text_input("Phòng ban")
                        nlv = st.text_input("Nơi làm việc", value="Cảng THQT Hòn La")
                    st.divider()
                    st.caption("💼 Hợp đồng & BHXH")
                    c4, c5, c6 = st.columns(3)
                    with c4:
                        lhd = st.selectbox("Loại HĐ *", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"])
                        nvl = st.text_input("Ngày vào làm (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                        nkt = st.text_input("Ngày kết thúc", placeholder="dd/mm/yyyy", max_chars=10)
                        mbh = st.text_input("Mã BHXH")
                        tbd = st.text_input("Bắt đầu BH (dd/mm/yyyy)", placeholder="dd/mm/yyyy", max_chars=10)
                    with c5:
                        lbh = st.text_input("Lương BH")
                        hsl = st.text_input("Hệ số lương")
                        pcv = st.text_input("PC chức vụ")
                        ptv = st.text_input("PC TNVK (%)")
                        ptn = st.text_input("PC TNN (%)")
                    with c6:
                        mhb = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"])
                        tld = st.text_input("Tỷ lệ đóng (%)")
                        mtd = st.text_input("Mức tiền đóng")
                        ptd = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"])
                        nbh = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"])
                    st.divider()
                    st.caption("🏦 Ngân hàng & KCB")
                    c7, c8, c9 = st.columns(3)
                    with c7:
                        stk = st.text_input("STK")
                        cnh = st.text_input("Chi nhánh NH")
                        tkb = st.text_input("Tỉnh KCB")
                        nkb = st.text_input("Nơi KCB")
                    with c8:
                        ths = st.text_input("Tỉnh/TP nhận HS")
                        phs = st.text_input("Phường/Xã nhận HS")
                        dhs = st.text_area("Địa chỉ nhận HS", height=100)
                    with c9:
                        dks = st.selectbox("ĐK nhận sổ", ["Có", "Không"])
                        hso = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"])
                    
                    if st.form_submit_button("💾 LƯU"):
                        if htn:
                            ngay_loi = []
                            if nsn and not parse_date(nsn):
                                ngay_loi.append("Ngày sinh")
                            if ncc and not parse_date(ncc):
                                ngay_loi.append("Ngày cấp CCCD")
                            if nvl and not parse_date(nvl):
                                ngay_loi.append("Ngày vào làm")
                            if nkt and not parse_date(nkt):
                                ngay_loi.append("Ngày kết thúc")
                            if tbd and not parse_date(tbd):
                                ngay_loi.append("Bắt đầu BH")
                            if ngay_loi:
                                st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                            else:
                                try:
                                    db = get_connection()
                                    c = db.cursor()
                                    c.execute("SELECT COALESCE(MAX(STT),0)+1 FROM nhan_vien")
                                    stt_moi = c.fetchone()[0]
                                    ma_nv = f"NV{stt_moi:03d}"
                                    nhl = parse_date(nvl) or date.today()
                                    if lhd == "Thử việc":
                                        ttnv, ttbh, tbd_val = 'THU_VIEC', 'CHUA_DONG', None
                                        c.execute("SELECT COUNT(*) FROM nhan_vien WHERE so_hdld LIKE '%/HĐTV-CHL'")
                                        tv_cnt = c.fetchone()[0] + 1
                                        so_hd = f"{tv_cnt:02d}/{nhl.year}/HĐTV-CHL"
                                    else:
                                        ttnv, ttbh = 'DANG_LAM', 'DANG_DONG'
                                        tbd_val = parse_date(tbd) or parse_date(nvl)
                                        c.execute("SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)),0)+1 FROM nhan_vien WHERE so_hdld LIKE '%/HĐLĐ-CHL'")
                                        so_hd = f"{int(c.fetchone()[0]):02d}/{nhl.year}/HĐLĐ-CHL"
                                    c.execute("""INSERT INTO nhan_vien (STT, ma_nv, so_hdld, ho_ten, chuc_danh_nghe, ngay_sinh, gioi_tinh,
                                        so_cccd, ngay_cap_cccd, noi_cap_cccd, nguyen_quan, thuong_tru,
                                        dien_thoai, email, email_lien_he, ho_so, luong_bao_hiem, ma_so_bhxh, ngay_vao_lam,
                                        noi_lam_viec, so_tai_khoan_nh, chi_nhanh_nh, ngay_ky_hd, loai_hop_dong,
                                        nhom_bhxh, thang_bat_dau_bh, thang_ket_thuc_bh, trang_thai, trang_thai_bhxh,
                                        phong_ban_lam_viec, ngay_ket_thuc, quoc_tich, dan_toc, he_so_luong, phu_cap_chuc_vu,
                                        phu_cap_tnvk, phu_cap_tnn, muc_huong_bhyt, ty_le_dong, muc_tien_dong, phuong_thuc_dong,
                                        tinh_nhan_hs, phuong_nhan_hs, dia_chi_nhan_hs, tinh_kcb, noi_dang_ky_kcb, dang_ky_nhan_so)
                                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s, %s)""",
                                          (stt_moi, ma_nv, so_hd, htn, cdn, parse_date(nsn), gtn, scc, parse_date(ncc), ncc2, nqn, ttn,
                                           dtn2, emn, emn, '', lbh, mbh, parse_date(nvl), nlv, stk, cnh, parse_date(nvl), lhd,
                                           nbh, tbd_val, None, ttnv, ttbh, pbn, parse_date(nkt), qtn, dtn, hsl, pcv, ptv, ptn,
                                           mhb, tld, mtd, ptd, ths, phs, dhs, tkb, nkb, dks))
                                    db.commit()
                                    db.close()
                                    st.success(f"✅ Đã lưu nhân viên mới thành công! {htn} - {ma_nv}")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Lỗi: {e}")
                        else:
                            st.error("Họ tên không được để trống!")
                st.divider()
        
        db_f = get_connection()
        c_f = db_f.cursor()
        c_f.execute("SELECT DISTINCT chuc_danh_nghe FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND chuc_danh_nghe IS NOT NULL AND chuc_danh_nghe != '' ORDER BY chuc_danh_nghe")
        ds_chuc_danh = [row[0] for row in c_f.fetchall()]
        c_f.execute("SELECT DISTINCT loai_hop_dong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') AND loai_hop_dong IS NOT NULL AND loai_hop_dong != '' ORDER BY loai_hop_dong")
        ds_loai_hd = [row[0] for row in c_f.fetchall()]
        db_f.close()
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filter_chuc_danh = st.selectbox("🔍 Lọc Chức danh:", ["Tất cả"] + ds_chuc_danh, key="filter_cd_danglam")
        with col_f2:
            filter_loai_hd = st.selectbox("🔍 Lọc Loại HĐ:", ["Tất cả"] + ds_loai_hd, key="filter_lhd_danglam")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        sql = "SELECT * FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC')"
        params = []
        if sn:
            sql += " AND (ho_ten LIKE %s OR dien_thoai LIKE %s OR so_cccd LIKE %s OR ma_nv LIKE %s)"
            params.extend([f'%{sn}%'] * 4)
        if filter_chuc_danh != "Tất cả":
            sql += " AND chuc_danh_nghe = %s"
            params.append(filter_chuc_danh)
        if filter_loai_hd != "Tất cả":
            sql += " AND loai_hop_dong = %s"
            params.append(filter_loai_hd)
        sql += " ORDER BY id DESC"
        c.execute(sql, tuple(params))
        ds = c.fetchall()
        db.close()
        
        if ds:
            df = pd.DataFrame(ds)
            for col in df.columns:
                if 'ngay' in col.lower():
                    df[col] = df[col].apply(format_date)
            
            if 'selected' not in df.columns:
                df.insert(0, 'selected', False)
            
            display_cols = ['selected', 'ma_nv', 'ho_ten', 'ngay_sinh', 'gioi_tinh', 'so_hdld', 'so_cccd', 'dien_thoai',
                            'thuong_tru', 'chuc_danh_nghe', 'loai_hop_dong', 'ngay_vao_lam', 'ma_so_bhxh', 'thang_bat_dau_bh']
            available_cols = [c for c in display_cols if c in df.columns]
            df_show = df[available_cols]
            
            col_map = {
                'selected': 'Chọn',
                'ma_nv': 'Mã NV',
                'ho_ten': 'Họ và tên',
                'ngay_sinh': 'Ngày sinh',
                'gioi_tinh': 'Giới tính',
                'so_hdld': 'Số HĐLĐ',
                'so_cccd': 'CCCD',
                'dien_thoai': 'SĐT',
                'thuong_tru': 'Thường trú',
                'chuc_danh_nghe': 'Chức danh',
                'loai_hop_dong': 'Loại HĐ',
                'ngay_vao_lam': 'Ngày vào làm',
                'ma_so_bhxh': 'Mã số BHXH',
                'thang_bat_dau_bh': 'Bắt đầu BH',
            }
            df_show.rename(columns=col_map, inplace=True)
            
            st.caption(f"📌 {len(ds)} kết quả. Tick chọn 1 nhân viên để thao tác.")
            
            # Nếu là viewer, hiển thị bảng không có checkbox chọn
            if st.session_state.role == "admin":
                edited_df = st.data_editor(
                    df_show,
                    column_config={
                        "Chọn": st.column_config.CheckboxColumn("Chọn để sửa", default=False)
                    },
                    disabled=[col for col in df_show.columns if col != 'Chọn'],
                    hide_index=True,
                    height=400,
                    key="nv_editor_danglam"
                )
            else:
                # Viewer: hiển thị bảng đơn thuần, không có checkbox
                st.dataframe(df_show.drop(columns=['Chọn'], errors='ignore'), use_container_width=True, hide_index=True, height=400)
                edited_df = None
            
            selected_nv = None
            if edited_df is not None and st.session_state.role == "admin" and 'Chọn' in edited_df.columns:
                selected_rows = edited_df[edited_df['Chọn'] == True]
                if len(selected_rows) > 0:
                    if len(selected_rows) > 1:
                        st.error("⚠️ Chỉ được chọn 1 nhân viên!")
                    else:
                        selected_idx = selected_rows.index[0]
                        selected_nv = df.iloc[selected_idx]
                        nv_id_key = selected_nv['id']
                        
                        # Hiển thị các nút chức năng (chỉ admin mới thấy và mới click được)
                        col_btn1, col_btn2, col_btn3, col_btn4, col_btn5 = st.columns(5)
                        
                        with col_btn1:
                            if st.button(f"✏️ SỬA '{selected_nv['ho_ten']}'", key=f"edit_nv_btn_{nv_id_key}", use_container_width=True):
                                st.session_state['selected_nv_id'] = int(selected_nv['id'])
                                st.rerun()
                        
                        with col_btn2:
                            trang_thai_nv = selected_nv.get('trang_thai', '')
                            if trang_thai_nv == 'DANG_LAM':
                                if st.button(f"🖨️ IN HĐLĐ - {selected_nv['ho_ten']}", key=f"print_hdld_btn_{nv_id_key}", use_container_width=True):
                                    db = get_connection()
                                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                    c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                    nv_full = c.fetchone()
                                    db.close()
                                    if nv_full:
                                        fp = tao_hop_dong(nv_full)
                                        with open(fp, "rb") as f:
                                            st.download_button(
                                                label="📥 TẢI HĐLĐ",
                                                data=f,
                                                file_name=f"HDLD_{nv_full['ho_ten']}_{datetime.now().strftime('%Y%m%d')}.docx",
                                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                                key=f"download_hdld_{nv_id_key}"
                                            )
                                    else:
                                        st.error("Không tìm thấy thông tin nhân viên!")
                            elif trang_thai_nv == 'THU_VIEC':
                                if st.button(f"🖨️ IN HĐTV - {selected_nv['ho_ten']}", key=f"print_hdtv_btn_{nv_id_key}", use_container_width=True):
                                    db = get_connection()
                                    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                    c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                    nv_full = c.fetchone()
                                    db.close()
                                    if nv_full:
                                        fp = tao_hop_dong_thu_viec(nv_full)
                                        with open(fp, "rb") as f:
                                            st.download_button(
                                                label="📥 TẢI HĐTV",
                                                data=f,
                                                file_name=f"HDTV_{nv_full['ho_ten']}_{datetime.now().strftime('%Y%m%d')}.docx",
                                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                                key=f"download_hdtv_{nv_id_key}"
                                            )
                                    else:
                                        st.error("Không tìm thấy thông tin nhân viên!")
                            else:
                                st.button(f"📄 {selected_nv['ho_ten']} (Không thể in HĐ)", disabled=True, use_container_width=True, key=f"disabled_btn_{nv_id_key}")
                        
                        with col_btn3:
                            if st.button(f"📱 GỬI ZALO - {selected_nv['ho_ten']}", key=f"zalo_btn_{nv_id_key}", use_container_width=True):
                                db = get_connection()
                                c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                c.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                nv_full = c.fetchone()
                                db.close()
                                if nv_full:
                                    ph = nv_full.get('dien_thoai', '')
                                    if ph:
                                        ph = ph.replace('+84', '0').replace(' ', '').strip()
                                        st.code(tao_noi_dung_zalo(nv_full))
                                        st.markdown(f"[👉 MỞ ZALO](https://zalo.me/{ph})")
                                    else:
                                        st.error("Chưa có SĐT!")
                                else:
                                    st.error("Không tìm thấy thông tin nhân viên!")
                        
                        with col_btn4:
                            ma_bhxh = selected_nv.get('ma_so_bhxh', '')
                            chua_co_bhxh = not bool(ma_bhxh and str(ma_bhxh).strip())
                            if chua_co_bhxh:
                                if st.button(f"🏠 NHẬP THÔNG TIN HỘ GIA ĐÌNH - {selected_nv['ho_ten']}", key=f"bhxh_family_btn_{nv_id_key}", use_container_width=True, type="primary"):
                                    st.session_state['bhxh_family_nv_id'] = int(selected_nv['id'])
                                    st.session_state['bhxh_family_nv_name'] = selected_nv['ho_ten']
                                    st.rerun()
                            else:
                                st.button(f"✅ ĐÃ CÓ BHXH - {selected_nv['ho_ten']}", disabled=True, use_container_width=True, key=f"has_bhxh_btn_{nv_id_key}")
                        
                        with col_btn5:
                            trang_thai_nv = selected_nv.get('trang_thai', '')
                            if trang_thai_nv == 'THU_VIEC':
                                if f'convert_open_{nv_id_key}' not in st.session_state:
                                    st.session_state[f'convert_open_{nv_id_key}'] = False
                                
                                if not st.session_state[f'convert_open_{nv_id_key}']:
                                    if st.button(f"🔄 CHUYỂN HĐLĐ KHÔNG XĐTH - {selected_nv['ho_ten']}", 
                                                key=f"convert_hdld_btn_{nv_id_key}", 
                                                use_container_width=True, type="primary"):
                                        st.session_state[f'convert_open_{nv_id_key}'] = True
                                        st.rerun()
                                else:
                                    st.markdown("---")
                                    st.markdown("### 📝 CHUYỂN ĐỔI HỢP ĐỒNG LAO ĐỘNG")
                                    st.caption("Vui lòng nhập đầy đủ thông tin cho quyết định chuyển đổi")
                                    
                                    db_temp = get_connection()
                                    c_temp = db_temp.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                                    c_temp.execute("SELECT * FROM nhan_vien WHERE id = %s", (int(selected_nv['id']),))
                                    nv_data = c_temp.fetchone()
                                    db_temp.close()
                                    
                                    if nv_data:
                                        ngay_quyet_dinh = st.date_input(
                                            "📅 Ngày quyết định:", 
                                            value=date.today(),
                                            key=f"ngay_qd_{nv_id_key}"
                                        )
                                        
                                        ngay_hieu_luc = st.date_input(
                                            "📅 Ngày hiệu lực (bắt đầu HĐLĐ):", 
                                            value=ngay_quyet_dinh,
                                            key=f"ngay_hl_{nv_id_key}"
                                        )
                                        
                                        current_year = datetime.now().year
                                        
                                        db_temp2 = get_connection()
                                        c_temp2 = db_temp2.cursor()
                                        c_temp2.execute("""
                                            SELECT COALESCE(MAX(CAST(SPLIT_PART(so_hdld, '/', 1) AS INTEGER)), 0) as max_stt
                                            FROM nhan_vien 
                                            WHERE so_hdld LIKE %s AND trang_thai IN ('DANG_LAM', 'THU_VIEC')
                                        """, (f'%/{current_year}/HĐLĐ-CHL',))
                                        result = c_temp2.fetchone()
                                        max_stt = result[0] if result else 0
                                        db_temp2.close()
                                        
                                        next_stt = max_stt + 1
                                        stt_str = str(next_stt).zfill(2)
                                        so_hd_moi = f"{stt_str}/{current_year}/HĐLĐ-CHL"
                                        
                                        st.info(f"📄 **Số HĐLĐ mới:** {so_hd_moi} (tự động sinh)")
                                        
                                        ngay_bat_dau_bh = st.date_input(
                                            "📅 Bắt đầu đóng BHXH:", 
                                            value=ngay_hieu_luc,
                                            key=f"ngay_bhxh_{nv_id_key}"
                                        )
                                        st.caption("⚠️ Đây là ngày bắt đầu tính đóng BHXH")
                                        
                                        ly_do_chuyen = st.text_area(
                                            "📝 Lý do/ Nội dung quyết định:", 
                                            value="Hoàn thành thời gian thử việc, chuyển sang hợp đồng lao động không xác định thời hạn",
                                            key=f"ly_do_{nv_id_key}",
                                            height=80
                                        )
                                        
                                        col_confirm1, col_confirm2, col_confirm3 = st.columns([1, 2, 1])
                                        with col_confirm2:
                                            if st.button("✅ XÁC NHẬN CHUYỂN ĐỔI", key=f"confirm_convert_{nv_id_key}", use_container_width=True, type="primary"):
                                                try:
                                                    db = get_connection()
                                                    c = db.cursor()
                                                    
                                                    c.execute("""
                                                        UPDATE nhan_vien SET 
                                                            trang_thai = 'DANG_LAM',
                                                            loai_hop_dong = 'Không xác định thời hạn',
                                                            so_hdld = %s,
                                                            ngay_ky_hd = %s,
                                                            ngay_chinh_thuc = %s,
                                                            thang_bat_dau_bh = %s,
                                                            trang_thai_bhxh = 'DANG_DONG',
                                                            ngay_ket_thuc = NULL
                                                        WHERE id = %s
                                                    """, (so_hd_moi, ngay_quyet_dinh, ngay_hieu_luc, ngay_bat_dau_bh, int(selected_nv['id'])))
                                                    
                                                    c.execute("""
                                                        INSERT INTO quyet_dinh_nhan_su (
                                                            nhan_vien_id, loai_quyet_dinh, ngay_quyet_dinh, ngay_hieu_luc,
                                                            noi_dung, so_quyet_dinh, loai_hop_dong_cu, loai_hop_dong_moi,
                                                            he_so_luong_cu, he_so_luong_moi
                                                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                                    """, (
                                                        int(selected_nv['id']),
                                                        'CHINH_THUC',
                                                        ngay_quyet_dinh,
                                                        ngay_hieu_luc,
                                                        ly_do_chuyen,
                                                        f"QD{ngay_quyet_dinh.strftime('%Y%m%d')}_{selected_nv['ma_nv']}",
                                                        nv_data.get('loai_hop_dong', 'Thử việc'),
                                                        'Không xác định thời hạn',
                                                        nv_data.get('he_so_luong', 0),
                                                        nv_data.get('he_so_luong', 0)
                                                    ))
                                                    
                                                    c.execute("""
                                                        UPDATE lich_su_cong_tac 
                                                        SET den_ngay = %s 
                                                        WHERE nhan_vien_id = %s AND den_ngay IS NULL
                                                    """, (ngay_hieu_luc - timedelta(days=1), int(selected_nv['id'])))
                                                    
                                                    c.execute("""
                                                        INSERT INTO lich_su_cong_tac (
                                                            nhan_vien_id, tu_ngay, chuc_danh, phong_ban, 
                                                            noi_lam_viec, loai_hop_dong, he_so_luong
                                                        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
                                                    """, (
                                                        int(selected_nv['id']),
                                                        ngay_hieu_luc,
                                                        nv_data.get('chuc_danh_nghe', ''),
                                                        nv_data.get('phong_ban_lam_viec', ''),
                                                        nv_data.get('noi_lam_viec', 'Cảng THQT Hòn La'),
                                                        'Không xác định thời hạn',
                                                        nv_data.get('he_so_luong', 0)
                                                    ))
                                                    
                                                    db.commit()
                                                    db.close()
                                                    
                                                    st.success(f"✅ Đã chuyển {nv_data['ho_ten']} sang HĐLĐ không xác định thời hạn!")
                                                    st.info(f"📄 Số HĐLĐ mới: {so_hd_moi}")
                                                    st.info(f"📅 Ngày hiệu lực: {ngay_hieu_luc.strftime('%d/%m/%Y')}")
                                                    st.info(f"💰 Bắt đầu đóng BHXH từ: {ngay_bat_dau_bh.strftime('%d/%m/%Y')}")
                                                    
                                                    st.session_state[f'convert_open_{nv_id_key}'] = False
                                                    st.rerun()
                                                except Exception as e:
                                                    st.error(f"❌ Lỗi: {str(e)}")
                                        
                                        if st.button("❌ HỦY", key=f"cancel_convert_{nv_id_key}", use_container_width=True):
                                            st.session_state[f'convert_open_{nv_id_key}'] = False
                                            st.rerun()
                            else:
                                st.button(f"✅ ĐÃ LÀ HĐLĐ", disabled=True, use_container_width=True, key=f"already_hdld_btn_{nv_id_key}")
                        
                        st.divider()
            
            # Form sửa nhân viên (chỉ admin)
            if 'selected_nv_id' in st.session_state and st.session_state.role == "admin":
                nid = int(st.session_state['selected_nv_id'])
                db = get_connection()
                c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                c.execute("SELECT * FROM nhan_vien WHERE id=%s", (nid,))
                nd = c.fetchone()
                db.close()
                
                if nd:
                    st.subheader(f"✏️ Cập nhật: {nd.get('ho_ten', '')} ({nd.get('ma_nv', '')})")
                    with st.form("edit_nv"):
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            hnv = st.text_input("Họ tên *", value=nd.get('ho_ten', ''))
                            nsnv = st.text_input("Ngày sinh (dd/mm/yyyy)", value=format_date(nd.get('ngay_sinh')), placeholder="dd/mm/yyyy", max_chars=10)
                            gtnv = st.selectbox("Giới tính", ["", "Nam", "Nữ", "Khác"], index=["", "Nam", "Nữ", "Khác"].index(nd.get('gioi_tinh', '')) if nd.get('gioi_tinh') in ["Nam", "Nữ", "Khác"] else 0)
                            qtnv = st.text_input("Quốc tịch", value=nd.get('quoc_tich', 'Việt Nam'))
                            dtnv = st.text_input("Dân tộc", value=nd.get('dan_toc', 'Kinh'))
                        with col2:
                            sccv = st.text_input("CCCD", value=nd.get('so_cccd', ''))
                            nccv = st.text_input("Ngày cấp CCCD (dd/mm/yyyy)", value=format_date(nd.get('ngay_cap_cccd')), placeholder="dd/mm/yyyy", max_chars=10)
                            ncv = st.text_input("Nơi cấp CCCD", value=nd.get('noi_cap_cccd', ''))
                            nqnv = st.text_input("Nguyên quán", value=nd.get('nguyen_quan', ''))
                            ttnv = st.text_input("Thường trú", value=nd.get('thuong_tru', ''))
                        with col3:
                            dtnv2 = st.text_input("SĐT", value=nd.get('dien_thoai', ''))
                            emnv = st.text_input("Email", value=nd.get('email_lien_he', ''))
                            cdnv = st.text_input("Chức danh", value=nd.get('chuc_danh_nghe', ''))
                            pbnv = st.text_input("Phòng ban", value=nd.get('phong_ban_lam_viec', ''))
                            nlv2 = st.text_input("Nơi làm việc", value=nd.get('noi_lam_viec', 'Cảng THQT Hòn La'))
                        
                        st.divider()
                        st.caption("💼 Hợp đồng & BHXH")
                        col4, col5, col6 = st.columns(3)
                        with col4:
                            lhdv = st.selectbox("Loại HĐ", ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"], index=["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"].index(nd.get('loai_hop_dong', 'Thử việc')) if nd.get('loai_hop_dong') in ["Thử việc", "Xác định thời hạn", "Không xác định thời hạn"] else 0)
                            nvlv = st.text_input("Ngày vào làm (dd/mm/yyyy)", value=format_date(nd.get('ngay_vao_lam')), placeholder="dd/mm/yyyy", max_chars=10)
                            nktv = st.text_input("Ngày kết thúc (dd/mm/yyyy)", value=format_date(nd.get('ngay_ket_thuc')), placeholder="dd/mm/yyyy", max_chars=10)
                            mbhv = st.text_input("Mã BHXH", value=nd.get('ma_so_bhxh', ''))
                            tbdv = st.text_input("Bắt đầu BH (dd/mm/yyyy)", value=format_date(nd.get('thang_bat_dau_bh')), placeholder="dd/mm/yyyy", max_chars=10)
                        with col5:
                            lbhv = st.text_input("Lương BH", value=nd.get('luong_bao_hiem', ''))
                            hslv = st.text_input("Hệ số lương", value=str(nd.get('he_so_luong', '')))
                            pcvv = st.text_input("PC chức vụ", value=str(nd.get('phu_cap_chuc_vu', '')))
                            ptvv = st.text_input("PC TNVK (%)", value=str(nd.get('phu_cap_tnvk', '')))
                            ptnv = st.text_input("PC TNN (%)", value=str(nd.get('phu_cap_tnn', '')))
                        with col6:
                            mhbv = st.selectbox("Mức hưởng BHYT", ["80%", "95%", "100%"], index=["80%", "95%", "100%"].index(nd.get('muc_huong_bhyt', '80%')) if nd.get('muc_huong_bhyt') in ["80%", "95%", "100%"] else 0)
                            tldv = st.text_input("Tỷ lệ đóng (%)", value=str(nd.get('ty_le_dong', '')))
                            mtdv = st.text_input("Mức tiền đóng", value=str(nd.get('muc_tien_dong', '')))
                            ptdv = st.selectbox("PT đóng", ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"], index=["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"].index(nd.get('phuong_thuc_dong', 'Hàng tháng')) if nd.get('phuong_thuc_dong') in ["Hàng tháng", "3 tháng", "6 tháng", "12 tháng"] else 0)
                            nbhv = st.selectbox("Nhóm BHXH", ["", "Văn phòng", "Lao động trực tiếp"], index=["", "Văn phòng", "Lao động trực tiếp"].index(nd.get('nhom_bhxh', '')) if nd.get('nhom_bhxh') in ["Văn phòng", "Lao động trực tiếp"] else 0)
                        
                        st.divider()
                        st.caption("🏦 Ngân hàng & KCB")
                        col7, col8, col9 = st.columns(3)
                        with col7:
                            stkv = st.text_input("STK", value=nd.get('so_tai_khoan_nh', ''))
                            cnhv = st.text_input("Chi nhánh NH", value=nd.get('chi_nhanh_nh', ''))
                            tkbv = st.text_input("Tỉnh KCB", value=nd.get('tinh_kcb', ''))
                            nkbv = st.text_input("Nơi KCB", value=nd.get('noi_dang_ky_kcb', ''))
                        with col8:
                            thsv = st.text_input("Tỉnh/TP nhận HS", value=nd.get('tinh_nhan_hs', ''))
                            phsv = st.text_input("Phường/Xã nhận HS", value=nd.get('phuong_nhan_hs', ''))
                            dhsv = st.text_area("Địa chỉ nhận HS", value=nd.get('dia_chi_nhan_hs', ''), height=100)
                        with col9:
                            dksv = st.selectbox("ĐK nhận sổ", ["Có", "Không"], index=["Có", "Không"].index(nd.get('dang_ky_nhan_so', 'Có')) if nd.get('dang_ky_nhan_so') in ["Có", "Không"] else 0)
                            hsov = st.selectbox("Hồ sơ", ["", "Đã có HS", "Chưa có"], index=["", "Đã có HS", "Chưa có"].index(nd.get('ho_so', '')) if nd.get('ho_so') in ["Đã có HS", "Chưa có"] else 0)
                        
                        col_save, col_quit, col_delete = st.columns(3)
                        with col_save:
                            if st.form_submit_button("💾 CẬP NHẬT"):
                                if hnv:
                                    ngay_loi = []
                                    if nsnv and not parse_date(nsnv):
                                        ngay_loi.append("Ngày sinh")
                                    if nccv and not parse_date(nccv):
                                        ngay_loi.append("Ngày cấp CCCD")
                                    if nvlv and not parse_date(nvlv):
                                        ngay_loi.append("Ngày vào làm")
                                    if nktv and not parse_date(nktv):
                                        ngay_loi.append("Ngày kết thúc")
                                    if tbdv and not parse_date(tbdv):
                                        ngay_loi.append("Bắt đầu BH")
                                    if ngay_loi:
                                        st.error(f"Sai định dạng dd/mm/yyyy: {', '.join(ngay_loi)}")
                                    else:
                                        try:
                                            db = get_connection()
                                            c = db.cursor()
                                            nhl = parse_date(nvlv) or date.today()
                                            ngay_bat_dau_bh = parse_date(tbdv) if tbdv and tbdv.strip() else None
                                            if lhdv == "Thử việc":
                                                tt_nv, tt_bh, tbd_val = 'THU_VIEC', 'CHUA_DONG', None
                                            else:
                                                tt_nv, tt_bh = 'DANG_LAM', 'DANG_DONG'
                                                if ngay_bat_dau_bh:
                                                    tbd_val = ngay_bat_dau_bh
                                                else:
                                                    tbd_val = parse_date(nvlv)
                                            c.execute("""UPDATE nhan_vien SET ho_ten=%s,chuc_danh_nghe=%s,ngay_sinh=%s,gioi_tinh=%s,
                                                so_cccd=%s,ngay_cap_cccd=%s,noi_cap_cccd=%s,nguyen_quan=%s,thuong_tru=%s,dien_thoai=%s,
                                                email=%s,email_lien_he=%s,ho_so=%s,luong_bao_hiem=%s,ma_so_bhxh=%s,ngay_vao_lam=%s,noi_lam_viec=%s,
                                                so_tai_khoan_nh=%s,chi_nhanh_nh=%s,ngay_ky_hd=%s,loai_hop_dong=%s,nhom_bhxh=%s,
                                                thang_bat_dau_bh=%s,trang_thai=%s,trang_thai_bhxh=%s,phong_ban_lam_viec=%s,
                                                ngay_ket_thuc=%s,quoc_tich=%s,dan_toc=%s,he_so_luong=%s,phu_cap_chuc_vu=%s,
                                                phu_cap_tnvk=%s,phu_cap_tnn=%s,muc_huong_bhyt=%s,ty_le_dong=%s,muc_tien_dong=%s,
                                                phuong_thuc_dong=%s,tinh_nhan_hs=%s,phuong_nhan_hs=%s,dia_chi_nhan_hs=%s,
                                                tinh_kcb=%s,noi_dang_ky_kcb=%s,dang_ky_nhan_so=%s WHERE id=%s""",
                                                  (hnv, cdnv, parse_date(nsnv), gtnv, sccv, parse_date(nccv), ncv, nqnv, ttnv, dtnv2,
                                                   emnv, emnv, hsov, lbhv, mbhv, parse_date(nvlv), nlv2, stkv, cnhv, parse_date(nvlv), lhdv,
                                                   nbhv, tbd_val, tt_nv, tt_bh, pbnv, parse_date(nktv), qtnv, dtnv, hslv, pcvv, ptvv, ptnv,
                                                   mhbv, tldv, mtdv, ptdv, thsv, phsv, dhsv, tkbv, nkbv, dksv, nid))
                                            db.commit()
                                            db.close()
                                            st.success(f"✅ Đã cập nhật: {hnv}")
                                            del st.session_state['selected_nv_id']
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi: {e}")
                                else:
                                    st.error("Họ tên không được để trống!")                        
                        with col_quit:
                            if f'nghi_viec_open_{nid}' not in st.session_state:
                                st.session_state[f'nghi_viec_open_{nid}'] = False
                            
                            if not st.session_state[f'nghi_viec_open_{nid}']:
                                if st.form_submit_button("🚫 NGHỈ VIỆC", use_container_width=True, type="secondary"):
                                    st.session_state[f'nghi_viec_open_{nid}'] = True
                                    st.rerun()
                            else:
                                st.markdown("---")
                                st.markdown("### 📝 XÁC NHẬN NGHỈ VIỆC")
                                
                                default_ngay_nghi = date.today()
                                ngay_ket_thuc_hien_tai = nd.get('ngay_ket_thuc')
                                if ngay_ket_thuc_hien_tai and ngay_ket_thuc_hien_tai != '':
                                    try:
                                        if hasattr(ngay_ket_thuc_hien_tai, 'strftime'):
                                            default_ngay_nghi = ngay_ket_thuc_hien_tai
                                    except:
                                        pass
                                
                                ngay_nghi = st.date_input(
                                    "📅 Ngày quyết định nghỉ việc (Ngày kết thúc HĐLĐ):", 
                                    value=default_ngay_nghi,
                                    key=f"ngay_nghi_{nid}"
                                )
                                
                                ly_do_nghi = st.text_area(
                                    "📝 Lý do nghỉ việc:", 
                                    value=nd.get('ly_do_nghi', '') if 'ly_do_nghi' in nd else '',
                                    placeholder="VD: Xin nghỉ theo nguyện vọng cá nhân, Chuyển công tác, Hết hạn hợp đồng...",
                                    key=f"ly_do_nghi_{nid}",
                                    height=100
                                )
                                
                                col_xac_nhan, col_huy = st.columns(2)
                                
                                with col_xac_nhan:
                                    if st.form_submit_button("✅ XÁC NHẬN NGHỈ VIỆC", use_container_width=True, type="primary"):
                                        try:
                                            db = get_connection()
                                            c = db.cursor()
                                            
                                            c.execute("""
                                                UPDATE nhan_vien 
                                                SET trang_thai = 'NGHI_VIEC', 
                                                    ngay_ket_thuc = %s,
                                                    ly_do_nghi = %s
                                                WHERE id = %s
                                            """, (ngay_nghi, ly_do_nghi if ly_do_nghi else None, nid))
                                            
                                            db.commit()
                                            db.close()
                                            
                                            st.session_state[f'nghi_viec_open_{nid}'] = False
                                            if 'selected_nv_id' in st.session_state:
                                                del st.session_state['selected_nv_id']
                                            
                                            st.success(f"✅ Đã cập nhật nghỉ việc cho {nd.get('ho_ten', '')}!")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"❌ Lỗi khi cập nhật nghỉ việc: {e}")
                                
                                with col_huy:
                                    if st.form_submit_button("❌ HỦY", use_container_width=True):
                                        st.session_state[f'nghi_viec_open_{nid}'] = False
                                        st.rerun()                          
                        with col_delete:
                            if st.form_submit_button("🗑️ XÓA"):
                                db = get_connection()
                                c = db.cursor()
                                c.execute("DELETE FROM ho_so_nhan_vien WHERE nhan_vien_id=%s", (nid,))
                                c.execute("DELETE FROM nhan_vien WHERE id=%s", (nid,))
                                db.commit()
                                db.close()
                                st.success("🗑️ Đã xóa!")
                                del st.session_state['selected_nv_id']
                                st.rerun()
                    
                    if st.button("❌ HỦY SỬA", use_container_width=True):
                        del st.session_state['selected_nv_id']
                        st.rerun()
            
            # Form nhập thông tin hộ gia đình (chỉ admin)
            if 'bhxh_family_nv_id' in st.session_state and st.session_state.role == "admin":
                nv_id = st.session_state['bhxh_family_nv_id']
                nv_name = st.session_state['bhxh_family_nv_name']
                st.divider()
                st.subheader(f"🏠 NHẬP THÔNG TIN HỘ GIA ĐÌNH CHO: {nv_name}")
                st.caption("Vui lòng nhập đầy đủ thông tin chủ hộ và các thành viên trong hộ gia đình")
                
                db = get_connection()
                c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                c.execute("SELECT * FROM nhan_vien WHERE id = %s", (nv_id,))
                nv_data = c.fetchone()
                db.close()
                
                if 'bhxh_family_members' not in st.session_state:
                    db_temp = get_connection()
                    c_temp = db_temp.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                    c_temp.execute("SELECT * FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_id,))
                    existing_members = c_temp.fetchall()
                    db_temp.close()
                    st.session_state['bhxh_family_members'] = []
                    for tv in existing_members:
                        st.session_state['bhxh_family_members'].append({
                            'ho_ten': tv['ho_ten'], 'ngay_sinh': tv['ngay_sinh'], 'gioi_tinh': tv['gioi_tinh'],
                            'quoc_tich': tv['quoc_tich'], 'dan_toc': tv['dan_toc'], 'quan_he': tv['quan_he_voi_chu_ho'],
                            'tinh': tv['tinh_thanh_pho'], 'phuong_xa': tv['phuong_xa']
                        })
                
                db_temp = get_connection()
                c_temp = db_temp.cursor()
                c_temp.execute("SELECT ma_tinh, ten_tinh FROM danh_muc_tinh ORDER BY ten_tinh")
                ds_tinh = c_temp.fetchall()
                db_temp.close()
                tinh_options = {ten: ma for ma, ten in ds_tinh}
                
                if st.session_state.bhxh_family_members:
                    st.markdown("**Danh sách thành viên đã thêm:**")
                    tv_data = []
                    for i, tv in enumerate(st.session_state.bhxh_family_members):
                        tv_data.append({"STT": i+1, "Họ và tên": tv['ho_ten'], "Ngày sinh": format_date(tv['ngay_sinh']),
                                        "Giới tính": tv['gioi_tinh'], "Quốc tịch": tv['quoc_tich'], "Dân tộc": tv['dan_toc'],
                                        "Quan hệ chủ hộ": tv['quan_he'], "Tỉnh/TP": tv['tinh'], "Phường/Xã": tv['phuong_xa']})
                    df_tv = pd.DataFrame(tv_data)
                    st.dataframe(df_tv, use_container_width=True, hide_index=True)
                    col_del1, col_del2, col_del3 = st.columns([1,1,1])
                    with col_del2:
                        tv_to_delete = st.number_input("Nhập STT thành viên cần xóa:", min_value=1, max_value=len(st.session_state.bhxh_family_members), step=1, key="tv_delete_family")
                        if st.button("🗑️ Xóa thành viên", key="btn_del_tv_family"):
                            st.session_state.bhxh_family_members.pop(tv_to_delete - 1)
                            st.rerun()
                
                with st.form(key=f"bhxh_family_form_{nv_id}"):
                    st.markdown("**I. THÔNG TIN CHỦ HỘ:**")
                    col1, col2 = st.columns(2)
                    with col1:
                        ho_ten_chu_ho = st.text_input("Họ và tên chủ hộ", value=nv_data.get('ho_ten_chu_ho', '') if nv_data else '')
                        so_cccd_chu_ho = st.text_input("Số CCCD chủ hộ", value=nv_data.get('so_cccd_chu_ho', '') if nv_data else '')
                        tinh_chu_ho_index = 0
                        tinh_chu_ho_current = nv_data.get('tinh_thanh_pho_chu_ho', '') if nv_data else ''
                        if tinh_chu_ho_current in tinh_options:
                            tinh_chu_ho_index = list(tinh_options.keys()).index(tinh_chu_ho_current) + 1
                        tinh_chu_ho = st.selectbox("Tỉnh/Thành phố (chủ hộ)", options=[""] + list(tinh_options.keys()), index=tinh_chu_ho_index)
                    with col2:
                        phuong_xa_options = []
                        phuong_xa_current = nv_data.get('phuong_xa_chu_ho', '') if nv_data else ''
                        if tinh_chu_ho and tinh_chu_ho != "":
                            ma_tinh = tinh_options.get(tinh_chu_ho)
                            db_temp2 = get_connection()
                            c_temp2 = db_temp2.cursor()
                            c_temp2.execute("SELECT ten_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh,))
                            phuong_xa_options = [row[0] for row in c_temp2.fetchall()]
                            db_temp2.close()
                        phuong_xa_index = 0
                        if phuong_xa_current in phuong_xa_options:
                            phuong_xa_index = phuong_xa_options.index(phuong_xa_current) + 1
                        phuong_xa_chu_ho = st.selectbox("Phường/Xã (chủ hộ)", options=[""] + phuong_xa_options, index=phuong_xa_index)
                    
                    st.markdown("**Thông tin thường trú:**")
                    col_tt1, col_tt2 = st.columns(2)
                    with col_tt1:
                        tinh_thuong_tru_index = 0
                        tinh_thuong_tru_current = nv_data.get('tinh_thanh_pho_thuong_tru', '') if nv_data else ''
                        if tinh_thuong_tru_current in tinh_options:
                            tinh_thuong_tru_index = list(tinh_options.keys()).index(tinh_thuong_tru_current) + 1
                        tinh_thuong_tru = st.selectbox("Tỉnh/Thành phố thường trú", options=[""] + list(tinh_options.keys()), index=tinh_thuong_tru_index)
                        ma_tinh_thuong_tru = tinh_options.get(tinh_thuong_tru, "") if tinh_thuong_tru else ""
                    with col_tt2:
                        phuong_xa_tt_options = []
                        phuong_xa_tt_current = nv_data.get('phuong_xa_thuong_tru', '') if nv_data else ''
                        if tinh_thuong_tru and tinh_thuong_tru != "":
                            ma_tinh_tt = tinh_options.get(tinh_thuong_tru)
                            db_temp3 = get_connection()
                            c_temp3 = db_temp3.cursor()
                            c_temp3.execute("SELECT ten_xa, ma_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh_tt,))
                            phuong_xa_tt_options = c_temp3.fetchall()
                            db_temp3.close()
                        phuong_xa_tt_index = 0
                        ma_phuong_xa_thuong_tru = ""
                        for i, px in enumerate(phuong_xa_tt_options):
                            if px[0] == phuong_xa_tt_current:
                                phuong_xa_tt_index = i + 1
                                ma_phuong_xa_thuong_tru = px[1]
                                break
                        phuong_xa_display_list = [""] + [px[0] for px in phuong_xa_tt_options]
                        phuong_xa_thuong_tru = st.selectbox("Phường/Xã thường trú", options=phuong_xa_display_list, index=phuong_xa_tt_index)
                        for px in phuong_xa_tt_options:
                            if px[0] == phuong_xa_thuong_tru:
                                ma_phuong_xa_thuong_tru = px[1]
                                break
                    
                    st.markdown("**II. THÊM THÀNH VIÊN MỚI:**")
                    st.caption("Điền thông tin vào các cột bên dưới, sau đó bấm '➕ Thêm thành viên'")
                    col_tv1, col_tv2, col_tv3, col_tv4, col_tv5, col_tv6, col_tv7, col_tv8 = st.columns([2,1.3,1,1,1,1.5,1.8,1.8])
                    with col_tv1:
                        ho_ten_tv = st.text_input("Họ và tên", key="tv_ho_ten_family", placeholder="Nguyễn Văn A")
                    with col_tv2:
                        ngay_sinh_tv = st.text_input("Ngày sinh", key="tv_ngay_sinh_family", placeholder="dd/mm/yyyy")
                    with col_tv3:
                        gioi_tinh_tv = st.selectbox("Giới tính", ["Nam", "Nữ"], key="tv_gioi_tinh_family")
                    with col_tv4:
                        quoc_tich_tv = st.text_input("Quốc tịch", value="Việt Nam", key="tv_quoc_tich_family")
                    with col_tv5:
                        dan_toc_tv = st.text_input("Dân tộc", value="Kinh", key="tv_dan_toc_family")
                    with col_tv6:
                        quan_he_tv = st.selectbox("Quan hệ chủ hộ", ["", "Vợ", "Chồng", "Con", "Bố", "Mẹ", "Anh", "Chị", "Em", "Ông", "Bà", "Khác"], key="tv_quan_he_family")
                    with col_tv7:
                        tinh_tv = st.selectbox("Tỉnh/Thành phố", options=[""] + list(tinh_options.keys()), key="tv_tinh_family")
                    with col_tv8:
                        phuong_xa_tv_options = []
                        if tinh_tv and tinh_tv != "":
                            ma_tinh_tv = tinh_options.get(tinh_tv)
                            db_temp4 = get_connection()
                            c_temp4 = db_temp4.cursor()
                            c_temp4.execute("SELECT ten_xa FROM danh_muc_phuong_xa WHERE ma_tinh = %s ORDER BY ten_xa", (ma_tinh_tv,))
                            phuong_xa_tv_options = [row[0] for row in c_temp4.fetchall()]
                            db_temp4.close()
                        phuong_xa_tv = st.selectbox("Phường/Xã", options=[""] + phuong_xa_tv_options, key="tv_phuong_xa_family")
                    
                    col_btn_add1, col_btn_add2, col_btn_add3 = st.columns([1,1,1])
                    with col_btn_add2:
                        if st.form_submit_button("➕ Thêm thành viên vào danh sách", use_container_width=True):
                            if ho_ten_tv:
                                st.session_state.bhxh_family_members.append({
                                    'ho_ten': ho_ten_tv, 'ngay_sinh': parse_date(ngay_sinh_tv), 'gioi_tinh': gioi_tinh_tv,
                                    'quoc_tich': quoc_tich_tv, 'dan_toc': dan_toc_tv, 'quan_he': quan_he_tv,
                                    'tinh': tinh_tv, 'phuong_xa': phuong_xa_tv
                                })
                                st.rerun()
                            else:
                                st.error("Vui lòng nhập họ tên thành viên")
                    
                    st.markdown("---")
                    col_save1, col_save2, col_save3 = st.columns([1,2,1])
                    with col_save2:
                        if st.form_submit_button("💾 LƯU THÔNG TIN CHỦ HỘ", use_container_width=True, type="primary"):
                            try:
                                db = get_connection()
                                c = db.cursor()
                                c.execute("""UPDATE nhan_vien SET ho_ten_chu_ho=%s, so_cccd_chu_ho=%s, tinh_thanh_pho_chu_ho=%s, phuong_xa_chu_ho=%s,
                                    tinh_thanh_pho_thuong_tru=%s, ma_tinh_thuong_tru=%s, phuong_xa_thuong_tru=%s, ma_phuong_xa_thuong_tru=%s WHERE id=%s""",
                                    (ho_ten_chu_ho, so_cccd_chu_ho, tinh_chu_ho, phuong_xa_chu_ho, tinh_thuong_tru, ma_tinh_thuong_tru, phuong_xa_thuong_tru, ma_phuong_xa_thuong_tru, nv_id))
                                c.execute("DELETE FROM phu_luc_gia_dinh WHERE nhan_vien_id = %s", (nv_id,))
                                for tv in st.session_state.bhxh_family_members:
                                    c.execute("""INSERT INTO phu_luc_gia_dinh (nhan_vien_id, ho_ten, ngay_sinh, gioi_tinh, quoc_tich, dan_toc, quan_he_voi_chu_ho, tinh_thanh_pho, phuong_xa) 
                                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                                        (nv_id, tv['ho_ten'], tv['ngay_sinh'], tv['gioi_tinh'], tv['quoc_tich'], tv['dan_toc'], tv['quan_he'], tv['tinh'], tv['phuong_xa']))
                                db.commit()
                                db.close()
                                del st.session_state['bhxh_family_nv_id']
                                del st.session_state['bhxh_family_nv_name']
                                del st.session_state['bhxh_family_members']
                                st.success(f"✅ Đã lưu thông tin hộ gia đình cho nhân viên {nv_name}")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi lưu: {e}")
                
                col_cancel1, col_cancel2, col_cancel3 = st.columns([1,2,1])
                with col_cancel2:
                    if st.button("❌ HỦY BỎ", use_container_width=True):
                        del st.session_state['bhxh_family_nv_id']
                        del st.session_state['bhxh_family_nv_name']
                        if 'bhxh_family_members' in st.session_state:
                            del st.session_state['bhxh_family_members']
                        st.rerun()
        
        else:
            st.info("Không có nhân viên nào đang làm việc")
    
    with tab_da_nghi:
        st.caption("📋 Danh sách nhân viên đã nghỉ việc (có thông tin ngày nghỉ)")
        
        col_filter1, col_filter2, col_filter3 = st.columns([2, 2, 1])
        with col_filter1:
            search_nghi = st.text_input("🔍 Tìm kiếm (Tên, Mã NV, SĐT, CCCD)", key="search_da_nghi")
        with col_filter2:
            db_temp = get_connection()
            c_temp = db_temp.cursor()
            c_temp.execute("SELECT DISTINCT EXTRACT(YEAR FROM ngay_ket_thuc) as nam FROM nhan_vien WHERE trang_thai='NGHI_VIEC' AND ngay_ket_thuc IS NOT NULL ORDER BY nam DESC")
            years = [int(row[0]) for row in c_temp.fetchall() if row[0] is not None]
            db_temp.close()
            filter_nam = st.selectbox("📅 Lọc theo năm nghỉ:", ["Tất cả"] + [str(y) for y in years] if years else ["Tất cả"])
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        sql = """
            SELECT id, ma_nv, ho_ten, ngay_sinh, gioi_tinh, so_cccd, dien_thoai, 
                   chuc_danh_nghe, loai_hop_dong, so_hdld, ngay_vao_lam, ngay_ket_thuc,
                   ma_so_bhxh, thang_bat_dau_bh, ly_do_nghi
            FROM nhan_vien 
            WHERE trang_thai = 'NGHI_VIEC'
        """
        params = []
        
        if search_nghi:
            sql += " AND (ho_ten LIKE %s OR ma_nv LIKE %s OR dien_thoai LIKE %s OR so_cccd LIKE %s)"
            params.extend([f'%{search_nghi}%'] * 4)
        
        if filter_nam != "Tất cả" and filter_nam.isdigit():
            sql += " AND EXTRACT(YEAR FROM ngay_ket_thuc) = %s"
            params.append(int(filter_nam))
        
        sql += " ORDER BY ngay_ket_thuc DESC, id DESC"
        c.execute(sql, tuple(params))
        ds_nghi = c.fetchall()
        db.close()
        
        if ds_nghi:
            df_nghi = pd.DataFrame(ds_nghi)
            for col in df_nghi.columns:
                if 'ngay' in col.lower():
                    df_nghi[col] = df_nghi[col].apply(format_date)
            
            display_cols_nghi = ['ma_nv', 'ho_ten', 'ngay_sinh', 'gioi_tinh', 'chuc_danh_nghe', 
                                 'so_hdld', 'ngay_vao_lam', 'ngay_ket_thuc', 'dien_thoai', 'ma_so_bhxh']
            available_cols_nghi = [c for c in display_cols_nghi if c in df_nghi.columns]
            df_show_nghi = df_nghi[available_cols_nghi]
            
            col_map_nghi = {
                'ma_nv': 'Mã NV',
                'ho_ten': 'Họ và tên',
                'ngay_sinh': 'Ngày sinh',
                'gioi_tinh': 'Giới tính',
                'chuc_danh_nghe': 'Chức danh',
                'so_hdld': 'Số HĐLĐ',
                'ngay_vao_lam': 'Ngày vào làm',
                'ngay_ket_thuc': '📅 Ngày nghỉ việc',
                'dien_thoai': 'SĐT',
                'ma_so_bhxh': 'Mã BHXH',
            }
            df_show_nghi.rename(columns=col_map_nghi, inplace=True)
            
            st.caption(f"📌 Tổng số: **{len(ds_nghi)}** nhân viên đã nghỉ việc")
            st.dataframe(df_show_nghi, use_container_width=True, hide_index=True, height=400)
            
            st.divider()
            st.subheader("🔍 Xem chi tiết / Khôi phục nhân viên")
            
            nv_options = {f"{nv['ma_nv']} - {nv['ho_ten']} (Nghỉ: {format_date(nv.get('ngay_ket_thuc'))})": nv['id'] for nv in ds_nghi}
            selected_nghi_name = st.selectbox("Chọn nhân viên đã nghỉ:", list(nv_options.keys()))
            selected_nghi_id = nv_options[selected_nghi_name]
            
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("SELECT * FROM nhan_vien WHERE id = %s", (selected_nghi_id,))
            nv_nghi_detail = c.fetchone()
            db.close()
            
            if nv_nghi_detail:
                col_detail1, col_detail2 = st.columns(2)
                with col_detail1:
                    st.markdown("**📋 Thông tin cá nhân**")
                    st.write(f"- **Mã NV:** {nv_nghi_detail.get('ma_nv', '')}")
                    st.write(f"- **Họ tên:** {nv_nghi_detail.get('ho_ten', '')}")
                    st.write(f"- **Ngày sinh:** {format_date(nv_nghi_detail.get('ngay_sinh'))}")
                    st.write(f"- **Giới tính:** {nv_nghi_detail.get('gioi_tinh', '')}")
                    st.write(f"- **CCCD:** {nv_nghi_detail.get('so_cccd', '')}")
                    st.write(f"- **SĐT:** {nv_nghi_detail.get('dien_thoai', '')}")
                    st.write(f"- **Chức danh:** {nv_nghi_detail.get('chuc_danh_nghe', '')}")
                
                with col_detail2:
                    st.markdown("**📅 Thông tin hợp đồng & nghỉ việc**")
                    st.write(f"- **Số HĐLĐ:** {nv_nghi_detail.get('so_hdld', '')}")
                    st.write(f"- **Loại HĐ:** {nv_nghi_detail.get('loai_hop_dong', '')}")
                    st.write(f"- **Ngày vào làm:** {format_date(nv_nghi_detail.get('ngay_vao_lam'))}")
                    st.write(f"- **📅 Ngày nghỉ việc:** **{format_date(nv_nghi_detail.get('ngay_ket_thuc'))}**")
                    st.write(f"- **Mã BHXH:** {nv_nghi_detail.get('ma_so_bhxh', '')}")
                    st.write(f"- **Lý do nghỉ:** {nv_nghi_detail.get('ly_do_nghi', 'Chưa có thông tin')}")
                
                if st.session_state.role == "admin":
                    st.divider()
                    col_restore1, col_restore2, col_restore3 = st.columns([1, 2, 1])
                    with col_restore2:
                        if st.button(f"🔄 KHÔI PHỤC NHÂN VIÊN - {nv_nghi_detail['ho_ten']}", use_container_width=True, type="primary"):
                            try:
                                db = get_connection()
                                c = db.cursor()
                                loai_hd = nv_nghi_detail.get('loai_hop_dong', '')
                                if loai_hd == 'Thử việc':
                                    trang_thai_moi = 'THU_VIEC'
                                else:
                                    trang_thai_moi = 'DANG_LAM'
                                c.execute("""
                                    UPDATE nhan_vien 
                                    SET trang_thai = %s, 
                                        ngay_ket_thuc = NULL
                                    WHERE id = %s
                                """, (trang_thai_moi, selected_nghi_id))
                                db.commit()
                                db.close()
                                st.success(f"✅ Đã khôi phục nhân viên {nv_nghi_detail['ho_ten']}!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Lỗi khi khôi phục: {e}")
        else:
            st.info("📭 Không có nhân viên nào đã nghỉ việc")
    
    with tab_qtct:
        st.caption("📜 Lịch sử công tác và quyết định nhân sự")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien ORDER BY id DESC")
        all_nv = c.fetchall()
        db.close()
        
        if all_nv:
            nv_options = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in all_nv}
            selected_nv_history = st.selectbox("🔍 Chọn nhân viên:", list(nv_options.keys()), key="history_nv")
            nv_id_history = nv_options[selected_nv_history]
            
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("SELECT * FROM nhan_vien WHERE id = %s", (nv_id_history,))
            nv_current = c.fetchone()
            
            st.markdown(f"""
            ### 📌 Thông tin hiện tại của {nv_current['ho_ten']} ({nv_current['ma_nv']})
            | Trường | Giá trị |
            |--------|---------|
            | Trạng thái | {'🟢 Đang làm' if nv_current['trang_thai'] == 'DANG_LAM' else '🔵 Thử việc' if nv_current['trang_thai'] == 'THU_VIEC' else '🔴 Đã nghỉ'} |
            | Loại hợp đồng | {nv_current['loai_hop_dong']} |
            | Ngày vào làm | {format_date(nv_current['ngay_vao_lam'])} |
            | Ngày chính thức | {format_date(nv_current.get('ngay_chinh_thuc')) or 'Chưa có'} |
            | Chức danh | {nv_current['chuc_danh_nghe']} |
            | Phòng ban | {nv_current['phong_ban_lam_viec']} |
            """)
            
            c.execute("""
                SELECT * FROM quyet_dinh_nhan_su 
                WHERE nhan_vien_id = %s 
                ORDER BY ngay_quyet_dinh DESC
            """, (nv_id_history,))
            quyet_dinh_list = c.fetchall()
            
            if quyet_dinh_list:
                st.markdown("### 📋 Các quyết định nhân sự")
                qd_data = []
                for i, qd in enumerate(quyet_dinh_list, 1):
                    loai_qd_map = {
                        'THU_VIEC': '📝 Quyết định thử việc',
                        'CHINH_THUC': '✅ Quyết định chính thức',
                        'DIEU_CHUYEN': '🔄 Quyết định điều chuyển',
                        'BO_NHIEM': '⭐ Quyết định bổ nhiệm',
                        'TANG_LUONG': '💰 Quyết định tăng lương',
                        'NGHI_VIEC': '🚫 Quyết định nghỉ việc'
                    }
                    qd_data.append({
                        "STT": i,
                        "Loại quyết định": loai_qd_map.get(qd['loai_quyet_dinh'], qd['loai_quyet_dinh']),
                        "Số quyết định": qd['so_quyet_dinh'] or '...',
                        "Ngày quyết định": format_date(qd['ngay_quyet_dinh']),
                        "Ngày hiệu lực": format_date(qd['ngay_hieu_luc']),
                        "Nội dung": (qd['noi_dung'][:50] + "...") if qd['noi_dung'] and len(qd['noi_dung']) > 50 else qd['noi_dung']
                    })
                df_qd = pd.DataFrame(qd_data)
                st.dataframe(df_qd, use_container_width=True, hide_index=True)
                
                with st.expander("🔍 Xem chi tiết quyết định"):
                    qd_options = {f"{format_date(qd['ngay_quyet_dinh'])} - {qd['loai_quyet_dinh']}": qd for qd in quyet_dinh_list}
                    selected_qd_name = st.selectbox("Chọn quyết định:", list(qd_options.keys()), key="qd_detail")
                    selected_qd = qd_options[selected_qd_name]
                    st.markdown(f"""
                    **📄 Chi tiết quyết định:**
                    - **Số quyết định:** {selected_qd.get('so_quyet_dinh', '...')}
                    - **Ngày quyết định:** {format_date(selected_qd.get('ngay_quyet_dinh'))}
                    - **Ngày hiệu lực:** {format_date(selected_qd.get('ngay_hieu_luc'))}
                    - **Loại quyết định:** {selected_qd.get('loai_quyet_dinh')}
                    - **Nội dung:** {selected_qd.get('noi_dung', '...')}
                    - **Người ký:** {selected_qd.get('nguoi_ky', COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC'))}
                    
                    **📊 Thay đổi:**
                    - Chức danh: {selected_qd.get('chuc_danh_cu', '...')} → {selected_qd.get('chuc_danh_moi', '...')}
                    - Phòng ban: {selected_qd.get('phong_ban_cu', '...')} → {selected_qd.get('phong_ban_moi', '...')}
                    - Loại HĐ: {selected_qd.get('loai_hop_dong_cu', '...')} → {selected_qd.get('loai_hop_dong_moi', '...')}
                    """)
            else:
                st.info("📭 Nhân viên này chưa có quyết định nào.")
            
            c.execute("""
                SELECT * FROM lich_su_cong_tac 
                WHERE nhan_vien_id = %s 
                ORDER BY tu_ngay ASC
            """, (nv_id_history,))
            lich_su_list = c.fetchall()
            
            if lich_su_list:
                st.markdown("### 📅 Lịch sử công tác")
                ls_data = []
                for i, ls in enumerate(lich_su_list, 1):
                    ls_data.append({
                        "STT": i,
                        "Từ ngày": format_date(ls['tu_ngay']),
                        "Đến ngày": format_date(ls['den_ngay']) if ls['den_ngay'] else "Đang làm",
                        "Chức danh": ls['chuc_danh'] or '',
                        "Phòng ban": ls['phong_ban'] or '',
                        "Loại HĐ": ls['loai_hop_dong'] or '',
                        "Hệ số lương": ls['he_so_luong'] or ''
                    })
                df_ls = pd.DataFrame(ls_data)
                st.dataframe(df_ls, use_container_width=True, hide_index=True, height=400)
            else:
                st.info("📭 Chưa có lịch sử công tác. Đang tạo từ dữ liệu hiện tại...")
                loai_hd_dung = nv_current['loai_hop_dong']
                if nv_current['trang_thai'] == 'THU_VIEC':
                    loai_hd_dung = 'Thử việc'
                c.execute("""
                    INSERT INTO lich_su_cong_tac (nhan_vien_id, tu_ngay, chuc_danh, phong_ban, noi_lam_viec, loai_hop_dong, he_so_luong)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (nv_id_history, nv_current['ngay_vao_lam'], nv_current['chuc_danh_nghe'], 
                      nv_current['phong_ban_lam_viec'], nv_current['noi_lam_viec'], 
                      loai_hd_dung, nv_current['he_so_luong']))
                db.commit()
                st.rerun()
            db.close()
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")
    
    st.divider()
    st.subheader("📊 Báo cáo tăng/giảm nhân sự trong kỳ")
    
    col_from, col_to, col_btn = st.columns([2, 2, 1])
    with col_from:
        tu_ngay_bc = st.date_input("Từ ngày:", value=date.today().replace(day=1), key="bc_tu")
    with col_to:
        den_ngay_bc = st.date_input("Đến ngày:", value=date.today(), key="bc_den")
    with col_btn:
        xuat_bc = st.button("📄 XUẤT BÁO CÁO WORD", use_container_width=True)
    
    if xuat_bc:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT ho_ten, chuc_danh_nghe, phong_ban_lam_viec, loai_hop_dong, ngay_vao_lam,
                   ngay_sinh, so_hdld, ngay_ky_hd
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND ngay_vao_lam BETWEEN %s AND %s
            ORDER BY ngay_vao_lam ASC
        """, (tu_ngay_bc, den_ngay_bc))
        tang_list = c.fetchall()
        c.execute("""
            SELECT ho_ten, chuc_danh_nghe, phong_ban_lam_viec, loai_hop_dong, ngay_vao_lam, ngay_ket_thuc,
                   ngay_sinh, so_hdld, ngay_ky_hd
            FROM nhan_vien 
            WHERE trang_thai = 'NGHI_VIEC'
            AND ngay_ket_thuc BETWEEN %s AND %s
            ORDER BY ngay_ket_thuc ASC
        """, (tu_ngay_bc, den_ngay_bc))
        giam_list = c.fetchall()
        db.close()
        if tang_list or giam_list:
            file_path = tao_bao_cao_tang_giam(tang_list, giam_list, tu_ngay_bc, den_ngay_bc)
            with open(file_path, "rb") as f:
                st.download_button(
                    label="📥 TẢI FILE BÁO CÁO (Word)",
                    data=f,
                    file_name=f"Bao_cao_tang_giam_{tu_ngay_bc.strftime('%d%m%Y')}_{den_ngay_bc.strftime('%d%m%Y')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.info("Không có biến động nhân sự trong kỳ.")

# ========== UPLOAD ==========
elif menu=="📁 Upload hồ sơ" and st.session_state.role=="admin":
    st.title("📁 Quản lý hồ sơ nhân viên")
    tab_upload, tab_list = st.tabs(["📤 UPLOAD HỒ SƠ", "📋 DANH SÁCH HỒ SƠ"])
    
    with tab_upload:
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien WHERE trang_thai IN ('DANG_LAM','THU_VIEC') ORDER BY id DESC")
        nvl = c.fetchall()
        db.close()
        
        if nvl:
            nd = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in nvl}
            cn = st.selectbox("📌 Chọn nhân viên:", list(nd.keys()))
            
            col1, col2 = st.columns(2)
            with col1:
                lh = st.selectbox("📂 Loại hồ sơ:", ["BANG_CAP", "CHUNG_CHI", "CCCD", "HOP_DONG", "SO_YEU_LY_LICH", "KHAC"])
            with col2:
                st.markdown("💡 **Hướng dẫn:**")
                st.caption("- BANG_CAP: Bằng cấp, chứng chỉ")
                st.caption("- CHUNG_CHI: Chứng chỉ nghề")
                st.caption("- CCCD: Căn cước công dân")
                st.caption("- HOP_DONG: Hợp đồng lao động")
                st.caption("- SO_YEU_LY_LICH: Sơ yếu lý lịch")
            
            fl = st.file_uploader("📎 Chọn file:", type=['pdf', 'jpg', 'png', 'jpeg', 'doc', 'docx'])
            
            if fl:
                st.info(f"📄 Tên file: {fl.name} | 📏 Kích thước: {fl.size/1024:.1f} KB")
            
            if fl and st.button("📤 UPLOAD", type="primary", use_container_width=True):
                nid = nd[cn]
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                fn = f"{nid}_{timestamp}_{lh}_{fl.name}"
                fp = os.path.join(UPLOAD_FOLDER, fn)
                
                with open(fp, "wb") as f:
                    f.write(fl.getbuffer())
                
                db = get_connection()
                c = db.cursor()
                c.execute("""
                    INSERT INTO ho_so_nhan_vien (nhan_vien_id, loai_ho_so, ten_file, duong_dan_file, ngay_upload) 
                    VALUES (%s, %s, %s, %s, CURRENT_DATE)
                """, (nid, lh, fl.name, fp))
                db.commit()
                db.close()
                
                st.success(f"✅ Đã upload thành công!\n📁 Lưu tại: {fp}")
                st.rerun()
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")
    
    with tab_list:
        st.subheader("📋 Danh sách hồ sơ đã upload")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ma_nv, ho_ten FROM nhan_vien ORDER BY id DESC")
        nvl = c.fetchall()
        db.close()
        
        if nvl:
            nd = {f"{x['ma_nv']} - {x['ho_ten']}": x['id'] for x in nvl}
            selected_nv = st.selectbox("🔍 Chọn nhân viên để xem hồ sơ:", list(nd.keys()), key="view_hoso")
            nv_id = nd[selected_nv]
            
            db = get_connection()
            c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            c.execute("""
                SELECT id, loai_ho_so, ten_file, duong_dan_file, ngay_upload 
                FROM ho_so_nhan_vien 
                WHERE nhan_vien_id = %s 
                ORDER BY ngay_upload DESC, id DESC
            """, (nv_id,))
            hs_list = c.fetchall()
            db.close()
            
            if hs_list:
                st.caption(f"📌 Tổng số: **{len(hs_list)}** hồ sơ")
                hs_data = []
                for i, hs in enumerate(hs_list, 1):
                    hs_data.append({
                        "STT": i,
                        "Loại hồ sơ": hs['loai_ho_so'],
                        "Tên file gốc": hs['ten_file'],
                        "Ngày upload": format_date(hs['ngay_upload']),
                        "ID": hs['id'],
                        "Đường dẫn": hs['duong_dan_file']
                    })
                df_hs = pd.DataFrame(hs_data)
                st.dataframe(df_hs[['STT', 'Loại hồ sơ', 'Tên file gốc', 'Ngày upload']], use_container_width=True, hide_index=True)
                
                st.divider()
                st.subheader("📄 Xem chi tiết & Tải xuống")
                hs_options = {f"{hs['loai_ho_so']} - {hs['ten_file']} (Ngày: {format_date(hs['ngay_upload'])})": hs for hs in hs_list}
                selected_hs_name = st.selectbox("Chọn hồ sơ:", list(hs_options.keys()))
                selected_hs = hs_options[selected_hs_name]
                
                col_info, col_download = st.columns([2, 1])
                with col_info:
                    st.markdown(f"""
                    **📋 Thông tin hồ sơ:**
                    - **Loại:** {selected_hs['loai_ho_so']}
                    - **Tên file gốc:** {selected_hs['ten_file']}
                    - **Đường dẫn:** `{selected_hs['duong_dan_file']}`
                    - **Ngày upload:** {format_date(selected_hs['ngay_upload'])}
                    """)
                
                with col_download:
                    if os.path.exists(selected_hs['duong_dan_file']):
                        with open(selected_hs['duong_dan_file'], "rb") as f:
                            st.download_button(
                                label="📥 TẢI HỒ SƠ",
                                data=f,
                                file_name=selected_hs['ten_file'],
                                mime="application/octet-stream",
                                use_container_width=True
                            )
                    else:
                        st.error("❌ File không tồn tại trên máy chủ!")
                
                st.divider()
                col_del1, col_del2, col_del3 = st.columns([1, 2, 1])
                with col_del2:
                    if st.button("🗑️ XÓA HỒ SƠ NÀY", use_container_width=True, type="secondary"):
                        try:
                            if os.path.exists(selected_hs['duong_dan_file']):
                                os.remove(selected_hs['duong_dan_file'])
                            db = get_connection()
                            c = db.cursor()
                            c.execute("DELETE FROM ho_so_nhan_vien WHERE id = %s", (selected_hs['id'],))
                            db.commit()
                            db.close()
                            st.success(f"✅ Đã xóa hồ sơ: {selected_hs['ten_file']}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Lỗi khi xóa: {e}")
            else:
                st.info(f"📭 Nhân viên này chưa có hồ sơ nào được upload.")
        else:
            st.info("⚠️ Chưa có nhân viên nào trong hệ thống!")

# ========== DANH MỤC CHỨC DANH ==========
elif menu == "⚙️ Danh mục" and st.session_state.role == "admin":
    st.title("⚙️ Quản lý danh mục Chức danh")
    if st.session_state.role == "admin":
        with st.expander("➕ Thêm chức danh mới", expanded=False):
            with st.form("add_chuc_danh"):
                ten_moi = st.text_input("Tên chức danh *"); mo_ta = st.text_area("Mô tả")
                if st.form_submit_button("💾 LƯU"):
                    if ten_moi:
                        db = get_connection(); c = db.cursor()
                        c.execute("SELECT COALESCE(MIN(t1.id + 1), 1) FROM vi_tri_cong_tac t1 LEFT JOIN vi_tri_cong_tac t2 ON t1.id + 1 = t2.id WHERE t2.id IS NULL AND t1.id >= 1")
                        id_trong = c.fetchone()[0]
                        c.execute("SELECT COALESCE(MAX(id),0) FROM vi_tri_cong_tac")
                        id_max = c.fetchone()[0]
                        id_moi = id_trong if id_trong <= id_max + 1 else id_max + 1
                        c.execute("INSERT INTO vi_tri_cong_tac (id, ten_vi_tri, ghi_chu) VALUES (%s, %s, %s)", (id_moi, ten_moi, mo_ta))
                        db.commit(); db.close(); st.success(f"✅ Đã thêm: {ten_moi}"); st.rerun()
                    else: st.error("Tên chức danh không được để trống!")
        st.subheader("📋 Danh sách chức danh")
        db = get_connection(); c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("SELECT id, ten_vi_tri, ghi_chu FROM vi_tri_cong_tac ORDER BY id")
        ds = c.fetchall(); db.close()
        if ds:
            df = pd.DataFrame(ds); df.columns = ['ID', 'Tên chức danh', 'Ghi chú']; st.dataframe(df, use_container_width=True, hide_index=True)
            st.divider(); cdx = st.number_input("Nhập ID cần xóa:", min_value=1, step=1)
            if st.button("🗑️ XÓA", key="del_cd"):
                db = get_connection(); c = db.cursor()
                c.execute("DELETE FROM vi_tri_cong_tac WHERE id=%s", (cdx,)); db.commit(); db.close(); st.success("🗑️ Đã xóa!"); st.rerun()
        else: st.info("Chưa có chức danh nào")

# ========== BHXH ==========
elif menu == "📋 BHXH":
    st.title("📋 Quản lý BHXH")
    
    t1, t2 = st.tabs(["📊 Tổng quan", "📝 Báo cáo tăng/giảm D02-LT"])
    
    with t1:
        st.subheader("📊 Tổng quan tình hình đóng BHXH")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        # Thống kê chung
        c.execute("SELECT COUNT(*) as tong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')")
        tong_ld = c.fetchone()['tong']
        
        c.execute("SELECT COUNT(*) as dang_dong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND trang_thai_bhxh = 'DANG_DONG'")
        dang_dong = c.fetchone()['dang_dong']
        
        c.execute("SELECT COUNT(*) as chua_dong FROM nhan_vien WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND trang_thai_bhxh = 'CHUA_DONG'")
        chua_dong = c.fetchone()['chua_dong']
        
        c.execute("SELECT COUNT(*) as da_nghi FROM nhan_vien WHERE trang_thai = 'NGHI_VIEC'")
        da_nghi = c.fetchone()['da_nghi']
        
        db.close()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("👥 Tổng lao động", tong_ld)
        col2.metric("✅ Đang đóng BHXH", dang_dong, delta=f"{dang_dong/tong_ld*100:.0f}%" if tong_ld > 0 else None)
        col3.metric("⏳ Chưa đóng BHXH", chua_dong, delta=f"-{chua_dong}" if chua_dong > 0 else None, delta_color="inverse")
        col4.metric("📋 Đã nghỉ việc", da_nghi)
        
        st.divider()
        
        # Danh sách lao động chưa đóng BHXH
        st.subheader("⚠️ Lao động chưa đóng BHXH")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        c.execute("""
            SELECT ma_nv, ho_ten, chuc_danh_nghe, ngay_vao_lam, loai_hop_dong, thang_bat_dau_bh
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC') AND trang_thai_bhxh = 'CHUA_DONG'
            ORDER BY ngay_vao_lam ASC
        """)
        chua_dong_list = c.fetchall()
        db.close()
        
        if chua_dong_list:
            df_chua_dong = pd.DataFrame(chua_dong_list)
            for col in df_chua_dong.columns:
                if 'ngay' in col.lower():
                    df_chua_dong[col] = df_chua_dong[col].apply(format_date)
            st.dataframe(df_chua_dong, use_container_width=True, hide_index=True)
            
            if st.session_state.role == "admin":
                st.warning("💡 Hướng dẫn: Vào menu '✅ Nhân viên' -> chọn nhân viên -> sửa thông tin -> cập nhật 'Bắt đầu BH' và chuyển trạng thái BHXH thành 'ĐANG ĐÓNG'")
        else:
            st.success("✅ Tất cả lao động đã được đăng ký đóng BHXH!")
    
    with t2:
        st.subheader("📝 Báo cáo tăng/giảm lao động tham gia BHXH (Mẫu D02-LT)")
        st.caption("Theo Thông tư 56/2017/TT-BYT và Quyết định 595/QĐ-BHXH")
        
        col_from, col_to = st.columns(2)
        with col_from:
            tu_ngay = st.date_input("📅 Từ ngày:", value=date(date.today().year, 1, 1), key="d02_tu")
        with col_to:
            den_ngay = st.date_input("📅 Đến ngày:", value=date.today(), key="d02_den")
        
        db = get_connection()
        c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        
        # Lao động tăng trong kỳ
        c.execute("""
            SELECT 
                ma_nv, ho_ten, ma_so_bhxh, ngay_sinh, gioi_tinh, so_cccd,
                chuc_danh_nghe, phong_ban_lam_viec, luong_bao_hiem, he_so_luong,
                COALESCE(thang_bat_dau_bh, ngay_vao_lam) as ngay_bat_dau,
                loai_hop_dong, so_hdld, ngay_vao_lam, thuong_tru
            FROM nhan_vien 
            WHERE trang_thai IN ('DANG_LAM', 'THU_VIEC')
            AND COALESCE(thang_bat_dau_bh, ngay_vao_lam) BETWEEN %s AND %s
            ORDER BY COALESCE(thang_bat_dau_bh, ngay_vao_lam) ASC
        """, (tu_ngay, den_ngay))
        tang_list = c.fetchall()
        
        # Lao động giảm trong kỳ
        c.execute("""
            SELECT 
                ma_nv, ho_ten, ma_so_bhxh, ngay_sinh, gioi_tinh, so_cccd,
                chuc_danh_nghe, phong_ban_lam_viec, luong_bao_hiem, he_so_luong,
                thang_ket_thuc_bh as ngay_ket_thuc,
                loai_hop_dong, so_hdld, ngay_vao_lam, thuong_tru, ly_do_nghi
            FROM nhan_vien 
            WHERE trang_thai = 'NGHI_VIEC'
            AND thang_ket_thuc_bh BETWEEN %s AND %s
            ORDER BY thang_ket_thuc_bh ASC
        """, (tu_ngay, den_ngay))
        giam_list = c.fetchall()
        db.close()
        
        col_tang, col_giam = st.columns(2)
        with col_tang:
            st.markdown(f"### 🟢 LAO ĐỘNG TĂNG ({len(tang_list)})")
            if tang_list:
                df_tang = pd.DataFrame(tang_list)
                for col in df_tang.columns:
                    if 'ngay' in col.lower():
                        df_tang[col] = df_tang[col].apply(format_date)
                st.dataframe(df_tang, use_container_width=True, hide_index=True, height=300)
            else:
                st.info("📭 Không có lao động tăng trong kỳ")
        
        with col_giam:
            st.markdown(f"### 🔴 LAO ĐỘNG GIẢM ({len(giam_list)})")
            if giam_list:
                df_giam = pd.DataFrame(giam_list)
                for col in df_giam.columns:
                    if 'ngay' in col.lower():
                        df_giam[col] = df_giam[col].apply(format_date)
                st.dataframe(df_giam, use_container_width=True, hide_index=True, height=300)
            else:
                st.info("📭 Không có lao động giảm trong kỳ")
        
        st.divider()
        
        # Chỉ admin mới được xuất Excel
        if st.session_state.role == "admin":
            if tang_list or giam_list:
                if st.button("📥 XUẤT EXCEL D02-LT (Mẫu báo cáo BHXH)", type="primary", use_container_width=True):
                    from openpyxl import Workbook
                    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
                    from openpyxl.utils import get_column_letter
                    
                    thin_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "D02-LT"
                    
                    # Header thông tin đơn vị
                    ten_cong_ty = COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA")
                    ma_don_vi_bhxh = COMPANY_CONFIG.get("ma_don_vi_BHXH", "................")
                    
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
                    ws['A1'] = ten_cong_ty
                    ws['A1'].font = Font(bold=True, size=13, name='Times New Roman')
                    ws['A1'].alignment = Alignment(horizontal='center')
                    
                    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
                    ws['A2'] = f"Mã đơn vị BHXH: {ma_don_vi_bhxh}"
                    ws['A2'].font = Font(size=11, name='Times New Roman')
                    ws['A2'].alignment = Alignment(horizontal='center')
                    
                    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=8)
                    ws['A3'] = f"BÁO CÁO TĂNG/GIẢM LAO ĐỘNG THAM GIA BHXH (Mẫu D02-LT)"
                    ws['A3'].font = Font(bold=True, size=12, name='Times New Roman')
                    ws['A3'].alignment = Alignment(horizontal='center')
                    
                    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=8)
                    ws['A4'] = f"(Từ ngày {tu_ngay.strftime('%d/%m/%Y')} đến ngày {den_ngay.strftime('%d/%m/%Y')})"
                    ws['A4'].font = Font(size=11, name='Times New Roman')
                    ws['A4'].alignment = Alignment(horizontal='center')
                    
                    # Tạo sheet riêng cho từng loại
                    current_row = 6
                    
                    # ===== DANH SÁCH TĂNG =====
                    ws.cell(row=current_row, column=1, value="I. DANH SÁCH LAO ĐỘNG TĂNG MỚI THAM GIA BHXH")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=12, name='Times New Roman')
                    current_row += 1
                    
                    if tang_list:
                        headers = ["STT", "Mã NV", "Họ và tên", "Mã số BHXH", "Ngày sinh", "Giới tính", "Số CCCD", "Ngày bắt đầu BH"]
                        for col_idx, header in enumerate(headers, 1):
                            cell = ws.cell(row=current_row, column=col_idx, value=header)
                            cell.font = Font(bold=True, size=10, name='Times New Roman')
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = thin_border
                            cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
                        
                        current_row += 1
                        for idx, nv in enumerate(tang_list, 1):
                            ws.cell(row=current_row, column=1, value=idx)
                            ws.cell(row=current_row, column=2, value=nv.get('ma_nv', ''))
                            ws.cell(row=current_row, column=3, value=nv.get('ho_ten', ''))
                            ws.cell(row=current_row, column=4, value=nv.get('ma_so_bhxh', ''))
                            ws.cell(row=current_row, column=5, value=format_date(nv.get('ngay_sinh')))
                            ws.cell(row=current_row, column=6, value='Nam' if nv.get('gioi_tinh') == 'Nam' else 'Nữ' if nv.get('gioi_tinh') == 'Nữ' else '')
                            ws.cell(row=current_row, column=7, value=nv.get('so_cccd', ''))
                            ws.cell(row=current_row, column=8, value=format_date(nv.get('ngay_bat_dau')))
                            
                            for col_idx in range(1, 9):
                                cell = ws.cell(row=current_row, column=col_idx)
                                cell.border = thin_border
                                if col_idx == 3:
                                    cell.alignment = Alignment(horizontal='left', vertical='center')
                                else:
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                            current_row += 1
                        
                        current_row += 1
                    else:
                        ws.cell(row=current_row, column=1, value="Không có lao động tăng trong kỳ")
                        ws.cell(row=current_row, column=1).font = Font(italic=True, size=10, name='Times New Roman')
                        current_row += 2
                    
                    # ===== DANH SÁCH GIẢM =====
                    ws.cell(row=current_row, column=1, value="II. DANH SÁCH LAO ĐỘNG GIẢM (NGHỈ VIỆC)")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=12, name='Times New Roman')
                    current_row += 1
                    
                    if giam_list:
                        headers = ["STT", "Mã NV", "Họ và tên", "Mã số BHXH", "Ngày sinh", "Giới tính", "Số CCCD", "Ngày kết thúc BH", "Lý do nghỉ"]
                        for col_idx, header in enumerate(headers, 1):
                            cell = ws.cell(row=current_row, column=col_idx, value=header)
                            cell.font = Font(bold=True, size=10, name='Times New Roman')
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = thin_border
                            cell.fill = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
                        
                        current_row += 1
                        for idx, nv in enumerate(giam_list, 1):
                            ws.cell(row=current_row, column=1, value=idx)
                            ws.cell(row=current_row, column=2, value=nv.get('ma_nv', ''))
                            ws.cell(row=current_row, column=3, value=nv.get('ho_ten', ''))
                            ws.cell(row=current_row, column=4, value=nv.get('ma_so_bhxh', ''))
                            ws.cell(row=current_row, column=5, value=format_date(nv.get('ngay_sinh')))
                            ws.cell(row=current_row, column=6, value='Nam' if nv.get('gioi_tinh') == 'Nam' else 'Nữ' if nv.get('gioi_tinh') == 'Nữ' else '')
                            ws.cell(row=current_row, column=7, value=nv.get('so_cccd', ''))
                            ws.cell(row=current_row, column=8, value=format_date(nv.get('ngay_ket_thuc')))
                            ws.cell(row=current_row, column=9, value=nv.get('ly_do_nghi', ''))
                            
                            for col_idx in range(1, 10):
                                cell = ws.cell(row=current_row, column=col_idx)
                                cell.border = thin_border
                                if col_idx in [3, 9]:
                                    cell.alignment = Alignment(horizontal='left', vertical='center')
                                else:
                                    cell.alignment = Alignment(horizontal='center', vertical='center')
                            current_row += 1
                        
                        current_row += 1
                    else:
                        ws.cell(row=current_row, column=1, value="Không có lao động giảm trong kỳ")
                        ws.cell(row=current_row, column=1).font = Font(italic=True, size=10, name='Times New Roman')
                        current_row += 2
                    
                    # Footer
                    ws.cell(row=current_row, column=1, value=f"Tổng số lao động tăng: {len(tang_list)}")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
                    current_row += 1
                    ws.cell(row=current_row, column=1, value=f"Tổng số lao động giảm: {len(giam_list)}")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=11, name='Times New Roman')
                    current_row += 2
                    
                    # Ký tên
                    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=8)
                    ws.cell(row=current_row, column=6, value="NGƯỜI LẬP BÁO CÁO")
                    ws.cell(row=current_row, column=6).font = Font(bold=True, size=11, name='Times New Roman')
                    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal='center')
                    current_row += 1
                    
                    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=8)
                    ws.cell(row=current_row, column=6, value="(Ký, ghi rõ họ tên)")
                    ws.cell(row=current_row, column=6).font = Font(size=10, name='Times New Roman', italic=True)
                    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal='center')
                    current_row += 2
                    
                    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=8)
                    ws.cell(row=current_row, column=6, value=COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC').upper())
                    ws.cell(row=current_row, column=6).font = Font(bold=True, size=11, name='Times New Roman')
                    ws.cell(row=current_row, column=6).alignment = Alignment(horizontal='center')
                    
                    # Điều chỉnh độ rộng cột
                    for col_idx in range(1, 10):
                        ws.column_dimensions[get_column_letter(col_idx)].width = 20
                    
                    filename = f"D02-LT_BHXH_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
                    wb.save(filename)
                    
                    with open(filename, "rb") as f:
                        st.download_button(
                            label="📥 TẢI FILE EXCEL D02-LT",
                            data=f,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    st.success(f"✅ Đã xuất báo cáo D02-LT với {len(tang_list)} lao động tăng và {len(giam_list)} lao động giảm")
            else:
                st.info("📭 Không có biến động lao động trong kỳ để xuất báo cáo")
        else:
            st.info("🔒 Chỉ Admin mới có quyền xuất file Excel báo cáo BHXH. Bạn đang ở chế độ xem (Viewer).")

# ========== BÁO CÁO TÌNH HÌNH SỬ DỤNG LAO ĐỘNG MẪU 01/PLI (EXCEL) ==========
elif menu == "📋 Báo cáo 01/PLI":
    st.title("📋 Báo cáo tình hình sử dụng lao động")
    st.caption("Theo mẫu 01/PLI Phụ lục I - Nghị định 145/2020/NĐ-CP (sửa đổi bởi Nghị định 35/2022/NĐ-CP)")
    
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    col1, col2 = st.columns(2)
    with col1:
        tu_ngay = st.date_input("📅 Từ ngày:", value=date(date.today().year, 1, 1), key="pli_tu")
    with col2:
        den_ngay = st.date_input("📅 Đến ngày:", value=date.today(), key="pli_den")
    
    db = get_connection()
    c = db.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    c.execute("""
        SELECT 
            nv.STT, nv.ma_nv, nv.ho_ten, nv.ma_so_bhxh, nv.ngay_sinh, nv.gioi_tinh,
            nv.so_cccd, nv.chuc_danh_nghe, nv.luong_bao_hiem, nv.he_so_luong,
            nv.phu_cap_chuc_vu, nv.phu_cap_tnvk, nv.phu_cap_tnn, nv.loai_hop_dong,
            nv.ngay_vao_lam, nv.ngay_ky_hd, nv.ngay_ket_thuc, nv.thang_bat_dau_bh,
            nv.thang_ket_thuc_bh, nv.so_hdld, nv.phong_ban_lam_viec, nv.noi_lam_viec
        FROM nhan_vien nv
        WHERE nv.trang_thai IN ('DANG_LAM', 'THU_VIEC')
        AND nv.ngay_vao_lam <= %s
        AND (nv.ngay_ket_thuc IS NULL OR nv.ngay_ket_thuc > %s)
        ORDER BY nv.STT ASC
    """, (den_ngay, den_ngay))
    ds_lao_dong = c.fetchall()
    db.close()
    
    st.info(f"📊 Tổng số lao động đang làm việc: **{len(ds_lao_dong)}** người")
    
    # Hiển thị bảng dữ liệu trước khi xuất (cho cả admin và viewer)
    if ds_lao_dong:
        st.subheader("📋 Danh sách lao động")
        df_preview = pd.DataFrame(ds_lao_dong)
        for col in df_preview.columns:
            if 'ngay' in col.lower():
                df_preview[col] = df_preview[col].apply(format_date)
        
        preview_cols = ['ma_nv', 'ho_ten', 'chuc_danh_nghe', 'loai_hop_dong', 'ngay_vao_lam', 'ma_so_bhxh']
        available_preview = [c for c in preview_cols if c in df_preview.columns]
        df_display = df_preview[available_preview]
        col_map_preview = {
            'ma_nv': 'Mã NV',
            'ho_ten': 'Họ tên',
            'chuc_danh_nghe': 'Chức danh',
            'loai_hop_dong': 'Loại HĐ',
            'ngay_vao_lam': 'Ngày vào làm',
            'ma_so_bhxh': 'Mã BHXH'
        }
        df_display.rename(columns=col_map_preview, inplace=True)
        st.dataframe(df_display, use_container_width=True, hide_index=True, height=400)
        
        st.divider()
        
        # Chỉ admin mới được xuất Excel
        if st.session_state.role == "admin":
            if st.button("📥 XUẤT EXCEL MẪU 01/PLI", type="primary", use_container_width=True):
                wb = Workbook()
                ws = wb.active
                ws.title = "BC_Tinh_hinh_su_dung_LD"
                
                ten_cong_ty = COMPANY_CONFIG.get("ten_cong_ty", "CÔNG TY CỔ PHẦN CẢNG HÒN LA")
                dia_chi = COMPANY_CONFIG.get("dia_chi", "")
                ma_so_thue = COMPANY_CONFIG.get("ma_so_thue", "")
                dien_thoai_cty = COMPANY_CONFIG.get("dien_thoai_cty", "")
                
                # Header
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
                ws['A1'] = ten_cong_ty
                ws['A1'].font = Font(bold=True, size=13, name='Times New Roman')
                ws['A1'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=1, start_column=20, end_row=1, end_column=26)
                ws['T1'] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
                ws['T1'].font = Font(bold=True, size=13, name='Times New Roman')
                ws['T1'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
                ws['A2'] = f"Số: 01/BC-01PLI-{datetime.now().year}/CHL"
                ws['A2'].font = Font(size=12, name='Times New Roman')
                ws['A2'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=2, start_column=20, end_row=2, end_column=26)
                ws['T2'] = "Độc lập - Tự do - Hạnh phúc"
                ws['T2'].font = Font(italic=True, size=12, name='Times New Roman')
                ws['T2'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=3, start_column=20, end_row=3, end_column=26)
                ws['T3'] = f"Quảng Trị, ngày {date.today().day} tháng {date.today().month} năm {date.today().year}"
                ws['T3'].font = Font(italic=True, size=12, name='Times New Roman')
                ws['T3'].alignment = Alignment(horizontal='right')
                
                ws.merge_cells('A5:AA5')
                ws['A5'] = "BÁO CÁO TÌNH HÌNH SỬ DỤNG LAO ĐỘNG"
                ws['A5'].font = Font(bold=True, size=13, name='Times New Roman')
                ws['A5'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells('A6:AA6')
                ws['A6'] = f"(Từ ngày {tu_ngay.strftime('%d/%m/%Y')} đến ngày {den_ngay.strftime('%d/%m/%Y')})"
                ws['A6'].font = Font(size=12, name='Times New Roman')
                ws['A6'].alignment = Alignment(horizontal='center')
                
                ws.merge_cells('A8:AA8')
                ws['A8'] = "Kính gửi: SỞ NỘI VỤ TỈNH QUẢNG TRỊ"
                ws['A8'].font = Font(bold=True, size=11, name='Times New Roman')
                ws['A8'].alignment = Alignment(horizontal='left')
                
                ws['A10'] = "1. Thông tin chung về doanh nghiệp:"
                ws['A10'].font = Font(bold=True, size=11, name='Times New Roman')
                
                row_info = 11
                for label in [f"- Tên doanh nghiệp: {ten_cong_ty}", f"- Địa chỉ: {dia_chi}", 
                             f"- Mã số thuế: {ma_so_thue}", f"- Điện thoại: {dien_thoai_cty}"]:
                    ws[f'A{row_info}'] = label
                    ws[f'A{row_info}'].font = Font(size=11, name='Times New Roman')
                    row_info += 1
                
                ws[f'A{row_info + 1}'] = "2. Thông tin tình hình sử dụng lao động của đơn vị:"
                ws[f'A{row_info + 1}'].font = Font(bold=True, size=11, name='Times New Roman')
                
                header_row = 18
                col_widths = [5, 25, 18, 15, 8, 18, 25, 12, 18, 18, 12, 15, 
                             12, 12, 12, 12, 15, 12, 12, 18, 18, 18, 18, 18, 18, 18, 20]
                for i, w in enumerate(col_widths, 1):
                    ws.column_dimensions[get_column_letter(i)].width = w
                
                stt_row = header_row + 3
                for col in range(1, 28):
                    ws.cell(row=stt_row, column=col, value=f"({col})")
                    ws.cell(row=stt_row, column=col).font = Font(size=9, name='Times New Roman')
                    ws.cell(row=stt_row, column=col).alignment = Alignment(horizontal='center')
                    ws.cell(row=stt_row, column=col).border = thin_border
                
                # Merge cells header
                ws.merge_cells(start_row=header_row, start_column=1, end_row=header_row+2, end_column=1)
                ws.cell(row=header_row, column=1, value="STT")
                
                ws.merge_cells(start_row=header_row, start_column=2, end_row=header_row+2, end_column=2)
                ws.cell(row=header_row, column=2, value="Họ và tên")
                
                ws.merge_cells(start_row=header_row, start_column=3, end_row=header_row+2, end_column=3)
                ws.cell(row=header_row, column=3, value="Mã số BHXH")
                
                ws.merge_cells(start_row=header_row, start_column=4, end_row=header_row+2, end_column=4)
                ws.cell(row=header_row, column=4, value="Ngày sinh")
                
                ws.merge_cells(start_row=header_row, start_column=5, end_row=header_row+2, end_column=5)
                ws.cell(row=header_row, column=5, value="Giới tính")
                
                ws.merge_cells(start_row=header_row, start_column=6, end_row=header_row+2, end_column=6)
                ws.cell(row=header_row, column=6, value="Số CCCD/Hộ chiếu")
                
                ws.merge_cells(start_row=header_row, start_column=7, end_row=header_row+2, end_column=7)
                ws.cell(row=header_row, column=7, value="Chức danh nghề, vị trí, công việc")
                
                ws.merge_cells(start_row=header_row, start_column=8, end_row=header_row, end_column=11)
                ws.cell(row=header_row, column=8, value="Vị trí việc làm (2)")
                
                ws.merge_cells(start_row=header_row, start_column=12, end_row=header_row, end_column=17)
                ws.cell(row=header_row, column=12, value="Tiền lương")
                
                ws.merge_cells(start_row=header_row, start_column=20, end_row=header_row, end_column=24)
                ws.cell(row=header_row, column=20, value="Loại và hiệu lực hợp đồng")
                
                ws.merge_cells(start_row=header_row, start_column=18, end_row=header_row+1, end_column=19)
                ws.cell(row=header_row, column=18, value="Ngành nghề nặng nhọc, độc hại")
                
                ws.merge_cells(start_row=header_row+1, start_column=13, end_row=header_row+1, end_column=17)
                ws.cell(row=header_row+1, column=13, value="Phụ cấp")
                
                ws.merge_cells(start_row=header_row+1, start_column=21, end_row=header_row+1, end_column=22)
                ws.cell(row=header_row+1, column=21, value="Hiệu lực HĐLĐ xác định thời hạn")
                
                ws.merge_cells(start_row=header_row+1, start_column=23, end_row=header_row+1, end_column=24)
                cell = ws.cell(row=header_row+1, column=23, value="Hiệu lực HĐLĐ khác (dưới 1 tháng, thử việc)")
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                
                ws.merge_cells(start_row=header_row+1, start_column=8, end_row=header_row+2, end_column=8)
                ws.cell(row=header_row+1, column=8, value="Nhà quản lý")
                
                ws.merge_cells(start_row=header_row+1, start_column=9, end_row=header_row+2, end_column=9)
                ws.cell(row=header_row+1, column=9, value="Chuyên môn kỹ thuật bậc cao")
                
                ws.merge_cells(start_row=header_row+1, start_column=10, end_row=header_row+2, end_column=10)
                ws.cell(row=header_row+1, column=10, value="Chuyên môn kỹ thuật bậc trung")
                
                ws.merge_cells(start_row=header_row+1, start_column=11, end_row=header_row+2, end_column=11)
                ws.cell(row=header_row+1, column=11, value="Khác")
                
                ws.merge_cells(start_row=header_row+1, start_column=12, end_row=header_row+2, end_column=12)
                ws.cell(row=header_row+1, column=12, value="Mức lương/Hệ số lương")
                
                ws.merge_cells(start_row=header_row+1, start_column=20, end_row=header_row+2, end_column=20)
                ws.cell(row=header_row+1, column=20, value="Ngày bắt đầu HĐLĐ không xác định thời hạn")
                
                ws.merge_cells(start_row=header_row, start_column=25, end_row=header_row+2, end_column=25)
                ws.cell(row=header_row, column=25, value="Thời điểm bắt đầu đóng BHXH")
                
                ws.merge_cells(start_row=header_row, start_column=26, end_row=header_row+2, end_column=26)
                ws.cell(row=header_row, column=26, value="Thời điểm kết thúc đóng BHXH")
                
                ws.merge_cells(start_row=header_row, start_column=27, end_row=header_row+2, end_column=27)
                ws.cell(row=header_row, column=27, value="Ghi chú")
                
                for row in range(header_row, header_row + 3):
                    for col in range(1, 28):
                        cell = ws.cell(row=row, column=col)
                        cell.border = thin_border
                
                ws.cell(row=header_row+2, column=13, value="Phụ cấp chức vụ")
                ws.cell(row=header_row+2, column=14, value="Phụ cấp thâm niên VK(%)")
                ws.cell(row=header_row+2, column=15, value="Phụ cấp thâm niên nghề (%)")
                ws.cell(row=header_row+2, column=16, value="Phụ cấp thâm niên nghề (%)")
                ws.cell(row=header_row+2, column=17, value="Các khoản bổ sung")
                ws.cell(row=header_row+2, column=18, value="Ngày bắt đầu")
                ws.cell(row=header_row+2, column=19, value="Ngày kết thúc")
                ws.cell(row=header_row+2, column=21, value="Ngày bắt đầu")
                ws.cell(row=header_row+2, column=22, value="Ngày kết thúc")
                ws.cell(row=header_row+2, column=23, value="Ngày bắt đầu")
                ws.cell(row=header_row+2, column=24, value="Ngày kết thúc")
                
                for row in range(header_row, header_row + 3):
                    for col in range(1, 28):
                        cell = ws.cell(row=row, column=col)
                        if cell.value:
                            cell.font = Font(bold=True, size=10, name='Times New Roman')
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            cell.border = thin_border
                
                for col in range(1, 28):
                    cell = ws.cell(row=stt_row, column=col)
                    cell.font = Font(size=9, name='Times New Roman')
                    cell.alignment = Alignment(horizontal='center')
                    cell.border = thin_border
                
                data_row = stt_row + 1
                for idx, nv in enumerate(ds_lao_dong, 1):
                    row = data_row + idx - 1
                    
                    ws.cell(row=row, column=1, value=idx)
                    ws.cell(row=row, column=2, value=nv.get('ho_ten', ''))
                    ws.cell(row=row, column=3, value=nv.get('ma_so_bhxh', ''))
                    ws.cell(row=row, column=4, value=format_date(nv.get('ngay_sinh')))
                    gt = nv.get('gioi_tinh', '')
                    ws.cell(row=row, column=5, value='Nam' if gt == 'Nam' else 'Nữ' if gt == 'Nữ' else '')
                    ws.cell(row=row, column=6, value=nv.get('so_cccd', ''))
                    ws.cell(row=row, column=7, value=nv.get('chuc_danh_nghe', ''))
                    
                    cd = (nv.get('chuc_danh_nghe') or '').lower()
                    is_quan_ly = any(x in cd for x in ['giám đốc', 'trưởng phòng', 'phó', 'quản lý'])
                    ws.cell(row=row, column=8, value='x' if is_quan_ly else '')
                    is_chuyen_mon_cao = (any(x in cd for x in ['kỹ thuật', 'kĩ thuật']) and any(x in cd for x in ['cao', 'chính', 'kỹ sư']))
                    ws.cell(row=row, column=9, value='x' if is_chuyen_mon_cao else '')
                    is_khac = any(x in cd for x in ['phổ thông', 'lao động', 'tạp vụ', 'bảo vệ', 'tạp vụ', 'lái xe'])
                    ws.cell(row=row, column=11, value='x' if is_khac else '')
                    is_trung = (not is_quan_ly) and (not is_chuyen_mon_cao) and (not is_khac)
                    ws.cell(row=row, column=10, value='x' if is_trung else '')
                    
                    luong = nv.get('luong_bao_hiem', '')
                    heso = nv.get('he_so_luong', '')
                    ws.cell(row=row, column=12, value=f"Hệ số: {heso}" if heso and str(heso).strip() else str(luong) if luong else '')
                    ws.cell(row=row, column=13, value=str(nv.get('phu_cap_chuc_vu', '')) if nv.get('phu_cap_chuc_vu') else '')
                    ws.cell(row=row, column=14, value=f"{nv.get('phu_cap_tnvk', '')}%" if nv.get('phu_cap_tnvk') else '')
                    ws.cell(row=row, column=15, value=f"{nv.get('phu_cap_tnn', '')}%" if nv.get('phu_cap_tnn') else '')
                    ws.cell(row=row, column=16, value='')
                    ws.cell(row=row, column=17, value='')
                    ws.cell(row=row, column=18, value='')
                    ws.cell(row=row, column=19, value='')
                    
                    loai_hd = nv.get('loai_hop_dong', '')
                    ngay_bd = nv.get('ngay_ky_hd') or nv.get('ngay_vao_lam')
                    ngay_kt = nv.get('ngay_ket_thuc')
                    ws.cell(row=row, column=20, value=format_date(ngay_bd) if loai_hd == 'Không xác định thời hạn' else '')
                    
                    if loai_hd == 'Xác định thời hạn':
                        ws.cell(row=row, column=21, value=format_date(ngay_bd))
                        ws.cell(row=row, column=22, value=format_date(ngay_kt) if ngay_kt else '')
                    else:
                        ws.cell(row=row, column=21, value='')
                        ws.cell(row=row, column=22, value='')
                    
                    if loai_hd == 'Thử việc':
                        ws.cell(row=row, column=23, value=format_date(ngay_bd))
                        ws.cell(row=row, column=24, value=format_date(ngay_kt) if ngay_kt else '')
                    else:
                        ws.cell(row=row, column=23, value='')
                        ws.cell(row=row, column=24, value='')
                    
                    ws.cell(row=row, column=25, value=format_date(nv.get('thang_bat_dau_bh')))
                    ws.cell(row=row, column=26, value=format_date(nv.get('thang_ket_thuc_bh')))
                    ws.cell(row=row, column=27, value=nv.get('so_hdld', ''))
                    
                    for col in range(1, 28):
                        cell = ws.cell(row=row, column=col)
                        cell.border = thin_border
                        cell.font = Font(size=10, name='Times New Roman')
                        if col in [1, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                
                total_row = data_row + len(ds_lao_dong)
                ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
                ws.cell(row=total_row, column=1, value=f"Tổng cộng: {len(ds_lao_dong)} người")
                ws.cell(row=total_row, column=1).font = Font(bold=True, size=10, name='Times New Roman')
                ws.cell(row=total_row, column=1).border = thin_border
                
                sign_row = total_row + 3
                ws.merge_cells(start_row=sign_row, start_column=23, end_row=sign_row, end_column=27)
                ws.cell(row=sign_row, column=23, value="ĐẠI DIỆN DOANH NGHIỆP")
                ws.cell(row=sign_row, column=23).font = Font(bold=True, size=11, name='Times New Roman')
                ws.cell(row=sign_row, column=23).alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=sign_row+1, start_column=23, end_row=sign_row+1, end_column=27)
                ws.cell(row=sign_row+1, column=23, value="(Ký, đóng dấu, ghi rõ họ tên)")
                ws.cell(row=sign_row+1, column=23).font = Font(size=10, name='Times New Roman')
                ws.cell(row=sign_row+1, column=23).alignment = Alignment(horizontal='center')
                
                ws.merge_cells(start_row=sign_row+2, start_column=23, end_row=sign_row+2, end_column=27)
                ws.cell(row=sign_row+2, column=23, value=COMPANY_CONFIG.get('dai_dien', 'GIÁM ĐỐC').upper())
                ws.cell(row=sign_row+2, column=23).font = Font(bold=True, size=11, name='Times New Roman')
                ws.cell(row=sign_row+2, column=23).alignment = Alignment(horizontal='center')
                
                filename = f"Bao_cao_01_PLI_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
                wb.save(filename)
                
                with open(filename, "rb") as f:
                    st.download_button(
                        label="📥 TẢI FILE EXCEL",
                        data=f,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                st.success(f"✅ Đã xuất báo cáo với {len(ds_lao_dong)} lao động")
        else:
            st.info("🔒 Chỉ Admin mới có quyền xuất file Excel báo cáo 01/PLI. Bạn đang ở chế độ xem (Viewer).")
            st.caption("💡 Với quyền Viewer, bạn có thể xem danh sách lao động ở trên nhưng không thể tải file Excel.")
    else:
        st.warning("⚠️ Không có lao động nào đang làm việc trong kỳ báo cáo!")
            
st.sidebar.divider()
st.sidebar.caption("© 2026 HRM-Port | Cảng biển quốc tế Hòn La")