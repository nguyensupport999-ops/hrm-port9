# chuan_hoa_the_thuc.py
"""Chuẩn hóa thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP (Phụ lục I, mục V).

Luồng: upload (.docx/.txt/.pdf) -> đọc thành danh sách dòng -> NHẬN DIỆN thành phần thể thức
(quốc hiệu, tên cơ quan, số, địa danh-ngày, tên loại, trích yếu, căn cứ, Phần/Chương/Mục/Điều/Khoản/Điểm,
nơi nhận, chữ ký...) -> DỰNG LẠI file Word đúng chuẩn.

Quy tắc cỡ chữ: văn bản 1-2 trang -> cỡ chữ lớn nhất 14 (bộ "14"); dài hơn 2 trang -> lớn nhất 13 (bộ "13").
Hai bộ cỡ chữ đều theo chú thích số 1 của NĐ 30: cỡ chữ trong cùng một văn bản phải thống nhất.
"""
import io
import math
import re
import shutil
import subprocess
import tempfile
import unicodedata
from datetime import date

import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
import numpy as np
import cv2

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    _CO_OCR = True
except ImportError:
    _CO_OCR = False


# ============================================================
# TIỀN XỬ LÝ ẢNH: LỌC CON DẤU ĐỎ + VÙNG CHỮ KÝ TAY
# ============================================================
def _loc_dau_do_va_chu_ky(img_bgr):
    """
    Nhận 1 ảnh BGR (numpy array từ OpenCV), trả về ảnh nhị phân đã loại:
      - Toàn bộ nét mực ĐỎ (con dấu, dấu CÔNG VĂN ĐẾN, dấu MẬT...)
      - Vùng chữ ký tay (mật độ nét dày, nét cong liên tục)
    Cách làm:
      1) Tách kênh HSV, mask vùng đỏ (2 dải hue: 0-10 và 170-180).
      2) Xoá vùng đỏ khỏi ảnh (tô trắng).
      3) Chuyển xám -> nhị phân adaptive.
      4) Loại các "blob" có mật độ nét bất thường (chữ ký tay, hình vẽ).
    """
    # --- 1. Lọc đỏ (con dấu, dấu đến, dấu mật) ---
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    m1 = cv2.inRange(hsv, (0, 60, 60), (10, 255, 255))
    m2 = cv2.inRange(hsv, (170, 60, 60), (180, 255, 255))
    mask_do = cv2.bitwise_or(m1, m2)
    # Nới rộng mask đỏ ra 1 chút để bắt cả viền mờ của con dấu
    mask_do = cv2.dilate(mask_do, np.ones((3, 3), np.uint8), iterations=1)
    # Xoá đỏ khỏi ảnh gốc (tô trắng — vì nền giấy vốn trắng)
    img_sach = img_bgr.copy()
    img_sach[mask_do > 0] = (255, 255, 255)

    # --- 2. Nhị phân hoá ---
    gray = cv2.cvtColor(img_sach, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    nhi_phan = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV,
        blockSize=25, C=15
    )

    # --- 3. Loại blob mật độ dày (chữ ký tay) ---
    # Chữ ký tay: mật độ điểm ảnh foreground trong bounding box rất cao (nét vẽ dày, cong),
    # diện tích lớn, tỷ lệ khung rộng.
    h_img, w_img = nhi_phan.shape
    contours, _ = cv2.findContours(nhi_phan, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    mask_xoa = np.zeros_like(nhi_phan)
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        dien_tich_bb = w * h
        if dien_tich_bb < 400:                    # bỏ qua đốm nhỏ
            continue
        vung = nhi_phan[y:y + h, x:x + w]
        mat_do = cv2.countNonZero(vung) / max(dien_tich_bb, 1)  # 0..1
        # Nét chữ in thường có mật độ ~0.15-0.35; chữ ký tay/ hình vẽ > 0.5
        # Đồng thời kích thước lớn bất thường (>5% diện tích trang) -> nghi chữ ký
        if mat_do > 0.5 and dien_tich_bb > 0.02 * h_img * w_img:
            cv2.rectangle(mask_xoa, (x, y), (x + w, y + h), 255, -1)
        # Chữ ký thường nằm ở góc dưới-phải, cao 80-200px, ngang 200-500px, nét mảnh nhưng cong
        elif (dien_tich_bb > 0.005 * h_img * w_img and
              mat_do > 0.35 and
              y > 0.6 * h_img and x > 0.4 * w_img):
            cv2.rectangle(mask_xoa, (x, y), (x + w, y + h), 255, -1)

    nhi_phan[mask_xoa > 0] = 0
    # Trả về ảnh đảo lại nền trắng chữ đen cho OCR
    return cv2.bitwise_not(nhi_phan)


def _ocr_pdf_scan(data):
    """PDF dạng scan (không có text layer). Render từng trang -> tiền xử lý -> OCR tiếng Việt."""
    if not _CO_OCR:
        raise ValueError(
            "PDF này là bản scan (không có lớp văn bản). Cần thư viện pdf2image + pytesseract "
            "và gói tesseract-ocr-vie để OCR. Vui lòng cài thêm hoặc gửi file .docx gốc."
        )
    pages = convert_from_bytes(data, dpi=300, fmt='jpeg')
    ket_qua = []
    for page in pages:
        img = cv2.cvtColor(np.array(page), cv2.COLOR_RGB2BGR)
        img_sach = _loc_dau_do_va_chu_ky(img)
        # OCR tiếng Việt; psm 6 = giả định 1 khối văn bản thống nhất (phù hợp văn bản HC)
        txt = pytesseract.image_to_string(img_sach, lang='vie', config='--psm 6')
        ket_qua.append(txt)
    return '\n'.join(ket_qua)


# ============================================================
# doc_dong() — điểm vào, tự phát hiện PDF có text layer hay scan
# ============================================================
def doc_dong(data, ten_file):
    ten = (ten_file or '').lower()
    if ten.endswith('.docx'):
        return doc_dong_tu_docx(data)

    if ten.endswith('.pdf'):
        txt = ''
        try:
            try:
                from pypdf import PdfReader
            except Exception:
                from PyPDF2 import PdfReader
            rd = PdfReader(io.BytesIO(data))
            txt = '\n'.join((pg.extract_text() or '') for pg in rd.pages)
        except Exception:
            txt = ''

        # Nếu PDF không có text layer (scan ảnh) -> chuyển sang OCR
        if len(_chuan_dong(txt)) < 30:
            txt = _ocr_pdf_scan(data)
    else:
        txt = ''
        for enc in ('utf-8-sig', 'utf-16', 'cp1258', 'latin-1'):
            try:
                txt = data.decode(enc)
                break
            except Exception:
                continue

    return [x for x in (_chuan_dong(l) for l in txt.splitlines()) if x]
    
# ============================================================
# LỌC HẬU KỲ: các dòng rác sinh ra từ con dấu / chữ ký tay / ghi chú tay
# ============================================================
RE_RAC_KY_TU = re.compile(
    r'^[^A-Za-zÀ-ỹ0-9]+$'                       # chỉ có ký tự lạ, không chữ/số
)
RE_RAC_MAT_DO = re.compile(
    r'[^\w\sÀ-ỹ,.()\[\]{}\-/:;%&"\'+*=]{3,}'    # 3+ ký tự không phải chữ/số/dấu câu thông dụng
)
RE_RAC_QUA_NGAN = re.compile(r'^.{1,3}$')       # 1-3 ký tự lẻ loi
RE_DAU_VP_CHAN = re.compile(
    r'^(m\.?s\.?d\.?n|mst|số\s*đt|đt|fax)\b',   # các nhãn hay thấy trên con dấu
    re.I
)
RE_CHU_KY_TAY_HINT = re.compile(
    r'(ký\s*tên|ký,\s*ghi|ghi\s*rõ\s*họ\s*tên|đã\s*ký)',  # gợi ý chữ ký tay
    re.I
)


def _la_dong_rac(t, tong_chu, tong_so):
    """True nếu dòng gần như chắc chắn là rác từ con dấu/chữ ký/ghi chú tay."""
    if not t:
        return True
    if RE_BANG.match(t):      # chỗ giữ vị trí của bảng dữ liệu — không phải rác
        return False
    if re.match(r'^[-–—•+]?\s*lưu\s*:', t, re.I):   # dòng "- Lưu: VT." của Nơi nhận ít chữ nên dễ bị tính nhầm là rác
        return False

    # 1. Dòng toàn ký tự lạ
    if RE_RAC_KY_TU.match(t) and len(t) >= 2:
        return True

    # 2. Dòng quá ngắn mà không phải số/ký hiệu hợp lệ
    if RE_RAC_QUA_NGAN.match(t) and not re.match(r'^\d+[.)]?$', t):
        return True

    # 3. Tỷ lệ ký tự "chữ + số hợp lệ" quá thấp (< 60%) => rác
    tong_hop_le = tong_chu + tong_so
    if len(t) >= 8 and tong_hop_le / max(len(t), 1) < 0.55:
        return True

    # 4. Chứa cụm ký tự lạ dài (VD "«¥§€…" lẫn vào)
    if RE_RAC_MAT_DO.search(t):
        return True

    # 5. Nhãn đặc trưng của con dấu công ty (M.S.D.N, Mã số thuế trên dấu)
    if RE_DAU_VP_CHAN.match(t) and len(t) < 40:
        return True

    return False


def _loc_dong_rac(lines):
    """Loại các dòng nghi là rác do con dấu / chữ ký tay / ghi chú tay sinh ra."""
    ket_qua = []
    for t in lines:
        chu = sum(1 for c in t if c.isalpha())
        so = sum(1 for c in t if c.isdigit())
        if _la_dong_rac(t, chu, so):
            continue
        ket_qua.append(t)
    return ket_qua

FONT = 'Times New Roman'
_CAN = {'left': WD_ALIGN_PARAGRAPH.LEFT, 'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}

# Bộ cỡ chữ theo chú thích 1 (NĐ 30): quốc hiệu/tiêu ngữ/địa danh thống nhất trong 1 văn bản
CO_CHU = {
    14: dict(qh=13, tn=14, cq=13, so=13, ngay=14, loai=14, ty=14, vv=12, nd=14, ky=14, nn=12, nn_ds=11),
    13: dict(qh=12, tn=13, cq=12, so=13, ngay=13, loai=13, ty=13, vv=12, nd=13, ky=13, nn=12, nn_ds=11),
}

# ───────────────────────── Tiện ích văn bản ─────────────────────────


def _nfc(t):
    return unicodedata.normalize('NFC', t or '')


def khong_dau(t):
    t = unicodedata.normalize('NFD', _nfc(t).replace('đ', 'd').replace('Đ', 'D'))
    return ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn').lower()


def _chuan_dong(t):
    return re.sub(r'\s+', ' ', _nfc(t).replace('\u00a0', ' ')).strip()


def _la_in_hoa(t):
    chu = [c for c in t if c.isalpha()]
    return bool(chu) and sum(1 for c in chu if c.isupper()) / len(chu) >= 0.85


LOAI_VB = {
    'quyet dinh', 'nghi quyet', 'chi thi', 'thong bao', 'bao cao', 'to trinh', 'ke hoach', 'quy che', 'quy dinh',
    'huong dan', 'bien ban', 'giay moi', 'thu moi', 'giay gioi thieu', 'chuong trinh', 'de an', 'phuong an',
    'thong cao', 'thong tri', 'cong dien', 'giay uy quyen', 'quy trinh', 'noi quy', 'cam ket', 'dieu le',
    'hop dong', 'thoa thuan', 'bien ban ban giao', 'bien ban hop', 'danh sach', 'don xin', 'giay xac nhan',
}
CHUC_VU_TU_KHOA = ('giam doc', 'chu tich', 'truong phong', 'pho', 'tong giam doc', 'chu ho', 'truong', 'ke toan',
                   'thu truong', 'bo truong', 'hieu truong', 'chu nhiem', 'nguoi dai dien')

# ───────────────────────── 1. Đọc file ─────────────────────────


def doc_dong_tu_docx(data):
    d = Document(io.BytesIO(data))
    dong = []

    def _tu_paragraph(p):
        t = _chuan_dong(p.text)
        if t:
            dong.append(t)

    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for child in d.element.body.iterchildren():
        if child.tag == qn('w:p'):
            _tu_paragraph(Paragraph(child, d))
        elif child.tag == qn('w:tbl'):
            tb = Table(child, d)
            for row in tb.rows:
                da_xet = set()
                for cell in row.cells:
                    if id(cell._tc) in da_xet:
                        continue
                    da_xet.add(id(cell._tc))
                    for p in cell.paragraphs:
                        _tu_paragraph(p)
    return dong

def _la_bang_bo_cuc(tb):
    """True nếu bảng chỉ dùng để bố cục (quốc hiệu | tên cơ quan, nơi nhận | chữ ký, kính gửi...) — khi đó
    app đọc phẳng như cũ. Bảng còn lại là BẢNG DỮ LIỆU -> giữ lưới."""
    try:
        so_hang, so_cot = len(tb.rows), len(tb.columns)
        for r in tb.rows:
            for c in r.cells:
                k = khong_dau(c.text.strip())
                if RE_QH.search(k) or RE_NOI_NHAN.match(k) or k.startswith('kinh gui'):
                    return True
        return so_hang == 1 and so_cot <= 2
    except Exception:
        return False


def _doc_luoi_bang(tb):
    """Đọc bảng Word thành lưới: list các hàng; mỗi ô = dict(t=chữ, span=số cột gộp, vm=None|'restart'|'cont')."""
    from docx.text.paragraph import Paragraph
    rows = []
    for tr in tb._tbl.tr_lst:
        hang = []
        for tc in tr.tc_lst:
            span, vm = 1, None
            tcPr = tc.tcPr
            if tcPr is not None:
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    span = int(gs.get(qn('w:val')) or 1)
                v = tcPr.find(qn('w:vMerge'))
                if v is not None:
                    vm = 'restart' if v.get(qn('w:val')) == 'restart' else 'cont'
            parts = [_chuan_dong(Paragraph(p, tb).text) for p in tc.p_lst]
            hang.append(dict(t='\n'.join(x for x in parts if x), span=max(span, 1), vm=vm))
        rows.append(hang)
    return rows


def doc_dong_va_bang(data, ten_file):
    """Giống doc_dong, nhưng với .docx thì GIỮ LẠI bảng dữ liệu để dựng lại lưới.
    Trả về (lines, bang): bảng dữ liệu thứ n được đại diện trong lines bằng dòng '[[#n]]'."""
    if not (ten_file or '').lower().endswith('.docx'):
        return doc_dong(data, ten_file), []
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = Document(io.BytesIO(data))
    dong, bang = [], []

    def _them(p):
        t = _chuan_dong(p.text)
        if t:
            dong.append(t)

    for child in d.element.body.iterchildren():
        if child.tag == qn('w:p'):
            _them(Paragraph(child, d))
        elif child.tag == qn('w:tbl'):
            tb = Table(child, d)
            if _la_bang_bo_cuc(tb):
                for row in tb.rows:
                    da_xet = set()
                    for cell in row.cells:
                        if cell._tc in da_xet:
                            continue
                        da_xet.add(cell._tc)
                        for p in cell.paragraphs:
                            _them(p)
            else:
                rows = _doc_luoi_bang(tb)
                if any(c['t'] for h in rows for c in h):
                    dong.append(f'[[#{len(bang)}]]')
                    bang.append(rows)
    return dong, bang

def _la_file_scan(data):
    """True nếu PDF không có text layer (chỉ có ảnh)."""
    try:
        try:
            from pypdf import PdfReader
        except Exception:
            from PyPDF2 import PdfReader
        rd = PdfReader(io.BytesIO(data))
        tong = sum(len((pg.extract_text() or '').strip()) for pg in rd.pages)
        return tong < 30
    except Exception:
        return False
        
def _doc_pdf_co_text(data):
    """Đọc PDF có lớp text sẵn (PDF gốc từ Word/Google Docs xuất ra, không phải scan).
    Trả về chuỗi text; rỗng nếu PDF không có lớp text (PDF scan)."""
    try:
        from pypdf import PdfReader
    except Exception:
        try:
            from PyPDF2 import PdfReader
        except Exception:
            return None  # không có thư viện nào -> báo hiệu để caller xử lý
    rd = PdfReader(io.BytesIO(data))
    return '\n'.join((pg.extract_text() or '') for pg in rd.pages)


def _doc_pdf_bang_ocr(data, ngon_ngu='vie+eng', dpi=300):
    """OCR PDF scan (ảnh) bằng pdf2image + pytesseract.
    ngon_ngu: 'vie+eng' = nhận dạng tiếng Việt + tiếng Anh (Tesseract cần gói tesseract-ocr-vie).
    dpi: càng cao càng chính xác nhưng chậm hơn; 300 là mức cân bằng tốt.
    Trả về chuỗi text, hoặc ném ValueError nếu thiếu thư viện/gói hệ thống."""
    try:
        from pdf2image import convert_from_bytes
    except Exception as e:
        raise ValueError(
            "Đây là PDF dạng ảnh quét (scan), cần OCR để đọc. "
            "Nhưng app chưa cài `pdf2image`. Vui lòng thêm `pdf2image` vào requirements.txt."
        ) from e
    try:
        import pytesseract
    except Exception as e:
        raise ValueError(
            "Đây là PDF dạng ảnh quét (scan), cần OCR để đọc. "
            "Nhưng app chưa cài `pytesseract`. Vui lòng thêm `pytesseract` vào requirements.txt."
        ) from e

    try:
        cac_trang = convert_from_bytes(data, dpi=dpi)
    except Exception as e:
        msg = str(e).lower()
        if 'poppler' in msg or 'pdftoppm' in msg or 'page count' in msg:
            raise ValueError(
                "Đây là PDF dạng ảnh quét (scan), cần OCR để đọc. "
                "Nhưng hệ thống thiếu `poppler-utils` (công cụ chuyển PDF sang ảnh). "
                "Vui lòng thêm `poppler-utils` vào file packages.txt."
            ) from e
        raise ValueError(f"Không chuyển được PDF sang ảnh để OCR: {e}") from e

    ds_text = []
    for i, anh in enumerate(cac_trang, 1):
        try:
            t = pytesseract.image_to_string(anh, lang=ngon_ngu)
            ds_text.append(t)
        except pytesseract.TesseractError as e:
            # Gói ngôn ngữ 'vie' chưa cài -> thử lại chỉ với 'eng'
            if 'vie' in str(e).lower() or 'failed loading language' in str(e).lower():
                try:
                    t = pytesseract.image_to_string(anh, lang='eng')
                    ds_text.append(t)
                    continue
                except Exception:
                    pass
            raise ValueError(
                f"Lỗi OCR trang {i}: {e}. "
                "Kiểm tra đã cài `tesseract-ocr` và `tesseract-ocr-vie` trong packages.txt chưa."
            ) from e
    return '\n'.join(ds_text)


def doc_dong(data, ten_file):
    ten = (ten_file or '').lower()
    if ten.endswith('.docx'):
        return doc_dong_tu_docx(data)

    if ten.endswith('.pdf'):
        # --- Bước 1: thử đọc lớp text sẵn có (PDF gốc từ Word) ---
        txt = _doc_pdf_co_text(data)
        if txt is None:
            raise ValueError(
                "Không đọc được PDF: thiếu thư viện `pypdf`. "
                "Vui lòng thêm `pypdf` vào requirements.txt."
            )
        # --- Bước 2: nếu PDF không có text (PDF scan) -> tự động OCR ---
        if len(txt.strip()) < 50:
            # Rất ít/không có chữ -> gần như chắc chắn là PDF scan, tự động OCR
            txt = _doc_pdf_bang_ocr(data)
        # --- Bước 3: nếu text quá ngắn sau cả 2 bước -> báo lỗi rõ ràng ---
        if len(txt.strip()) < 20:
            raise ValueError(
                "Đọc PDF xong nhưng không trích được nội dung. "
                "File có thể là ảnh scan chất lượng thấp / chữ viết tay / bảng biểu phức tạp. "
                "Vui lòng kiểm tra lại file gốc hoặc chuyển sang PDF có lớp text (xuất từ Word)."
            )
    else:
        # .txt: thử lần lượt các bảng mã phổ biến
        txt = None
        for enc in ('utf-8-sig', 'utf-16', 'cp1258', 'latin-1'):
            try:
                txt = data.decode(enc)
                break
            except Exception:
                continue
        if txt is None:
            raise ValueError("Không đọc được file text: bảng mã không nhận dạng được.")

    return [x for x in (_chuan_dong(l) for l in txt.splitlines()) if x]


# ───────────────────────── 2. Nhận diện thành phần ─────────────────────────
RE_QH = re.compile(r'cong\s*hoa\s*xa\s*hoi\s*chu\s*nghia\s*viet\s*nam')
RE_TN = re.compile(r'doc\s*lap\s*[-–—]*\s*tu\s*do\s*[-–—]*\s*hanh\s*phuc')
RE_SO = re.compile(r'^so\s*[:.]')
RE_NGAY = re.compile(r'^[^,]{2,60},\s*ngay(?:\s+[^\s]+)?\s+thang(?:\s+[^\s]+)?\s+nam(?:\s+[^\s]+)?')
RE_CAN_CU = re.compile(r'^(can cu|xet de nghi|theo de nghi|xet|sau khi|nhan thay)\b')
RE_DIEU = re.compile(r'^dieu\s+\d+\s*[\.:]')
RE_KHOAN = re.compile(r'^\d+\.\s+\S')
RE_DIEM = re.compile(r'^[a-zđ]\)\s+\S')
RE_ROMAN = re.compile(r'^[ivx]{1,4}\.\s+\S')
RE_PHAN = re.compile(r'^phan\s+([ivxlc]+|\d+|thu\s+\S+)\b')
RE_CHUONG = re.compile(r'^chuong\s+([ivxlc]+|\d+)\b')
RE_MUC = re.compile(r'^muc\s+(\d+|[ivxlc]+)\b')
RE_TIEUMUC = re.compile(r'^tieu\s*muc\s+\d+')
RE_PHULUC = re.compile(r'^phu\s*luc\b')
RE_KHAN = re.compile(r'^(hoa toc|thuong khan|khan|hoa toc hen gio)$')
RE_QUYEN_HAN = re.compile(r'^(tm|kt|tl|tuq|q)\.\s*')
RE_NOI_NHAN = re.compile(r'^noi nhan\s*:?')
RE_GACH = re.compile(r'^[-–—•+]\s*')
RE_BANG = re.compile(r'^\[\[#(\d+)\]\]$')   # chỗ giữ vị trí của 1 bảng dữ liệu trong danh sách dòng
RE_DINH_KEM = re.compile(r'^(phu\s*luc|danh\s*sach|bang\s*ke|bieu\s*mau|bang\s*(cham|luong|tong|tinh)|mau\s*so\b|bieu\s*so\b)')
RE_TIEUDE_PL = re.compile(r'^phu\s*luc(\s+([ivxlc]+|\d+|[a-z]))?\s*[:.\-–]?\s*$')
RE_DIEN_GIAI = re.compile(r'^\(?\s*(ban hanh kem theo|kem theo|ban hanh|dinh kem)\b')


def _la_loai_vb(dong):
    kd = khong_dau(dong).rstrip(' :')
    return _la_in_hoa(dong) and kd in LOAI_VB


def _la_dong_ky(dong):
    kd = khong_dau(dong)
    return bool(RE_QUYEN_HAN.match(kd)) or (_la_in_hoa(dong) and len(dong) <= 70)


def chuan_vv(t):
    """'V/v: Đề nghị ... .' -> 'V/v Đề nghị ...' (NĐ 30: không dấu hai chấm, không chấm cuối trích yếu)."""
    t = re.sub(r'^\s*v\s*/\s*v\s*[:.]?\s*', 'V/v ', t, flags=re.I)
    return _chuan_dong(t).rstrip(' .')


def _la_bang(t):
    return bool(RE_BANG.match(t or ''))


def _la_tieu_de_phu_luc(t):
    k = khong_dau(t)
    if not RE_PHULUC.match(k) or len(t) > 80:
        return False
    return (_la_in_hoa(t) or bool(RE_TIEUDE_PL.match(k))
            or bool(re.match(r'^phu\s*luc\s+([ivxlc]+|\d+|[a-z])\s*[:.\-–]', k)))


def _la_dau_dinh_kem(t, chat=False):
    """Dòng mở đầu phần đính kèm (Phụ lục / Danh sách / Bảng kê / 1 bảng dữ liệu) nằm SAU phần chữ ký.
    chat=True (không có mốc 'Nơi nhận' để định vị): chỉ nhận dòng 'Phụ lục ...' cho an toàn."""
    if chat:
        return _la_tieu_de_phu_luc(t)
    if _la_bang(t) or _la_tieu_de_phu_luc(t):
        return True
    return bool(RE_DINH_KEM.match(khong_dau(t))) and len(t) <= 80


def _vi_tri_dinh_kem(lines, chat=False, tu=0):
    for i in range(tu, len(lines)):
        if _la_dau_dinh_kem(lines[i], chat):
            return i
    return None


def _phan_than(body_lines, co_loai):
    """Phân loại từng dòng thân văn bản (Căn cứ/Điều/Khoản/Điểm/bảng...). Trả về list [loai, text]."""
    body = []
    da_den_dieu_hoac_qd = False
    n = len(body_lines)
    i = 0
    while i < n:
        t = body_lines[i]
        k = khong_dau(t)
        loai = 'text'
        if _la_bang(t):
            loai = 'bang'
        elif RE_KHAN.match(k) and _la_in_hoa(t):
            loai = 'khan'
        elif RE_PHULUC.match(k):
            loai = 'phu_luc'
        elif RE_PHAN.match(k) and len(t) <= 40:
            loai = 'cap_tieu_de'
        elif RE_CHUONG.match(k) and len(t) <= 40:
            loai = 'cap_tieu_de'
        elif RE_TIEUMUC.match(k) and len(t) <= 40:
            loai = 'cap_tieu_de'
        elif RE_MUC.match(k) and len(t) <= 40:
            loai = 'cap_tieu_de'
        elif RE_DIEU.match(k):
            loai = 'dieu'; da_den_dieu_hoac_qd = True
        elif _la_in_hoa(t) and k.rstrip().endswith(':') and len(t) <= 40 and k.rstrip(' :') in (
                'quyet dinh', 'quyet nghi', 'ra quyet dinh', 'quyet dinh nhu sau'):
            loai = 'quyet_dinh_op'; da_den_dieu_hoac_qd = True
        elif not da_den_dieu_hoac_qd and RE_CAN_CU.match(k) and (co_loai or t.rstrip()[-1:] in ';,') and len(t) <= 300:
            loai = 'can_cu'
        elif co_loai and not da_den_dieu_hoac_qd and _la_in_hoa(t) and len(t) <= 120 and not body \
                and not RE_ROMAN.match(k) and not RE_KHAN.match(k):
            loai = 'tham_quyen'
        elif RE_ROMAN.match(k) and _la_in_hoa(t.split('.', 1)[1].strip() if '.' in t else t):
            loai = 'roman'
        elif RE_KHOAN.match(k):
            phan_sau = re.sub(r'^\d+\.\s+', '', t)
            loai = 'khoan_tieu_de' if (len(phan_sau) <= 90 and not re.search(r'[.;:,]$', phan_sau)) else 'khoan'
        elif RE_DIEM.match(k):
            loai = 'diem'
        elif RE_GACH.match(t):
            loai = 'gach'
        body.append([loai, t])
        # dòng tiêu đề đi kèm Phần/Chương/Mục/Phụ lục: dòng IN HOA kế tiếp
        if loai in ('cap_tieu_de', 'phu_luc') and i + 1 < n and _la_in_hoa(body_lines[i + 1]) \
                and len(body_lines[i + 1]) <= 120 and not RE_KHAN.match(khong_dau(body_lines[i + 1])):
            body.append(['tieu_de_in_hoa', body_lines[i + 1]])
            i += 1
        i += 1
    return body


def _doc_nhom_phu_luc(g):
    """1 phụ lục/danh sách -> dict(tieu_de=[...], dien_giai=[...], noi_dung=[[loai, text], ...])."""
    tieu_de, dien_giai = [], []
    i, n = 0, len(g)
    if g and not _la_bang(g[0]):
        tieu_de.append(g[0]); i = 1
        # tiêu đề có thể có thêm 1-2 dòng IN HOA kế tiếp (VD: "PHỤ LỤC I" + "DANH SÁCH ...")
        while i < n and len(tieu_de) < 3 and not _la_bang(g[i]) and _la_in_hoa(g[i]) and len(g[i]) <= 120 \
                and not RE_DIEN_GIAI.match(khong_dau(g[i])):
            tieu_de.append(g[i]); i += 1
    # diễn giải: "(Ban hành kèm theo ...)" — có thể ngắt thành nhiều dòng
    while i < n and not _la_bang(g[i]) and (RE_DIEN_GIAI.match(khong_dau(g[i])) or g[i].startswith('(')):
        t = g[i]; i += 1
        dem = 0
        while t.count('(') > t.count(')') and i < n and dem < 3 and not _la_bang(g[i]):
            t += ' ' + g[i]; i += 1; dem += 1
        dien_giai.append(t)
    return dict(tieu_de=tieu_de, dien_giai=dien_giai, noi_dung=_phan_than(g[i:], False))


def _tach_phu_luc(lines):
    """Phần sau chữ ký -> list các phụ lục (mỗi dòng 'Phụ lục ...' mở một phụ lục mới)."""
    nhom, cur = [], []
    for t in lines:
        if cur and _la_tieu_de_phu_luc(t):
            nhom.append(cur); cur = []
        cur.append(t)
    if cur:
        nhom.append(cur)
    return [_doc_nhom_phu_luc(g) for g in nhom]


def phan_tich(lines, bang=None):
    """Nhận diện thành phần thể thức. Trả về dict mô tả cấu trúc + danh sách nhận diện.
    bang: danh sách bảng dữ liệu (từ doc_dong_va_bang) — các dòng '[[#n]]' trong lines trỏ tới nó."""
    L = [_chuan_dong(x) for x in lines if _chuan_dong(x)]
    kq = dict(quoc_hieu=False, tieu_ngu=False, co_quan=[], so=None, ngay=None, loai=None, trich_yeu=None,
              v_v=None, kinh_gui=[], kg_items=[], body=[], noi_nhan=[], ky=dict(quyen_han=[], chuc_vu=[], ho_ten=''),
              la_cong_van=False, dinh_kem=[], bang=list(bang or []))
    if not L:
        return kq
    kd = [khong_dau(x) for x in L]
    dung = set()          # chỉ số dòng đã được gán vào khối đầu

    # --- vùng đầu: tối đa 16 dòng đầu hoặc tới khi gặp tên loại/Kính gửi/Căn cứ/Điều ---
    gioi_han = min(len(L), 16)
    for i in range(len(L)):
        if _la_loai_vb(L[i]) and not L[i].rstrip().endswith(':') or kd[i].startswith('kinh gui') \
                or RE_CAN_CU.match(kd[i]) or RE_DIEU.match(kd[i]):
            gioi_han = min(gioi_han, i + (1 if _la_loai_vb(L[i]) else 0))
            break
    for i in range(gioi_han):
        if RE_QH.search(kd[i]):
            kq['quoc_hieu'] = True; dung.add(i)
        elif RE_TN.search(kd[i]):
            kq['tieu_ngu'] = True; dung.add(i)
        elif RE_SO.match(kd[i]) and kq['so'] is None:
            kq['so'] = L[i]; dung.add(i)
        elif RE_NGAY.match(kd[i]) and kq['ngay'] is None:
            kq['ngay'] = L[i]; dung.add(i)
        elif kd[i].startswith('v/v') and kq['v_v'] is None:
            kq['v_v'] = L[i]; kq['la_cong_van'] = True; dung.add(i)
    # V/v (công văn) có thể bị ngắt thành nhiều dòng: nối các dòng tiếp theo bắt đầu bằng chữ thường
    if kq['v_v']:
        i_vv = next(i for i in range(gioi_han) if i in dung and kd[i].startswith('v/v'))
        j = i_vv + 1
        while j < gioi_han and j not in dung and L[j][:1].islower() and len(L[j]) < 160 \
                and not RE_SO.match(kd[j]) and not kd[j].startswith('kinh gui'):
            kq['v_v'] += ' ' + L[j]; dung.add(j); j += 1
        kq['v_v'] = chuan_vv(kq['v_v'])
    # tên cơ quan: dòng IN HOA trong vùng đầu chưa được gán, nằm trước tên loại
    i_loai = None
    for i in range(len(L)):
        if _la_loai_vb(L[i]) and not L[i].rstrip().endswith(':'):
            i_loai = i
            break
    gioi_cq = i_loai if i_loai is not None else gioi_han
    for i in range(min(gioi_cq, 10)):
        if i not in dung and _la_in_hoa(L[i]) and not RE_NOI_NHAN.match(kd[i]) and len(L[i]) <= 80 \
                and not _la_loai_vb(L[i]) and len(kq['co_quan']) < 2:
            kq['co_quan'].append(L[i]); dung.add(i)
    pos = (max(dung) + 1) if dung else 0
    if i_loai is not None and i_loai >= pos - 0:
        kq['loai'] = L[i_loai].rstrip(' :')
        pos = i_loai + 1
        if pos < len(L) and re.match(r'^(ve viec|v/v|ve\b)', kd[pos]):
            kq['trich_yeu'] = L[pos]
            pos += 1
    # Kính gửi: 1 nơi ("Kính gửi: A") hoặc nhiều nơi (dòng "Kính gửi:" + danh sách gạch đầu dòng / kết thúc ; ,)
    if pos < len(L) and kd[pos].startswith('kinh gui'):
        dong_kg = L[pos]; pos += 1
        kq['kinh_gui'].append(dong_kg)
        phan_dau = re.sub(r'^\s*k[ií]nh\s+g[ửu]i\s*[:：]?\s*', '', dong_kg, flags=re.I).strip()
        items = [phan_dau] if phan_dau else []
        while pos < len(L) and len(L[pos]) <= 200:
            t_ = L[pos]
            if RE_GACH.match(t_) or (items and items[-1].rstrip()[-1:] in ';,') \
                    or (not items and len(t_) <= 100 and t_.rstrip()[-1:] not in '.:'
                        and not RE_CAN_CU.match(khong_dau(t_)) and not khong_dau(t_).startswith(('ngay ', 'thuc hien', 'qua '))):
                items.append(RE_GACH.sub('', t_).strip()); kq['kinh_gui'].append(t_); pos += 1
            else:
                break
        kq['kg_items'] = [x for x in items if x]
        kq['la_cong_van'] = kq['la_cong_van'] or kq['loai'] is None

    rest = L[pos:]
    # --- Nơi nhận + chữ ký + PHẦN ĐÍNH KÈM (Phụ lục/Danh sách...) ---
    krest = [khong_dau(x) for x in rest]
    i_nn = next((i for i, k in enumerate(krest) if RE_NOI_NHAN.match(k)), None)
    body_lines, sig_lines, dk_lines = rest, [], []
    if i_nn is not None:
        body_lines = rest[:i_nn]
        tail = rest[i_nn + 1:]
        j = 0
        while j < len(tail) and (RE_GACH.match(tail[j]) or khong_dau(tail[j]).startswith('luu')):
            kq['noi_nhan'].append(RE_GACH.sub('', tail[j]).strip()); j += 1
        sig_lines = tail[j:]
        c = _vi_tri_dinh_kem(sig_lines)                  # mọi thứ từ đây trở đi là phụ lục/danh sách đính kèm
        if c is not None:
            dk_lines, sig_lines = sig_lines[c:], sig_lines[:c]
    else:
        c = _vi_tri_dinh_kem(body_lines, chat=True, tu=1)
        if c is not None:
            dk_lines, body_lines = body_lines[c:], body_lines[:c]
    # chữ ký không bao giờ nằm trước 1 bảng dữ liệu -> chặn dò ngược ở dòng bảng cuối cùng
    san = max([i for i, x in enumerate(body_lines) if _la_bang(x)] + [-1]) + 1
    if i_nn is not None:
        if not sig_lines:          # chữ ký nằm TRƯỚC "Nơi nhận" (cuối thân văn bản)
            k = len(body_lines)
            while k > san and (len(body_lines) - k) < 6 and (
                    (len(body_lines[k - 1]) <= 70 and _la_dong_ky(body_lines[k - 1]))
                    or (len(body_lines[k - 1]) <= 40 and not re.search(r'[.;:,]$', body_lines[k - 1]))):
                k -= 1
            blk = body_lines[k:]
            if blk and any(_la_in_hoa(x) or RE_QUYEN_HAN.match(khong_dau(x)) for x in blk):
                sig_lines, body_lines = blk, body_lines[:k]
    else:
        k = len(body_lines)
        while k > san and (len(body_lines) - k) < 6 and len(body_lines[k - 1]) <= 70 \
                and not re.search(r'[.;:,]$', body_lines[k - 1]) and not RE_GACH.match(body_lines[k - 1]):
            k -= 1
        blk = body_lines[k:]
        if blk and any(_la_in_hoa(x) for x in blk) and len(blk) <= 5:
            sig_lines, body_lines = blk, body_lines[:k]

    # phân tách chữ ký — KHÔNG lấy họ tên nếu nghi là chữ ký tay (tên bị OCR méo)
    ky = kq['ky']
    if sig_lines:
        sl = list(sig_lines)
        ung_vien_ten = None
        if len(sl) >= 2 and not (_la_in_hoa(sl[-1]) and
                                 any(khong_dau(sl[-1]).startswith(t) for t in CHUC_VU_TU_KHOA)):
            ung_vien_ten = sl.pop()
        elif len(sl) == 1 and not _la_in_hoa(sl[0]):
            ung_vien_ten = sl.pop()

        # Chỉ nhận họ tên khi nó "sạch" (toàn chữ cái, có dấu hợp lệ, 2-5 từ)
        if ung_vien_ten:
            tu = ung_vien_ten.split()
            la_ten_sach = (
                2 <= len(tu) <= 6
                and all(re.fullmatch(r"[A-Za-zÀ-ỹ]+", w) for w in tu)
            )
            if la_ten_sach:
                ky['ho_ten'] = ung_vien_ten.title() if _la_in_hoa(ung_vien_ten) else ung_vien_ten
        for x in sl:
            (ky['quyen_han'] if RE_QUYEN_HAN.match(khong_dau(x)) else ky['chuc_vu']).append(x.upper())

    kq['body'] = _phan_than(body_lines, kq['loai'] is not None)
    kq['dinh_kem'] = _tach_phu_luc(dk_lines) if dk_lines else []
    return kq


# ───────────────────────── 3. Dựng Word ─────────────────────────


def _font_run(r, size, bold=None, italic=None):
    r.font.name = FONT
    rf = r._element.get_or_add_rPr().find(qn('w:rFonts'))
    if rf is not None:
        for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            rf.set(qn(a), FONT)
    r.font.size = Pt(size)
    if bold is not None:
        r.bold = bold
    if italic is not None:
        r.italic = italic
    return r


def _p(container, text='', size=14, bold=False, italic=False, align='justify', indent=True,
       before=6, after=6, p=None, upper=False):
    p = p if p is not None else container.add_paragraph()
    p.alignment = _CAN[align]
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.first_line_indent = Cm(1.27) if indent else Cm(0)
    if text:
        _font_run(p.add_run(text.upper() if upper else text), size, bold, italic)
    return p


def _ke_ngang(container, rong_cm, ben_cm):
    p = container.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr'); b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '6'); b.set(qn('w:space'), '1'); b.set(qn('w:color'), '000000')
    bdr.append(b); pPr.append(bdr)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ind = max((ben_cm - rong_cm) / 2.0, 0)
    pf = p.paragraph_format
    pf.left_indent = Cm(ind); pf.right_indent = Cm(ind)
    pf.space_before = Pt(0); pf.space_after = Pt(0)
    pf.line_spacing = Pt(5)
    _font_run(p.add_run(''), 1)
    rpr = OxmlElement('w:rPr'); sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '2')
    rpr.append(sz); pPr.append(rpr)
    return p


def _khung(p):
    """Khung viền 4 cạnh quanh đoạn (dấu mức độ khẩn: HỎA TỐC / THƯỢNG KHẨN / KHẨN)."""
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    for canh in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement(f'w:{canh}')
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '8'); e.set(qn('w:space'), '3'); e.set(qn('w:color'), '000000')
        bdr.append(e)
    pPr.insert(0, bdr)


def _bo_vien_bang(tbl):
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            b = tcPr.find(qn('w:tcBorders'))
            if b is not None:
                tcPr.remove(b)

def _ghi_o(cell, text, size, la_tieu_de=False):
    """Ghi chữ vào 1 ô bảng: hàng tiêu đề căn giữa + đậm; ô số/STT căn giữa; còn lại căn trái."""
    for idx, x in enumerate(text.split('\n') if text else ['']):
        so = len(x) <= 12 and bool(re.fullmatch(r'[\d\s.,/%()+\-–:]*', x))
        _p(cell, x, size, bold=la_tieu_de, align='center' if (la_tieu_de or so) else 'left', indent=False,
           before=2, after=2, p=cell.paragraphs[0] if idx == 0 else None)


def _dung_bang(doc, rows, size):
    """Dựng lại bảng dữ liệu dạng lưới (có viền, gộp ô ngang/dọc nếu có).
    rows = [[dict(t, span, vm), ...], ...] — lấy từ _doc_luoi_bang()."""
    rows = [h for h in rows if h]
    if not rows:
        return
    n_cot = max(1, max(sum(c['span'] for c in h) for h in rows))
    # độ rộng cột theo độ dài chữ, tổng = 16 cm (bề rộng vùng in)
    trong = [3] * n_cot
    for h in rows:
        c0 = 0
        for c in h:
            if c['span'] == 1 and c0 < n_cot:
                dai = max((len(x) for x in c['t'].split('\n')), default=0)
                trong[c0] = max(trong[c0], min(dai, 40))
            c0 += c['span']
    rong = [max(1.2, 16.0 * x / float(sum(trong))) for x in trong]
    k = 16.0 / sum(rong)
    rong = [x * k for x in rong]

    tb = doc.add_table(rows=len(rows), cols=n_cot)
    tb.style = 'Table Grid'
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = False
    for ci, w in enumerate(rong):
        tb.columns[ci].width = Cm(w)
    for row in tb.rows:
        for ci, cell in enumerate(row.cells):
            cell.width = Cm(rong[ci])

    # hàng đầu là tiêu đề nếu không có ô nào là số thuần (STT, số tiền...)
    co_tieu_de = len(rows) > 1 and not any(re.fullmatch(r'\d+([.,]\d+)*', c['t'].strip()) for c in rows[0] if c['t'])
    goc = {}                                   # cột -> hàng bắt đầu của ô gộp dọc
    for ri, h in enumerate(rows):
        c0 = 0
        for c in h:
            if c0 >= n_cot:
                break
            c1 = min(c0 + c['span'], n_cot) - 1
            if c['vm'] == 'cont' and c0 in goc:
                try:
                    tb.cell(goc[c0], c0).merge(tb.cell(ri, c1))
                except Exception:
                    pass
                c0 += c['span']
                continue
            if c['vm'] == 'restart':
                goc[c0] = ri
            else:
                goc.pop(c0, None)
            cell = tb.cell(ri, c0)
            if c1 > c0:
                try:
                    cell = cell.merge(tb.cell(ri, c1))
                except Exception:
                    pass
            _ghi_o(cell, c['t'], size, la_tieu_de=(co_tieu_de and ri == 0))
            c0 += c['span']
    if co_tieu_de:                             # lặp lại hàng tiêu đề khi bảng tràn sang trang sau
        trPr = tb.rows[0]._tr.get_or_add_trPr()
        th = OxmlElement('w:tblHeader'); th.set(qn('w:val'), 'true'); trPr.append(th)
    _p(doc, '', 4, before=0, after=0)          # đoạn đệm nhỏ để 2 bảng liền nhau không bị Word dính làm một

def _so_trang(sec, size):
    sec.different_first_page_header_footer = True
    sec.header_distance = Cm(1.0)
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hp.add_run(); _font_run(r, size, italic=False)
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = 'PAGE'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
    r._r.append(f1); r._r.append(it); r._r.append(f2)


def _thiet_lap_trang(doc, nd):
    st_ = doc.styles['Normal']
    st_.font.name = FONT; st_.font.size = Pt(nd)
    rf = st_.element.get_or_add_rPr().find(qn('w:rFonts'))
    if rf is not None:
        for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            rf.set(qn(a), FONT)
    pf = st_.paragraph_format
    pf.space_before = Pt(6); pf.space_after = Pt(6)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sec = doc.sections[0]
    sec.page_width = Cm(21); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2); sec.bottom_margin = Cm(2); sec.left_margin = Cm(3); sec.right_margin = Cm(2)
    _so_trang(sec, nd)


def _khoi_dau(doc, S, co_quan, so, ngay, v_v):
    # Quốc hiệu cỡ 13 cần cột phải rộng ~10,4cm; cỡ 12 thì ~9,6cm -> cột trái (Số, V/v) rộng thêm khi có thể
    L, R = (5.6, 10.4) if S['qh'] >= 13 else (6.4, 9.6)
    ht = doc.add_table(rows=1, cols=2)
    ht.alignment = WD_TABLE_ALIGNMENT.CENTER; ht.autofit = False
    _bo_vien_bang(ht)
    cl, cr = ht.rows[0].cells
    cl.width = Cm(L); cr.width = Cm(R)
    ht.columns[0].width = Cm(L); ht.columns[1].width = Cm(R)
    # trái: (cơ quan chủ quản đứng) + cơ quan ban hành đứng đậm + kẻ + số + V/v
    for idx, ten in enumerate(co_quan):
        _p(cl, ten.upper(), S['cq'], bold=True, italic=False, align='center', indent=False, before=0, after=0,
           p=cl.paragraphs[0] if idx == 0 else None)
    _ke_ngang(cl, 3.0, L)
    _p(cl, so, S['so'], align='center', indent=False, before=6, after=0)
    if v_v:
        # Khung V/v: cả khối căn giữa dưới "Số" (lề trái = lề phải), chữ trong khối căn đều 2 bên, không thụt đầu dòng
        pv = _p(cl, v_v, S['vv'], align='justify', indent=False, before=4, after=0)
        pv.paragraph_format.left_indent = Cm(0.15); pv.paragraph_format.right_indent = Cm(0.15)
    _p(cr, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', S['qh'], True, align='center', indent=False, before=0, after=0,
       p=cr.paragraphs[0])
    _p(cr, 'Độc lập - Tự do - Hạnh phúc', S['tn'], True, align='center', indent=False, before=0, after=0)
    _ke_ngang(cr, 5.6, R)
    _p(cr, ngay, S['ngay'], italic=True, align='right', indent=False, before=6, after=0)


def _khoi_kinh_gui(doc, S, items):
    """Khối "Kính gửi": 2 cột (nhãn | danh sách nơi nhận), căn trái, KHÔNG căn giữa.
    1 nơi: viết liền; nhiều nơi: mỗi nơi 1 dòng gạch đầu dòng."""
    _p(doc, '', 6, before=6, after=0)
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    _bo_vien_bang(t)
    c1, c2 = t.rows[0].cells
    W1, W2 = 4.0, 12.0
    c1.width = Cm(W1); c2.width = Cm(W2)
    t.columns[0].width = Cm(W1); t.columns[1].width = Cm(W2)
    pl = _p(c1, 'Kính gửi:', S['nd'], align='left', indent=False, before=0, after=0, p=c1.paragraphs[0])
    pl.paragraph_format.left_indent = Cm(1.27)
    multi = len(items) > 1
    for idx, x in enumerate(items):
        _p(c2, ('- ' + x) if multi else x, S['nd'], align='left', indent=False, before=0, after=0,
           p=c2.paragraphs[0] if idx == 0 else None)
    _p(doc, '', 6, before=0, after=6)


def _khoi_cuoi(doc, S, noi_nhan, ky):
    _p(doc, '', 6, before=0, after=6)
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
    _bo_vien_bang(t)
    cl, cr = t.rows[0].cells
    cl.width = Cm(5.5); cr.width = Cm(10.5)
    t.columns[0].width = Cm(5.5); t.columns[1].width = Cm(10.5)
    _p(cl, 'Nơi nhận:', S['nn'], True, True, align='left', indent=False, before=0, after=0, p=cl.paragraphs[0])
    for d in noi_nhan:
        _p(cl, '- ' + d.lstrip('-– ').strip(), S['nn_ds'], align='left', indent=False, before=0, after=0)
    dong = list(ky.get('quyen_han', [])) + list(ky.get('chuc_vu', []))
    for idx, x in enumerate(dong):
        _p(cr, x, S['ky'], True, align='center', indent=False, before=0, after=0, p=cr.paragraphs[0] if idx == 0 else None)
    _p(cr, ky.get('ho_ten', ''), S['ky'], True, align='center', indent=False, before=72 if dong else 0, after=0,
       p=cr.paragraphs[0] if not dong else None)


def _dung_than(doc, S, body, bang=None):
    nd = S['nd']
    n = len(body)
    for i, (loai, t) in enumerate(body):
        tiep_la_can_cu = i + 1 < n and body[i + 1][0] == 'can_cu'
        if loai == 'bang':
            m = RE_BANG.match(t)
            if m and bang and int(m.group(1)) < len(bang):
                _dung_bang(doc, bang[int(m.group(1))], max(nd - 1, 12))
            continue
        if loai == 'khan':
            p = _p(doc, t.upper(), S['ky'], True, align='right', indent=False, before=6, after=6)
            _khung(p)
        elif loai == 'tham_quyen':
            _p(doc, t, S['ky'], True, align='center', indent=False, before=12, after=6, upper=True)
        elif loai == 'can_cu':
            _p(doc, t, nd, italic=True, before=0, after=0 if tiep_la_can_cu else 6)
        elif loai == 'quyet_dinh_op':
            _p(doc, t.upper(), nd, True, align='center', indent=False, before=12, after=6)
        elif loai == 'cap_tieu_de':
            _p(doc, t, nd, True, align='center', indent=False, before=12, after=0)
        elif loai == 'phu_luc':
            _p(doc, t, nd, True, align='center', indent=False, before=12, after=0)
        elif loai == 'tieu_de_in_hoa':
            _p(doc, t, nd, True, align='center', indent=False, before=0, after=6, upper=True)
        elif loai == 'roman':
            _p(doc, t.upper(), nd, True, align='left', indent=False, before=6, after=6)
        elif loai == 'dieu':
            m = re.match(r'^(Điều\s+\d+\s*[\.:])\s*(.*)$', t, re.I | re.S)
            if m and m.group(2) and (len(m.group(2)) > 90 or re.search(r'[.;:,]$', m.group(2))):
                p = _p(doc, '', nd)
                _font_run(p.add_run(m.group(1) + ' '), nd, True)
                _font_run(p.add_run(m.group(2)), nd, False)
            else:
                _p(doc, t, nd, True)
        elif loai == 'khoan_tieu_de':
            _p(doc, t, nd, True)
        elif loai in ('khoan', 'diem', 'gach', 'text'):
            _p(doc, t, nd)

def _dung_dinh_kem(doc, S, dinh_kem, bang):
    """Phần đính kèm sau chữ ký (Phụ lục / Danh sách / Bảng kê...): mỗi mục sang TRANG MỚI;
    tiêu đề căn giữa, đậm, đứng; dòng diễn giải (Ban hành kèm theo...) căn giữa, nghiêng; bảng dựng lại lưới."""
    nd = S['nd']
    for pl in dinh_kem:
        ds = [(t, True, False) for t in pl['tieu_de']] + [(t, False, True) for t in pl['dien_giai']]
        if not ds:
            ds = [('', False, False)]
        for idx, (t, dam, nghieng) in enumerate(ds):
            p = _p(doc, t, nd, bold=dam, italic=nghieng, align='center', indent=False,
                   before=0, after=6 if idx == len(ds) - 1 else 0)
            if idx == 0:
                p.paragraph_format.page_break_before = True
        _dung_than(doc, S, pl['noi_dung'], bang)

def uoc_luong_chieu_cao(cau_truc, co=14):
    """Ước tính chiều cao (cm) nội dung khi dàn theo bộ cỡ chữ `co` (dùng khi không có LibreOffice để đếm trang thật).
    Phần đính kèm (Phụ lục...) luôn sang trang mới nên được cộng theo số trang riêng."""
    cpl = 76 if co == 14 else 82
    dong_cm = 0.57 if co == 14 else 0.53
    cach_doan = 0.42
    bang = cau_truc.get('bang') or []

    def _cao(ds):
        h_ = 0.0
        for loai, t in ds:
            m = RE_BANG.match(t) if loai == 'bang' else None
            if m:
                h_ += (len(bang[int(m.group(1))]) * 0.9) if int(m.group(1)) < len(bang) else 0
            else:
                h_ += max(1, math.ceil(len(t) / cpl)) * dong_cm + cach_doan
        return h_

    h = 5.0                                # khối đầu
    if cau_truc['loai']:
        h += 1.4
    if cau_truc['trich_yeu']:
        h += 0.9
    h += _cao(cau_truc['body'])
    h += 6.7                               # nơi nhận + chữ ký (đã tính khoảng ký cao thêm 20%)
    cm = math.ceil(h / 25.7) * 25.7
    for pl in cau_truc.get('dinh_kem', []):
        hp = 2.5 + _cao(pl['noi_dung'])
        cm += math.ceil(hp / 25.7) * 25.7
    return cm


def _dem_trang_bang_libreoffice(docx_bytes):
    exe = shutil.which('soffice') or shutil.which('libreoffice')
    if not exe:
        return None
    try:
        with tempfile.TemporaryDirectory() as td:
            src = f"{td}/vb.docx"
            open(src, 'wb').write(docx_bytes)
            subprocess.run([exe, '--headless', '--convert-to', 'pdf', '--outdir', td, src],
                           capture_output=True, timeout=90)
            pdf = open(f"{td}/vb.pdf", 'rb').read()
            return len(re.findall(rb'/Type\s*/Page[^s]', pdf)) or None
    except Exception:
        return None


def dung_docx(cau_truc, co, cfg=None, tu_bo_sung=True):
    """Dựng file Word từ cấu trúc nhận diện. Trả về (bytes, danh_sach_bo_sung)."""
    cfg = cfg or {}
    S = CO_CHU[co]
    bo_sung = []
    doc = Document()
    _thiet_lap_trang(doc, S['nd'])

    co_quan = list(cau_truc['co_quan'])
    if not co_quan:
        ten_ty = (cfg.get('ten_cong_ty') or '').strip()
        if tu_bo_sung and ten_ty:
            low = ten_ty.lower()
            tach = None
            for kw, nhan in (("công ty cổ phần", "CÔNG TY CỔ PHẦN"), ("công ty tnhh", "CÔNG TY TNHH"),
                             ("hộ kinh doanh", "HỘ KINH DOANH")):
                if kw in low:
                    tach = [nhan, ten_ty[low.index(kw) + len(kw):].strip().upper()]
                    break
            co_quan = [x for x in (tach or [ten_ty.upper()]) if x]
            bo_sung.append("Tên cơ quan (lấy từ cấu hình công ty)")
    so = cau_truc['so'] or ('Số: ....../......' if tu_bo_sung else 'Số:')
    if not cau_truc['so']:
        bo_sung.append("Số, ký hiệu (để trống điền tay)")
    ngay = cau_truc['ngay']
    if not ngay:
        dd = cfg.get('dia_diem') or ''
        ngay = f"{dd + ', ' if dd else ''}ngày ... tháng ... năm ......"
        bo_sung.append("Địa danh, ngày tháng (để trống điền tay)")
    if not cau_truc['quoc_hieu']:
        bo_sung.append("Quốc hiệu")
    if not cau_truc['tieu_ngu']:
        bo_sung.append("Tiêu ngữ")
    _khoi_dau(doc, S, co_quan, so, ngay, cau_truc['v_v'] if cau_truc['la_cong_van'] else None)

    if cau_truc['loai']:
        _p(doc, cau_truc['loai'].upper(), S['loai'], True, align='center', indent=False, before=18, after=0)
        if cau_truc['trich_yeu']:
            _p(doc, cau_truc['trich_yeu'], S['ty'], True, align='center', indent=False, before=0, after=0)
            _ke_ngang(doc, 4.0, 16.0)
    if cau_truc.get('kg_items'):
        _khoi_kinh_gui(doc, S, cau_truc['kg_items'])
    elif cau_truc['kinh_gui']:
        _khoi_kinh_gui(doc, S, [re.sub(r'^\s*k[ií]nh\s+g[ửu]i\s*[:：]?\s*', '', cau_truc['kinh_gui'][0], flags=re.I) or '...'])
    _dung_than(doc, S, cau_truc['body'], cau_truc.get('bang'))

    noi_nhan = list(cau_truc['noi_nhan'])
    if not noi_nhan and tu_bo_sung:
        noi_nhan = ['Như trên;' if cau_truc['la_cong_van'] else 'Lưu: VT.']
        if cau_truc['la_cong_van']:
            noi_nhan.append('Lưu: VT.')
        bo_sung.append("Nơi nhận (mặc định)")
    ky = dict(cau_truc['ky'])
    if not (ky.get('chuc_vu') or ky.get('quyen_han') or ky.get('ho_ten')) and tu_bo_sung:
        ky['chuc_vu'] = [(cfg.get('chuc_vu') or 'GIÁM ĐỐC').upper()]
        bo_sung.append("Chức vụ người ký (lấy từ cấu hình công ty; họ tên để trống ký tay)")
    _khoi_cuoi(doc, S, noi_nhan, ky)
    _dung_dinh_kem(doc, S, cau_truc.get('dinh_kem', []), cau_truc.get('bang'))
    bio = io.BytesIO(); doc.save(bio)
    return bio.getvalue(), bo_sung



# ───────────────────────── Rà soát chính tả & câu ─────────────────────────
LOI_CHINH_TA = {
    # lỗi gõ/chính tả thường gặp trong văn bản hành chính (chỉ liệt kê các cặp sai-đúng chắc chắn)
    'sử lý': 'xử lý', 'xữ lý': 'xử lý', 'xữ lí': 'xử lí', 'sữ dụng': 'sử dụng', 'chỉnh sữa': 'chỉnh sửa',
    'sữa chữa': 'sửa chữa', 'kiễm tra': 'kiểm tra', 'quyết đinh': 'quyết định', 'trân thành': 'chân thành',
    'chân trọng cảm ơn': 'trân trọng cảm ơn', 'giám sác': 'giám sát', 'tăng cuờng': 'tăng cường',
    'tăng cưỡng': 'tăng cường', 'đẩy mạng': 'đẩy mạnh', 'chủ chương': 'chủ trương', 'triển khải': 'triển khai',
    'thực hiên': 'thực hiện', 'dưt điểm': 'dứt điểm', 'dứt điễm': 'dứt điểm', 'xin chân thành cảm ơn': 'xin chân thành cảm ơn',
    'rà xoát': 'rà soát', 'dà soát': 'rà soát', 'giãi quyết': 'giải quyết', 'giãi trình': 'giải trình',
    'phối hơp': 'phối hợp', 'hổ trợ': 'hỗ trợ', 'hõ trợ': 'hỗ trợ', 'chử ký': 'chữ ký', 'nghĩ việc': 'nghỉ việc',
    'nghĩ phép': 'nghỉ phép', 'xét duyêt': 'xét duyệt', 'cấp bách bách': 'cấp bách',
}
_LOI_CT_THAT = {k: v for k, v in LOI_CHINH_TA.items() if k != v}
TU_VIET_TAT_CHO_PHEP = {'tp', 'ts', 'gs', 'pgs', 'ths', 'bs', 'ks', 'cn', 'đc', 'v.v', 'tt', 'qđ', 'nđ', 'tb', 'stt'}


def sua_dau_cau(t):
    """Sửa an toàn khoảng trắng/dấu câu: bỏ khoảng trắng trước , ; : ! ? và thêm khoảng trắng sau , ; khi dính chữ.
    Không động tới số thập phân (1,5), giờ (10:30), dấu ba chấm, đường dẫn. Trả về (văn bản mới, số chỗ đã sửa)."""
    n = 0
    t2, k = re.subn(r'(?<=\S)\s+([,;:!?])(?!\d)', r'\1', t); n += k
    t3, k = re.subn(r'([,;])(?=[A-Za-zÀ-ỹ])', r'\1 ', t2); n += k
    return t3, n


def ra_soat_ngoai_tuyen(lines, ct=None):
    """Rà soát bằng quy tắc (không cần mạng): trả về danh sách lỗi dict(muc, trich, van_de, de_xuat, nguon)."""
    from datetime import datetime as _dt
    loi = []

    def them(muc, trich, van_de, de_xuat=''):
        loi.append(dict(muc=muc, trich=(trich or '')[:140], van_de=van_de, de_xuat=de_xuat, nguon='Quy tắc'))

    for t in lines:
        tl = t.lower()
        # a) lỗi chính tả thường gặp / lặp từ
        for sai, dung_ in _LOI_CT_THAT.items():
            if sai in tl:
                them('Chính tả', t, f"Có thể sai: “{sai}”", f"Nên viết: “{dung_}”")
        for m in re.finditer(r'(?<![\w/])([^\W\d_]{2,})\s+\1(?![\w/])', t, flags=re.I):
            them('Chính tả', t, f"Lặp từ: “{m.group(0)}”", f"Nên bỏ bớt: “{m.group(1)}”")
        # b) dấu câu / ngoặc
        if re.search(r'([,;:])\1', t) or re.search(r'(?<!\.)\.\.(?!\.)', t):
            them('Dấu câu', t, "Dấu câu bị lặp (,, ;; :: hoặc '..')", "Kiểm tra lại dấu câu")
        if t.count('(') != t.count(')'):
            them('Dấu câu', t, "Ngoặc đơn không khớp cặp “(” “)”", "Bổ sung ngoặc còn thiếu")
        if t.count('“') != t.count('”') or t.count('"') % 2 == 1:
            them('Dấu câu', t, "Dấu ngoặc kép không khớp cặp", "Bổ sung dấu ngoặc kép còn thiếu")
        # c) viết hoa đầu câu
        for m in re.finditer(r'([^\W\d_]{3,})([.!?])\s+([a-zà-ỹ])', t):
            if m.group(1).lower() not in TU_VIET_TAT_CHO_PHEP:
                them('Chính tả', t, f"Đầu câu chưa viết hoa sau dấu “{m.group(2)}”: “…{m.group(1)}{m.group(2)} {m.group(3)}…”",
                     "Viết hoa chữ cái đầu câu")
        # d) ngày tháng không tồn tại
        for m in re.finditer(r'(?<!\d)(\d{1,2})/(\d{1,2})/(\d{4})(?!\d)', t):
            d_, mo, y = map(int, m.groups())
            try:
                _dt(y, mo, d_)
            except ValueError:
                them('Logic', t, f"Ngày không tồn tại: {m.group(0)}", "Kiểm tra lại ngày/tháng/năm")
        for m in re.finditer(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})', t, flags=re.I):
            d_, mo, y = map(int, m.groups())
            try:
                _dt(y, mo, d_)
            except ValueError:
                them('Logic', t, f"Ngày không tồn tại: {m.group(0)}", "Kiểm tra lại ngày/tháng/năm")
        # e) khoảng thời gian đảo ngược
        for m in re.finditer(r'từ\s+(?:ngày\s+)?(\d{1,2})/(\d{1,2})/(\d{4})\s+(?:đến|tới)\s+(?:ngày\s+)?(\d{1,2})/(\d{1,2})/(\d{4})', t, flags=re.I):
            try:
                a_ = _dt(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                b_ = _dt(int(m.group(6)), int(m.group(5)), int(m.group(4)))
                if b_ < a_:
                    them('Logic', t, f"Khoảng thời gian bị ngược: “{m.group(0)}” (ngày kết thúc trước ngày bắt đầu)",
                         "Đổi lại thứ tự hoặc sửa ngày")
            except ValueError:
                pass
    # f) đoạn dài không kết thúc bằng dấu câu (chỉ xét thân văn bản)
    if ct:
        for loai, t in ct.get('body', []):
            if loai in ('text', 'khoan', 'diem', 'gach') and len(t) > 80 and t.rstrip()[-1:] not in '.;:!?…)”"':
                them('Câu', t, "Đoạn/câu dài không kết thúc bằng dấu câu", "Thêm dấu chấm (.) hoặc dấu chấm phẩy (;)")
    # loại trùng
    seen, kq = set(), []
    for x in loi:
        k = (x['muc'], x['van_de'], x['trich'][:60])
        if k not in seen:
            seen.add(k); kq.append(x)
    return kq


def ra_soat_ai(lines, ai_call):
    """Rà soát chính tả, ngữ pháp, câu bất hợp lý bằng AI (dùng hàm gọi Claude sẵn có của app).
    Trả về (danh_sach_loi, thong_bao_loi). AI chỉ ĐỀ XUẤT — không tự sửa nội dung văn bản."""
    import json as _json
    if ai_call is None:
        return [], "Chưa kết nối dịch vụ AI."
    van_ban = '\n'.join(f"[{i + 1}] {t}" for i, t in enumerate(lines))
    cat = len(van_ban) > 14000
    van_ban = van_ban[:14000]
    system = (
        "Bạn là biên tập viên văn bản hành chính tiếng Việt, rất cẩn trọng. Hãy rà soát văn bản người dùng gửi và CHỈ báo cáo các lỗi THẬT SỰ: "
        "(1) lỗi chính tả, dùng sai dấu, sai từ; (2) lỗi ngữ pháp, câu thiếu chủ ngữ/vị ngữ, câu cụt, câu lặp ý; "
        "(3) câu hoặc nội dung bất hợp lý, mâu thuẫn (ngày tháng, số liệu, tên cơ quan/người không nhất quán, 'đề nghị' nhưng không nêu đối tượng...). "
        "KHÔNG nhận xét về văn phong, KHÔNG báo lỗi nếu không chắc, KHÔNG viết lại văn bản, KHÔNG bịa. "
        "Trả lời DUY NHẤT bằng JSON hợp lệ, không kèm giải thích hay markdown, đúng dạng: "
        '{"loi":[{"muc":"Chính tả|Ngữ pháp|Logic","dong":<số thứ tự dòng trong ngoặc vuông>,'
        '"trich":"<trích nguyên văn đoạn có lỗi, tối đa 120 ký tự>","van_de":"<mô tả ngắn lỗi>","de_xuat":"<cách sửa đề xuất>"}]}. '
        'Nếu không có lỗi trả về {"loi":[]}.')
    try:
        kq = ai_call(system, [{"role": "user", "content": van_ban}])
    except Exception as e:
        return [], f"Không gọi được AI: {e}"
    if not isinstance(kq, str):
        return [], "AI trả về dữ liệu không hợp lệ."
    if kq.strip().startswith(('⚠️', '❌')):
        return [], re.sub(r'[*`]', '', kq).strip()
    txt = re.sub(r'^```(?:json)?|```$', '', kq.strip(), flags=re.M).strip()
    i, j = txt.find('{'), txt.rfind('}')
    try:
        data = _json.loads(txt[i:j + 1])
    except Exception:
        return [], "AI trả về kết quả không đọc được (không phải JSON)."
    out = []
    for x in data.get('loi', []) if isinstance(data, dict) else []:
        if isinstance(x, dict) and x.get('van_de'):
            out.append(dict(muc=str(x.get('muc') or 'Khác'), trich=str(x.get('trich') or '')[:140],
                            van_de=str(x.get('van_de')), de_xuat=str(x.get('de_xuat') or ''), nguon='AI'))
    note = "Văn bản dài nên AI chỉ rà soát phần đầu (~14.000 ký tự)." if cat else ""
    return out, note


# ───────────────────────── PDF -> Word bằng AI (BƯỚC 1) ─────────────────────────
LENH_PDF_SANG_WORD = (
    "Hãy trích xuất toàn bộ nội dung từ file PDF tôi đính kèm và tạo thành một file Microsoft Word (.docx).\n"
    "Yêu cầu định dạng file Word nghiêm ngặt theo chuẩn Nghị định 30/2020/NĐ-CP như sau:\n"
    "Font chữ: Times New Roman.\n"
    "Cỡ chữ: 14 cho phần nội dung (các tiêu đề có thể tùy chỉnh to/nhỏ hoặc in đậm/in nghiêng tương ứng theo cấu trúc văn bản gốc).\n"
    "Căn lề trang giấy (Page Setup - khổ A4): Lề trên 2 cm; Lề dưới 2 cm; Lề trái 3 cm; Lề phải 2 cm.\n"
    "Định dạng đoạn văn (Paragraph):\n"
    "Thụt đầu dòng (First line indent): 1.27 cm.\n"
    "Khoảng cách giữa các đoạn (Spacing Before/After): 6pt.\n"
    "Căn lề văn bản (Alignment): Căn đều hai bên (Justify).\n"
    "Bố cục: Giữ nguyên tối đa cấu trúc phân chương, phân mục, các đoạn văn, danh sách (bullet/numbering) và vị trí của "
    "các tiêu đề (căn giữa, căn phải) giống như văn bản gốc.\n"
    "Đối với bảng biểu, hãy cố gắng tái tạo lại cấu trúc lưới (grid) và điền dữ liệu tương ứng.")

QUY_UOC_TRA_KET_QUA = (
    "\n\nQUY ƯỚC TRẢ KẾT QUẢ (bắt buộc — vì app sẽ tự dựng file Word đúng các yêu cầu định dạng ở trên từ nội dung bạn trả):\n"
    "- Mỗi đoạn văn / mỗi dòng của văn bản gốc = 1 dòng. Giữ đúng thứ tự đọc (trên xuống dưới, trái sang phải), kể cả khối đầu "
    "(quốc hiệu, tiêu ngữ, tên cơ quan, số, địa danh ngày tháng), nơi nhận, chức vụ và họ tên người ký, các phụ lục/danh sách đính kèm.\n"
    "- Có thể thêm tiền tố ở đầu dòng để giữ vị trí/kiểu chữ gốc: [C] căn giữa, [R] căn phải, [L] căn trái, [B] in đậm, "
    "[I] in nghiêng (ghép được, VD [C][B]). Không có tiền tố = căn đều hai bên.\n"
    "- Bảng biểu: mở đầu bằng dòng [[BANG]], mỗi hàng của bảng 1 dòng, các ô cách nhau bằng dấu |, ô trống để trống, "
    "ô gộp thì ghi nội dung ở ô đầu rồi để trống các ô còn lại; kết thúc bằng dòng [[/BANG]]. Số ô các hàng phải bằng nhau.\n"
    "- Bỏ qua con dấu, chữ ký tay, số trang, tiêu đề/chân trang lặp lại. Chỉ ghi họ tên người ký nếu được đánh máy rõ ràng.\n"
    "- Không thêm lời dẫn, chú thích hay markdown nào ngoài nội dung văn bản.")


def _parse_khoi_ai(txt):
    """Kết quả AI -> list khối: ('p', (text, tags)) | ('bang', [[ô, ô, ...], ...])."""
    khoi, cur = [], None
    txt = re.sub(r'^```[a-z]*\s*$', '', txt or '', flags=re.M)
    for raw in txt.splitlines():
        l = raw.strip()
        if re.fullmatch(r'\[\[BANG\]\]', l, flags=re.I):
            cur = []
            continue
        if re.fullmatch(r'\[\[/BANG\]\]', l, flags=re.I):
            if cur:
                khoi.append(('bang', cur))
            cur = None
            continue
        if cur is not None:
            if not l:
                continue
            if l.startswith('|'):
                l = l[1:]
            if l.endswith('|'):
                l = l[:-1]
            cells = [_chuan_dong(c) for c in l.split('|')]
            if any(cells) and all(re.fullmatch(r':?-{2,}:?', c) for c in cells if c):
                continue                         # dòng ngăn kiểu markdown |---|---|
            cur.append(cells)
            continue
        m = re.match(r'^((?:\[[CRLBI]\])+)\s*(.*)$', l)
        tags = set(re.findall(r'\[([CRLBI])\]', m.group(1))) if m else set()
        text = _chuan_dong(m.group(2) if m else l)
        if text:
            khoi.append(('p', (text, tags)))
    if cur:
        khoi.append(('bang', cur))               # AI quên đóng [[/BANG]]
    return khoi


def _dung_docx_trung_gian(khoi):
    """Dựng file Word trung gian đúng yêu cầu ở LENH_PDF_SANG_WORD (TNR 14; lề 2/2/3/2; thụt 1,27; giãn 6pt; căn đều)."""
    doc = Document()
    _thiet_lap_trang(doc, 14)
    for loai, nd in khoi:
        if loai == 'bang':
            _dung_bang(doc, [[dict(t=c, span=1, vm=None) for c in h] for h in nd], 13)
            continue
        text, tags = nd
        align = 'center' if 'C' in tags else 'right' if 'R' in tags else 'left' if 'L' in tags else 'justify'
        _p(doc, text, 14, bold='B' in tags, italic='I' in tags, align=align, indent=(align == 'justify'))
    bio = io.BytesIO(); doc.save(bio)
    return bio.getvalue()


def pdf_sang_docx_bang_ai(data, ai_call):
    """BƯỚC 1 với file PDF: gửi PDF cho AI kèm lệnh trích xuất (LENH_PDF_SANG_WORD) -> dựng file Word trung gian."""
    import base64
    b64 = base64.standard_b64encode(data).decode('ascii')
    system = ("Bạn là công cụ trích xuất nội dung văn bản từ PDF. Chỉ trả về nội dung văn bản theo đúng quy ước được nêu, "
              "không giải thích, không tóm tắt, không bỏ sót.")
    noi_dung = [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
                {"type": "text", "text": LENH_PDF_SANG_WORD + QUY_UOC_TRA_KET_QUA}]
    kq = ai_call(system, [{"role": "user", "content": noi_dung}])
    if not isinstance(kq, str):
        raise ValueError("AI trả về dữ liệu không hợp lệ.")
    if kq.strip().startswith(('⚠️', '❌')):
        raise ValueError(re.sub(r'[*`]', '', kq).strip())
    khoi = _parse_khoi_ai(kq)
    if sum(len(x[1][0]) for x in khoi if x[0] == 'p') < 30:
        raise ValueError("AI không trích được nội dung từ PDF.")
    return _dung_docx_trung_gian(khoi)


def chuan_hoa(data, ten_file, cfg=None, che_do='auto', tu_bo_sung=True, ai_call=None, dung_ai=False, pdf_qua_ai=True):
    """Pipeline đầy đủ. che_do: 'auto' | 14 | 13. Trả về dict kết quả.
    File PDF (khi pdf_qua_ai và có ai_call): BƯỚC 1 = AI trích xuất PDF -> Word trung gian;
    BƯỚC 2 = chuẩn hóa thể thức file Word đó theo logic chung của app."""
    ghi_chu_pdf, docx_tg = '', None
    la_pdf = (ten_file or '').lower().endswith('.pdf')
    if la_pdf and pdf_qua_ai and ai_call is not None:
        try:
            docx_tg = pdf_sang_docx_bang_ai(data, ai_call)                      # BƯỚC 1
            data, ten_file = docx_tg, 'trich_xuat_tu_pdf.docx'
            ghi_chu_pdf = "Đã trích xuất PDF → Word bằng AI (bước 1), rồi chuẩn hóa thể thức (bước 2)."
        except Exception as e:
            docx_tg = None
            ghi_chu_pdf = f"Không trích xuất PDF bằng AI được ({e}) → dùng cách đọc PDF thông thường/OCR."
    lines, bang = doc_dong_va_bang(data, ten_file)                              # BƯỚC 2 bắt đầu
    la_scan = la_pdf and docx_tg is None and _la_file_scan(data)
    if not lines:
        raise ValueError("Văn bản trống hoặc không đọc được nội dung.")
    # --- LỌC RÁC do con dấu / chữ ký tay TRƯỚC khi nhận diện thể thức ---
    lines = _loc_dong_rac(lines)
    # Bước rà soát 1: tự sửa an toàn khoảng trắng/dấu câu (không đổi nội dung), rồi quét lỗi bằng quy tắc
    so_cho_sua = 0
    lines_sach = []
    for t in lines:
        t2, k = sua_dau_cau(t)
        so_cho_sua += k
        lines_sach.append(t2)
    lines = lines_sach
    ct = phan_tich(lines, bang)
    loi = ra_soat_ngoai_tuyen(lines, ct)
    ai_note = ''
    if dung_ai:
        loi_ai, ai_note = ra_soat_ai(lines, ai_call)
        loi += loi_ai
    ghi_chu_trang = ''
    if che_do in (13, 14):
        co = che_do
        ghi_chu_trang = f"Ép cỡ chữ lớn nhất {co} theo lựa chọn."
        out, bs = dung_docx(ct, co, cfg, tu_bo_sung)
    else:
        out14, bs14 = dung_docx(ct, 14, cfg, tu_bo_sung)
        so_trang = _dem_trang_bang_libreoffice(out14)
        if so_trang is not None:
            nguon = f"đếm thực tế: {so_trang} trang"
        else:
            so_trang = uoc_luong_chieu_cao(ct, 14) / 25.7
            nguon = f"ước tính: ~{so_trang:.1f} trang"
        if so_trang <= 2.0 + 1e-9:
            co, out, bs = 14, out14, bs14
        else:
            co = 13
            out, bs = dung_docx(ct, 13, cfg, tu_bo_sung)
        ghi_chu_trang = f"Văn bản {'1-2 trang' if co == 14 else 'dài hơn 2 trang'} ({nguon}) → cỡ chữ lớn nhất {co}."
    return dict(docx=out, co=co, ghi_chu_trang=ghi_chu_trang, bo_sung=bs, cau_truc=ct, so_dong=len(lines),
                ra_soat=loi, so_cho_sua=so_cho_sua, la_scan=la_scan, ai_note=ai_note, da_dung_ai=bool(dung_ai),
                ghi_chu_pdf=ghi_chu_pdf, docx_trung_gian=docx_tg)


# ───────────────────────── 4. Giao diện Streamlit ─────────────────────────
_NHAN = {
    'can_cu': 'Căn cứ', 'tham_quyen': 'Thẩm quyền ban hành', 'quyet_dinh_op': 'QUYẾT ĐỊNH:', 'cap_tieu_de': 'Phần/Chương/Mục',
    'tieu_de_in_hoa': 'Tiêu đề (in hoa)', 'dieu': 'Điều', 'khoan': 'Khoản', 'khoan_tieu_de': 'Khoản có tiêu đề',
    'diem': 'Điểm', 'roman': 'Mục La Mã', 'gach': 'Gạch đầu dòng', 'text': 'Đoạn văn', 'khan': 'Dấu mức độ khẩn',
    'phu_luc': 'Phụ lục',
}


def render(cfg=None, ai_call=None):
    cfg = cfg or {}
    st.title("📐 Chuẩn hóa thể thức VB")
    st.caption("Tải lên 1 văn bản → bấm **Render** → app nhận diện các thành phần thể thức và xuất file Word đúng "
               "Nghị định 30/2020/NĐ-CP (khổ A4; lề trên 2 / dưới 2 / trái 3 / phải 2 cm; Times New Roman; giãn đoạn 6pt; "
               "giãn dòng Single; thụt đầu dòng 1,27 cm; số trang từ trang 2). "
               "Văn bản 1–2 trang: cỡ chữ lớn nhất **14**; dài hơn 2 trang: cỡ chữ lớn nhất **13**.")
    up = st.file_uploader("📎 Chọn văn bản cần chuẩn hóa (.docx / .txt / .pdf)", type=['docx', 'txt', 'pdf'], key='cht_upload')
    c1, c2 = st.columns(2)
    with c1:
        che_do = st.radio("Cỡ chữ", ["Tự động (theo số trang)", "Ép cỡ 14", "Ép cỡ 13"], horizontal=True, key='cht_che_do')
    with c2:
        tu_bo_sung = st.checkbox("Tự bổ sung thành phần còn thiếu (quốc hiệu, tiêu ngữ, số, nơi nhận...)", value=True, key='cht_bs')
    dung_ai = st.checkbox("🔎 Rà soát chính tả & câu bất hợp lý bằng AI (chậm hơn vài giây; cần cấu hình ANTHROPIC_API_KEY)",
                          value=ai_call is not None, disabled=ai_call is None, key='cht_ai')
    pdf_ai = st.checkbox("📄 File PDF: nhờ AI trích xuất sang Word trước (bước 1), rồi mới chuẩn hóa thể thức (bước 2) — "
                         "khuyên dùng, giữ được bảng biểu", value=ai_call is not None, disabled=ai_call is None, key='cht_pdf_ai')
    if st.button("🔄 Render", type="primary", disabled=up is None, key='cht_btn'):
        try:
            with st.spinner("Đang nhận diện và dựng lại văn bản..."):
                kq = chuan_hoa(up.getvalue(), up.name, cfg,
                               che_do={'Ép cỡ 14': 14, 'Ép cỡ 13': 13}.get(che_do, 'auto'), tu_bo_sung=tu_bo_sung,
                               ai_call=ai_call, dung_ai=dung_ai, pdf_qua_ai=pdf_ai)
            kq['ten_goc'] = up.name
            st.session_state['cht_kq'] = kq
        except Exception as e:
            st.session_state.pop('cht_kq', None)
            st.error(f"❌ Không chuẩn hóa được: {e}")
    kq = st.session_state.get('cht_kq')
    if not kq:
        return
        
    # Cảnh báo nếu file vào là PDF scan (không có text layer)
    if kq.get('la_scan'):
        st.warning(
            "⚠️ File này là **bản scan (ảnh)**, không phải PDF văn bản. Em đã chạy OCR tiếng Việt "
            "và tự động loại bỏ con dấu đỏ + vùng chữ ký tay. Tuy nhiên, **họ tên người ký tay không "
            "được lấy tự động** (chữ ký viết ngoáy OCR không chính xác) — anh/chị vui lòng kiểm tra "
            "và tự ký lại khi in. Nếu có file **.docx gốc**, nên dùng file đó để kết quả chính xác nhất."
        )
    ct = kq['cau_truc']
    st.success(f"✅ Đã chuẩn hóa. {kq['ghi_chu_trang']}")
    ten_ra = re.sub(r'\.[^.]+$', '', kq['ten_goc']) + '_NĐ30.docx'
    st.download_button("📥 Tải văn bản đã chuẩn hóa (Word)", data=kq['docx'], file_name=ten_ra,
                       mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                       type="primary", key='cht_dl')
    if kq.get('ghi_chu_pdf'):
        st.info("📄 " + kq['ghi_chu_pdf'])
    if kq.get('docx_trung_gian'):
        st.download_button("📥 Tải file Word trung gian (kết quả bước 1 từ PDF)", data=kq['docx_trung_gian'],
                           file_name=re.sub(r'\.[^.]+$', '', kq['ten_goc']) + '_buoc1.docx',
                           mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', key='cht_dl_tg')
    if kq['bo_sung']:
        st.warning("ℹ️ Văn bản gốc thiếu các thành phần sau, app đã bổ sung/để trống: " + "; ".join(kq['bo_sung']))
    if kq.get('so_cho_sua'):
        st.caption(f"🧹 Đã tự chỉnh {kq['so_cho_sua']} chỗ khoảng trắng/dấu câu (không đổi nội dung).")
    loi = kq.get('ra_soat') or []
    st.markdown(f"#### 🔎 Rà soát chính tả & câu: {len(loi)} điểm cần xem" if loi else "#### 🔎 Rà soát chính tả & câu: chưa phát hiện lỗi")
    if loi:
        import pandas as _pd
        st.caption("Các mục dưới đây chỉ là ĐỀ XUẤT để anh/chị đối chiếu — app không tự sửa nội dung văn bản gốc. "
                   "Sửa trong văn bản gốc rồi Render lại nếu cần.")
        st.dataframe(_pd.DataFrame([{'Loại': x['muc'], 'Nguồn': x['nguon'], 'Đoạn có vấn đề': x['trich'],
                                     'Vấn đề': x['van_de'], 'Đề xuất': x['de_xuat']} for x in loi]),
                     hide_index=True, width='stretch')
    if kq.get('da_dung_ai') and kq.get('ai_note'):
        st.info(f"ℹ️ AI: {kq['ai_note']}")
    elif not kq.get('da_dung_ai'):
        st.caption("Chưa bật rà soát bằng AI — chỉ kiểm tra bằng quy tắc (lỗi thường gặp, lặp từ, ngoặc, ngày tháng...).")
    with st.expander("🔍 Kết quả nhận diện thành phần (kiểm tra trước khi dùng)", expanded=False):
        st.markdown(
            f"- **Quốc hiệu:** {'✔' if ct['quoc_hieu'] else '✘'}  |  **Tiêu ngữ:** {'✔' if ct['tieu_ngu'] else '✘'}\n"
            f"- **Tên cơ quan:** {' / '.join(ct['co_quan']) or '—'}\n"
            f"- **Số, ký hiệu:** {ct['so'] or '—'}  |  **Địa danh, ngày:** {ct['ngay'] or '—'}\n"
            f"- **Tên loại:** {ct['loai'] or ('Công văn' if ct['la_cong_van'] else '—')}  |  "
            f"**Trích yếu:** {ct['trich_yeu'] or ct['v_v'] or '—'}\n"
            f"- **Kính gửi:** {'; '.join(ct['kg_items']) or '—'}\n"
            f"- **Nơi nhận:** {'; '.join(ct['noi_nhan']) or '—'}\n"
            f"- **Người ký:** {' / '.join(ct['ky']['quyen_han'] + ct['ky']['chuc_vu']) or '—'} — {ct['ky']['ho_ten'] or '—'}")
        thong_ke = {}
        for loai, _ in ct['body']:
            thong_ke[loai] = thong_ke.get(loai, 0) + 1
        st.caption("Thân văn bản: " + ", ".join(f"{_NHAN.get(k, k)}: {v}" for k, v in thong_ke.items()))
        st.caption("Nếu nhận diện sai (VD: dòng tiêu đề bị coi là đoạn văn), hãy sửa văn bản gốc cho rõ "
                   "(vd. viết 'Điều 1.', '1.', 'a)', tiêu đề in HOA) rồi Render lại.")
