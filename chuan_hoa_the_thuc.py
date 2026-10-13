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


def doc_dong(data, ten_file):
    ten = (ten_file or '').lower()
    if ten.endswith('.docx'):
        return doc_dong_tu_docx(data)
    if ten.endswith('.pdf'):
        try:
            try:
                from pypdf import PdfReader
            except Exception:
                from PyPDF2 import PdfReader
            rd = PdfReader(io.BytesIO(data))
            txt = '\n'.join((pg.extract_text() or '') for pg in rd.pages)
        except Exception as e:
            raise ValueError(f"Không đọc được PDF (cần thư viện pypdf): {e}")
    else:
        for enc in ('utf-8-sig', 'utf-16', 'cp1258', 'latin-1'):
            try:
                txt = data.decode(enc)
                break
            except Exception:
                continue
    return [x for x in (_chuan_dong(l) for l in txt.splitlines()) if x]


# ───────────────────────── 2. Nhận diện thành phần ─────────────────────────
RE_QH = re.compile(r'cong\s*hoa\s*xa\s*hoi\s*chu\s*nghia\s*viet\s*nam')
RE_TN = re.compile(r'doc\s*lap\s*[-–—]*\s*tu\s*do\s*[-–—]*\s*hanh\s*phuc')
RE_SO = re.compile(r'^so\s*[:.]')
RE_NGAY = re.compile(r'^[^,]{2,60},\s*ngay\s+\S+\s+thang\s+\S+\s+nam\s+\S+')
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


def _la_loai_vb(dong):
    kd = khong_dau(dong).rstrip(' :')
    return _la_in_hoa(dong) and kd in LOAI_VB


def _la_dong_ky(dong):
    kd = khong_dau(dong)
    return bool(RE_QUYEN_HAN.match(kd)) or (_la_in_hoa(dong) and len(dong) <= 70)


def phan_tich(lines):
    """Nhận diện thành phần thể thức. Trả về dict mô tả cấu trúc + danh sách nhận diện."""
    L = [_chuan_dong(x) for x in lines if _chuan_dong(x)]
    kq = dict(quoc_hieu=False, tieu_ngu=False, co_quan=[], so=None, ngay=None, loai=None, trich_yeu=None,
              v_v=None, kinh_gui=[], body=[], noi_nhan=[], ky=dict(quyen_han=[], chuc_vu=[], ho_ten=''),
              la_cong_van=False)
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
    # Kính gửi
    if pos < len(L) and kd[pos].startswith('kinh gui'):
        kq['kinh_gui'].append(L[pos]); pos += 1
        while pos < len(L) and RE_GACH.match(L[pos]) and kd[pos - 1].startswith(('kinh gui', '-', '–')):
            kq['kinh_gui'].append(L[pos]); pos += 1
        kq['la_cong_van'] = kq['la_cong_van'] or kq['loai'] is None
    elif kq['v_v'] and pos < len(L) and kd[pos].startswith('kinh gui'):
        kq['kinh_gui'].append(L[pos]); pos += 1

    rest = L[pos:]
    # --- Nơi nhận + chữ ký ---
    krest = [khong_dau(x) for x in rest]
    i_nn = next((i for i, k in enumerate(krest) if RE_NOI_NHAN.match(k)), None)
    body_lines, sig_lines = rest, []
    if i_nn is not None:
        body_lines = rest[:i_nn]
        tail = rest[i_nn + 1:]
        j = 0
        while j < len(tail) and (RE_GACH.match(tail[j]) or khong_dau(tail[j]).startswith('luu')):
            kq['noi_nhan'].append(RE_GACH.sub('', tail[j]).strip()); j += 1
        sig_lines = tail[j:]
        if not sig_lines:          # chữ ký nằm TRƯỚC "Nơi nhận" (cuối thân văn bản)
            k = len(body_lines)
            while k > 0 and (len(body_lines) - k) < 6 and len(body_lines[k - 1]) <= 70 and _la_dong_ky(body_lines[k - 1]) \
                    or (k > 0 and (len(body_lines) - k) < 6 and len(body_lines[k - 1]) <= 40
                        and not re.search(r'[.;:,]$', body_lines[k - 1])):
                k -= 1
            blk = body_lines[k:]
            if blk and any(_la_in_hoa(x) or RE_QUYEN_HAN.match(khong_dau(x)) for x in blk):
                sig_lines, body_lines = blk, body_lines[:k]
    else:
        k = len(body_lines)
        while k > 0 and (len(body_lines) - k) < 6 and len(body_lines[k - 1]) <= 70 \
                and not re.search(r'[.;:,]$', body_lines[k - 1]) and not RE_GACH.match(body_lines[k - 1]):
            k -= 1
        blk = body_lines[k:]
        if blk and any(_la_in_hoa(x) for x in blk) and len(blk) <= 5:
            sig_lines, body_lines = blk, body_lines[:k]
    # phân tách chữ ký
    ky = kq['ky']
    if sig_lines:
        sl = list(sig_lines)
        if len(sl) >= 2 and not (_la_in_hoa(sl[-1]) and any(khong_dau(sl[-1]).startswith(t) for t in CHUC_VU_TU_KHOA)):
            ten = sl.pop()
            ky['ho_ten'] = ten.title() if _la_in_hoa(ten) else ten
        elif len(sl) == 1 and not _la_in_hoa(sl[0]):
            ky['ho_ten'] = sl.pop()
        for x in sl:
            (ky['quyen_han'] if RE_QUYEN_HAN.match(khong_dau(x)) else ky['chuc_vu']).append(x.upper())

    # --- Thân văn bản ---
    body = []
    da_den_dieu_hoac_qd = False
    n = len(body_lines)
    i = 0
    co_loai = kq['loai'] is not None
    while i < n:
        t = body_lines[i]
        k = khong_dau(t)
        loai = 'text'
        if RE_KHAN.match(k) and _la_in_hoa(t):
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
        elif not da_den_dieu_hoac_qd and RE_CAN_CU.match(k):
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
    # "Điều N." — chỉ in đậm cả dòng khi là tiêu đề ngắn; còn lại in đậm nhãn "Điều N."
    kq['body'] = body
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
    L, R = 5.6, 10.4
    ht = doc.add_table(rows=1, cols=2)
    ht.alignment = WD_TABLE_ALIGNMENT.CENTER; ht.autofit = False
    _bo_vien_bang(ht)
    cl, cr = ht.rows[0].cells
    cl.width = Cm(L); cr.width = Cm(R)
    ht.columns[0].width = Cm(L); ht.columns[1].width = Cm(R)
    # trái: (cơ quan chủ quản đứng) + cơ quan ban hành đứng đậm + kẻ + số + V/v
    for idx, ten in enumerate(co_quan):
        la_cuoi = idx == len(co_quan) - 1
        _p(cl, ten.upper(), S['cq'], bold=la_cuoi, align='center', indent=False, before=0, after=0,
           p=cl.paragraphs[0] if idx == 0 else None)
    _ke_ngang(cl, 3.0, L)
    _p(cl, so, S['so'], align='center', indent=False, before=6, after=0)
    if v_v:
        _p(cl, v_v, S['vv'], align='center', indent=False, before=4, after=0)
    _p(cr, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', S['qh'], True, align='center', indent=False, before=0, after=0,
       p=cr.paragraphs[0])
    _p(cr, 'Độc lập - Tự do - Hạnh phúc', S['tn'], True, align='center', indent=False, before=0, after=0)
    _ke_ngang(cr, 5.6, R)
    _p(cr, ngay, S['ngay'], italic=True, align='center', indent=False, before=6, after=0)


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
    _p(cr, ky.get('ho_ten', ''), S['ky'], True, align='center', indent=False, before=60 if dong else 0, after=0,
       p=cr.paragraphs[0] if not dong else None)


def _dung_than(doc, S, body):
    nd = S['nd']
    n = len(body)
    for i, (loai, t) in enumerate(body):
        tiep_la_can_cu = i + 1 < n and body[i + 1][0] == 'can_cu'
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


def uoc_luong_chieu_cao(cau_truc, co=14):
    """Ước tính chiều cao (cm) nội dung khi dàn theo bộ cỡ chữ `co` (dùng khi không có LibreOffice để đếm trang thật)."""
    cpl = 76 if co == 14 else 82
    dong_cm = 0.57 if co == 14 else 0.53
    cach_doan = 0.42
    h = 5.0                                # khối đầu
    if cau_truc['loai']:
        h += 1.4
    if cau_truc['trich_yeu']:
        h += 0.9
    for loai, t in cau_truc['body']:
        h += max(1, math.ceil(len(t) / cpl)) * dong_cm + cach_doan
    h += 5.5                               # nơi nhận + chữ ký
    return h


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
    if cau_truc['kinh_gui']:
        for idx, x in enumerate(cau_truc['kinh_gui']):
            _p(doc, x, S['nd'], align='center', indent=False, before=18 if idx == 0 else 0, after=0)
        _p(doc, '', 6, before=0, after=0)
    _dung_than(doc, S, cau_truc['body'])

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
    bio = io.BytesIO(); doc.save(bio)
    return bio.getvalue(), bo_sung


def chuan_hoa(data, ten_file, cfg=None, che_do='auto', tu_bo_sung=True):
    """Pipeline đầy đủ. che_do: 'auto' | 14 | 13. Trả về dict kết quả."""
    lines = doc_dong(data, ten_file)
    if not lines:
        raise ValueError("Văn bản trống hoặc không đọc được nội dung.")
    ct = phan_tich(lines)
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
    return dict(docx=out, co=co, ghi_chu_trang=ghi_chu_trang, bo_sung=bs, cau_truc=ct, so_dong=len(lines))


# ───────────────────────── 4. Giao diện Streamlit ─────────────────────────
_NHAN = {
    'can_cu': 'Căn cứ', 'tham_quyen': 'Thẩm quyền ban hành', 'quyet_dinh_op': 'QUYẾT ĐỊNH:', 'cap_tieu_de': 'Phần/Chương/Mục',
    'tieu_de_in_hoa': 'Tiêu đề (in hoa)', 'dieu': 'Điều', 'khoan': 'Khoản', 'khoan_tieu_de': 'Khoản có tiêu đề',
    'diem': 'Điểm', 'roman': 'Mục La Mã', 'gach': 'Gạch đầu dòng', 'text': 'Đoạn văn', 'khan': 'Dấu mức độ khẩn',
    'phu_luc': 'Phụ lục',
}


def render(cfg=None):
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
    if st.button("🔄 Render", type="primary", disabled=up is None, key='cht_btn'):
        try:
            with st.spinner("Đang nhận diện và dựng lại văn bản..."):
                kq = chuan_hoa(up.getvalue(), up.name, cfg,
                               che_do={'Ép cỡ 14': 14, 'Ép cỡ 13': 13}.get(che_do, 'auto'), tu_bo_sung=tu_bo_sung)
            kq['ten_goc'] = up.name
            st.session_state['cht_kq'] = kq
        except Exception as e:
            st.session_state.pop('cht_kq', None)
            st.error(f"❌ Không chuẩn hóa được: {e}")
    kq = st.session_state.get('cht_kq')
    if not kq:
        return
    ct = kq['cau_truc']
    st.success(f"✅ Đã chuẩn hóa. {kq['ghi_chu_trang']}")
    ten_ra = re.sub(r'\.[^.]+$', '', kq['ten_goc']) + '_NĐ30.docx'
    st.download_button("📥 Tải văn bản đã chuẩn hóa (Word)", data=kq['docx'], file_name=ten_ra,
                       mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                       type="primary", key='cht_dl')
    if kq['bo_sung']:
        st.warning("ℹ️ Văn bản gốc thiếu các thành phần sau, app đã bổ sung/để trống: " + "; ".join(kq['bo_sung']))
    with st.expander("🔍 Kết quả nhận diện thành phần (kiểm tra trước khi dùng)", expanded=False):
        st.markdown(
            f"- **Quốc hiệu:** {'✔' if ct['quoc_hieu'] else '✘'}  |  **Tiêu ngữ:** {'✔' if ct['tieu_ngu'] else '✘'}\n"
            f"- **Tên cơ quan:** {' / '.join(ct['co_quan']) or '—'}\n"
            f"- **Số, ký hiệu:** {ct['so'] or '—'}  |  **Địa danh, ngày:** {ct['ngay'] or '—'}\n"
            f"- **Tên loại:** {ct['loai'] or ('Công văn' if ct['la_cong_van'] else '—')}  |  "
            f"**Trích yếu:** {ct['trich_yeu'] or ct['v_v'] or '—'}\n"
            f"- **Kính gửi:** {'; '.join(ct['kinh_gui']) or '—'}\n"
            f"- **Nơi nhận:** {'; '.join(ct['noi_nhan']) or '—'}\n"
            f"- **Người ký:** {' / '.join(ct['ky']['quyen_han'] + ct['ky']['chuc_vu']) or '—'} — {ct['ky']['ho_ten'] or '—'}")
        thong_ke = {}
        for loai, _ in ct['body']:
            thong_ke[loai] = thong_ke.get(loai, 0) + 1
        st.caption("Thân văn bản: " + ", ".join(f"{_NHAN.get(k, k)}: {v}" for k, v in thong_ke.items()))
        st.caption("Nếu nhận diện sai (VD: dòng tiêu đề bị coi là đoạn văn), hãy sửa văn bản gốc cho rõ "
                   "(vd. viết 'Điều 1.', '1.', 'a)', tiêu đề in HOA) rồi Render lại.")
