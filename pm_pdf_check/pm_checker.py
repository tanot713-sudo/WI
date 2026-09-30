# -*- coding: utf-8 -*-
"""
ตรวจสอบใบรายงานการตรวจสอบสภาพอุปกรณ์ (PM) ที่สแกนเป็น PDF เทียบกับ Excel

ทำงานแบบออฟไลน์ (Tesseract + OpenCV) ไม่ใช้ API ภายนอก
สิ่งที่ตรวจ:
  1. เลขหน้า  n/N มุมบนขวา  (N = จำนวนหน้าจริงของ section)
  2. section  = ข้อความบรรทัดที่ 5 ของหัวกระดาษ (สำนักงานเขต...)
  3. วันที่   (ลายมือ) เทียบกับ Excel  -> ต้องยืนยันด้วยคนหนึ่งครั้ง (ดู review)
  4. ลำดับที่ เรียงต่อเนื่องข้ามหน้า 1..N
  5. ลายเซ็นครบ 2 ช่องที่หน้าสุดท้ายของ section

ผลลัพธ์เป็น list ของ Finding (หน้า / หมวด / พบ / ควรเป็น / วิธีแก้)
"""
import re
import difflib
import unicodedata
from collections import Counter
from dataclasses import dataclass, field, asdict
from datetime import datetime

import cv2
import numpy as np
import pymupdf
import pytesseract

# ----------------------------------------------------------------- CONFIG --
CONFIG = dict(
    dpi=200,
    title_line_index=5,          # บรรทัดที่ 5 ของหัวกระดาษ = ชื่อ section
    year_be=2569,                # ปี พ.ศ. ที่คาดหวัง (ใช้เมื่อผู้ใช้ยืนยันวันที่เฉพาะเลขวัน)
    expected_signatures=("ผู้ให้บริการ", "เจ้าหน้าที่เขต"),
    sig_ink_min=0.008,           # สัดส่วนหมึกขั้นต่ำในกรอบลายเซ็นจึงถือว่า "มีลายเซ็น"
    sig_boxes=dict(              # (x0, x1) เป็นสัดส่วนความกว้างหน้า
        provider=(0.26, 0.48),
        officer=(0.64, 0.92),
    ),
    date_band=(0.175, 0.235),    # ช่วง y (สัดส่วนความสูงหน้า) ของบรรทัด "วันที่ ... เดือน ... พ.ศ."
    date_xrange=(0.25, 0.75),
)

THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
THAI_MONTH_ABBR = {"ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4, "พ.ค.": 5, "มิ.ย.": 6,
                   "ก.ค.": 7, "ส.ค.": 8, "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12}


# ------------------------------------------------------------------ MODELS --
@dataclass
class Finding:
    section: str
    page: int                 # หน้าในไฟล์ PDF (เริ่มที่ 1)
    category: str             # เลขหน้า / วันที่ / ลำดับที่ / ลายเซ็น / อ่านไม่ได้
    found: str                # สิ่งที่พบในเอกสาร
    expected: str             # สิ่งที่ควรเป็น
    fix: str                  # สิ่งที่ต้องแก้ไข
    severity: str = "ผิด"     # ผิด / ตรวจซ้ำ (ระบบไม่มั่นใจ)

    def key(self):
        return (self.section, self.page, self.category)


@dataclass
class PageInfo:
    page: int
    page_no: int | None = None
    page_total: int | None = None
    page_text: str = ""
    title: str = ""
    seq_numbers: list = field(default_factory=list)   # [(เลข|None, ความมั่นใจ)]
    has_sig: dict = field(default_factory=dict)       # {'provider': bool, 'officer': bool}
    sig_ink: dict = field(default_factory=dict)
    date_crop: np.ndarray | None = None
    seq_crops: list = field(default_factory=list)
    sig_crops: dict = field(default_factory=dict)


# ------------------------------------------------------------ IMAGE / OCR ---
def norm_th(s: str) -> str:
    """ทำให้ข้อความไทยเปรียบเทียบได้: NFC, แก้ สระอำ ที่ Tesseract แยกเป็น นิคหิต+า"""
    s = unicodedata.normalize("NFC", s or "")
    s = s.replace("ํา", "ำ")   # ํา -> ำ
    s = re.sub(r"\s+", "", s)
    return s


def render_pages(pdf_path, dpi=200):
    doc = pymupdf.open(pdf_path)
    for p in doc:
        pm = p.get_pixmap(dpi=dpi)
        a = np.frombuffer(pm.samples, np.uint8).reshape(pm.h, pm.w, pm.n)
        yield p.number + 1, cv2.cvtColor(a, cv2.COLOR_RGB2GRAY if pm.n == 1 else
                                         (cv2.COLOR_RGB2BGR if pm.n == 3 else cv2.COLOR_RGBA2BGR))


def deskew(bgr):
    """แก้ภาพสแกนเอียงจากเส้นแนวนอนของตาราง (มุมเล็กๆ)"""
    g = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY) if bgr.ndim == 3 else bgr
    edges = cv2.Canny(g, 60, 160)
    H, W = g.shape
    lines = cv2.HoughLinesP(edges, 1, np.pi / 720, threshold=200,
                            minLineLength=int(.3 * W), maxLineGap=15)
    if lines is None:
        return bgr, 0.0
    angs = []
    for x1, y1, x2, y2 in lines.reshape(-1, 4):
        a = np.degrees(np.arctan2(y2 - y1, x2 - x1))
        if abs(a) < 5:
            angs.append(a)
    if not angs:
        return bgr, 0.0
    ang = float(np.median(angs))
    if abs(ang) < 0.05:
        return bgr, ang
    M = cv2.getRotationMatrix2D((W / 2, H / 2), ang, 1.0)
    out = cv2.warpAffine(bgr, M, (W, H), flags=cv2.INTER_CUBIC,
                         borderMode=cv2.BORDER_REPLICATE)
    return out, ang


def ocr_page_number(g):
    H, W = g.shape
    crop = g[: int(.06 * H), int(.8 * W):]
    crop = cv2.resize(crop, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    txt = pytesseract.image_to_string(crop, lang="eng", config="--psm 7").strip()
    m = re.search(r"(\d+)\s*[/／]\s*(\d+)", txt)
    return (int(m.group(1)), int(m.group(2)), txt) if m else (None, None, txt)


TITLE_KEYS = ("สำนัก", "กอง", "เขต", "มหาวิทยาลัย", "สภา", "ศูนย์", "โรงพยาบาล", "สถาบัน", "ศาลา")


def ocr_title(g, line_index=5):
    """คืนข้อความบรรทัดที่ 5 ของหัวกระดาษ (ไล่จากบรรทัดของ 'จ้างบำรุง...' ไปอีก 1 บรรทัด)"""
    H, W = g.shape
    hdr = g[int(.03 * H): int(.24 * H)]
    txt = pytesseract.image_to_string(hdr, lang="tha", config="--psm 6")
    lines = [unicodedata.normalize("NFC", l).replace("ํา", "ำ").strip()
             for l in txt.splitlines()]
    lines = [l for l in lines if len(re.sub(r"[\W_]", "", l)) >= 4]
    for i, l in enumerate(lines):
        if "จ้าง" in l and "บำรุง" in l.replace("บํารุง", "บำรุง") and i + 1 < len(lines):
            return lines[i + 1], lines
    for l in lines:                       # สำรอง: บรรทัดแรกที่ดูเป็นชื่อหน่วยงาน
        if any(k in l for k in TITLE_KEYS) and "จำกัด" not in l and "บริษัท" not in l:
            return l, lines
    return "", lines


def crop_date_band(g):
    H, W = g.shape
    y0, y1 = CONFIG["date_band"]
    x0, x1 = CONFIG["date_xrange"]
    return g[int(y0 * H): int(y1 * H), int(x0 * W): int(x1 * W)].copy()


# --- ลำดับที่ ---------------------------------------------------------------
def _header_bands(g):
    """แถวสีเทา (หัวข้ออุปกรณ์/สถานที่) ในคอลัมน์ 'ลำดับที่' -> [(y0,y1)]"""
    H, W = g.shape
    x0, x1 = int(.066 * W), int(.12 * W)
    p80 = np.percentile(g[:, x0:x1].astype(float), 80, axis=1)
    gray = (p80 > 165) & (p80 < 232)
    out, s = [], None
    for y in range(H):
        if gray[y] and s is None:
            s = y
        if (not gray[y]) and s is not None:
            if y - s >= 15:
                out.append((s, y))
            s = None
    return out, (x0, x1)


def _read_digit(cell):
    """อ่านเลขในช่อง โดยโหวตจากหลายโหมดของ Tesseract คืน (เลข|None, ink)"""
    h, w = cell.shape
    c = cell[int(.1 * h): int(.9 * h), int(.15 * w): int(.85 * w)]
    c = cv2.resize(c, None, fx=6, fy=6, interpolation=cv2.INTER_CUBIC)
    _, inv = cv2.threshold(c, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    n, lab, st, _ = cv2.connectedComponentsWithStats(inv)
    keep = np.zeros_like(inv)
    for i in range(1, n):                 # ตัดเส้นขอบตาราง/จุดเล็กๆ ออก
        _, _, cw, ch, a = st[i]
        if a > 300 and .25 * c.shape[0] < ch < .95 * c.shape[0] and cw < .8 * c.shape[1]:
            keep[lab == i] = 255
    ink = float((keep > 0).mean())
    if ink < .004:
        return None, ink
    img = cv2.copyMakeBorder(255 - keep, 60, 60, 60, 60, cv2.BORDER_CONSTANT, value=255)
    votes = []
    for cfg in ("--psm 10", "--psm 7", "--psm 10 -c tessedit_char_whitelist=0123456789",
                "--psm 8 -c tessedit_char_whitelist=0123456789"):
        t = re.sub(r"\D", "", pytesseract.image_to_string(img, lang="eng", config=cfg))
        if t:
            votes.append(t)
    if not votes:
        return "?", ink
    return Counter(votes).most_common(1)[0][0], ink


def read_seq_numbers(g):
    """คืน [(เลข|None|'?', y0, y1, crop)] ของแถวหัวข้อในหน้า
    None = ไม่มีเลข (แถวชื่อสถานที่), '?' = มีตัวเลขแต่อ่านไม่ออก"""
    bands, (x0, x1) = _header_bands(g)
    out = []
    for a, b in bands:
        cell = g[a:b, x0:x1]
        val, _ = _read_digit(cell)
        out.append((val, a, b, cell.copy()))
    # แถวชื่อสถานที่ (เช่น "ฝ่ายปกครอง ชั้น 2") ไม่มีเลข และอยู่ติดกับแถวอุปกรณ์แถวแรก -> ตัดทิ้ง
    keep = []
    for i, o in enumerate(out):
        if i + 1 < len(out) and out[i + 1][1] - o[2] < 12:
            continue
        keep.append(o)
    return keep


# --- ลายเซ็น ----------------------------------------------------------------
def _long_hlines(g, thr=110, frac=.35):
    H, W = g.shape
    bw = (g < thr).astype(np.uint8)
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (int(frac * W), 1))
    op = cv2.morphologyEx(bw, cv2.MORPH_OPEN, k)
    ys = np.where(op.sum(1) > 0)[0]
    out = []
    for y in ys:
        if not out or y - out[-1][1] > 3:
            out.append([y, y])
        else:
            out[-1][1] = y
    return out


def signature_boxes(g):
    """หากรอบลายเซ็น 2 ช่อง จากเส้นล่างสุดของตาราง"""
    H, W = g.shape
    lines = [l for l in _long_hlines(g) if l[0] < .8 * H]
    if not lines:
        return None
    table_bottom = lines[-1][1]
    y0 = table_bottom + int(.004 * H)
    y1 = table_bottom + int(.05 * H)       # เหนือแถวชื่อตัวบรรจง/ตราประทับ
    boxes = {}
    for k, (xa, xb) in CONFIG["sig_boxes"].items():
        boxes[k] = (int(xa * W), y0, int(xb * W), y1)
    return boxes


def sig_ink(bgr, box):
    """สัดส่วนหมึกลายมือ (สีน้ำเงิน หรือเส้นเข้ม) ในกรอบ หลังตัดเส้นประ/จุดพิมพ์"""
    x0, y0, x1, y1 = box
    c = bgr[y0:y1, x0:x1]
    hsv = cv2.cvtColor(c, cv2.COLOR_BGR2HSV)
    gray = cv2.cvtColor(c, cv2.COLOR_BGR2GRAY)
    bw = (((hsv[..., 1] > 70) & (hsv[..., 2] > 40)) | (gray < 110)).astype(np.uint8) * 255
    n, lab, st, _ = cv2.connectedComponentsWithStats(bw)
    keep = np.zeros_like(bw)
    for i in range(1, n):
        _, _, cw, ch, a = st[i]
        if a > 25 and max(cw, ch) > 14:        # เส้นประ = จุดเล็กๆ แยกกัน จึงถูกตัดทิ้ง
            keep[lab == i] = 255
    return float((keep > 0).mean()), c.copy()


# ------------------------------------------------------------ ANALYSE PDF --
def analyse_pdf(pdf_path, progress=None):
    infos = []
    for pno, bgr in render_pages(pdf_path, CONFIG["dpi"]):
        bgr, _ = deskew(bgr)
        g = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        pi = PageInfo(page=pno)
        pi.page_no, pi.page_total, pi.page_text = ocr_page_number(g)
        pi.title, _ = ocr_title(g, CONFIG["title_line_index"])
        pi.date_crop = crop_date_band(g) if pi.page_no in (1, None) else None
        nums = read_seq_numbers(g)
        pi.seq_numbers = [n[0] for n in nums]
        pi.seq_crops = [n[3] for n in nums]
        boxes = signature_boxes(g)
        if boxes:
            for k, box in boxes.items():
                ink, crop = sig_ink(bgr, box)
                pi.sig_ink[k] = ink
                pi.sig_crops[k] = crop
                pi.has_sig[k] = ink >= CONFIG["sig_ink_min"]
        infos.append(pi)
        if progress:
            progress(pno)
    return infos


# ---------------------------------------------------------------- SECTIONS --
def group_sections(infos):
    """แบ่ง section: หน้าที่ title (บรรทัดที่ 5) เหมือนกันและอยู่ติดกัน = section เดียว
    ถ้า title อ่านไม่ออก ใช้เลขหน้า n/N ช่วย (n=1 คือเริ่ม section ใหม่)"""
    sections, cur = [], None
    for pi in infos:
        t = norm_th(pi.title)
        new = cur is None
        if cur is not None:
            same = t and cur["key"] and difflib.SequenceMatcher(None, t, cur["key"]).ratio() >= .85
            if pi.page_no == 1 and cur["pages"][-1].page_no != 1:
                same = False
            if not t and pi.page_no not in (None, 1):
                same = True
            new = not same
        if new:
            cur = {"key": t, "title": pi.title, "pages": []}
            sections.append(cur)
        elif not cur["key"] and t:
            cur["key"], cur["title"] = t, pi.title
        cur["pages"].append(pi)
    return sections


# ------------------------------------------------------------------- EXCEL --
def _norm_unit(s):
    s = norm_th(str(s or ""))
    return re.sub(r"^(สำนักงานเขต|สำนักงาน|เขต)", "", s)


def load_excel(xlsx_path, sheet_name):
    """อ่านชีตรายเดือน -> list ของ dict(seq, date, unit, unit_raw, operators)
    (Excel เก็บปี พ.ศ. ไว้ในช่องปี เช่น 2569 จึงไม่แปลงเป็น ค.ศ.)"""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name]
    hdr = next(i for i, r in enumerate(ws.iter_rows(max_row=10, values_only=True), 1) if r[0] == "ลำดับ")
    rows = []
    for r in ws.iter_rows(min_row=hdr + 1, values_only=True):
        if r[0] is None:
            continue
        d = r[1]
        date = (d.day, d.month, d.year) if hasattr(d, "day") else None
        ops = [o.strip() for o in re.split(r"[\n\r]+", str(r[11] or "")) if o.strip()]
        rows.append(dict(seq=r[0], date=date, unit_raw=str(r[2] or ""), unit=_norm_unit(r[2]),
                         operators=ops))
    return rows


def match_unit(title, excel_rows, cutoff=.6):
    key = _norm_unit(title)
    if not key:
        return None, 0.0
    best, score = None, 0.0
    for r in excel_rows:
        cand = r["unit"]
        u = cand.split("ศาลาว่าการ")[0] if "ศาลาว่าการ" in cand else cand
        sc = max(difflib.SequenceMatcher(None, key, cand).ratio(),
                 difflib.SequenceMatcher(None, key, u).ratio())
        if sc > score:
            best, score = r, sc
    return (best, score) if score >= cutoff else (None, score)


def fmt_date(d):
    if not d:
        return "-"
    day, month, year = d
    return f"{day} {THAI_MONTHS[month - 1]} {year}"


def parse_manual_date(text, default_month=None, default_year=None):
    """รับรูปแบบ '9' , '9/9' , '9/9/2569' -> (d,m,y)"""
    t = (text or "").strip()
    if not t:
        return None
    p = [x for x in re.split(r"[/\-. ]+", t) if x]
    d = int(p[0])
    m = int(p[1]) if len(p) > 1 else default_month
    y = int(p[2]) if len(p) > 2 else default_year
    if y and y < 100:
        y += 2500
    if y and y < 2400:
        y += 543
    return (d, m, y)


# ------------------------------------------------------------------- CHECK --
def check_sections(infos, excel_rows, manual_dates=None, default_month=None):
    """คืน (findings, summary_rows)
    manual_dates: {ชื่อ section(norm) หรือหน้าแรก(int): (d,m,y)} ที่คนยืนยันแล้ว"""
    manual_dates = manual_dates or {}
    findings, summary = [], []
    for sec in group_sections(infos):
        pages = sec["pages"]
        title = sec["title"] or "(อ่านชื่อ section ไม่ได้)"
        first = pages[0].page
        n_pages = len(pages)
        add = lambda pg, cat, found, exp, fix, sev="ผิด": findings.append(
            Finding(title, pg, cat, found, exp, fix, sev))

        if not sec["title"]:
            add(first, "ชื่อ section", "อ่านบรรทัดที่ 5 ไม่ได้", "ชื่อหน่วยงาน", "ตรวจด้วยตา", "ตรวจซ้ำ")

        # 1) เลขหน้า
        for i, pi in enumerate(pages, 1):
            if pi.page_no is None:
                add(pi.page, "เลขหน้า", f"อ่านไม่ได้ ('{pi.page_text}')", f"{i}/{n_pages}",
                    "ตรวจด้วยตา / เขียนเลขหน้ามุมบนขวา", "ตรวจซ้ำ")
                continue
            if pi.page_no != i or pi.page_total != n_pages:
                add(pi.page, "เลขหน้า", f"{pi.page_no}/{pi.page_total}", f"{i}/{n_pages}",
                    f"แก้เลขหน้าจาก {pi.page_no}/{pi.page_total} เป็น {i}/{n_pages}")

        # 2) วันที่ เทียบ Excel
        row, score = match_unit(sec["title"], excel_rows)
        if row is None:
            add(first, "วันที่", "-", "-", f"หา '{sec['title']}' ใน Excel ไม่พบ (ความคล้าย {score:.2f})", "ตรวจซ้ำ")
        else:
            got = manual_dates.get(first)
            if got is None:
                add(first, "วันที่", "ยังไม่ได้ยืนยันวันที่จากลายมือ",
                    fmt_date(row["date"]), "กรอกวันที่ที่เห็นในเอกสารในขั้นตอน review", "ตรวจซ้ำ")
            else:
                d, m, y = got
                m = m or default_month
                got = (d, m, y or CONFIG["year_be"])
                if got != row["date"]:
                    add(first, "วันที่", fmt_date(got), fmt_date(row["date"]),
                        f"แก้วันที่ในเอกสารจาก {fmt_date(got)} เป็น {fmt_date(row['date'])}")

        # 3) ลำดับที่ ต่อเนื่องข้ามหน้า
        expect = 1
        for pi in pages:
            for v in pi.seq_numbers:
                if v is None or v == "":
                    continue           # แถวชื่อสถานที่ (ไม่มีเลข)
                if v == "?":
                    add(pi.page, "ลำดับที่", "อ่านตัวเลขไม่ออก", str(expect),
                        "ตรวจด้วยตา", "ตรวจซ้ำ")
                    expect += 1
                    continue
                if v != str(expect):
                    add(pi.page, "ลำดับที่", v, str(expect),
                        f"แก้ลำดับที่จาก {v} เป็น {expect}")
                expect += 1

        # 4) ลายเซ็น ที่หน้าสุดท้ายของ section
        last = pages[-1]
        names = {"provider": "เจ้าหน้าที่ผู้ให้บริการ", "officer": "เจ้าหน้าที่ (ฝั่งเขต/สำนัก)"}
        if not last.sig_ink:
            add(last.page, "ลายเซ็น", "หาตำแหน่งช่องลายเซ็นไม่พบ", "ลายเซ็น 2 ช่อง", "ตรวจด้วยตา", "ตรวจซ้ำ")
        else:
            for k, label in names.items():
                if not last.has_sig.get(k):
                    add(last.page, "ลายเซ็น", f"ไม่พบลายเซ็น {label}", f"มีลายเซ็น {label}",
                        f"ให้ {label} ลงลายมือชื่อ")
        # ลายเซ็นในหน้าก่อนหน้า (ไม่ควรมี) ไม่ตรวจ

        summary.append(dict(section=title, first_page=first, pages=n_pages,
                            excel_date=fmt_date(row["date"]) if row else "-",
                            unit_score=round(score, 2)))
    return findings, summary


# ================================================================ REPORTS ==
import io
import json
import pandas as pd

HDR_FIND = ["#", "Section", "หน้า", "หมวด", "พบ", "ควรเป็น", "สิ่งที่ต้องแก้ไข", "ระดับ"]
HDR_HIST = ["รอบ", "วันที่ตรวจ", "Section", "หน้า", "หมวด", "แก้จาก", "แก้เป็น", "สถานะ"]


def findings_df(findings):
    return pd.DataFrame([asdict(f) for f in findings]).rename(columns=dict(
        section="Section", page="หน้า", category="หมวด", found="พบ", expected="ควรเป็น",
        fix="สิ่งที่ต้องแก้ไข", severity="ระดับ"))


def _png(img, max_w=520):
    h, w = img.shape[:2]
    if w > max_w:
        img = cv2.resize(img, (max_w, int(h * max_w / w)))
    ok, buf = cv2.imencode(".png", img)
    return io.BytesIO(buf.tobytes())


# ---------- อ่านรายงาน Word ฉบับก่อน (โหมด import Word) ----------
def read_previous_report(docx_path):
    """คืน (round_no, findings_ก่อนหน้า, history_rows_ก่อนหน้า)"""
    import docx
    d = docx.Document(docx_path)
    prev, hist, rnd = [], [], 0
    for t in d.tables:
        head = [c.text.strip() for c in t.rows[0].cells]
        if head[:len(HDR_FIND)] == HDR_FIND:
            for r in t.rows[1:]:
                c = [x.text.strip() for x in r.cells]
                try:
                    prev.append(Finding(c[1], int(c[2]), c[3], c[4], c[5], c[6], c[7] or "ผิด"))
                except ValueError:
                    pass
        elif head[:len(HDR_HIST)] == HDR_HIST:
            for r in t.rows[1:]:
                c = [x.text.strip() for x in r.cells]
                hist.append(c[:len(HDR_HIST)])
                if c[0].isdigit():
                    rnd = max(rnd, int(c[0]))
    for p in d.paragraphs:                       # "รอบที่ตรวจ: N"
        m = re.search(r"รอบที่ตรวจ:\s*(\d+)", p.text)
        if m:
            rnd = max(rnd, int(m.group(1)))
    return rnd, prev, hist


def compare_with_previous(prev, new, round_no, when):
    """เทียบรอบก่อนกับรอบนี้ -> แถวประวัติการแก้ไข (แก้จากอะไรเป็นอะไร)"""
    nk = {}
    for f in new:
        nk.setdefault((norm_th(f.section), f.category), []).append(f)
    rows = []
    used = set()
    for o in prev:
        cands = [f for f in nk.get((norm_th(o.section), o.category), []) if f.page == o.page] \
            or nk.get((norm_th(o.section), o.category), [])
        if not cands:
            rows.append([str(round_no), when, o.section, str(o.page), o.category,
                         o.found, o.expected, "แก้ไขแล้ว (ตรวจซ้ำผ่าน)"])
        else:
            n = cands[0]
            used.add(id(n))
            if n.found == o.found:
                rows.append([str(round_no), when, o.section, str(o.page), o.category,
                             o.found, o.found, "ยังไม่ได้แก้ไข"])
            else:
                rows.append([str(round_no), when, o.section, str(n.page), o.category,
                             o.found, n.found, f"แก้แล้วแต่ยังไม่ถูก (ควรเป็น {n.expected})"])
    for n in new:
        if id(n) not in used and not any(
                norm_th(o.section) == norm_th(n.section) and o.category == n.category for o in prev):
            rows.append([str(round_no), when, n.section, str(n.page), n.category,
                         "-", n.found, "ปัญหาใหม่ที่พบรอบนี้"])
    return rows


def build_word_report(path, findings, summary, meta, history_rows=None, evidence=None, round_no=1):
    """เขียนรายงาน Word: สรุป / รายการที่ต้องแก้ / ประวัติการแก้ไข / หลักฐานภาพ"""
    import docx
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    d = docx.Document()
    sec = d.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(1.5))
    st = d.styles["Normal"]
    st.font.name = "TH SarabunPSK"
    st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "TH SarabunPSK")

    def shade(cell, color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), color)
        tcPr.append(shd)

    def table(head, rows, widths=None, color_col=None):
        t = d.add_table(rows=1, cols=len(head))
        t.style = "Table Grid"
        for i, h in enumerate(head):
            c = t.rows[0].cells[i]
            c.text = h
            c.paragraphs[0].runs[0].bold = True
            shade(c, "D9E2F3")
        for r in rows:
            cells = t.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = str(v)
            if color_col is not None:
                v = str(r[color_col])
                if v.startswith("ผิด") or "ยังไม่" in v:
                    shade(cells[color_col], "F8CBAD")
                elif v.startswith("ตรวจซ้ำ") or "ยังไม่ถูก" in v or "ใหม่" in v:
                    shade(cells[color_col], "FFE699")
                elif "แก้ไขแล้ว" in v:
                    shade(cells[color_col], "C6E0B4")
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t

    d.add_heading("รายงานผลการตรวจสอบใบรายงานการตรวจสอบสภาพอุปกรณ์ (PM)", 0)
    d.add_paragraph(f"รอบที่ตรวจ: {round_no}")
    for k, v in meta.items():
        d.add_paragraph(f"{k}: {v}")
    n_wrong = sum(f.severity == "ผิด" for f in findings)
    n_chk = sum(f.severity != "ผิด" for f in findings)
    d.add_paragraph(f"สรุป: พบจุดผิด {n_wrong} รายการ, ต้องตรวจซ้ำด้วยตา {n_chk} รายการ, "
                    f"จำนวน section {len(summary)}")

    d.add_heading("1. สรุปรายหน่วยงาน (section)", 1)
    bad = Counter(f.section for f in findings)
    table(["Section", "หน้าแรก (ในไฟล์ PDF)", "จำนวนหน้า", "วันที่ใน Excel", "จุดที่พบ"],
          [[s["section"], s["first_page"], s["pages"], s["excel_date"], bad.get(s["section"], 0)]
           for s in summary])

    d.add_heading("2. รายการที่ต้องแก้ไข", 1)
    if findings:
        table(HDR_FIND,
              [[i, f.section, f.page, f.category, f.found, f.expected, f.fix, f.severity]
               for i, f in enumerate(findings, 1)],
              widths=[1, 5, 1.5, 2.5, 4.5, 4.5, 6.5, 2], color_col=7)
    else:
        d.add_paragraph("ไม่พบจุดผิด")

    d.add_heading("3. ประวัติการแก้ไข (แก้จากอะไรเป็นอะไร)", 1)
    table(HDR_HIST, history_rows or [], widths=[1.2, 3, 5, 1.5, 2.5, 4, 4, 5], color_col=7)

    if evidence:
        d.add_page_break()
        d.add_heading("4. หลักฐานภาพ (ส่วนที่ระบบตัดมาตรวจ)", 1)
        for title, imgs in evidence:
            d.add_paragraph(title).runs[0].bold = True
            for label, im in imgs:
                d.add_paragraph(label)
                d.add_picture(_png(im), width=Cm(11))
    d.save(path)
    return path


def build_excel_report(path, findings, summary):
    with pd.ExcelWriter(path) as xw:
        findings_df(findings).to_excel(xw, sheet_name="จุดที่ต้องแก้", index=False)
        pd.DataFrame(summary).to_excel(xw, sheet_name="สรุป section", index=False)
    return path


def run_check(pdf_path, xlsx_path, sheet_name, manual_dates=None, default_month=None,
              prev_docx=None, out_docx="รายงานตรวจ.docx", out_xlsx="รายงานตรวจ.xlsx", progress=None, infos=None):
    """เรียกครั้งเดียวจบ: วิเคราะห์ PDF -> ตรวจ -> เขียน Word/Excel (พร้อมประวัติถ้ามีรายงานเก่า)"""
    infos = infos or analyse_pdf(pdf_path, progress)
    rows = load_excel(xlsx_path, sheet_name)
    findings, summary = check_sections(infos, rows, manual_dates, default_month)
    when = datetime.now().strftime("%d/%m/%Y %H:%M")
    round_no, hist = 1, []
    if prev_docx:
        r, prev, old_hist = read_previous_report(prev_docx)
        round_no = r + 1
        hist = old_hist + compare_with_previous(prev, findings, round_no, when)
    evidence = []
    for sec in group_sections(infos):
        imgs = []
        if sec["pages"][0].date_crop is not None:
            imgs.append((f"วันที่ (หน้า {sec['pages'][0].page})", sec["pages"][0].date_crop))
        last = sec["pages"][-1]
        for k, lab in (("provider", "ช่องลายเซ็นผู้ให้บริการ"), ("officer", "ช่องลายเซ็นเจ้าหน้าที่เขต")):
            if k in last.sig_crops:
                imgs.append((f"{lab} (หน้า {last.page}, หมึก={last.sig_ink[k]:.3f})", last.sig_crops[k]))
        evidence.append((sec["title"], imgs))
    meta = {"ไฟล์ PDF": pdf_path.split("/")[-1], "ไฟล์ Excel": xlsx_path.split("/")[-1],
            "ชีตที่ใช้เทียบ": sheet_name, "วันเวลาที่ตรวจ": when}
    build_word_report(out_docx, findings, summary, meta, hist, evidence, round_no)
    build_excel_report(out_xlsx, findings, summary)
    return infos, findings, summary
