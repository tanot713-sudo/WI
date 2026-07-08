"""
Pipeline for building a PM (Preventive Maintenance) equipment/tools workbook
directly from a "system" Work Instruction .docx file (e.g. DLTWI01AEP_system.docx).

Unlike the older notebook flow, this file does NOT rely on the WI .docx already
containing a table with 'ขั้นตอนการปฏิบัติงาน/ข้อควรระวัง' / 'รูปภาพประกอบ' /
'อุปกรณ์/เครื่องมือที่ใช้' columns. Instead it reads the "system" WI document
(narrative steps grouped under maintenance-interval headings, a dedicated
equipment section, and inline pictures with 'รูปที่N ...' captions) and builds
that table itself, using the REAL equipment list and REAL embedded pictures
found in the document instead of guessed placeholder text.
"""
import re
from dataclasses import dataclass, field

import docx
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

FIGURE_RE = re.compile(r"รูปที่\s*\d+")
EQUIPMENT_SECTION_TITLES = ("อุปกรณ์",)
PROCEDURE_SECTION_TITLES = ("วิธีปฏิบัติ",)

# Concept -> keyword patterns that identify the concept in a step's text, and
# keyword substrings that identify the concept among the document's own
# (real) equipment names. Used to link a step to the equipment that really
# exists in this WI's equipment list, instead of inventing generic tool text.
CONCEPT_RULES = [
    ("cleaning", r"ทำความสะอาด|คราบ|ฝุ่น", ("ผ้า", "สเปรย์", "แปรง")),
    ("fastener", r"น็อต|สกรู|ขัน", ("ไขควง", "ประแจ")),
    ("lubrication", r"หล่อลื่น|จาระบี|น้ำมันหล่อลื่น", ("จาระบี", "น้ำมัน")),
    ("drilling", r"เจาะ|สว่าน", ("สว่าน",)),
    ("measuring", r"วัด|บันทึก|จดบันทึก|เกจ", ("วัด", "เกจ", "บันทึก")),
]


@dataclass
class Step:
    group: str
    text: str
    images: list = field(default_factory=list)  # list of bytes
    equipment: str = ""


def iter_block_items(document):
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def _image_parts(paragraph, document):
    parts = []
    for blip in paragraph._p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid and rid in document.part.rels:
            rel = document.part.rels[rid]
            if "image" in rel.reltype:
                parts.append(rel.target_part)
    return parts


def _split_captions(text):
    """Split a paragraph like 'รูปที่1 A     รูปที่2 B' into ['A', 'B']."""
    matches = list(FIGURE_RE.finditer(text))
    if not matches:
        return []
    captions = []
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        captions.append(text[start:end].strip(" \t　-–—:"))
    return captions


def extract_equipment(document):
    """Return the list of equipment/tool names from the 'อุปกรณ์' section."""
    equipment = []
    in_section = False
    for item in iter_block_items(document):
        if isinstance(item, Table):
            if len(item.rows) == 1 and len(item.columns) == 1:
                title = item.rows[0].cells[0].text.strip()
                if in_section:
                    break
                if any(t in title for t in EQUIPMENT_SECTION_TITLES):
                    in_section = True
            continue
        if in_section:
            text = item.text.strip()
            if text:
                equipment.append(text)
    return equipment


def extract_steps(document):
    """
    Walk the 'วิธีปฏิบัติ' section and return a list of Step objects (text +
    matched images), in document order, with real embedded pictures attached
    to the step they best describe.
    """
    groups = []          # [(group_title, [Step, ...]), ...]
    current_group = None
    pending_images = []    # images seen but not yet captioned/attached
    pending_captions = []  # caption strings waiting to be matched to images
    in_section = False

    def flush_pending(target_steps):
        # Attach any images collected so far to the best-matching step among
        # the ones gathered in the *current* group only - figures illustrate
        # steps physically near them, not steps from an unrelated section.
        n = min(len(pending_images), len(pending_captions))
        for img, caption in zip(pending_images[:n], pending_captions[:n]):
            best_step = _best_match(caption, target_steps)
            if best_step is not None:
                best_step.images.append(img.blob)
        pending_images.clear()
        pending_captions.clear()

    for item in iter_block_items(document):
        if isinstance(item, Table):
            if len(item.rows) == 1 and len(item.columns) == 1:
                title = item.rows[0].cells[0].text.strip()
                if in_section:
                    if current_group is not None:
                        flush_pending(current_group[1])
                    break
                if any(t in title for t in PROCEDURE_SECTION_TITLES):
                    in_section = True
                    current_group = ["(ทั่วไป)", []]
                    groups.append(current_group)
            continue
        if not in_section:
            continue

        text = item.text.strip()
        images = _image_parts(item, document)
        is_caption = bool(FIGURE_RE.search(text))

        if is_caption:
            pending_captions.extend(_split_captions(text))
            continue
        if images:
            pending_images.extend(images)
            continue
        if not text:
            continue
        if item.style.name == "Normal":
            # A new maintenance-interval heading, e.g. "การตรวจสอบรอบ 1 สัปดาห์"
            # Resolve any pending figures against the group that's ending
            # before starting the new one.
            if current_group is not None:
                flush_pending(current_group[1])
            current_group = [text, []]
            groups.append(current_group)
            continue

        # Otherwise: a real step line.
        if current_group is None:
            current_group = ["(ทั่วไป)", []]
            groups.append(current_group)
        current_group[1].append(Step(group=current_group[0], text=text))

    return groups


def _trigrams(text):
    text = re.sub(r"\s+", "", text).lower()
    if len(text) < 3:
        return {text} if text else set()
    return {text[i:i + 3] for i in range(len(text) - 2)}


def _best_match(caption, steps):
    """Pick the step whose text contains the most of the caption's content
    (containment score relative to the short caption, not a symmetric
    ratio - a short caption is expected to be a fragment of a longer step)."""
    if not steps:
        return None
    caption_grams = _trigrams(caption)
    if not caption_grams:
        return None
    best, best_score = None, -1.0
    for step in steps:
        step_grams = _trigrams(step.text)
        score = len(caption_grams & step_grams) / len(caption_grams)
        if score > best_score:
            best, best_score = step, score
    return best


def match_equipment(steps, equipment_list):
    """Fill in step.equipment using concept-based matching against the
    document's own (real) equipment list."""
    for step in steps:
        desc = step.text
        matched = []
        for _concept, step_pattern, equip_keywords in CONCEPT_RULES:
            if re.search(step_pattern, desc):
                for equip_name in equipment_list:
                    if any(kw in equip_name for kw in equip_keywords) and equip_name not in matched:
                        matched.append(equip_name)
        if not matched and re.search(r"สายตา|สัมผัส", desc):
            matched.append("ตรวจสอบด้วยสายตา/สัมผัส (ไม่ใช้เครื่องมือ)")
        step.equipment = " + ".join(matched)


def process_wi_document(path):
    """
    Parse one *_system.docx WI file and return:
      equipment_list: list[str]
      groups: list[(group_title, list[Step])]
    """
    document = docx.Document(path)
    equipment_list = extract_equipment(document)
    groups = extract_steps(document)
    all_steps = [s for _, steps in groups for s in steps]
    match_equipment(all_steps, equipment_list)
    return equipment_list, groups


# ---------------------------------------------------------------------------
# Output: build the "หัวข้อ 5" (step-by-step) table as an Excel workbook, one
# sheet per PM, with real embedded pictures in 'รูปภาพประกอบ' and the real
# matched equipment names in 'อุปกรณ์/เครื่องมือที่ใช้'.
# ---------------------------------------------------------------------------

HEADERS = ["ลำดับ", "ขั้นตอนการปฏิบัติงาน/ข้อควรระวัง", "รูปภาพประกอบ", "อุปกรณ์/เครื่องมือที่ใช้"]

IMG_CELL_WIDTH_PX = 220
IMG_ROW_HEIGHT_PT = 130


def _sheet_name(pm_name):
    name = re.sub(r"[\[\]:*?/\\]", "_", pm_name).strip()
    return name[:31] or "PM"


def add_pm_sheet(workbook, pm_name, equipment_list, groups):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ws = workbook.create_sheet(title=_sheet_name(pm_name))
    ws.append(HEADERS)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    border = Border(*(Side(style="thin"),) * 4)
    wrap = Alignment(vertical="center", horizontal="left", wrap_text=True)

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = IMG_CELL_WIDTH_PX / 7.0
    ws.column_dimensions["D"].width = 45

    row_idx = 1
    step_no = 1
    for group_title, steps in groups:
        row_idx += 1
        ws.cell(row=row_idx, column=1, value=group_title)
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(HEADERS))
        group_cell = ws.cell(row=row_idx, column=1)
        group_cell.font = Font(bold=True)
        group_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        for step in steps:
            row_idx += 1
            ws.cell(row=row_idx, column=1, value=step_no)
            ws.cell(row=row_idx, column=2, value=step.text)
            ws.cell(row=row_idx, column=4, value=step.equipment)
            for col in range(1, len(HEADERS) + 1):
                c = ws.cell(row=row_idx, column=col)
                c.border = border
                c.alignment = wrap

            if step.images:
                _embed_images(ws, row_idx, col=3, blobs=step.images)
            step_no += 1

    return ws


def _fit(width, height, max_w, max_h):
    scale = min(max_w / width, max_h / height, 1.0)
    return int(width * scale), int(height * scale)


def _embed_images(ws, row_idx, col, blobs):
    """Embed one or more real pictures into a single cell, stacking them
    vertically (with a OneCellAnchor + pixel offset) so multiple pictures on
    the same step don't fully overlap, and grow the row to fit them."""
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU
    import io

    max_img_w = IMG_CELL_WIDTH_PX - 10
    max_img_h = IMG_ROW_HEIGHT_PT * 1.6
    y_offset_px = 0
    for blob in blobs:
        img = XLImage(io.BytesIO(blob))
        w, h = _fit(img.width, img.height, max_img_w, max_img_h)
        img.width, img.height = w, h
        marker = AnchorMarker(col=col - 1, colOff=pixels_to_EMU(5), row=row_idx - 1, rowOff=pixels_to_EMU(y_offset_px))
        img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
        ws.add_image(img)
        y_offset_px += h + 6

    ws.row_dimensions[row_idx].height = max(IMG_ROW_HEIGHT_PT, y_offset_px * 0.75)


def build_pm_workbook(pm_results):
    """pm_results: dict[pm_name -> (equipment_list, groups)] -> openpyxl Workbook"""
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    for pm_name, (equipment_list, groups) in pm_results.items():
        add_pm_sheet(wb, pm_name, equipment_list, groups)
    return wb


def save_pm_workbook(paths, output_path="PM_Equipment_Tools_Revised.xlsx"):
    """
    paths: list of *_system.docx file paths, one per PM.
    Parses each, builds the workbook, and saves it to output_path.
    """
    pm_results = {}
    for path in paths:
        pm_name = re.sub(r"_system$", "", path.rsplit("/", 1)[-1].rsplit(".", 1)[0])
        pm_results[pm_name] = process_wi_document(path)

    wb = build_pm_workbook(pm_results)
    wb.save(output_path)
    return output_path
