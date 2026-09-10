"""เขียนภาพลงช่อง Actual Result ของ Validation Report

โครงเอกสาร: ตารางมีคอลัมน์ No / Description / Expected / Actual / Accept / Tester
แถวที่ช่อง No เป็นตัวเลข = ขั้นตอนจริง (1..15) แถวหัวตารางและแถวคำอธิบายข้าม
"""

import os
import shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Twips

import config

# margin ในเซลล์ที่ Word ใช้เป็น default เมื่อตารางไม่ได้กำหนดเอง (~0.19cm/ข้าง)
_DEFAULT_CELL_MARGIN_TWIPS = 108


def _cell_margins(table):
    """(ซ้าย, ขวา) margin ในเซลล์ หน่วย twips — จากตารางถ้ากำหนด ไม่งั้น default"""
    left = right = _DEFAULT_CELL_MARGIN_TWIPS
    tblPr = table._tbl.find(qn("w:tblPr"))
    mar = tblPr.find(qn("w:tblCellMar")) if tblPr is not None else None
    if mar is not None:
        for side, ref in (("left", "left"), ("right", "right")):
            el = mar.find(qn(f"w:{side}"))
            if el is not None and el.get(qn("w:w")) is not None:
                val = int(el.get(qn("w:w")))
                if side == "left":
                    left = val
                else:
                    right = val
    return left, right


def fit_width(table, cell):
    """ความกว้างรูปที่พอดีคอลัมน์ของ cell นี้ — คืน docx width หรือ None ถ้าคำนวณไม่ได้

    อ่านความกว้างคอลัมน์จริงจาก tblGrid ลบ margin สองข้าง เผื่อ 1 twip กันปัดเกิน
    None = ให้ผู้เรียก fallback ไปใช้ค่าคงที่เดิม
    """
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is None:
        return None
    cols = grid.findall(qn("w:gridCol"))
    # grid_offset = index ใน tblGrid จริง ถูกต้องแม้มี merged cell
    ci = getattr(cell._tc, "grid_offset", None)
    if ci is None or ci >= len(cols):
        return None
    col_twips = int(cols[ci].get(qn("w:w")))
    # เต็มความกว้างคอลัมน์ — ผู้เรียก (place_image_full) ตั้ง indent ลบชดเชย
    # margin ซ้าย ให้รูปชิดเส้นคอลัมน์ทั้งสองข้าง ไม่เหลือช่องว่าง
    return Twips(col_twips)


def place_full_width(para, image_path, table, cell):
    """วางรูปเต็มความกว้างคอลัมน์ ชิดเส้นทั้งสองข้าง ในย่อหน้า para

    รูปกว้าง = คอลัมน์เต็ม แล้วดัน indent ย่อหน้าเป็นลบเท่า margin ซ้าย
    ให้รูปเริ่มชิดเส้นคอลัมน์ซ้าย (ไม่เว้น margin) จบชิดเส้นขวาพอดี
    fit_width คืน None (อ่านคอลัมน์ไม่ได้) -> ใช้ค่าคงที่เดิม ไม่ตั้ง indent

    ใช้ run แรกของ para (เติมให้ถ้ายังไม่มี) — โหมด pic เคลียร์ text ใน run
    เดิมไว้แล้วต้องวางลง run นั้น ไม่ใช่สร้างใหม่
    """
    if not para.runs:
        para.add_run()
    width = fit_width(table, cell)
    if width is None:
        para.runs[0].add_picture(str(image_path), width=Cm(config.WORD_IMAGE_CM))
        return
    left, _right = _cell_margins(table)
    para.paragraph_format.left_indent = Twips(-left)
    para.runs[0].add_picture(str(image_path), width=width)

class NumberingTracker:
    """ช่วยอ่านตัวเลขจาก Auto-numbering list ของ Word (เช่น 1.1) ที่ python-docx อ่านไม่ได้"""
    def __init__(self, doc):
        self.nums = {} 
        self.abstract_nums = {} 
        self.state = {}
        self.p_to_num = {} 
        
        self._parse_numbering(doc)
        if self.nums:
            self._traverse_doc(doc)
            
    def _parse_numbering(self, doc):
        try:
            numbering_part = doc.part.numbering_part
        except NotImplementedError:
            return
            
        if not numbering_part:
            return
            
        element = numbering_part.element
        ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        
        for num in element.findall(f'.//{ns_w}num'):
            numId = num.get(f'{ns_w}numId')
            absNum = num.find(f'.//{ns_w}abstractNumId')
            if absNum is not None:
                self.nums[numId] = absNum.get(f'{ns_w}val')
                
        for abs_num in element.findall(f'.//{ns_w}abstractNum'):
            absNumId = abs_num.get(f'{ns_w}abstractNumId')
            levels = {}
            for lvl in abs_num.findall(f'.//{ns_w}lvl'):
                ilvl = lvl.get(f'{ns_w}ilvl')
                start = lvl.find(f'.//{ns_w}start')
                start_val = int(start.get(f'{ns_w}val')) if start is not None else 1
                
                fmt = lvl.find(f'.//{ns_w}numFmt')
                fmt_val = fmt.get(f'{ns_w}val') if fmt is not None else 'decimal'
                
                text = lvl.find(f'.//{ns_w}lvlText')
                text_val = text.get(f'{ns_w}val') if text is not None else ''
                
                levels[ilvl] = {'start': start_val, 'fmt': fmt_val, 'text': text_val}
            self.abstract_nums[absNumId] = levels
            self.state[absNumId] = {}

    def _traverse_doc(self, doc):
        ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        # ค้นหา <w:p> ทั้งหมดในเอกสารตามลำดับ เพื่อจะได้นับเลขได้ถูกต้อง
        for p in doc.element.xpath('.//w:p'):
            numPr = p.find(f'.//{ns_w}numPr')
            if numPr is None:
                continue
                
            numId_el = numPr.find(f'.//{ns_w}numId')
            ilvl_el = numPr.find(f'.//{ns_w}ilvl')
            
            if numId_el is None:
                continue
                
            numId = numId_el.get(f'{ns_w}val')
            ilvl = ilvl_el.get(f'{ns_w}val') if ilvl_el is not None else '0'
            
            if numId not in self.nums:
                continue
                
            absNumId = self.nums[numId]
            if absNumId not in self.abstract_nums or ilvl not in self.abstract_nums[absNumId]:
                continue
                
            lvl_info = self.abstract_nums[absNumId][ilvl]
            
            st = self.state[absNumId]
            ilvl_int = int(ilvl)
            
            if ilvl not in st:
                st[ilvl] = lvl_info['start']
            else:
                st[ilvl] += 1
                
            # รีเซ็ตเลขที่อยู่ระดับลึกกว่า
            for k in list(st.keys()):
                if int(k) > ilvl_int:
                    del st[k]
                    
            if lvl_info['fmt'] == 'decimal':
                txt = lvl_info['text']
                for i in range(ilvl_int + 1):
                    lvl_str = str(i)
                    val = st.get(lvl_str, self.abstract_nums[absNumId].get(lvl_str, {}).get('start', 1))
                    txt = txt.replace(f'%{i+1}', str(val))
                # key ด้วย XML string ไม่ใช่ตัว element — lxml สร้าง proxy ใหม่
                # ทุกครั้งที่ access ทำให้ p (จาก xpath) != paragraph._p (จาก
                # python-docx) แม้เป็น node เดียวกัน dict lookup จึงพลาดเสมอ
                self.p_to_num[str(p.xml)] = txt

    def get_numbering(self, paragraph):
        return self.p_to_num.get(str(paragraph._p.xml), "")

    def cell_number(self, cell):
        """เลขขั้นตอนของแถวจากช่อง No. — text ถ้ามี ไม่งั้นดึงจาก auto-numbering

        คืน string ที่ตัด '.' ท้ายแล้ว (เช่น '1.1') — '' ถ้าไม่มีเลข
        """
        no = cell.text.strip()
        if not no:
            for para in cell.paragraphs:
                num = self.get_numbering(para)
                if num:
                    no = num
                    break
        return no.rstrip(".")

ACTUAL_HEADER = "actual result"


def ensure_working_copy(template):
    """คัดลอกต้นฉบับเป็นไฟล์ทำงาน ครั้งแรกครั้งเดียว — ไม่แตะต้นฉบับเลย"""
    template = Path(template)
    if not template.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ต้นฉบับ: {template}")
    working = config.word_working(template)
    if not working.exists():
        shutil.copy2(template, working)
    return working


def actual_column(table):
    """หา index ของคอลัมน์ Actual Result — None ถ้าตารางนี้ไม่ใช่ตารางผลทดสอบ"""
    if not table.rows:
        return None
    for i, cell in enumerate(table.rows[0].cells):
        if ACTUAL_HEADER in cell.text.strip().lower():
            return i
    return None


def _run_is_bold(run):
    """ตัวหนาจริงไหม — ดูทั้งที่ run และที่ style

    run.bold มักเป็น None เพราะตัวหนามาจาก character style (เช่น 'Strong')
    ไม่ได้ตั้งที่ run ตรงๆ ต้องไล่ base_style ขึ้นไปด้วย
    """
    if run.bold is not None:
        return run.bold
    style = run.style
    while style is not None:
        if style.font.bold is not None:
            return style.font.bold
        style = style.base_style
    return False


def _step_name(cell, number):
    """ชื่อขั้นตอนจากช่อง Description/Action

    ถ้าชื่อขั้นตอนทำตัวหนา (คำอธิบายตัวบาง) — เอาตัวหนาช่วงแรก
    ถ้าไม่มีตัวหนา และมีเครื่องหมาย – คั่น — เอาส่วนหน้า –
    ถ้าไม่มีทั้งคู่ (ทั้งช่องเป็นประโยคยาว) — โชว์เลขขั้นตอนแทน
    """
    bold = []
    for para in cell.paragraphs:
        for run in para.runs:
            if not run.text.strip():
                continue
            if _run_is_bold(run):
                bold.append(run.text.strip())
            elif bold:
                # เจอตัวบางหลังตัวหนาแล้ว = จบชื่อขั้นตอน
                return " ".join(bold)
    if bold:
        return " ".join(bold)

    desc = cell.text.strip().replace("\n", " ")
    if "–" in desc:
        return desc.split("–")[0].strip()
    # ไม่มีชื่อแยก — ประโยคยาวทั้งช่อง เอามาเป็นชื่อไม่ได้ โชว์เลขแทน
    return number


def _is_step_number(text):
    """ช่อง No เป็นเลขขั้นตอนไหม — รับ 1  1.  1.1  1.1.1

    แถวหัวตาราง ('No') และแถวหัวข้อกลุ่ม ('Verification 1.1: ...') จะไม่ผ่าน
    """
    t = text.strip().rstrip(".")
    if not t:
        return False
    return all(part.isdigit() for part in t.split("."))


def step_rows(table, tracker=None):
    """คืน [(เลขขั้นตอน, index แถว, ชื่อขั้นตอน)] — อ่านจากไฟล์ ไม่ hardcode

    เลขขั้นตอนเก็บเป็น string เพราะมีทั้ง '1' และ '1.1' '1.1.1'
    """
    steps = []
    for ri, row in enumerate(table.rows):
        cell = row.cells[0]
        no = tracker.cell_number(cell) if tracker else cell.text.strip().rstrip(".")
        if not _is_step_number(no):
            continue
        steps.append((no, ri, _step_name(row.cells[1], no)))
    return steps


# คำที่ไม่ใช่ชื่อตาราง — เป็นหัวข้อ/คำอธิบายมาตรฐาน ข้ามไปหาบรรทัดที่มีเนื้อหาจริง
_LABEL_SKIP = ("acceptance criteria", "all verification", "expected result")


def _pick_label(recent_lines):
    """เลือกชื่อตารางจากบรรทัดก่อนหน้า — ข้ามหัวข้อ/คำอธิบายมาตรฐาน

    เอาบรรทัดล่าสุดที่ 'ไม่ใช่' คำมาตรฐาน (ไล่จากใกล้ตารางขึ้นไป)
    """
    for line in reversed(recent_lines):
        low = line.lower()
        if any(skip in low for skip in _LABEL_SKIP):
            continue
        return line
    return recent_lines[-1] if recent_lines else ""


def find_tables(template):
    """ตารางที่ใส่รูปได้ — คืน [(index, ป้ายชื่อ, จำนวนขั้นตอน)]

    ป้ายชื่อดึงจากข้อความก่อนตาราง (เช่น "Test result for index No.1")
    เพื่อให้เลือกได้โดยไม่ต้องรู้เลข index
    """
    from docx.oxml.ns import qn

    doc = Document(str(template))
    # ไล่ body ตามลำดับจริง เพื่อจับคู่ข้อความ -> ตารางที่ตามมา
    # เก็บหลายบรรทัดก่อนตาราง เพราะบางเอกสารมี "Acceptance Criteria:" คั่น
    # ระหว่างชื่อจริงกับตาราง — เอาบรรทัดสุดท้ายจะได้คำอธิบายซ้ำๆ ไม่ใช่ชื่อ
    labels = {}
    recent = []
    ti = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            t = "".join(n.text or "" for n in child.iter(qn("w:t"))).strip()
            if t:
                recent.append(t)
                if len(recent) > 4:
                    recent.pop(0)
        elif child.tag == qn("w:tbl"):
            labels[ti] = _pick_label(recent)
            ti += 1

    out = []
    tracker = NumberingTracker(doc)
    for i, table in enumerate(doc.tables):
        if actual_column(table) is None:
            continue
        steps = step_rows(table, tracker)
        if not steps:
            continue
        label = labels.get(i, "") or f"ตารางที่ {i}"
        out.append((i, label, len(steps)))
    return out


def load_steps(template, table_index):
    """รายการขั้นตอนของตารางที่ระบุ"""
    doc = Document(str(ensure_working_copy(template)))
    table = doc.tables[table_index]
    if actual_column(table) is None:
        raise ValueError(f"ตาราง #{table_index} ไม่มีคอลัมน์ Actual Result")
    tracker = NumberingTracker(doc)
    return step_rows(table, tracker)


def count_images(template, table_index):
    """นับรูปที่มีอยู่จริงในช่อง Actual Result ของแต่ละขั้นตอน

    คืน {เลขขั้นตอน: จำนวนรูป}

    อ่านจากไฟล์ ไม่ใช่จำไว้ในโปรแกรม — ปิดเปิดใหม่จึงรู้ว่าทำอะไรไปแล้ว
    และตรงกับความจริงเสมอแม้มีคนไปแก้เอกสารเอง
    """
    working = config.word_working(template)
    if not working.exists():
        return {}

    doc = Document(str(working))
    table = doc.tables[table_index]
    col = actual_column(table)
    if col is None:
        return {}

    counts = {}
    tracker = NumberingTracker(doc)
    for no, ri, _ in step_rows(table, tracker):
        xml = table.rows[ri].cells[col]._tc.xml
        counts[no] = xml.count("<wp:extent")
    return counts


def open_at_table(template, table_index):
    """เปิดไฟล์ทำงานใน Word แล้วเลื่อนไปหัวเอกสาร

    คืน (สำเร็จไหม, ข้อความ) — ถ้าสั่ง Word ไม่ได้ ยังเปิดไฟล์ให้ตามปกติ
    """
    working = config.word_working(template)
    if not working.exists():
        return False, "ยังไม่มีไฟล์"

    try:
        import win32com.client as com
    except ImportError:
        os.startfile(working)
        return True, "เปิดไฟล์แล้ว"

    try:
        app = com.Dispatch("Word.Application")
        app.Visible = True
        doc = app.Documents.Open(str(working))
        # เลื่อนไปหัวเอกสาร — wdStory=6, wdMove=0
        app.Selection.HomeKey(Unit=6, Extend=0)
        app.Activate()
        return True, "เปิด Word แล้ว"
    except Exception as e:
        # Word อาจถูกปิด/ถูกล็อก/COM พัง — เปิดไฟล์ธรรมดาแทน ดีกว่าไม่เปิดเลย
        try:
            os.startfile(working)
        except OSError:
            return False, f"เปิดไฟล์ไม่ได้ ({type(e).__name__})"
        return True, f"เปิดไฟล์แล้ว (เลื่อนไปตารางไม่ได้: {type(e).__name__})"


def insert_image(template, table_index, step_no, image_path):
    """ใส่ภาพลงช่อง Actual Result ของขั้นตอนที่ระบุ

    คืน (สำเร็จไหม, ข้อความ)
    """
    path = ensure_working_copy(template)
    doc = Document(str(path))
    table = doc.tables[table_index]

    col = actual_column(table)
    if col is None:
        return False, f"ตาราง #{table_index} ไม่มีคอลัมน์ Actual Result"

    tracker = NumberingTracker(doc)
    match = [(n, ri, nm) for n, ri, nm in step_rows(table, tracker) if n == step_no]
    if not match:
        return False, f"ไม่พบขั้นตอนที่ {step_no} ในตาราง"
    _, row_idx, name = match[0]

    cell = table.rows[row_idx].cells[col]
    # ย่อหน้าแรกของช่องมักว่างอยู่แล้ว ใช้ต่อได้ ไม่งั้นสร้างใหม่
    para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    if para.text.strip() or "graphic" in para._p.xml:
        para = cell.add_paragraph()
    place_full_width(para, image_path, table, cell)

    doc.save(str(path))
    return True, name
