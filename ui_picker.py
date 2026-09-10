"""หน้าต่างเลือกรูปจากโฟลเดอร์ ใส่ลงจุด <pic> — คลิกจุดทางซ้าย แล้วคลิกรูปทางขวา

คนจับคู่เองตอนเห็นรูปจริงกับชื่อขั้นตอนพร้อมกัน จึงไม่ต้องตั้งชื่อไฟล์ตามเลข
และไม่ต้องเติมเลขที่ marker ในเอกสาร — ไม่มีอะไรให้ sync ผิด

รูปที่ถูกใช้ไปแล้วดูจากการเทียบเนื้อรูปกับรูปที่ฝังอยู่ในเอกสารจริง ไม่ใช่จากไฟล์บันทึก
จึงบอกได้แม้รูปถูกใส่ไว้ตั้งแต่ก่อนมีฟีเจอร์นี้

เรียกผ่าน ui_word.py
"""

import io
import tkinter as tk
from pathlib import Path
from tkinter import ttk

from PIL import Image, ImageTk

BG = "#f7f7f5"
CARD = "#ffffff"
BLACK = "#1a1d21"
GREY = "#6b7280"
LINE = "#dcdcd6"
GREEN = "#2e7d32"
AMBER = "#e65100"
RED = "#c62828"
BLUE = "#2c5f8d"
BLUE_BG = "#e3edf6"
FONT = "Leelawadee UI"

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp"}
THUMB = (150, 110)
COLS = 3
# โหลดทีละก้อนคั่นด้วย after() — หน้าต่างเปิดใช้ได้ทันที ไม่ค้างรอรูปครบ
CHUNK = 6
DOC_CHUNK = 3

SIG_SIZE = (16, 16)
# วัดกับงานจริง: รูปเดียวกันได้ 0.15-0.49 คนละรูปที่ใกล้สุดได้ 10.8
# ตั้งไว้ 4 จึงห่างจากทั้งสองฝั่งพอ ไม่ต้องกลัวทั้ง false match และหลุด
SIG_TOL = 4.0


def signature(img):
    """ลายเซ็นรูป 16x16 เทา — เทียบว่ารูปเดียวกันไหมโดยไม่สนขนาด/การย่อ"""
    return list(img.convert("L").resize(SIG_SIZE, Image.BILINEAR).getdata())


def _distance(a, b):
    return sum(abs(x - y) for x, y in zip(a, b)) / len(a)


def _ellipsis(text, limit):
    return text if len(text) <= limit else text[:limit - 1] + "…"


class Picker:
    """tables = [(idx, label, จำนวนจุด)]
    load_cb(table_index) -> (slots, {slot_no: bytes รูปที่ฝังอยู่})
       slots = [(slot_no, ชื่อ, ข้อความกำกับรูป, มีรูปแล้วไหม)]
    place_cb(table_index, slot_no, path) -> (ok, ข้อความ)
    """

    def __init__(self, win, folder, tables, table_index, load_cb, place_cb):
        self.win = win
        self.folder = Path(folder)
        self.tables = list(tables)
        self.ti = table_index
        self.load_cb = load_cb
        self.place_cb = place_cb

        self.placed = 0
        self.slots = []
        self.rows = []
        self.sel = 0
        # PhotoImage ต้องถูกอ้างอิงไว้ ไม่งั้น GC เก็บแล้วรูปหายทั้งกริด
        self.thumbs = []
        self.cells = []
        self.used_labels = {}
        self.file_sigs = {}
        self.doc_sigs = {}
        self._loading_docs = False

        win.title("แทรกรูปจากโฟลเดอร์")
        win.geometry("980x660")
        win.minsize(820, 560)
        win.configure(bg=BG)
        ttk.Style().configure("Thai.TButton", font=(FONT, 9))
        ttk.Style().configure("Thai.TCombobox", font=(FONT, 10))

        self._build_head()
        self._build_foot()
        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=(8, 0))
        self._build_slots(body)
        self._build_grid(body)

        self._reload_table()
        self._build_cells()

    # ---------- ส่วนประกอบ ----------

    def _build_head(self):
        head = tk.Frame(self.win, bg=BG)
        head.pack(fill="x", padx=14, pady=(12, 0))
        # path ยาวๆ ตัดหัวทิ้ง เก็บชื่อโฟลเดอร์ท้ายสุดไว้ — ส่วนที่คนใช้ดูว่ามาถูกที่
        path = str(self.folder)
        if len(path) > 62:
            path = "…" + path[-61:]
        tk.Label(head, text=path, font=(FONT, 9), bg=BG, fg=GREY,
                 anchor="w").pack(side="left", fill="x", expand=True)
        self.count_lbl = tk.Label(head, text="", font=(FONT, 10, "bold"),
                                  bg=BG, fg=BLACK)
        self.count_lbl.pack(side="right")

        bar = tk.Frame(self.win, bg=BG)
        bar.pack(fill="x", padx=14, pady=(6, 0))
        tk.Label(bar, text="ตาราง:", font=(FONT, 10), bg=BG, fg=BLACK).pack(side="left")
        self.table_var = tk.StringVar()
        self.table_box = ttk.Combobox(bar, textvariable=self.table_var,
                                      state="readonly", font=(FONT, 10),
                                      style="Thai.TCombobox")
        self.table_box.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.table_box.config(values=[f"{lb}   ({n} จุด)" for _, lb, n in self.tables])
        self.table_box.bind("<<ComboboxSelected>>", self._on_table_pick)

    def _scroller(self, parent, width=None):
        """canvas + scrollbar + frame ข้างใน — คืน (canvas, inner)"""
        canvas = tk.Canvas(parent, bg=CARD, highlightthickness=0,
                           **({"width": width} if width else {}))
        sb = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=CARD)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cw = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return canvas, inner

    def _build_slots(self, body):
        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="y")

        tk.Label(left, text="จุดที่ต้องใส่รูป", font=(FONT, 10, "bold"),
                 bg=BG, fg=BLACK).pack(anchor="w", pady=(0, 4))

        wrap = tk.Frame(left, bg=CARD, highlightbackground=LINE, highlightthickness=1,
                        width=290)
        wrap.pack(fill="y", expand=True)
        wrap.pack_propagate(False)
        self.slot_canvas, self.slot_inner = self._scroller(wrap, width=288)

    def _build_grid(self, body):
        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))

        tk.Label(right, text="รูปในโฟลเดอร์ — คลิกเพื่อวางลงจุดที่เลือก",
                 font=(FONT, 10, "bold"), bg=BG, fg=BLACK).pack(anchor="w", pady=(0, 4))

        wrap = tk.Frame(right, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        wrap.pack(fill="both", expand=True)
        self.grid_canvas, self.grid_inner = self._scroller(wrap)

        # bind_all เพราะ Windows ส่ง wheel ไปที่ widget ที่ focus ไม่ใช่ใต้เมาส์
        # หน้าต่างนี้ modal อยู่แล้ว ฝั่งแม่ผูกคืนเองหลังปิด
        self.win.bind_all("<MouseWheel>", self._on_wheel)

    def _on_wheel(self, event):
        """เลือกฝั่งที่จะเลื่อนจากพิกัดเมาส์ในอีเวนต์

        ห้ามใช้ winfo_containing() ตรงนี้ — มันไล่ widget tree ทุกครั้งที่หมุน
        พอมีรูปเป็นร้อย การเลื่อนจะหน่วงจนใช้ไม่ได้
        """
        left = self.slot_canvas.winfo_rootx()
        inside = left <= event.x_root <= left + self.slot_canvas.winfo_width()
        canvas = self.slot_canvas if inside else self.grid_canvas
        canvas.yview_scroll(int(-event.delta / 120), "units")

    def _build_foot(self):
        foot = tk.Frame(self.win, bg=BG)
        foot.pack(side="bottom", fill="x", padx=14, pady=(6, 12))
        ttk.Button(foot, text="ปิด", style="Thai.TButton",
                   command=self.win.destroy).pack(side="right")
        self.status = tk.Label(foot, text="", font=(FONT, 9), bg=BG, fg=GREY,
                               anchor="w", wraplength=740, justify="left")
        self.status.pack(side="left", fill="x", expand=True)

    # ---------- ตาราง ----------

    def _on_table_pick(self, _event=None):
        i = self.table_box.current()
        if i < 0 or self.tables[i][0] == self.ti:
            return
        self.ti = self.tables[i][0]
        self._reload_table()
        self.status.config(text=f"เปลี่ยนไปตาราง: {self.tables[i][1][:44]}", fg=GREEN)

    def _reload_table(self):
        """อ่านจุดของตารางที่เลือก แล้วเริ่มเทียบรูปในเอกสารใหม่"""
        self.slots, blobs = self.load_cb(self.ti)
        self.sel = next((i for i, s in enumerate(self.slots) if not s[3]), 0)
        self.doc_sigs = {}
        self._doc_blobs = list(blobs.items())

        for i, (idx, _lb, _n) in enumerate(self.tables):
            if idx == self.ti:
                self.table_box.current(i)
                break

        self._build_slot_rows()
        self._refresh_slots()
        self._clear_marks()
        if self._doc_blobs and not self._loading_docs:
            self._loading_docs = True
            self.win.after(10, self._load_doc_sigs, 0)

    def _build_slot_rows(self):
        for w in self.slot_inner.winfo_children():
            w.destroy()
        self.rows = []
        for i, (no, name, caption, _has) in enumerate(self.slots):
            row = tk.Frame(self.slot_inner, bg=CARD, cursor="hand2")
            row.pack(fill="x")
            mark = tk.Label(row, text="○", font=(FONT, 11), bg=CARD, width=3)
            mark.pack(side="left", padx=(8, 0))

            texts = tk.Frame(row, bg=CARD)
            texts.pack(side="left", fill="x", expand=True, pady=4)
            nm = tk.Label(texts, text=f"{no}. {name}", font=(FONT, 10), bg=CARD,
                          fg=BLACK, anchor="w")
            nm.pack(fill="x")
            # ข้อความกำกับรูปในเอกสาร — บอกว่าจุดนี้ต้องเป็นรูปอะไร
            cap = tk.Label(texts, text=_ellipsis(caption or "", 36), font=(FONT, 8),
                           bg=CARD, fg=GREY, anchor="w", justify="left")
            cap.pack(fill="x")

            for w in (row, mark, texts, nm, cap):
                w.bind("<Button-1>", lambda e, k=i: self.select(k))
            tk.Frame(self.slot_inner, bg=LINE, height=1).pack(fill="x")
            self.rows.append((row, mark, texts, nm, cap))

    def _load_doc_sigs(self, start):
        """ถอดรูปที่ฝังในเอกสารมาทำลายเซ็น — ทีละก้อน ไม่ให้หน้าต่างค้าง"""
        if not self.win.winfo_exists():
            return

        for slot_no, blob in self._doc_blobs[start:start + DOC_CHUNK]:
            try:
                img = Image.open(io.BytesIO(blob))
                img.draft("L", (64, 64))
                self.doc_sigs[slot_no] = signature(img)
            except (OSError, ValueError):
                continue

        nxt = start + DOC_CHUNK
        if nxt < len(self._doc_blobs):
            self.win.after(1, self._load_doc_sigs, nxt)
        else:
            self._loading_docs = False
            self._mark_all()

    # ---------- รูปในโฟลเดอร์ ----------

    def _build_cells(self):
        """สร้างช่องเปล่าให้ครบก่อน แล้วค่อยทยอยใส่รูป — เปิดหน้าต่างแล้วใช้ได้เลย"""
        try:
            files = sorted(p for p in self.folder.iterdir()
                           if p.suffix.lower() in IMAGE_EXT)
        except OSError as e:
            self.status.config(text=f"อ่านโฟลเดอร์ไม่ได้: {e}", fg=RED)
            return

        if not files:
            self.status.config(text="โฟลเดอร์นี้ไม่มีไฟล์รูป (.png .jpg .jpeg .bmp)",
                               fg=RED)
            return

        for i, f in enumerate(files):
            self._add_cell(f, i)
        self.win.after(10, self._load_chunk, 0)

    def _add_cell(self, path, i):
        cell = tk.Frame(self.grid_inner, bg=CARD, cursor="hand2",
                        highlightbackground=LINE, highlightthickness=1)
        cell.grid(row=i // COLS, column=i % COLS, padx=6, pady=6, sticky="n")

        pic = tk.Label(cell, bg=CARD, width=20, height=6)
        pic.pack(padx=4, pady=(4, 0))
        name = tk.Label(cell, text=path.name[:22], font=(FONT, 8), bg=CARD, fg=GREY)
        name.pack()
        used = tk.Label(cell, text="", font=(FONT, 8, "bold"), bg=CARD, fg=GREEN)
        used.pack(pady=(0, 4))
        self.used_labels[str(path)] = used

        for w in (cell, pic, name, used):
            w.bind("<Button-1>", lambda e, p=path: self.place(p))
        self.cells.append((pic, path))

    def _load_chunk(self, start):
        if not self.win.winfo_exists():
            return

        for i in range(start, min(start + CHUNK, len(self.cells))):
            pic, path = self.cells[i]
            try:
                img = Image.open(path)
                # draft ให้ JPEG ถอดรหัสที่ความละเอียดต่ำตั้งแต่แรก เร็วกว่าย่อทีหลังมาก
                img.draft("RGB", THUMB)
                img.thumbnail(THUMB, Image.BILINEAR)
                photo = ImageTk.PhotoImage(img)
            except (OSError, ValueError):
                pic.config(text="เปิดไม่ได้", fg=RED)   # ไฟล์เสีย ข้ามไป ไม่ล้มทั้งหน้าต่าง
                continue
            self.thumbs.append(photo)
            pic.config(image=photo, width=photo.width(), height=photo.height())
            # ลายเซ็นคิดจาก thumbnail ที่ถอดรหัสไว้แล้ว จึงแทบไม่มีต้นทุนเพิ่ม
            self.file_sigs[str(path)] = signature(img)
            self._mark_one(str(path))

        nxt = start + CHUNK
        if nxt < len(self.cells):
            self.status.config(text=f"กำลังโหลดรูป {nxt}/{len(self.cells)}…", fg=GREY)
            self.win.after(1, self._load_chunk, nxt)
        else:
            self.status.config(text="คลิกจุดทางซ้าย แล้วคลิกรูปเพื่อวาง", fg=GREY)

    # ---------- ธง "ใช้แล้ว" ----------

    def _slot_of(self, path_key):
        """จุดที่รูปนี้ถูกใส่อยู่ — None ถ้าไม่ตรงกับรูปไหนในเอกสาร"""
        sig = self.file_sigs.get(path_key)
        if sig is None or not self.doc_sigs:
            return None
        slot_no, d = min(((k, _distance(sig, s)) for k, s in self.doc_sigs.items()),
                         key=lambda kv: kv[1])
        return slot_no if d < SIG_TOL else None

    def _mark_one(self, path_key):
        lbl = self.used_labels.get(path_key)
        if lbl is None:
            return
        slot_no = self._slot_of(path_key)
        name = next((nm for no, nm, _c, _h in self.slots if no == slot_no), None)
        lbl.config(text=f"✓ อยู่ที่จุด {name or slot_no}" if slot_no else "")

    def _mark_all(self):
        for key in self.used_labels:
            self._mark_one(key)

    def _clear_marks(self):
        for lbl in self.used_labels.values():
            lbl.config(text="")

    # ---------- การกระทำ ----------

    def select(self, i):
        self.sel = i
        self._refresh_slots()
        no, name, _caption, has = self.slots[i]
        if has:
            self.status.config(text=f"{no}. {name} มีรูปแล้ว — คลิกรูปใหม่จะ"
                                    f"แทนที่รูปเดิม (1 จุด = 1 รูป)", fg=AMBER)
        else:
            self.status.config(text=f"เลือกจุด {no}. {name} — คลิกรูปที่จะวาง", fg=GREY)

    def place(self, path):
        if not self.slots:
            return
        no, name, caption, had = self.slots[self.sel]

        ok, msg = self.place_cb(self.ti, no, str(path))
        if not ok:
            self.status.config(text=msg, fg=RED)
            return

        self.placed += 1
        self.slots[self.sel] = (no, name, caption, True)
        # อัปเดตลายเซ็นฝั่งเอกสารเอง ไม่ต้องอ่านไฟล์ .docx ใหม่ทั้งก้อน
        sig = self.file_sigs.get(str(path))
        if sig is not None:
            self.doc_sigs[no] = sig
        self._mark_all()

        verb = "แทนที่รูปเดิมที่" if had else "→"
        self.status.config(text=f"{path.name}  {verb}  {no}. {name}", fg=GREEN)

        # ไปจุดว่างถัดไป ถ้าหมดแล้ววนกลับหาจุดว่างที่ค้างอยู่ข้างบน
        nxt = next((i for i in range(self.sel + 1, len(self.slots))
                    if not self.slots[i][3]), None)
        if nxt is None:
            nxt = next((i for i, s in enumerate(self.slots) if not s[3]), self.sel)
        self.sel = nxt
        self._refresh_slots()

    def _refresh_slots(self):
        done = sum(1 for s in self.slots if s[3])
        self.count_lbl.config(text=f"ใส่รูปแล้ว {done} / {len(self.slots)}")

        for i, (row, mark, texts, nm, cap) in enumerate(self.rows):
            cur = (i == self.sel)
            bg = BLUE_BG if cur else CARD
            if self.slots[i][3]:
                mark.config(text="✓", fg=GREEN)
                nm.config(fg=GREY)
            else:
                mark.config(text="○", fg=LINE)
                nm.config(fg=BLACK)
            nm.config(font=(FONT, 10, "bold") if cur else (FONT, 10))
            for w in (row, mark, texts, nm, cap):
                w.config(bg=bg)

        self._scroll_to(self.sel)

    def _scroll_to(self, i):
        if not (0 <= i < len(self.rows)):
            return
        self.win.update_idletasks()
        total = max(self.slot_inner.winfo_height(), 1)
        y = self.rows[i][0].winfo_y()
        self.slot_canvas.yview_moveto(max(0, (y - 80) / total))


def open_picker(parent, folder, tables, table_index, load_cb, place_cb):
    """เปิดหน้าต่างเลือกรูปแบบ modal — คืน (จำนวนรูปที่วาง, ตารางที่ค้างไว้ตอนปิด)"""
    win = tk.Toplevel(parent)
    win.transient(parent)
    win.grab_set()
    picker = Picker(win, folder, tables, table_index, load_cb, place_cb)
    win.wait_window()
    return picker.placed, picker.ti
