"""หน้าต่างสำหรับลง Validation Report — รายการขั้นตอนคือหน้าจอหลัก

เรียกผ่าน main.py
"""

import datetime
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import keyboard

import capture
import config
import pic_target
import ui_picker
import word_target

BG = "#f7f7f5"
CARD = "#ffffff"
BLACK = "#1a1d21"
GREY = "#6b7280"
LINE = "#dcdcd6"
GREEN = "#2e7d32"
GREEN_BG = "#e8f3e9"
AMBER = "#e65100"
AMBER_BG = "#fdf3e7"
RED = "#c62828"
BLUE = "#2c5f8d"
BLUE_BG = "#e3edf6"
FONT = "Leelawadee UI"

DONE, TODO = "done", "todo"


def _ellipsis(text, limit):
    """ตัดข้อความยาวๆ กันดันของข้างๆ ตกขอบ"""
    return text if len(text) <= limit else text[:limit - 1] + "…"


def choose_document(parent=None):
    """ให้ผู้ใช้เลือกไฟล์ Word และตาราง — คืน (template, table_index) หรือ None"""
    if parent is not None:
        # ดัน dialog มาหน้าสุด ไม่งั้นไปโผล่หลังหน้าต่างอื่น หาไม่เจอ
        parent.lift()
        parent.attributes("-topmost", True)
        parent.update()
        parent.attributes("-topmost", False)

    template = filedialog.askopenfilename(
        parent=parent,
        title="เลือกไฟล์ Word",
        filetypes=[("เอกสาร Word", "*.docx"), ("ทุกไฟล์", "*.*")],
        initialdir=str(config.word_template().parent) if config.word_template()
        else str(Path.home() / "Downloads"),
    )
    if not template:
        return None

    try:
        # เอกสารมี <pic> = โหมด pic, ไม่มี = โหมดขั้นตอนเดิม
        if pic_target.has_pic_markers(template):
            tables = pic_target.find_pic_tables(template)
        else:
            tables = word_target.find_tables(template)
    except Exception as e:
        messagebox.showerror("เปิดไฟล์ไม่ได้",
                             f"อ่านไฟล์นี้ไม่ได้\n\n{type(e).__name__}: {e}",
                             parent=parent)
        return None

    if not tables:
        messagebox.showerror(
            "ใช้ไฟล์นี้ไม่ได้",
            "ไม่พบตารางที่ใส่รูปได้\n\n"
            "ไฟล์ต้องมี <pic> ในตาราง หรือเป็นตารางแบบ Validation Report "
            "(มีคอลัมน์ Actual Result และขั้นตอนเป็นตัวเลข)",
            parent=parent)
        return None

    if len(tables) == 1:
        return template, tables[0][0]
    return _pick_table(parent, template, tables)


def _pick_table(parent, template, tables):
    """เลือกตารางจากป้ายชื่อจริงในเอกสาร ไม่ใช่เลข index"""
    win = tk.Toplevel(parent) if parent else tk.Tk()
    win.title("เลือกตาราง")
    win.configure(bg=BG)
    win.resizable(False, False)
    if parent:
        win.transient(parent)
        win.grab_set()

    tk.Label(win, text=Path(template).name, font=(FONT, 9), bg=BG, fg=GREY,
             wraplength=460).pack(anchor="w", padx=16, pady=(14, 0))
    tk.Label(win, text="เอกสารนี้มีหลายตาราง — จะใส่รูปลงตารางไหน",
             font=(FONT, 11, "bold"), bg=BG, fg=BLACK).pack(anchor="w", padx=16, pady=(4, 10))

    chosen = {"idx": tables[0][0]}
    var = tk.IntVar(value=tables[0][0])

    box = tk.Frame(win, bg=CARD, highlightbackground=LINE, highlightthickness=1)
    box.pack(fill="x", padx=16)
    for idx, label, n in tables:
        tk.Radiobutton(box, text=f"{label}   ({n} ขั้นตอน)", variable=var, value=idx,
                       font=(FONT, 10), bg=CARD, fg=BLACK, anchor="w",
                       activebackground=CARD, selectcolor=CARD,
                       highlightthickness=0).pack(fill="x", padx=8, pady=4)

    def ok():
        chosen["idx"] = var.get()
        win.destroy()

    def cancel():
        chosen["idx"] = None
        win.destroy()

    bar = tk.Frame(win, bg=BG)
    bar.pack(fill="x", padx=16, pady=12)
    ttk.Button(bar, text="ตกลง", style="Thai.TButton", command=ok).pack(side="right")
    ttk.Button(bar, text="ยกเลิก", style="Thai.TButton",
               command=cancel).pack(side="right", padx=(0, 6))

    win.protocol("WM_DELETE_WINDOW", cancel)
    win.wait_window()
    if chosen["idx"] is None:
        return None
    return template, chosen["idx"]


class App:
    def __init__(self, root, template, table_index):
        self.root = root
        self.template = Path(template)
        self.table_index = table_index
        self._last_shot = 0.0                # กันกด F9 รัวเกินไป
        # โหมด pic = เอกสารมี <pic> marker / โหมด step = ตารางขั้นตอนแบบเดิม
        self.pic_mode = pic_target.has_pic_markers(self.template)

        root.title("AutoCapture — Validation Report")
        root.geometry("560x730")
        root.minsize(520, 640)
        root.configure(bg=BG)

        # ปุ่มกับสถานะจองที่ล่างสุดก่อน ไม่งั้นรายการยาวๆ ดันตกขอบ
        self._build_header(root)
        self._build_now(root)
        self._build_footer(root)
        self.list_wrap = None
        self._load_document()

        keyboard.add_hotkey(config.hotkey(), self.on_hotkey, suppress=False)

    def _fill_tables(self):
        """ใส่ตารางทั้งหมดของเอกสารนี้ลง dropdown

        เก็บ self.tables ไว้ใช้ซ้ำ — หา table อ่านไฟล์ทั้งก้อน
        เรียกทุกครั้งที่ refresh จะช้า
        """
        if self.pic_mode:
            self.tables = pic_target.find_pic_tables(self.template)
            unit = "รูป"
        else:
            self.tables = word_target.find_tables(self.template)
            unit = "ขั้นตอน"
        labels = [f"{lb}   ({n} {unit})" for _, lb, n in self.tables]
        self.table_box.config(values=labels)
        for i, (idx, _, _) in enumerate(self.tables):
            if idx == self.table_index:
                self.table_box.current(i)
                break

    def _on_table_pick(self, _event=None):
        i = self.table_box.current()
        if i < 0 or self.tables[i][0] == self.table_index:
            return
        self.table_index = self.tables[i][0]
        config.save_settings(word_table_index=self.table_index)
        self._load_document()
        self.status.config(text=f"เปลี่ยนไปตาราง: {self.tables[i][1][:44]}", fg=GREEN)
        # ให้ focus กลับไปที่หน้าต่าง ไม่งั้น dropdown ยังคาอยู่ กด F9 อาจไม่เข้า
        self.root.focus_set()

    def _inserter(self, item_id):
        """คืน callback ที่รับ path รูป แล้วใส่ลง Word ตามโหมด"""
        tpl, ti = self.template, self.table_index
        if self.pic_mode:
            return lambda p: pic_target.place_image(tpl, ti, item_id, p)
        return lambda p: word_target.insert_image(tpl, ti, item_id, p)

    def _read_items(self):
        """อ่านรายการจากไฟล์ตามโหมด — คืน (steps, shots)

        steps = [(id, name)]  shots = [จำนวนรูปที่ใส่แล้ว]
        อ่านจากไฟล์จริง ปิดเปิดใหม่จึงรู้ว่าทำอะไรไปแล้ว
        """
        if self.pic_mode:
            slots = pic_target.load_slots(self.template, self.table_index)
            steps = [(no, name) for no, name, _total, _placed in slots]
            shots = [1 if placed else 0 for _no, _name, _total, placed in slots]
        else:
            rows = word_target.load_steps(self.template, self.table_index)
            counts = word_target.count_images(self.template, self.table_index)
            steps = [(no, name) for no, _ri, name in rows]
            shots = [counts.get(no, 0) for no, _name in steps]
        return steps, shots

    def _load_document(self):
        """อ่านรายการจากเอกสารที่เลือก แล้วสร้าง list ใหม่"""
        self.steps, self.shots = self._read_items()
        self.state = [DONE if n else TODO for n in self.shots]
        # ชี้จุดแรกที่ยังไม่มีรูป
        self.idx = next((i for i, n in enumerate(self.shots) if not n),
                        len(self.steps))
        self.rows = []

        self._fill_tables()
        if self.list_wrap is not None:
            self.list_wrap.destroy()
        self._build_list(self.root)
        self.refresh()

    def change_document(self):
        picked = choose_document(self.root)
        if not picked:
            return
        template, table_index = picked
        config.save_settings(word_template=str(template), word_table_index=table_index)
        self.template = Path(template)
        self.table_index = table_index
        try:
            self._load_document()
        except Exception as e:
            messagebox.showerror("เปิดเอกสารไม่ได้", f"{type(e).__name__}: {e}",
                                 parent=self.root)
            return
        self.status.config(text=f"เปลี่ยนเป็น {self.template.name}", fg=GREEN)

    # ---------- ส่วนประกอบ ----------

    def _build_header(self, root):
        style = ttk.Style()
        style.configure("Thai.TButton", font=(FONT, 9))

        # ปุ่มและตัวนับจองที่ก่อน (side="right") แล้วป้ายที่เหลือค่อยยืดเติม
        # ไม่งั้นชื่อไฟล์ยาวๆ ดันของพวกนี้ตกขอบ
        top = tk.Frame(root, bg=BG)
        top.pack(fill="x", padx=16, pady=(10, 0))
        ttk.Button(top, text="เปลี่ยนเอกสาร", style="Thai.TButton",
                   command=self.change_document).pack(side="right")
        self.file_lbl = tk.Label(top, text="", font=(FONT, 9), bg=BG, fg=GREY,
                                 anchor="w")
        self.file_lbl.pack(side="left", fill="x", expand=True)

        bar = tk.Frame(root, bg=BG)
        bar.pack(fill="x", padx=16, pady=(6, 0))
        self.count_lbl = tk.Label(bar, text="", font=(FONT, 10), bg=BG, fg=GREY)
        self.count_lbl.pack(side="right", padx=(8, 0))

        style = ttk.Style()
        style.configure("Thai.TCombobox", font=(FONT, 10))
        self.table_var = tk.StringVar()
        self.table_box = ttk.Combobox(bar, textvariable=self.table_var,
                                      state="readonly", font=(FONT, 10),
                                      style="Thai.TCombobox")
        self.table_box.pack(side="left", fill="x", expand=True)
        self.table_box.bind("<<ComboboxSelected>>", self._on_table_pick)

    def _build_now(self, root):
        """กล่องขั้นตอนปัจจุบัน — สิ่งเดียวที่ต้องอ่านตอนกด F9"""
        self.now = tk.Frame(root, bg=BLUE_BG, highlightbackground=BLUE,
                            highlightthickness=1)
        self.now.pack(fill="x", padx=16, pady=(10, 0))

        tk.Label(self.now, text="กำลังจะแคปลงขั้นตอนนี้", font=(FONT, 9),
                 bg=BLUE_BG, fg=BLUE).pack(anchor="w", padx=14, pady=(10, 0))

        line = tk.Frame(self.now, bg=BLUE_BG)
        line.pack(fill="x", padx=14, pady=(2, 12))
        self.now_no = tk.Label(line, text="", font=(FONT, 34, "bold"),
                               bg=BLUE_BG, fg=BLUE)
        self.now_no.pack(side="left")
        self.now_name = tk.Label(line, text="", font=(FONT, 15, "bold"),
                                 bg=BLUE_BG, fg=BLACK, wraplength=400,
                                 justify="left")
        self.now_name.pack(side="left", padx=(12, 0))

        self.hint = tk.Label(root, text=f"กด {config.hotkey().upper()}",
                             font=(FONT, 11, "bold"), bg=BG, fg=BLACK)
        self.hint.pack(pady=(8, 0))

    def _build_list(self, root):
        wrap = tk.Frame(root, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        wrap.pack(fill="both", expand=True, padx=16, pady=(10, 0))
        self.list_wrap = wrap

        canvas = tk.Canvas(wrap, bg=CARD, highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        self.inner = tk.Frame(canvas, bg=CARD)

        self.inner.bind("<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.canvas = canvas

        self._bind_wheel()

        for i, (no, name) in enumerate(self.steps):
            row = tk.Frame(self.inner, bg=CARD, cursor="hand2")
            row.pack(fill="x")

            mark = tk.Label(row, text="", font=(FONT, 11), bg=CARD, width=3)
            mark.pack(side="left", padx=(10, 0))
            num = tk.Label(row, text=f"{no}.", font=(FONT, 10), bg=CARD,
                           fg=GREY, width=3, anchor="w")
            num.pack(side="left")
            nm = tk.Label(row, text=name, font=(FONT, 10), bg=CARD,
                          fg=BLACK, anchor="w", padx=0)
            nm.pack(side="left", fill="x", expand=True, pady=5)

            for w in (row, mark, num, nm):
                w.bind("<Button-1>", lambda e, k=i: self.jump(k))

            tk.Frame(self.inner, bg=LINE, height=1).pack(fill="x")
            self.rows.append((row, mark, num, nm))

    def _bind_wheel(self):
        """ผูก wheel แบบ bind_all — Windows ส่ง wheel ไปที่ widget ที่ focus ไม่ใช่ใต้เมาส์

        หน้าต่างเลือกรูปยึด binding นี้ไปตอนเปิด ต้องผูกคืนหลังมันปิด
        """
        self.canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))

    def _build_footer(self, root):
        style = ttk.Style()
        style.configure("Thai.TButton", font=(FONT, 9))

        # ยึดขอบล่าง — รายการขั้นตอนที่ยาวจะไม่ดันปุ่มตกจอ
        btns = tk.Frame(root, bg=BG)
        btns.pack(side="bottom", fill="x", padx=16, pady=(0, 12))

        self.status = tk.Label(root, text="พร้อมทำงาน", font=(FONT, 9),
                               bg=BG, fg=GREY, wraplength=520, anchor="w",
                               justify="left")
        self.status.pack(side="bottom", fill="x", padx=16, pady=(8, 4))
        # เฉพาะโหมด <pic> — โหมด step ลบไม่ได้ว่าจะถอดรูปไหนถ้ามีหลายรูปในช่องเดียว
        self.revert_btn = None
        if self.pic_mode:
            self.revert_btn = ttk.Button(btns, text="คืนค่า <pic>", style="Thai.TButton",
                                         command=self.revert)
            self.revert_btn.pack(side="left")
            ttk.Button(btns, text="นำเข้าด้วยรูปภาพ", style="Thai.TButton",
                       command=self.import_from_folder).pack(side="left", padx=(6, 0))
        ttk.Button(btns, text="เปิด Word", style="Thai.TButton",
                   command=self.open_word).pack(side="right")

        # แถวตั้งปุ่มลัด — วางเหนือปุ่ม ยึดขอบล่างเช่นกัน
        hk_row = tk.Frame(root, bg=BG)
        hk_row.pack(side="bottom", fill="x", padx=16, pady=(0, 4))
        tk.Label(hk_row, text="ปุ่มลัด:", font=(FONT, 9), bg=BG, fg=GREY).pack(side="left")
        self.hotkey_var = tk.StringVar(value=config.hotkey())
        entry = tk.Entry(hk_row, textvariable=self.hotkey_var, font=(FONT, 9), width=14)
        entry.pack(side="left", padx=6)
        entry.bind("<Return>", lambda e: self._change_hotkey())
        ttk.Button(hk_row, text="ใช้ปุ่มนี้", style="Thai.TButton",
                   command=self._change_hotkey).pack(side="left")

    def _change_hotkey(self):
        new = config.set_hotkey(self.hotkey_var.get())
        if new is None:
            self.status.config(
                text="ปุ่มไม่ถูกต้อง — พิมพ์เช่น f9, f10, ctrl+alt+c", fg=RED)
            self.hotkey_var.set(config.hotkey())
            return
        # ถอดปุ่มเก่าก่อนผูกใหม่ ไม่งั้นทั้งสองปุ่มยิงพร้อมกัน
        keyboard.unhook_all_hotkeys()
        keyboard.add_hotkey(new, self.on_hotkey, suppress=False)
        self.hotkey_var.set(new)
        self.status.config(text=f"เปลี่ยนปุ่มลัดเป็น {new.upper()} แล้ว", fg=GREEN)
        self.refresh()   # อัปเดตข้อความ hint ให้ตรงปุ่มใหม่

    # ---------- สถานะ ----------

    def refresh(self):
        done = sum(1 for s in self.state if s == DONE)
        self.count_lbl.config(text=f"ใส่รูปแล้ว {done} / {len(self.steps)}")
        self.file_lbl.config(text=_ellipsis(self.template.name, 54))

        if self.idx >= len(self.steps):
            self.now.config(bg=GREEN_BG, highlightbackground=GREEN)
            for w in self.now.winfo_children():
                self._paint(w, GREEN_BG)
            self.now_no.config(text="✓", fg=GREEN, bg=GREEN_BG)
            self.now_name.config(text="ครบทุกจุดแล้ว", bg=GREEN_BG)
            self.hint.config(text="กดที่ขั้นตอนในรายการเพื่อกลับไปแก้")
        else:
            no, name = self.steps[self.idx]
            self.now.config(bg=BLUE_BG, highlightbackground=BLUE)
            for w in self.now.winfo_children():
                self._paint(w, BLUE_BG)
            self.now_no.config(text=str(no), fg=BLUE, bg=BLUE_BG)
            n = self.shots[self.idx]
            self.now_name.config(text=f"{name}   ({n} รูป)" if n else name, bg=BLUE_BG)
            self.hint.config(text=f"กด {config.hotkey().upper()} = ใส่รูปลงขั้นตอนนี้")

        if self.revert_btn is not None:
            has_image = self.idx < len(self.steps) and self.shots[self.idx] > 0
            self.revert_btn.config(state="normal" if has_image else "disabled")

        for i, (row, mark, num, nm) in enumerate(self.rows):
            st = self.state[i]
            cur = (i == self.idx)
            bg = BLUE_BG if cur else CARD
            if st == DONE:
                n = self.shots[i]
                mark.config(text=f"{n}✓" if n > 1 else "✓", fg=GREEN)
                nm.config(fg=GREY)
            else:
                mark.config(text="○", fg=LINE)
                nm.config(fg=BLACK)
            nm.config(font=(FONT, 10, "bold") if cur else (FONT, 10))
            for w in (row, mark, num, nm):
                w.config(bg=bg)

        self._scroll_to(self.idx)

    def _paint(self, widget, bg):
        try:
            widget.config(bg=bg)
        except tk.TclError:
            pass
        for c in widget.winfo_children():
            self._paint(c, bg)

    def _scroll_to(self, i):
        if not (0 <= i < len(self.rows)):
            return
        self.root.update_idletasks()
        total = max(self.inner.winfo_height(), 1)
        y = self.rows[i][0].winfo_y()
        self.canvas.yview_moveto(max(0, (y - 60) / total))

    # ---------- การกระทำ ----------

    def on_hotkey(self):
        # hotkey มาจากอีก thread — โยนเข้า main loop ก่อนแตะ widget
        self.root.after(0, self.shoot)

    def shoot(self):
        # กันกดรัวเกินไป
        now_t = time.monotonic()
        if now_t - self._last_shot < 0.4:
            return
        self._last_shot = now_t

        if self.idx >= len(self.steps):
            self.status.config(text="ครบทุกจุดแล้ว — กดที่ขั้นตอนในรายการถ้าต้องการแก้",
                               fg=AMBER)
            return

        no, name = self.steps[self.idx]
        result = capture.capture_to_word(self._inserter(no))
        now = datetime.datetime.now()

        if not result.ok:
            self.status.config(text=result.message, fg=RED)
            return

        # อ่านกลับจากไฟล์ ไม่บวกเอาเอง — เลขจะไม่มีวันเพี้ยนจากของจริง
        _steps, self.shots = self._read_items()
        self.state[self.idx] = DONE
        n = self.shots[self.idx]

        self.status.config(text=f"{now:%H:%M:%S}  ใส่ {no} ({name}) "
                                f"{'ครบ ' + str(n) + ' รูป' if n > 1 else 'แล้ว'}",
                           fg=GREEN)
        self.idx += 1
        self.refresh()

    def import_from_folder(self):
        """เลือกรูปที่มีอยู่แล้วจากโฟลเดอร์ มาใส่ทีละจุดในหน้าต่างเลือกรูป"""
        folder = filedialog.askdirectory(parent=self.root,
                                         title="เลือกโฟลเดอร์ที่เก็บรูป")
        if not folder:
            return

        n, ti = ui_picker.open_picker(self.root, folder, self.tables,
                                      self.table_index, self._picker_slots,
                                      self._place_file)
        self._bind_wheel()

        if ti != self.table_index:
            # ผู้ใช้สลับตารางในหน้าเลือกรูป — หน้าหลักต้องตามไปตารางเดียวกัน
            self.table_index = ti
            config.save_settings(word_table_index=ti)
            self._load_document()
        else:
            _steps, self.shots = self._read_items()
            self.state = [DONE if c else s for c, s in zip(self.shots, self.state)]
            self.idx = next((i for i, c in enumerate(self.shots) if not c),
                            len(self.steps))
            self.refresh()
        if n:
            self.status.config(text=f"ใส่รูปจากโฟลเดอร์แล้ว {n} รูป", fg=GREEN)

    def _picker_slots(self, table_index):
        """จุดของตารางนั้น + รูปที่ฝังอยู่จริง ให้หน้าเลือกรูปเทียบว่าใช้ไฟล์ไหนไปแล้ว"""
        slots = pic_target.load_slots(self.template, table_index)
        return ([(no, name, placed) for no, name, _total, placed in slots],
                pic_target.slot_images(self.template, table_index))

    def _place_file(self, table_index, slot_no, src):
        """ย่อรูปเก็บเข้า captures/ แล้ววางลง slot — คืน (สำเร็จ, ข้อความ)"""
        try:
            path = capture.import_image(src, max_width=config.WORD_IMAGE_WIDTH)
        except OSError:
            return False, f"เปิดไฟล์รูปไม่ได้: {Path(src).name}"

        try:
            return pic_target.place_image(self.template, table_index,
                                          slot_no, str(path))
        except PermissionError:
            return False, "Word เปิดไฟล์ค้างอยู่ — ปิดแล้วลองใหม่ (รูปเก็บไว้แล้ว ไม่หาย)"
        except Exception as e:
            return False, f"ใส่ Word ไม่ได้ ({type(e).__name__}) — รูปเก็บใน captures ไม่หาย"

    def revert(self):
        """เอารูปที่จุดปัจจุบันออก คืนเป็น <pic> — ใช้ตอนแคปผิดแล้วอยากแคปใหม่"""
        if self.idx >= len(self.steps) or not self.shots[self.idx]:
            return
        no, name = self.steps[self.idx]
        ok, msg = pic_target.revert_image(self.template, self.table_index, no)
        now = datetime.datetime.now()

        if not ok:
            self.status.config(text=msg, fg=RED)
            return

        _steps, self.shots = self._read_items()
        self.state[self.idx] = TODO
        self.status.config(
            text=f"{now:%H:%M:%S}  เอารูปออกจาก {no} ({name}) แล้ว — กด "
                 f"{config.hotkey().upper()} เพื่อแคปใหม่", fg=AMBER)
        self.refresh()

    def jump(self, i):
        self.idx = i
        self.refresh()
        no, name = self.steps[i]
        n = self.shots[i]
        if n and self.pic_mode:
            # โหมด <pic> 1 จุด = 1 รูป การวางทับคือแทนที่ ไม่ใช่ต่อท้าย
            self.status.config(
                text=f"{no} ({name}) มีรูปแล้ว — แคปอีกจะแทนที่รูปเดิม "
                     f"(กด 'คืนค่า <pic>' ถ้าอยากเอาออกเฉยๆ)", fg=AMBER)
        elif n:
            self.status.config(
                text=f"{no} ({name}) มี {n} รูปแล้ว — กดอีกจะเพิ่มรูปที่ {n + 1} ต่อท้าย",
                fg=AMBER)
        else:
            self.status.config(text=f"ไปที่ขั้น {no} ({name})", fg=GREY)

    def open_word(self):
        # ไฟล์ทำงานถูกสร้างตั้งแต่เปิดโปรแกรมแล้ว (ensure_working_copy)
        # ไม่ต้องเช็คว่ามีไหม
        self.status.config(text="กำลังเปิด Word…", fg=GREY)
        self.root.update_idletasks()
        ok, msg = word_target.open_at_table(self.template, self.table_index)
        self.status.config(text=msg, fg=GREEN if ok else RED)



def run():
    root = tk.Tk()
    ttk.Style().configure("Thai.TButton", font=(FONT, 9))

    template = config.word_template()
    table_index = config.word_table_index()

    # ไฟล์ที่จำไว้อาจถูกย้าย/ลบ — ถามใหม่แทนที่จะพัง
    if not template or not template.exists():
        # ห้าม withdraw ตรงนี้ — หน้าต่างที่ซ่อนอยู่ทำให้ dialog ไปโผล่หลังจอ
        # มองไม่เห็น กดไม่ได้ และไม่มีปุ่มใน taskbar = โปรแกรมค้างเงียบ
        root.title("AutoCapture")
        root.geometry("360x120+80+80")
        root.configure(bg=BG)
        tk.Label(root, text="เลือกไฟล์ Word ที่จะใส่รูป", font=(FONT, 11, "bold"),
                 bg=BG, fg=BLACK).pack(pady=(24, 4))
        tk.Label(root, text="กำลังเปิดหน้าต่างเลือกไฟล์…", font=(FONT, 9),
                 bg=BG, fg=GREY).pack()
        root.update()

        picked = choose_document(root)
        if not picked:
            root.destroy()      # ไม่เลือกไฟล์ = ไม่มีอะไรให้ทำ ปิดไปเลย
            return
        template, table_index = picked
        config.save_settings(word_template=str(template), word_table_index=table_index)

        for w in root.winfo_children():
            w.destroy()

    def close():
        # ต้องถอด hotkey เอง — keyboard lib เก็บ callback ไว้ระดับ process
        # root.destroy() ไม่ถอดให้ ถ้าไม่ถอด callback เก่าจะยิงใส่ root ที่ตายแล้ว
        # -> RuntimeError: main thread is not in main loop
        keyboard.unhook_all_hotkeys()
        root.destroy()

    try:
        App(root, template, table_index)
        root.protocol("WM_DELETE_WINDOW", close)
    except Exception as e:
        messagebox.showerror("เปิดเอกสารไม่ได้", f"{type(e).__name__}: {e}",
                             parent=root)
        root.destroy()
        return

    root.lift()
    root.mainloop()
