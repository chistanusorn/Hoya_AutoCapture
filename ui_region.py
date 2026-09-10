"""ลากเลือกพื้นที่บนภาพนิ่งของทั้งจอ — ใช้ตอนไม่อยากได้ทั้งหน้าต่าง

แคปทั้งจอเก็บไว้ก่อนแล้วค่อยให้ลากบนภาพนั้น ไม่ใช่ลากแล้วค่อยแคป เพราะ
  - หน้าจอโปรแกรมเป้าหมายอาจเปลี่ยนระหว่างที่ยังลากอยู่
  - ตัวกรอบเลือกพื้นที่เองจะไม่ติดเข้าไปในรูป

เรียกผ่าน ui_word.py
"""

import tkinter as tk

import mss
from PIL import Image, ImageTk

HINT = "ลากเพื่อเลือกพื้นที่ที่จะแคป   •   Esc = ยกเลิก"
LINE = "#c62828"
FONT = "Leelawadee UI"

# ลากได้เล็กกว่านี้ถือว่าเผลอคลิก ไม่ใช่ตั้งใจเลือกพื้นที่
MIN_SIZE = 20


class _Overlay:
    def __init__(self, win, full, origin):
        self.win = win
        self.full = full
        self.left, self.top = origin
        self.start = None
        self.rect = None
        self.result = None

        # ครอบทุกจอ — พิกัดติดลบได้ถ้าจอรองอยู่ซ้ายของจอหลัก Tk รับรูปแบบ +-1920+0
        win.geometry(f"{full.width}x{full.height}+{self.left}+{self.top}")
        # ต้อง map หน้าต่างให้ได้ขนาดก่อนถอดขอบ ไม่งั้น Windows ปล่อยค้างไว้ที่ 1x1
        # แล้วจะไม่มีอะไรโผล่บนจอเลย
        win.update_idletasks()
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.config(cursor="crosshair")

        self.photo = ImageTk.PhotoImage(full)
        self.canvas = tk.Canvas(win, width=full.width, height=full.height,
                                highlightthickness=0, cursor="crosshair")
        self.canvas.pack()
        self.canvas.create_image(0, 0, image=self.photo, anchor="nw")

        # แถบทึบรองข้อความ ไม่งั้นตัวหนังสือขาวจมหายถ้าพื้นหลังจอสว่าง
        mid = full.width // 2
        self.canvas.create_rectangle(mid - 230, 8, mid + 230, 44,
                                     fill="#1a1d21", outline="")
        self.canvas.create_text(mid, 26, text=HINT, fill="#ffffff",
                                font=(FONT, 12, "bold"))

        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        win.bind("<Escape>", lambda e: win.destroy())

        win.deiconify()
        win.lift()
        win.update()
        win.focus_force()

    def on_press(self, event):
        self.start = (event.x, event.y)
        if self.rect is not None:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(event.x, event.y, event.x, event.y,
                                                 outline=LINE, width=2)

    def on_drag(self, event):
        if self.start is None:
            return
        self.canvas.coords(self.rect, self.start[0], self.start[1], event.x, event.y)

    def on_release(self, event):
        if self.start is None:
            return
        x1, y1 = self.start
        x2, y2 = event.x, event.y
        box = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        if box[2] - box[0] >= MIN_SIZE and box[3] - box[1] >= MIN_SIZE:
            # พิกัดบน canvas เริ่มที่มุมเดียวกับภาพที่แคปไว้ ตัดได้ตรงๆ
            self.result = self.full.crop(box)
        self.win.destroy()


def select_region(parent):
    """คืนภาพส่วนที่ผู้ใช้ลากเลือก — None ถ้ายกเลิกหรือลากเล็กเกินไป"""
    with mss.mss() as sct:
        mon = sct.monitors[0]          # [0] = กรอบรวมทุกจอ ไม่ใช่จอหลักจอเดียว
        shot = sct.grab(mon)
    full = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

    # ไม่ตั้ง transient — หน้าต่างเต็มจอที่เป็นลูกของหน้าต่างที่ซ่อน/ย่ออยู่
    # จะไม่ถูก map ขึ้นจอ
    win = tk.Toplevel(parent)
    overlay = _Overlay(win, full, (mon["left"], mon["top"]))
    win.grab_set()
    win.wait_window()
    return overlay.result
