"""แคปหน้าต่างที่อยู่หน้าสุด แล้วเก็บภาพไว้ให้ฝังลง Word

ไฟล์นี้เป็น logic ล้วน ไม่แสดงผลเอง — เรียกผ่าน ui_word.py
"""

import ctypes
import datetime

import mss
from PIL import Image

import config

user32 = ctypes.windll.user32
user32.SetProcessDPIAware()


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


def active_window():
    """คืน (ชื่อหน้าต่าง, กรอบ) ของหน้าต่างหน้าสุด — None ถ้าหาไม่ได้"""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None

    length = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)

    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    if rect.right <= rect.left or rect.bottom <= rect.top:
        return None

    return buf.value, (rect.left, rect.top, rect.right, rect.bottom)


def capture(bbox):
    left, top, right, bottom = bbox
    with mss.mss() as sct:
        shot = sct.grab({"left": left, "top": top,
                         "width": right - left, "height": bottom - top})
    return Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")


def import_image(src, max_width):
    """คัดลอกรูปจากไฟล์ภายนอกเข้า captures/ ย่อขนาดเหมือนภาพที่แคปเอง

    ไม่ฝังไฟล์ต้นทางตรงๆ เพราะรูปจากกล้อง/tool อื่นมักใหญ่เกิน ทำ .docx บวม
    และ captures/ ต้องมีครบทุกรูปที่เข้าเอกสาร ไว้กู้ตอน Word พัง
    """
    img = Image.open(src)
    if img.mode != "RGB":
        img = img.convert("RGB")
    return save_image(img, datetime.datetime.now(), max_width)


def save_image(img, now, max_width):
    """ย่อภาพแล้วบันทึก — microsecond กันชื่อชนตอนกดรัว"""
    img.thumbnail((max_width, 10_000), Image.LANCZOS)
    config.IMAGE_DIR.mkdir(exist_ok=True)
    path = config.IMAGE_DIR / f"{now:%Y-%m-%d_%H-%M-%S-%f}.png"
    img.save(path)
    return path


class Result:
    """ผลของการกดแคป 1 ครั้ง"""

    def __init__(self, ok, message):
        self.ok = ok
        self.message = message


def capture_to_word(insert):
    """แคป 1 ครั้ง แล้วใส่ลง Word ผ่าน insert(image_path) -> (สำเร็จ, ข้อความ)

    รับ insert เป็น callback เพราะปลายทางต่างกันตามโหมด
    (ช่อง Actual Result ของขั้นตอน หรือ แทน <pic>) แต่ขั้นตอนแคป/กัน error
    เหมือนกัน ภาพเก็บลง captures/ เสมอ ถ้า Word พังก็หยิบไปแปะเองได้
    """
    now = datetime.datetime.now()

    win = active_window()
    if win is None:
        return Result(False, "หาหน้าต่างไม่เจอ ข้าม")
    _, bbox = win

    img = capture(bbox)
    path = save_image(img, now, config.WORD_IMAGE_WIDTH)

    try:
        ok, msg = insert(str(path))
    except PermissionError:
        return Result(False, "Word เปิดไฟล์ค้างอยู่ — ปิดแล้วกดใหม่ (ภาพเก็บไว้แล้ว ไม่หาย)")
    except Exception as e:
        return Result(False, f"ใส่ Word ไม่ได้ ({type(e).__name__}) — ภาพเก็บใน captures ไม่หาย")

    if not ok:
        return Result(False, msg)
    return Result(True, msg)
