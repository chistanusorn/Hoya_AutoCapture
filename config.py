"""ตั้งค่า AutoCapture

ค่าที่เลือกจากหน้าโปรแกรม (ไฟล์ Word, ตาราง) เก็บใน settings.json
ค่าที่เหลือแก้ในไฟล์นี้
"""

import json
import sys
from pathlib import Path

# โฟลเดอร์ของโปรแกรม — ยึดตำแหน่งไฟล์นี้ ไม่ใช่ cwd
# ถ้าใช้ cwd เวลาเปิดจาก shortcut/Task Scheduler ไฟล์จะไปโผล่ที่อื่น
# และภาพที่ค้างอยู่จะถูกมองว่าหาย
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

# ปุ่มเริ่มต้นที่กดเพื่อแคป — ใช้เมื่อยังไม่เคยตั้งใน settings.json
# ผู้ใช้เปลี่ยนได้จากช่องตั้งปุ่มในหน้าโปรแกรม ค่าที่ตั้งเก็บใน settings.json
# ถ้า F9 ชนกับโปรแกรมอื่น ให้เปลี่ยนเป็น "ctrl+alt+f9"
DEFAULT_HOTKEY = "f9"


# โฟลเดอร์เก็บภาพต้นฉบับ สำรองไว้เผื่อไฟล์ Word เสีย
IMAGE_DIR = BASE_DIR / "captures"

# ความกว้างภาพในช่อง (cm) — ใช้เป็น fallback ถ้าอ่านความกว้างคอลัมน์จริงไม่ได้
# ปกติ fit_width() คำนวณจากคอลัมน์เอง ค่านี้จึงไม่ค่อยถูกใช้
WORD_IMAGE_CM = 8.0

# pixel ที่เก็บใน .png — เก็บไว้สูงเพราะฝังในคอลัมน์แคบ (~7.7cm) แล้ว Word
# ย่ออีกชั้น เก็บ pixel มากไว้ภาพจึงคม อ่านออก
# ยิ่งมากยิ่งคมแต่ไฟล์ .docx ใหญ่ขึ้น — 1600 คมพอสำหรับ screenshot ตัวหนังสือ
WORD_IMAGE_WIDTH = 1600

# โฟลเดอร์เก็บไฟล์ที่โปรแกรมเขียน — ต้นฉบับไม่เคยถูกแตะ
WORD_OUTPUT_DIR = BASE_DIR / "output"

# จำค่าที่เลือกจากหน้าโปรแกรม
SETTINGS_FILE = BASE_DIR / "settings.json"


def load_settings():
    if not SETTINGS_FILE.exists():
        return {}
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        # ไฟล์เสียก็แค่เริ่มใหม่ ไม่ควรทำให้โปรแกรมเปิดไม่ขึ้น
        return {}


def save_settings(**kw):
    data = load_settings()
    data.update(kw)
    SETTINGS_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False),
                             encoding="utf-8")


def word_template():
    """ไฟล์ต้นฉบับที่เลือกไว้ — None ถ้ายังไม่เคยเลือก"""
    p = load_settings().get("word_template")
    return Path(p) if p else None


def word_table_index():
    return load_settings().get("word_table_index", 0)


def hotkey():
    """ปุ่มที่ผู้ใช้ตั้งไว้ — DEFAULT_HOTKEY ถ้ายังไม่เคยตั้ง"""
    hk = load_settings().get("hotkey")
    return hk if hk else DEFAULT_HOTKEY


def valid_hotkey(hk):
    """เช็คว่า keyboard lib รับปุ่มนี้ได้ไหม — คืน combo ที่ปรับให้เป็นมาตรฐาน หรือ None

    ตรวจก่อนบันทึก ไม่งั้นตั้งปุ่มพังแล้วเปิดโปรแกรมครั้งหน้า add_hotkey โยน
    ValueError ทั้งที่ยังไม่ทันสร้างหน้าต่าง = เปิดไม่ขึ้นเงียบๆ
    """
    hk = (hk or "").strip().lower()
    if not hk:
        return None
    try:
        import keyboard
        # parse_hotkey โยน ValueError ถ้าชื่อปุ่มไม่รู้จัก ไม่แตะสถานะ hook
        keyboard.parse_hotkey(hk)
        return hk
    except (ValueError, ImportError):
        return None


def set_hotkey(hk):
    """บันทึกปุ่มใหม่ถ้าใช้ได้ — คืน combo ที่บันทึก หรือ None ถ้าปุ่มไม่ถูกต้อง"""
    valid = valid_hotkey(hk)
    if valid is None:
        return None
    save_settings(hotkey=valid)
    return valid


def word_working(template):
    """ไฟล์ทำงานของต้นฉบับนี้ — คนละต้นฉบับได้คนละไฟล์ ไม่ทับกัน"""
    WORD_OUTPUT_DIR.mkdir(exist_ok=True)
    return WORD_OUTPUT_DIR / f"{Path(template).stem}_capture.docx"
