"""กันเปิดโปรแกรมซ้อนกัน

hotkey เป็นแบบทั้งเครื่อง ถ้าเปิด 2 ตัว กด F9 ครั้งเดียวจะแคป 2 รูป
และถ้าชี้ไฟล์เดียวกัน ทั้งคู่จะเขียนทับกัน = งานหาย

ใช้ mutex ของ Windows แทนไฟล์ล็อก เพราะ OS คืนให้เองตอน process ตาย
ไฟล์ล็อกจะค้างถ้าโปรแกรมพัง แล้วเปิดใหม่ไม่ได้อีกเลย
"""

import ctypes
from ctypes import wintypes

ERROR_ALREADY_EXISTS = 183

_handle = None


def acquire(name="AutoCapture"):
    """จองสิทธิ์เป็นตัวเดียวที่รันอยู่ — True ถ้าได้, False ถ้ามีตัวอื่นเปิดอยู่แล้ว"""
    global _handle

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]

    # Local\ = เห็นเฉพาะ session ของผู้ใช้นี้ ไม่กวน user อื่นที่ remote เข้ามา
    _handle = kernel32.CreateMutexW(None, True, f"Local\\{name}")
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        return False
    return True


def focus_existing(title):
    """ดึงหน้าต่างของตัวที่เปิดอยู่แล้วขึ้นมาหน้าสุด"""
    user32 = ctypes.windll.user32
    found = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def cb(hwnd, _):
        buf = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(hwnd, buf, 256)
        if title in buf.value and user32.IsWindowVisible(hwnd):
            found.append(hwnd)
        return True

    user32.EnumWindows(cb, 0)
    if not found:
        return False

    hwnd = found[0]
    SW_RESTORE = 9
    user32.ShowWindow(hwnd, SW_RESTORE)
    user32.SetForegroundWindow(hwnd)
    return True
