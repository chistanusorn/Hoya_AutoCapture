"""AutoCapture — กด F9 ตอนจบ Lot แล้วแคปหน้าจอลง Validation Report (Word)

รัน:
    python main.py
"""

import single_instance


def main():
    # hotkey เป็นแบบทั้งเครื่อง เปิดซ้อนกันจะแคปซ้ำและเขียนไฟล์ทับกัน
    if not single_instance.acquire():
        if not single_instance.focus_existing("AutoCapture"):
            print("AutoCapture เปิดอยู่แล้ว — ใช้หน้าต่างเดิม")
        _warn_already_running()
        return

    import ui_word
    ui_word.run()


def _warn_already_running():
    """บอกให้รู้ว่าทำไมโปรแกรมไม่เปิด — ไม่งั้นดับเบิลคลิกแล้วเงียบ ไม่รู้เรื่อง"""
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showwarning(
        "AutoCapture เปิดอยู่แล้ว",
        "โปรแกรมเปิดอยู่แล้ว 1 หน้าต่าง\n\n"
        "เปิดซ้อนกันไม่ได้ เพราะกด F9 ครั้งเดียวจะแคป 2 รูป\n"
        "และถ้าใช้ไฟล์เดียวกัน งานจะเขียนทับกันจนรูปหาย")
    root.destroy()


if __name__ == "__main__":
    main()
