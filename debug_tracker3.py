import sys
from docx import Document
from word_target import find_tables, step_rows, NumberingTracker

def debug():
    doc_path = r"d:\hoya\AutoCapture\output\VPR-2603-003_F-CP.GE-X34.01_Validation Report_PQ_FogGuard_draft_V1_20260324_capture.docx"
    doc = Document(doc_path)
    tracker = NumberingTracker(doc)
    
    tables = find_tables(doc_path)
    for ti, label, n in tables:
        print(f"Table index {ti} in dropdown: {label}")
        
        table = doc.tables[ti]
        for ri, row in enumerate(table.rows):
            if not row.cells: continue
            c0 = row.cells[0]
            c1 = row.cells[1] if len(row.cells) > 1 else None
            
            t0 = c0.text.strip().replace('\n', ' ')
            t1 = c1.text.strip().replace('\n', ' ') if c1 else ""
            
            print(f"  Row {ri}: No='{t0}' | Desc='{t1}'")
        break # just check the first one

if __name__ == "__main__":
    debug()
