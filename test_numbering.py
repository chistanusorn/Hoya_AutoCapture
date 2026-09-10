import sys
import json
from docx import Document
from docx.oxml.ns import qn

def test(docx_path):
    doc = Document(docx_path)
    
    numbering_part = doc.part.numbering_part
    if not numbering_part:
        return
        
    element = numbering_part.element
    
    nums = {}
    for num in element.findall(qn('w:num')):
        numId = num.get(qn('w:numId'))
        absNumId = num.find(qn('w:abstractNumId')).get(qn('w:val'))
        nums[numId] = absNumId
        
    abs_nums = {}
    for abs_num in element.findall(qn('w:abstractNum')):
        absNumId = abs_num.get(qn('w:abstractNumId'))
        levels = {}
        for lvl in abs_num.findall(qn('w:lvl')):
            ilvl = lvl.get(qn('w:ilvl'))
            start = lvl.find(qn('w:start'))
            start_val = int(start.get(qn('w:val'))) if start is not None else 1
            
            fmt = lvl.find(qn('w:numFmt'))
            fmt_val = fmt.get(qn('w:val')) if fmt is not None else 'decimal'
            
            text = lvl.find(qn('w:lvlText'))
            text_val = text.get(qn('w:val')) if text is not None else ''
            
            levels[ilvl] = {'start': start_val, 'fmt': fmt_val, 'text': text_val}
        abs_nums[absNumId] = levels
        
    with open('test_out.json', 'w', encoding='utf-8') as f:
        json.dump({'nums': nums, 'abs_nums': abs_nums}, f, indent=2)

if __name__ == '__main__':
    test(sys.argv[1])
