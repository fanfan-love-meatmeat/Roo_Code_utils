import sys, os
sys.path.insert(0, r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils')
from tools.patent_md_to_docx import build_patent_docx
from docx.oxml.ns import qn

md_path = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_md\完整专利申请书_20260531_1235_v2.1.md'
with open(md_path, 'r', encoding='utf-8') as f:
    doc = build_patent_docx(f.read())

print('=== Hotfix: HeiTi -> KaiTi_GB2312 ===')
for p in doc.paragraphs[:12]:
    if p.runs and p.runs[0].font.size and p.runs[0].font.size >= 177800:
        rPr = p.runs[0]._element.rPr
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            ul = p.runs[0].underline
            ali = 'C' if p.alignment == 1 else 'L'
            print(f'  [{ali}] {ea:16s} ul={ul} | {p.text[:35]}')

from collections import Counter
fonts = Counter()
for p in doc.paragraphs:
    for r in p.runs[:1]:
        rPr = r._element.rPr
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            if ea:
                fonts[ea] += 1

print(f'\nFont distribution:')
for k, v in fonts.most_common():
    print(f'  {k}: {v} paragraphs')

out_path = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_docx\完整专利申请书_20260531.docx'
doc.save(out_path)
print(f'\nSaved: {out_path} ({os.path.getsize(out_path)} bytes)')
