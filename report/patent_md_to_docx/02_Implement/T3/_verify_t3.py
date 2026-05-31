"""T3 verification: convert v2.3 MD and validate all T3 acceptance criteria"""
import sys, os, datetime
sys.path.insert(0, r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils')
from tools.patent_md_to_docx import build_patent_docx
from docx.oxml.ns import qn

md_path = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_md\完整专利申请书_20260531_1525_v2.3.md'
with open(md_path, 'r', encoding='utf-8') as f:
    doc = build_patent_docx(f.read())

paras = doc.paragraphs
print('=' * 55)
print('T3 ACCEPTANCE VERIFICATION')
print('=' * 55)

# A1: H1 - paragraph borders + 3-space + no figure number
print('\n--- A1: H1 section titles ---')
for p in paras:
    if p.runs and p.runs[0].font.size and p.runs[0].font.size >= 228600:
        rPr = p.runs[0]._element.rPr
        rFonts = rPr.find(qn('w:rFonts'))
        ea = rFonts.get(qn('w:eastAsia')) if rFonts is not None else ''
        # Check paragraph border
        pPr = p._element.find(qn('w:pPr'))
        has_border = False
        if pPr is not None:
            pBdr = pPr.find(qn('w:pBdr'))
            has_border = pBdr is not None
        # Check underline removed
        has_underline = p.runs[0].underline
        # Check 3-space spacing
        text = p.text
        has_spaces = '   ' in text
        has_fig = '图' in text
        print(f'  [{ea}] border={has_border} ul={has_underline} '
              f'spaces={has_spaces} fig={has_fig} | {text[:40]}')

# A2: H2 - 宋体 Pt(12), conditional alignment
print('\n--- A2: H2 sub titles ---')
for p in paras:
    if p.runs and p.runs[0].font.size:
        sz = p.runs[0].font.size / 12700
        if abs(sz - 12) < 0.5 and p.runs[0].bold and ('技术领域' in p.text or '背景技术' in p.text or '发明内容' in p.text or '附图说明' in p.text or '具体实施方式' in p.text or '摘要附图' in p.text):
            rPr = p.runs[0]._element.rPr
            rFonts = rPr.find(qn('w:rFonts'))
            ea = rFonts.get(qn('w:eastAsia')) if rFonts is not None else ''
            ali = 'C' if p.alignment == 1 else 'L'
            indent = p.paragraph_format.first_line_indent
            ind = '0' if indent is None or indent == 0 else f'{indent/12700:.0f}pt'
            print(f'  [{ali}] {ea} pt({sz:.0f}) bold={p.runs[0].bold} indent={ind} | {p.text[:30]}')

# A3: H3 - 宋体 Pt(12)
print('\n--- A3: H3 embodiment titles ---')
for p in paras:
    if p.runs and p.runs[0].font.size:
        sz = p.runs[0].font.size / 12700
        if abs(sz - 12) < 0.5 and p.runs[0].bold and '实施例' in p.text:
            rPr = p.runs[0]._element.rPr
            rFonts = rPr.find(qn('w:rFonts'))
            ea = rFonts.get(qn('w:eastAsia')) if rFonts is not None else ''
            indent = p.paragraph_format.first_line_indent
            ind = '0' if indent is None or indent == 0 else f'{indent/12700:.0f}pt'
            print(f'  {ea} pt({sz:.0f}) bold={p.runs[0].bold} indent={ind} | {p.text[:40]}')

# Font distribution
from collections import Counter
fonts = Counter()
sizes = Counter()
for p in paras:
    for r in p.runs[:1]:
        rPr = r._element.rPr
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia')) or ''
            if ea:
                fonts[ea] += 1
        if r.font.size:
            sizes[int(r.font.size/12700)] += 1
print(f'\n--- Font distribution ---')
for k, v in fonts.most_common():
    print(f'  {k}: {v} paragraphs')
print(f'\n--- Size distribution ---')
for k in sorted(sizes.keys()):
    print(f'  {k}pt: {sizes[k]} paragraphs')

# Save
out_dir = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_docx'
base = '完整专利申请书'
now = datetime.datetime.now().strftime('%Y%m%d_%H%M')
version = 'v2.4'
out_path = os.path.join(out_dir, f'{base}_{now}_{version}.docx')
doc.save(out_path)
print(f'\nSaved: {out_path}')
print(f'Size: {os.path.getsize(out_path)} bytes')
