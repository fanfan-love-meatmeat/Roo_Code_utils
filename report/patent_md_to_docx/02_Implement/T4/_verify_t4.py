"""T4 verification: convert v2.4 MD and validate bottom-border + auto-append"""
import sys, os, datetime
sys.path.insert(0, r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils')
from tools.patent_md_to_docx import build_patent_docx
from docx.oxml.ns import qn

md_path = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_md\完整专利申请书_20260531_2010_v2.4.md'
with open(md_path, 'r', encoding='utf-8') as f:
    doc = build_patent_docx(f.read())

paras = doc.paragraphs
print('=' * 55)
print('T4 ACCEPTANCE VERIFICATION')

# A1: H1 - only bottom border, thick (sz=18)
print('\n--- A1: H1 bottom-only border + thickened ---')
for p in paras:
    if p.runs and p.runs[0].font.size and p.runs[0].font.size >= 228600:
        pPr = p._element.find(qn('w:pPr'))
        has_top = False
        has_bottom = False
        bottom_sz = 'N/A'
        if pPr is not None:
            pBdr = pPr.find(qn('w:pBdr'))
            if pBdr is not None:
                top_el = pBdr.find(qn('w:top'))
                bot_el = pBdr.find(qn('w:bottom'))
                has_top = top_el is not None
                has_bottom = bot_el is not None
                if bot_el is not None:
                    bottom_sz = bot_el.get(qn('w:sz'))
        text = p.text
        has_3space = '      ' in text or '   ' in text[:10]
        print(f'  top={has_top} bottom={has_bottom} bot_sz={bottom_sz} '
              f'spaces={has_3space} | {text[:40]}')

# A2: check last paragraphs for 说明书附图 appendix
print('\n--- A2: Auto-appended 说明书附图 ---')
last_paras = [p for p in paras if p.runs][-6:]
for i, p in enumerate(last_paras):
    sz = p.runs[0].font.size / 12700 if p.runs and p.runs[0].font.size else 0
    text = p.text[:50]
    print(f'  last-{6-len(last_paras)+i}: pt({sz:.0f}) | {text}')

# Check that the very last H1-like paragraph is 说明书附图
last_h1 = None
for p in reversed(paras):
    if p.runs and p.runs[0].font.size and p.runs[0].font.size >= 228600:
        last_h1 = p.text[:40]
        break
print(f'\n  Last H1 title: {last_h1}')
print(f'  Is 说明书附图: {"说明书附图" in (last_h1 or "")}')

# Save
out_dir = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_docx'
base = '完整专利申请书'
now = datetime.datetime.now().strftime('%Y%m%d_%H%M')
version = 'v2.5'
out_path = os.path.join(out_dir, f'{base}_{now}_{version}.docx')
doc.save(out_path)
print(f'\nSaved: {out_path}')
print(f'Size: {os.path.getsize(out_path)} bytes')
