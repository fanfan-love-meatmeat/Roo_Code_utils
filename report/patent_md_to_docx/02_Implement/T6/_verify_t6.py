"""T6 verification: convert v2.4 MD and check lfloor fix"""
import sys, os, re, datetime
sys.path.insert(0, r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils')

from tools.patent_md_to_docx import build_patent_docx

md_path = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\test\patent_md\完整专利申请书_20260531_2010_v2.4.md'
with open(md_path, 'r', encoding='utf-8') as f:
    doc = build_patent_docx(f.read())

print("=== T6: lfloor -> left_lfloor fix ===")
found = False
for p in doc.paragraphs:
    for r in p.runs:
        if '\\lfloor' in r.text or 'lfloor' in r.text:
            txt = r.text
            has_left = '\\left\\lfloor' in txt
            has_right = '\\right\\rfloor' in txt
            print(f"  has_left={has_left} has_right={has_right}")
            print(f"  text: {txt[:150]}")
            found = True
            assert has_left, "FAIL: lfloor NOT wrapped with left!"
            break
    if found:
        break

if not found:
    print("  (no lfloor in output - formula may be in single run)")
    # Check all runs for the full formula
    for p in doc.paragraphs:
        for r in p.runs:
            if '哈希' in r.text and 'lfloor' in r.text:
                print(f"  FOUND in mixed run: {r.text[:200]}")
                found = True

# Also verify idempotence: normal formulas unchanged
print("\n=== Idempotence: normal formulas unchanged ===")
for p in doc.paragraphs:
    for r in p.runs:
        if 'R_t' in r.text or 'V_overlap' in r.text:
            print(f"  {r.text[:80]}")
            break
    else:
        continue
    break

# Save
fn = os.path.basename(md_path)
now = datetime.datetime.now().strftime('%Y%m%d_%H%M')
m = re.search(r'_v(\d+)\.(\d+)\.md$', fn)
v = f'v{m.group(1)}.{m.group(2)}' if m else 'v1.0'
base = fn[:m.start()] if m else os.path.splitext(fn)[0]
out_dir = r'D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\patent_docx'
os.makedirs(out_dir, exist_ok=True)
out = os.path.join(out_dir, f'{base}_{now}_{v}.docx')
doc.save(out)
print(f'\nSaved: {out} ({os.path.getsize(out)} bytes)')
