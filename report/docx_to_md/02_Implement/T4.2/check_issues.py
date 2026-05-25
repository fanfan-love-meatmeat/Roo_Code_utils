"""Find K+xi formulas and dump hex"""
import sys, io, zipfile, olefile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")
import mtef_parser, mtef_fast, os

raw_dir = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw"
for fname in os.listdir(raw_dir):
    if fname.endswith(".docx") and chr(33539) in fname:
        docx_path = os.path.join(raw_dir, fname)
        break

print(f"DOCX: {docx_path}")
with zipfile.ZipFile(docx_path, "r") as zf:
    entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
    found = 0
    for ei, entry in enumerate(entries):
        with zf.open(entry) as f:
            ole = olefile.OleFileIO(f)
            raw = ole.openstream("Equation Native").read()
            mtef = raw[28:]
        p = mtef_parser.MTEFParser(mtef)
        r = p.parse()
        inline = mtef_fast.is_inline_equation(mtef)
        if "xi" in r and "K" in r:
            found += 1
            name = entry.split("/")[-1]
            p2 = mtef_parser.MTEFParser(mtef)
            p2._skip_header()
            hdr = p2.pos
            eqn = mtef[hdr:hdr+20]
            hx = " ".join(format(b, "02X") for b in eqn)
            mtef_head = " ".join(format(b, "02X") for b in mtef[:20])
            eq_opt = mtef[11]
            print(f"\n#{ei} {name} inline={inline} eq_opt={eq_opt} bit0={eq_opt&1}")
            print(f"  eqn({hdr}): {hx}")
            print(f"  mtef_head: {mtef_head}")
            print(f"  result: {repr(r[:120])}")
            if found >= 8:
                break

# Also find formulas that should be inline but have BIT0=0
print(f"\n\n=== Inline/Display check: formulas with bit0=0 (should be display) but short content ===")
with zipfile.ZipFile(docx_path, "r") as zf:
    entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
    count = 0
    for ei, entry in enumerate(entries):
        with zf.open(entry) as f:
            ole = olefile.OleFileIO(f)
            raw = ole.openstream("Equation Native").read()
            mtef = raw[28:]
        eq_opt = mtef[11]
        inline = mtef_fast.is_inline_equation(mtef)
        p = mtef_parser.MTEFParser(mtef)
        r = p.parse()
        # Show if it's marked as display but content is short (likely inline)
        if not inline and len(r) < 20:
            count += 1
            name = entry.split("/")[-1]
            print(f"#{ei} {name} DISPLAY(marked) inline(real)={inline} eq_opt={eq_opt} result={repr(r[:60])}")
            if count >= 10:
                break
