"""Find K+xi formulas and dump hex"""
import sys, io, zipfile, olefile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")
import mtef_parser, mtef_fast

DOCX = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"
with zipfile.ZipFile(DOCX, "r") as zf:
    entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
    for ei, entry in enumerate(entries):
        with zf.open(entry) as f:
            ole = olefile.OleFileIO(f)
            raw = ole.openstream("Equation Native").read()
            mtef = raw[28:]
        p = mtef_parser.MTEFParser(mtef)
        r = p.parse()
        inline = mtef_fast.is_inline_equation(mtef)
        if "xi" in r and "K" in r:
            name = entry.split("/")[-1]
            p2 = mtef_parser.MTEFParser(mtef)
            p2._skip_header()
            hdr = p2.pos
            eqn = mtef[hdr:hdr+20]
            hx = " ".join(format(b, "02X") for b in eqn)
            mtef_head = " ".join(format(b, "02X") for b in mtef[:20])
            print(f"#{ei} {name} inline={inline} hdr={hdr}")
            print(f"  eqn: {hx}")
            print(f"  mtef_head: {mtef_head}")
            print(f"  result: {repr(r[:100])}")
            print(f"  eq_options byte={mtef[11]} (bit0={'inline' if mtef[11]&1 else 'display'})")
            print()
            if ei > 80:
                break
