"""Dump formula 2 TMPL data"""
import sys, io, zipfile, olefile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")
import mtef_parser

DOCX = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"
with zipfile.ZipFile(DOCX, "r") as zf:
    entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
    for ei in [0, 1, 2, 5]:
        with zf.open(entries[ei]) as f:
            ole = olefile.OleFileIO(f)
            raw = ole.openstream("Equation Native").read()
            mtef = raw[28:]
        p = mtef_parser.MTEFParser(mtef)
        p._skip_header()
        hdr = p.pos
        eqn = mtef[hdr:hdr+20]
        r = p.parse()
        hx = " ".join(f"{b:02X}" for b in eqn)
        print(f"#{ei} hdr={hdr} eqn={hx} -> {repr(r[:60])}")
