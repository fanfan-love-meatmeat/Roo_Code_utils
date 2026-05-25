"""Find formula producing 뢄/햄 and dump hex"""
import sys, io, zipfile, olefile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")
import mtef_parser

DOCX = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"

with zipfile.ZipFile(DOCX, "r") as zf:
    entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
    found = 0
    for ei, entry in enumerate(entries):
        with zf.open(entry) as f:
            ole = olefile.OleFileIO(f)
            raw = ole.openstream("Equation Native").read()
            mtef = raw[28:]
        parser = mtef_parser.MTEFParser(mtef)
        result = parser.parse()
        if "뢄" in result or "햄" in result:
            found += 1
            name = entry.split("/")[-1]
            print(f"Formula #{ei}: {name} ({len(mtef)}B)")
            print(f"  Output: {repr(result[:200])}")
            print(f"  Full MTEF hex ({len(mtef)} bytes):")
            for i in range(0, len(mtef), 16):
                chunk = mtef[i:i+16]
                hx = " ".join(f"{b:02X}" for b in chunk)
                asc = "".join(chr(b) if 0x20 <= b < 0x7F else "." for b in chunk)
                print(f"  {i:04X}: {hx:<48s} |{asc}|")
            # Also show after skip_header + prefix
            p2 = mtef_parser.MTEFParser(mtef)
            p2._skip_header()
            hdr_end = p2.pos
            p2._consume_size_prefix()
            eqn_start = p2.pos
            eqn = mtef[eqn_start:]
            print(f"  _skip_header→{hdr_end}, prefix→{eqn_start}")
            print(f"  Equation data ({len(eqn)}B) from pos {eqn_start}:")
            for i in range(0, len(eqn), 16):
                chunk = eqn[i:i+16]
                hx = " ".join(f"{b:02X}" for b in chunk)
                asc = "".join(chr(b) if 0x20 <= b < 0x7F else "." for b in chunk)
                print(f"  {eqn_start+i:04X}: {hx:<48s} |{asc}|")
            print()
            if found >= 3:
                break
