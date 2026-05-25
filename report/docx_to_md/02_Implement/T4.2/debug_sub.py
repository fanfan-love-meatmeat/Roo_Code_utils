"""Debug selector=0x00 subscript"""
import sys, io, zipfile, olefile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")
import mtef_parser

orig = mtef_parser.MTEFParser._template_to_latex

def patched(self, selector, variation, opts, base=""):
    if selector == 0x00:
        hx = self.data[self.pos:self.pos+10].hex(" ") if self.pos < len(self.data) else "EOF"
        print(f"[TMPL] sel=00 base={repr(base)} pos={self.pos} data={hx}")
        if self.pos < len(self.data) and self.data[self.pos] == 0x00:
            self.pos += 1
            print(f"[TMPL] skipped leading END, new pos={self.pos}")
        subs = self._parse_line()
        ret = f'{base}_{{{subs[0]}}}' if subs else base
        print(f"[TMPL] subs={subs} ret={repr(ret)}")
        return ret
    return orig(self, selector, variation, opts, base)

mtef_parser.MTEFParser._template_to_latex = patched

DOCX = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"
with zipfile.ZipFile(DOCX, "r") as zf:
    entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
    with zf.open(entries[2]) as f:
        ole = olefile.OleFileIO(f)
        raw = ole.openstream("Equation Native").read()
        mtef = raw[28:]
    p = mtef_parser.MTEFParser(mtef)
    p._skip_header()
    p._consume_size_prefix()
    parts = p._parse_line()
    print(f"FINAL parts={parts}")
