"""Brute-force alignment scan: find correct equation data start position"""
import sys, io, struct, zipfile, olefile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")

import mtef_parser
DOCX = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"

def parse_from(data, start):
    """Try parsing MTEF from start position, return output"""
    parser = mtef_parser.MTEFParser(data)
    parser.pos = start
    if start >= len(data):
        return ""
    try:
        parser._consume_size_prefix()
        if parser.pos >= len(data):
            return ""
        parts = parser._parse_line()
        return ''.join(parts)
    except:
        return "[ERROR]"

def scan_offsets(data, name):
    """Scan different start positions and show results"""
    print(f"\n--- {name} ({len(data)}B) ---")
    for offset in range(190, min(230, len(data))):
        if offset >= len(data):
            break
        result = parse_from(data, offset)
        if result and len(result) > 0 and not result.startswith('['):
            # Filter: show results that look like valid LaTeX (no CJK, has LaTeX commands)
            has_cjk = any(0x4E00 <= ord(c) <= 0x9FFF or 0x3400 <= ord(c) <= 0x4DBF for c in result)
            has_latex = any(p in result for p in ['\\', '{', '}', '^', '_'])
            if not has_cjk and len(result) > 1:
                print(f"  offset={offset}: {repr(result[:80])}")


def trace_at_offset(data, start_pos, label):
    """Byte-level trace at a specific start position"""
    print(f"\n--- {label} (offset={start_pos}) ---")
    print(f"  Hex: {data[start_pos:min(start_pos+20, len(data))].hex(' ')}")
    pos = start_pos
    for step in range(10):
        if pos >= len(data):
            break
        tag = data[pos]
        rt = tag & 0x0F
        names = {0:"END",1:"LINE",2:"CHAR",3:"TMPL",4:"PILE",5:"MATRIX",
                 6:"EMBELL",7:"RULER",9:"SIZE",0xA:"FULL",0xB:"SUB",
                 0xC:"SUB2",0xD:"SYM",0xE:"SUBSYM"}
        name = names.get(rt, f"0x{rt:02X}")
        print(f"  [0x{pos:04X}] {name} tag=0x{tag:02X}", end="")
        if rt == 0:
            print()
            pos += 1
        elif rt == 1:
            print()
            pos += 1
        elif rt == 2:
            pos += 1
            tf = data[pos]; pos += 1
            b1 = data[pos]; b2 = data[pos+1] if pos+1 < len(data) else 0
            cv = struct.unpack('<H', data[pos:pos+2])[0]
            cv_mtc = b1 & 0xFF
            print(f" tf=0x{tf:02X} cv(uint16)=0x{cv:04X} cv(mtcode)=0x{cv_mtc:02X}")
            pos += 2
        elif rt == 3:
            pos += 1
            if pos + 3 <= len(data):
                sel = data[pos]; var = data[pos+1]; to = data[pos+2]
                sel_name = mtef_parser.TMPL_SELECTORS.get(sel, (f"0x{sel:02X}",))[0]
                print(f" sel={sel_name} var={var}")
                pos += 3
            else:
                print()
                pos = len(data)
        else:
            print()
            pos += 1


def main():
    with zipfile.ZipFile(DOCX, "r") as zf:
        entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]

        # Test formula #0 (should be a simple variable like x_c)
        for ei in [0, 5]:
            with zf.open(entries[ei]) as f:
                ole = olefile.OleFileIO(f)
                raw = ole.openstream("Equation Native").read()
                mtef = raw[28:]

            # Run full parser
            parser = mtef_parser.MTEFParser(mtef)
            result = parser.parse()
            print(f"\n{'='*60}")
            print(f"Formula {ei}: {entries[ei].split('/')[-1]} ({len(mtef)}B)")
            print(f"Full output: {repr(result)}")
            print(f"_skip_header ends at: parser._skip_header() wasn't traced")

            # Scan full data for the equation data boundary
            # Show hex from where _skip_header probably ends
            parser2 = mtef_parser.MTEFParser(mtef)
            hdr_end_before = parser2.pos
            parser2.pos = 0
            parser2._skip_header()
            hdr_end = parser2.pos

            # Show full equation data
            print(f"\n_skip_header stops at: {hdr_end}")
            eqn_data = mtef[hdr_end:min(hdr_end+40, len(mtef))]
            print(f"Equation data hex ({hdr_end}→{hdr_end+len(eqn_data)}):")
            for i in range(0, len(eqn_data), 16):
                chunk = eqn_data[i:i+16]
                hx = " ".join(f"{b:02X}" for b in chunk)
                print(f"  {hdr_end+i:04X}: {hx}")

            # Show where total size prefix ends
            parser3 = mtef_parser.MTEFParser(mtef)
            parser3.pos = 0
            parser3._skip_header()
            prefix_before = parser3.pos
            parser3._consume_size_prefix()
            print(f"After _consume_size_prefix: pos={parser3.pos} (+{parser3.pos-prefix_before})")

            # Trace from that position
            if parser3.pos < len(mtef):
                print(f"\nTrace from {parser3.pos}:")
                trace_at_offset(mtef, parser3.pos, "After prefix")

            # Test: manually align by skipping one byte
            print(f"\n--- Manual alignment test (skip leading 00) ---")
            test_pos = hdr_end
            while test_pos < len(mtef) and mtef[test_pos] == 0x00:
                test_pos += 1
                print(f"  Skipped 00 at {test_pos-1}")

            # Bypass 00 bytes between CHAR tag and typeface
            # The data seems to be: 02 [padding?] typeface char_val
            # Try manually parsing as: tag → skip 1 → typeface → char
            if test_pos < len(mtef):
                print(f"  Manual trace from {test_pos}:")
                trace_at_offset(mtef, test_pos, "After 00 cleanup")


if __name__ == "__main__":
    main()
