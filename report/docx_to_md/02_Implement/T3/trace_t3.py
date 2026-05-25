"""T3 fix verification: trace bytes after EQN_PREFS to locate real equation data"""
import sys, io, os, struct, zipfile, olefile
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")

import mtef_parser
DOCX = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"

def hex_dump(data, offset=0, length=256):
    lines = []
    end = min(offset + length, len(data))
    for i in range(offset, end, 16):
        chunk = data[i:i+16]
        hx = " ".join(f"{b:02X}" for b in chunk)
        asc = "".join(chr(b) if 0x20 <= b < 0x7F else "." for b in chunk)
        lines.append(f"  {i:04X}: {hx:<48s} |{asc}|")
    return "\n".join(lines)


def trace_parse(data, start_pos):
    """Trace records from start_pos, show tag names"""
    pos = start_pos
    depth = 0
    for step in range(40):
        if pos >= len(data):
            break
        tag = data[pos]
        rt = tag & 0x0F
        opts = tag >> 4
        names = {0:"END",1:"LINE",2:"CHAR",3:"TMPL",4:"PILE",5:"MATRIX",
                 6:"EMBELL",7:"RULER",8:"FONT_STYLE_DEF",9:"SIZE",
                 0xA:"FULL",0xB:"SUB",0xC:"SUB2",0xD:"SYM",0xE:"SUBSYM",0xF:"COLOR"}
        name = names.get(rt, f"0x{rt:02X}")
        indent = "  " * depth
        print(f"{indent}[0x{pos:04X}] {name}(rt={rt}) tag=0x{tag:02X} opts={opts}", end="")

        if rt == 0:  # END
            print()
            pos += 1
            depth -= 1
            if depth < 0:
                break
        elif rt == 1:  # LINE
            print()
            pos += 1
            depth += 1
            if opts & 0x8:
                dx=mtef_parser.MTEFParser(data)._read_nudge.__code__  # just skip
                pos += 2 if (data[pos] != 128 or data[pos+1] != 128) else 6
            if opts & 0x4:
                pos += 2
        elif rt == 2:  # CHAR
            pos += 1
            if opts & 0x8:
                pos += 2 if (data[pos] != 128 or data[pos+1] != 128) else 6
            tf = data[pos]; pos += 1
            if opts & 0x4:
                cv = data[pos]; pos += 1
            else:
                cv = struct.unpack('<H', data[pos:pos+2])[0]; pos += 2
            ch = chr(cv) if 0x20 <= cv < 0x7F else f"\\x{cv:02X}"
            if cv > 127:
                ch += f" U+{cv:04X}"
            print(f" tf={tf} val={cv}({ch})")
            if opts & 0x2:
                while pos < len(data):
                    et = data[pos] & 0x0F
                    if et == 0: pos += 1; break
                    elif et == 6: pos += 1; pos += 1 + (2 if (data[pos-1]>>4)&8 else 0)
                    else: pos += 1
        elif rt == 3:  # TMPL
            pos += 1
            if opts & 0x8: pos += 2
            sel = data[pos]; pos += 1
            var = data[pos]; pos += 1
            to = data[pos]; pos += 1
            sel_name = mtef_parser.TMPL_SELECTORS.get(sel, (f"0x{sel:02X}",))[0]
            print(f" sel={sel_name} var={var} opts=0x{to:02X}")
            depth += 1
        elif rt == 9:  # SIZE
            pos += 1
            b = data[pos]; pos += 1
            if b == 101: pos += 2
            elif b == 100: pos += 3
            else: pos += 1
            print(f" b={b}")
        elif rt in (0xA,0xB,0xC,0xD,0xE):  # typesize tags
            pos += 1
            print()
        else:
            pos += 1
            print()


def main():
    print("=" * 80)
    print("T3 FIX: EQN_PREFS + Pre-AST Trace")
    print("=" * 80)

    with zipfile.ZipFile(DOCX, "r") as zf:
        entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]

        for ei in range(3):
            with zf.open(entries[ei]) as f:
                ole = olefile.OleFileIO(f)
                raw = ole.openstream("Equation Native").read()
                mtef_data = raw[28:]

            # Run the FIXED parser
            parser = mtef_parser.MTEFParser(mtef_data)
            latex = parser.parse()

            print(f"\n--- Formula {ei+1}: {os.path.basename(entries[ei])} ({len(mtef_data)}B) ---")
            print(f"Parser output: {repr(latex[:150])}")

            # Manual trace: where does _skip_header end?
            parser2 = mtef_parser.MTEFParser(mtef_data)
            parser2._skip_header()
            hdr_end = parser2.pos
            print(f"\n_skip_header ends at pos={hdr_end}")

            # Show full hex after header
            print(f"Full data from {hdr_end} to end:")
            print(hex_dump(mtef_data, hdr_end, len(mtef_data) - hdr_end))

            # Trace _consume_size_prefix effect
            parser3 = mtef_parser.MTEFParser(mtef_data)
            parser3._skip_header()
            print(f"\nAfter _skip_header: pos={parser3.pos}")
            parser3._consume_size_prefix()
            print(f"After _consume_size_prefix: pos={parser3.pos} (+{parser3.pos - hdr_end})")

            # Trace records from the final position
            print(f"\nTrace from pos={parser3.pos}:")
            trace_parse(mtef_data, parser3.pos)


if __name__ == "__main__":
    main()
