# -*- coding: utf-8 -*-
"""Byte-level trace of MTEF equation data parsing"""
import os
import sys
import struct
import zipfile
import olefile
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")

import mtef_parser
import mtef_fast

DOCX_PATH = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"

def hex_dump(data: bytes, offset: int = 0, length: int = 256) -> str:
    lines = []
    end = min(offset + length, len(data))
    for i in range(offset, end, 16):
        chunk = data[i:i+16]
        hex_str = " ".join(f"{b:02X}" for b in chunk)
        ascii_str = "".join(chr(b) if 0x20 <= b < 0x7F else "." for b in chunk)
        lines.append(f"  {i:06X}: {hex_str:<48s} |{ascii_str}|")
    return "\n".join(lines)


def trace_skip_header(data: bytes) -> int:
    """Replicate _skip_header and return where it stops"""
    pos = 0

    # version check
    if data[0] != 5:
        return pos
    print(f"  [0x0000] version={data[0]}, platform={data[1]}, product={data[2]}, "
          f"major={data[3]}, minor={data[4]}")

    pos = 5
    # app key
    end = data.index(0, pos)
    app_key = data[pos:end].decode('ascii', errors='replace')
    print(f"  [0x{pos:04X}] app_key='{app_key}', null at {end}, next pos={end+1}")
    pos = end + 1

    # equation options
    eq_opt = data[pos]
    print(f"  [0x{pos:04X}] eq_options=0x{eq_opt:02X} (inline={eq_opt&1}, has_value={bool(eq_opt&0x80)})")
    pos += 1
    if eq_opt & 0x80:
        val_len = data[pos]
        print(f"  [0x{pos:04X}] options_value_len={val_len}")
        pos += val_len + 1

    record_count = 0
    while pos < len(data):
        tag = data[pos]
        record_type = tag & 0x0F
        options = tag >> 4
        tag_name = {
            0x0: "END", 0x1: "LINE", 0x2: "CHAR", 0x3: "TMPL",
            0x4: "PILE", 0x5: "MATRIX", 0x6: "EMBELL", 0x7: "RULER",
            0x8: "FONT_STYLE_DEF", 0x9: "SIZE", 0xA: "FULL", 0xB: "SUB",
            0xC: "SUB2", 0xD: "SYM", 0xE: "SUBSYM",
            0x10: "COLOR_DEF", 0x11: "FONT_DEF", 0x12: "EQN_PREFS", 0x13: "ENCODING_DEF"
        }.get(record_type, f"UNKNOWN(0x{record_type:02X})")

        # Check if this is equation data (should break)
        is_eqn_data = record_type in (0x1, 0x2, 0x3, 0x4, 0x5, 0x6, 0x7, 0x9, 0xA, 0xB, 0xC, 0xD, 0xE)
        
        if is_eqn_data:
            print(f"  [0x{pos:04X}] BREAK: first eqn record = {tag_name} (tag=0x{tag:02X}, opts={options})")
            break

        record_count += 1
        print(f"  [0x{pos:04X}] #{record_count}: {tag_name} (tag=0x{tag:02X}, opts={options})", end="")

        if record_type == 0x11:  # FONT_DEF
            pos += 1
            idx = data[pos]
            pos += 1
            style = data[pos]
            pos += 1
            try:
                name_end = data.index(0, pos)
                name = data[pos:name_end].decode('utf-8', errors='replace')
                print(f" idx={idx}, style={style}, name='{name}'")
                pos = name_end + 1
            except ValueError:
                remaining = min(len(data)-pos, 30)
                print(f" idx={idx}, style={style}, NO NULL FOUND! Next 30B: {data[pos:pos+remaining].hex()}")
                pos += remaining
        elif record_type == 0x13:  # ENCODING_DEF
            pos += 1
            try:
                name_end = data.index(0, pos)
                name = data[pos:name_end].decode('utf-8', errors='replace')
                print(f" name='{name}'")
                pos = name_end + 1
            except ValueError:
                remaining = min(len(data)-pos, 30)
                print(f" NO NULL FOUND! Next 30B: {data[pos:pos+remaining].hex()}")
                pos += remaining
        elif record_type == 0x12:  # EQN_PREFS
            pos += 1
            skip = min(80, len(data)-pos)
            print(f" skip {skip}B")
            pos += skip
        elif record_type == 0x00:  # END
            pos += 1
            print(" -> break")
            break
        else:
            pos += 1
            print()

    return pos


def trace_parse_eqn(data: bytes, start_pos: int):
    """Byte-level trace of equation data parsing"""
    pos = start_pos
    depth = 0
    step = 0

    def indent():
        return "  " * depth

    while pos < len(data) and step < 60:
        tag = data[pos]
        record_type = tag & 0x0F
        options = tag >> 4

        step += 1

        if record_type == 0x0:  # END
            print(f"{indent()}[0x{pos:04X}] END (tag=0x{tag:02X})")
            pos += 1
            depth -= 1
            if depth < 0:
                break
        elif record_type == 0x1:  # LINE
            print(f"{indent()}[0x{pos:04X}] LINE (tag=0x{tag:02X}, opts={options})")
            pos += 1
            depth += 1
            if options & 0x8:  # xfLMOVE
                dx, dy, consumed = mtef_fast.read_nudge(data, pos)
                print(f"{indent()}  nudge: dx={dx}, dy={dy}")
                pos += consumed
            if options & 0x4:  # xfLSPACE
                lspace = struct.unpack('<H', data[pos:pos+2])[0]
                print(f"{indent()}  lspace: {lspace}")
                pos += 2
        elif record_type == 0x2:  # CHAR
            pos += 1
            if options & 0x8:  # xfLMOVE
                dx, dy, consumed = mtef_fast.read_nudge(data, pos)
                pos += consumed
            typeface = data[pos]
            pos += 1
            if options & 0x4:  # mtefOPT_CHAR_ENC_CHAR_8
                char_val = data[pos]
                pos += 1
            else:
                char_val = struct.unpack('<H', data[pos:pos+2])[0]
                pos += 2

            # Try to make sense of the character
            ch_repr = chr(char_val) if 0x20 <= char_val <= 0x7E else f"\\x{char_val:02X}"
            if char_val > 0x7E:
                ch_repr += f" (U+{char_val:04X})"
            print(f"{indent()}[0x{pos:X}] CHAR tf={typeface} val=0x{char_val:04X} '{ch_repr}'")
            
            if options & 0x2:  # xfEMBELL
                # skip embellishment list
                print(f"{indent()}  (embellishment list follows)")
                while pos < len(data):
                    etag = data[pos]
                    etype = etag & 0x0F
                    if etype == 0x0:
                        pos += 1
                        break
                    elif etype == 0x6:  # EMBELL
                        pos += 1
                        if (etag >> 4) & 0x8:
                            dx, dy, consumed = mtef_fast.read_nudge(data, pos)
                            pos += consumed
                        pos += 1  # embell type
                    else:
                        pos += 1

        elif record_type == 0x3:  # TMPL
            pos += 1
            if options & 0x8:
                dx, dy, consumed = mtef_fast.read_nudge(data, pos)
                print(f"{indent()}  nudge: dx={dx}, dy={dy}")
                pos += consumed
            selector = data[pos]; pos += 1
            variation = data[pos]; pos += 1
            tmpl_opts = data[pos]; pos += 1
            sel_name = mtef_parser.TMPL_SELECTORS.get(selector, (f"0x{selector:02X}", "unknown"))[0]
            print(f"{indent()}[0x{pos:X}] TMPL sel={sel_name} var={variation} opts=0x{tmpl_opts:02X}")
            depth += 1

        elif record_type == 0x9:  # SIZE
            pos += 1
            b = data[pos]; pos += 1
            if b == 101:
                sz = struct.unpack('<H', data[pos:pos+2])[0]; pos += 2
                print(f"{indent()}[0x{pos:X}] SIZE 101 val={sz}")
            elif b == 100:
                lsb = data[pos]; pos += 1
                msb = struct.unpack('<H', data[pos:pos+2])[0]; pos += 2
                print(f"{indent()}[0x{pos:X}] SIZE 100 lsb={lsb} msb={msb}")
            else:
                sb = data[pos]; pos += 1
                print(f"{indent()}[0x{pos:X}] SIZE {b} sb={sb}")

        elif record_type in (0xA, 0xB, 0xC, 0xD, 0xE):  # typesize tags
            names = {0xA: "FULL", 0xB: "SUB", 0xC: "SUB2", 0xD: "SYM", 0xE: "SUBSYM"}
            print(f"{indent()}[0x{pos:04X}] {names[record_type]} (typesize)")
            pos += 1

        elif record_type == 0x8:  # FONT_STYLE_DEF
            print(f"{indent()}[0x{pos:04X}] FONT_STYLE_DEF")
            pos += 1
            pos += 2  # skip 2 bytes

        elif record_type == 0xF:  # COLOR
            print(f"{indent()}[0x{pos:04X}] COLOR")
            pos += 1
            pos += 4  # skip RGBA

        else:
            print(f"{indent()}[0x{pos:04X}] UNKNOWN tag=0x{tag:02X} type={record_type} opts={options}")
            pos += 1
            if step > 50:
                break

    print(f"\n  Stopped at pos={pos}/{len(data)}")


def main():
    print("=" * 80)
    print("MTEF BYTE-LEVEL TRACE")
    print("=" * 80)

    with zipfile.ZipFile(DOCX_PATH, "r") as zf:
        ole_entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]

        # Sample a few formulas: simple inline, display, complex
        for ei in [0, 1, 2, 5, 9]:
            entry_path = ole_entries[ei]
            entry_name = os.path.basename(entry_path)

            with zf.open(entry_path) as f:
                ole = olefile.OleFileIO(f)
                raw = ole.openstream("Equation Native").read()
                mtef_data = raw[28:]

            print(f"\n{'='*60}")
            print(f"FORMULA {ei+1}: {entry_name} ({len(mtef_data)}B)")
            print(f"{'='*60}")

            # Show full MTEF data hex
            print(f"\nFull MTEF hex dump:")
            print(hex_dump(mtef_data, 0, len(mtef_data)))

            # Trace skip_header
            print(f"\n_skip_header trace:")
            eqn_start = trace_skip_header(mtef_data)
            print(f"\nEquation data starts at pos={eqn_start}/{len(mtef_data)}")

            # Show equation data hex
            if eqn_start < len(mtef_data):
                print(f"\nEquation data hex (from pos {eqn_start}):")
                print(hex_dump(mtef_data, eqn_start, len(mtef_data) - eqn_start))

            # Trace equation parsing
            print(f"\n_parse_line trace from pos {eqn_start}:")
            trace_parse_eqn(mtef_data, eqn_start)


if __name__ == "__main__":
    main()
