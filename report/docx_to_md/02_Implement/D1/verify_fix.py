# -*- coding: utf-8 -*-
"""Verify the fix: trace _skip_header with full-tag check first"""
import os
import sys
import struct
import zipfile
import olefile
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")

DOCX_PATH = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"

# The corrected _skip_header logic
def fixed_skip_header(data: bytes) -> int:
    pos = 0
    if data[0] != 5:
        return 0

    pos = 5
    end = data.index(0, pos)
    pos = end + 1

    eq_opt = data[pos]
    pos += 1
    if eq_opt & 0x80:
        val_len = data[pos]
        pos += val_len + 1

    while pos < len(data):
        tag = data[pos]
        record_type = tag & 0x0F

        # === FIX: check extended records first ===
        if tag == 0x11:  # FONT_DEF
            pos += 1
            idx = data[pos]; pos += 1
            style = data[pos]; pos += 1
            try:
                null_pos = data.index(0, pos)
                name = data[pos:null_pos]
                pos = null_pos + 1
            except ValueError:
                pos = len(data)
        elif tag == 0x13:  # ENCODING_DEF
            pos += 1
            try:
                null_pos = data.index(0, pos)
                pos = null_pos + 1
            except ValueError:
                pos = len(data)
        elif tag == 0x12:  # EQN_PREFS
            pos += 1
            pos = min(len(data), pos + 80)
        elif tag == 0x10:  # COLOR_DEF
            pos += 1
            pos += 4
        # === FIX END ===
        elif record_type in (0x1, 0x2, 0x3, 0x4, 0x5, 0x6, 0x7, 0x9, 0xA, 0xB, 0xC, 0xD, 0xE):
            break
        elif record_type == 0x0:  # END
            pos += 1
            break
        else:
            pos += 1
    return pos


def hex_dump(data: bytes, offset: int = 0, length: int = 128) -> str:
    lines = []
    end = min(offset + length, len(data))
    for i in range(offset, end, 16):
        chunk = data[i:i+16]
        hex_str = " ".join(f"{b:02X}" for b in chunk)
        ascii_str = "".join(chr(b) if 0x20 <= b < 0x7F else "." for b in chunk)
        lines.append(f"  {i:06X}: {hex_str:<48s} |{ascii_str}|")
    return "\n".join(lines)


def main():
    print("=" * 80)
    print("FIX VERIFICATION: corrected _skip_header start position")
    print("=" * 80)

    with zipfile.ZipFile(DOCX_PATH, "r") as zf:
        ole_entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]

        # Test first 5 formulas
        for ei in range(5):
            entry_path = ole_entries[ei]
            entry_name = os.path.basename(entry_path)

            with zf.open(entry_path) as f:
                ole = olefile.OleFileIO(f)
                raw = ole.openstream("Equation Native").read()
                mtef_data = raw[28:]

            old_pos = 12  # current _skip_header always stops at 12
            new_pos = fixed_skip_header(mtef_data)

            print(f"\nFormula {ei+1}: {entry_name} ({len(mtef_data)}B MTEF)")
            print(f"  OLD _skip_header stops at: {old_pos}")
            print(f"  NEW _skip_header stops at: {new_pos} (+{new_pos - old_pos}B)")
            print(f"  Eqn data remaining: {len(mtef_data) - new_pos}B")

            if new_pos < len(mtef_data):
                eqn_head = mtef_data[new_pos:min(new_pos+30, len(mtef_data))]
                print(f"  First eqn bytes: {eqn_head.hex(' ')}")

                # Quick analysis: MTEF eqn data should start with a recognizable record
                tag = mtef_data[new_pos]
                rec_type = tag & 0x0F
                opts = tag >> 4
                rec_names = {0x1:"LINE", 0x2:"CHAR", 0x3:"TMPL", 0x9:"SIZE",
                            0xA:"FULL", 0xB:"SUB", 0xC:"SUB2", 0xD:"SYM", 0xE:"SUBSYM"}
                print(f"  First record: tag=0x{tag:02X} → {rec_names.get(rec_type, f'type={rec_type}')} opts={opts}")

            print(f"  Eqn data hex (first 64B):")
            print(hex_dump(mtef_data, new_pos, 64))


if __name__ == "__main__":
    main()
