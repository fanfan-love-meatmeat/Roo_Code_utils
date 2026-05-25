# -*- coding: utf-8 -*-
"""
MTEF formula conversion full-chain debug script
Usage: python debug_mtef.py
Purpose: dissect word/embeddings/oleObject*.bin -> MTEF bytes -> LaTeX chain
Output: structured log + diagnostic stats
"""
import os
import sys
import struct
import zipfile
import olefile
import json
import io
from collections import Counter

# Force UTF-8 stdout to avoid GBK encoding errors on Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Add tool paths
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")

DOCX_PATH = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "debug_output.jsonl")

import mtef_parser
import mtef_fast

def _safe_mtef_to_latex(data: bytes) -> str:
    """Equivalent to mtef_to_latex but avoids relative import issue"""
    # Track 1: Fast-Path
    tex = mtef_fast.extract_tex_metadata(data)
    if tex:
        return tex
    # Track 2: Slow-Path
    try:
        parser = mtef_parser.MTEFParser(data)
        return parser.parse()
    except Exception:
        return "[FORMULA: MTEF PARSE FAILED]"


# ============================================================================
# MTEF Header Analysis
# ============================================================================

def analyze_mtef_header(data: bytes) -> dict:
    """Parse MTEF v5 header fields"""
    result = {}
    if len(data) < 5:
        return {"error": "data too short", "len": len(data)}
    result["version"] = data[0]
    result["platform"] = "Windows" if data[1] == 1 else "Mac" if data[1] == 2 else f"unknown({data[1]})"
    result["product"] = data[2]
    result["product_version"] = data[3]
    result["product_subversion"] = data[4]

    # app key
    if len(data) > 5:
        try:
            end = data.index(0, 5)
            result["app_key"] = data[5:end].decode("ascii", errors="replace")
            pos = end + 1
        except ValueError:
            result["app_key"] = repr(data[5:20])
            pos = len(data)
    else:
        pos = 5

    # equation options
    if pos < len(data):
        eq_opt = data[pos]
        result["eq_options"] = eq_opt
        result["inline"] = (eq_opt & 0x01) != 0
        result["has_options_value"] = (eq_opt & 0x80) != 0
        pos += 1

    result["data_start_offset"] = pos
    result["total_bytes"] = len(data)
    return result


def hex_dump(data: bytes, offset: int = 0, length: int = 128) -> str:
    """Hex dump formatter"""
    lines = []
    end = min(offset + length, len(data))
    for i in range(offset, end, 16):
        chunk = data[i:i+16]
        hex_str = " ".join(f"{b:02X}" for b in chunk)
        ascii_str = "".join(chr(b) if 0x20 <= b < 0x7F else "." for b in chunk)
        lines.append(f"  {i:06X}: {hex_str:<48s} |{ascii_str}|")
    return "\n".join(lines)


# ============================================================================
# Single formula diagnosis
# ============================================================================

def diagnose_single_formula(data: bytes, idx: int) -> dict:
    """Full-chain diagnosis for a single formula"""
    diag = {
        "index": idx,
        "total_bytes": len(data),
    }

    # 1. Header analysis
    header = analyze_mtef_header(data)
    diag["header"] = header

    # 2. Fast-Path attempt
    fast_tex = mtef_fast.extract_tex_metadata(data)
    diag["fast_path"] = {
        "has_tex": fast_tex is not None,
        "tex": fast_tex,
    }

    # 3. Slow-Path attempt
    try:
        parser = mtef_parser.MTEFParser(data)
        slow_latex = parser.parse()
    except Exception as e:
        slow_latex = f"[PARSE_ERROR: {e}]"
    diag["slow_path"] = {
        "latex": slow_latex,
        "len": len(slow_latex) if slow_latex else 0,
    }

    # 4. Via full entry point
    final = _safe_mtef_to_latex(data)
    diag["final"] = {
        "latex": final,
        "len": len(final) if final else 0,
        "same_as_fast": (final == fast_tex) if fast_tex else False,
        "same_as_slow": (final == slow_latex),
    }

    # 5. Garbled character detection
    garbled_chars = []
    if final:
        for i, ch in enumerate(final):
            cp = ord(ch)
            if 0x0000 <= cp <= 0x001F:
                garbled_chars.append(f"pos={i}: U+{cp:04X} (control char)")
            elif 0xE000 <= cp <= 0xF8FF:
                garbled_chars.append(f"pos={i}: U+{cp:04X} (private use)")
    diag["garbled_count"] = len(garbled_chars)
    if garbled_chars[:20]:
        diag["garbled_samples"] = garbled_chars[:20]

    # 6. Hex dump of equation data section
    data_start = header.get("data_start_offset", 0)
    diag["hex_dump_eqn"] = hex_dump(data, data_start, 96)

    # 7. Character breakdown of parser output
    if slow_latex and not slow_latex.startswith("[PARSE_ERROR"):
        char_analysis = {}
        for ch in slow_latex[:200]:
            cp = ord(ch)
            cat = "ASCII" if cp < 128 else "Latin1" if cp < 256 else "CJK" if 0x4E00 <= cp <= 0x9FFF else f"U+{cp:04X}"
            char_analysis[cat] = char_analysis.get(cat, 0) + 1
        diag["char_categories"] = char_analysis

    return diag


# ============================================================================
# Main process
# ============================================================================

def main():
    print("=" * 80)
    print("MTEF FORMULA FULL-CHAIN DEBUG")
    print(f"DOCX: {DOCX_PATH}")
    print("=" * 80)

    if not os.path.exists(DOCX_PATH):
        print("ERROR: DOCX not found")
        sys.exit(1)

    results = []
    total_stats = Counter()

    # First, let's also manually read the first few bytes more carefully
    # to understand the exact MTEF binary structure

    with zipfile.ZipFile(DOCX_PATH, "r") as zf:
        ole_entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
        print(f"\nTotal OLE objects: {len(ole_entries)}")
        total_stats["ole_objects"] = len(ole_entries)

        for ei, entry_path in enumerate(ole_entries[:10]):  # Limit to first 10 for deep analysis
            entry_name = os.path.basename(entry_path)
            print(f"\n--- Formula {ei+1}: {entry_name} ---")

            with zf.open(entry_path) as f:
                if not olefile.isOleFile(f):
                    print("  SKIP: not OLE")
                    total_stats["not_ole"] += 1
                    continue

                ole = olefile.OleFileIO(f)

                if not ole.exists("Equation Native"):
                    print("  SKIP: no Equation Native stream")
                    total_stats["no_eqn_native"] += 1
                    continue

                raw = ole.openstream("Equation Native").read()
                ole_header = raw[:28]
                mtef_data = raw[28:]

                print(f"  Total: {len(raw)}B (OLE header=28B, MTEF={len(mtef_data)}B)")
                print(f"  OLE header hex: {ole_header.hex()[:80]}...")

                diag = diagnose_single_formula(mtef_data, ei)
                results.append(diag)

                # Quick summary
                header = diag["header"]
                print(f"  MTEF v{header.get('version')}, App: {header.get('app_key', '?')}, Inline: {header.get('inline')}")

                fast = diag["fast_path"]
                print(f"  Fast-Path (TeX metadata): {'YES' if fast['has_tex'] else 'NO'}")
                if fast["has_tex"]:
                    print(f"    TeX: {repr(fast['tex'][:120])}")

                slow = diag["slow_path"]
                print(f"  Slow-Path (recursive parse): {slow['len']} chars")
                preview = slow["latex"][:100] if slow["latex"] else "(empty)"
                print(f"    Preview: {repr(preview)}")

                final = diag["final"]
                if final["same_as_slow"]:
                    print(f"  Final: same as slow-path")
                else:
                    print(f"  Final: {final['len']} chars, preview: {repr(final['latex'][:60])}")

                if diag["garbled_count"] > 0:
                    print(f"  !! GARBLED CHARS: {diag['garbled_count']}")
                    for g in diag.get("garbled_samples", [])[:3]:
                        print(f"     {g}")

                # Character categories
                if "char_categories" in diag:
                    print(f"  Char categories: {diag['char_categories']}")

                # Hex dump
                print(f"  MTEF eqn data hex (first 96B):")
                print(diag["hex_dump_eqn"])

                total_stats["parsed_ok"] += 1

    # Summary stats
    print(f"\n{'=' * 80}")
    print(f"SUMMARY STATS")
    print(f"{'=' * 80}")
    stats = {
        "total_ole_objects": total_stats["ole_objects"],
        "parsed_ok": total_stats.get("parsed_ok", 0),
        "not_ole": total_stats.get("not_ole", 0),
        "no_eqn_native": total_stats.get("no_eqn_native", 0),
    }

    fast_count = sum(1 for r in results if r["fast_path"]["has_tex"])
    slow_count = sum(1 for r in results if r["slow_path"]["latex"] and not r["slow_path"]["latex"].startswith("[PARSE_ERROR"))
    garbled_count = sum(1 for r in results if r.get("garbled_count", 0) > 0)

    stats["fast_path_hits"] = fast_count
    stats["slow_path_ok"] = slow_count
    stats["garbled_formulas"] = garbled_count
    if results:
        stats["garbled_ratio"] = f"{garbled_count}/{len(results)} = {100*garbled_count/len(results):.1f}%"

    for k, v in stats.items():
        print(f"  {k}: {v}")

    # Sample comparison
    print(f"\n  --- First 5 formula samples ---")
    for r in results[:5]:
        fast = r["fast_path"]["tex"] or "(none)"
        slow = r["slow_path"]["latex"][:80] if r["slow_path"]["latex"] else "(empty)"
        print(f"  [{r['index']}] Fast: {repr(fast[:60])}")
        print(f"       Slow: {repr(slow)}")

    # Write detailed output
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        for r in results:
            f.write(json.dumps(r, ensure_ascii=False, default=str) + "\n")
    print(f"\nDetailed output: {OUTPUT_FILE}")

    return results, stats


if __name__ == "__main__":
    main()
