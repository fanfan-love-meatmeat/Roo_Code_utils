# -*- coding: utf-8 -*-
"""Regression test: verify nibble conflict fix on all 263 formulas"""
import os
import sys
import zipfile
import olefile
import io
import re
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils")
sys.path.insert(0, r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\Roo_Code_utils\tools")

import mtef_parser
import mtef_fast

DOCX = r"D:\FJL\Projects\Kilo_Code_Gobal_Settings\raw\范加力-一种基于人体姿势交互的大田自动跟随运输平台控制方法8.2 V1 0325 FJL.docx"

def is_garbled(latex: str) -> bool:
    """Detect if output contains garbled characters from metadata parsing"""
    for ch in latex:
        cp = ord(ch)
        # Private use area (0xE000-0xF8FF) = likely garbled
        if 0xE000 <= cp <= 0xF8FF:
            return True
        # High codepoints that look like misread MTEF bytes
        if 0x0200 <= cp <= 0x02FF:  # Latin Extended-B
            return True
        if 0x2000 <= cp <= 0x2BFF and ch not in "−×÷≠≤≥≈∞√∫∂∇∑∏∪∩∈∉⊂⊃⊆⊇⊕⊗⊥∥∠△▽◊":
            return True
    return False

def has_valid_tex_structure(latex: str) -> bool:
    """Check if output has recognizable LaTeX structure"""
    if not latex or len(latex) < 2:
        return False
    # Should contain at least some LaTeX-like patterns
    patterns = [r'\\[a-zA-Z]+', r'\{', r'\}', r'\^', r'_', r'\$']
    return any(re.search(p, latex) for p in patterns)

def main():
    print("=" * 80)
    print("REGRESSION TEST: Nibble Conflict Fix")
    print("=" * 80)

    results = []
    stats = Counter()

    with zipfile.ZipFile(DOCX, "r") as zf:
        ole_entries = [n for n in zf.namelist() if n.startswith("word/embeddings/oleObject")]
        total = len(ole_entries)
        print(f"\nTotal OLE objects: {total}")

        for ei, entry_path in enumerate(ole_entries):
            entry_name = os.path.basename(entry_path)

            try:
                with zf.open(entry_path) as f:
                    if not olefile.isOleFile(f):
                        stats["not_ole"] += 1
                        continue
                    ole = olefile.OleFileIO(f)
                    if not ole.exists("Equation Native"):
                        stats["no_stream"] += 1
                        continue
                    raw = ole.openstream("Equation Native").read()
                    mtef_data = raw[28:]

                # Parse with fixed parser
                parser = mtef_parser.MTEFParser(mtef_data)
                latex = parser.parse()

                # Check inline status
                inline = mtef_fast.is_inline_equation(mtef_data)

                garbled = is_garbled(latex)
                valid = has_valid_tex_structure(latex)

                result = {
                    "index": ei,
                    "entry": entry_name,
                    "len": len(mtef_data),
                    "latex_len": len(latex),
                    "latex": latex,
                    "inline": inline,
                    "garbled": garbled,
                    "valid_structure": valid,
                }
                results.append(result)

                if garbled:
                    stats["garbled"] += 1
                else:
                    stats["clean"] += 1
                stats["total"] += 1

            except Exception as e:
                stats["error"] += 1
                results.append({
                    "index": ei,
                    "entry": entry_name,
                    "error": str(e),
                })

            # Progress
            if (ei + 1) % 50 == 0:
                print(f"  Progress: {ei+1}/{total}")

    print(f"\n{'=' * 60}")
    print(f"RESULTS")
    print(f"{'=' * 60}")
    print(f"  Total parsed:         {stats['total']}")
    print(f"  Clean (no garbled):   {stats['clean']}")
    print(f"  Garbled:              {stats['garbled']}")
    print(f"  Errors:               {stats['error']}")
    if stats['total'] > 0:
        pass_rate = stats['clean'] / stats['total'] * 100
        print(f"  Pass rate:            {pass_rate:.1f}%")
    print(f"  Not OLE:              {stats['not_ole']}")
    print(f"  No Equation Native:   {stats['no_stream']}")

    # Show sample outputs
    print(f"\n{'=' * 60}")
    print(f"SAMPLE OUTPUTS (first 10 clean formulas)")
    print(f"{'=' * 60}")
    clean_results = [r for r in results if not r.get("garbled") and "latex" in r]
    for r in clean_results[:10]:
        inline_mark = "INLINE" if r.get("inline") else "DISPLAY"
        print(f"\n  [{r['index']}] {r['entry']} ({r['len']}B) [{inline_mark}]")
        print(f"    {repr(r['latex'][:100])}")

    # Show garbled samples if any
    garbled_results = [r for r in results if r.get("garbled")]
    if garbled_results:
        print(f"\n{'=' * 60}")
        print(f"GARBLED SAMPLES ({len(garbled_results)} total)")
        print(f"{'=' * 60}")
        for r in garbled_results[:5]:
            print(f"\n  [{r['index']}] {r['entry']} ({r['len']}B)")
            print(f"    {repr(r['latex'][:120])}")
    elif stats['total'] > 0:
        print(f"\n  >>> ALL {stats['total']} FORMULAS PASSED CLEAN <<<")

    return results, stats


if __name__ == "__main__":
    main()
