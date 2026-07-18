"""
NTTH Thesis — DOCX Merge Script
================================
Merges all 7 individual DOCX parts into one final document
using docxcompose (which properly handles images and styles).

Usage:
    python ntth_merge_docx.py

Input:  docs/docx_parts/ntth_part1.docx ... ntth_part7.docx
Output: docs/NTTH_Thesis_Final.docx
"""

import sys
import os
from pathlib import Path

try:
    from docxcompose.composer import Composer
    from docx import Document
    from docx.shared import Pt, Inches
except ImportError:
    print("Installing docxcompose...")
    os.system(f"{sys.executable} -m pip install docxcompose python-docx")
    from docxcompose.composer import Composer
    from docx import Document
    from docx.shared import Pt, Inches

DOCS_DIR   = Path(__file__).parent
PARTS_DIR  = DOCS_DIR / "docx_parts"
OUTPUT_DOC = DOCS_DIR / "NTTH_Thesis_Final.docx"

PARTS = [
    "ntth_part1.docx",
    "ntth_part2.docx",
    "ntth_part3.docx",
    "ntth_part4.docx",
    "ntth_part5.docx",
    "ntth_part6.docx",
    "ntth_part7.docx",
]


def merge():
    print("=" * 60)
    print("  NTTH THESIS — DOCX MERGE (docxcompose)")
    print("=" * 60)

    # Verify all parts exist
    missing = []
    for part in PARTS:
        path = PARTS_DIR / part
        if not path.exists():
            missing.append(part)
            print(f"  [MISSING] {part}")
        else:
            size_kb = path.stat().st_size / 1024
            print(f"  [OK] {part} ({size_kb:.1f} KB)")

    if missing:
        print(f"\nERROR: {len(missing)} parts missing. Run ntth_build_docx.py first.")
        sys.exit(1)

    print(f"\nMerging into {OUTPUT_DOC.name}...")

    # Use first part as base
    base_path = PARTS_DIR / PARTS[0]
    master = Document(str(base_path))
    composer = Composer(master)

    for part_name in PARTS[1:]:
        part_path = PARTS_DIR / part_name
        print(f"  Appending: {part_name}")
        doc = Document(str(part_path))
        composer.append(doc)

    composer.save(str(OUTPUT_DOC))

    size_mb = OUTPUT_DOC.stat().st_size / 1024 / 1024
    print(f"\n{'='*60}")
    print(f"  MERGED OUTPUT : {OUTPUT_DOC.name}")
    print(f"  File size     : {size_mb:.2f} MB")
    print(f"  Location      : {OUTPUT_DOC}")
    print(f"{'='*60}")
    print("\nOpen NTTH_Thesis_Final.docx in Microsoft Word.")


if __name__ == "__main__":
    merge()
