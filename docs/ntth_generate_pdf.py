"""
NTTH Thesis PDF Generator using Playwright (Headless Chromium)
==============================================================
Converts all 7 HTML thesis parts to PDF then merges into one.

Requirements (auto-installed by this script if needed):
    pip install playwright PyPDF2
    python -m playwright install chromium

Usage:
    python ntth_generate_pdf.py

Output:
    NTTH_Thesis_Final.pdf
"""

import asyncio
import sys
import subprocess
import os
from pathlib import Path

# Force UTF-8 output on Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCS_DIR = Path(__file__).parent

PARTS = [
    "ntth_thesis_part1.html",
    "ntth_thesis_part2.html",
    "ntth_thesis_part3.html",
    "ntth_thesis_part4.html",
    "ntth_thesis_part5.html",
    "ntth_thesis_part6.html",
    "ntth_thesis_part7.html",
]

OUTPUT_PDF = DOCS_DIR / "NTTH_Thesis_Final.pdf"


async def convert_html_to_pdf(playwright, html_path: Path, pdf_path: Path):
    """Convert a single HTML file to PDF using Playwright."""
    browser = await playwright.chromium.launch(headless=True)
    page = await browser.new_page()

    file_url = html_path.as_uri()
    await page.goto(file_url, wait_until='networkidle')

    # Wait for fonts to load
    await asyncio.sleep(1)

    await page.pdf(
        path=str(pdf_path),
        format='Letter',
        margin={
            'top': '1in',
            'right': '1in',
            'bottom': '0.7in',
            'left': '1in',
        },
        print_background=True,
    )
    await browser.close()
    print(f"  Converted: {html_path.name} -> {pdf_path.name}")


def merge_pdfs(pdf_paths: list, output: Path):
    """Merge multiple PDFs into one using PyPDF2."""
    import PyPDF2
    merger = PyPDF2.PdfMerger()
    total_pages = 0

    for pdf_path in pdf_paths:
        reader = PyPDF2.PdfReader(str(pdf_path))
        pages = len(reader.pages)
        print(f"  Merging: {pdf_path.name} ({pages} pages)")
        merger.append(str(pdf_path))
        total_pages += pages

    merger.write(str(output))
    merger.close()

    size_mb = output.stat().st_size / 1024 / 1024
    print(f"\n{'=' * 60}")
    print(f"  MERGED OUTPUT : {output.name}")
    print(f"  Total pages   : {total_pages}")
    print(f"  File size     : {size_mb:.1f} MB")
    print(f"  Location      : {output}")
    print(f"{'=' * 60}\n")
    return total_pages


async def main_async():
    print("=" * 60)
    print("  NTTH THESIS — PDF GENERATOR (Playwright + PyPDF2)")
    print("=" * 60)
    print(f"  Docs folder: {DOCS_DIR}")
    print()

    # Verify all parts exist
    missing = []
    for part in PARTS:
        path = DOCS_DIR / part
        status = "[OK]" if path.exists() else "[MISSING]"
        print(f"  {status}  {part}")
        if not path.exists():
            missing.append(part)
    print()

    if missing:
        print(f"ERROR: {len(missing)} HTML parts are missing:")
        for m in missing:
            print(f"  - {m}")
        sys.exit(1)

    # Convert each HTML part to PDF
    print("Converting HTML parts to PDF...")
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("ERROR: playwright not installed. Run: pip install playwright && python -m playwright install chromium")
        sys.exit(1)

    part_pdfs = []
    async with async_playwright() as pw:
        for part_html in PARTS:
            html_path = DOCS_DIR / part_html
            pdf_path = DOCS_DIR / part_html.replace(".html", "_gen.pdf")
            try:
                await convert_html_to_pdf(pw, html_path, pdf_path)
                part_pdfs.append(pdf_path)
            except Exception as e:
                print(f"  ERROR converting {part_html}: {str(e)[:200]}")

    if not part_pdfs:
        print("ERROR: No PDFs were generated.")
        sys.exit(1)

    # Merge all PDFs
    print(f"\nMerging {len(part_pdfs)} PDF parts...")
    total = merge_pdfs(part_pdfs, OUTPUT_PDF)

    # Clean up temporary PDFs
    print("Cleaning up temporary part PDFs...")
    for p in part_pdfs:
        try:
            p.unlink()
        except Exception:
            pass

    print(f"Done! Final thesis PDF: {OUTPUT_PDF}")
    print(f"Target was 99 pages (sample_report), generated: {total} pages")


def main():
    asyncio.run(main_async())


if __name__ == "__main__":
    main()
