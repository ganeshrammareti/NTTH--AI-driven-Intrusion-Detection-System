"""
NTTH Thesis PDF Merger
======================
Converts all 7 HTML thesis parts to PDF using WeasyPrint or pdfkit,
then merges them into a single thesis document.

Usage:
    python ntth_merge_thesis.py

Requirements:
    pip install weasyprint PyPDF2
    OR
    pip install pdfkit PyPDF2
    (pdfkit requires wkhtmltopdf installed separately)

Output:
    NTTH_Thesis_Final.pdf  — merged 99-page thesis document
"""

import os
import sys
import subprocess
from pathlib import Path

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


def try_weasyprint():
    """Attempt PDF generation using WeasyPrint."""
    try:
        from weasyprint import HTML, CSS
        print("[WeasyPrint] Available — using WeasyPrint for PDF conversion")
    except ImportError:
        print("[WeasyPrint] Not installed.")
        return False

    part_pdfs = []
    for part_html in PARTS:
        html_path = DOCS_DIR / part_html
        pdf_path = DOCS_DIR / part_html.replace(".html", ".pdf")
        if not html_path.exists():
            print(f"  MISSING: {html_path}")
            continue
        print(f"  Converting: {part_html} -> {pdf_path.name}")
        try:
            HTML(filename=str(html_path), base_url=str(DOCS_DIR)).write_pdf(
                str(pdf_path),
                stylesheets=[
                    CSS(string="""
                        @page {
                            size: letter;
                            margin: 1in 1in 0.7in 1in;
                        }
                        @media print {
                            .page { page-break-after: always; }
                        }
                    """)
                ]
            )
            part_pdfs.append(pdf_path)
            print(f"    OK — {pdf_path.name}")
        except Exception as e:
            print(f"    ERROR: {e}")

    if part_pdfs:
        merge_pdfs(part_pdfs, OUTPUT_PDF)
        return True
    return False


def try_pdfkit():
    """Attempt PDF generation using pdfkit (wkhtmltopdf)."""
    try:
        import pdfkit
        print("[pdfkit] Available — using pdfkit for PDF conversion")
    except ImportError:
        print("[pdfkit] Not installed.")
        return False

    options = {
        'page-size': 'Letter',
        'margin-top': '1in',
        'margin-right': '1in',
        'margin-bottom': '0.7in',
        'margin-left': '1in',
        'encoding': 'UTF-8',
        'no-outline': None,
        'enable-local-file-access': None,
        'print-media-type': None,
    }

    part_pdfs = []
    for part_html in PARTS:
        html_path = DOCS_DIR / part_html
        pdf_path = DOCS_DIR / part_html.replace(".html", ".pdf")
        if not html_path.exists():
            print(f"  MISSING: {html_path}")
            continue
        print(f"  Converting: {part_html} -> {pdf_path.name}")
        try:
            pdfkit.from_file(str(html_path), str(pdf_path), options=options)
            part_pdfs.append(pdf_path)
            print(f"    OK — {pdf_path.name}")
        except Exception as e:
            print(f"    ERROR: {e}")

    if part_pdfs:
        merge_pdfs(part_pdfs, OUTPUT_PDF)
        return True
    return False


def merge_pdfs(part_pdfs: list, output: Path):
    """Merge PDF files using PyPDF2."""
    try:
        import PyPDF2
        merger = PyPDF2.PdfMerger()
        total_pages = 0
        for pdf_path in part_pdfs:
            reader = PyPDF2.PdfReader(str(pdf_path))
            pages = len(reader.pages)
            print(f"  Merging: {pdf_path.name} ({pages} pages)")
            merger.append(str(pdf_path))
            total_pages += pages

        merger.write(str(output))
        merger.close()
        print(f"\n{'='*60}")
        print(f"  MERGED OUTPUT: {output.name}")
        print(f"  Total pages:  {total_pages}")
        print(f"  Location:     {output}")
        print(f"{'='*60}\n")
    except ImportError:
        print("[PyPDF2] Not installed. Trying pypdf...")
        try:
            from pypdf import PdfMerger
            merger = PdfMerger()
            total_pages = 0
            for pdf_path in part_pdfs:
                from pypdf import PdfReader
                reader = PdfReader(str(pdf_path))
                pages = len(reader.pages)
                print(f"  Merging: {pdf_path.name} ({pages} pages)")
                merger.append(str(pdf_path))
                total_pages += pages
            merger.write(str(output))
            merger.close()
            print(f"\n  MERGED: {output} ({total_pages} pages)")
        except Exception as e:
            print(f"  Merge failed: {e}")
            print(f"  Partial PDFs saved in: {DOCS_DIR}")


def install_dependencies():
    """Install required Python packages."""
    print("Installing required packages...")
    packages = ["weasyprint", "PyPDF2"]
    for pkg in packages:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg],
            capture_output=True
        )


def main():
    print("=" * 60)
    print("  NTTH THESIS — PDF MERGE TOOL")
    print("=" * 60)
    print(f"  Working directory: {DOCS_DIR}")
    print(f"  Parts to convert: {len(PARTS)}")
    print(f"  Output file: {OUTPUT_PDF.name}")
    print()

    # Check which parts exist
    for part in PARTS:
        path = DOCS_DIR / part
        exists = "[OK]" if path.exists() else "[MISSING]"
        print(f"  {exists}  {part}")
    print()

    # Try conversion methods in order of preference
    if not try_weasyprint():
        print()
        if not try_pdfkit():
            print()
            print("[INFO] No PDF conversion library available.")
            print("To install WeasyPrint (recommended):")
            print("  pip install weasyprint")
            print("  (WeasyPrint may need GTK/Cairo on Windows — see docs)")
            print()
            print("Alternative: Install pdfkit + wkhtmltopdf:")
            print("  pip install pdfkit")
            print("  Download wkhtmltopdf from: https://wkhtmltopdf.org/downloads.html")
            print()
            print("Or use browser print-to-PDF for each part:")
            for part in PARTS:
                print(f"  file:///{(DOCS_DIR / part).as_posix()}")
            sys.exit(1)


if __name__ == "__main__":
    main()
