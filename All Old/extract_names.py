"""
debug_pdf.py
------------
Debug tool to inspect what's actually inside a PDF.
Shows: raw text, form fields, and positional word data.

Usage:
    python debug_pdf.py yourfile.pdf
"""

import sys
import subprocess
from pypdf import PdfReader

def separator(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print('='*60)

def check_pdftotext(pdf_path):
    separator("1. RAW TEXT (pdftotext)")
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, text=True
        )
        text = result.stdout.strip()
        if text:
            print(text[:2000])  # first 2000 chars
        else:
            print("❌ No text extracted — PDF may be scanned/image-based")
    except FileNotFoundError:
        print("❌ pdftotext not found — install poppler")

def check_form_fields(pdf_path):
    separator("2. FORM FIELDS (pypdf)")
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields()
        if fields:
            print(f"✅ Found {len(fields)} form fields:\n")
            for field_name, field_obj in fields.items():
                value = field_obj.get("/V", "(empty)")
                ftype = field_obj.get("/FT", "unknown")
                print(f"  Field : {field_name}")
                print(f"  Type  : {ftype}")
                print(f"  Value : {value}")
                print()
        else:
            print("❌ No form fields found — not a fillable PDF form")
    except Exception as e:
        print(f"❌ Error reading form fields: {e}")

def check_pypdf_text(pdf_path):
    separator("3. PAGE TEXT (pypdf per page)")
    try:
        reader = PdfReader(pdf_path)
        print(f"Total pages: {len(reader.pages)}\n")
        # Show first 3 pages only
        for i, page in enumerate(reader.pages[:3]):
            print(f"--- Page {i+1} ---")
            text = page.extract_text()
            if text:
                print(text[:500])
            else:
                print("(no text on this page)")
            print()
    except Exception as e:
        print(f"❌ Error: {e}")

def check_pdfplumber(pdf_path):
    separator("4. POSITIONAL WORDS (pdfplumber)")
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]  # first page only
            words = page.extract_words()
            if words:
                print(f"✅ Found {len(words)} words on page 1:\n")
                print(f"  {'TEXT':<20} {'X0':>6} {'Y0':>6} {'X1':>6} {'Y1':>6}")
                print(f"  {'-'*50}")
                for w in words[:40]:  # first 40 words
                    print(f"  {w['text']:<20} {w['x0']:>6.1f} {w['top']:>6.1f} "
                          f"{w['x1']:>6.1f} {w['bottom']:>6.1f}")
            else:
                print("❌ No words found via pdfplumber")
    except ImportError:
        print("❌ pdfplumber not installed — run: pip install pdfplumber")
    except Exception as e:
        print(f"❌ Error: {e}")

def check_fonts(pdf_path):
    separator("5. FONT INFO (pdffonts)")
    try:
        result = subprocess.run(
            ["pdffonts", pdf_path],
            capture_output=True, text=True
        )
        print(result.stdout or "No font info returned")
    except FileNotFoundError:
        print("❌ pdffonts not found — install poppler")

# ── main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python debug_pdf.py <yourfile.pdf>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    print(f"\n🔍 Debugging: {pdf_path}\n")

    check_pdftotext(pdf_path)
    check_form_fields(pdf_path)
    check_pypdf_text(pdf_path)
    check_pdfplumber(pdf_path)
    check_fonts(pdf_path)

    print(f"\n{'='*60}")
    print("  DEBUG COMPLETE")
    print(f"{'='*60}\n")
