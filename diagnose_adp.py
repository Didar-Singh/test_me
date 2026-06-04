"""
ADP Diagnostic - Show raw text from PDF pages

Usage:
  python diagnose_adp.py yourfile.pdf 1              (show page 1 only)
  python diagnose_adp.py yourfile.pdf 1 3            (show pages 1-3)
  python diagnose_adp.py yourfile.pdf 17 19          (show pages 17-19)
"""
import sys
import pdfplumber

if len(sys.argv) < 3:
    print(__doc__)
    sys.exit(1)

pdf_path = sys.argv[1]

try:
    page_from = int(sys.argv[2])
    page_to   = int(sys.argv[3]) if len(sys.argv) > 3 else page_from
except ValueError:
    print("Error: page numbers must be integers")
    print(__doc__)
    sys.exit(1)

with pdfplumber.open(pdf_path) as pdf:
    total = len(pdf.pages)
    
    if page_from < 1 or page_to > total or page_from > page_to:
        print(f"Error: Invalid page range. PDF has {total} pages.")
        sys.exit(1)
    
    print(f"\nPDF: {pdf_path}")
    print(f"Total pages: {total}")
    print(f"Showing: {page_from} to {page_to}\n")
    
    for page_num in range(page_from, page_to + 1):
        page = pdf.pages[page_num - 1]
        text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
        
        print(f"{'='*70}")
        print(f"PAGE {page_num} ({len(text)} chars)")
        print(f"{'='*70}\n")
        print(text)
        print(f"\n[END PAGE {page_num}]\n")

print("="*70)
print(f"Done. Showed {page_to - page_from + 1} page(s).\n")
