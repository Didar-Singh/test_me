"""
ADP Diagnostic - run this first to see exact raw text from your PDF.
Usage:  python diagnose_adp.py yourfile.pdf
        python diagnose_adp.py yourfile.pdf --pages 1-2
"""
import sys, re
import pdfplumber

args = [a for a in sys.argv[1:] if not a.startswith("--")]
PAGE_FROM = PAGE_TO = None
for arg in sys.argv[1:]:
    m = re.match(r"--pages[=:]?(\d+)[-:](\d+)$", arg)
    if m:
        PAGE_FROM, PAGE_TO = int(m.group(1)), int(m.group(2))

if len(args) < 1:
    print("Usage: python diagnose_adp.py yourfile.pdf [--pages 1-3]")
    sys.exit(1)

pdf_path = args[0]

with pdfplumber.open(pdf_path) as pdf:
    total = len(pdf.pages)
    p_from = (PAGE_FROM or 1)
    p_to   = min(total, PAGE_TO or min(3, total))   # default: first 3 pages
    
    print(f"\nPDF: {pdf_path}")
    print(f"Total pages: {total}")
    print(f"Showing pages: {p_from} to {p_to}")
    print("="*70)

    for i in range(p_from-1, p_to):
        page = pdf.pages[i]
        text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
        print(f"\n{'='*70}")
        print(f"PAGE {i+1} — {len(text)} chars")
        print(f"{'='*70}")
        print(text)
        print(f"\n[END PAGE {i+1}]")

print("\nDone. Copy the output above and share it.")
