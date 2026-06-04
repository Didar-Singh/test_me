"""
ADP Diagnostic - show raw text from PDF pages
Usage:  python diagnose_adp.py yourfile.pdf --pages 17-19
        python diagnose_adp.py yourfile.pdf --pages 17-20
"""
import sys, re
import pdfplumber

args = [a for a in sys.argv[1:] if not a.startswith("--")]
PAGE_FROM = PAGE_TO = None
for arg in sys.argv[1:]:
    m = re.match(r"--pages[=:]?(\d+)[-:](\d+)$", arg)
    if m:
        PAGE_FROM, PAGE_TO = int(m.group(1)), int(m.group(2))
        break

if len(args) < 1:
    print("Usage: python diagnose_adp.py yourfile.pdf --pages 17-19")
    sys.exit(1)

pdf_path = args[0]

with pdfplumber.open(pdf_path) as pdf:
    total = len(pdf.pages)
    
    # Use --pages if provided, otherwise default to 1-3
    if PAGE_FROM and PAGE_TO:
        p_from = PAGE_FROM
        p_to   = min(total, PAGE_TO)
    else:
        p_from = 1
        p_to   = min(total, 3)
    
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
        print(f"\n[END PAGE {i+1}]\n")

print("\n" + "="*70)
print(f"Done. Showed {p_to - p_from + 1} pages.")
