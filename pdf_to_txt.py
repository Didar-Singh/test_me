"""
Scanned PDF → TXT Extractor
For each PDF in a folder, creates a matching .txt file with all extracted text.

Setup on new PC:
    pip install pymupdf pytesseract Pillow tqdm
    Install Tesseract: https://github.com/UB-Mannheim/tesseract/wiki

Run:
    python pdf_to_txt.py                        # prompts you to enter folder path
    python pdf_to_txt.py "C:\path\to\pdfs"      # pass folder directly
"""

import os
import sys
import time
import fitz
import pytesseract
from PIL import Image
from tqdm import tqdm

# ── Tesseract path ────────────────────────────────────────────────────────────
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\DidarSingh\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
)

DPI = 300


def pdf_to_txt(pdf_path: str, output_folder: str, file_index: int, total_files: int):
    doc      = fitz.open(pdf_path)
    filename = os.path.splitext(os.path.basename(pdf_path))[0]
    txt_path = os.path.join(output_folder, filename + ".txt")
    zoom     = DPI / 72.0
    matrix   = fitz.Matrix(zoom, zoom)
    total_pages = len(doc)

    print(f"\n[{file_index}/{total_files}] {os.path.basename(pdf_path)}  ({total_pages} page(s))")

    lines     = []
    start     = time.time()

    with tqdm(total=total_pages, unit="page", ncols=70,
              bar_format="  {l_bar}{bar}| {n_fmt}/{total_fmt} pages  [{elapsed}<{remaining}]") as pbar:
        for page_num, page in enumerate(doc, start=1):
            pix  = page.get_pixmap(matrix=matrix, alpha=False)
            img  = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            text = pytesseract.image_to_string(img).strip()
            lines.append(f"=== Page {page_num} ===\n{text}\n")
            pbar.update(1)

    doc.close()

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    elapsed = time.time() - start
    print(f"  Saved  → {txt_path}")
    print(f"  Time   → {elapsed:.1f}s  ({elapsed/total_pages:.1f}s/page)")


def main():
    # ── Get folder ────────────────────────────────────────────────────────────
    if len(sys.argv) > 1:
        folder = sys.argv[1].strip()
    else:
        print("=" * 60)
        print("  Scanned PDF → TXT Extractor")
        print("=" * 60)
        folder = input("\nEnter folder path containing PDF files:\n> ").strip().strip('"')

    if not os.path.isdir(folder):
        print(f"\nError: '{folder}' is not a valid folder.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    pdf_files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))

    if not pdf_files:
        print(f"\nNo PDF files found in:\n  {folder}")
        input("\nPress Enter to exit...")
        sys.exit(0)

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\nFolder : {folder}")
    print(f"PDFs   : {len(pdf_files)} file(s) found")
    print("-" * 60)
    for i, f in enumerate(pdf_files, 1):
        print(f"  {i}. {f}")
    print("-" * 60)
    input("\nPress Enter to start extraction...")

    # ── Process ───────────────────────────────────────────────────────────────
    overall_start = time.time()

    for i, pdf_file in enumerate(pdf_files, 1):
        try:
            pdf_to_txt(os.path.join(folder, pdf_file), folder, i, len(pdf_files))
        except Exception as e:
            print(f"  ERROR: {e}")

    total_time = time.time() - overall_start
    print("\n" + "=" * 60)
    print(f"  All done!  {len(pdf_files)} file(s) processed in {total_time:.1f}s")
    print(f"  Output folder: {folder}")
    print("=" * 60)
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
