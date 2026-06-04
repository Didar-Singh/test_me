"""
PDF → Scanned PDF Converter
Renders every page of each PDF as a raster image and saves it as a new
image-only PDF (no selectable text), ready for OCR processing.

No extra installs — uses pymupdf and tqdm which are already required by
pdf_to_txt_smart.py.

Run:
    python pdf_to_scanned.py                        # prompts for folder path
    python pdf_to_scanned.py "C:\path\to\pdfs"     # pass folder directly

Output:
    Each input file  →  <name>_scanned.pdf  in the same folder.
    Originals are never modified.
"""

import os
import sys
import time
import fitz
from tqdm import tqdm

DPI = 200   # 200 dpi: good OCR quality, reasonable file size


def rasterise(pdf_path: str, output_folder: str, file_index: int, total_files: int):
    doc        = fitz.open(pdf_path)
    out_doc    = fitz.open()
    matrix     = fitz.Matrix(DPI / 72, DPI / 72)
    total_pages = len(doc)
    stem       = os.path.splitext(os.path.basename(pdf_path))[0]
    out_path   = os.path.join(output_folder, stem + "_scanned.pdf")

    print(f"\n[{file_index}/{total_files}] {os.path.basename(pdf_path)}  ({total_pages} page(s))")

    start = time.time()
    with tqdm(total=total_pages, unit="page", ncols=70,
              bar_format="  {l_bar}{bar}| {n_fmt}/{total_fmt} pages  [{elapsed}<{remaining}]") as pbar:
        for page in doc:
            pix      = page.get_pixmap(matrix=matrix, alpha=False)
            img_page = out_doc.new_page(width=pix.width, height=pix.height)
            img_page.insert_image(img_page.rect, pixmap=pix)
            pbar.update(1)

    out_doc.save(out_path, garbage=4, deflate=True)
    doc.close()
    out_doc.close()

    elapsed = time.time() - start
    print(f"  Saved  → {out_path}")
    print(f"  Time   → {elapsed:.1f}s  ({elapsed / total_pages:.1f}s/page)")


def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1].strip()
    else:
        print("=" * 60)
        print("  PDF → Scanned PDF Converter")
        print("=" * 60)
        folder = input("\nEnter folder path containing PDF files:\n> ").strip().strip('"')

    if not os.path.isdir(folder):
        print(f"\nError: '{folder}' is not a valid folder.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # Exclude files already produced by this script
    pdf_files = sorted(
        f for f in os.listdir(folder)
        if f.lower().endswith(".pdf") and not f.lower().endswith("_scanned.pdf")
    )

    if not pdf_files:
        print(f"\nNo PDF files found in:\n  {folder}")
        input("\nPress Enter to exit...")
        sys.exit(0)

    print(f"\nFolder : {folder}")
    print(f"PDFs   : {len(pdf_files)} file(s) found  (DPI = {DPI})")
    print("-" * 60)
    for i, f in enumerate(pdf_files, 1):
        print(f"  {i}. {f}  →  {os.path.splitext(f)[0]}_scanned.pdf")
    print("-" * 60)
    input("\nPress Enter to start conversion...")

    overall_start = time.time()

    for i, pdf_file in enumerate(pdf_files, 1):
        try:
            rasterise(os.path.join(folder, pdf_file), folder, i, len(pdf_files))
        except Exception as e:
            print(f"  ERROR: {e}")

    total_time = time.time() - overall_start
    print("\n" + "=" * 60)
    print(f"  All done!  {len(pdf_files)} file(s) converted in {total_time:.1f}s")
    print(f"  Output folder: {folder}")
    print("=" * 60)
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
