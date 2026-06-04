"""
Scanned PDF → Smart TXT Extractor
Extracts text via OCR and auto-aligns key:value fields (Employee ID, Name, etc.)
so formatting is consistent across all output files.

Setup on new PC:
    pip install pymupdf pytesseract Pillow tqdm
    Install Tesseract: https://github.com/UB-Mannheim/tesseract/wiki

Run:
    python pdf_to_txt_smart.py                      # prompts for folder path
    python pdf_to_txt_smart.py "C:\path\to\pdfs"   # pass folder directly
"""

import os
import sys
import re
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

# Matches lines like:  "Employee Id : 123"  or  "EMP Name: Test Test"
FIELD_RE = re.compile(r'^([A-Za-z][A-Za-z0-9 /_()-]{0,50}?)\s*:\s*(.*)$')


def normalise_blank_lines(text: str) -> str:
    """Collapse 3+ consecutive newlines → one blank line."""
    return re.sub(r'\n{3,}', '\n\n', text)


def merge_broken_fields(lines: list[str]) -> list[str]:
    """
    Sometimes OCR splits a key:value pair across two lines, e.g.:
        Employee
        Id : 123
    Detect this by checking if a line has no colon but the next line
    starts with a colon-bearing pattern, then merge them.
    """
    merged = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # If this line has no colon AND the next line looks like ": value", merge
        if (
            i + 1 < len(lines)
            and ':' not in line
            and line.strip()
            and re.match(r'^\s*:', lines[i + 1])
        ):
            merged.append(line.rstrip() + ' ' + lines[i + 1].lstrip())
            i += 2
        else:
            merged.append(line)
            i += 1
    return merged


def align_fields(text: str) -> str:
    """
    Parse key:value lines, find the longest key per page block,
    then reformat so all values start at the same column.
    Non-field lines are passed through unchanged.
    """
    lines = merge_broken_fields(text.splitlines())

    parsed = []  # (is_field, key, value, original)
    for line in lines:
        m = FIELD_RE.match(line.strip())
        if m:
            parsed.append((True, m.group(1).strip(), m.group(2).strip(), line))
        else:
            parsed.append((False, '', '', line))

    field_entries = [p for p in parsed if p[0]]
    if not field_entries:
        return text  # nothing to align, return as-is

    max_key = max(len(p[1]) for p in field_entries)

    out = []
    for is_field, key, value, original in parsed:
        if is_field:
            out.append(f"{key:<{max_key}} : {value}")
        else:
            out.append(original)
    return '\n'.join(out)


def pdf_to_txt(pdf_path: str, output_folder: str, file_index: int, total_files: int):
    doc         = fitz.open(pdf_path)
    filename    = os.path.splitext(os.path.basename(pdf_path))[0]
    txt_path    = os.path.join(output_folder, filename + ".txt")
    zoom        = DPI / 72.0
    matrix      = fitz.Matrix(zoom, zoom)
    total_pages = len(doc)

    print(f"\n[{file_index}/{total_files}] {os.path.basename(pdf_path)}  ({total_pages} page(s))")

    pages = []
    start = time.time()

    with tqdm(total=total_pages, unit="page", ncols=70,
              bar_format="  {l_bar}{bar}| {n_fmt}/{total_fmt} pages  [{elapsed}<{remaining}]") as pbar:
        for page_num, page in enumerate(doc, start=1):
            pix  = page.get_pixmap(matrix=matrix, alpha=False)
            img  = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            raw  = pytesseract.image_to_string(img).strip()

            clean   = normalise_blank_lines(raw)
            aligned = align_fields(clean)

            pages.append(f"=== Page {page_num} ===\n{aligned}\n")
            pbar.update(1)

    doc.close()

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(pages))

    elapsed = time.time() - start
    print(f"  Saved  → {txt_path}")
    print(f"  Time   → {elapsed:.1f}s  ({elapsed / total_pages:.1f}s/page)")


def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1].strip()
    else:
        print("=" * 60)
        print("  Scanned PDF → Smart TXT Extractor")
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

    print(f"\nFolder : {folder}")
    print(f"PDFs   : {len(pdf_files)} file(s) found")
    print("-" * 60)
    for i, f in enumerate(pdf_files, 1):
        print(f"  {i}. {f}")
    print("-" * 60)
    input("\nPress Enter to start extraction...")

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
