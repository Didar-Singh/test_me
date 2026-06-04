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

# Matches "Label : value"  (value must be non-empty)
FIELD_RE = re.compile(r'^([A-Za-z][A-Za-z0-9 /_()-]{0,50}?)\s*:\s*(.+)$')
# Matches "Label :"  with NO value — orphan label from two-column form layouts
ORPHAN_RE = re.compile(r'^([A-Za-z][A-Za-z0-9 /_()-]{0,50}?)\s*:\s*$')


def merge_broken_fields(lines: list[str]) -> list[str]:
    """Merge lines where OCR split a label across two lines (next line starts with ':')."""
    merged = []
    i = 0
    while i < len(lines):
        line = lines[i]
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


def pair_columns(lines: list[str]) -> list[str]:
    """
    Fix two-column form layouts where Tesseract reads all labels first, then all values:

        Name     :          →  Name     : John Doe
        SSN      :          →  SSN      : 123-45-6789
        Address  :          →  Address  : 123 Main St
        John Doe
        123-45-6789
        123 Main St

    Detects a run of ≥2 consecutive orphan labels, then pairs them with the
    same number of plain-text lines that follow.
    """
    result = []
    i = 0
    while i < len(lines):
        # Collect consecutive orphan-label lines
        labels = []
        j = i
        while j < len(lines) and ORPHAN_RE.match(lines[j].strip()):
            labels.append(ORPHAN_RE.match(lines[j].strip()).group(1).strip())
            j += 1

        if len(labels) >= 2:
            # Skip blank lines between label block and value block
            while j < len(lines) and not lines[j].strip():
                j += 1
            # Collect the next N non-field, non-blank lines as values
            values = []
            k = j
            while k < len(lines) and len(values) < len(labels):
                s = lines[k].strip()
                if s and not ORPHAN_RE.match(s) and not FIELD_RE.match(s):
                    values.append(s)
                    k += 1
                elif not s:
                    k += 1  # allow stray blank lines between values
                else:
                    break

            if len(values) == len(labels):
                for key, val in zip(labels, values):
                    result.append(f"{key} : {val}")
                i = k
                continue  # successfully paired — move past both blocks

        # No pairing possible — pass the current line through unchanged
        result.append(lines[i])
        i += 1

    return result


def tag_lines(lines: list[str]) -> list[tuple]:
    """Return (tag, original_line, key, value). tag = 'field' | 'blank' | 'text'."""
    tagged = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            tagged.append(('blank', line, '', ''))
        else:
            m = FIELD_RE.match(stripped)
            if m:
                tagged.append(('field', line, m.group(1).strip(), m.group(2).strip()))
            else:
                tagged.append(('text', line, '', ''))
    return tagged


def remove_inter_field_blanks(tagged: list[tuple]) -> list[tuple]:
    """Drop blank lines whose nearest non-blank neighbours on both sides are field lines."""
    result = []
    for i, item in enumerate(tagged):
        if item[0] == 'blank':
            prev_tag = next((t[0] for t in reversed(tagged[:i]) if t[0] != 'blank'), None)
            next_tag = next((t[0] for t in tagged[i + 1:] if t[0] != 'blank'), None)
            if prev_tag == 'field' and next_tag == 'field':
                continue  # OCR-injected gap — drop it
        result.append(item)
    return result


def process_all_pages(raw_pages: list[str]) -> list[str]:
    """
    Process all pages together for document-wide consistent alignment:
    1. Repair split labels and two-column form layouts.
    2. Remove OCR-injected blank lines between field lines.
    3. Align all fields to a single column width computed across all pages.
    """
    all_tagged = []
    for raw in raw_pages:
        lines = merge_broken_fields(raw.splitlines())
        lines = pair_columns(lines)            # fix two-column layout
        tagged = tag_lines(lines)
        tagged = remove_inter_field_blanks(tagged)
        all_tagged.append(tagged)

    all_keys = [key for page in all_tagged for tag, _, key, _ in page if tag == 'field']
    max_key = max((len(k) for k in all_keys), default=0)

    formatted = []
    for tagged in all_tagged:
        out = []
        for tag, line, key, val in tagged:
            if tag == 'field' and max_key:
                out.append(f"{key:<{max_key}} : {val}")
            else:
                out.append(line)
        formatted.append('\n'.join(out))
    return formatted


def pdf_to_txt(pdf_path: str, output_folder: str, file_index: int, total_files: int):
    doc         = fitz.open(pdf_path)
    filename    = os.path.splitext(os.path.basename(pdf_path))[0]
    txt_path    = os.path.join(output_folder, filename + ".txt")
    zoom        = DPI / 72.0
    matrix      = fitz.Matrix(zoom, zoom)
    total_pages = len(doc)

    print(f"\n[{file_index}/{total_files}] {os.path.basename(pdf_path)}  ({total_pages} page(s))")

    raw_pages = []
    start     = time.time()

    with tqdm(total=total_pages, unit="page", ncols=70,
              bar_format="  {l_bar}{bar}| {n_fmt}/{total_fmt} pages  [{elapsed}<{remaining}]") as pbar:
        for page in doc:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            raw_pages.append(pytesseract.image_to_string(img).strip())
            pbar.update(1)

    doc.close()

    formatted = process_all_pages(raw_pages)
    output_lines = [
        f"=== Page {i} ===\n{text}\n"
        for i, text in enumerate(formatted, start=1)
    ]

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))

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
