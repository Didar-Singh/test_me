"""
extract_names.py
----------------
Extracts names from a searchable PDF (or falls back to Tesseract OCR).
Expected PDF format:  LASTNAME, FIRSTNAME [MI].
Handles:
  - Single last names:        LOAYZA, JOSE A.
  - Double-barrelled names:   DE LA CRUZ, JUAN B.
  - No middle initial:        SMITH, JOHN

Output:  names_output.txt  (comma-separated: LAST,FIRST,MI)
         names_output.csv  (same, with header row — paste into Excel)

Usage:
  python extract_names.py yourfile.pdf
"""

import sys
import os
import re
import subprocess
import csv

# ── helpers ──────────────────────────────────────────────────────────────────

def extract_text_pdftotext(pdf_path: str) -> str:
    """Use pdftotext (poppler) — works on searchable PDFs."""
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, text=True, check=True
        )
        return result.stdout
    except (subprocess.CalledProcessError, FileNotFoundError):
        return ""


def extract_text_tesseract(pdf_path: str) -> str:
    """
    Fall back to Tesseract OCR.
    Converts each page to an image via pdftoppm, then OCRs with pytesseract.
    Requires: pdftoppm (poppler), tesseract, pytesseract, Pillow
    """
    try:
        import pytesseract
        from PIL import Image
        import tempfile, glob

        with tempfile.TemporaryDirectory() as tmpdir:
            subprocess.run(
                ["pdftoppm", "-jpeg", "-r", "300", pdf_path,
                 os.path.join(tmpdir, "page")],
                check=True
            )
            pages = sorted(glob.glob(os.path.join(tmpdir, "page-*.jpg")) or
                           glob.glob(os.path.join(tmpdir, "page-*.jpeg")))
            if not pages:
                # pdftoppm sometimes uses different zero-padding
                pages = sorted(glob.glob(os.path.join(tmpdir, "page*")))

            all_text = []
            for page_img in pages:
                img = Image.open(page_img)
                text = pytesseract.image_to_string(
                    img,
                    config="--psm 6 -c tessedit_char_whitelist="
                           "ABCDEFGHIJKLMNOPQRSTUVWXYZ., "
                )
                all_text.append(text)
            return "\n".join(all_text)

    except Exception as e:
        print(f"[ERROR] Tesseract fallback failed: {e}")
        return ""


def parse_names(raw_text: str) -> list[dict]:
    """
    Parse lines matching the pattern:
        LASTNAME, FIRSTNAME [MI[.]]
    or multi-word last names like:
        DE LA CRUZ, JUAN B.

    Returns list of dicts: {last, first, mi, raw}
    """
    # Pattern breakdown:
    #   ^([A-Z][A-Z ,]*?)   → last name (1+ uppercase words, may contain spaces)
    #   ,\s*                → the comma separator
    #   ([A-Z]+)            → first name
    #   (?:\s+([A-Z])\.?)?  → optional middle initial (with or without trailing dot)
    pattern = re.compile(
        r"^([A-Z][A-Z ]+?),\s*([A-Z]+)(?:\s+([A-Z])\.?)?\s*$"
    )

    results = []
    for line in raw_text.splitlines():
        line = line.strip()
        # Skip blank lines or lines that are clearly not names
        if not line or len(line) < 3:
            continue
        # Normalise multiple spaces → single space (OCR artifact cleanup)
        line_clean = re.sub(r"\s{2,}", " ", line).strip(".")

        m = pattern.match(line_clean)
        if m:
            last  = m.group(1).strip()
            first = m.group(2).strip()
            mi    = (m.group(3) or "").strip()
            results.append({"last": last, "first": first, "mi": mi, "raw": line})

    return results


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_names.py <yourfile.pdf>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"[ERROR] File not found: {pdf_path}")
        sys.exit(1)

    print(f"[1/3] Trying pdftotext on: {pdf_path}")
    raw = extract_text_pdftotext(pdf_path)

    # Quick sanity check — if we got very little text, fall back to OCR
    if len(raw.strip()) < 20:
        print("[2/3] pdftotext returned little/no text → falling back to Tesseract OCR ...")
        raw = extract_text_tesseract(pdf_path)
    else:
        print("[2/3] pdftotext succeeded, skipping OCR.")

    if not raw.strip():
        print("[ERROR] Could not extract any text. Check poppler/tesseract installation.")
        sys.exit(1)

    print("[3/3] Parsing names ...")
    names = parse_names(raw)

    if not names:
        print("\n[WARNING] No names matched the expected pattern.")
        print("Raw extracted text preview:\n")
        print(raw[:500])
        sys.exit(1)

    # ── output ──
    base = os.path.splitext(pdf_path)[0]
    txt_out = base + "_names.txt"
    csv_out = base + "_names.csv"

    # Plain text: one name per line, comma-separated fields
    with open(txt_out, "w", encoding="utf-8") as f:
        for n in names:
            line = f"{n['last']},{n['first']},{n['mi']}" if n['mi'] \
                   else f"{n['last']},{n['first']}"
            f.write(line + "\n")

    # CSV with header — open directly in Excel
    with open(csv_out, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["last", "first", "mi"])
        writer.writeheader()
        for n in names:
            writer.writerow({"last": n["last"], "first": n["first"], "mi": n["mi"]})

    print(f"\n✅ Done! {len(names)} names extracted.")
    print(f"   Text file : {txt_out}")
    print(f"   CSV file  : {csv_out}  ← open this in Excel\n")

    # Preview
    print(f"{'LAST':<25} {'FIRST':<20} {'MI'}")
    print("-" * 50)
    for n in names[:20]:
        print(f"{n['last']:<25} {n['first']:<20} {n['mi']}")
    if len(names) > 20:
        print(f"  ... and {len(names)-20} more rows.")


if __name__ == "__main__":
    main()
