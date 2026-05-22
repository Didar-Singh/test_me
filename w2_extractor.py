# w2_extractor.py
# Extracts Employee Name, SSN, and Address from all pages of
# W-2 PDF files (ADP layout and standard IRS layouts).
# Uses coordinate-based word extraction for accurate field detection.
# Returns one row per employee per page. If extraction is wrong, run with --debug first.
#
# Outputs results to a CSV file.
#
# Requires: pdfplumber, pandas, tqdm
#   pip install pdfplumber pandas tqdm
#
# ---------------------------------------------------------------------------
# USAGE
# ---------------------------------------------------------------------------
#
# Debug mode — print raw word positions pdfplumber reads from the first page.
# Run this first on a new PDF to verify label detection before extracting.
#
#   python pythonScripts/w2_extractor.py "C:\YourFolder\sample_w2.pdf" --debug
#
# Single file extraction:
#
#   python pythonScripts/w2_extractor.py "C:\YourFolder\sample_w2.pdf"
#   python pythonScripts/w2_extractor.py "C:\YourFolder\sample_w2.pdf" "C:\Output\results.csv"
#
# Batch folder extraction (all PDFs in a directory):
#
#   python pythonScripts/w2_extractor.py "C:\YourFolder\PDFs\" "C:\Output\results.csv"
#
# ---------------------------------------------------------------------------
# DO NOT commit real W-2 documents or any file containing PII as test data.
# Use synthetic/anonymized samples only.
# ---------------------------------------------------------------------------

import re
import sys
import os
from pathlib import Path

import pdfplumber
import pandas as pd
from tqdm import tqdm


# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# SSN: XXX-XX-XXXX (not EIN format XX-XXXXXXX)
_SSN_PATTERN = re.compile(r"^\d{3}-\d{2}-\d{4}$")
_SSN_IN_LINE = re.compile(r"\b(\d{3}-\d{2}-\d{4})\b")

# City, ST ZIP pattern
_CITY_STATE_ZIP_RE = re.compile(r"^.+,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\s*$")

# US ZIP / ZIP+4
_ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")


# ---------------------------------------------------------------------------
# Layout constants (tuned for ADP W-2 layout)
# The 'e/f' block label sits at top < 280 and x0 < 60
# ---------------------------------------------------------------------------

_EF_LABEL_MAX_Y = 280
_EF_LABEL_MAX_X = 60
_SSN_OFFSET_MIN = 30   # points below e/f label where SSN appears
_SSN_OFFSET_MAX = 80
_BOTTOM_ROW_Y_MIN = 450  # SSNs below this are employer copies — ignore


# ---------------------------------------------------------------------------
# Coordinate-based extraction (primary path — matches extract_w2_gui.py logic)
# ---------------------------------------------------------------------------

def _find_ef_blocks(words: list) -> list:
    """Return position dicts for each 'e/f' label found in the top region of the page."""
    blocks = []
    for w in words:
        if w["top"] > _EF_LABEL_MAX_Y:
            continue
        if w["x0"] > _EF_LABEL_MAX_X:
            continue
        text_lower = w["text"].lower()
        if text_lower == "e/f" or text_lower.startswith("e/f"):
            blocks.append({"top": w["top"], "x0": w["x0"]})
    return blocks


def _extract_employee_from_block(words: list, label_top: float, label_x0: float) -> dict | None:
    """
    Given the position of an 'e/f' label, extract name/street/city-state-zip
    from the 3 lines below it, and the SSN from the area further below.
    """
    # Words in the ~35pt window below the label, within the left column
    block_words = [
        w for w in words
        if label_top + 2 < w["top"] < label_top + 35
        and w["x0"] < 200
        and w["x0"] > label_x0 - 5
    ]
    block_words.sort(key=lambda w: (w["top"], w["x0"]))

    # Group into visual lines by proximity of 'top' coordinate
    lines = []
    LINE_TOL = 4
    for w in block_words:
        if not lines:
            lines.append([w])
            continue
        if abs(w["top"] - lines[-1][0]["top"]) <= LINE_TOL:
            lines[-1].append(w)
        else:
            lines.append([w])

    text_lines = []
    for line in lines:
        line.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in line).strip()
        if text and len(text) > 1:
            text_lines.append(text)

    if len(text_lines) < 2:
        return None

    # Find SSN in the y-offset band below the block label
    ssn_candidates = []
    for w in words:
        if label_top + _SSN_OFFSET_MIN < w["top"] < label_top + _SSN_OFFSET_MAX:
            if w["x0"] < 200:
                m = _SSN_IN_LINE.search(w["text"])
                if m and _SSN_PATTERN.match(m.group(1)):
                    ssn_candidates.append(m.group(1))

    # Fallback: search full text of that region (handles fused tokens)
    if not ssn_candidates:
        region_words = [
            w for w in words
            if label_top + _SSN_OFFSET_MIN < w["top"] < label_top + _SSN_OFFSET_MAX
            and w["x0"] < 250
        ]
        line_text = " ".join(w["text"] for w in sorted(region_words, key=lambda w: w["x0"]))
        for m in _SSN_IN_LINE.finditer(line_text):
            if _SSN_PATTERN.match(m.group(1)):
                ssn_candidates.append(m.group(1))

    ssn = ssn_candidates[0] if ssn_candidates else ""

    # Identify city/state/zip line, work backward for name and street
    name = street = csz = ""
    csz_idx = -1
    for i, ln in enumerate(text_lines):
        if _CITY_STATE_ZIP_RE.match(ln):
            csz_idx = i
            csz = ln
            break

    if csz_idx == -1:
        # Positional fallback when no city/state/zip pattern matched
        if len(text_lines) >= 1:
            name = text_lines[0]
        if len(text_lines) >= 2:
            street = text_lines[1]
        if len(text_lines) >= 3:
            csz = text_lines[2]
    elif csz_idx == 1:
        name = text_lines[0]
    elif csz_idx >= 2:
        name = text_lines[0]
        street = " ".join(text_lines[1:csz_idx]).strip()

    return {
        "name": name.strip(),
        "street": street.strip(),
        "city_state_zip": csz.strip(),
        "ssn": ssn,
    }


# ---------------------------------------------------------------------------
# Text-based fallback (for PDFs where coordinate extraction finds nothing)
# ---------------------------------------------------------------------------

_SSN_RE_LOOSE = re.compile(r"\b(\d{3}[-\s]\d{2}[-\s]\d{4})\b")
_ZIP_RE_TEXT = re.compile(r"\b\d{5}(?:-\d{4})?\b")
_NAME_RE = re.compile(r"\b([A-Z][a-zA-Z\-']+(?:\s+[A-Z][a-zA-Z\-']*){1,4})\b")

_NON_NAME_WORDS = {
    "EARNINGS", "SUMMARY", "WAGES", "TIPS", "COMPENSATION", "FEDERAL", "STATE",
    "LOCAL", "INCOME", "TAX", "WITHHELD", "SECURITY", "MEDICARE", "DEPENDENT",
    "CARE", "BENEFITS", "NONQUALIFIED", "PLANS", "ALLOCATED", "ADVANCE", "EIC",
    "PAYMENT", "SALARY", "BONUS", "TOTAL", "EMPLOYEE", "EMPLOYER", "GROSS",
    "NET", "PAY", "PERIOD", "YTD", "YEAR", "DATE", "AMOUNT", "COPY", "VOID",
    "CORRECTED", "DEPARTMENT", "TREASURY", "REVENUE", "SERVICE", "INTERNAL",
    "STATEMENT", "FORM", "WAGE", "AND", "TAX",
}

_EF_LABELS_TEXT = [
    "e/f employee's name, address, and zip code",
    "e/f employee's name, address, zip code",
    "employee's name, address, and zip code",
    "employee's name, address, zip code",
    "employee's name, address",
]

_SSN_LABELS_TEXT = [
    "employee's social security number",
    "social security number",
    "ssn",
    "social security no",
]


def _is_form_label(text: str) -> bool:
    return bool({w.upper() for w in text.split()} & _NON_NAME_WORDS)


def _extract_name_from_text(block: str) -> str:
    candidates = [c for c in _NAME_RE.findall(block) if not _is_form_label(c)]
    if not candidates:
        return ""
    return max(candidates, key=lambda s: len(s.split()))


def _text_fallback(page_text: str) -> dict:
    """Text-based extraction used when coordinate method yields no results."""
    lines = [ln for ln in page_text.splitlines() if ln.strip()]

    # SSN
    ssn = ""
    for line in lines:
        norm = line.lower().strip()
        for lbl in _SSN_LABELS_TEXT:
            if lbl in norm:
                idx = norm.index(lbl)
                remainder = line[idx + len(lbl):].strip().lstrip(":- ").strip()
                m = _SSN_RE_LOOSE.search(remainder or page_text)
                if m:
                    ssn = m.group(1)
                break
    if not ssn:
        m = _SSN_RE_LOOSE.search(page_text)
        ssn = m.group(1) if m else ""

    # Name + address via e/f label
    name = street = csz = ""
    for i, line in enumerate(lines):
        norm = line.lower().strip()
        if any(lbl in norm for lbl in _EF_LABELS_TEXT):
            subsequent = [ln.strip() for ln in lines[i + 1:] if ln.strip()]
            if subsequent and not _is_form_label(subsequent[0]):
                name = subsequent[0]
                for j in range(1, len(subsequent)):
                    if _ZIP_RE_TEXT.search(subsequent[j]):
                        csz = subsequent[j]
                        if j >= 2 and not _is_form_label(subsequent[j - 1]):
                            street = subsequent[j - 1]
                        break
            break

    return {"EmployeeName": name, "SSN": ssn, "Address": ", ".join(p for p in [street, csz] if p)}


# ---------------------------------------------------------------------------
# Core page extraction — coordinate-based primary, text fallback
# ---------------------------------------------------------------------------

def _extract_page_words(words: list, page_text: str) -> list[dict]:
    """
    Extract all employee records from a page.
    Returns a list of dicts (one per employee found).
    Uses coordinate-based extraction; falls back to text if nothing found.
    """
    blocks = _find_ef_blocks(words)
    records = []
    seen_ssns: set = set()
    seen_keys: set = set()

    for block in blocks:
        result = _extract_employee_from_block(words, block["top"], block["x0"])
        if not result:
            continue
        if result["ssn"] and result["ssn"] in seen_ssns:
            continue
        key = (result["name"], result["street"], result["city_state_zip"])
        if key == ("", "", ""):
            continue
        if not result["ssn"] and key in seen_keys:
            continue
        if result["ssn"]:
            seen_ssns.add(result["ssn"])
        seen_keys.add(key)

        address_parts = [p for p in [result["street"], result["city_state_zip"]] if p]
        records.append({
            "EmployeeName": result["name"],
            "SSN": result["ssn"],
            "Address": ", ".join(address_parts),
        })

    # If coordinate extraction found nothing, use text fallback
    if not records and page_text.strip():
        fb = _text_fallback(page_text)
        if fb["EmployeeName"] or fb["SSN"]:
            records.append(fb)

    return records


def extract_w2(file_path: str, max_pages: int = 0) -> pd.DataFrame:
    """
    Extract Employee Name, SSN, and Address from a W-2 or Earnings Summary PDF.
    Returns one row per employee per page.

    Parameters
    ----------
    file_path : str
        Path to the PDF file.
    max_pages : int
        Maximum number of pages to process. 0 (default) means all pages.

    Returns
    -------
    pd.DataFrame
        DataFrame with columns: File, Page, EmployeeName, SSN, Address.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {file_path}")

    print(f"Opening {path.name} ...", flush=True)
    rows = []
    with pdfplumber.open(str(path)) as pdf:
        if not pdf.pages:
            raise ValueError(f"PDF has no pages: {file_path}")

        total = len(pdf.pages)
        limit = min(total, max_pages) if max_pages > 0 else total
        pages_to_process = pdf.pages[:limit]

        if max_pages > 0 and max_pages < total:
            print(f"  {total} page(s) found. Processing first {limit} page(s)...", flush=True)
        else:
            print(f"  {total} page(s) found. Extracting...", flush=True)

        for page_num, page in enumerate(
            tqdm(pages_to_process, desc=f"  {path.name}", unit="pg", leave=True,
                 dynamic_ncols=True, miniters=1),
            start=1,
        ):
            try:
                words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            except Exception:
                words = []
            page_text = page.extract_text() or ""

            records = _extract_page_words(words, page_text)
            for record in records:
                record["File"] = path.name
                record["Page"] = page_num
                rows.append(record)

    if not rows:
        return pd.DataFrame(columns=["File", "Page", "EmployeeName", "SSN", "Address"])

    df = pd.DataFrame(rows)[["File", "Page", "EmployeeName", "SSN", "Address"]]
    return df


# ---------------------------------------------------------------------------
# Batch helper
# ---------------------------------------------------------------------------

def _save_parts(df: pd.DataFrame, output_csv: str, part_size: int = 1000) -> None:
    """Split *df* into chunks of *part_size* rows and write one CSV per chunk."""
    out = Path(output_csv)
    stem = out.stem
    suffix = out.suffix
    total = len(df)
    num_parts = (total + part_size - 1) // part_size

    for part_num in range(1, num_parts + 1):
        chunk = df.iloc[(part_num - 1) * part_size : part_num * part_size]
        part_path = out.parent / f"{stem}_Part{part_num}{suffix}"
        chunk.to_csv(part_path, index=False)
        print(f"  Part {part_num}/{num_parts}: {len(chunk)} records -> {part_path.name}")


def extract_w2_batch(input_dir: str, output_csv: str, part_size: int = 1000,
                     max_pages: int = 0) -> None:
    """
    Process every PDF in *input_dir* and write results split into part files.

    Parameters
    ----------
    input_dir : str
        Directory containing W-2 PDF files.
    output_csv : str
        Base output path; parts are saved as <stem>_Part1.csv, _Part2.csv, etc.
    part_size : int
        Number of records per output file (default 1000).
    max_pages : int
        Maximum pages to process per PDF. 0 (default) means all pages.
    """
    pdf_files = list(Path(input_dir).glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in: {input_dir}")
        return

    frames = []
    for idx, pdf_path in enumerate(sorted(pdf_files), start=1):
        print(f"\n[{idx}/{len(pdf_files)}] {pdf_path.name}", flush=True)
        try:
            df = extract_w2(str(pdf_path), max_pages=max_pages)
            frames.append(df)
            print(f"  [OK]  {len(df)} records", flush=True)
        except Exception as exc:
            print(f"  [ERR] {exc}", flush=True)

    if frames:
        combined = pd.concat(frames, ignore_index=True)
        print(f"\nExtracted {len(combined)} record(s) total")
        _save_parts(combined, output_csv, part_size)
    else:
        print("No records extracted.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _debug_lines(file_path: str) -> None:
    """
    Print word positions pdfplumber extracts from the first page.
    Use this to verify e/f label coordinates and tune layout constants.
    """
    with pdfplumber.open(file_path) as pdf:
        words = pdf.pages[0].extract_words(use_text_flow=False, keep_blank_chars=False)
    print(f"\n--- Word positions from first page ({len(words)} words) ---")
    print(f"  {'x0':>6}  {'top':>6}  text")
    print(f"  {'-'*6}  {'-'*6}  ----")
    for w in words[:120]:  # cap at 120 to avoid flooding the terminal
        print(f"  {w['x0']:>6.1f}  {w['top']:>6.1f}  {w['text']}")
    if len(words) > 120:
        print(f"  ... ({len(words) - 120} more words)")
    print("---")


if __name__ == "__main__":
    # Usage:
    #   python w2_extractor.py <pdf_or_directory> [output.csv] [--pages N] [--debug]
    #
    # Examples:
    #   python w2_extractor.py sample.pdf                        # all pages
    #   python w2_extractor.py sample.pdf --pages 10             # first 10 pages only
    #   python w2_extractor.py sample.pdf results.csv --pages 5
    #   python w2_extractor.py "C:\PDFs\" results.csv --pages 20 # batch, 20 pages per file
    #   python w2_extractor.py sample.pdf --debug                # show word positions
    #
    # DO NOT pass real W-2 files containing live PII — use anonymised samples only.

    import argparse

    parser = argparse.ArgumentParser(
        prog="w2_extractor",
        description="Extract Employee Name, SSN, and Address from W-2 PDFs.",
    )
    parser.add_argument("target", help="PDF file or folder of PDFs to process")
    parser.add_argument("output", nargs="?", default="w2_extracted.csv",
                        help="Output CSV path (default: w2_extracted.csv)")
    parser.add_argument("--pages", type=int, default=0, metavar="N",
                        help="Only process the first N pages per file (default: all pages)")
    parser.add_argument("--debug", action="store_true",
                        help="Print word positions from first page and exit")
    args = parser.parse_args()

    if args.debug:
        _debug_lines(args.target)
        sys.exit(0)

    if args.pages < 0:
        print("Error: --pages must be a positive integer.")
        sys.exit(1)

    if os.path.isdir(args.target):
        extract_w2_batch(args.target, args.output, max_pages=args.pages)
    else:
        result = extract_w2(args.target, max_pages=args.pages)
        print(result.to_string(index=False))
        print(f"\nExtracted {len(result)} record(s) total")
        _save_parts(result, args.output)
