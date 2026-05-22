# w2_extractor.py
# Extracts Employee Name, SSN, and Address from all pages of
# W-2 and Earnings Summary PDF files (2013–2016 formats, searchable PDFs).
# Returns one row per page. If extraction is wrong, run with --debug first.
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
# Debug mode — print raw lines pdfplumber reads from the first page.
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

# SSN: 123-45-6789 or 123 45 6789 (with optional spaces around separators)
_SSN_RE = re.compile(r"\b(\d{3}[-\s]\d{2}[-\s]\d{4})\b")

# Multi-word name: Title Case ("John Doe") or ALL-CAPS ("JOHN DOE"), 2–5 tokens
_NAME_RE = re.compile(r"\b([A-Z][a-zA-Z\-']+(?:\s+[A-Z][a-zA-Z\-']*){1,4})\b")

# US ZIP / ZIP+4 anchors an address line
_ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")


# ---------------------------------------------------------------------------
# Words that disqualify a text token from being an employee name.
# These are W-2 form labels, financial terms, and document header words.
# ---------------------------------------------------------------------------

_NON_NAME_WORDS = {
    "EARNINGS", "SUMMARY", "WAGES", "TIPS", "COMPENSATION", "FEDERAL", "STATE",
    "LOCAL", "INCOME", "TAX", "WITHHELD", "SECURITY", "MEDICARE", "DEPENDENT",
    "CARE", "BENEFITS", "NONQUALIFIED", "PLANS", "ALLOCATED", "ADVANCE", "EIC",
    "PAYMENT", "SALARY", "BONUS", "TOTAL", "EMPLOYEE", "EMPLOYER", "GROSS",
    "NET", "PAY", "PERIOD", "YTD", "YEAR", "DATE", "AMOUNT", "COPY", "VOID",
    "CORRECTED", "DEPARTMENT", "TREASURY", "REVENUE", "SERVICE", "INTERNAL",
    "STATEMENT", "FORM", "WAGE", "AND", "TAX",
}


# ---------------------------------------------------------------------------
# Label-based field anchors found on IRS W-2 / earnings summary layouts
# ---------------------------------------------------------------------------

_SSN_LABELS = [
    "employee's social security number",
    "social security number",
    "ssn",
    "social security no",
]

# The IRS "e/f" box label combines name + address on one label line.
# The three data lines that follow are: Name, Street, City/State/ZIP.
_EF_LABELS = [
    "e/f employee's name, address, and zip code",
    "e/f employee's name, address, zip code",
    "employee's name, address, and zip code",
    "employee's name, address, zip code",
    "employee's name, address",
]

_NAME_LABELS = [
    "employee's first name and initials",
    "employee name",
    "employee's name",
    "first name",
]

_ADDR_LABELS = [
    "employee's address and zip code",
    "employee address",
    "city, state, zip",
]


def _normalise(text: str) -> str:
    return text.lower().strip()


def _find_field(lines: list[str], labels: list[str]) -> str:
    """
    Find a field value by label — checks same line first, then next line.

    Same-line case (common in W-2 multi-column layouts):
        "a Employee's social security number  123-45-6789"

    Next-line case:
        "e/f Employee's Name, Address, and ZIP Code"
        "ALBERTO RAJ ANTONY"
    """
    for i, line in enumerate(lines):
        norm = _normalise(line)
        for lbl in labels:
            if lbl in norm:
                # Value after the label on the same line
                idx = norm.index(lbl)
                remainder = line[idx + len(lbl):].strip().lstrip(":- ").strip()
                if remainder:
                    return remainder
                # Value on the next non-blank line
                for candidate in lines[i + 1:]:
                    if candidate.strip():
                        return candidate.strip()
    return ""


def _find_ef_block(lines: list[str]) -> tuple[str, str, str]:
    """
    Locate the e/f combined label and extract name, street, city/state/ZIP.

    Instead of fixed index offsets, we anchor on the ZIP code line so that
    any extra lines pdfplumber inserts between fields don't shift the result.

    W-2 box e/f layout:
        e/f Employee's Name, Address, and ZIP Code   <- label
        ALBERTO RAJ ANTONY                           <- name  (line after label)
        519 TRADITION PKWY 4200                      <- street (line before ZIP)
        PLEASANTON CA 94566-4477                     <- city/state/zip (ZIP anchor)
    """
    for i, line in enumerate(lines):
        norm = _normalise(line)
        if any(lbl in norm for lbl in _EF_LABELS):
            subsequent = [ln.strip() for ln in lines[i + 1:] if ln.strip()]

            # First subsequent non-blank line = employee name
            name = subsequent[0] if subsequent else ""
            if not name or _is_form_label(name):
                return "", "", ""

            # Scan forward from after the name for the first line with a ZIP
            street = ""
            city_st_zip = ""
            for j in range(1, len(subsequent)):
                if _ZIP_RE.search(subsequent[j]):
                    city_st_zip = subsequent[j]
                    # The line immediately before the ZIP line is the street
                    if j >= 2 and not _is_form_label(subsequent[j - 1]):
                        street = subsequent[j - 1]
                    elif j == 1:
                        # Name and city/state/zip are adjacent — no street line
                        street = ""
                    break

            return name, street, city_st_zip
    return "", "", ""


def _extract_ssn(text: str) -> str:
    """Return the first SSN-shaped token found in *text*."""
    match = _SSN_RE.search(text)
    return match.group(1) if match else ""


def _is_form_label(text: str) -> bool:
    """Return True if any word in *text* is a known W-2 form/financial term."""
    return bool({w.upper() for w in text.split()} & _NON_NAME_WORDS)


def _extract_name_from_block(block: str) -> str:
    """Return the longest multi-word name candidate from *block*, excluding form labels."""
    candidates = [c for c in _NAME_RE.findall(block) if not _is_form_label(c)]
    if not candidates:
        return ""
    return max(candidates, key=lambda s: len(s.split()))


def _extract_address_near(lines: list[str], anchor_name: str) -> str:
    """
    Find the employee address by looking for a ZIP code in lines near where
    the employee name appears. Avoids picking up the employer's address which
    appears earlier on the page.
    """
    # Find the line index where the employee name appears
    name_idx = -1
    if anchor_name:
        for i, line in enumerate(lines):
            if anchor_name.split()[0] in line:  # match on first word of name
                name_idx = i
                break

    # If we found the name, search for ZIP in the lines after it
    search_lines = lines[name_idx:] if name_idx >= 0 else lines
    for i, line in enumerate(search_lines):
        if _ZIP_RE.search(line):
            street = search_lines[i - 1].strip() if i > 0 else ""
            city_state_zip = line.strip()
            # Reject if the street line looks like a form label
            if street and _is_form_label(street):
                street = ""
            parts = [p for p in [street, city_state_zip] if p]
            return ", ".join(parts)
    return ""


# ---------------------------------------------------------------------------
# Core extraction
# ---------------------------------------------------------------------------

def _extract_page(page_text: str) -> dict:
    """Extract name, SSN, and address fields from a single page's raw text."""
    lines = [ln for ln in page_text.splitlines() if ln.strip()]

    # --- SSN ---------------------------------------------------------------
    # Check same line as label AND next line, then fall back to full-page scan
    ssn_raw = _find_field(lines, _SSN_LABELS)
    ssn = _extract_ssn(ssn_raw) if ssn_raw else _extract_ssn(page_text)

    # --- Name + Address (e/f combined box — primary path) ------------------
    ef_name, ef_street, ef_city = _find_ef_block(lines)

    if ef_name:
        name = ef_name
        parts = [p for p in [ef_street, ef_city] if p]
        address = ", ".join(parts)
        # If ef_block found name but missed address, try the smarter fallback
        if not address:
            address = _extract_address_near(lines, name)
    else:
        # Fallback: separate label lookups
        name_raw = _find_field(lines, _NAME_LABELS)
        name = _extract_name_from_block(name_raw) if name_raw else ""

        if not name:
            # Search within 5 lines of the SSN — name is always near SSN on W-2
            for i, line in enumerate(lines):
                if _SSN_RE.search(line):
                    window = lines[max(0, i - 5): i + 5]
                    for candidate_line in window:
                        name = _extract_name_from_block(candidate_line)
                        if name:
                            break
                if name:
                    break

        # Address: check "f" box label (same line or next), then near-name scan
        addr_raw = _find_field(lines, _ADDR_LABELS)
        if addr_raw and _ZIP_RE.search(addr_raw):
            address = addr_raw
        else:
            address = _extract_address_near(lines, name)

    return {"EmployeeName": name, "SSN": ssn, "Address": address}


def extract_w2(file_path: str) -> pd.DataFrame:
    """
    Extract Employee Name, SSN, and Address from all pages of a
    W-2 or Earnings Summary PDF. Returns one row per page that yields data.

    Parameters
    ----------
    file_path : str
        Path to the PDF file.

    Returns
    -------
    pd.DataFrame
        DataFrame with columns: File, Page, EmployeeName, SSN, Address.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {file_path}")

    rows = []
    with pdfplumber.open(str(path)) as pdf:
        if not pdf.pages:
            raise ValueError(f"PDF has no pages: {file_path}")

        for page_num, page in enumerate(
            tqdm(pdf.pages, desc=f"  {path.name}", unit="pg", leave=False),
            start=1,
        ):
            raw_text = page.extract_text() or ""
            if not raw_text.strip():
                continue
            record = _extract_page(raw_text)
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

def extract_w2_batch(input_dir: str, output_csv: str) -> None:
    """
    Process every PDF in *input_dir* and write combined results to *output_csv*.

    Parameters
    ----------
    input_dir : str
        Directory containing W-2 PDF files.
    output_csv : str
        Destination CSV path (will be created/overwritten).
    """
    pdf_files = list(Path(input_dir).glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in: {input_dir}")
        return

    frames = []
    for pdf_path in tqdm(sorted(pdf_files), desc="Processing PDFs", unit="file"):
        try:
            df = extract_w2(str(pdf_path))
            frames.append(df)
            tqdm.write(f"  [OK]  {pdf_path.name}")
        except Exception as exc:
            tqdm.write(f"  [ERR] {pdf_path.name}: {exc}")

    if frames:
        combined = pd.concat(frames, ignore_index=True)
        combined.to_csv(output_csv, index=False)
        print(f"\nExtracted {len(combined)} record(s) -> {output_csv}")
    else:
        print("No records extracted.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _debug_lines(file_path: str) -> None:
    """
    Print the raw lines pdfplumber extracts from the first page.
    Use this to diagnose label-matching issues on a new PDF layout.
    """
    with pdfplumber.open(file_path) as pdf:
        raw = pdf.pages[0].extract_text() or ""
    lines = [ln for ln in raw.splitlines() if ln.strip()]
    print(f"\n--- Raw lines extracted from first page ({len(lines)} lines) ---")
    for i, ln in enumerate(lines):
        print(f"  [{i:03d}] {ln}")
    print("---")


if __name__ == "__main__":
    # Usage:
    #   python pythonScripts/w2_extractor.py <pdf_or_directory> [output.csv]
    #   python pythonScripts/w2_extractor.py <pdf_file> --debug   (print raw lines only)
    #
    # DO NOT pass real W-2 files containing live PII — use anonymised samples only.

    if len(sys.argv) < 2:
        print("Usage: python w2_extractor.py <pdf_file_or_dir> [output.csv|--debug]")
        sys.exit(1)

    target = sys.argv[1]

    if len(sys.argv) > 2 and sys.argv[2] == "--debug":
        _debug_lines(target)
        sys.exit(0)

    out_csv = sys.argv[2] if len(sys.argv) > 2 else "w2_extracted.csv"

    if os.path.isdir(target):
        extract_w2_batch(target, out_csv)
    else:
        result = extract_w2(target)
        result.to_csv(out_csv, index=False)
        print(result.to_string(index=False))
        print(f"\nSaved -> {out_csv}")
