#!/usr/bin/env python3
"""
Form Field Extractor
Reads New Hire Form and Payroll Change Notice .txt files.
Outputs one row per file with all extracted fields as columns.

Usage:
  python form_field_extractor.py                  # process all .txt in current folder
  python form_field_extractor.py folder/path      # process all .txt in given folder
  python form_field_extractor.py a.txt b.txt ...  # process specific files

Output: output.csv  (File Number | all fields as columns)
"""

import re
import csv
import sys
from pathlib import Path

# ── Field definitions: (column_header, regex_pattern) ───────────────────────
# Patterns are intentionally flexible to handle OCR noise / slight variations.

NEW_HIRE_FIELDS = [
    ("Employee Name",          r"Employee Name"),
    ("Today's Date",           r"Today.s Date"),
    ("Social Security Number", r"Social Security Number"),
    ("Date of Hire",           r"Date of Hire"),
    ("Street Address",         r"Street Address"),
    ("City, State, Zip",       r"City,?\s*State,?\s*Zip"),
    ("Telephone No",           r"Telephone No"),
    ("Cell",                   r"Cell"),
    ("File No",                r"File No"),
    ("Email Address",          r"Email Address"),
    ("Badge No",               r"Badge No"),
    ("SSN Verified",           r"SSN Verified"),
    ("Gender",                 r"Gender"),
    ("Date of Birth",          r"Date of Birth"),
    ("Marital Status",         r"Marital Status"),
    ("Race",                   r"Race"),
    ("Office Location",        r"Office Location"),
    ("Annual Salary",          r"Annual Salary"),
    ("Department No",          r"Department No"),
    ("Hourly Rate",            r"Hourly Rate"),
    ("Job Title",              r"Job Title"),
    ("Standard Hours",         r"Standard Hours"),
    ("Reports To",             r"Reports To"),
    ("Employee Type",          r"Employee Type"),
    ("E-Time Approver",        r"E-?Time Approver"),
    ("Pay Cycle",              r"Pay Cycle"),
    ("Union Code",             r"Union Code"),
    ("Car Allowance",          r"Car Allowance"),
    ("EEO-1 Code",             r"EEO-1 Code"),
    ("FLSA Overtime",          r"FLSA Overtime"),
    ("Job Group",              r"Job Group"),
    ("Vacation Entitlement",   r"Vacation Entitlement"),
    ("Harassment Prevention",  r"Harassment Prevention"),
    ("Current Year Vacation",  r"Current Year Vacation"),
    ("Training",               r"Training"),
]

PAYROLL_FIELDS = [
    ("Employee Name",           r"Employee Name"),
    ("Today's Date",            r"Today.s Date"),
    ("Office Location",         r"Office Location"),
    ("File #",                  r"File #"),
    ("FLSA Effective Date",     r"FLSAD? Effective Date"),
    ("Job Title",               r"Job Title"),
    ("Reports To",              r"Reports To"),
    ("Employee Type",           r"Empl(?:oyee)? Type"),
    ("Department",              r"Department"),
    ("Car Allowance",           r"Car Allowance"),
    ("Current Rate",            r"Current Rate"),
    ("Reason for Change",       r"Reason for Change"),
    ("Future Rate",             r"Future Rate"),
    ("EEO-1 Code",              r"EEO-1 Code"),
    ("Job Group",               r"Job Group"),
    ("Pay Cycle",               r"Pay Cycle"),
    ("Standard Hours",          r"Standard Hours"),
    ("Period Start",            r"(?:Additional Earnings )?Period Start"),
    ("Period End",              r"Period End"),
    ("Amount",                  r"Amount"),
    ("Leave of Absence Reason", r"Leave of Absence Reason(?:\s+for LOA)?"),
]


def detect_form_type(text: str):
    upper = text.upper()
    if "NEW HIRE" in upper:
        return "New Hire Form", NEW_HIRE_FIELDS
    elif "PAYROLL CHANGE" in upper:
        return "Payroll Change Notice", PAYROLL_FIELDS
    return "Unknown", NEW_HIRE_FIELDS + PAYROLL_FIELDS


def extract_fields(text: str, fields: list) -> dict:
    """
    Extracts field:value pairs. Handles multiple fields on the same line
    by using a lookahead that stops the value at the next known field name.
    """
    all_patterns = sorted([pat for _, pat in fields], key=len, reverse=True)
    lookahead = "|".join(f"(?:{p})" for p in all_patterns)

    results = {}
    for col_name, field_pattern in fields:
        if col_name in results:
            continue

        pattern = (
            rf"(?:{field_pattern})"
            rf"\s*[:\-]\s*"
            rf"([^\n]*?)"
            rf"(?=\s*(?:{lookahead})\s*[:\-]|\s*$)"
        )
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            value = match.group(1).strip().strip("|").strip()
            if value:
                results[col_name] = value

    return results


def parse_file(path: Path) -> dict:
    text = path.read_text(encoding="utf-8", errors="ignore")
    form_type, fields = detect_form_type(text)
    extracted = extract_fields(text, fields)
    return {
        "File Number": path.stem,   # filename without extension
        "Form Type":   form_type,
        **extracted,
    }


def collect_files(args: list) -> list:
    """Resolve CLI args to a list of .txt file paths."""
    if not args:
        return sorted(Path(".").glob("*.txt"))

    paths = []
    for arg in args:
        p = Path(arg)
        if p.is_dir():
            paths.extend(sorted(p.glob("*.txt")))
        elif p.is_file():
            paths.append(p)
        else:
            print(f"WARNING: not found, skipping: {arg}", flush=True)
    return paths


def main():
    args = [a for a in sys.argv[1:] if a not in ("-h", "--help")]

    if "-h" in sys.argv or "--help" in sys.argv:
        print(__doc__)
        return

    files = collect_files(args)
    if not files:
        print("No .txt files found.")
        return

    rows = []
    for path in files:
        try:
            row = parse_file(path)
            rows.append(row)
            print(f"  Parsed: {path.name}  ({row['Form Type']}, {len(row)-2} fields)")
        except Exception as e:
            print(f"  ERROR {path.name}: {e}")

    if not rows:
        print("Nothing to write.")
        return

    # Build column order: File Number + Form Type + all unique field columns
    # in the order they first appear across all rows
    col_order = ["File Number", "Form Type"]
    seen = set(col_order)
    for row in rows:
        for k in row:
            if k not in seen:
                col_order.append(k)
                seen.add(k)

    out_path = Path("output.csv")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=col_order, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    print(f"\nSaved: {out_path}  ({len(rows)} rows x {len(col_order)} columns)")


if __name__ == "__main__":
    main()
