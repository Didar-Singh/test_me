#!/usr/bin/env python3
"""
Form Field Extractor
Reads three form types and outputs one row per file with all fields as columns.

Supported formats:
  1. New Hire Form          (Lindenmeyr Munroe)
  2. Payroll Change Notice  (Lindenmeyr Munroe)
  3. ADP Pay Profile        (ADP Workforce Now)

Commands:
  python form_field_extractor.py                          process all .txt in current folder
  python form_field_extractor.py -i "C:/forms"            process all .txt in a folder
  python form_field_extractor.py -i a.txt b.txt           process specific files
  python form_field_extractor.py -o results.csv           set output filename (default: output.csv)
  python form_field_extractor.py -f table                 print table to screen instead of CSV
  python form_field_extractor.py -f json                  print JSON to screen
  python form_field_extractor.py -v                       verbose: show found/missing fields
  python form_field_extractor.py --list                   list detected .txt files only
  python form_field_extractor.py -i "C:/forms" -o results.csv -v   combine options
"""

import re
import csv
import json
import sys
import argparse
from pathlib import Path

# ── Field definitions for inline formats (Field: Value on same line) ─────────

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

# ADP format: consecutive non-empty lines form the label block,
# the next block is the value. e.g. "Affected User\nName" -> "Affected User Name"
ADP_LABEL_MAP = {
    "Event Name":           "Event Name",
    "Affected User Name":   "Name",
    "Affected Position ID": "Position ID",
    "Submit By":            "Submit By",
    "Submit On":            "Submit On",
}

ADP_SKIP_PATTERNS = [
    r"^===",
    r"ADP Workforce Now",
    r"^Review$",
    r"^> ?Print$",
    r"^Print$",
    r"Label\s+Original Value\s+New Value",
]


def detect_form_type(text):
    upper = text.upper()
    if "ADP WORKFORCE" in upper or "PAY PROFILE" in upper:
        return "ADP Pay Profile", None
    elif "NEW HIRE" in upper:
        return "New Hire Form", NEW_HIRE_FIELDS
    elif "PAYROLL CHANGE" in upper:
        return "Payroll Change Notice", PAYROLL_FIELDS
    return "Unknown", NEW_HIRE_FIELDS + PAYROLL_FIELDS


def extract_inline_fields(text, fields, verbose=False):
    all_patterns = sorted([pat for _, pat in fields], key=len, reverse=True)
    lookahead = "|".join(f"(?:{p})" for p in all_patterns)
    results = {}
    missing = []
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
            elif verbose:
                missing.append(col_name)
        elif verbose:
            missing.append(col_name)
    if verbose and missing:
        print(f"    Missing: {', '.join(missing)}")
    return results


def extract_adp_fields(text, verbose=False):
    blocks = []
    current = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped:
            current.append(stripped)
        else:
            if current:
                blocks.append(" ".join(current))
                current = []
    if current:
        blocks.append(" ".join(current))

    def is_skip(block):
        return any(re.search(p, block, re.IGNORECASE) for p in ADP_SKIP_PATTERNS)

    blocks = [b for b in blocks if not is_skip(b)]

    results = {}
    i = 0
    while i < len(blocks):
        label = blocks[i]
        if label in ADP_LABEL_MAP:
            col_name = ADP_LABEL_MAP[label]
            if i + 1 < len(blocks) and blocks[i + 1] not in ADP_LABEL_MAP:
                results[col_name] = blocks[i + 1]
                i += 2
                continue
        # Parse changes table rows e.g. "Marital Status D D" or "Additional Tax D"
        parts = label.rsplit(None, 2)
        if len(parts) >= 2 and len(parts[-1]) <= 15 and len(parts[-2]) <= 15:
            field = parts[0]
            orig  = parts[1] if len(parts) > 1 else ""
            new   = parts[2] if len(parts) > 2 else ""
            results[f"Change: {field}"] = f"{orig} -> {new}" if new else orig
        i += 1

    if verbose:
        print(f"    ADP fields found: {', '.join(results.keys()) or 'none'}")
    return results


def parse_file(path, verbose=False):
    text = path.read_text(encoding="utf-8", errors="ignore")
    form_type, fields = detect_form_type(text)
    if fields is None:
        extracted = extract_adp_fields(text, verbose=verbose)
    else:
        extracted = extract_inline_fields(text, fields, verbose=verbose)
    return {"File Number": path.stem, "Form Type": form_type, **extracted}


def collect_files(inputs):
    if not inputs:
        return sorted(Path(".").glob("*.txt"))
    paths = []
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            paths.extend(sorted(p.glob("*.txt")))
        elif p.is_file():
            paths.append(p)
        else:
            print(f"WARNING: not found, skipping: {item}")
    return paths


def build_col_order(rows):
    col_order = ["File Number", "Form Type"]
    seen = set(col_order)
    for row in rows:
        for k in row:
            if k not in seen:
                col_order.append(k)
                seen.add(k)
    return col_order


def write_csv(rows, out_path):
    col_order = build_col_order(rows)
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=col_order, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
    print(f"\nSaved: {out_path}  ({len(rows)} rows x {len(col_order)} columns)")


def print_table(rows):
    col_order = build_col_order(rows)
    col_widths = {c: len(c) for c in col_order}
    for row in rows:
        for c in col_order:
            col_widths[c] = max(col_widths[c], len(str(row.get(c, ""))))
    sep    = "+-" + "-+-".join("-" * col_widths[c] for c in col_order) + "-+"
    header = "| " + " | ".join(c.ljust(col_widths[c]) for c in col_order) + " |"
    print(sep)
    print(header)
    print(sep)
    for row in rows:
        line = "| " + " | ".join(str(row.get(c, "")).ljust(col_widths[c]) for c in col_order) + " |"
        print(line)
    print(sep)


def print_json(rows):
    print(json.dumps(rows, indent=2))


def build_parser():
    parser = argparse.ArgumentParser(
        prog="form_field_extractor",
        description="Extract fields from New Hire, Payroll Change, and ADP Pay Profile text files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python form_field_extractor.py
  python form_field_extractor.py -i "C:/forms"
  python form_field_extractor.py -i file1.txt file2.txt
  python form_field_extractor.py -i "C:/forms" -o results.csv -v
  python form_field_extractor.py -f table
  python form_field_extractor.py -f json
  python form_field_extractor.py --list
        """,
    )
    parser.add_argument(
        "-i", "--input", nargs="+", metavar="PATH",
        help="Input file(s) or folder. Defaults to all .txt in current folder.",
    )
    parser.add_argument(
        "-o", "--output", metavar="FILE", default="output.csv",
        help="Output CSV filename (default: output.csv).",
    )
    parser.add_argument(
        "-f", "--format", choices=["csv", "table", "json"], default="csv",
        help="Output format: csv (default), table (screen), json (screen).",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true",
        help="Show field counts and missing fields per file.",
    )
    parser.add_argument(
        "--list", action="store_true",
        help="List detected .txt files only — do not parse or write output.",
    )
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    files = collect_files(args.input or [])
    if not files:
        print("No .txt files found.")
        sys.exit(1)

    if args.list:
        print(f"Found {len(files)} file(s):")
        for f in files:
            print(f"  {f}")
        return

    rows = []
    for path in files:
        try:
            row = parse_file(path, verbose=args.verbose)
            rows.append(row)
            print(f"  Parsed: {path.name}  ({row['Form Type']}, {len(row) - 2} fields)")
        except Exception as e:
            print(f"  ERROR {path.name}: {e}")

    if not rows:
        print("Nothing to write.")
        sys.exit(1)

    if args.format == "csv":
        write_csv(rows, Path(args.output))
    elif args.format == "table":
        print_table(rows)
    elif args.format == "json":
        print_json(rows)


if __name__ == "__main__":
    main()
