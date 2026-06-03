"""
New Hire Field Extractor → Excel
Reads .txt files (output from pdf_to_txt.py) and extracts:
  - Name (Last, First)
  - Position ID

Output: new_hire_extracted.xlsx in the same folder.

Setup:
    pip install openpyxl

Run:
    python extract_fields_to_excel.py                        # prompts for folder
    python extract_fields_to_excel.py "C:\\path\\to\\txts"  # pass folder directly
"""

import os
import re
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ── Regex patterns ─────────────────────────────────────────────────────────────

# Name: "LAST, FIRST" or "LAST, FIRST MIDDLE" — all-caps format typical on HR forms
RE_NAME = re.compile(
    r"\b([A-Z][A-Z'\-]{1,}),\s+([A-Z][A-Z'\-]+(?:\s+[A-Z][A-Z'\-]+)*)\b"
)

# Position ID: common labels on new hire forms
RE_POSITION_ID = re.compile(
    r"(?:Position\s*(?:ID|No\.?|Number|#)?|Pos\.?\s*(?:ID|No\.?)?|"
    r"Job\s*(?:ID|Code|No\.?)|Requisition\s*(?:ID|No\.?)|Req\.?\s*(?:ID|No\.?)?)[\s:\-#]*"
    r"([A-Z0-9]{3,20})\b",
    re.IGNORECASE,
)


def extract(text: str) -> dict:
    result = {"last_name": "", "first_name": "", "position_id": ""}

    m = RE_NAME.search(text)
    if m:
        result["last_name"]  = m.group(1).strip()
        result["first_name"] = m.group(2).strip()

    m = RE_POSITION_ID.search(text)
    if m:
        result["position_id"] = m.group(1).strip()

    return result


def build_excel(rows: list[dict], output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "New Hires"

    headers = ["PDF File", "Last Name", "First Name", "Position ID"]

    header_fill = PatternFill("solid", fgColor="2F5496")
    header_font = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row in enumerate(rows, 2):
        ws.cell(row=row_idx, column=1, value=row["pdf_file"])
        ws.cell(row=row_idx, column=2, value=row["last_name"])
        ws.cell(row=row_idx, column=3, value=row["first_name"])
        ws.cell(row=row_idx, column=4, value=row["position_id"])

    for col, width in zip([1, 2, 3, 4], [30, 18, 22, 16]):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    ws.freeze_panes = "A2"
    wb.save(output_path)


def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1].strip()
    else:
        print("=" * 60)
        print("  New Hire Field Extractor → Excel")
        print("=" * 60)
        folder = input("\nEnter folder path containing .txt files:\n> ").strip().strip('"')

    if not os.path.isdir(folder):
        print(f"\nError: '{folder}' is not a valid folder.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    txt_files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".txt"))

    if not txt_files:
        print(f"\nNo .txt files found in:\n  {folder}")
        input("\nPress Enter to exit...")
        sys.exit(0)

    print(f"\nFolder : {folder}")
    print(f"Files  : {len(txt_files)} .txt file(s) found")
    print("-" * 60)

    rows = []
    for txt_file in txt_files:
        with open(os.path.join(folder, txt_file), "r", encoding="utf-8") as f:
            text = f.read()

        fields = extract(text)
        fields["pdf_file"] = os.path.splitext(txt_file)[0] + ".pdf"
        rows.append(fields)

        name = f"{fields['last_name']}, {fields['first_name']}".strip(", ")
        pos  = fields["position_id"] or "—"
        print(f"  {txt_file}  →  Name: {name or '?'}  |  Position ID: {pos}")

    output_path = os.path.join(folder, "new_hire_extracted.xlsx")
    build_excel(rows, output_path)

    print("\n" + "=" * 60)
    print(f"  Done!  {len(rows)} record(s) written.")
    print(f"  Output : {output_path}")
    print("=" * 60)
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
