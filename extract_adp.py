"""
ADP Employee Form Extractor
Extracts employee data from ADP personnel PDF forms into Excel.

Fields extracted per employee:
  Last Name, First Name, Mailing Address, City, State, Zip,
  File #, Dept, SSN status, Title, Hire Date, Birth Date,
  eVoucher Status, Sex, Race,
  Gross Pay, Salary, Bi-Weekly Rate, LWW, NWW, Std Hours, Pay Group,
  Federal Filing Status, State Tax, Direct Deposit Acct #, Tran/ABA,
  401K, Pre-Med, HSA, AD&D, Goal Deductions Limit, Goal Deductions To Date
"""

import re
import sys
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path


# ── helpers ──────────────────────────────────────────────────────────────────

def clean(s):
    return " ".join(s.split()) if s else ""

def first_match(pattern, text, group=1, default=""):
    m = re.search(pattern, text)
    return clean(m.group(group)) if m else default

def safe_lines(text):
    return [l.strip() for l in text.splitlines() if l.strip()]


# ── per-employee block parser ─────────────────────────────────────────────────

def parse_employee_block(block: str) -> dict:
    """Parse one employee's text block into a dict of fields."""
    rec = {}
    lines = safe_lines(block)

    # ── Name (bold, first meaningful line: LAST, FIRST or LAST FIRST)
    name_raw = ""
    for ln in lines[:6]:
        # Skip section headers
        if any(h in ln.upper() for h in ["PERSONNEL", "PAY", "TAX STATUS", "SCHEDULED"]):
            continue
        if re.search(r"[A-Z]{2,}", ln):
            name_raw = ln
            break

    if "," in name_raw:
        parts = name_raw.split(",", 1)
        rec["Last Name"] = clean(parts[0])
        rec["First Name"] = clean(parts[1])
    else:
        tokens = name_raw.split()
        rec["Last Name"] = tokens[0] if tokens else ""
        rec["First Name"] = " ".join(tokens[1:]) if len(tokens) > 1 else ""

    # ── Mailing / Home Address block
    addr_m = re.search(
        r"Mailing\s*&\s*Home\s*Address[-–\s]*(.*?)(?=File:|Dept:|eVoucher|Hire:)",
        block, re.DOTALL | re.IGNORECASE
    )
    if addr_m:
        addr_lines = safe_lines(addr_m.group(1))
        rec["Address Line 1"] = addr_lines[0] if len(addr_lines) > 0 else ""
        rec["Address Line 2"] = addr_lines[1] if len(addr_lines) > 1 else ""
        # Try to parse city/state/zip from last address line
        city_line = addr_lines[-1] if addr_lines else ""
        csz = re.search(r"^(.*?),?\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line)
        if csz:
            rec["City"] = clean(csz.group(1))
            rec["State"] = csz.group(2)
            rec["Zip"] = csz.group(3)
        else:
            rec["City/State/Zip"] = city_line
    else:
        rec["Address Line 1"] = ""
        rec["Address Line 2"] = ""

    # ── PERSONNEL section fields
    rec["File #"]   = first_match(r"File:\s*([\d\-]+)", block)
    rec["Dept"]     = first_match(r"Dept:\s*(\S+)", block)
    rec["Cntl"]     = first_match(r"Cntl:\s*(\S+)", block)
    rec["SSN"]      = first_match(r"SSN:\s*(.+?)(?:\n|Sex:|Race:)", block)
    rec["Title"]    = first_match(r"Title:\s*(\S+)", block)
    rec["Cost"]     = first_match(r"Cost:\s*(\S+)", block)
    rec["eVoucher Status"] = first_match(r"Status:\s*(\w+)", block)
    rec["Sex"]      = first_match(r"Sex:\s*(\w+)", block)
    rec["Race"]     = first_match(r"Race:\s*(\w+)", block)
    rec["Occup"]    = first_match(r"Occup:\s*(\w+)", block)

    # ── Dates
    rec["Hire Date"]  = first_match(r"Hire:\s*([\d/]+)", block)
    rec["Birth Date"] = first_match(r"Birth:\s*([\d/]+)", block)
    rec["Date 6"]     = first_match(r"Date\s*6:\s*([\d/]+)", block)
    rec["Date 9"]     = first_match(r"Date\s*9:\s*([\d/]+)", block)

    # ── PAY section
    rec["Gross Pay"]    = first_match(r"Gross:\s*([\d,\.]+)", block)
    rec["Salary"]       = first_match(r"Salary:\s*([\d,\.]+)", block)
    rec["Rate 2"]       = first_match(r"Rate\s*2:\s*([\d,\.]+)", block)
    rec["Rate Calc"]    = first_match(r"Rate\s*Calc:\s*(\w+)", block)
    rec["LWW"]          = first_match(r"LWW:\s*(\d+)", block)
    rec["NWW"]          = first_match(r"NWW:\s*(\d+)", block)
    rec["Std Hours"]    = first_match(r"Std\s*Hours:\s*([\d\.]+)", block)
    rec["Pay Group"]    = first_match(r"Pay\s*Group:\s*(\d+)", block)
    rec["Pay Frequency"]= first_match(r"(Bi-Wkly|Weekly|Semi-Monthly|Monthly)", block)

    # ── TAX STATUS
    rec["Federal Filing"] = first_match(
        r"(D-Single/Married[^,\n]*|Single[^,\n]*|Married[^,\n]*)", block)
    rec["Dependents"]     = first_match(r"Dependents\s*\$([\d]+)", block)
    rec["Extra W/H"]      = first_match(r"Extra\s*W/H\s*\$([\d,\.]+)", block)
    rec["State Tax"]      = first_match(r"(\d{2}\s+[A-Z]{2}\s+\(Lived\s+in\))", block)
    rec["GTL Coverage"]   = first_match(r"GTL\s*Cov\s*([\d,\.]+)", block)

    # ── DIRECT DEPOSITS
    rec["Acct #"]    = first_match(r"Acct\s*#\s*[:\-]?\s*([\d\-x*]+)", block, default="")
    rec["Tran/ABA"]  = first_match(r"Tran/ABA:\s*([\d\s]+)", block)
    rec["DD Code"]   = first_match(r"Code\s+([A-Z])\b", block)
    rec["DD Type"]   = first_match(r"(Full\s+Deposit|Partial\s+Deposit)", block)

    # ── SCHEDULED AMOUNTS (pick common ones)
    rec["401K"]      = first_match(r"K\s*401K\s+([\d,\.]+)", block)
    rec["Pre-Med"]   = first_match(r"35\s*PREMED\s+([\d,\.]+)", block)
    rec["ADDLCH"]    = first_match(r"42\s*ADDLCH\s+([\d,\.]+)", block)
    rec["AD&D"]      = first_match(r"57\s*AD&?D\s+([\d,\.]+)", block)
    rec["HSA"]       = first_match(r"HSA\s+HCCACT\s+([\d,\.]+)", block)
    rec["Goal Limit"]    = first_match(r"Limit:\s*([\d,\.]+)", block)
    rec["Goal To Date"]  = first_match(r"To\s*Date:\s*([\d,\.]+)", block)

    return rec


# ── PDF reader ────────────────────────────────────────────────────────────────

EMPLOYEE_SPLIT_RE = re.compile(
    r"(?=PERSONNEL\s*\n)", re.IGNORECASE
)

def extract_employees(pdf_path: str) -> list[dict]:
    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text(layout=True) or ""
            all_text += t + "\n\n"

    # Split on each PERSONNEL header (one block per employee)
    blocks = EMPLOYEE_SPLIT_RE.split(all_text)
    employees = []
    for block in blocks:
        block = block.strip()
        if not block or "PERSONNEL" not in block.upper():
            continue
        try:
            rec = parse_employee_block(block)
            if rec.get("Last Name") or rec.get("First Name"):
                employees.append(rec)
        except Exception as e:
            print(f"  [warn] Skipped a block: {e}", file=sys.stderr)

    return employees


# ── Excel writer ──────────────────────────────────────────────────────────────

COLUMNS = [
    "Last Name", "First Name",
    "Address Line 1", "Address Line 2", "City", "State", "Zip",
    "File #", "Dept", "Cntl", "SSN", "Title", "Cost",
    "eVoucher Status", "Sex", "Race", "Occup",
    "Hire Date", "Birth Date", "Date 6", "Date 9",
    "Gross Pay", "Salary", "Rate 2", "Rate Calc",
    "LWW", "NWW", "Std Hours", "Pay Group", "Pay Frequency",
    "Federal Filing", "Dependents", "Extra W/H", "State Tax", "GTL Coverage",
    "Acct #", "Tran/ABA", "DD Code", "DD Type",
    "401K", "Pre-Med", "ADDLCH", "AD&D", "HSA",
    "Goal Limit", "Goal To Date",
]

HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
ALT_FILL     = PatternFill("solid", fgColor="D9E1F2")
NORMAL_FONT  = Font(name="Arial", size=10)
BORDER_SIDE  = Side(style="thin", color="BFBFBF")
CELL_BORDER  = Border(left=BORDER_SIDE, right=BORDER_SIDE,
                      top=BORDER_SIDE, bottom=BORDER_SIDE)

def write_excel(employees: list[dict], out_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ADP Employee Data"

    # Header row
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER
    ws.row_dimensions[1].height = 30

    # Data rows
    for row_idx, emp in enumerate(employees, start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else None
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            val = emp.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font   = NORMAL_FONT
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="center")
            if fill:
                cell.fill = fill

    # Column widths
    col_widths = {
        "Last Name": 18, "First Name": 18,
        "Address Line 1": 28, "Address Line 2": 20,
        "City": 18, "State": 8, "Zip": 10,
        "File #": 12, "Dept": 10, "Cntl": 8,
        "SSN": 12, "Title": 14, "Cost": 22,
        "eVoucher Status": 14, "Sex": 6, "Race": 6, "Occup": 8,
        "Hire Date": 12, "Birth Date": 12, "Date 6": 12, "Date 9": 12,
        "Gross Pay": 12, "Salary": 12, "Rate 2": 10, "Rate Calc": 10,
        "LWW": 8, "NWW": 8, "Std Hours": 10, "Pay Group": 10, "Pay Frequency": 14,
        "Federal Filing": 28, "Dependents": 12, "Extra W/H": 10,
        "State Tax": 22, "GTL Coverage": 14,
        "Acct #": 18, "Tran/ABA": 16, "DD Code": 8, "DD Type": 14,
        "401K": 10, "Pre-Med": 10, "ADDLCH": 10, "AD&D": 8, "HSA": 10,
        "Goal Limit": 12, "Goal To Date": 14,
    }
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 14)

    # Freeze header + enable autofilter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(out_path)
    print(f"✅ Saved {len(employees)} employees → {out_path}")


# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_adp.py <input.pdf> [output.xlsx]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else str(
        Path(pdf_path).stem + "_extracted.xlsx"
    )

    print(f"📄 Reading: {pdf_path}")
    employees = extract_employees(pdf_path)
    print(f"👥 Found {len(employees)} employee record(s)")

    if not employees:
        print("⚠️  No records found. Check PDF text layer (may be scanned).")
        sys.exit(1)

    write_excel(employees, out_path)
