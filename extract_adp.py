"""
ADP Master Control Report - Extractor (ROBUST VERSION)
========================================================
Fixes:
  ✓ Captures ALL employees per page (not just last)
  ✓ Captures full multi-line addresses
  ✓ Captures all account numbers
  ✓ Captures gross/net salary completely

Usage:
  python extract_adp.py yourfile.pdf
  python extract_adp.py yourfile.pdf output.xlsx
  python extract_adp.py yourfile.pdf --pages 17-20
  python extract_adp.py yourfile.pdf --pages 17-20 --debug
"""

import re, sys
import pdfplumber
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from pathlib import Path
from tqdm import tqdm

DEBUG = "--debug" in sys.argv
PAGE_FROM = PAGE_TO = None

# Parse --pages flag - handle: --pages 17-25, --pages=17-25, --pages 17 25
for i, arg in enumerate(sys.argv):
    if "--pages" in arg:
        if "=" in arg:
            val = arg.split("=")[1]
        elif i + 1 < len(sys.argv):
            val = sys.argv[i + 1]
        else:
            val = ""
        
        m = re.match(r"(\d+)[-\s:](\d+)", val)
        if m:
            PAGE_FROM, PAGE_TO = int(m.group(1)), int(m.group(2))
            break

args = [a for a in sys.argv[1:] if not a.startswith("--")]

# ── helpers ───────────────────────────────────────────────────────────────────

def clean(s):
    return " ".join(str(s).split()).strip() if s else ""

def grep(pattern, text, group=1, default="", flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return clean(m.group(group)) if m else default

def mgrep(patterns, text, default=""):
    for p in (patterns if isinstance(patterns, list) else [patterns]):
        v = grep(p, text, default="")
        if v: return v
    return default

# ── find ALL employee names in text ───────────────────────────────────────────

def find_all_names(text):
    """
    Find ALL employee name positions in text.
    Name pattern: LAST,FIRST or LAST FIRST (all caps, on own line or end of line)
    """
    name_positions = []
    
    # Pattern 1: Name on its own line
    for m in re.finditer(r"^([A-Z][A-Z'\-\.]+),([A-Z][A-Z'\-\.\s]+)$", text, re.MULTILINE):
        name_positions.append((m.start(), m.group(0)))
    
    # Pattern 2: Name at END of a line (after other text like coverage line)
    for m in re.finditer(r"\n([A-Z][A-Z'\-\.]+),([A-Z][A-Z'\-\.\s\.]+)\s*$", text, re.MULTILINE):
        name_positions.append((m.start(1), m.group(1) + "," + m.group(2)))
    
    # Sort by position and remove duplicates
    name_positions = sorted(set(name_positions))
    
    if DEBUG:
        print(f"\n[FOUND {len(name_positions)} NAMES]")
        for pos, name in name_positions:
            print(f"  @ {pos:5d}: {name}")
    
    return name_positions

def split_by_names(text):
    """Split text into blocks, one per employee name found."""
    names = find_all_names(text)
    if not names:
        return []
    
    blocks = []
    for i, (pos, name) in enumerate(names):
        start = pos
        end   = names[i+1][0] if i+1 < len(names) else len(text)
        blocks.append(text[start:end].strip())
    
    return blocks

# ── parse one employee block ──────────────────────────────────────────────────

def parse_block(block, block_num=1):
    rec = {"_block": block_num}
    txt = block
    lines = [l.strip() for l in block.splitlines()]
    
    if DEBUG:
        print(f"\n{'─'*70}")
        print(f"BLOCK {block_num} ({len(lines)} lines)")
        print(f"{'─'*70}")
        print(txt[:600])
        print(f"{'─'*70}")

    # ── NAME ──────────────────────────────────────────────────────────────────
    first_line = lines[0] if lines else ""
    if "," in first_line:
        parts = first_line.split(",", 1)
        rec["Last Name"]  = clean(parts[0])
        rec["First Name"] = clean(parts[1])
    else:
        toks = first_line.split()
        rec["Last Name"]  = " ".join(toks[:-1]) if len(toks) > 1 else first_line
        rec["First Name"] = toks[-1] if len(toks) > 1 else ""

    rec["Continued"] = "Yes" if re.search(r"\(continued\)", txt, re.I) else ""

    # ── FILE # (first one) ────────────────────────────────────────────────────
    file_m = re.search(r"File:\s*(\d+)", txt, re.I)
    rec["File #"] = file_m.group(1) if file_m else ""

    # ── PERSONNEL FIELDS ──────────────────────────────────────────────────────
    rec["Status"]   = grep(r"Status:\s*(\w+)", txt)
    rec["Dept"]     = grep(r"Dept:\s*(\S+)", txt)
    rec["Sex"]      = grep(r"Sex:\s*(\w)", txt)
    rec["Cntl"]     = grep(r"Cntl:\s*(\S+)", txt)
    rec["Race"]     = grep(r"Race:\s*([^\s\n]+)", txt)
    rec["SSN"]      = grep(r"SSN:\s*([^\n]+?)(?=\s{2,}|Occup|\n)", txt)
    rec["Occup"]    = grep(r"Occup:\s*([^\s\n]+)", txt)
    rec["Title"]    = grep(r"Title:\s*(\S+)", txt)
    rec["eVoucher"] = "Yes" if re.search(r"eVoucher", txt, re.I) else ""
    rec["Qualified Pension"] = "Yes" if re.search(r"Qualified\s+Pension", txt, re.I) else ""
    rec["Health Coverage"]   = grep(r"(Employee\s*[&and]+\s*Dependents[^\n]+)", txt)

    # ── COST CENTER (may span multiple lines) ─────────────────────────────────
    cost_m = re.search(
        r"Cost:\s*\n([\s\S]*?)(?=\nDates|\nHire:|\neVoucher|\nGross:|\nMailing|\n\n)",
        txt, re.I
    )
    if cost_m:
        rec["Cost"] = clean(cost_m.group(1).replace("\n", " "))
    else:
        rec["Cost"] = grep(r"Cost:\s*([^\n]+)", txt)

    # ── DATES ─────────────────────────────────────────────────────────────────
    rec["Hire Date"]  = grep(r"Hire:\s*([\d/]+)", txt)
    rec["Term Date"]  = grep(r"Term:\s*([\d/]+)", txt)
    rec["Birth Date"] = grep(r"Birth:\s*([\d/]+)", txt)
    rec["Date 1"]     = grep(r"Date\s+1:\s*([\d/]+)", txt)
    rec["Date 3"]     = grep(r"Date\s+3:\s*([\d/]+)", txt)
    rec["Date 6"]     = grep(r"Date\s+6:\s*([\d/]+)", txt)
    rec["Date 8"]     = grep(r"Date\s+8:\s*([\d/]+)", txt)
    rec["Date 9"]     = grep(r"Date\s+9:\s*([\d/]+)", txt)

    # ── ADDRESS (FULL multi-line capture) ─────────────────────────────────────
    addr_m = re.search(
        r"Mailing\s*[&]\s*Home\s*Address\s*\n(.*?)(?=\nGross:|\nSalary:|\nPAY:|\nTAX:|\n\n\n|\nPayroll|\Z)",
        txt, re.DOTALL | re.I
    )
    if addr_m:
        addr_block = addr_m.group(1).strip()
        addr_lines = [l.strip() for l in addr_block.splitlines() if l.strip()]
        rec["Address Line 1"] = addr_lines[0] if len(addr_lines) > 0 else ""
        rec["Address Line 2"] = addr_lines[1] if len(addr_lines) > 1 else ""
        rec["Address Line 3"] = addr_lines[2] if len(addr_lines) > 2 else ""
        
        # Parse city, state, zip from LAST address line
        city_line = addr_lines[-1] if addr_lines else ""
        # Try "City,ST ZIP"
        csz = re.search(r"^(.*?),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line, re.I)
        if csz:
            rec["City"]  = clean(csz.group(1))
            rec["State"] = csz.group(2).upper()
            rec["Zip"]   = csz.group(3)
        else:
            # Try "City ST ZIP" (no comma)
            csz2 = re.search(r"^(.*?)\s+([A-Z]{2})\s+(\d{5})$", city_line, re.I)
            if csz2:
                rec["City"]  = clean(csz2.group(1))
                rec["State"] = csz2.group(2).upper()
                rec["Zip"]   = csz2.group(3)
            else:
                rec["City/State/Zip"] = city_line
    else:
        rec["Address Line 1"] = rec["Address Line 2"] = rec["Address Line 3"] = ""

    # ── PAY (GREEDY - capture full numbers) ───────────────────────────────────
    # Gross may span multiple lines or have spaces
    rec["Gross"]     = mgrep([
        r"Gross:\s*([\d,\.]+(?:\s*[\d,\.]+)?)",
        r"Gross\s+([\d,\.]+)"
    ], txt)
    rec["Salary"]    = mgrep([
        r"Salary:\s*([\d,\.]+(?:\s*[\d,\.]+)?)",
        r"Salary\s+([\d,\.]+)"
    ], txt)
    rec["Bi-Wkly"]   = grep(r"Bi-Wkly\s*[:\s]*([\d,\.]+)", txt)
    rec["Rate Calc"] = grep(r"Rate\s*Calc:\s*(\w+)", txt)
    rec["LWW"]       = grep(r"LWW:\s*(\d+)", txt)
    rec["NWW"]       = grep(r"NWW:\s*(\d+)", txt)
    rec["Std Hours"] = grep(r"Std\s*Hours:\s*([\d\.]+)", txt)
    rec["Pay Group"] = grep(r"Pay\s*Group:\s*(\d+)", txt)
    
    # Net Pay (may be on its own line)
    rec["Net Pay"]   = mgrep([
        r"Net\s*Pay:\s*([\d,\.]+)",
        r"Net:\s*([\d,\.]+)"
    ], txt)

    # ── TAX ───────────────────────────────────────────────────────────────────
    rec["Marital Status"]     = grep(r"Marital\s+Status:\s*([^\n]+)", txt)
    rec["Federal Exemptions"] = mgrep([
        r"Federal[:\s]+(\d+)\s*Exemptions?",
        r"(\d+)\s*Exemptions?\s*Federal",
    ], txt)
    rec["Federal Extra W/H"]  = grep(r"Extra\s*W[/\\]H\s*\$?([\d,\.]+)", txt)

    # State lines
    state_lines = re.findall(
        r"^\s*(\d{2,4}[A-Z]?\s+[A-Z]{2}[\w\s\-]*?)$",
        txt, re.MULTILINE
    )
    state_lines = [clean(s) for s in state_lines if len(clean(s)) > 3][:4]
    rec["State Tax 1"] = state_lines[0] if len(state_lines) > 0 else ""
    rec["State Tax 2"] = state_lines[1] if len(state_lines) > 1 else ""
    rec["State Tax 3"] = state_lines[2] if len(state_lines) > 2 else ""
    rec["State Tax 4"] = state_lines[3] if len(state_lines) > 3 else ""

    # ── SCHEDULED AMOUNTS ─────────────────────────────────────────────────────
    rec["401K"]      = mgrep([r"K\s*401K\s+([\d,\.]+)", r"401K\s+([\d,\.]+)"], txt)
    rec["Pre-Med"]   = mgrep([r"35\s*PREMED\s+([\d,\.]+)", r"PREMED\s+([\d,\.]+)"], txt)
    rec["ADDLCH"]    = mgrep([r"42\s*ADDLCH\s+([\d,\.]+)", r"ADDLCH\s+([\d,\.]+)"], txt)
    rec["AD&D"]      = mgrep([r"57\s*AD&?D\s+([\d,\.]+)", r"AD&?D\s+([\d,\.]+)"], txt)
    rec["HSA"]       = mgrep([r"HSA\s+HCCACT\s+([\d,\.]+)", r"HSA\s+([\d,\.]+)"], txt)
    rec["Goal Limit"]    = grep(r"Limit:\s*([\d,\.]+)", txt)
    rec["Goal To Date"]  = grep(r"To\s*Date:\s*([\d,\.]+)", txt)

    # ── DIRECT DEPOSIT (ALL patterns) ─────────────────────────────────────────
    # Account number — try ALL variations
    rec["Acct #"]   = mgrep([
        r"Acct\s*#\s*[:\-]?\s*([\dXx*\-]+)",
        r"Account\s*#\s*[:\-]?\s*([\dXx*\-]+)",
        r"Account:\s*([\dXx*\-]+)",
        r"Acct:\s*([\dXx*\-]+)"
    ], txt)
    
    rec["Tran/ABA"] = mgrep([
        r"Tran[/\\]ABA:\s*([\d\s]+)",
        r"ABA[:\s]+([\d\s]+)",
        r"Routing:\s*([\d\s]+)"
    ], txt)
    rec["DD Code"]  = grep(r"Code\s+([A-Z])\b", txt)
    rec["DD Type"]  = mgrep([r"(Full\s+Deposit)", r"(Partial\s+Deposit)"], txt)

    if DEBUG:
        print(f"\nPARSED:")
        for k, v in rec.items():
            if v and not k.startswith('_'):
                print(f"  {k:22s}: {v}")

    return rec

# ── extract all ──────────────────────────────────────────────────────────────

def extract_all(pdf_path, page_from=None, page_to=None):
    employees = []

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        p_from = max(1, page_from) if page_from else 1
        p_to   = min(total, page_to) if page_to else total

        if p_from > total or p_from > p_to:
            print(f"\n  ❌ Invalid page range {p_from}-{p_to} (PDF has {total} pages)")
            return []

        page_indices = range(p_from - 1, p_to)
        print(f"  Total pages in PDF : {total}")
        if page_from or page_to:
            print(f"  Processing pages   : {p_from} to {p_to}  ({len(page_indices)} page(s))")
        else:
            print(f"  Processing pages   : all {total}")

        all_text = ""
        with tqdm(total=len(page_indices), desc="📄 Reading",
                  unit="pg", colour="cyan", ncols=65) as pbar:
            for page_num in page_indices:
                page = pdf.pages[page_num]
                t = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                all_text += t + "\n"
                pbar.update(1)

    if not all_text.strip():
        print("\n  ❌ No text found. File may be scanned.")
        return []

    blocks = split_by_names(all_text)
    print(f"  Employee blocks    : {len(blocks)}")

    if not blocks:
        print("\n  ⚠️  No employee blocks found.")
        return []

    with tqdm(total=len(blocks), desc="👤 Parsing",
              unit="emp", colour="green", ncols=65) as pbar:
        for i, block in enumerate(blocks):
            try:
                rec = parse_block(block, block_num=i+1)
                if rec.get("Last Name") or rec.get("File #"):
                    employees.append(rec)
            except Exception as e:
                tqdm.write(f"  ⚠️  Block {i+1}: {e}")
            pbar.update(1)

    return employees

# ── Excel writer (plain) ──────────────────────────────────────────────────────

COLUMNS = [
    "Last Name", "First Name", "Continued",
    "File #", "Status", "Dept", "Sex", "Cntl", "Race", "SSN", "Occup", "Title",
    "eVoucher", "Cost", "Qualified Pension", "Health Coverage",
    "Hire Date", "Term Date", "Birth Date", "Date 1", "Date 3", "Date 6", "Date 8", "Date 9",
    "Address Line 1", "Address Line 2", "Address Line 3", "City", "State", "Zip", "City/State/Zip",
    "Gross", "Salary", "Bi-Wkly", "Rate Calc", "LWW", "NWW", "Std Hours", "Pay Group", "Net Pay",
    "Marital Status", "Federal Exemptions", "Federal Extra W/H", "State Tax 1", "State Tax 2", "State Tax 3", "State Tax 4",
    "401K", "Pre-Med", "ADDLCH", "AD&D", "HSA", "Goal Limit", "Goal To Date",
    "Acct #", "Tran/ABA", "DD Code", "DD Type",
    "_block",
]

def write_excel(employees, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ADP Master Control"

    # Header
    for ci, col in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font = Font(bold=True, name="Arial", size=10)

    # Data
    with tqdm(total=len(employees), desc="💾 Excel",
              unit="row", colour="yellow", ncols=65) as pbar:
        for ri, emp in enumerate(employees, 2):
            for ci, col in enumerate(COLUMNS, 1):
                ws.cell(row=ri, column=ci, value=emp.get(col, ""))
            pbar.update(1)

    # Auto-width
    for ci, col in enumerate(COLUMNS, 1):
        max_len = len(col)
        for ri in range(2, min(len(employees)+2, 52)):
            val = ws.cell(row=ri, column=ci).value or ""
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 2, 45)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"
    wb.save(out_path)

# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(args) < 1:
        print(__doc__)
        sys.exit(0)

    pdf_path = args[0]
    out_path = args[1] if len(args) > 1 else str(Path(pdf_path).stem + "_employees.xlsx")

    if not Path(pdf_path).exists():
        print(f"\n❌ File not found: {pdf_path}")
        sys.exit(1)

    flags = []
    if PAGE_FROM or PAGE_TO: flags.append(f"pages {PAGE_FROM}-{PAGE_TO}")
    if DEBUG: flags.append("DEBUG")

    print(f"\n{'='*55}")
    print(f"  ADP Master Control Extractor")
    print(f"{'='*55}")
    print(f"  Input : {pdf_path}")
    print(f"  Output: {out_path}")
    if flags: print(f"  Flags : {' | '.join(flags)}")
    print(f"{'='*55}")

    employees = extract_all(pdf_path, page_from=PAGE_FROM, page_to=PAGE_TO)
    print(f"\n  👥 Records: {len(employees)}")

    if not employees:
        print("\n❌ Nothing extracted. Try:")
        print(f"   python extract_adp.py \"{pdf_path}\" --pages 17-18 --debug\n")
        sys.exit(1)

    write_excel(employees, out_path)

    print(f"\n{'='*55}")
    print(f"  ✅ Done! → {out_path}")
    print(f"{'='*55}\n")
