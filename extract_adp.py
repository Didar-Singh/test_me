"""
ADP Master Control Report - Extractor
======================================
Parses ADP Master Control PDFs where text reads linearly (top to bottom).

Each employee block structure (from raw text):
  LAST,FIRST                         <- Name line
  File: XXXXXX                       <- may be followed by (continued)
  eVoucher                           <- optional marker
  File: NNN  Status: ACTIVE/TERM
  Dept: NNN  Sex: M/F
  Cntl: NNN  Race: N
  SSN: On File  Occup: N
  Title: XXXXX
  Cost:
  COST-CENTER-VALUE
  Dates
  Hire: MM/DD/YYYY  Term: MM/DD/YYYY
  Birth: MM/DD/YY   Date 6: MM/DD/YYYY
  Date 8: MM/DD/YYYY
  Qualified Pension                  <- optional
  Employee & Dependents ...          <- optional coverage line
  SUPERVISOR,NAME                    <- optional
  Mailing & Home Address
  123 STREET ST
  City,ST ZIPCODE
  --- PAY section ---
  Gross: N.NN  Salary: N.NN  Bi-Wkly  Rate Calc: N  LWW: NN  NWW: NN
  --- TAX section ---
  Marital Status: S-SINGLE / M-MARRIED
  Federal: NN Exemptions
  59 PA SUIDI  / state lines
  ...

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 COMMANDS & EXAMPLES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

 1. Basic — process entire PDF, auto-name output:
       python extract_adp.py myfile.pdf

 2. Custom output filename:
       python extract_adp.py myfile.pdf results.xlsx

 3. Process specific page range (e.g. pages 2 to 5):
       python extract_adp.py myfile.pdf --pages 2-5

 4. Process specific page range + custom output:
       python extract_adp.py myfile.pdf results.xlsx --pages 2-5

 5. Single page test (e.g. just page 3):
       python extract_adp.py myfile.pdf --pages 3-3

 6. Debug mode — prints raw text + parsed fields per employee:
       python extract_adp.py myfile.pdf --debug

 7. Debug on a page range (best for troubleshooting):
       python extract_adp.py myfile.pdf --pages 1-2 --debug

 8. Install required libraries (run once):
       pip install pdfplumber openpyxl tqdm

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 FLAGS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  --pages FROM-TO   Only process pages FROM through TO (1-based)
  --debug           Print raw text and parsed fields per employee
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import re, sys
import pdfplumber
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from pathlib import Path
from tqdm import tqdm

DEBUG = "--debug" in sys.argv

# Parse --pages FROM-TO
PAGE_FROM = PAGE_TO = None
for arg in sys.argv[1:]:
    m = re.match(r"--pages[=:]?(\d+)[-:](\d+)$", arg)
    if m:
        PAGE_FROM, PAGE_TO = int(m.group(1)), int(m.group(2))
        break
    m2 = re.match(r"--pages[=:]?(\d+)$", arg)
    if m2:
        PAGE_FROM = PAGE_TO = int(m2.group(1))
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

# ── split full text into per-employee blocks ──────────────────────────────────

# Name line pattern: "LASTNAME,FIRSTNAME" or "LAST FIRST" — all caps, no digits
NAME_RE = re.compile(
    r"^([A-Z][A-Z'\-\.]+),([A-Z][A-Z'\-\.\s]+)$",
    re.MULTILINE
)

def split_into_blocks(text):
    """
    Split full PDF text into one block per employee.
    Handles two cases:
      1. Name on its own line:  "SINGH,DIDAR"
      2. Name at end of line:   "Employee & Dependents Health Care Coverage KUMAR,MOHIT M."
    In case 2, we normalise the text so the name starts on a new line before splitting.
    """
    # Normalise: if a NAME pattern appears mid-line (after other text),
    # insert a newline before it so NAME_RE can find it.
    # Pattern: any WORD text, then space, then LAST,FIRST at end of line
    INLINE_NAME_RE = re.compile(
        r"([^\n]+?)\s+([A-Z][A-Z'\-\.]{1,},[A-Z][A-Z'\-\.\s]{1,})$",
        re.MULTILINE
    )
    def _split_inline(m):
        prefix = m.group(1).strip()
        name   = m.group(2).strip()
        # Only split if the prefix is NOT itself a bare name line
        if re.match(r"^[A-Z][A-Z'\-\.,\s]{2,}$", prefix):
            return m.group(0)   # already a name line, leave alone
        return prefix + "\n" + name
    text = INLINE_NAME_RE.sub(_split_inline, text)

    matches = list(NAME_RE.finditer(text))
    if not matches:
        return []

    blocks = []
    for i, m in enumerate(matches):
        start = m.start()
        end   = matches[i+1].start() if i+1 < len(matches) else len(text)
        blocks.append(text[start:end].strip())

    return blocks

# ── parse one employee block ──────────────────────────────────────────────────

def parse_block(block):
    rec = {}
    lines = [l.strip() for l in block.splitlines()]
    # Full block as single searchable string
    txt = block

    if DEBUG:
        print(f"\n{'═'*60}")
        print(f"RAW BLOCK:\n{block[:800]}")
        print(f"{'─'*60}")

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

    # ── FILE # ────────────────────────────────────────────────────────────────
    # Two File: lines possible — first is the display file#, second has Status
    file_nums = re.findall(r"File:\s*(\d+)", txt, re.I)
    rec["File #"] = file_nums[0] if file_nums else ""

    # ── PERSONNEL FIELDS ──────────────────────────────────────────────────────
    rec["Status"]   = grep(r"Status:\s*(\w+)", txt)
    rec["Dept"]     = grep(r"Dept:\s*(\S+)", txt)
    rec["Sex"]      = grep(r"Sex:\s*(\w)", txt)
    rec["Cntl"]     = grep(r"Cntl:\s*(\S+)", txt)
    rec["Race"]     = grep(r"Race:\s*(\w+)", txt)
    rec["SSN"]      = grep(r"SSN:\s*([^\n]+?)(?=\s{2,}|\s+Occup|\n)", txt)
    rec["Occup"]    = grep(r"Occup:\s*(\w+)", txt)
    rec["Title"]    = grep(r"Title:\s*(\S+)", txt)
    rec["eVoucher"] = "Yes" if re.search(r"eVoucher", txt, re.I) else ""
    rec["Qualified Pension"] = "Yes" if re.search(r"Qualified\s+Pension", txt, re.I) else ""
    rec["Health Coverage"]   = grep(r"(Employee\s*[&and]+\s*Dependents[^\n]+)", txt)

    # ── COST CENTER ───────────────────────────────────────────────────────────
    # "Cost:" on one line, value on next line(s) until "Dates"
    cost_m = re.search(r"Cost:\s*\n(.*?)(?=\nDates|\nHire:|\neVoucher)", txt, re.DOTALL | re.I)
    if cost_m:
        rec["Cost"] = clean(cost_m.group(1).replace("\n", " "))
    else:
        rec["Cost"] = grep(r"Cost:\s*([^\n]+)", txt)

    # ── DATES ─────────────────────────────────────────────────────────────────
    rec["Hire Date"]  = grep(r"Hire:\s*([\d/]+)", txt)
    rec["Term Date"]  = grep(r"Term:\s*([\d/]+)", txt)
    rec["Birth Date"] = grep(r"Birth:\s*([\d/]+)", txt)
    rec["Date 6"]     = grep(r"Date\s*6:\s*([\d/]+)", txt)
    rec["Date 8"]     = grep(r"Date\s*8:\s*([\d/]+)", txt)
    rec["Date 9"]     = grep(r"Date\s*9:\s*([\d/]+)", txt)
    rec["Date 1"]     = grep(r"Date\s*1:\s*([\d/]+)", txt)
    rec["Date 3"]     = grep(r"Date\s*3:\s*([\d/]+)", txt)

    # ── ADDRESS ───────────────────────────────────────────────────────────────
    addr_m = re.search(
        r"Mailing\s*[&]\s*Home\s*Address\s*\n(.*?)(?=\n\n|\nGross:|\nSalary:|\nPAY|\nTAX|\Z)",
        txt, re.DOTALL | re.I
    )
    if addr_m:
        addr_lines = [l.strip() for l in addr_m.group(1).splitlines() if l.strip()]
        rec["Address Line 1"] = addr_lines[0] if len(addr_lines) > 0 else ""
        rec["Address Line 2"] = addr_lines[1] if len(addr_lines) > 1 else ""
        # city,state zip — last address line
        city_line = addr_lines[-1] if addr_lines else ""
        csz = re.search(r"^(.*?),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line, re.I)
        if csz:
            rec["City"]  = clean(csz.group(1))
            rec["State"] = csz.group(2).upper()
            rec["Zip"]   = csz.group(3)
        else:
            # Try without comma: "New York NY 10466"
            csz2 = re.search(r"^(.*?)\s+([A-Z]{2})\s+(\d{5})$", city_line, re.I)
            if csz2:
                rec["City"]  = clean(csz2.group(1))
                rec["State"] = csz2.group(2).upper()
                rec["Zip"]   = csz2.group(3)
            else:
                rec["City/State/Zip"] = city_line
    else:
        rec["Address Line 1"] = rec["Address Line 2"] = ""

    # ── PAY ───────────────────────────────────────────────────────────────────
    rec["Gross"]     = mgrep([r"Gross:\s*([\d,\.]+)", r"Gross\s+([\d,\.]+)"], txt)
    rec["Salary"]    = grep(r"Salary:\s*([\d,\.]+)", txt)
    rec["Bi-Wkly"]   = grep(r"Bi-Wkly\s*[:\s]*([\d,\.]+)", txt)
    rec["Rate Calc"] = grep(r"Rate\s*Calc:\s*(\w+)", txt)
    rec["LWW"]       = grep(r"LWW:\s*(\d+)", txt)
    rec["NWW"]       = grep(r"NWW:\s*(\d+)", txt)
    rec["Std Hours"] = grep(r"Std\s*Hours:\s*([\d\.]+)", txt)
    rec["Pay Group"] = grep(r"Pay\s*Group:\s*(\d+)", txt)

    # ── TAX ───────────────────────────────────────────────────────────────────
    rec["Marital Status"]     = grep(r"Marital\s+Status:\s*([^\n]+)", txt)
    rec["Federal Exemptions"] = mgrep([
        r"Federal[:\s]+(\d+)\s*Exemptions?",
        r"(\d+)\s*Exemptions?\s*\n.*Federal",
    ], txt)
    rec["Federal Extra W/H"]  = grep(r"Extra\s*W[/\\]H\s*\$?([\d,\.]+)", txt)

    # State tax lines — collect all state-looking lines
    state_lines = re.findall(
        r"^\s*(\d{2,4}[A-Z]?\s+[A-Z]{2}[\w\s\-]*?)$",
        txt, re.MULTILINE
    )
    state_lines = [clean(s) for s in state_lines if len(clean(s)) > 3][:4]
    rec["State Tax 1"] = state_lines[0] if len(state_lines) > 0 else ""
    rec["State Tax 2"] = state_lines[1] if len(state_lines) > 1 else ""
    rec["State Tax 3"] = state_lines[2] if len(state_lines) > 2 else ""
    rec["State Tax 4"] = state_lines[3] if len(state_lines) > 3 else ""

    # ── SCHEDULED AMOUNTS ────────────────────────────────────────────────────
    rec["401K"]      = mgrep([r"K\s*401K\s+([\d,\.]+)", r"401K\s+([\d,\.]+)"], txt)
    rec["Pre-Med"]   = mgrep([r"35\s*PREMED\s+([\d,\.]+)", r"PREMED\s+([\d,\.]+)"], txt)
    rec["ADDLCH"]    = mgrep([r"42\s*ADDLCH\s+([\d,\.]+)", r"ADDLCH\s+([\d,\.]+)"], txt)
    rec["AD&D"]      = mgrep([r"57\s*AD&?D\s+([\d,\.]+)", r"AD&?D\s+([\d,\.]+)"], txt)
    rec["HSA"]       = mgrep([r"HSA\s+HCCACT\s+([\d,\.]+)", r"HSA\s+([\d,\.]+)"], txt)
    rec["Goal Limit"]    = grep(r"Limit:\s*([\d,\.]+)", txt)
    rec["Goal To Date"]  = grep(r"To\s*Date:\s*([\d,\.]+)", txt)

    # ── DIRECT DEPOSIT ────────────────────────────────────────────────────────
    rec["Acct #"]   = mgrep([
        r"Acct\s*#\s*[:\-]?\s*([\dXx*\-]+)",
        r"Account\s*#?\s*[:\-]?\s*([\dXx*\-]+)"
    ], txt)
    rec["Tran/ABA"] = grep(r"Tran/ABA:\s*([\d\s]+)", txt)
    rec["DD Code"]  = grep(r"Code\s+([A-Z])\b", txt)
    rec["DD Type"]  = mgrep([r"(Full\s+Deposit)", r"(Partial\s+Deposit)"], txt)

    if DEBUG:
        for k, v in rec.items():
            if v:
                print(f"  {k:22s}: {v}")

    return rec

# ── read PDF and extract all employees ────────────────────────────────────────

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

        # Collect all text first
        all_text = ""
        with tqdm(total=len(page_indices), desc="📄 Reading pages",
                  unit="pg", colour="cyan", ncols=65) as pbar:
            for page_num in page_indices:
                page = pdf.pages[page_num]
                t = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                all_text += t + "\n"
                pbar.update(1)

    if not all_text.strip():
        print("\n  ❌ No text found in PDF. File may be scanned (image-only).")
        return []

    if DEBUG:
        print(f"\n  === FULL EXTRACTED TEXT (first 1000 chars) ===")
        print(all_text[:1000])
        print("  ===")

    # Split into employee blocks
    blocks = split_into_blocks(all_text)
    print(f"  Employee blocks found: {len(blocks)}")

    if not blocks:
        print("\n  ⚠️  Could not detect employee blocks.")
        print("  Run with --debug --pages 1-1 to see raw text.")
        return []

    # Parse each block
    with tqdm(total=len(blocks), desc="👤 Parsing records",
              unit="emp", colour="green", ncols=65) as pbar:
        for i, block in enumerate(blocks):
            try:
                rec = parse_block(block)
                rec["_block"] = i + 1
                if rec.get("Last Name") or rec.get("File #"):
                    employees.append(rec)
                else:
                    tqdm.write(f"  ⚠️  Block {i+1}: no name or file# — skipped")
            except Exception as e:
                tqdm.write(f"  ⚠️  Block {i+1} error: {e}")
            pbar.update(1)

    return employees

# ── Excel writer ──────────────────────────────────────────────────────────────

COLUMNS = [
    # Identity
    "Last Name", "First Name", "Continued",
    # Personnel
    "File #", "Status", "Dept", "Sex", "Cntl", "Race",
    "SSN", "Occup", "Title", "eVoucher",
    "Cost", "Qualified Pension", "Health Coverage",
    # Dates
    "Hire Date", "Term Date", "Birth Date",
    "Date 1", "Date 3", "Date 6", "Date 8", "Date 9",
    # Address
    "Address Line 1", "Address Line 2", "City", "State", "Zip", "City/State/Zip",
    # Pay
    "Gross", "Salary", "Bi-Wkly", "Rate Calc", "LWW", "NWW", "Std Hours", "Pay Group",
    # Tax
    "Marital Status", "Federal Exemptions", "Federal Extra W/H",
    "State Tax 1", "State Tax 2", "State Tax 3", "State Tax 4",
    # Scheduled
    "401K", "Pre-Med", "ADDLCH", "AD&D", "HSA", "Goal Limit", "Goal To Date",
    # Direct Deposit
    "Acct #", "Tran/ABA", "DD Code", "DD Type",
    # Meta
    "_block",
]

SEC_COLOR = {
    "Last Name":"1F4E79","First Name":"1F4E79","Continued":"1F4E79",
    "File #":"375623","Status":"375623","Dept":"375623","Sex":"375623",
    "Cntl":"375623","Race":"375623","SSN":"375623","Occup":"375623",
    "Title":"375623","eVoucher":"375623","Cost":"375623",
    "Qualified Pension":"375623","Health Coverage":"375623",
    "Hire Date":"7B2C2C","Term Date":"7B2C2C","Birth Date":"7B2C2C",
    "Date 1":"7B2C2C","Date 3":"7B2C2C","Date 6":"7B2C2C",
    "Date 8":"7B2C2C","Date 9":"7B2C2C",
    "Address Line 1":"1F3864","Address Line 2":"1F3864","City":"1F3864",
    "State":"1F3864","Zip":"1F3864","City/State/Zip":"1F3864",
    "Gross":"7B4A00","Salary":"7B4A00","Bi-Wkly":"7B4A00",
    "Rate Calc":"7B4A00","LWW":"7B4A00","NWW":"7B4A00",
    "Std Hours":"7B4A00","Pay Group":"7B4A00",
    "Marital Status":"4A235A","Federal Exemptions":"4A235A",
    "Federal Extra W/H":"4A235A","State Tax 1":"4A235A",
    "State Tax 2":"4A235A","State Tax 3":"4A235A","State Tax 4":"4A235A",
    "401K":"0D4C6E","Pre-Med":"0D4C6E","ADDLCH":"0D4C6E","AD&D":"0D4C6E",
    "HSA":"0D4C6E","Goal Limit":"0D4C6E","Goal To Date":"0D4C6E",
    "Acct #":"5C3317","Tran/ABA":"5C3317","DD Code":"5C3317","DD Type":"5C3317",
    "_block":"555555",
}

SEC_LABELS = {
    "Last Name":       "👤 Identity",
    "File #":          "🗂 Personnel",
    "Hire Date":       "📅 Dates",
    "Address Line 1":  "📍 Address",
    "Gross":           "💰 Pay",
    "Marital Status":  "🧾 Tax",
    "401K":            "📊 Scheduled",
    "Acct #":          "🏦 Direct Deposit",
    "_block":          "ℹ Meta",
}

def write_excel(employees, out_path):
    """Plain simple Excel — no colors, no merged cells, just headers + data."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ADP Master Control"

    # Header row — plain bold
    for ci, col in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font = Font(bold=True, name="Arial", size=10)

    # Data rows
    with tqdm(total=len(employees), desc="💾 Writing Excel",
              unit="row", colour="yellow", ncols=65) as pbar:
        for ri, emp in enumerate(employees, 2):
            for ci, col in enumerate(COLUMNS, 1):
                c = ws.cell(row=ri, column=ci, value=emp.get(col, ""))
                c.font = Font(name="Arial", size=10)
            pbar.update(1)

    # Auto-width based on content
    for ci, col in enumerate(COLUMNS, 1):
        max_len = len(col)
        for ri in range(2, min(len(employees)+2, 52)):   # sample up to 50 rows
            val = ws.cell(row=ri, column=ci).value or ""
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"
    wb.save(out_path)
# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(args) < 1:
        print(__doc__)
        sys.exit(0)

    pdf_path = args[0]
    out_path = args[1] if len(args) > 1 else \
               str(Path(pdf_path).stem + "_employees.xlsx")

    if not Path(pdf_path).exists():
        print(f"\n❌ File not found: {pdf_path}")
        sys.exit(1)

    flags = []
    if PAGE_FROM or PAGE_TO: flags.append(f"pages {PAGE_FROM}-{PAGE_TO}")
    if DEBUG:                 flags.append("DEBUG")

    print(f"\n{'='*55}")
    print(f"  ADP Master Control Extractor")
    print(f"{'='*55}")
    print(f"  Input : {pdf_path}")
    print(f"  Output: {out_path}")
    if flags: print(f"  Flags : {' | '.join(flags)}")
    print(f"{'='*55}")

    employees = extract_all(pdf_path, page_from=PAGE_FROM, page_to=PAGE_TO)
    print(f"\n  👥 Records found: {len(employees)}")

    if not employees:
        print("\n  ❌ No records extracted. Try:")
        print(f"     python extract_adp.py \"{pdf_path}\" --pages 1-1 --debug\n")
        sys.exit(1)

    write_excel(employees, out_path)

    print(f"\n{'='*55}")
    print(f"  ✅ Done! → {out_path}")
    print(f"{'='*55}\n")
