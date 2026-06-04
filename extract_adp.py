"""
ADP Master Control Report - Employee Data Extractor
=====================================================
Handles the 4-column ADP Master Control layout:
  Col 1: PERSONNEL  |  Col 2: PAY  |  Col 3: TAX STATUS  |  Col 4: SCHEDULED AMOUNTS

Each employee record is one "row" of 4 columns on the page.
Multiple employees per page, multiple pages per file.

Usage:
  python extract_adp.py yourfile.pdf
  python extract_adp.py yourfile.pdf employees.xlsx
  python extract_adp.py yourfile.pdf --debug
"""

import re, sys, io
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
from tqdm import tqdm

DEBUG = "--debug" in sys.argv
args  = [a for a in sys.argv[1:] if not a.startswith("--")]

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

# ── extract text from one bounding box on a page ─────────────────────────────

def crop_text(page, x0, y0, x1, y1):
    """Extract text from a rectangular region of a page."""
    try:
        cropped = page.crop((x0, y0, x1, y1))
        return cropped.extract_text() or ""
    except Exception:
        return ""

# ── find employee record boundaries on a page ─────────────────────────────────

def find_record_boundaries(page):
    """
    Each employee record starts with a PERSONNEL box.
    We detect vertical positions of each PERSONNEL header on the page
    to split the page into individual employee record strips.
    """
    page_width  = page.width
    page_height = page.height

    # Extract all words with their bounding boxes
    words = page.extract_words()

    # Find y-positions of "PERSONNEL" labels (top of each employee record)
    personnel_ys = []
    for w in words:
        if w["text"].upper() == "PERSONNEL":
            personnel_ys.append(w["top"])

    if not personnel_ys:
        # Fallback: try extracting full page text
        if DEBUG:
            print(f"  [page] No PERSONNEL found. Full text sample:")
            print(page.extract_text()[:500] if page.extract_text() else "  (empty)")
        return []

    # Sort and deduplicate (within 5px = same line)
    personnel_ys = sorted(set(round(y/5)*5 for y in personnel_ys))

    if DEBUG:
        print(f"  [page] Found {len(personnel_ys)} PERSONNEL headers at y={personnel_ys}")

    # Build vertical slices: each record goes from its PERSONNEL y to next one
    boundaries = []
    for i, y_top in enumerate(personnel_ys):
        # Start slightly above the PERSONNEL label
        y0 = max(0, y_top - 5)
        # End just before next record, or at page bottom
        y1 = (personnel_ys[i+1] - 5) if i+1 < len(personnel_ys) else page_height
        boundaries.append((y0, y1))

    return boundaries

# ── parse one employee record strip ──────────────────────────────────────────

def parse_record(page, y0, y1):
    """
    Given vertical bounds of one employee record on the page,
    extract each of the 4 column sections separately.
    """
    pw = page.width

    # ADP Master Control: 4 columns roughly at these x% positions:
    # PERSONNEL: 0-25% | PAY: 25-50% | TAX STATUS: 50-73% | SCHEDULED: 73-100%
    col_bounds = [
        (0,        pw*0.26),   # PERSONNEL
        (pw*0.26,  pw*0.50),   # PAY
        (pw*0.50,  pw*0.73),   # TAX STATUS
        (pw*0.73,  pw*1.00),   # SCHEDULED AMOUNTS
    ]

    col_texts = []
    for x0, x1 in col_bounds:
        txt = crop_text(page, x0, y0, x1, y1)
        col_texts.append(txt)

    personnel_txt, pay_txt, tax_txt, sched_txt = col_texts

    if DEBUG:
        print(f"\n{'─'*55} RECORD y={y0:.0f}-{y1:.0f}")
        print(f"[PERSONNEL]\n{personnel_txt}")
        print(f"[PAY]\n{pay_txt}")
        print(f"[TAX]\n{tax_txt}")
        print(f"[SCHEDULED]\n{sched_txt}")

    rec = {}

    # ═══════════════════════════════════════════════════
    # PERSONNEL COLUMN
    # ═══════════════════════════════════════════════════
    p_lines = [l.strip() for l in personnel_txt.splitlines() if l.strip()]

    # Name: first non-header line, all caps, format "LAST, FIRST" or "LAST FIRST"
    SKIP = {"PERSONNEL","PAY","TAX","STATUS","SCHEDULED","AMOUNTS",
            "MAILING","HOME","ADDRESS","DIRECT","DEPOSITS","DATES","GOAL"}
    name_raw = ""
    for ln in p_lines[:8]:
        words_up = set(ln.upper().split())
        if words_up & SKIP and not re.search(r"\d", ln):
            continue
        if re.match(r"^[A-Z][A-Z\s,\.\'\-]{2,}$", ln.upper()):
            name_raw = ln.strip()
            break

    if name_raw:
        if "," in name_raw:
            parts = name_raw.split(",", 1)
            rec["Last Name"]  = clean(parts[0])
            rec["First Name"] = clean(parts[1])
        else:
            toks = name_raw.split()
            rec["Last Name"]  = " ".join(toks[:-1]) if len(toks)>1 else name_raw
            rec["First Name"] = toks[-1] if len(toks)>1 else ""
    else:
        rec["Last Name"] = rec["First Name"] = ""

    # Address — lines between "Mailing & Home Address" and "File:"
    addr_block = grep(
        r"Mailing\s*[&]\s*Home\s*Address[-–\s]*(.*?)(?=File\s*:|Dept\s*:|eVoucher|Hire\s*:|$)",
        personnel_txt, default="", flags=re.DOTALL|re.IGNORECASE)
    addr_lines = [l.strip() for l in addr_block.splitlines() if l.strip()]
    rec["Address Line 1"] = addr_lines[0] if len(addr_lines) > 0 else ""
    rec["Address Line 2"] = addr_lines[1] if len(addr_lines) > 1 else ""
    # City/State/Zip from last address line
    city_line = next((l for l in reversed(addr_lines)
                      if re.search(r"[A-Z]{2}\s+\d{5}", l)), "")
    csz = re.search(r"^(.*?),?\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line)
    if csz:
        rec["City"]  = clean(csz.group(1))
        rec["State"] = csz.group(2)
        rec["Zip"]   = csz.group(3)
    else:
        rec["City/State/Zip"] = city_line

    rec["File #"]          = mgrep([r"File\s*[:#]?\s*([\w\-]+)", r"File\s+([\d\-]+)"], personnel_txt)
    rec["Dept"]            = grep(r"Dept\s*[:#]?\s*(\S+)", personnel_txt)
    rec["Cntl"]            = grep(r"Cntl\s*[:#]?\s*(\S+)", personnel_txt)
    rec["SSN"]             = grep(r"SSN\s*[:#]?\s*([^\n]+)", personnel_txt)
    rec["Title"]           = grep(r"Title\s*[:#]?\s*(\S+)", personnel_txt)
    rec["Cost"]            = grep(r"Cost\s*[:#]?\s*(.+)", personnel_txt)
    rec["eVoucher Status"] = grep(r"Status\s*[:#]?\s*(\w+)", personnel_txt)
    rec["Sex"]             = grep(r"Sex\s*[:#]?\s*(\w)", personnel_txt)
    rec["Race"]            = grep(r"Race\s*[:#]?\s*(\w+)", personnel_txt)
    rec["Occup"]           = grep(r"Occup\s*[:#]?\s*(\w+)", personnel_txt)
    rec["eW2"]             = grep(r"eW2\s*[:#]?\s*(\w+)", personnel_txt)

    # Dates
    rec["Hire Date"]  = mgrep([r"Hire\s*[:#]?\s*([\d/]+)", r"Hire\s+Date\s*[:#]?\s*([\d/]+)"], personnel_txt)
    rec["Birth Date"] = mgrep([r"Birth\s*[:#]?\s*([\d/]+)", r"DOB\s*[:#]?\s*([\d/]+)"], personnel_txt)
    rec["Date 6"]     = grep(r"Date\s*6\s*[:#]?\s*([\d/]+)", personnel_txt)
    rec["Date 9"]     = grep(r"Date\s*9\s*[:#]?\s*([\d/]+)", personnel_txt)
    rec["Qualified Pension"] = "Yes" if re.search(r"Qualified\s*Pension", personnel_txt, re.I) else ""

    # ═══════════════════════════════════════════════════
    # PAY COLUMN
    # ═══════════════════════════════════════════════════
    rec["Gross Pay"]     = mgrep([r"Gross\s*[:#]?\s*([\d,\.]+)", r"Gross\s+Pay\s*[:#]?\s*([\d,\.]+)"], pay_txt)
    rec["Salary"]        = grep(r"Salary\s*[:#]?\s*([\d,\.]+)", pay_txt)
    rec["Pay Frequency"] = mgrep([r"(Bi-Wkly|Bi-Weekly|Weekly|Semi-Monthly|Monthly)"], pay_txt)
    rec["Rate 2"]        = grep(r"Rate\s*2\s*[:#]?\s*([\d,\.]+)", pay_txt)
    rec["Rate Calc"]     = grep(r"Rate\s*Calc\s*[:#]?\s*(\w+)", pay_txt)
    rec["LWW"]           = grep(r"LWW\s*[:#]?\s*(\d+)", pay_txt)
    rec["NWW"]           = grep(r"NWW\s*[:#]?\s*(\d+)", pay_txt)
    rec["Std Hours"]     = grep(r"Std\s*Hours\s*[:#]?\s*([\d\.]+)", pay_txt)
    rec["Pay Group"]     = grep(r"Pay\s*Group\s*[:#]?\s*(\d+)", pay_txt)
    rec["Paid"]          = grep(r"Paid\s+(.+?)(?:\n|$)", pay_txt)
    rec["Prior Qtr"]     = grep(r"Prior\s+Qtr\s+(.+?)(?:\n|$)", pay_txt)

    # ═══════════════════════════════════════════════════
    # TAX STATUS COLUMN
    # ═══════════════════════════════════════════════════
    rec["Federal"]        = grep(r"Federal\s*[:#]?\s*(.+?)(?:\n|$)", tax_txt)
    rec["W4 Year"]        = grep(r"(\d{4})\s*Form\s*W-?4", tax_txt)
    rec["Federal Filing"] = mgrep([
        r"(D-Single[^,\n]{0,40})",
        r"(Single[^,\n]{0,40})",
        r"(Married[^,\n]{0,40})",
        r"(HH[^,\n]{0,40})"
    ], tax_txt)
    rec["Dependents"]     = grep(r"Dependents?\s*\$?\s*([\d\.]+)", tax_txt)
    rec["Extra W/H"]      = grep(r"Extra\s*W[/\\]H\s*\$?\s*([\d,\.]+)", tax_txt)
    rec["Filing Status"]  = grep(r"Filing\s*Status\s*(.+?)(?:\n|$)", tax_txt)
    rec["State Tax"]      = mgrep([
        r"(\d{2}\s+[A-Z]{2}\s+\(Lived\s+in\))",
        r"(\d{2}\s+[A-Z]{2}\s+Lived)",
        r"(0[0-9]\s+[A-Z]{2}[^\n]*)"
    ], tax_txt)
    rec["Extra State Tax"]= grep(r"\$([\d,\.]+)\s*Extra\s*State\s*Tax", tax_txt)
    rec["SUI/DI"]         = grep(r"(\d{2}\s+[A-Z]{2}\s+SUI/DI)", tax_txt)
    rec["GTL Coverage"]   = grep(r"GTL\s*Cov(?:erage)?\s*([\d,\.]+)", tax_txt)
    rec["FLI"]            = "Exempt" if re.search(r"Exempt\s*FLI", tax_txt, re.I) else ""

    # ═══════════════════════════════════════════════════
    # SCHEDULED AMOUNTS COLUMN
    # ═══════════════════════════════════════════════════
    rec["401K"]         = mgrep([r"401K\s+([\d,\.]+)", r"K\s*401K\s+([\d,\.]+)"], sched_txt)
    rec["V5 VIS"]       = grep(r"V5\s*VIS\s+([\d,\.]+)", sched_txt)
    rec["Pre-Med"]      = mgrep([r"35\s*PREMED\s+([\d,\.]+)", r"PREMED\s+([\d,\.]+)"], sched_txt)
    rec["ADDLF"]        = grep(r"37\s*ADDLF\s+([\d,\.]+)", sched_txt)
    rec["ADDLCH"]       = mgrep([r"42\s*ADDLCH\s+([\d,\.]+)", r"ADDLCH\s+([\d,\.]+)"], sched_txt)
    rec["ADDSP"]        = grep(r"65\s*ADDSP\s+([\d,\.]+)", sched_txt)
    rec["AD&D"]         = mgrep([r"57\s*AD&?D\s+([\d,\.]+)", r"AD&?D\s+([\d,\.]+)"], sched_txt)
    rec["PREDN"]        = grep(r"65\s*PREDN\s+([\d,\.]+)", sched_txt)
    rec["HSA"]          = mgrep([r"HSA\s+HCCACT\s+([\d,\.]+)", r"HSA\s+([\d,\.]+)"], sched_txt)
    rec["Goal Limit"]   = grep(r"Limit\s*[:#]?\s*([\d,\.]+)", sched_txt)
    rec["Goal To Date"] = grep(r"To\s*Date\s*[:#]?\s*([\d,\.]+)", sched_txt)

    # Direct Deposits
    rec["Acct #"]   = mgrep([
        r"Acct\s*#\s*[:\-]?\s*([\dXx*\-]+)",
        r"Account\s*#?\s*[:\-]?\s*([\dXx*\-]+)"
    ], sched_txt)
    rec["Tran/ABA"] = mgrep([r"Tran/ABA\s*[:#]?\s*([\d\s]+)", r"ABA\s*[:#]?\s*([\d\s]+)"], sched_txt)
    rec["DD Code"]  = grep(r"Code\s+([A-Z])\b", sched_txt)
    rec["DD Type"]  = mgrep([r"(Full\s+Deposit)", r"(Partial\s+Deposit)"], sched_txt)

    return rec

# ── process full PDF ──────────────────────────────────────────────────────────

def extract_all(pdf_path):
    employees = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"  Pages: {total_pages}")

        with tqdm(total=total_pages, desc="📄 Reading pages",
                  unit="pg", colour="cyan", ncols=65) as pbar:
            for page_num, page in enumerate(pdf.pages):
                boundaries = find_record_boundaries(page)

                if DEBUG:
                    print(f"\n══ PAGE {page_num+1} — {len(boundaries)} records ══")

                for y0, y1 in boundaries:
                    try:
                        rec = parse_record(page, y0, y1)
                        rec["_page"] = page_num + 1
                        # Accept if we got at least a name or file number
                        if rec.get("Last Name") or rec.get("File #") or rec.get("Gross Pay"):
                            employees.append(rec)
                        elif DEBUG:
                            print(f"  [skip] Empty record at y={y0:.0f}-{y1:.0f}")
                    except Exception as e:
                        tqdm.write(f"  ⚠️  Page {page_num+1} y={y0:.0f}: {e}")
                pbar.update(1)

    return employees

# ── write Excel ───────────────────────────────────────────────────────────────

COLUMNS = [
    # Identity
    "Last Name","First Name",
    # Address
    "Address Line 1","Address Line 2","City","State","Zip","City/State/Zip",
    # Personnel
    "File #","Dept","Cntl","SSN","Title","Cost","eVoucher Status",
    "Sex","Race","Occup","eW2","Qualified Pension",
    # Dates
    "Hire Date","Birth Date","Date 6","Date 9",
    # Pay
    "Gross Pay","Salary","Pay Frequency","Rate 2","Rate Calc",
    "LWW","NWW","Std Hours","Pay Group","Paid","Prior Qtr",
    # Tax
    "Federal","W4 Year","Federal Filing","Dependents","Extra W/H",
    "Filing Status","State Tax","Extra State Tax","SUI/DI","GTL Coverage","FLI",
    # Scheduled Amounts
    "401K","V5 VIS","Pre-Med","ADDLF","ADDLCH","ADDSP","AD&D","PREDN","HSA",
    "Goal Limit","Goal To Date",
    # Direct Deposit
    "Acct #","Tran/ABA","DD Code","DD Type",
    # Meta
    "_page",
]

# Section header colors for column groups
SECTION_COLORS = {
    "Last Name":"1F4E79","First Name":"1F4E79",
    "Address Line 1":"1F3864","Address Line 2":"1F3864",
    "City":"1F3864","State":"1F3864","Zip":"1F3864","City/State/Zip":"1F3864",
    "File #":"375623","Dept":"375623","Cntl":"375623","SSN":"375623",
    "Title":"375623","Cost":"375623","eVoucher Status":"375623",
    "Sex":"375623","Race":"375623","Occup":"375623","eW2":"375623",
    "Qualified Pension":"375623",
    "Hire Date":"7B2C2C","Birth Date":"7B2C2C","Date 6":"7B2C2C","Date 9":"7B2C2C",
    "Gross Pay":"7B4A00","Salary":"7B4A00","Pay Frequency":"7B4A00",
    "Rate 2":"7B4A00","Rate Calc":"7B4A00","LWW":"7B4A00","NWW":"7B4A00",
    "Std Hours":"7B4A00","Pay Group":"7B4A00","Paid":"7B4A00","Prior Qtr":"7B4A00",
    "Federal":"4A235A","W4 Year":"4A235A","Federal Filing":"4A235A",
    "Dependents":"4A235A","Extra W/H":"4A235A","Filing Status":"4A235A",
    "State Tax":"4A235A","Extra State Tax":"4A235A","SUI/DI":"4A235A",
    "GTL Coverage":"4A235A","FLI":"4A235A",
    "401K":"0D4C6E","V5 VIS":"0D4C6E","Pre-Med":"0D4C6E","ADDLF":"0D4C6E",
    "ADDLCH":"0D4C6E","ADDSP":"0D4C6E","AD&D":"0D4C6E","PREDN":"0D4C6E",
    "HSA":"0D4C6E","Goal Limit":"0D4C6E","Goal To Date":"0D4C6E",
    "Acct #":"5C3317","Tran/ABA":"5C3317","DD Code":"5C3317","DD Type":"5C3317",
    "_page":"444444",
}

SECTION_LABELS = {
    "Last Name":        "👤 Identity",
    "Address Line 1":   "📍 Address",
    "File #":           "🗂 Personnel",
    "Hire Date":        "📅 Dates",
    "Gross Pay":        "💰 Pay",
    "Federal":          "🧾 Tax Status",
    "401K":             "📊 Scheduled Amounts",
    "Acct #":           "🏦 Direct Deposit",
    "_page":            "ℹ Meta",
}

def write_excel(employees, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ADP Master Control"

    BD = Side(style="thin", color="BFBFBF")
    BDR = Border(left=BD, right=BD, top=BD, bottom=BD)
    N_FONT = Font(name="Arial", size=10)
    A_FILL = PatternFill("solid", fgColor="EEF2F7")

    # Row 1: Section labels (merged)
    # Row 2: Column headers
    # Row 3+: Data

    # Build section spans
    section_spans = {}
    for ci, col in enumerate(COLUMNS, 1):
        label = SECTION_LABELS.get(col)
        if label:
            section_spans[label] = {"start": ci, "end": ci, "color": SECTION_COLORS.get(col,"333333")}
        else:
            # Extend last section
            for lbl in reversed(list(section_spans.keys())):
                section_spans[lbl]["end"] = ci
                break

    # Write section header row
    for label, span in section_spans.items():
        ws.merge_cells(start_row=1, start_column=span["start"],
                       end_row=1, end_column=span["end"])
        c = ws.cell(row=1, column=span["start"], value=label)
        c.fill = PatternFill("solid", fgColor=span["color"])
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BDR
    ws.row_dimensions[1].height = 22

    # Column header row
    for ci, col in enumerate(COLUMNS, 1):
        color = SECTION_COLORS.get(col, "1F4E79")
        c = ws.cell(row=2, column=ci, value=col)
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BDR
    ws.row_dimensions[2].height = 28

    # Data rows
    with tqdm(total=len(employees), desc="💾 Writing Excel",
              unit="row", colour="yellow", ncols=65) as pbar:
        for ri, emp in enumerate(employees, 3):
            fill = A_FILL if ri % 2 == 0 else None
            for ci, col in enumerate(COLUMNS, 1):
                val = emp.get(col, "")
                c = ws.cell(row=ri, column=ci, value=val)
                c.font = N_FONT
                c.border = BDR
                c.alignment = Alignment(vertical="center")
                if fill: c.fill = fill
            pbar.update(1)

    # Column widths
    W = {"Last Name":16,"First Name":16,"Address Line 1":26,"Address Line 2":18,
         "City":15,"State":6,"Zip":9,"City/State/Zip":20,"File #":11,"Dept":9,
         "Cntl":7,"SSN":14,"Title":13,"Cost":22,"eVoucher Status":13,"Sex":5,
         "Race":7,"Occup":7,"eW2":6,"Qualified Pension":14,"Hire Date":11,
         "Birth Date":11,"Date 6":11,"Date 9":11,"Gross Pay":11,"Salary":11,
         "Pay Frequency":13,"Rate 2":9,"Rate Calc":9,"LWW":7,"NWW":7,
         "Std Hours":9,"Pay Group":9,"Paid":18,"Prior Qtr":12,"Federal":16,
         "W4 Year":8,"Federal Filing":26,"Dependents":11,"Extra W/H":9,
         "Filing Status":20,"State Tax":20,"Extra State Tax":14,"SUI/DI":13,
         "GTL Coverage":13,"FLI":8,"401K":9,"V5 VIS":9,"Pre-Med":9,
         "ADDLF":8,"ADDLCH":9,"ADDSP":8,"AD&D":7,"PREDN":8,"HSA":9,
         "Goal Limit":11,"Goal To Date":13,"Acct #":16,"Tran/ABA":15,
         "DD Code":7,"DD Type":13,"_page":6}
    for ci, col in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = W.get(col, 12)

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(COLUMNS))}2"
    wb.save(out_path)

# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(args) < 1:
        print(__doc__)
        sys.exit(1)

    pdf_path = args[0]
    out_path = args[1] if len(args) > 1 else str(Path(pdf_path).stem + "_employees.xlsx")

    if not Path(pdf_path).exists():
        print(f"\n❌ File not found: {pdf_path}")
        sys.exit(1)

    print(f"\n{'='*55}")
    print(f"  ADP Master Control Extractor")
    print(f"{'='*55}")
    print(f"  Input : {pdf_path}")
    print(f"  Output: {out_path}")
    if DEBUG: print(f"  Mode  : DEBUG ON")
    print(f"{'='*55}")

    employees = extract_all(pdf_path)

    print(f"\n  👥 Records found: {len(employees)}")

    if not employees:
        print("\n  ❌ No records extracted.")
        print("  Try: python extract_adp.py yourfile.pdf --debug")
        print("  This will show the raw text per page so we can diagnose.\n")
        sys.exit(1)

    write_excel(employees, out_path)

    print(f"\n{'='*55}")
    print(f"  ✅ Done! Saved → {out_path}")
    print(f"{'='*55}\n")
