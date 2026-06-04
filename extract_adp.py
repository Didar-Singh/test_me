"""
ADP Master Control - Extractor
===============================
Single file OR batch mode - auto-detects!

Usage:
  python extract_adp.py file.pdf                    (single file)
  python extract_adp.py file.pdf output.xlsx        (single + custom output)
  python extract_adp.py C:\path\to\folder           (batch mode - all PDFs)
  python extract_adp.py file.pdf --pages 1-5        (single + page range)
  python extract_adp.py file.pdf --debug            (debug mode)
"""
import re, sys, os
from pathlib import Path
import pdfplumber
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from tqdm import tqdm

DEBUG = "--debug" in sys.argv
PAGE_FROM = PAGE_TO = None

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

# ── HELPERS ───────────────────────────────────────────────────────────────────

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

# ── NAME FINDER ────────────────────────────────────────────────────────────────

def find_all_names(text):
    """Find ALL names: LAST,FIRST format"""
    positions = []
    for m in re.finditer(r"([A-Z][A-Z\'\-\.]+),([A-Z][A-Z\'\-\.\s]+?)(?=\s|$)", text):
        positions.append((m.start(), m.group(0)))
    return sorted(set(positions))

def split_by_names(text):
    """Split text into blocks by employee names"""
    names = find_all_names(text)
    if not names:
        return []
    
    blocks = []
    for i, (pos, name) in enumerate(names):
        start = pos
        end = names[i+1][0] if i+1 < len(names) else len(text)
        blocks.append(text[start:end].strip())
    
    if DEBUG:
        print(f"\n[FOUND {len(blocks)} EMPLOYEE BLOCKS]")
    
    return blocks

# ── PARSE EMPLOYEE ─────────────────────────────────────────────────────────────

def parse_block(block, num=1):
    """Parse messy single-line format"""
    rec = {"_block": num}
    txt = block
    
    # NAME - capture full first name including middle initials
    name_m = re.search(r"([A-Z][A-Z\'\-\.]+),([A-Z][A-Z\'\-\.\s]+?)(?=\s+File:|\s+Marital|\s+Gross:|\s+Mailing|$)", txt, re.I)
    if name_m:
        rec["Last Name"] = clean(name_m.group(1))
        rec["First Name"] = clean(name_m.group(2))
    
    rec["Continued"] = "Yes" if "(continued)" in txt else ""
    rec["File #"] = grep(r"File:\s*(\d+)", txt)
    rec["Status"] = mgrep([r"Status:\s*(\w+)", r"Status\s+(\w+)"], txt)
    rec["Dept"] = grep(r"Dept:\s*(\S+)", txt)
    rec["Sex"] = grep(r"Sex:\s*(\w)", txt)
    rec["Cntl"] = grep(r"Cntl:\s*(\S+)", txt)
    rec["Race"] = grep(r"Race:\s*(\S+)", txt)
    rec["Occup"] = grep(r"Occup:\s*(\S+)", txt)
    rec["SSN"] = grep(r"SSN:\s*([^\n]+?)(?=\s{2,}|GTL|Title|$)", txt)
    rec["Title"] = grep(r"Title:\s*(\S+)", txt)
    
    # DATES
    rec["Hire Date"] = grep(r"Hire:\s*([\d/]+)", txt)
    rec["Term Date"] = grep(r"Term:\s*([\d/]+)", txt)
    rec["Birth Date"] = grep(r"Birth:\s*([\d/]+)", txt)
    rec["Date 6"] = grep(r"Date\s+6:\s*([\d/]+)", txt)
    rec["Date 8"] = grep(r"Date\s+8:\s*([\d/]+)", txt)
    rec["Date 9"] = grep(r"Date\s+9:\s*([\d/]+)", txt)
    
    # ADDRESS - stop at pipes, Monthly, Exemptions, etc
    addr_lines = []
    for m in re.finditer(r"^(\d+\s+[A-Z][A-Z\s\.\-]+?)(?=\s*\||Monthly|Exemptions|Federal|Form|Rate|$)", txt, re.MULTILINE | re.I):
        street = clean(m.group(1))
        if street and not re.search(r"(Monthly|Exemptions|Federal|Form)$", street, re.I):
            addr_lines.append(street)
    
    for m in re.finditer(r"(?:[A-Z][\s])*(?:Q|Y|SS)?\s*([A-Z]{2,}?)\s+([A-Z]{2})\s+(\d{5})", txt):
        city_raw = m.group(1).strip()
        if len(city_raw) > 2 and city_raw not in ["SS", "YY", "QQ"]:
            city_clean = re.sub(r"\s+[A-Z]$", "", city_raw).strip()
            if len(city_clean) > 2:
                rec["City"] = clean(city_clean)
                rec["State"] = m.group(2).upper()
                rec["Zip"] = m.group(3)
                break
    
    if addr_lines:
        rec["Address Line 1"] = addr_lines[0] if len(addr_lines) > 0 else ""
        rec["Address Line 2"] = addr_lines[1] if len(addr_lines) > 1 else ""
    
    # PAY
    gross_matches = re.findall(r"Gross:\s*([\d\s,\.]+?)(?=\s+(?:Federal|Salary|State|Form|Rate|Std|Dept|File|Marital|Q\s|$))", txt, re.I)
    if gross_matches:
        rec["Gross"] = clean(gross_matches[0]).replace(" ", "")
    else:
        rec["Gross"] = ""
    
    salary_m = re.search(r"Salary:\s*([\d\s,\.]+?)(?=\s+(?:Federal|Form|Rate|2020|Std|Dept|File|Monthly|$))", txt, re.I)
    rec["Salary"] = clean(salary_m.group(1)).replace(" ", "") if salary_m else ""
    
    rec["Rate Calc"] = grep(r"Rate\s+Calc:\s*(\S+)", txt)
    rec["Std Hours"] = grep(r"Std\s+Hours:\s*([\d\.]+)", txt)
    rec["Pay Group"] = grep(r"Pay\s+Group:\s*(\d+)", txt)
    
    # TAX
    rec["Marital Status"] = grep(r"(J?-?Married|Single)", txt)
    rec["Federal Exemptions"] = mgrep([r"Exemptions[^0-9]*(\d+)", r"(\d+)\s+Exemptions"], txt)
    
    state_m = re.search(r"([\d]{2}\s+[A-Z]{2}(?:\s+[A-Z\s]+)?)", txt)
    if state_m:
        rec["State Tax"] = clean(state_m.group(1))
    
    rec["GTL Cov"] = grep(r"GTL\s+Cov[^0-9]*?([\d\s]+)", txt)
    rec["401K"] = grep(r"401K\s+([\d,\.]+)", txt)
    
    # DIRECT DEPOSIT
    rec["Acct #"] = mgrep([r"Acct\s*#\s*[:\-]?\s*([\dXx*\-]+)", r"Acct:\s*([\dXx*\-]+)"], txt)
    rec["Tran/ABA"] = grep(r"Tran[/\\]ABA:\s*([\d\s]+)", txt)
    rec["DD Code"] = grep(r"Code\s+([A-Z])\b", txt)
    rec["DD Type"] = grep(r"(Full|Partial)\s+Deposit", txt)
    
    return rec

# ── EXTRACT FROM PDF ───────────────────────────────────────────────────────────

def extract_all(pdf_path, page_from=None, page_to=None):
    employees = []
    
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        p_from = max(1, page_from) if page_from else 1
        p_to = min(total, page_to) if page_to else total
        
        if p_from > total or p_from > p_to:
            print(f"\n  ❌ Invalid pages {p_from}-{p_to} (PDF has {total})")
            return []
        
        all_text = ""
        with tqdm(total=p_to-p_from+1, desc="📄 Reading", unit="pg", colour="cyan", ncols=65) as pbar:
            for page_num in range(p_from-1, p_to):
                page = pdf.pages[page_num]
                t = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                all_text += t + "\n"
                pbar.update(1)
        
        if not all_text.strip():
            return []
        
        blocks = split_by_names(all_text)
        
        if not blocks:
            return []
        
        with tqdm(total=len(blocks), desc="👤 Parsing", unit="emp", colour="green", ncols=65) as pbar:
            for i, block in enumerate(blocks):
                try:
                    rec = parse_block(block, i+1)
                    if rec.get("Last Name") or rec.get("File #"):
                        employees.append(rec)
                except Exception as e:
                    tqdm.write(f"  ⚠️  Block {i+1}: {e}")
                pbar.update(1)
    
    return employees

# ── EXCEL WRITER ───────────────────────────────────────────────────────────────

COLUMNS = [
    "Last Name", "First Name", "Continued",
    "File #", "Status", "Dept", "Sex", "Cntl", "Race", "Occup", "SSN", "Title",
    "Hire Date", "Term Date", "Birth Date", "Date 6", "Date 8", "Date 9",
    "Address Line 1", "Address Line 2", "City", "State", "Zip",
    "Gross", "Salary", "Rate Calc", "Std Hours", "Pay Group",
    "Marital Status", "Federal Exemptions", "State Tax",
    "GTL Cov", "401K", "Acct #", "Tran/ABA", "DD Code", "DD Type",
    "_block",
]

def write_excel(employees, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ADP Master Control"
    
    for ci, col in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font = Font(bold=True, name="Arial", size=10)
    
    with tqdm(total=len(employees), desc="💾 Excel", unit="row", colour="yellow", ncols=65) as pbar:
        for ri, emp in enumerate(employees, 2):
            for ci, col in enumerate(COLUMNS, 1):
                ws.cell(row=ri, column=ci, value=emp.get(col, ""))
            pbar.update(1)
    
    for ci, col in enumerate(COLUMNS, 1):
        max_len = len(col)
        for ri in range(2, min(len(employees)+2, 52)):
            val = ws.cell(row=ri, column=ci).value or ""
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 2, 50)
    
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"
    wb.save(out_path)

# ── MAIN ───────────────────────────────────────────────────────────────────────

def process_single_file(pdf_path, out_path=None):
    """Process ONE PDF file"""
    if not out_path:
        out_path = str(Path(pdf_path).stem + "_employees.xlsx")
    
    print(f"\n{'='*55}")
    print(f"  ADP Master Control Extractor")
    print(f"{'='*55}")
    print(f"  Input : {pdf_path}")
    print(f"  Output: {out_path}")
    print(f"{'='*55}")
    
    employees = extract_all(pdf_path, page_from=PAGE_FROM, page_to=PAGE_TO)
    print(f"\n  👥 Records: {len(employees)}")
    
    if not employees:
        print("\n❌ No records extracted.")
        return False
    
    write_excel(employees, out_path)
    print(f"\n{'='*55}")
    print(f"  ✅ Done! → {out_path}")
    print(f"{'='*55}\n")
    return True

def process_folder(folder_path):
    """Process ALL PDFs in FOLDER"""
    folder = Path(folder_path)
    pdf_files = sorted(folder.glob("*.pdf"))
    
    if not pdf_files:
        print(f"❌ No PDF files in: {folder}")
        return
    
    print(f"\n{'='*55}")
    print(f"  ADP Batch Extractor")
    print(f"{'='*55}")
    print(f"  Folder: {folder.absolute()}")
    print(f"  Files : {len(pdf_files)} PDF(s)")
    print(f"{'='*55}\n")
    
    success = 0
    failed = 0
    
    for pdf_file in tqdm(pdf_files, desc="Processing", unit="file", colour="cyan", ncols=70):
        out_file = pdf_file.parent / f"{pdf_file.stem}_employees.xlsx"
        try:
            employees = extract_all(str(pdf_file))
            if not employees:
                tqdm.write(f"  ⚠️  {pdf_file.name}: No records")
                failed += 1
                continue
            write_excel(employees, str(out_file))
            tqdm.write(f"  ✅ {pdf_file.name} ({len(employees)} records)")
            success += 1
        except Exception as e:
            tqdm.write(f"  ❌ {pdf_file.name}: {e}")
            failed += 1
    
    print(f"\n{'='*55}")
    print(f"  ✅ Success: {success}  |  ❌ Failed: {failed}")
    print(f"{'='*55}\n")

# ── ENTRY POINT ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(args) < 1:
        print(__doc__)
        sys.exit(0)
    
    target = args[0]
    target_path = Path(target)
    
    # Auto-detect: FILE or FOLDER?
    if target_path.is_file() and target_path.suffix.lower() == ".pdf":
        # Single PDF file
        out_path = args[1] if len(args) > 1 and not args[1].startswith("--") else None
        success = process_single_file(str(target_path), out_path)
        sys.exit(0 if success else 1)
    
    elif target_path.is_dir():
        # Folder - batch mode
        process_folder(target_path)
        sys.exit(0)
    
    else:
        print(f"❌ Error: {target}")
        print(f"   Not a valid PDF file or folder")
        sys.exit(1)
