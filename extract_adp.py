"""
ADP Master Control - Extract from MESSY LINEAR TEXT
Handles text where all fields are jumbled on single lines.

Usage:
  python extract_adp.py file.pdf
  python extract_adp.py file.pdf output.xlsx
  python extract_adp.py file.pdf output.xlsx --pages 1-5
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

# Find ALL employee names in full text
def find_all_names(text):
    """Find ALL names: LAST,FIRST format - any position"""
    positions = []
    # Much simpler: just find LAST,FIRST pattern anywhere
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
        for i, b in enumerate(blocks):
            print(f"  Block {i+1}: {b[:60]}...")
    
    return blocks

def parse_block(block, num=1):
    """Parse messy single-line format"""
    rec = {"_block": num}
    txt = block
    
    # NAME - just find LAST,FIRST (with optional middle) - capture until File: or other field
    name_m = re.search(r"([A-Z][A-Z\'\-\.]+),([A-Z][A-Z\'\-\.\s]+?)(?=\s+File:|\s+Marital|\s+Gross:|\s+Mailing|$)", txt, re.I)
    if name_m:
        rec["Last Name"] = clean(name_m.group(1))
        rec["First Name"] = clean(name_m.group(2))  # This now includes middle initials/names
    
    rec["Continued"] = "Yes" if "(continued)" in txt else ""
    
    # FILE # (first occurrence)
    rec["File #"] = grep(r"File:\s*(\d+)", txt)
    
    # Status, Dept, Sex, etc. - all on mixed lines
    rec["Status"] = mgrep([r"Status:\s*(\w+)", r"Status\s+(\w+)"], txt)
    rec["Dept"] = grep(r"Dept:\s*(\S+)", txt)
    rec["Sex"] = grep(r"Sex:\s*(\w)", txt)
    rec["Cntl"] = grep(r"Cntl:\s*(\S+)", txt)
    rec["Race"] = grep(r"Race:\s*(\S+)", txt)
    rec["Occup"] = grep(r"Occup:\s*(\S+)", txt)
    rec["SSN"] = grep(r"SSN:\s*([^\n]+?)(?=\s{2,}|GTL|Title|$)", txt)
    rec["Title"] = grep(r"Title:\s*(\S+)", txt)
    
    # DATES - all mixed in
    rec["Hire Date"] = grep(r"Hire:\s*([\d/]+)", txt)
    rec["Term Date"] = grep(r"Term:\s*([\d/]+)", txt)
    rec["Birth Date"] = grep(r"Birth:\s*([\d/]+)", txt)
    rec["Date 6"] = grep(r"Date\s+6:\s*([\d/]+)", txt)
    rec["Date 8"] = grep(r"Date\s+8:\s*([\d/]+)", txt)
    rec["Date 9"] = grep(r"Date\s+9:\s*([\d/]+)", txt)
    
    # ADDRESS - look for street address pattern (digits + words)
    # Address might be on separate lines or mixed with other fields
    addr_lines = []
    for m in re.finditer(r"^(\d+\s+[A-Z][A-Z\s\.\-]+(?:ST|AVE|RD|DR|LN|BLVD|CT|CRESCT|CRECENT|DRIVE)?)", txt, re.MULTILINE | re.I):
        addr_lines.append(clean(m.group(1)))
    
    # City, State, Zip pattern - look for: [optional codes] CITY STATE ZIP
    # Remove codes like Q, Y, SS, etc that come before city
    for m in re.finditer(r"(?:[A-Z][\s])*(?:Q|Y|SS)?\s*([A-Z]{2,}?)\s+([A-Z]{2})\s+(\d{5})", txt):
        city_raw = m.group(1).strip()
        # Filter: city must be real word (2+ chars, not just codes)
        if len(city_raw) > 2 and city_raw not in ["SS", "YY", "QQ"]:
            # Remove trailing code letters
            city_clean = re.sub(r"\s+[A-Z]$", "", city_raw).strip()
            if len(city_clean) > 2:
                rec["City"] = clean(city_clean)
                rec["State"] = m.group(2).upper()
                rec["Zip"] = m.group(3)
                break
    
    if addr_lines:
        rec["Address Line 1"] = addr_lines[0] if len(addr_lines) > 0 else ""
        rec["Address Line 2"] = addr_lines[1] if len(addr_lines) > 1 else ""
    
    # PAY - greedy to get full numbers even if mixed with other fields
    # Gross: followed by numbers (handle multiple Gross entries, take first meaningful one)
    gross_matches = re.findall(r"Gross:\s*([\d\s,\.]+?)(?=\s+(?:Federal|Salary|State|Form|Rate|Std|Dept|File|Marital|Q\s|$))", txt, re.I)
    if gross_matches:
        # Take first match, clean up spaces
        rec["Gross"] = clean(gross_matches[0]).replace(" ", "")
    else:
        rec["Gross"] = ""
    
    # Salary: followed by numbers
    salary_m = re.search(r"Salary:\s*([\d\s,\.]+?)(?=\s+(?:Federal|Form|Rate|2020|Std|Dept|File|Monthly|$))", txt, re.I)
    rec["Salary"] = clean(salary_m.group(1)).replace(" ", "") if salary_m else ""
    
    rec["Rate Calc"] = grep(r"Rate\s+Calc:\s*(\S+)", txt)
    rec["Std Hours"] = grep(r"Std\s+Hours:\s*([\d\.]+)", txt)
    rec["Pay Group"] = grep(r"Pay\s+Group:\s*(\d+)", txt)
    
    # TAX
    rec["Marital Status"] = grep(r"(J?-?Married|Single)", txt)
    rec["Federal Exemptions"] = mgrep([
        r"Exemptions[^0-9]*(\d+)",
        r"(\d+)\s+Exemptions",
    ], txt)
    
    # STATE TAX - extract state codes
    state_m = re.search(r"([\d]{2}\s+[A-Z]{2}(?:\s+[A-Z\s]+)?)", txt)
    if state_m:
        rec["State Tax"] = clean(state_m.group(1))
    
    # DEDUCTIONS
    rec["GTL Cov"] = grep(r"GTL\s+Cov[^0-9]*?([\d\s]+)", txt)
    rec["401K"] = grep(r"401K\s+([\d,\.]+)", txt)
    
    # DIRECT DEPOSIT - extract all
    rec["Acct #"] = mgrep([
        r"Acct\s*#\s*[:\-]?\s*([\dXx*\-]+)",
        r"Acct:\s*([\dXx*\-]+)",
    ], txt)
    rec["Tran/ABA"] = grep(r"Tran[/\\]ABA:\s*([\d\s]+)", txt)
    rec["DD Code"] = grep(r"Code\s+([A-Z])\b", txt)
    rec["DD Type"] = grep(r"(Full|Partial)\s+Deposit", txt)
    
    if DEBUG:
        print(f"\nBLOCK {num}:")
        for k, v in rec.items():
            if v and not k.startswith("_"):
                print(f"  {k:20s}: {v}")
    
    return rec

def extract_all(pdf_path, page_from=None, page_to=None):
    employees = []
    
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        p_from = max(1, page_from) if page_from else 1
        p_to = min(total, page_to) if page_to else total
        
        if p_from > total or p_from > p_to:
            print(f"\n  ❌ Invalid pages {p_from}-{p_to} (PDF has {total})")
            return []
        
        print(f"  PDF pages: {total}")
        print(f"  Processing: {p_from} to {p_to}")
        
        all_text = ""
        with tqdm(total=p_to-p_from+1, desc="📄 Reading", unit="pg", colour="cyan", ncols=65) as pbar:
            for page_num in range(p_from-1, p_to):
                page = pdf.pages[page_num]
                t = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                all_text += t + "\n"
                pbar.update(1)
        
        if not all_text.strip():
            print("\n  ❌ No text found")
            return []
        
        blocks = split_by_names(all_text)
        print(f"  Employees found: {len(blocks)}")
        
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

COLUMNS = [
    "Last Name", "First Name", "Continued",
    "File #", "Status", "Dept", "Sex", "Cntl", "Race", "Occup", "SSN", "Title",
    "Hire Date", "Term Date", "Birth Date", "Date 6", "Date 8", "Date 9",
    "Address Line 1", "Address Line 2", "City", "State", "Zip",
    "Gross", "Salary", "Rate Calc", "Std Hours", "Pay Group",
    "Marital Status", "Federal Exemptions", "State Tax",
    "GTL Cov", "401K",
    "Acct #", "Tran/ABA", "DD Code", "DD Type",
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

if __name__ == "__main__":
    if len(args) < 1:
        print(__doc__)
        sys.exit(0)
    
    pdf_path = args[0]
    out_path = args[1] if len(args) > 1 else str(Path(pdf_path).stem + "_employees.xlsx")
    
    if not Path(pdf_path).exists():
        print(f"\n❌ File not found: {pdf_path}\n")
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
        print("\n❌ No records. Try debug mode:")
        print(f"   python extract_adp.py \"{pdf_path}\" --debug\n")
        sys.exit(1)
    
    write_excel(employees, out_path)
    
    print(f"\n{'='*55}")
    print(f"  ✅ Done! → {out_path}")
    print(f"{'='*55}\n")
