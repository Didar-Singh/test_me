import pdfplumber
import pandas as pd
import re
from pathlib import Path
import sys

def extract_earnings_data(pdf_path):
    """Extract name, address, gross salary, and net pay from PDF"""
    data = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()

                if not text:
                    continue

                # Extract Name (usually appears early in the document)
                name = extract_name(text)

                # Extract Address
                address = extract_address(text)

                # Extract Gross Salary
                gross_salary = extract_gross_salary(text)

                # Extract Net Pay
                net_pay = extract_net_pay(text)

                if name or gross_salary or net_pay:  # Only add if we found something
                    data.append({
                        'Name': name or 'N/A',
                        'Address': address or 'N/A',
                        'Gross Salary': gross_salary or 'N/A',
                        'Net Pay': net_pay or 'N/A'
                    })

    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")

    return data

def extract_name(text):
    """Extract person's name from text"""
    lines = text.split('\n')

    # Common patterns for names in earnings statements
    for line in lines[:10]:  # Check first 10 lines
        line = line.strip()
        # Skip common header text
        if any(skip in line.lower() for skip in ['earnings', 'statement', 'employee', 'pay', 'date', 'period']):
            continue
        # If line has 2-4 words and doesn't contain numbers or special chars
        words = line.split()
        if 2 <= len(words) <= 4 and not any(char.isdigit() for char in line):
            return line

    return None

def extract_address(text):
    """Extract address from text"""
    lines = text.split('\n')
    address_lines = []

    # Look for address patterns after "Employee's name, address, and zip code"
    capture = False
    for line in lines:
        line = line.strip()
        if "Employee's name" in line or "employee's address" in line.lower():
            capture = True
            continue
        if capture and line and not any(skip in line.lower() for skip in ['wage', 'tax', 'social', 'medicare']):
            if len(address_lines) < 2 and line != "":
                address_lines.append(line)
            elif len(address_lines) >= 2:
                break

    return ', '.join(address_lines) if address_lines else None

def extract_gross_salary(text):
    """Extract gross salary/income from text"""
    # Look for "1 Wages, tips, other compensation" pattern in W-2
    patterns = [
        r'1\s+Wages[^0-9]*?([\d,]+\.\d{2})',
        r'Wages[^0-9]*?([\d,]+\.\d{2})',
        r'Gross[^0-9]*?([\d,]+\.\d{2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)

    return None

def extract_net_pay(text):
    """Extract net pay from text"""
    # Look for "2 Federal income tax withheld" pattern in W-2
    patterns = [
        r'2\s+Federal\s+income\s+tax[^0-9]*?([\d,]+\.\d{2})',
        r'Federal\s+income\s+tax[^0-9]*?([\d,]+\.\d{2})',
        r'Net\s+Pay[^0-9]*?([\d,]+\.\d{2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)

    return None

def process_folder(folder_path, output_excel):
    """Process all PDFs in a folder and save to Excel"""
    folder = Path(folder_path)
    all_data = []

    # Find all PDF files
    pdf_files = list(folder.rglob('*.pdf'))

    if not pdf_files:
        print(f"No PDF files found in {folder_path}")
        return

    print(f"Found {len(pdf_files)} PDF file(s)")

    for pdf_file in pdf_files:
        print(f"Processing: {pdf_file}")
        data = extract_earnings_data(str(pdf_file))
        all_data.extend(data)

    if all_data:
        # Create DataFrame and save to Excel
        df = pd.DataFrame(all_data)
        df.to_excel(output_excel, index=False, sheet_name='Earnings Data')
        print(f"\n[SUCCESS] Data extracted successfully!")
        print(f"  Output saved to: {output_excel}")
        print(f"  Total records: {len(df)}")
        print("\nPreview:")
        print(df.to_string())
    else:
        print("No data extracted from PDFs")

if __name__ == "__main__":
    # Specify your PDF folder here
    pdf_folder = "All Old"  # Change this to your PDF folder path
    output_file = "earnings_data.xlsx"  # Output Excel file

    if len(sys.argv) > 1:
        pdf_folder = sys.argv[1]

    if len(sys.argv) > 2:
        output_file = sys.argv[2]

    process_folder(pdf_folder, output_file)
