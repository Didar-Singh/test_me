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

    # Look for "e/f" marker which indicates employee info section
    for i, line in enumerate(lines):
        if "e/f" in line.lower() and "employee" in line.lower():
            # Name is the next non-empty line after the e/f marker
            if i + 1 < len(lines):
                name = lines[i + 1].strip()
                if name:
                    return name

    return None

def extract_address(text):
    """Extract address from text"""
    lines = text.split('\n')

    # Look for "e/f" marker which indicates employee info section
    for i, line in enumerate(lines):
        if "e/f" in line.lower() and "employee" in line.lower():
            # Skip name (line i+1), then get address lines
            # Typically address is in lines i+2 and i+3
            address_parts = []

            # Get line i+2 (street address)
            if i + 2 < len(lines):
                addr_line = lines[i + 2].strip()
                if addr_line and not any(skip in addr_line.lower() for skip in ['wage', 'tax', 'social', 'medicare']):
                    address_parts.append(addr_line)

            # Get line i+3 (city, state, zip)
            if i + 3 < len(lines):
                addr_line = lines[i + 3].strip()
                if addr_line and not any(skip in addr_line.lower() for skip in ['wage', 'tax', 'social', 'medicare']):
                    address_parts.append(addr_line)

            return ', '.join(address_parts) if address_parts else None

    return None

def extract_gross_salary(text):
    """Extract gross salary/income from text"""
    # Look for "1 Wages, tips, other compensation" followed by amount
    match = re.search(r'1\s+Wages[^$]*?\s+([\d,]+\.?\d*)', text, re.IGNORECASE | re.DOTALL)
    if match:
        return match.group(1).strip()

    # Fallback patterns
    patterns = [
        r'Wages[,\s]+[^$]*?\s+([\d,]+\.\d{2})',
        r'Gross\s+[^$]*?\s+([\d,]+\.\d{2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)

    return None

def extract_net_pay(text):
    """Extract net pay from text"""
    # Look for "2 Federal income tax withheld" followed by amount
    match = re.search(r'2\s+Federal\s+income\s+tax[^$]*?\s+([\d,]+\.?\d*)', text, re.IGNORECASE | re.DOTALL)
    if match:
        return match.group(1).strip()

    # Fallback patterns
    patterns = [
        r'Federal\s+income\s+tax[^$]*?\s+([\d,]+\.\d{2})',
        r'Net\s+Pay[^$]*?\s+([\d,]+\.\d{2})',
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
