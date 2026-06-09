import pandas as pd
import csv
import re

# Input / Output files
input_csv = "input.csv"
output_excel = "output_names.xlsx"

# Suffixes to recognize
SUFFIXES = ['Jr', 'Sr', 'II', 'III', 'IV', 'V', 'MD', 'Esq', 'PhD', 'DDS', 'DVM', 'Ph.D', 'M.D']

def extract_employee_id(text):
    """Extract Employee ID from 'File:' pattern"""
    match = re.search(r'File:\s*(\d+)', text, re.IGNORECASE)
    if match:
        return match.group(1)
    return ""

def is_valid_name(name):
    """Check if name looks like a person's name (alphabetic, reasonable length)"""
    if not name:
        return False
    # Must be alphabetic only, 2-30 chars
    if not re.match(r'^[A-Za-z\s\-\.]+$', name):
        return False
    if len(name) < 2 or len(name) > 30:
        return False
    return True

def parse_name(last_name_line, first_name_line=""):
    """
    Parse first name, middle name, last name, and suffix

    last_name_line: Text with comma, e.g., "HARISH," or "HARISH, 5 000 00 YEB..."
    first_name_line: Next line text, e.g., "PURI 5 000 00 968..."
    """

    # Extract LAST NAME (before comma)
    if ',' in last_name_line:
        last_name = last_name_line.split(',')[0].strip()
    else:
        last_name = ""

    # Extract text after comma (may have first name)
    text_after_comma = ""
    if ',' in last_name_line:
        text_after_comma = last_name_line.split(',', 1)[1].strip()

    # Determine which text to use for first name
    # If text after comma starts with a number/special char, use next line instead
    if text_after_comma and not re.match(r'^\d', text_after_comma):
        first_name_text = text_after_comma
    else:
        first_name_text = first_name_line

    # Extract words before numbers/special chars
    words = []
    for word in first_name_text.split():
        # Stop if we hit numbers or special patterns
        if re.match(r'^\d+', word) or word.lower() in ['yeb', 'n-', 'x', 'mcttax', 'fit', 'ny', 'w', 'chk', 'voucher', 'file']:
            break
        words.append(word)

    first_name = ""
    middle_name = ""
    suffix = ""

    if words:
        # Check if last word is a suffix
        if words[-1] in SUFFIXES or words[-1].rstrip(',') in SUFFIXES:
            suffix = words[-1].rstrip(',')
            words = words[:-1]

        # First word is first name
        if words:
            first_name = words[0]

        # Remaining words (or single letters) are middle names
        if len(words) > 1:
            middle_parts = words[1:]
            middle_name = ' '.join(middle_parts)

    return {
        'first_name': first_name,
        'middle_name': middle_name,
        'last_name': last_name,
        'suffix': suffix
    }

# Store final records
records = []

# Read CSV
with open(input_csv, 'r', encoding='utf-8', errors='ignore') as file:
    reader = csv.reader(file)
    rows = list(reader)

# Loop through rows
i = 0

while i < len(rows):
    row = rows[i]

    # Skip bad rows
    if len(row) < 3:
        i += 1
        continue

    file_number = row[0].strip()
    page_num = row[1].strip()
    text = row[2].strip()

    # Look for comma (indicates Last Name)
    if ',' in text:
        # Get next line for potential first name
        next_text = ""
        if i + 1 < len(rows) and len(rows[i + 1]) >= 3:
            next_text = rows[i + 1][2].strip()

        # Parse the name
        parsed = parse_name(text, next_text)

        # Look forward for Employee ID (File: should appear after name)
        employee_id = ""
        for j in range(i + 1, min(i + 5, len(rows))):  # Check next 4 rows
            if len(rows[j]) >= 3:
                future_text = rows[j][2].strip()
                if 'File:' in future_text:
                    employee_id = extract_employee_id(future_text)
                    break

        # Only add if we got BOTH Employee ID AND Last Name AND First Name
        # AND both names look like actual person names
        if (parsed['last_name'] and employee_id and parsed['first_name'] and
            is_valid_name(parsed['last_name']) and is_valid_name(parsed['first_name'])):
            records.append([
                file_number,
                page_num,
                employee_id,
                parsed['first_name'],
                parsed['middle_name'],
                parsed['last_name'],
                parsed['suffix'],
                text  # Keep original text for reference
            ])

    i += 1

# Excel columns
columns = [
    "File Number",
    "Page Number",
    "Employee ID",
    "First Name",
    "Middle Name",
    "Last Name",
    "Suffix",
    "Original Text"
]

# Create dataframe
df = pd.DataFrame(records, columns=columns)

# Save Excel
df.to_excel(output_excel, index=False)

print(f"Done! {len(records)} names extracted to {output_excel}")
