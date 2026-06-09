import pandas as pd
import csv
import re

# Input / Output files
input_csv = "sample_input.csv"
output_excel = "@Latest Template.xlsx"

def parse_address_data(text):
    """
    Parse Street, Apt, City State ZIP from extracted text
    Adjust regex patterns based on your actual data format
    """
    street = ""
    apt = ""
    city_state_zip = ""

    # TODO: Add your parsing logic here based on your data format
    # Example patterns:
    # - Street address pattern
    # - Apartment/Unit pattern
    # - City, State ZIP pattern

    return street, apt, city_state_zip

def extract_name(text):
    """Extract name from extracted text"""
    # Look for name patterns (simplistic - adjust as needed)
    words = text.split()
    name = ""

    for word in words:
        # Stop at numbers or special keywords
        if re.match(r'^\d+', word) or word.lower() in ['yeb', 'fit', 'ny', 'file', 'voucher']:
            break
        name += word + " "

    return name.strip()

# Store final records
records = []

# Read CSV
with open(input_csv, 'r', encoding='utf-8', errors='ignore') as file:
    reader = csv.reader(file)
    rows = list(reader)

# Loop through rows (skip header)
for i in range(1, len(rows)):
    row = rows[i]

    if len(row) < 3:
        continue

    file_name = row[0].strip()
    page_number = row[1].strip()
    extracted_text = row[2].strip()

    # Extract fields
    name = extract_name(extracted_text)
    street, apt, city_state_zip = parse_address_data(extracted_text)

    if name:  # Only add if we have a name
        records.append({
            'File Name': file_name,
            'Page Number': page_number,
            'Name': name,
            'Street': street,
            'Apt': apt,
            'City State ZIP': city_state_zip
        })

# Create dataframe
df_output = pd.DataFrame(records)

# Save to Excel
df_output.to_excel(output_excel, sheet_name='Data', index=False)

print(f"Done! {len(records)} records extracted to {output_excel}")
