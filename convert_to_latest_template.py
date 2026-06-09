import pandas as pd
import re
from pathlib import Path

def clean_text(text):
    """Clean and trim whitespace"""
    if pd.isna(text):
        return ""
    text = str(text).strip()
    # Remove extra spaces
    text = re.sub(r'\s+', ' ', text)
    return text

def parse_name(name_str):
    """Parse Name into First Name and Last Name"""
    if not name_str:
        return "", ""

    name_str = clean_text(name_str)
    parts = name_str.split()

    if len(parts) == 0:
        return "", ""
    elif len(parts) == 1:
        return parts[0], ""
    else:
        # Last word is last name, rest is first name
        first_name = " ".join(parts[:-1])
        last_name = parts[-1]
        return first_name, last_name

def parse_city_state_zip(city_state_zip_str):
    """Parse City State ZIP into separate fields"""
    if not city_state_zip_str:
        return "", "", ""

    city_state_zip_str = clean_text(city_state_zip_str)

    # Try to extract Zip Code (last 5 digits or 9-digit zip)
    zip_match = re.search(r'(\d{5}(?:-\d{4})?)\s*$', city_state_zip_str)
    zip_code = zip_match.group(1) if zip_match else ""

    # Remove zip from string to get city and state
    remaining = re.sub(r'\s*\d{5}(?:-\d{4})?\s*$', '', city_state_zip_str).strip()

    # Try to extract state (2 uppercase letters)
    state_match = re.search(r'([A-Z]{2})\s*$', remaining)
    state = state_match.group(1) if state_match else ""

    # Rest is city
    city = re.sub(r'\s+[A-Z]{2}\s*$', '', remaining).strip()

    return city, state, zip_code

def convert_file(input_file_path, output_file_path="@Latest Template.xlsx"):
    """
    Convert input file to Latest Template format

    Args:
        input_file_path: Path to input Excel file with columns:
                        File Name, Page Number, Name, Street, Apt, City State ZIP
        output_file_path: Output file path (default: @Latest Template.xlsx)
    """

    # Read input file
    print(f"Reading {input_file_path}...")
    df_input = pd.read_excel(input_file_path)

    # Expected input columns
    expected_columns = ['File Name', 'Page Number', 'Name', 'Street', 'Apt', 'City State ZIP']

    # Check if columns exist
    missing_cols = [col for col in expected_columns if col not in df_input.columns]
    if missing_cols:
        print(f"Warning: Missing columns: {missing_cols}")

    # Create output dataframe with template structure
    output_data = []

    for idx, row in df_input.iterrows():
        # Extract and clean fields
        file_name = clean_text(row.get('File Name', ''))
        first_name, last_name = parse_name(row.get('Name', ''))
        street = clean_text(row.get('Street', ''))
        apt = clean_text(row.get('Apt', ''))
        city, state, zip_code = parse_city_state_zip(row.get('City State ZIP', ''))

        # Combine Street and Apt for Residential Address
        residential_address = street
        if apt:
            residential_address = f"{street} {apt}".strip() if street else apt

        # Create output row
        output_row = {
            'DOCID': file_name,
            'Last Name': last_name,
            'First Name': first_name,
            'Middle Name': '',
            'Suffix': '',
            'Data Subject Type': '',
            'Residential Address': residential_address,
            'State of Residence (if US)': state,
            'Country of Residence': 'USA',
            'City': city,
            'Province of Residence (if Canada)': '',
            'Zip Code': zip_code,
            'Address Comments': f"Apt: {apt}" if apt else '',
            'Phone Number': '',
            'Email Address - Personal': '',
            'PI Notes': '',
            'Contact Information': '',
            'Government- Issued Identification': '',
            'Social Security Number': '',
            'Passport Number': '',
            'Passport Country': '',
            "Driver's License Number": '',
            'DL Issuing Country': '',
            'DL Issuing Province (if Canada)': '',
            'DL Issuing State (if US)': '',
            'Government-Issued ID Number': '',
            'Government ID Issuing Country': '',
            'Health Related Information': '',
            'Birth Information': '',
            'Full Date of Birth (MM/DD/YYYY)': '',
            'Financial Account Information': '',
            'Access Credentials (Non-Financial Account)': '',
            'Biometric Data': '',
            'Demographic Information': '',
            'Family Information': '',
            'Student-Related Information': '',
            'Work-Related Information': '',
            'Employee Identification Number': ''
        }

        output_data.append(output_row)

    # Create output dataframe
    df_output = pd.DataFrame(output_data)

    # Save to Excel
    output_path = Path(output_file_path)
    print(f"Writing {output_path.name}...")
    df_output.to_excel(output_file_path, sheet_name='Data', index=False)

    print(f"\n✓ Done! Converted {len(output_data)} rows")
    print(f"Output: {output_file_path}")
    print(f"\nSample output:")
    print(df_output[['DOCID', 'Last Name', 'First Name', 'Residential Address', 'City', 'State of Residence (if US)', 'Zip Code']].head())

if __name__ == "__main__":
    # Usage: python convert_to_latest_template.py <input_file_path>
    import sys

    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        # Default input file
        input_file = "input.xlsx"

    convert_file(input_file)
