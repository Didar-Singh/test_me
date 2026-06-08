import pandas as pd
import csv
import re

# Files
input_csv = "input.csv"
output_excel = "output.xlsx"

# Search text
keyword = "PURCHASE, NY 10577"

# Store results
records = []

# Read CSV
with open(input_csv, 'r', encoding='utf-8', errors='ignore') as file:
    reader = csv.reader(file)

    # Flatten all cells
    all_data = []

    for row in reader:
        for cell in row:
            cell = cell.strip()

            if cell:
                all_data.append(cell)

# Function to filter junk rows
def is_valid_text(text):

    # Ignore PDF filenames
    if text.lower().endswith(".pdf"):
        return False

    # Ignore only numbers
    if re.fullmatch(r"\d+", text):
        return False

    return True

# Process data
for i in range(len(all_data)):

    if keyword.lower() in all_data[i].lower():

        extracted = []

        # Add matched row
        extracted.append(all_data[i])

        # Get next valid 5 rows
        j = i + 1

        while j < len(all_data) and len(extracted) < 6:

            if is_valid_text(all_data[j]):
                extracted.append(all_data[j])

            j += 1

        # Fill blanks if needed
        while len(extracted) < 6:
            extracted.append("")

        records.append(extracted)

# Create Excel
columns = [
    "Line_1",
    "Line_2",
    "Line_3",
    "Line_4",
    "Line_5",
    "Line_6"
]

df = pd.DataFrame(records, columns=columns)

# Save
df.to_excel(output_excel, index=False)

print(f"Done! {len(records)} records saved to {output_excel}")
