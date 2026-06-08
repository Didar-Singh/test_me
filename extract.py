import pandas as pd
import csv

# File names
input_csv = "input.csv"
output_excel = "output.xlsx"

# Keyword to search
keyword = "PURCHASE, NY 10577"

# Store final records
records = []

# Read CSV
with open(input_csv, 'r', encoding='utf-8', errors='ignore') as file:
    reader = csv.reader(file)

    # Convert all rows into simple list
    all_rows = [row for row in reader]

# Flatten all cells into one list
flat_data = []

for row in all_rows:
    for cell in row:
        if cell.strip():
            flat_data.append(cell.strip())

# Search keyword and extract next 5 rows
for i in range(len(flat_data)):

    if keyword.lower() in flat_data[i].lower():

        # Take matched row + next 5 rows
        extracted = flat_data[i:i+6]

        # Fill blank if less than 6 rows
        while len(extracted) < 6:
            extracted.append("")

        records.append(extracted)

# Create dataframe
columns = [
    "Row_1",
    "Row_2",
    "Row_3",
    "Row_4",
    "Row_5",
    "Row_6"
]

df = pd.DataFrame(records, columns=columns)

# Export to Excel
df.to_excel(output_excel, index=False)

print(f"Done! {len(records)} records exported to {output_excel}")
