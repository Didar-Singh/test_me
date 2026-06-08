import pandas as pd
import csv
import re

# Files
input_csv = "input.csv"
output_excel = "output.xlsx"

# Keyword
keyword = "PURCHASE, NY 10577"

# Store records
records = []

# Read CSV
with open(input_csv, 'r', encoding='utf-8', errors='ignore') as file:
    reader = csv.reader(file)

    # Convert all rows
    rows = list(reader)

# Process row by row
for row in rows:

    if len(row) < 3:
        continue

    file_name = row[0].strip()
    page_num = row[1].strip()
    text = row[2].strip()

    # Search keyword
    if keyword.lower() in text.lower():

        extracted = []
        extracted.append(text)

        # Start checking next rows
        current_index = rows.index(row) + 1

        while current_index < len(rows) and len(extracted) < 6:

            next_row = rows[current_index]

            if len(next_row) >= 3:

                next_text = next_row[2].strip()

                # Skip blank lines
                if next_text:

                    # Skip PDF filenames
                    if not next_text.lower().endswith(".pdf"):

                        # Skip only numbers
                        if not re.fullmatch(r"\d+", next_text):

                            extracted.append(next_text)

            current_index += 1

        # Add blanks if needed
        while len(extracted) < 6:
            extracted.append("")

        # Final row
        records.append([
            file_name,
            page_num,
            extracted[0],
            extracted[1],
            extracted[2],
            extracted[3],
            extracted[4],
            extracted[5]
        ])

# Create dataframe
columns = [
    "File Name",
    "Page Num",
    "Line_1",
    "Line_2",
    "Line_3",
    "Line_4",
    "Line_5",
    "Line_6"
]

df = pd.DataFrame(records, columns=columns)

# Save Excel
df.to_excel(output_excel, index=False)

print(f"Done! {len(records)} records saved to {output_excel}")
