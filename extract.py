import pandas as pd
import csv
import re

# Input / Output files
input_csv = "input.csv"
output_excel = "output.xlsx"

# Search keyword
keyword = "PURCHASE, NY 10577"

# Store final records
records = []

# Read CSV
with open(input_csv, 'r', encoding='utf-8', errors='ignore') as file:
    reader = csv.reader(file)
    rows = list(reader)

# Loop using index
for i in range(len(rows)):

    row = rows[i]

    # Skip bad rows
    if len(row) < 3:
        continue

    file_name = row[0].strip()
    page_num = row[1].strip()
    text = row[2].strip()

    # Find keyword
    if keyword.lower() in text.lower():

        extracted = [text]

        # Capture next 5 valid lines
        j = i + 1

        while j < len(rows) and len(extracted) < 6:

            next_row = rows[j]

            if len(next_row) >= 3:

                next_text = next_row[2].strip()

                # Skip blanks
                if next_text:

                    # Skip pdf names
                    if not next_text.lower().endswith(".pdf"):

                        # Skip only numbers
                        if not re.fullmatch(r"\d+", next_text):

                            extracted.append(next_text)

            j += 1

        # Fill blanks
        while len(extracted) < 6:
            extracted.append("")

        # Save record with correct filename
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

# Excel columns
columns = [
    "File Name",
    "Page Number",
    "Line_1",
    "Line_2",
    "Line_3",
    "Line_4",
    "Line_5",
    "Line_6"
]

# Create dataframe
df = pd.DataFrame(records, columns=columns)

# Save Excel
df.to_excel(output_excel, index=False)

print(f"Done! {len(records)} records exported to {output_excel}")
