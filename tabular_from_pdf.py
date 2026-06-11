import pdfplumber
import pandas as pd
import os

pdf_folder = r"C:\Users\dsingh5\Downloads\1136_DocumentDownload_04172026222716\DocumentDownload_04172026222716\Native\S2"
output_csv = r"NNN1.csv"

all_rows = []

for filename in os.listdir(pdf_folder):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, filename)
        print(f"Processing file: {filename}")
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                print(f"  Extracting tables from page {page_num}...")
                try:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            all_rows.append([filename, page_num] + row)
                except Exception as e:
                    print(f"❌ Error in {filename} page {page_num}: {e}")
if all_rows:
    max_cols = max(len(row) for row in all_rows)
    columns = ["File Name", "Page Number"] + [f"Col_{i+1}" for i in range(max_cols-2)]
    df = pd.DataFrame(all_rows, columns=columns)
    df.to_csv(output_csv, index=False, encoding='utf-8-sig')
    print(f"✅ Saved: {output_csv}")
else:
    print("No tables found.")