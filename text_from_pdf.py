import pdfplumber
import pandas as pd
from tqdm import tqdm
import os

pdf_folder = r"C:\Users\DSINGH5\Downloads\Mass Arbitration Claimants\Mass Arbitration Claimants\New folder"
output_csv = "18_April.csv"


records = []


pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]

for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_folder, pdf_file)
    print(f"\n📄 Processing: {pdf_file}")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(tqdm(pdf.pages, desc="Extracting", unit="page"), start=1):
                try:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            records.append({
                                "File Name": pdf_file,
                                "Page Number": page_num,
                                "Extracted Text": line.strip()
                            })
                except Exception as e:
                    print(f"   ❌ Error on page {page_num}: {e}")
                    continue
    except Exception as e:
        print(f"🚫 Failed to open {pdf_file}: {e}")

df = pd.DataFrame(records)
df.to_csv(output_csv, index=False, encoding='utf-8-sig')
print(f"\n✅ All data saved to: {output_csv}")