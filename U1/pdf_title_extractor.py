"""
PDF Title Extractor
Extracts the document title (shown in PDF tab) from all PDFs in a folder.

Setup:
    pip install pymupdf

Run:
    python pdf_title_extractor.py
    python pdf_title_extractor.py "C:\path\to\pdfs"
"""

import os
import sys
import csv


def extract_titles(folder: str) -> list[dict]:
    pdf_files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))

    if not pdf_files:
        print(f"No PDF files found in: {folder}")
        sys.exit(0)

    try:
        import fitz
    except ImportError:
        print("pymupdf not installed. Run: pip install pymupdf")
        sys.exit(1)

    results = []

    print(f"\n{'#':<5} {'Filename':<40} {'PDF Title'}")
    print("-" * 90)

    for i, filename in enumerate(pdf_files, 1):
        path = os.path.join(folder, filename)
        try:
            doc = fitz.open(path)
            title = doc.metadata.get("title", "").strip()
            doc.close()
        except Exception as e:
            title = f"ERROR: {e}"

        print(f"{i:<5} {filename:<40} {title or '(no title)'}")
        results.append({"filename": filename, "title": title})

    return results


def save_csv(results: list[dict], folder: str):
    out_path = os.path.join(folder, "pdf_titles.csv")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["filename", "title"])
        writer.writeheader()
        writer.writerows(results)
    print(f"\nSaved → {out_path}")


def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1].strip().strip('"')
    else:
        print("=" * 60)
        print("  PDF Title Extractor")
        print("=" * 60)
        folder = input("\nEnter folder path containing PDF files:\n> ").strip().strip('"')

    if not os.path.isdir(folder):
        print(f"Error: '{folder}' is not a valid folder.")
        sys.exit(1)

    results = extract_titles(folder)

    print(f"\nTotal: {len(results)} PDF(s) processed.")

    save = input("\nSave results to pdf_titles.csv? (y/n): ").strip().lower()
    if save == "y":
        save_csv(results, folder)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
