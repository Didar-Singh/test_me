"""
╔═══════════════════════════════════════════════════════════════════════════════╗
║                          CSV SPLITTER SCRIPT                                  ║
║                   Split Large CSV Files into Chunks                           ║
╚═══════════════════════════════════════════════════════════════════════════════╝

USAGE EXAMPLES:
==============

1. BASIC USAGE - Split with default settings (800,000 rows per file):
   ─────────────────────────────────────────────────────────────────
   from split_csv import split_csv
   split_csv("large_file.csv")

2. CUSTOM CHUNK SIZE - Split into 500,000 rows per file:
   ─────────────────────────────────────────────────────────────────
   split_csv("large_file.csv", chunk_size=500000)

3. CUSTOM OUTPUT DIRECTORY - Save splits in specific folder:
   ─────────────────────────────────────────────────────────────────
   split_csv("large_file.csv", chunk_size=800000, output_dir="./output")

4. MEMORY OPTIMIZED MODE - For 10M+ rows (recommended):
   ─────────────────────────────────────────────────────────────────
   from split_csv import split_csv_memory_optimized
   split_csv_memory_optimized("large_file.csv")

5. COMMAND LINE - Run directly with configuration:
   ─────────────────────────────────────────────────────────────────
   python split_csv.py

FEATURES:
=========
✓ Splits large CSV files into manageable chunks
✓ Preserves headers in every output file
✓ Auto-numbered output files (part_1.csv, part_2.csv, etc.)
✓ Progress indicator showing each chunk created
✓ Two methods: Fast (pandas) & Memory-Optimized (csv module)
✓ Handles 10M+ rows efficiently

REQUIREMENTS:
=============
- Python 3.6+
- pandas (optional, for Method 1 - install: pip install pandas)
- csv module (built-in Python)

CONFIGURATION:
==============
Edit the section at the bottom marked "CONFIGURATION" to set:
  - csv_file: Path to your large CSV file
  - rows_per_chunk: Rows per output file (default: 800000)
  - output_directory: Where to save splits (default: same folder as input)

"""

import pandas as pd
import os
from pathlib import Path

def split_csv(input_file, chunk_size=800000, output_dir=None):
    """
    Split a large CSV file into smaller chunks.
    
    Args:
        input_file: Path to the input CSV file
        chunk_size: Number of rows per output file (default: 800,000)
        output_dir: Directory to save output files (default: same as input file)
    """
    
    # Validate input file
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found!")
        return
    
    # Set output directory
    if output_dir is None:
        output_dir = os.path.dirname(input_file) or "."
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Get base filename without extension
    base_name = Path(input_file).stem
    
    print(f"Reading {input_file}...")
    print(f"Chunk size: {chunk_size:,} rows")
    print(f"Output directory: {output_dir}\n")
    
    chunk_number = 1
    total_rows = 0
    
    # Read CSV in chunks
    for chunk in pd.read_csv(input_file, chunksize=chunk_size):
        output_file = os.path.join(
            output_dir, 
            f"{base_name}_part_{chunk_number}.csv"
        )
        
        chunk.to_csv(output_file, index=False)
        rows_in_chunk = len(chunk)
        total_rows += rows_in_chunk
        
        print(f"✓ Part {chunk_number}: {rows_in_chunk:,} rows → {output_file}")
        chunk_number += 1
    
    print(f"\n{'='*60}")
    print(f"Split complete!")
    print(f"Total rows processed: {total_rows:,}")
    print(f"Total files created: {chunk_number - 1}")
    print(f"{'='*60}")


def split_csv_memory_optimized(input_file, chunk_size=800000, output_dir=None):
    """
    Alternative method using csv module for even lower memory usage.
    Better for extremely large files.
    """
    import csv
    
    # Validate input file
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found!")
        return
    
    # Set output directory
    if output_dir is None:
        output_dir = os.path.dirname(input_file) or "."
    
    os.makedirs(output_dir, exist_ok=True)
    base_name = Path(input_file).stem
    
    print(f"Reading {input_file} (memory-optimized mode)...")
    print(f"Chunk size: {chunk_size:,} rows")
    print(f"Output directory: {output_dir}\n")
    
    with open(input_file, 'r', encoding='utf-8') as infile:
        reader = csv.reader(infile)
        
        # Read header
        header = next(reader)
        
        chunk_number = 1
        row_count = 0
        total_rows = 0
        
        # Create first output file
        output_file = os.path.join(
            output_dir, 
            f"{base_name}_part_{chunk_number}.csv"
        )
        outfile = open(output_file, 'w', newline='', encoding='utf-8')
        writer = csv.writer(outfile)
        writer.writerow(header)
        
        # Process rows
        for row in reader:
            writer.writerow(row)
            row_count += 1
            total_rows += 1
            
            # When chunk is full, start new file
            if row_count >= chunk_size:
                outfile.close()
                print(f"✓ Part {chunk_number}: {row_count:,} rows → {output_file}")
                
                chunk_number += 1
                row_count = 0
                
                output_file = os.path.join(
                    output_dir, 
                    f"{base_name}_part_{chunk_number}.csv"
                )
                outfile = open(output_file, 'w', newline='', encoding='utf-8')
                writer = csv.writer(outfile)
                writer.writerow(header)
        
        # Close last file
        outfile.close()
        if row_count > 0:
            print(f"✓ Part {chunk_number}: {row_count:,} rows → {output_file}")
    
    print(f"\n{'='*60}")
    print(f"Split complete!")
    print(f"Total rows processed: {total_rows:,}")
    print(f"Total files created: {chunk_number}")
    print(f"{'='*60}")


if __name__ == "__main__":
    # ════════════════════════════════════════════════════════════════════════
    #                         CONFIGURATION SECTION
    # ════════════════════════════════════════════════════════════════════════
    
    print("\n" + "="*80)
    print("CSV SPLITTER - QUICK START GUIDE")
    print("="*80 + "\n")
    
    print("STEP 1: Edit this script and replace 'your_large_file.csv' with your file path")
    print("STEP 2: Optionally adjust 'rows_per_chunk' (default: 800,000)")
    print("STEP 3: Run: python split_csv.py\n")
    
    print("-"*80)
    print("EXAMPLES:")
    print("-"*80)
    print("  Example 1: csv_file = 'data.csv'")
    print("             → Creates: data_part_1.csv, data_part_2.csv, etc.")
    print()
    print("  Example 2: csv_file = '/home/user/Downloads/sales_data.csv'")
    print("             → Absolute path to file")
    print()
    print("  Example 3: csv_file = './data/large_dataset.csv'")
    print("             → Relative path from current directory")
    print()
    print("  Example 4: output_directory = './split_output'")
    print("             → All parts saved to split_output folder")
    print("-"*80 + "\n")
    
    # ════════════════════════════════════════════════════════════════════════
    #                    EDIT YOUR SETTINGS HERE ⬇️
    # ════════════════════════════════════════════════════════════════════════
    
    # SET YOUR CSV FILE PATH HERE (REQUIRED)
    # Examples:
    #   csv_file = "data.csv"
    #   csv_file = "C:\\Users\\YourName\\Downloads\\data.csv"  (Windows)
    #   csv_file = "/Users/yourname/data.csv"  (Mac)
    #   csv_file = "/home/user/data.csv"  (Linux)
    
    csv_file = "your_large_file.csv"  # 👈 CHANGE THIS TO YOUR FILE PATH
    
    # SET ROWS PER CHUNK (default: 800,000 = 8 lakh rows)
    # Examples:
    #   rows_per_chunk = 100000     → 1 lakh rows per file
    #   rows_per_chunk = 500000     → 5 lakh rows per file
    #   rows_per_chunk = 800000     → 8 lakh rows per file (default)
    #   rows_per_chunk = 1000000    → 10 lakh rows per file
    
    rows_per_chunk = 800000
    
    # SET OUTPUT DIRECTORY (optional)
    # Leave as None to save in same folder as input file
    # Examples:
    #   output_directory = None            → Save in same folder as input
    #   output_directory = "./output"      → Create 'output' folder
    #   output_directory = "C:\\splits"    → Windows path
    
    output_directory = None
    
    # ════════════════════════════════════════════════════════════════════════
    #                      SELECT METHOD TO RUN
    # ════════════════════════════════════════════════════════════════════════
    
    print("\n📊 Starting CSV split operation...\n")
    
    # CHOOSE ONE METHOD:
    
    # ✅ METHOD 1: PANDAS (Faster, requires: pip install pandas)
    split_csv(csv_file, chunk_size=rows_per_chunk, output_dir=output_directory)
    
    # ⭐ METHOD 2: MEMORY OPTIMIZED (Better for 10M+ rows, no dependencies)
    # Uncomment the line below and comment out METHOD 1 above to use this
    # split_csv_memory_optimized(csv_file, chunk_size=rows_per_chunk, output_dir=output_directory)
    
    print("\n" + "="*80)
    print("✨ CSV SPLIT COMPLETE! ✨")
    print("="*80)
