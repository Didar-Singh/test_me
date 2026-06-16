#!/usr/bin/env python3
"""
Simple PDF extraction for Windows - No font dependencies
"""

import sys
from pathlib import Path

def extract_with_pdfplumber(pdf_path):
    """Extract text with pdfplumber (works on Windows)"""
    try:
        import pdfplumber
        print("✓ Using pdfplumber...")
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text += f"=== PAGE {i+1} ===\n"
                text += page.extract_text() or ""
                text += "\n"
        return text, "pdfplumber"
    except ImportError:
        print("Install: pip install pdfplumber")
        return None, None

def extract_with_pypdf(pdf_path):
    """Extract text with pypdf (works on Windows)"""
    try:
        from pypdf import PdfReader
        print("✓ Using pypdf...")
        reader = PdfReader(pdf_path)
        text = ""
        for i, page in enumerate(reader.pages):
            text += f"=== PAGE {i+1} ===\n"
            text += page.extract_text() or ""
            text += "\n"
        return text, "pypdf"
    except ImportError:
        print("Install: pip install pypdf")
        return None, None

def fix_missing_commas(text):
    """
    Try to recover missing commas in LASTNAME-FIRSTNAME pattern
    Adjust this regex based on your actual name format
    """
    import re
    
    # Pattern: WORD + WORD where first is all caps, second is mixed case
    # Example: FORSYTHCHARLES becomes FORSYTH,CHARLES
    fixed = re.sub(r'([A-Z]+)([A-Z][a-z]+)', r'\1,\2', text)
    return fixed

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_pdf.py your_file.pdf")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    
    if not Path(pdf_path).exists():
        print(f"❌ File not found: {pdf_path}")
        sys.exit(1)
    
    print(f"📄 Processing: {pdf_path}\n")
    
    # Try pdfplumber first
    text, lib = extract_with_pdfplumber(pdf_path)
    
    if text is None:
        # Fall back to pypdf
        text, lib = extract_with_pypdf(pdf_path)
    
    if text is None:
        print("❌ Install pdfplumber or pypdf")
        sys.exit(1)
    
    # Save raw output
    output_file = f"output_raw.txt"
    with open(output_file, "w") as f:
        f.write(text)
    print(f"✓ Saved to: {output_file}")
    
    # Try to fix missing commas
    print("\n🔧 Attempting to recover missing commas...")
    fixed_text = fix_missing_commas(text)
    
    fixed_file = f"output_fixed.txt"
    with open(fixed_file, "w") as f:
        f.write(fixed_text)
    print(f"✓ Saved to: {fixed_file}")
    
    # Show preview
    print("\n" + "="*60)
    print("PREVIEW (first 500 chars):")
    print("="*60)
    print(text[:500])
    
    print("\n" + "="*60)
    print("PREVIEW WITH COMMA FIX (first 500 chars):")
    print("="*60)
    print(fixed_text[:500])

if __name__ == "__main__":
    main()
