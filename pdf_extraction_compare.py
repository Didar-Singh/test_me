#!/usr/bin/env python3
"""
PDF Text Extraction Comparison Tool
Tests multiple libraries and outputs results with library names
"""

import sys
import os
from pathlib import Path

# Color codes for output
GREEN = "\033[92m"
BLUE = "\033[94m"
YELLOW = "\033[93m"
RED = "\033[91m"
RESET = "\033[0m"
DIVIDER = "=" * 80

def test_pypdf(pdf_path):
    """Extract using pypdf library"""
    try:
        from pypdf import PdfReader
        reader = PdfReader(pdf_path)
        text = ""
        for i, page in enumerate(reader.pages):
            text += f"--- PAGE {i+1} ---\n"
            text += page.extract_text() or ""
            text += "\n"
        return text
    except ImportError:
        return "❌ pypdf not installed: pip install pypdf"
    except Exception as e:
        return f"❌ Error: {str(e)}"

def test_pdfplumber(pdf_path):
    """Extract using pdfplumber library"""
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text += f"--- PAGE {i+1} ---\n"
                text += page.extract_text() or ""
                text += "\n"
        return text
    except ImportError:
        return "❌ pdfplumber not installed: pip install pdfplumber"
    except Exception as e:
        return f"❌ Error: {str(e)}"

def test_fitz(pdf_path):
    """Extract using PyMuPDF (fitz) library"""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        text = ""
        for i, page in enumerate(doc):
            text += f"--- PAGE {i+1} ---\n"
            text += page.get_text()
            text += "\n"
        return text
    except ImportError:
        return "❌ PyMuPDF not installed: pip install PyMuPDF"
    except Exception as e:
        return f"❌ Error: {str(e)}"

def test_pdftotext_cli(pdf_path):
    """Extract using pdftotext command-line tool"""
    try:
        import subprocess
        import tempfile
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt') as tmp:
            tmp_path = tmp.name
        
        result = subprocess.run(
            ["pdftotext", "-layout", pdf_path, tmp_path],
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            with open(tmp_path, 'r') as f:
                text = f.read()
            os.unlink(tmp_path)
            return text
        else:
            os.unlink(tmp_path)
            return f"❌ pdftotext CLI failed: {result.stderr}"
    except FileNotFoundError:
        return "❌ pdftotext CLI not installed: apt-get install poppler-utils"
    except Exception as e:
        return f"❌ Error: {str(e)}"

def test_pdfminer(pdf_path):
    """Extract using pdfminer library"""
    try:
        from pdfminer.high_level import extract_text
        text = extract_text(pdf_path)
        return text
    except ImportError:
        return "❌ pdfminer.six not installed: pip install pdfminer.six"
    except Exception as e:
        return f"❌ Error: {str(e)}"

def save_results(results, pdf_name):
    """Save all results to separate files"""
    output_dir = Path("/mnt/user-data/outputs")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"\n{BLUE}📁 Saving individual outputs:{RESET}")
    
    for lib_name, content in results.items():
        filename = f"extract_{lib_name}_{pdf_name}.txt"
        filepath = output_dir / filename
        with open(filepath, 'w') as f:
            f.write(content)
        print(f"   ✓ {filename}")
    
    # Save comparison report
    report_file = output_dir / f"extraction_comparison_{pdf_name}.txt"
    with open(report_file, 'w') as f:
        f.write("PDF TEXT EXTRACTION COMPARISON REPORT\n")
        f.write(f"File: {pdf_name}\n")
        f.write(DIVIDER + "\n\n")
        
        for lib_name, content in results.items():
            f.write(f"\n{'='*60}\n")
            f.write(f"LIBRARY: {lib_name}\n")
            f.write(f"{'='*60}\n\n")
            f.write(content)
            f.write("\n\n")
    
    print(f"   ✓ extraction_comparison_{pdf_name}.txt")

def main():
    if len(sys.argv) < 2:
        print(f"{RED}Usage: python pdf_extraction_compare.py <pdf_file>{RESET}")
        print(f"\nExample: python pdf_extraction_compare.py document.pdf")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    
    if not os.path.exists(pdf_path):
        print(f"{RED}❌ File not found: {pdf_path}{RESET}")
        sys.exit(1)
    
    pdf_name = Path(pdf_path).stem
    
    print(f"\n{GREEN}{DIVIDER}{RESET}")
    print(f"{GREEN}PDF TEXT EXTRACTION COMPARISON TOOL{RESET}")
    print(f"{GREEN}{DIVIDER}{RESET}")
    print(f"\n📄 Testing: {YELLOW}{pdf_path}{RESET}")
    print(f"Processing with all available libraries...\n")
    
    # Test all libraries
    results = {}
    
    libraries = [
        ("pypdf", test_pypdf),
        ("pdfplumber", test_pdfplumber),
        ("PyMuPDF (fitz)", test_fitz),
        ("pdftotext (CLI)", test_pdftotext_cli),
        ("pdfminer.six", test_pdfminer),
    ]
    
    for lib_name, test_func in libraries:
        print(f"Testing {BLUE}{lib_name}...{RESET}", end=" ", flush=True)
        output = test_func(pdf_path)
        results[lib_name] = output
        
        if output.startswith("❌"):
            print(f"{RED}{output}{RESET}")
        else:
            lines = output.count('\n')
            print(f"{GREEN}✓ ({lines} lines){RESET}")
    
    # Display results
    print(f"\n{DIVIDER}")
    print(f"{GREEN}EXTRACTION RESULTS{RESET}")
    print(f"{DIVIDER}\n")
    
    for lib_name, content in results.items():
        print(f"\n{BLUE}{'='*60}{RESET}")
        print(f"{BLUE}LIBRARY: {lib_name}{RESET}")
        print(f"{BLUE}{'='*60}{RESET}\n")
        
        if content.startswith("❌"):
            print(f"{RED}{content}{RESET}")
        else:
            # Show first 500 chars of output
            preview = content[:500]
            print(preview)
            if len(content) > 500:
                print(f"\n{YELLOW}... (truncated - see output files for full content){RESET}")
    
    # Save all results
    print(f"\n{DIVIDER}")
    save_results(results, pdf_name)
    
    print(f"\n{GREEN}✓ Comparison complete!{RESET}")
    print(f"{YELLOW}Check /mnt/user-data/outputs/ for full results{RESET}\n")

if __name__ == "__main__":
    main()
