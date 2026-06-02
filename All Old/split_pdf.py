import os
import sys
import math
from pathlib import Path


def split_pdf(input_path: str, pages_per_chunk: int = 900, output_dir: str = None):
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter
        except ImportError:
            print("Error: Install pypdf first:  pip install pypdf")
            sys.exit(1)

    input_path = Path(input_path)
    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    if output_dir is None:
        output_dir = input_path.parent / f"{input_path.stem}_split"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Reading: {input_path}")
    reader = PdfReader(str(input_path))
    total_pages = len(reader.pages)
    total_chunks = math.ceil(total_pages / pages_per_chunk)

    print(f"Total pages : {total_pages}")
    print(f"Pages/file  : {pages_per_chunk}")
    print(f"Output files: {total_chunks}")
    print(f"Output dir  : {output_dir}\n")

    for chunk_idx in range(total_chunks):
        start = chunk_idx * pages_per_chunk
        end = min(start + pages_per_chunk, total_pages)

        writer = PdfWriter()
        for page_num in range(start, end):
            writer.add_page(reader.pages[page_num])

        out_filename = f"{input_path.stem}_part{chunk_idx + 1:03d}_pages{start + 1}-{end}.pdf"
        out_path = output_dir / out_filename

        with open(out_path, "wb") as f:
            writer.write(f)

        print(f"  [{chunk_idx + 1}/{total_chunks}] {out_filename}  ({end - start} pages)")

    print(f"\nDone. {total_chunks} files written to: {output_dir}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python split_pdf.py <input.pdf> [pages_per_chunk] [output_dir]")
        print()
        print("Examples:")
        print("  python split_pdf.py large.pdf")
        print("  python split_pdf.py large.pdf 900")
        print("  python split_pdf.py large.pdf 900 C:/output/folder")
        sys.exit(0)

    input_file = sys.argv[1]
    chunk_size = int(sys.argv[2]) if len(sys.argv) > 2 else 900
    out_dir = sys.argv[3] if len(sys.argv) > 3 else None

    split_pdf(input_file, chunk_size, out_dir)
