# generate_sample_w2.py
# Creates a synthetic W-2 PDF with procedurally generated placeholder data.
# No real names, SSNs, or addresses — all values are clearly fabricated.
#
# Usage:
#   python generate_sample_w2.py              # writes sample_w2.pdf (1 page)
#   python generate_sample_w2.py 5            # writes sample_w2.pdf (5 pages)

import sys
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

_FIRST   = ["ALPHA", "BETA", "GAMMA", "DELTA", "EPSILON"]
_LAST    = ["TESTUSER", "SAMPLE", "PLACEHOLDER", "EXAMPLE", "DEMO"]
_STREETS = ["100 TEST STREET", "200 SAMPLE AVE", "300 DEMO BLVD", "400 PLACEHOLDER RD", "500 EXAMPLE LN"]
_CITIES  = ["TESTCITY ST 00001", "SAMPLETOWN ST 00002", "DEMOCITY ST 00003", "PLACEHOLDER ST 00004", "EXAMPLEVILLE ST 00005"]
_EMPLOYERS = ["TEST EMPLOYER A", "SAMPLE CORP B", "DEMO COMPANY C", "PLACEHOLDER INC D", "EXAMPLE LLC E"]


def _make_employee(i: int) -> dict:
    n = i % 5
    ssn_base = (n + 1) * 111
    return {
        "name":           f"{_FIRST[n]} {_LAST[n]}",
        "ssn":            f"{ssn_base:03d}-{(n+1)*11:02d}-{(n+1)*1111:04d}",
        "street":         _STREETS[n],
        "city_state_zip": _CITIES[n],
        "employer":       _EMPLOYERS[n],
        "ein":            f"{(n+1)*11:02d}-{(n+1)*1111111:07d}",
    }


def _draw_w2_page(c: canvas.Canvas, emp: dict, page_width: float, page_height: float) -> None:
    y = page_height - 50

    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, y, "W-2 Wage and Tax Statement")
    y -= 30

    c.setFont("Helvetica", 9)
    c.drawString(50, y, "a  Employee's social security number")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(300, y, emp["ssn"])
    y -= 25

    c.setFont("Helvetica", 9)
    c.drawString(50, y, "b  Employer identification number (EIN)")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(300, y, emp["ein"])
    y -= 25

    c.setFont("Helvetica", 9)
    c.drawString(50, y, "c  Employer's name, address, and ZIP code")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(300, y, emp["employer"])
    y -= 25

    c.setFont("Helvetica", 9)
    c.drawString(50, y, "e/f Employee's name, address, and zip code")
    y -= 18
    c.setFont("Helvetica-Bold", 10)
    c.drawString(50, y, emp["name"])
    y -= 16
    c.drawString(50, y, emp["street"])
    y -= 16
    c.drawString(50, y, emp["city_state_zip"])
    y -= 30

    c.setFont("Helvetica", 9)
    boxes = [
        ("1  Wages, tips, other compensation", "00,000.00"),
        ("2  Federal income tax withheld",      "00,000.00"),
        ("3  Social security wages",            "00,000.00"),
        ("4  Social security tax withheld",     "00,000.00"),
        ("5  Medicare wages and tips",          "00,000.00"),
        ("6  Medicare tax withheld",            "00,000.00"),
    ]
    for label, value in boxes:
        c.drawString(50, y, label)
        c.drawString(320, y, value)
        y -= 18

    c.setFont("Helvetica-Oblique", 8)
    c.drawString(50, y - 20, "SYNTHETIC DATA - NOT A REAL W-2 - FOR TESTING ONLY")


def generate(output_path: str = "sample_w2.pdf", num_pages: int = 1) -> None:
    c = canvas.Canvas(output_path, pagesize=letter)
    w, h = letter

    for i in range(num_pages):
        _draw_w2_page(c, _make_employee(i), w, h)
        c.showPage()

    c.save()
    print(f"Generated: {output_path}  ({num_pages} page(s))")
    print("\nTest with:")
    print(f"  python w2_extractor.py {output_path}")
    print(f"  python w2_extractor.py {output_path} --debug")


if __name__ == "__main__":
    pages = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    out = sys.argv[2] if len(sys.argv) > 2 else "sample_w2.pdf"
    generate(out, pages)
