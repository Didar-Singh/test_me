import subprocess
from PIL import Image
import pytesseract
from pathlib import Path

pdf_path = "your_file.pdf"
pdf_name = Path(pdf_path).stem

# Step 1: Convert PDF pages to images
subprocess.run([
    "pdftoppm", "-png", "-r", "300", 
    pdf_path, f"/tmp/{pdf_name}"
])

# Step 2: Use OCR to read the images
for image_file in Path("/tmp/").glob(f"{pdf_name}*.png"):
    text = pytesseract.image_to_string(str(image_file))
    print(f"\n--- {image_file.name} ---")
    print(text)
    
    # Save to file
    with open(f"ocr_output_{image_file.stem}.txt", "w") as f:
        f.write(text)
