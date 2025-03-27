import os
import pytesseract
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
docx_output = r"C:/MyPDFs/output.docx"
excel_output = r"C:/MyPDFs/output.xlsx"
image_output_dir = r"C:/MyPDFs/pdf_pages"

os.makedirs(image_output_dir, exist_ok=True)

# Optional: Tesseract path (Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# === STEP 1: Convert PDF to image temporarily to feed into OCR ===
pages = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(pages):
    img_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(img_path, "PNG")
    image_paths.append(img_path)

# === STEP 2: Create Word and Excel from OCR (editable content) ===
doc = Document()
wb = Workbook()
wb.remove(wb.active)

for idx, img_path in enumerate(image_paths):
    text = pytesseract.image_to_string(img_path, config="--psm 6").strip()

    # === Add to Word ===
    doc.add_paragraph(f"Page {idx+1}", style='Heading1')
    for line in text.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line)
            p.style.font.size = Pt(10)
    doc.add_page_break()

    # === Add to Excel ===
    lines = [[line.strip()] for line in text.split('\n') if line.strip()]
    df = pd.DataFrame(lines)
    df = df.replace('', pd.NA).dropna(how='all', axis=1)

    sheet = wb.create_sheet(title=f"Page_{idx+1}")
    for row in dataframe_to_rows(df, index=False, header=False):
        sheet.append(row)

# === Save both files ===
doc.save(docx_output)
wb.save(excel_output)

print(f"✅ Word saved to: {docx_output}")
print(f"✅ Excel saved to: {excel_output}")
