import os
from pdf2image import convert_from_path
import pytesseract
from docx import Document
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Path to tesseract (optional, for Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# File paths
pdf_path = "your_file.pdf"
image_output_dir = "pdf_pages"
docx_output_dir = "docx_pages"
excel_output_file = "final_output.xlsx"

os.makedirs(image_output_dir, exist_ok=True)
os.makedirs(docx_output_dir, exist_ok=True)

# ---------- STEP 1: Convert PDF pages to images ----------
print("📄 Converting PDF pages to images...")
pages = convert_from_path(pdf_path, dpi=300)
image_files = []

for i, page in enumerate(pages):
    img_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(img_path, "PNG")
    image_files.append(img_path)
    print(f"✅ Saved image: {img_path}")

# ---------- STEP 2: OCR Each Image and Save as Word Pages ----------
print("🔍 Running OCR and saving as Word files...")
for i, img_path in enumerate(image_files):
    text = pytesseract.image_to_string(img_path, config='--psm 6')  # preserve layout
    doc = Document()
    for line in text.split('\n'):
        doc.add_paragraph(line)
    docx_path = os.path.join(docx_output_dir, f"page_{i+1}.docx")
    doc.save(docx_path)
    print(f"✅ Saved Word: {docx_path}")

# ---------- STEP 3: Convert Word Pages to Excel Sheets ----------
print("📊 Creating Excel file with one sheet per page...")
wb = Workbook()
wb.remove(wb.active)

for i, docx_file in enumerate(sorted(os.listdir(docx_output_dir))):
    if docx_file.endswith(".docx"):
        doc_path = os.path.join(docx_output_dir, docx_file)
        doc = Document(doc_path)
        sheet_name = f"Page_{i+1}"
        rows = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                rows.append([text])

        df = pd.DataFrame(rows)
        ws = wb.create_sheet(title=sheet_name)
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)
        print(f"📄 Sheet added: {sheet_name}")

wb.save(excel_output_file)
print(f"✅ Excel saved at: {excel_output_file}")
