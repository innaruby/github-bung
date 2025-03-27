import os
import cv2
import pytesseract
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook

# === CONFIGURATION ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
word_output_path = r"C:/MyPDFs/converted.docx"
excel_output_path = r"C:/MyPDFs/final_output.xlsx"

os.makedirs(image_output_dir, exist_ok=True)

# === STEP 1: Convert PDF pages to images ===
print("📄 Converting PDF to images...")
pages = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(pages):
    img_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(img_path, "PNG")
    image_paths.append(img_path)
    print(f"✅ Saved: {img_path}")

# === STEP 2: Create Word file with one image per page ===
print("\n📝 Creating Word document...")
doc = Document()

for img_path in image_paths:
    doc.add_picture(img_path, width=Inches(6.5))  # Fit within margins
    doc.add_page_break()

doc.save(word_output_path)
print(f"✅ Word file created: {word_output_path}")

# === STEP 3: Extract content from images (for each Word page) and write to Excel ===
print("\n📊 Extracting text for Excel...")
wb = Workbook()
wb.remove(wb.active)

for i, img_path in enumerate(image_paths):
    img = cv2.imread(img_path)
    text = pytesseract.image_to_string(img, config="--psm 6")

    # Split into rows and columns (simple line split)
    lines = text.strip().split("\n")
    data = [line.split() for line in lines if line.strip()]

    # Write to Excel
    sheet = wb.create_sheet(title=f"Page_{i+1}")
    for r_idx, row in enumerate(data, start=1):
        for c_idx, cell in enumerate(row, start=1):
            sheet.cell(row=r_idx, column=c_idx, value=cell)

    print(f"✅ Sheet created: Page_{i+1}")

wb.save(excel_output_path)
print(f"\n✅ Excel file saved at: {excel_output_path}")
