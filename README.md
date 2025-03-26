import os
import cv2
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

# CONFIG
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"

os.makedirs(image_output_dir, exist_ok=True)

# STEP 1: Convert PDF to images
print("📄 Converting PDF to images...")
pages = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(pages):
    img_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(img_path, "PNG")
    image_paths.append(img_path)
    print(f"✅ Image saved: {img_path}")

# STEP 2: Excel workbook init
wb = Workbook()
wb.remove(wb.active)

# STEP 3: Process each image
for i, image_path in enumerate(image_paths):
    print(f"🔍 Processing: {image_path}")
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Run OCR with layout info
    ocr_df = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DATAFRAME)
    ocr_df = ocr_df[(ocr_df.conf != -1) & (ocr_df.text.notna())].copy()
    ocr_df = ocr_df[ocr_df.text.str.strip() != ""]
    ocr_df = ocr_df.sort_values(by=["top", "left"])

    # Group lines by vertical position
    lines = []
    current_line = []
    last_top = None
    threshold = 15  # Adjust if lines are splitting incorrectly

    for _, row in ocr_df.iterrows():
        if last_top is None or abs(row['top'] - last_top) <= threshold:
            current_line.append(row)
        else:
            lines.append(current_line)
            current_line = [row]
        last_top = row['top']
    if current_line:
        lines.append(current_line)

    # Convert lines to row-wise text
    rows = []
    for line in lines:
        sorted_line = sorted(line, key=lambda x: x['left'])
        row_texts = [word['text'] for word in sorted_line]
        rows.append(row_texts)

    # Write to Excel
    ws = wb.create_sheet(title=f"Page_{i+1}")
    for row in rows:
        ws.append(row)

    print(f"📊 Sheet created: Page_{i+1}")

# Save Excel
wb.save(excel_output_file)
print(f"\n✅ DONE: Excel saved at {excel_output_file}")
