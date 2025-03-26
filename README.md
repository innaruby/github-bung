import os
from pdf2image import convert_from_path
from paddleocr import PaddleOCR
import cv2
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"
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

# === STEP 2: Load PaddleOCR (with layout detection) ===
ocr = PaddleOCR(use_angle_cls=True, lang='en', show_log=False)

# === STEP 3: Excel Init ===
wb = Workbook()
wb.remove(wb.active)

# === STEP 4: Process Each Page ===
for i, image_path in enumerate(image_paths):
    print(f"🔍 Processing structured OCR on: {image_path}")
    result = ocr.ocr(image_path, cls=True)

    # Flatten OCR results
    data = []
    for line in result[0]:
        if line:
            box = line[0]
            text = line[1][0]
            confidence = line[1][1]
            (x_min, y_min) = map(int, box[0])
            data.append((y_min, x_min, text, confidence))

    # Sort top to bottom, then left to right
    data = sorted(data, key=lambda x: (x[0], x[1]))

    # Group rows based on Y-axis proximity
    rows = []
    current_row = []
    prev_y = None
    row_threshold = 15

    for item in data:
        y, x, text, conf = item
        if prev_y is None or abs(y - prev_y) < row_threshold:
            current_row.append((x, text))
        else:
            rows.append(sorted(current_row, key=lambda r: r[0]))
            current_row = [(x, text)]
        prev_y = y
    if current_row:
        rows.append(sorted(current_row, key=lambda r: r[0]))

    # === Write to Excel ===
    ws = wb.create_sheet(title=f"Page_{i+1}")
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, (_, text) in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=text)

    print(f"📄 Structured page written: Page_{i+1}")

# === STEP 5: Save Excel ===
wb.save(excel_output_file)
print(f"\n✅ All pages processed. Excel saved at: {excel_output_file}")
