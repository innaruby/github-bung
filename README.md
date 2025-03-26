import os
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Optional: Set Tesseract path (Windows only)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"

# === STEP 1: Convert PDF pages to images ===
os.makedirs(image_output_dir, exist_ok=True)
print("📄 Converting PDF to images...")
pages = convert_from_path(pdf_path, dpi=300)

image_paths = []
for i, page in enumerate(pages):
    image_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(image_path, "PNG")
    image_paths.append(image_path)
    print(f"✅ Saved: {image_path}")

# === STEP 2: Run OCR with layout detection and write to Excel ===
wb = Workbook()
wb.remove(wb.active)

for i, image_path in enumerate(image_paths):
    print(f"🔍 Processing OCR for: {image_path}")
    img = cv2.imread(image_path)

    # Run Tesseract OCR with layout info
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)

    # Clean data
    data = data[(data.conf != -1) & (data.text.notna())]

    # Group by block/paragraph/line
    grouped = data.groupby(['block_num', 'par_num', 'line_num'])

    structured_rows = []
    for _, line in grouped:
        line_words = line.sort_values('left')['text'].tolist()
        structured_rows.append(line_words)

    df = pd.DataFrame(structured_rows)

    # Create Excel sheet for this page
    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    print(f"📄 Sheet created: {sheet_name}")

# Save Excel workbook
wb.save(excel_output_file)
print(f"\n✅ Done! Excel saved at: {excel_output_file}")
