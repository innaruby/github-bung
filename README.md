import os
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Optional: Set Tesseract path (for Windows)
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

# === Utility: Group words into rows based on Y-position ===
def cluster_to_rows(data, y_threshold=10):
    data = data.sort_values(by='top')
    rows = []
    current_row = []
    last_top = None

    for _, row in data.iterrows():
        if last_top is None or abs(row['top'] - last_top) <= y_threshold:
            current_row.append(row)
        else:
            rows.append(current_row)
            current_row = [row]
        last_top = row['top']

    if current_row:
        rows.append(current_row)
    return rows

# === STEP 2: Perform OCR and write to Excel ===
wb = Workbook()
wb.remove(wb.active)  # Remove default sheet

for i, image_path in enumerate(image_paths):
    print(f"🔍 Processing OCR for: {image_path}")
    img = cv2.imread(image_path)

    # Run OCR
    config = r'--oem 1 --psm 6'  # Assumes block of text
    data = pytesseract.image_to_data(img, config=config, output_type=pytesseract.Output.DATAFRAME)

    # Clean up data
    data = data[(data.conf != -1) & (data.text.notna())]

    if data.empty:
        print(f"⚠️ No text detected on {image_path}")
        continue

    # Cluster words into rows
    clustered_rows = cluster_to_rows(data, y_threshold=10)

    # Sort each row by horizontal position (left)
    structured_rows = []
    for row in clustered_rows:
        sorted_row = sorted(row, key=lambda r: r['left'])
        row_text = [r['text'] for r in sorted_row]
        structured_rows.append(row_text)

    df = pd.DataFrame(structured_rows)

    # Write to Excel
    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    print(f"📄 Sheet created: {sheet_name}")

# Save workbook
wb.save(excel_output_file)
print(f"\n✅ Done! Excel saved at: {excel_output_file}")
