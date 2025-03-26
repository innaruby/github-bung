import os
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# CONFIG
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"

# Optional: Set Tesseract path (Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

os.makedirs(image_output_dir, exist_ok=True)

# STEP 1: Convert PDF to images
print("📄 Converting PDF to images...")
pages = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(pages):
    image_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(image_path, "PNG")
    image_paths.append(image_path)
    print(f"✅ Saved: {image_path}")

# Utility to cluster and merge text
def cluster_to_rows_with_phrases(data, y_threshold=10, x_threshold=25):
    data = data.sort_values(by=['top', 'left'])
    rows, current_row, last_top = [], [], None

    for _, row in data.iterrows():
        if last_top is None or abs(row['top'] - last_top) <= y_threshold:
            current_row.append(row)
        else:
            rows.append(current_row)
            current_row = [row]
        last_top = row['top']
    if current_row:
        rows.append(current_row)

    structured_lines = []
    for row in rows:
        row = sorted(row, key=lambda r: r['left'])
        merged_line, last_right, current_phrase = [], None, ""

        for word in row:
            if last_right is None:
                current_phrase = word['text']
            elif word['left'] - last_right <= x_threshold:
                current_phrase += " " + word['text']
            else:
                merged_line.append(current_phrase)
                current_phrase = word['text']
            last_right = word['left'] + word['width']
        if current_phrase:
            merged_line.append(current_phrase)
        structured_lines.append(merged_line)
    return structured_lines

# STEP 2: OCR each image and write to Excel
wb = Workbook()
wb.remove(wb.active)

for i, image_path in enumerate(image_paths):
    print(f"🔍 OCR processing: {image_path}")
    img = cv2.imread(image_path)

    # Optional: Preprocess image (e.g. grayscale, threshold) for better OCR
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    data = pytesseract.image_to_data(gray, config='--oem 1 --psm 6', output_type=pytesseract.Output.DATAFRAME)
    data = data[(data.conf != -1) & (data.text.notna())]

    if data.empty:
        print(f"⚠️ No text found on page {i+1}")
        continue

    structured_lines = cluster_to_rows_with_phrases(data)
    df = pd.DataFrame(structured_lines)
    df = df.replace('', pd.NA).dropna(how='all', axis=1)

    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    print(f"📄 Sheet created: {sheet_name}")

# Save Excel file
wb.save(excel_output_file)
print(f"\n✅ Done! Excel saved at: {excel_output_file}")
