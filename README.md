import os
import cv2
import pandas as pd
import numpy as np
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_dir = r"C:/MyPDFs/pdf_images"
output_excel = r"C:/MyPDFs/structured_output.xlsx"
os.makedirs(image_dir, exist_ok=True)

# === STEP 1: Convert PDF to images ===
print("📄 Converting PDF pages to images...")
images = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(images):
    image_path = os.path.join(image_dir, f"page_{i+1}.png")
    page.save(image_path, "PNG")
    image_paths.append(image_path)
    print(f"✅ Saved: {image_path}")

# === STEP 2: Function to layout-align OCR results ===
def reconstruct_layout_from_ocr(ocr_data, row_thresh=10, col_thresh=50):
    ocr_data = ocr_data[(ocr_data.conf != -1) & (ocr_data.text.notna())]
    if ocr_data.empty:
        return pd.DataFrame([["No text found"]])

    ocr_data = ocr_data[['left', 'top', 'text']]
    ocr_data = ocr_data.sort_values(by=['top', 'left'])

    # === Group by visual rows ===
    lines = []
    current_line = []
    prev_top = -1000
    for _, row in ocr_data.iterrows():
        if abs(row['top'] - prev_top) > row_thresh:
            if current_line:
                lines.append(current_line)
            current_line = [row]
            prev_top = row['top']
        else:
            current_line.append(row)
    if current_line:
        lines.append(current_line)

    # === Estimate column anchors ===
    all_lefts = sorted(set(int(item['left']) for line in lines for item in line))
    all_lefts = np.array(all_lefts)
    col_bins = [all_lefts[0]]
    for l in all_lefts[1:]:
        if l - col_bins[-1] > col_thresh:
            col_bins.append(l)

    # === Build structured rows ===
    structured_rows = []
    for line in lines:
        row_dict = {}
        for word in line:
            col_idx = np.argmin([abs(word['left'] - c) for c in col_bins])
            row_dict[col_idx] = word['text']
        row = [row_dict.get(i, "") for i in range(len(col_bins))]
        structured_rows.append(row)

    return pd.DataFrame(structured_rows)

# === STEP 3: Process each image ===
wb = Workbook()
wb.remove(wb.active)

for i, img_path in enumerate(image_paths):
    print(f"🔍 OCR Processing Page {i+1}")
    img = cv2.imread(img_path)
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)

    df = reconstruct_layout_from_ocr(data)
    ws = wb.create_sheet(title=f"Page_{i+1}")
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)
    print(f"✅ Page {i+1} structured and added to Excel")

# === STEP 4: Save Excel ===
wb.save(output_excel)
print(f"\n🎉 All done! Excel saved to: {output_excel}")
