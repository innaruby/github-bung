import os
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Optional: Set tesseract path (Windows only)
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

# === Helper function: Cluster OCR words into table-like layout ===
def cluster_lines_into_table(data, col_threshold=50, row_threshold=10):
    data = data.sort_values(by=['top', 'left'])

    # Group words into lines based on vertical positions
    lines = []
    current_line = []
    prev_top = -1000

    for _, row in data.iterrows():
        if abs(row['top'] - prev_top) > row_threshold:
            if current_line:
                lines.append(current_line)
            current_line = [row]
            prev_top = row['top']
        else:
            current_line.append(row)
    if current_line:
        lines.append(current_line)

    # Identify potential column positions
    all_lefts = sorted(set(int(item['left']) for line in lines for item in line))
    all_lefts = np.array(all_lefts)
    col_bins = [all_lefts[0]]
    for l in all_lefts[1:]:
        if l - col_bins[-1] > col_threshold:
            col_bins.append(l)

    # Map words into appropriate column bins
    structured_rows = []
    for line in lines:
        row_dict = {}
        for word in line:
            col_idx = np.argmin([abs(word['left'] - c) for c in col_bins])
            row_dict[col_idx] = word['text']
        row = [row_dict.get(i, "") for i in range(len(col_bins))]
        structured_rows.append(row)

    return pd.DataFrame(structured_rows)

# === STEP 2: OCR with layout detection and write to Excel ===
wb = Workbook()
wb.remove(wb.active)

for i, image_path in enumerate(image_paths):
    print(f"🔍 Processing OCR for: {image_path}")
    img = cv2.imread(image_path)

    # Run Tesseract OCR with layout info
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)

    # Clean data
    data = data[(data.conf != -1) & (data.text.notna())]
    data = data[['left', 'top', 'text']]  # Keep only required columns

    if not data.empty:
        df = cluster_lines_into_table(data)
    else:
        df = pd.DataFrame([["No text found"]])

    # Create Excel sheet for this page
    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    print(f"📄 Sheet created: {sheet_name}")

# Save Excel workbook
wb.save(excel_output_file)
print(f"\n✅ Done! Excel saved at: {excel_output_file}")
