import os
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from sklearn.cluster import KMeans
import numpy as np

# Optional: Set Tesseract path (Windows only)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"
dpi = 300
max_columns = 10   # Max columns to guess for table structure

# === STEP 1: Convert PDF pages to images ===
os.makedirs(image_output_dir, exist_ok=True)
print("📄 Converting PDF to images...")
pages = convert_from_path(pdf_path, dpi=dpi)

image_paths = []
for i, page in enumerate(pages):
    image_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(image_path, "PNG")
    image_paths.append(image_path)
    print(f"✅ Saved: {image_path}")

# === STEP 2: Define helper function for column clustering ===
def cluster_columns(words_df, max_columns=6):
    if len(words_df) < max_columns:
        max_columns = len(words_df)
    x_coords = words_df['left'].values.reshape(-1, 1)
    kmeans = KMeans(n_clusters=max_columns, n_init='auto', random_state=0).fit(x_coords)
    words_df['col_id'] = kmeans.labels_
    words_df = words_df.sort_values(by=['top', 'col_id'])
    return words_df

# === STEP 3: Run OCR and build Excel ===
wb = Workbook()
wb.remove(wb.active)

for i, image_path in enumerate(image_paths):
    print(f"🔍 OCR on: {image_path}")
    img = cv2.imread(image_path)

    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)
    data = data[(data.conf != -1) & (data.text.notna())]
    data = data.sort_values(by=['top'])

    # Group words by similar 'top' values into rows
    rows = []
    current_row = []
    prev_top = None
    line_threshold = 10  # Tweak this value if lines are merging/splitting incorrectly

    for _, row in data.iterrows():
        if prev_top is None or abs(row['top'] - prev_top) < line_threshold:
            current_row.append(row)
        else:
            rows.append(pd.DataFrame(current_row))
            current_row = [row]
        prev_top = row['top']
    if current_row:
        rows.append(pd.DataFrame(current_row))

    # Structure rows by clustering into columns
    structured_data = []
    for row_df in rows:
        clustered = cluster_columns(row_df, max_columns)
        line_text = clustered.sort_values('col_id')['text'].tolist()
        structured_data.append(line_text)

    # Normalize row lengths
    max_len = max(len(r) for r in structured_data)
    for r in structured_data:
        r += [''] * (max_len - len(r))

    df = pd.DataFrame(structured_data)

    # Write to Excel sheet
    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)
    print(f"📄 Sheet created: {sheet_name}")

# Save Excel workbook
wb.save(excel_output_file)
print(f"\n✅ Done! Excel saved at: {excel_output_file}")
