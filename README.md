import os
import cv2
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"

os.makedirs(image_output_dir, exist_ok=True)

# === Convert PDF to images ===
print("📄 Converting PDF to images...")
pages = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(pages):
    img_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(img_path, "PNG")
    image_paths.append(img_path)
    print(f"✅ Saved: {img_path}")

# === Excel init ===
wb = Workbook()
wb.remove(wb.active)

# === Smart Text Structuring Function ===
def extract_structured_text(image_path):
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Get OCR output with bounding boxes
    ocr_df = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DATAFRAME)
    ocr_df = ocr_df[(ocr_df.conf != -1) & (ocr_df.text.notna())].copy()
    ocr_df = ocr_df[ocr_df.text.str.strip() != ""]

    # Sort by top -> left
    ocr_df = ocr_df.sort_values(by=['top', 'left'])

    # === Group lines using vertical (top) proximity ===
    rows = []
    current_row = []
    line_threshold = 15
    prev_top = None

    for _, row in ocr_df.iterrows():
        if prev_top is None or abs(row['top'] - prev_top) < line_threshold:
            current_row.append(row)
        else:
            rows.append(current_row)
            current_row = [row]
        prev_top = row['top']
    if current_row:
        rows.append(current_row)

    # === Convert to structured rows of text, aligned by left coordinate ===
    structured_data = []
    for row in rows:
        row_df = pd.DataFrame(row)
        row_df = pd.DataFrame(row) if not isinstance(row, pd.DataFrame) else row
        line = row_df.sort_values(by='left')['text'].tolist()
        structured_data.append(line)

    return pd.DataFrame(structured_data)

# === Process Each Page and Save to Excel ===
for i, image_path in enumerate(image_paths):
    print(f"🔍 Structuring data in: {image_path}")
    df = extract_structured_text(image_path)

    # Create Excel Sheet
    ws = wb.create_sheet(title=f"Page_{i+1}")
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    print(f"📄 Structured sheet saved: Page_{i+1}")

# Save Excel
wb.save(excel_output_file)
print(f"\n✅ Done! Excel with smart structure saved at: {excel_output_file}")
