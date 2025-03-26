import os
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"

# Optional: Tesseract path (for Windows)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

os.makedirs(image_output_dir, exist_ok=True)

# === STEP 1: Convert PDF to images ===
print("📄 Converting PDF to images...")
pages = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(pages):
    image_path = os.path.join(image_output_dir, f"page_{i+1}.png")
    page.save(image_path, "PNG")
    image_paths.append(image_path)
    print(f"✅ Saved: {image_path}")

# === Utility: layout-aware clustering ===
def extract_aligned_table(data, y_threshold=10, x_tolerance=25):
    data = data.sort_values(by=["top", "left"])
    rows = []
    current_row = []
    last_top = None

    # Step 1: Group into lines
    for _, word in data.iterrows():
        if last_top is None or abs(word["top"] - last_top) <= y_threshold:
            current_row.append(word)
        else:
            rows.append(current_row)
            current_row = [word]
        last_top = word["top"]
    if current_row:
        rows.append(current_row)

    # Step 2: Collect all column positions
    all_lefts = sorted(set(word["left"] for row in rows for word in row))

    # Step 3: Cluster `left` positions into columns
    column_bins = []
    for pos in all_lefts:
        for b in column_bins:
            if abs(pos - b) <= x_tolerance:
                break
        else:
            column_bins.append(pos)
    column_bins.sort()

    # Step 4: Align words to nearest column
    structured_rows = []
    for row in rows:
        aligned = [""] * len(column_bins)
        for word in row:
            for i, col_pos in enumerate(column_bins):
                if abs(word["left"] - col_pos) <= x_tolerance:
                    aligned[i] += (" " + word["text"]).strip()
                    break
        structured_rows.append(aligned)

    return structured_rows

# === STEP 2: OCR and layout → Excel ===
wb = Workbook()
wb.remove(wb.active)

for i, image_path in enumerate(image_paths):
    print(f"🔍 OCR: {image_path}")
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    ocr_data = pytesseract.image_to_data(
        gray, config="--oem 1 --psm 6", output_type=pytesseract.Output.DATAFRAME
    )
    ocr_data = ocr_data[(ocr_data.conf != -1) & (ocr_data.text.notna())]

    if ocr_data.empty:
        print(f"⚠️ No text found on page {i+1}")
        continue

    structured = extract_aligned_table(ocr_data)
    df = pd.DataFrame(structured)
    df = df.replace('', pd.NA).dropna(how='all', axis=1)

    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    print(f"📄 Sheet created: {sheet_name}")

# === Save Excel ===
wb.save(excel_output_file)
print(f"\n✅ Done! Excel saved at: {excel_output_file}")
