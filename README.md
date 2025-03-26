import os
from pdf2image import convert_from_path
import pytesseract
import cv2
import pandas as pd
import layoutparser as lp
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# === CONFIG ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
image_output_dir = r"C:/MyPDFs/pdf_pages"
excel_output_file = r"C:/MyPDFs/final_output.xlsx"

# Optional: Set Tesseract path (Windows users)
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

# === Layout Parser Model ===
model = lp.Detectron2LayoutModel(
    "lp://PubLayNet/faster_rcnn_R_50_FPN_3x/config",
    extra_config=["MODEL.ROI_HEADS.SCORE_THRESH_TEST", 0.5],
    label_map={0: "Text", 1: "Title", 2: "List", 3: "Table", 4: "Figure"}
)

# === Helper: Merge nearby words into phrases based on X/Y proximity ===
def cluster_to_rows_with_phrases(data, y_threshold=10, x_threshold=25):
    data = data.sort_values(by=['top', 'left'])
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

    structured_lines = []
    for row in rows:
        row = sorted(row, key=lambda r: r['left'])
        merged_line = []
        last_right = None
        current_phrase = ""

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

# === STEP 2: Layout detection, OCR, and Excel export ===
wb = Workbook()
wb.remove(wb.active)

for i, image_path in enumerate(image_paths):
    print(f"🔍 Processing: {image_path}")
    img = cv2.imread(image_path)

    layout = model.detect(img)
    layout = lp.Layout([b for b in layout if b.type in ['Text', 'Title', 'Table']])

    page_rows = []

    for block in layout:
        segment = block.crop_image(img)
        ocr_data = pytesseract.image_to_data(segment, config='--oem 1 --psm 6', output_type=pytesseract.Output.DATAFRAME)
        ocr_data = ocr_data[(ocr_data.conf != -1) & (ocr_data.text.notna())]

        if not ocr_data.empty:
            lines = cluster_to_rows_with_phrases(ocr_data)
            page_rows.extend(lines)

    if not page_rows:
        print(f"⚠️ No text found in: {image_path}")
        continue

    df = pd.DataFrame(page_rows)
    df = df.replace('', pd.NA).dropna(how='all', axis=1)

    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    print(f"📄 Sheet created: {sheet_name}")

# === Save Excel ===
wb.save(excel_output_file)
print(f"\n✅ Done! Excel saved at: {excel_output_file}")
