import os
import cv2
import pytesseract
from pdf2image import convert_from_path
from docx import Document
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

# === CONFIGURATION ===
pdf_path = "C:/MyPDFs/input.pdf"
output_folder = "C:/MyPDFs/"
word_path = os.path.join(output_folder, "output.docx")
excel_path = os.path.join(output_folder, "output.xlsx")
image_folder = os.path.join(output_folder, "images")

os.makedirs(image_folder, exist_ok=True)

# === STEP 1: Convert PDF to Images ===
pages = convert_from_path(pdf_path, dpi=300)
image_paths = []
for i, page in enumerate(pages):
    img_path = os.path.join(image_folder, f"page_{i+1}.png")
    page.save(img_path, "PNG")
    image_paths.append(img_path)

# === STEP 2: OCR and Create Word with One Page per PDF Page ===
doc = Document()

def ocr_image_to_lines(image_path):
    img = cv2.imread(image_path)
    ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)
    ocr_data = ocr_data[(ocr_data.conf != -1) & (ocr_data.text.notna())]
    ocr_data = ocr_data[['left', 'top', 'text']]
    ocr_data = ocr_data.sort_values(by=['top', 'left'])

    # Reconstruct lines by grouping close top coordinates
    lines = []
    current_line = []
    prev_top = -1000

    for _, row in ocr_data.iterrows():
        if abs(row['top'] - prev_top) > 10:
            if current_line:
                lines.append(current_line)
            current_line = [row['text']]
            prev_top = row['top']
        else:
            current_line.append(row['text'])

    if current_line:
        lines.append(current_line)

    return [" ".join(line) for line in lines]

all_page_texts = []

for img_path in image_paths:
    lines = ocr_image_to_lines(img_path)
    all_page_texts.append(lines)

    for line in lines:
        doc.add_paragraph(line)
    doc.add_page_break()

doc.save(word_path)
print(f"✅ Word document saved to: {word_path}")

# === STEP 3: Convert Each Word "Page" to a Sheet in Excel ===
wb = Workbook()
wb.remove(wb.active)

for i, lines in enumerate(all_page_texts):
    df = pd.DataFrame([[line] for line in lines])
    ws = wb.create_sheet(title=f"Page_{i+1}")
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

wb.save(excel_path)
print(f"✅ Excel document saved to: {excel_path}")
