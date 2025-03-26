import pdfplumber
import pytesseract
import cv2
import pandas as pd
import os
from pdf2image import convert_from_path
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# === CONFIG ===
input_pdf = r"C:/MyPDFs/searchable.pdf"  # use ocrmypdf if needed
fallback_images_dir = r"C:/MyPDFs/pdf_images"
output_excel = r"C:/MyPDFs/final_output.xlsx"

os.makedirs(fallback_images_dir, exist_ok=True)
wb = Workbook()
wb.remove(wb.active)

def extract_with_plumber(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            print(f"🔍 Processing Page {i+1} with PDFPlumber...")
            tables = page.extract_tables()
            if not tables:
                print("❗ No tables found, using OCR fallback.")
                yield i, None
            else:
                max_table = max(tables, key=lambda t: len(t))  # Pick largest table
                df = pd.DataFrame(max_table)
                yield i, df

def extract_with_tesseract_fallback(pdf_path, page_number):
    images = convert_from_path(pdf_path, dpi=300, first_page=page_number+1, last_page=page_number+1)
    image_path = os.path.join(fallback_images_dir, f"page_{page_number+1}.png")
    images[0].save(image_path, "PNG")

    img = cv2.imread(image_path)
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)
    data = data[(data.conf != -1) & (data.text.notna())]
    data = data[['left', 'top', 'text']]

    if data.empty:
        return pd.DataFrame([["No text found"]])
    
    # --- Cluster similar to earlier method ---
    def cluster_lines_into_table(data, col_threshold=50, row_threshold=10):
        data = data.sort_values(by=['top', 'left'])
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
        all_lefts = sorted(set(int(item['left']) for line in lines for item in line))
        col_bins = [all_lefts[0]]
        for l in all_lefts[1:]:
            if l - col_bins[-1] > col_threshold:
                col_bins.append(l)
        structured_rows = []
        for line in lines:
            row_dict = {}
            for word in line:
                col_idx = min(range(len(col_bins)), key=lambda i: abs(word['left'] - col_bins[i]))
                row_dict[col_idx] = word['text']
            row = [row_dict.get(i, "") for i in range(len(col_bins))]
            structured_rows.append(row)
        return pd.DataFrame(structured_rows)

    return cluster_lines_into_table(data)

# === Run Extraction ===
for i, df in extract_with_plumber(input_pdf):
    if df is None:
        df = extract_with_tesseract_fallback(input_pdf, i)

    sheet_name = f"Page_{i+1}"
    ws = wb.create_sheet(title=sheet_name)
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)
    print(f"✅ Sheet added: {sheet_name}")

# === Save Excel ===
wb.save(output_excel)
print(f"\n🎉 All done! Excel saved to: {output_excel}")
