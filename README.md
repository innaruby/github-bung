import os
from pdf2docx import Converter
from docx import Document
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import fitz  # PyMuPDF

# -------- CONFIG --------
pdf_path = "your_file.pdf"              # Your input PDF file
docx_output_dir = "docx_pages"          # Temporary folder to hold per-page Word files
final_excel_path = "final_output.xlsx"  # Final Excel output file

# -------- STEP 0: Count Pages Using PyMuPDF --------
def get_pdf_page_count(pdf_file):
    doc = fitz.open(pdf_file)
    return doc.page_count

# -------- STEP 1: PDF → Word (one DOCX per page) --------
def split_pdf_to_docx_per_page(pdf_path, output_dir, num_pages):
    os.makedirs(output_dir, exist_ok=True)
    cv = Converter(pdf_path)

    print(f"🔄 Converting PDF to {num_pages} Word pages...")
    for i in range(num_pages):
        page_docx = os.path.join(output_dir, f"page_{i+1}.docx")
        cv.convert(page_docx, start=i, end=i)
        print(f"✅ Saved: {page_docx}")
    cv.close()
    return num_pages

# -------- STEP 2: Word Pages → Excel Sheets --------
def convert_docx_pages_to_excel(docx_dir, excel_path, num_pages):
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    for i in range(1, num_pages + 1):
        filename = f"page_{i}.docx"
        docx_file = os.path.join(docx_dir, filename)
        if not os.path.exists(docx_file):
            print(f"⚠️ Skipping missing file: {filename}")
            continue

        doc = Document(docx_file)
        sheet_name = f"Page_{i}"
        rows = []

        # Extract paragraphs
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                rows.append([text])

        # Extract tables
        for table in doc.tables:
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                rows.append(row_data)
            rows.append([""])  # Add spacing after each table

        # Write to Excel
        df = pd.DataFrame(rows)
        ws = wb.create_sheet(title=sheet_name)
        for r in dataframe_to_rows(df, index=False, header=False):
            ws.append(r)

        print(f"📄 Added sheet: {sheet_name}")

    wb.save(excel_path)
    print(f"✅ Excel saved at: {excel_path}")

# -------- RUN THE PIPELINE --------
if __name__ == "__main__":
    if not os.path.exists(pdf_path):
        print(f"❌ PDF not found: {pdf_path}")
    else:
        page_count = get_pdf_page_count(pdf_path)
        split_pdf_to_docx_per_page(pdf_path, docx_output_dir, page_count)
        convert_docx_pages_to_excel(docx_output_dir, final_excel_path, page_count)
