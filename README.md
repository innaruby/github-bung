from pdf2docx import Converter
from docx import Document
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os

# -------- Step 1: Convert PDF to Word --------
def convert_pdf_to_docx(pdf_path, docx_path):
    print("🔄 Converting PDF to Word...")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)  # Convert all pages
    cv.close()
    print(f"✅ PDF converted to DOCX: {docx_path}")

# -------- Step 2: Convert Word to Excel --------
def convert_docx_to_excel(docx_path, excel_path):
    print("📤 Extracting from Word and writing to Excel...")
    doc = Document(docx_path)
    wb = Workbook()
    wb.remove(wb.active)  # Remove default empty sheet

    page_num = 1
    buffer_lines = []

    def flush_page(lines, page_num):
        if lines:
            df = pd.DataFrame(lines)
            ws = wb.create_sheet(title=f"Page_{page_num}")
            for row in dataframe_to_rows(df, index=False, header=False):
                ws.append(row)

    for para in doc.paragraphs:
        if para.text.strip() == "":
            flush_page(buffer_lines, page_num)
            buffer_lines = []
            page_num += 1
        else:
            buffer_lines.append([para.text])

    # Handle last page
    flush_page(buffer_lines, page_num)

    # Optionally extract tables as separate sheets
    for idx, table in enumerate(doc.tables):
        data = []
        for row in table.rows:
            data.append([cell.text.strip() for cell in row.cells])
        df = pd.DataFrame(data)
        ws = wb.create_sheet(title=f"Table_{idx+1}")
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)

    wb.save(excel_path)
    print(f"✅ Excel saved as: {excel_path}")

# -------- Run the pipeline --------
if __name__ == "__main__":
    pdf_file = "your_file.pdf"
    docx_file = "converted.docx"
    excel_file = "final_output.xlsx"

    if not os.path.exists(pdf_file):
        print(f"❌ PDF file not found: {pdf_file}")
    else:
        convert_pdf_to_docx(pdf_file, docx_file)
        convert_docx_to_excel(docx_file, excel_file)
