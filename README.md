import os
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# -------- CONFIG --------
pdf_path = "your_file.pdf"              # Your input PDF file
final_excel_path = "final_output.xlsx"  # Final Excel output file

# -------- STEP 0: Count Pages Using pdfplumber --------
def get_pdf_page_count(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        return len(pdf.pages)

# -------- STEP 1: PDF → Excel (Preserve Layout) --------
def convert_pdf_to_excel(pdf_path, excel_path):
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                print(f"⚠️ No text found on page {i+1}. Skipping.")
                continue

            # Split text into lines and create a DataFrame
            lines = text.split('\n')
            df = pd.DataFrame(lines, columns=["Content"])

            # Create a new sheet for each page
            sheet_name = f"Page_{i+1}"
            ws = wb.create_sheet(title=sheet_name)

            # Write DataFrame to Excel
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
        convert_pdf_to_excel(pdf_path, final_excel_path)
