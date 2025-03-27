import os
from pdf2docx import Converter
import PyPDF2
import docx
from openpyxl import Workbook

# === CONFIGURATION ===
# Path to your PDF file
pdf_file = r"C:\MyPDFs\your_file.pdf"
# Temporary folder to store per‑page DOCX files
temp_docx_dir = r"C:\MyPDFs\temp_docx"
# Final Excel output file
excel_output_file = r"C:\MyPDFs\final_output.xlsx"

# Create temporary directory if it does not exist
os.makedirs(temp_docx_dir, exist_ok=True)

# === STEP 1: DETERMINE PDF PAGE COUNT ===
with open(pdf_file, 'rb') as f:
    reader = PyPDF2.PdfFileReader(f)
    total_pages = reader.numPages

print(f"[INFO] Total pages in PDF: {total_pages}")

# === STEP 2: CONVERT EACH PDF PAGE TO A SEPARATE DOCX FILE ===
docx_files = []
for i in range(total_pages):
    output_docx = os.path.join(temp_docx_dir, f"page_{i+1}.docx")
    print(f"[INFO] Converting page {i+1} to DOCX: {output_docx}...")
    # Create Converter instance for the PDF file
    cv = Converter(pdf_file)
    # Convert only page i (pages are zero-indexed)
    cv.convert(output_docx, start=i, end=i)
    cv.close()
    docx_files.append(output_docx)
print("[SUCCESS] PDF-to-DOCX conversion complete.")

# === STEP 3: CREATE EXCEL FILE WITH EACH DOCX CONTROLLING A SEPARATE SHEET ===
wb = Workbook()
# Remove the default created sheet
default_sheet = wb.active
wb.remove(default_sheet)

for idx, docx_file in enumerate(docx_files, start=1):
    print(f"[INFO] Processing DOCX: {docx_file} ...")
    document = docx.Document(docx_file)
    
    # Create a new worksheet for the current page; name it accordingly.
    sheet_name = f"Page_{idx}"
    ws = wb.create_sheet(title=sheet_name)
    
    # Start putting content from cell A1.
    row = 1
    for para in document.paragraphs:
        text = para.text.strip()
        if text:  # Only write non-empty lines
            ws.cell(row=row, column=1, value=text)
            row += 1
    print(f"[SUCCESS] Content from {docx_file} copied to Excel sheet '{sheet_name}'.")

# Save the Excel workbook
wb.save(excel_output_file)
print(f"[FINAL] Excel workbook saved at: {excel_output_file}")
