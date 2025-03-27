import os
import subprocess
from docx import Document
from openpyxl import Workbook
from lxml import etree

# === CONFIGURATION ===
pdf_path = r"C:/MyPDFs/your_file.pdf"
output_dir = r"C:/MyPDFs/"
docx_path = os.path.join(output_dir, "converted.docx")
excel_path = os.path.join(output_dir, "final_output.xlsx")

# === STEP 1: Convert PDF to DOCX using LibreOffice ===
print("📄 Converting PDF to Word...")
subprocess.run([
    "soffice",
    "--headless",
    "--convert-to", "docx",
    pdf_path,
    "--outdir", output_dir
], check=True)

if not os.path.exists(docx_path):
    print(f"❌ Word file not created: {docx_path}")
    exit(1)

print(f"✅ Word file created: {docx_path}")

# === STEP 2: Load Word file and split by page breaks ===
print("📝 Reading Word file and detecting pages...")
doc = Document(docx_path)
wb = Workbook()
wb.remove(wb.active)

# Helper: split Word into pages using XML <w:lastRenderedPageBreak>
pages = []
current_page = []

for element in doc.element.body:
    if element.tag.endswith("}p") and "lastRenderedPageBreak" in str(element.xml):
        pages.append(current_page)
        current_page = []
    else:
        current_page.append(element)
if current_page:
    pages.append(current_page)

print(f"📄 Total pages detected: {len(pages)}")

# === STEP 3: Write each page's content to Excel sheet ===
for i, page_elements in enumerate(pages):
    ws = wb.create_sheet(title=f"Page_{i+1}")
    row_idx = 1
    for elem in page_elements:
        text_elements = elem.xpath('.//w:t', namespaces=elem.nsmap)
        text = ''.join([t.text for t in text_elements if t.text])
        if text.strip():
            words = text.strip().split()
            for col_idx, word in enumerate(words, start=1):
                ws.cell(row=row_idx, column=col_idx, value=word)
            row_idx += 1

    print(f"📊 Page {i+1} copied to Excel")

# === STEP 4: Save Excel ===
wb.save(excel_path)
print(f"\n✅ Done! Excel saved at: {excel_path}")
