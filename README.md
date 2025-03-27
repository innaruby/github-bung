from docx import Document
from openpyxl import Workbook

word_path = r"C:/MyPDFs/your_file.docx"
excel_output = r"C:/MyPDFs/final_output.xlsx"

# Load Word document
doc = Document(word_path)

# Excel setup
wb = Workbook()
wb.remove(wb.active)

# === Split Word content by page breaks ===
pages = []
current_page = []

for elem in doc.element.body:
    if elem.tag.endswith("}p"):
        p = elem
        if 'lastRenderedPageBreak' in str(p.xml):
            pages.append(current_page)
            current_page = []
        else:
            current_page.append(p)
if current_page:
    pages.append(current_page)

# === Process Each Page ===
for i, page in enumerate(pages):
    sheet = wb.create_sheet(title=f"Page_{i+1}")
    row_idx = 1
    for p_elem in page:
        para = p_elem.xpath('.//w:t', namespaces=p_elem.nsmap)
        text = ''.join([t.text for t in para if t.text])
        if text.strip():
            words = text.split()
            for col_idx, word in enumerate(words, start=1):
                sheet.cell(row=row_idx, column=col_idx, value=word)
            row_idx += 1

# Save Excel
wb.save(excel_output)
print(f"✅ Excel saved at: {excel_output}")
