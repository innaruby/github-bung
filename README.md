from docx import Document
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

doc = Document("output/searchable.docx")
wb = Workbook()
wb.remove(wb.active)

page_count = 1
page_text = []
for para in doc.paragraphs:
    text = para.text.strip()
    if text:
        page_text.append(text)
    if para.text.strip() == "" and len(page_text) > 20:  # crude page break signal
        df = pd.DataFrame([[line] for line in page_text])
        ws = wb.create_sheet(title=f"Page_{page_count}")
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)
        page_count += 1
        page_text = []

# Add last page
if page_text:
    df = pd.DataFrame([[line] for line in page_text])
    ws = wb.create_sheet(title=f"Page_{page_count}")
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

wb.save("final_output.xlsx")
print("✅ Done! Word to Excel conversion complete.")
