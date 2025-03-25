import pdfplumber
import pandas as pd

# Input and output files
pdf_path = "your_file.pdf"
excel_path = "output.xlsx"

with pdfplumber.open(pdf_path) as pdf:
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')

    for i, page in enumerate(pdf.pages):
        tables = page.extract_tables()

        if tables:
            # Combine all tables on the page into one
            full_page_data = []
            for table in tables:
                full_page_data.extend(table)
                full_page_data.append([""])  # Add a gap between tables for readability

            df = pd.DataFrame(full_page_data)
            sheet_name = f"Page_{i+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        else:
            print(f"⚠️ No table found on page {i+1}")

    writer.save()

print(f"✅ Done! Each page saved as a separate sheet in '{excel_path}'")
