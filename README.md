import pdfplumber
import pandas as pd

# File paths
pdf_path = "your_file.pdf"
excel_path = "output.xlsx"

with pdfplumber.open(pdf_path) as pdf:
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')

    for i, page in enumerate(pdf.pages):
        sheet_name = f"Page_{i+1}"
        tables = page.extract_tables()

        if tables:
            # Combine all tables into one
            full_page_data = []
            for table in tables:
                full_page_data.extend(table)
                full_page_data.append([""])  # gap between tables

            df = pd.DataFrame(full_page_data)
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        else:
            # If no tables, extract plain text and write each line in a row
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                df = pd.DataFrame(lines)
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            else:
                # Create a blank sheet if even text is missing
                df = pd.DataFrame([["No content found on this page"]])
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    writer.close()

print(f"✅ Done! Saved as '{excel_path}' with one sheet per PDF page.")
