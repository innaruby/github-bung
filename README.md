import pdfplumber
import pandas as pd
from collections import defaultdict

pdf_path = "your_file.pdf"
excel_path = "output_layout_preserved.xlsx"

def group_words_by_line(words, y_tolerance=3):
    lines = defaultdict(list)
    for word in words:
        y0 = round(word['top'] / y_tolerance) * y_tolerance
        lines[y0].append(word)
    return [sorted(line, key=lambda w: w['x0']) for y0, line in sorted(lines.items())]

with pdfplumber.open(pdf_path) as pdf:
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')

    for i, page in enumerate(pdf.pages):
        sheet_name = f"Page_{i+1}"
        tables = page.extract_tables()

        if tables:
            # Prefer actual table extraction if available
            full_page_data = []
            for table in tables:
                full_page_data.extend(table)
                full_page_data.append([""])  # gap between tables
            df = pd.DataFrame(full_page_data)
        else:
            words = page.extract_words()
            if words:
                grouped_lines = group_words_by_line(words)
                rows = []
                for line in grouped_lines:
                    row = [w['text'] for w in line]
                    rows.append(row)
                df = pd.DataFrame(rows)
            else:
                df = pd.DataFrame([["No content found on this page"]])

        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    writer.close()

print(f"✅ Excel saved with visual layout preserved: {excel_path}")
