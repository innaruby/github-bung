from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import openpyxl

def apply_light_grey_fill_final(directory):
    print("\n🎨 Applying light grey fill in Summe rows and PLAN columns, clearing other backgrounds...")
    current_year = datetime.now().year
    light_grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue

        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            visible_cols = [col for col in range(1, ws.max_column + 1)
                            if not ws.column_dimensions[get_column_letter(col)].hidden]
            end_row = find_end_row(ws, sheet_name)

            summe_rows = set()
            plan_col_target = None

            for col in visible_cols:
                header3 = str(ws.cell(row=3, column=col).value or "").strip().upper()
                header4 = str(ws.cell(row=4, column=col).value or "").replace("e", "").strip()
                if header3 == "PLAN" and header4 == str(current_year + 1):
                    plan_col_target = col

            for row in range(5, end_row + 1):
                if ws.row_dimensions[row].hidden:
                    continue
                cell_val = str(ws.cell(row=row, column=1).value or "").strip().lower()
                if "summe" in cell_val and ws.cell(row=row, column=1).font.bold:
                    summe_rows.add(row)

            for row in range(5, end_row + 1):
                if ws.row_dimensions[row].hidden:
                    continue
                for col in visible_cols:
                    cell = ws.cell(row=row, column=col)
                    if row in summe_rows or (plan_col_target == col):
                        cell.fill = light_grey_fill
                    else:
                        cell.fill = white_fill

        wb.save(file_path)
        print(f"✅ Updated fill colors in: {file}")
