import os
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from tkinter import Tk, filedialog

COLOR_MAP = {
    '#FFFFFF': 'White', '#FF0000': 'Red', '#00B050': 'Green', '#92D050': 'Light Green',
    '#0070C0': 'Blue', '#00B0F0': 'Light Blue', '#FFFF00': 'Yellow', '#FFC000': 'Orange',
    '#7030A0': 'Purple', '#D9D9D9': 'Gray', '#000000': 'Black', '#ED7D31': 'Dark Orange',
    '#A9D08E': 'Pale Green', '#F4B084': 'Peach', '#FFD966': 'Pale Yellow'
}

def rgb_to_hex_name(rgb):
    if rgb is None:
        return "No Color"
    if rgb.type == "rgb":
        hex_color = f"#{rgb.rgb[2:]}"
        return COLOR_MAP.get(hex_color.upper(), "Custom Color")
    elif rgb.type == "theme":
        return f"Theme Color {rgb.theme} (Tint {rgb.tint})"
    return "Unknown Format"

def get_sheet_tab_colors(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_colors = {}
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        color = sheet.sheet_properties.tabColor
        sheet_colors[sheet_name] = rgb_to_hex_name(color)
    return sheet_colors

def find_end_row(ws, sheet_name):
    sheet_name_lower = sheet_name.lower()
    for row in range(7, ws.max_row + 1):
        cell = ws[f"A{row}"]
        if cell.value and isinstance(cell.value, str) and cell.value.lower() == "summe" and cell.font.bold:
            return row
    for row in range(7, ws.max_row + 1):
        cell = ws[f"A{row}"]
        if cell.value and isinstance(cell.value, str) and cell.value.strip().lower() == sheet_name_lower:
            return row
    for row in range(7, ws.max_row + 1):
        if ws[f"A{row}"].value in (None, ""):
            return row - 1
    return ws.max_row

def find_merged_veraenderung_columns(ws):
    for row in [3, 4]:
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == row and merged_range.max_row == row:
                cell_value = ws.cell(row=row, column=merged_range.min_col).value
                if cell_value and "veränderung" in str(cell_value).lower():
                    return (merged_range.min_col, merged_range.max_col)
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value and "veränderung" in str(cell.value).lower():
                return (col, col)
    return None

def style_cell(cell):
    cell.font = Font(size=16, bold=True)
    cell.alignment = Alignment(horizontal="center")
    hair_border = Border(left=Side(style='hair'), right=Side(style='hair'),
                         top=Side(style='hair'), bottom=Side(style='hair'))
    cell.border = hair_border

def apply_veraenderung_formulas(ws, ist_col, plan_col, vera_start_col, end_row):
    diff_col = vera_start_col + 2
    perc_col = vera_start_col + 3

    for row in range(5, end_row + 1):
        plan_letter = get_column_letter(plan_col)
        ist_letter = get_column_letter(ist_col)
        diff_letter = get_column_letter(diff_col)

        ws.cell(row=row, column=diff_col).value = f"={plan_letter}{row}-{ist_letter}{row}"
        ws.cell(row=row, column=perc_col).value = f"=IF({ist_letter}{row}=0,0,({diff_letter}{row}/{ist_letter}{row})*100)"

def delete_columns_B_and_C(ws):
    ws.delete_cols(2, 2)

def process_excel_files(directory):
    current_year = datetime.now().year

    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue
        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)
        sheet_colors = get_sheet_tab_colors(file_path)

        for sheet_name in wb.sheetnames:
            tab_color = sheet_colors.get(sheet_name, "")
            if "green" not in tab_color.lower():
                continue

            ws = wb[sheet_name]
            delete_columns_B_and_C(ws)
            end_row = find_end_row(ws, sheet_name)
            vera_cols = find_merged_veraenderung_columns(ws)
            if vera_cols is None:
                continue

            vera_start_col, vera_end_col = vera_cols
            insert_col = vera_start_col

            merged_to_restore = []
            for merged_range in list(ws.merged_cells.ranges):
                if merged_range.min_row == 3 and merged_range.max_row == 3:
                    if merged_range.min_col == vera_start_col and merged_range.max_col == vera_end_col:
                        merged_to_restore.append(merged_range)
                        ws.unmerge_cells(str(merged_range))

            existing_plan = ws.cell(row=3, column=vera_start_col - 2).value
            existing_ist = ws.cell(row=3, column=vera_start_col - 1).value
            if str(existing_plan).strip().lower() == "plan" and str(existing_ist).strip().lower() == "ist":
                print(f"Skipping insertion in sheet '{sheet_name}' of file '{file}' as columns already exist.")
                continue

            ws.insert_cols(insert_col, 2)

            for merged_range in merged_to_restore:
                new_start = merged_range.min_col + 2
                new_end = merged_range.max_col + 2
                ws.merge_cells(start_row=3, start_column=new_start, end_row=3, end_column=new_end)

            ws.cell(row=3, column=insert_col).value = "IST"
            ws.cell(row=4, column=insert_col).value = f"{current_year}e"
            style_cell(ws.cell(row=3, column=insert_col))
            style_cell(ws.cell(row=4, column=insert_col))

            ws.cell(row=3, column=insert_col + 1).value = "Plan"
            ws.cell(row=4, column=insert_col + 1).value = current_year + 1
            style_cell(ws.cell(row=3, column=insert_col + 1))
            style_cell(ws.cell(row=4, column=insert_col + 1))

            for row in range(5, end_row + 1):
                ws.cell(row=row, column=insert_col).value = None
                ws.cell(row=row, column=insert_col + 1).value = None
                style_cell(ws.cell(row=row, column=insert_col))
                style_cell(ws.cell(row=row, column=insert_col + 1))

            apply_veraenderung_formulas(ws, ist_col=insert_col, plan_col=insert_col + 1,
                                        vera_start_col=vera_start_col, end_row=end_row)

            unhide_cols = {1, insert_col, insert_col + 1}
            unhide_cols.update(range(vera_start_col + 2, vera_end_col + 4))

            for col in range(1, ws.max_column + 1):
                header3 = ws.cell(row=3, column=col).value
                header4 = str(ws.cell(row=4, column=col).value)
                if (header3 == "PLAN" and header4.replace("e", "").strip() == str(current_year)) or \
                   (header3 == "IST" and header4.replace("e", "").strip() in [str(current_year), str(current_year - 1), str(current_year - 2)]):
                    unhide_cols.add(col)
                    if col != insert_col and col != insert_col + 1:
                        ws.cell(row=4, column=col).value = header4.replace("e", "").strip()

            for col in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].hidden = col not in unhide_cols

            for col in unhide_cols:
                if col != 1:
                    col_letter = get_column_letter(col)
                    ws.column_dimensions[col_letter].width = 18

        wb.save(file_path)


def extract_valid_tokens(cell_value):
    if not cell_value:
        return []

    tokens = []
    # Split by + and - to preserve math operators
    parts = re.split(r'([+\-])', cell_value)

    for part in parts:
        part = part.strip()
        if not part:
            continue
        if part in ['+', '-']:
            tokens.append(part)
        else:
            # Further split on commas, semicolons, and newlines
            subparts = re.split(r'[\n;,]+', part)
            for sub in subparts:
                # Check for the presence of #
                if '#' in sub:
                    sub = sub.split('#')[0]  # Ignore everything after #
                cleaned = re.sub(r'\s+', '', sub)  # remove inner spaces
                if cleaned and not cleaned.isalpha():
                    tokens.append(cleaned)
    return tokens
def perform_custom_vlookup(current_ws, kosten_ws, end_row, current_year, sheet_name):
    print(f"\n Processing VLOOKUP for sheet: {sheet_name}")
    ist_col_index = None
    for col in range(1, current_ws.max_column + 1):
        h1 = current_ws.cell(row=3, column=col).value
        h2 = str(current_ws.cell(row=4, column=col).value).replace("e", "").strip()
        if h1 and h1.strip().upper() == "IST" and h2 == str(current_year - 1):
            ist_col_index = col
            print(f" Found IST column for year {current_year - 1} → Column {get_column_letter(col)} (Index {col})")
            break
    if ist_col_index is None:
        print(" IST column with previous year not found.")
        return

    for row in range(5, end_row + 1):
        ab_value = str(current_ws.cell(row=row, column=28).value)
        if not ab_value.strip():
            continue
        print(f"\n🖎 Row {row}, AB value: {ab_value}")
        tokens = extract_valid_tokens(ab_value)
        print(f" Tokens extracted: {tokens}")

        expr = ""
        for token in tokens:
            if token in ['+', '-']:
                expr += f" {token} "
                continue

            match_value = None
            for kosten_row in range(2, kosten_ws.max_row + 1):
                key = str(kosten_ws.cell(row=kosten_row, column=1).value)
                if token == key or token in key:
                    match_value = kosten_ws.cell(row=kosten_row, column=4).value
                    if match_value is None:
                        print(f"   Matched '{token}' but D is None → using 0")
                        match_value = 0
                    print(f"   Matched '{token}' in row {kosten_row} → D: {match_value}")
                    break
            if match_value is None:
                print(f"   No match found for '{token}', using 0")
                match_value = 0

            expr += str(int(match_value))

        if not expr.strip():
            print(f" No valid tokens to evaluate at row {row} — skipping.")
            continue

        try:
            result = eval(expr)
            print(f" Final Expression: {expr} = {result}")
        except Exception as e:
            print(f" Error evaluating expression '{expr}': {e}")
            result = 0

        #  Write to the IST column for current_year - 1, if not merged
        cell = current_ws.cell(row=row, column=ist_col_index)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            print(f" Cannot write to merged cell at {get_column_letter(ist_col_index)}{row} — skipping.")
        else:
            if result >= 1000:
                cell.value = round(result / 1000, 3)  # Write the value in thousands with three decimal places
            else:
                cell.value = round(result)  # Write the value as is if less than 1000
            print(f" Value {cell.value} written to {get_column_letter(ist_col_index)}{row}")

def post_processing_with_vlookup(directory):
    kosten_file = None
    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") and file.endswith((".xlsx", ".xlsm")):
            kosten_file = os.path.join(directory, file)
            break
    if not kosten_file:
        print(" Kostenstelle file not found.")
        return

    print(f"\n Kostenstelle file found: {os.path.basename(kosten_file)}")
    kosten_wb = openpyxl.load_workbook(kosten_file, data_only=True)
    kosten_ws = kosten_wb.active

    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue
        file_path = os.path.join(directory, file)
        print(f"\n Processing file: {file}")
        wb = openpyxl.load_workbook(file_path)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            tab_color = rgb_to_hex_name(ws.sheet_properties.tabColor)
            if ws.sheet_properties.tabColor is None or "green" not in tab_color.lower():
                continue
            end_row = find_end_row(ws, sheet)
            perform_custom_vlookup(ws, kosten_ws, end_row, datetime.now().year, sheet)
        wb.save(file_path)
        print(f" File saved: {file}")

def main():
    root = Tk()
    root.withdraw()
    selected_directory = filedialog.askdirectory(title="Select Directory with Excel Files")
    if selected_directory:
        process_excel_files(selected_directory)
        post_processing_with_vlookup(selected_directory)

if __name__ == "__main__":
    main()
