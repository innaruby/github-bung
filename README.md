    the following code after execution i found  a problem . That problem can be explained through an example . In the excel file in the sheet which has yellow colour tab , it has the keyword Veränderung in row 3 but from both the columns V and W . From row 4 onwards there are no more merged cells in V and W. After processing with this python file , i found that the data in the column index V and W only shifted towards right when adding the two new columns it should not happend. The structure of the cells also should be taken along when shifting , now that merged cell is lost where the Veränderung keyword is present and that merged cell now belong to the newly added two columns which should not happend. Please modify the code                                                                                                                                                                                                                                                             import os 
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
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
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")

def extract_lookup_keys(value):
    if not value:
        return []
    parts = re.split(r'[,+\\-]', str(value))
    return [re.sub(r'\\s+', '', part.strip()) for part in parts if part.strip()]

def lookup_and_aggregate(keys, kosten_ws, column_letter):
    values = []
    for row in kosten_ws.iter_rows(min_row=2):
        ref = str(row[0].value) if row[0].value else ""
        for key in keys:
            if key and (ref == key or key in ref):
                val = row[ord(column_letter) - 65].value
                if isinstance(val, (int, float)):
                    values.append(val)
    return sum(values) if values else None

def process_excel_files(directory):
    current_year = datetime.now().year
    kostenstelle_path = None
    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") and file.endswith((".xlsx", ".xlsm")):
            kostenstelle_path = os.path.join(directory, file)
            break
    if not kostenstelle_path:
        print("Kostenstelle file not found.")
        return

    kosten_wb = openpyxl.load_workbook(kostenstelle_path, data_only=True)
    kosten_ws = kosten_wb.active

    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue
        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)
        sheet_colors = get_sheet_tab_colors(file_path)

        for sheet_name in wb.sheetnames:
            tab_color = sheet_colors.get(sheet_name, "")
            if "yellow" not in tab_color.lower():
                continue

            ws = wb[sheet_name]
            end_row = find_end_row(ws, sheet_name)
            vera_cols = find_merged_veraenderung_columns(ws)
            if vera_cols is None:
                continue

            vera_start_col, vera_end_col = vera_cols
            insert_col = vera_start_col

            # Insert two new columns before the merged Veränderung column
            ws.insert_cols(insert_col, 2)

            # Write headers
            ws.cell(row=3, column=insert_col).value = "Plan"
            ws.cell(row=4, column=insert_col).value = current_year + 1
            style_cell(ws.cell(row=3, column=insert_col))
            style_cell(ws.cell(row=4, column=insert_col))

            ws.cell(row=3, column=insert_col + 1).value = "IST"
            ws.cell(row=4, column=insert_col + 1).value = f"{current_year}e"
            style_cell(ws.cell(row=3, column=insert_col + 1))
            style_cell(ws.cell(row=4, column=insert_col + 1))

            # Lookup logic
            for row in range(5, end_row + 1):
                ab_value = ws[f"AB{row}"].value
                if not ab_value:
                    continue
                keys = extract_lookup_keys(ab_value)
                plan_value = lookup_and_aggregate(keys, kosten_ws, "D")
                ist_value = lookup_and_aggregate(keys, kosten_ws, "C")
                if plan_value is not None:
                    ws.cell(row=row, column=insert_col).value = plan_value
                if ist_value is not None:
                    ws.cell(row=row, column=insert_col + 1).value = ist_value

            # Columns to unhide
            unhide_cols = {1, insert_col, insert_col + 1}
            unhide_cols.update(range(vera_start_col + 2, vera_end_col + 3))  # Adjusted for shifted columns

            # Additional logic to preserve IST/Plan from previous years
            for col in range(1, ws.max_column + 1):
                header3 = ws.cell(row=3, column=col).value
                header4 = str(ws.cell(row=4, column=col).value).replace("e", "").strip()
                if (header3 == "Plan" and header4 == str(current_year + 1)) or \
                   (header3 == "IST" and header4 in [str(current_year), str(current_year - 1), str(current_year - 2)]):
                    unhide_cols.add(col)

            # Hide all other columns
            for col in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].hidden = col not in unhide_cols

        wb.save(file_path)

def main():
    root = Tk()
    root.withdraw()
    selected_directory = filedialog.askdirectory(title="Select Directory with Excel Files")
    if selected_directory:
        process_excel_files(selected_directory)

if __name__ == "__main__":
    main()
