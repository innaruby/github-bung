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
from openpyxl.utils import get_column_letter

def apply_final_sums(ws, end_row):
    sheet_name = ws.title
    print(f"\n🧮 Starting apply_final_sums for sheet: {sheet_name}")

    # Step 1: Locate Veränderung column and visible columns before it
    veraenderung_cols = find_merged_veraenderung_columns(ws)
    if not veraenderung_cols:
        print(f"❌ [Sheet: {sheet_name}] Veränderung columns not found.")
        return

    vera_col_start = veraenderung_cols[0]
    print(f"✅ [Sheet: {sheet_name}] Veränderung starts at column index: {vera_col_start} ({get_column_letter(vera_col_start)})")

    visible_cols = []
    for col in range(2, vera_col_start):
        col_letter = get_column_letter(col)
        hidden = ws.column_dimensions[col_letter].hidden
        print(f"👁️ [Sheet: {sheet_name}] Column {col_letter} (index {col}) hidden: {hidden}")
        if not hidden:
            visible_cols.append(col)

    print(f"✅ [Sheet: {sheet_name}] Visible columns before Veränderung: {[get_column_letter(c) for c in visible_cols]}")
    if not visible_cols:
        print(f"⚠️ [Sheet: {sheet_name}] No visible columns found before Veränderung. Skipping.")
        return

    # Step 2: Find "Summe" row (any cell in column A containing "summe" and bold)
    summe_row = None
    for row in range(5, end_row + 1):
        cell = ws.cell(row=row, column=1)
        val = str(cell.value).strip().lower() if cell.value else ""
        is_bold = cell.font.bold
        print(f"🔎 [Sheet: {sheet_name}] Row {row}, A: '{val}', Bold: {is_bold}")
        if "summe" in val and is_bold:
            summe_row = row
            break

    if not summe_row:
        print(f"❌ [Sheet: {sheet_name}] 'Summe' row not found in column A.")
        return
    print(f"✅ [Sheet: {sheet_name}] Found 'Summe' in row: {summe_row}")

    # Step 3: Identify visible rows (excluding the Summe row)
    visible_rows = []
    for row in range(5, end_row + 1):
        hidden = ws.row_dimensions[row].hidden
        print(f"👁️ [Sheet: {sheet_name}] Row {row} hidden: {hidden}")
        if row != summe_row and not hidden:
            visible_rows.append(row)

    print(f"✅ [Sheet: {sheet_name}] Final visible rows for summing: {visible_rows}")
    if not visible_rows:
        print(f"⚠️ [Sheet: {sheet_name}] No visible rows found for summing.")
        return

    # Step 4: Process each visible column and compute sum with formula handling
    for col in visible_cols:
        col_letter = get_column_letter(col)
        total = 0
        value_details = []

        for row in visible_rows:
            cell = ws.cell(row=row, column=col)
            val = cell.value
            original_val = val
            parsed_val = 0  # default

            if val is None or val == "":
                parsed_val = 0
            elif isinstance(val, (int, float)):
                parsed_val = val
            elif isinstance(val, str) and val.strip().startswith("="):
                try:
                    # Evaluate basic numeric expressions (e.g., "=1000+250-50")
                    parsed_val = eval(val.strip().lstrip("="))
                except Exception as e:
                    print(f"⚠️ [Sheet: {sheet_name}] Failed to eval formula in {col_letter}{row}: {val} → {e}")
                    continue
            else:
                try:
                    parsed_val = float(str(val).strip())
                except Exception as e:
                    print(f"⚠️ [Sheet: {sheet_name}] Non-numeric value ignored at {col_letter}{row}: {val}")
                    continue

            value_details.append(f"{col_letter}{row}={parsed_val}")
            total += parsed_val

        # Log what was used
        formula_trace = " + ".join([v.split("=")[0] for v in value_details])
        value_trace = ", ".join(value_details)
        print(f"🔢 [Sheet: {sheet_name}] Values used for {col_letter}{summe_row}: {value_trace}")
        print(f"🧾 [Sheet: {sheet_name}] Formula simulated: {formula_trace} = {total}")

        # Clear formula and write sum
        target_cell = ws.cell(row=summe_row, column=col)
        target_cell.value = None
        target_cell.value = total
        print(f"🟢 [Sheet: {sheet_name}] Wrote sum {total} to {col_letter}{summe_row}")





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

            ws.cell(row=3, column=insert_col + 1).value = "PLAN"
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
            apply_final_sums(ws, end_row)
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
from openpyxl.utils import get_column_letter

def perform_custom_vlookup(current_ws, kosten_ws, end_row, current_year, sheet_name):
    print(f"\n Processing VLOOKUP for sheet: {sheet_name}")

    def find_column(header_1, header_2):
        for col in range(1, current_ws.max_column + 1):
            h1 = str(current_ws.cell(row=3, column=col).value).strip().upper()
            h2 = str(current_ws.cell(row=4, column=col).value).replace("e", "").strip()
            if h1 == header_1 and h2 == header_2:
                print(f" Found column → {header_1} {header_2} → {get_column_letter(col)} (Index {col})")
                return col
        print(f" Column not found → {header_1} {header_2}")
        return None

    # Locate target columns
    ist_prev_col = find_column("IST", str(current_year - 1))
    ist_curr_col = find_column("IST", str(current_year))
    plan_next_col = find_column("PLAN", str(current_year + 1))

    if not ist_prev_col:
        print(" IST column with previous year not found.")
        return

    for row in range(5, end_row + 1):
        ab_value = str(current_ws.cell(row=row, column=28).value)
        if not ab_value.strip():
            continue
        print(f"\n🖎 Row {row}, AB value: {ab_value}")
        tokens = extract_valid_tokens(ab_value)
        print(f" Tokens extracted: {tokens}")

        expr_c = ""  # Column C
        expr_h = ""  # Column H
        expr_i = ""  # Column I

        for token in tokens:
            if token in ['+', '-']:
                expr_c += f" {token} "
                expr_h += f" {token} "
                expr_i += f" {token} "
                continue

            val_c = val_h = val_i = 0
            for kosten_row in range(2, kosten_ws.max_row + 1):
                key = str(kosten_ws.cell(row=kosten_row, column=1).value)
                if token == key or token in key:
                    val_c = kosten_ws.cell(row=kosten_row, column=3).value or 0
                    val_h = kosten_ws.cell(row=kosten_row, column=8).value or 0
                    val_i = kosten_ws.cell(row=kosten_row, column=9).value or 0
                    print(f"   Matched '{token}' in row {kosten_row} → C: {val_c}, H: {val_h}, I: {val_i}")
                    break

            expr_c += str(int(val_c))
            expr_h += str(int(val_h))
            expr_i += str(int(val_i))

        # Helper to evaluate and write to Excel
        def evaluate_and_write(expr, col_index, label):
            if not expr.strip() or not col_index:
                return
            try:
                result = eval(expr)
                if result >= 1000:
                    final_val = float(f"{result / 1000:.3f}")  # Force float with 3 decimals
                else:
                    final_val = int(round(result))
                print(f" Final Expression ({label}): {expr} = {final_val}")
                cell = current_ws.cell(row=row, column=col_index)
                if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                    cell.value = final_val
                    print(f"  → Value {final_val} written to {get_column_letter(col_index)}{row}")
                else:
                    print(f"  → Cannot write to merged cell at {get_column_letter(col_index)}{row}")
            except Exception as e:
                print(f"  → Error evaluating expression for {label}: {e}")

        evaluate_and_write(expr_c, ist_prev_col, f"IST {current_year - 1}")
        evaluate_and_write(expr_h, ist_curr_col, f"IST {current_year}")
        evaluate_and_write(expr_i, plan_next_col, f"PLAN {current_year + 1}")






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
