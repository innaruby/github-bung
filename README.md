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

from openpyxl.utils import get_column_letter

def apply_final_sums(ws, end_row):
    from openpyxl.cell.cell import MergedCell

    sheet_name = ws.title
    print(f"\n🧮 Starting apply_final_sums for sheet: {sheet_name}")

    veraenderung_cols = find_merged_veraenderung_columns(ws)
    if not veraenderung_cols:
        print(f"❌ [Sheet: {sheet_name}] Veränderung columns not found.")
        return

    vera_col_start = veraenderung_cols[0]
    visible_cols = [col for col in range(2, vera_col_start)
                    if not ws.column_dimensions[get_column_letter(col)].hidden]

    if not visible_cols:
        print(f"⚠️ [Sheet: {sheet_name}] No visible columns before Veränderung.")
        return

    visible_rows = [r for r in range(5, end_row + 1) if not ws.row_dimensions[r].hidden]
    zwischensumme_rows = []
    summe_rows = []

    for row in visible_rows:
        cell_val = str(ws.cell(row=row, column=1).value or "").strip().lower()
        if ws.cell(row=row, column=1).font.bold:
            if "zwischensumme" in cell_val:
                zwischensumme_rows.append(row)
            elif "summe" in cell_val:
                summe_rows.append(row)

    zwischensumme_values = {}  # {row: {col: value}}
    previous = 4
    for z_row in zwischensumme_rows:
        rows_to_sum = [r for r in visible_rows if previous < r < z_row]
        zwischensumme_values[z_row] = {}
        print(f"🧩 Zwischensumme row {z_row} → summing rows {rows_to_sum}")

        for col in visible_cols:
            total = 0
            value_details = []
            col_letter = get_column_letter(col)
            for r in rows_to_sum:
                cell = ws.cell(row=r, column=col)
                parsed = parse_numeric(cell.value)
                total += parsed
                value_details.append(f"{col_letter}{r}={parsed}")
            target_cell = ws.cell(row=z_row, column=col)
            target_cell.value = None
            target_cell.value = total
            zwischensumme_values[z_row][col] = total

            print(f"   🟢 Zwischensumme {col_letter}{z_row} = {' + '.join([v.split('=')[0] for v in value_details])} = {total}")
            print(f"     🔍 Values: {', '.join(value_details)}")
        previous = z_row

    for s_row in summe_rows:
        print(f"🧮 Summe row {s_row} processing...")
        prev_z_row = max([z for z in zwischensumme_rows if z < s_row], default=None)
        rows_to_sum = [r for r in visible_rows if (prev_z_row or 4) < r < s_row]

        for col in visible_cols:
            col_letter = get_column_letter(col)
            total = 0
            value_details = []

            if prev_z_row:
                zw_val = zwischensumme_values.get(prev_z_row, {}).get(col, 0)
                total += zw_val
                value_details.append(f"{col_letter}{prev_z_row}={zw_val}")

            for r in rows_to_sum:
                cell = ws.cell(row=r, column=col)
                parsed = parse_numeric(cell.value)
                total += parsed
                value_details.append(f"{col_letter}{r}={parsed}")

            target_cell = ws.cell(row=s_row, column=col)
            target_cell.value = None
            target_cell.value = total

            print(f"   ✅ Summe {col_letter}{s_row} = {' + '.join([v.split('=')[0] for v in value_details])} = {total}")
            print(f"     🔍 Values: {', '.join(value_details)}")

def parse_numeric(val):
    if val is None or val == "":
        return 0
    elif isinstance(val, (int, float)):
        return val
    elif isinstance(val, str) and val.strip().startswith("="):
        try:
            return eval(val.strip().lstrip("="))
        except:
            return 0
    try:
        return float(str(val).strip())
    except:
        return 0


def parse_numeric(val):
    if val is None or val == "":
        return 0
    elif isinstance(val, (int, float)):
        return val
    elif isinstance(val, str) and val.strip().startswith("="):
        try:
            return eval(val.strip().lstrip("="))
        except:
            return 0
    try:
        return float(str(val).strip())
    except:
        return 0




def process_sachaufwand_links(wb, file_path):
    # Step 1: Reload workbook WITHOUT data_only to access formulas
    wb_with_formulas = openpyxl.load_workbook(file_path, data_only=False)

    # Step 2: Get 'Sachaufwand' sheet (case-insensitive) from both workbooks
    sach_sheet = None
    sach_sheet_formula = None
    for sheet in wb.sheetnames:
        if sheet.lower() == "sachaufwand":
            sach_sheet = wb[sheet]
            break
    for sheet in wb_with_formulas.sheetnames:
        if sheet.lower() == "sachaufwand":
            sach_sheet_formula = wb_with_formulas[sheet]
            break

    if not sach_sheet or not sach_sheet_formula:
        print("❌ 'Sachaufwand' sheet not found in one or both workbooks.")
        return

    print("\n🔍 Starting process_sachaufwand_links for 'Sachaufwand'...")

    # Step 3: Clear values and formulas from row 5 onward, columns B+
    cleared_cells = 0
    max_row = sach_sheet_formula.max_row
    max_col = sach_sheet_formula.max_column

    for row in range(5, max_row + 1):
        for col in range(2, max_col + 1):  # Start from column B
            cell_formula = sach_sheet_formula.cell(row=row, column=col)
            cell_target = sach_sheet.cell(row=row, column=col)
            if isinstance(cell_formula.value, str) and cell_formula.value.strip().startswith("="):
                cell_target.value = None
                cleared_cells += 1
            elif cell_target.value is not None:
                cell_target.value = None
                cleared_cells += 1

    print(f"🧹 Cleared {cleared_cells} cells from 'Sachaufwand' (excluding headers and column A).")

    # Step 4: Prepare lowercase sheet name map
    sheet_map = {s.lower(): s for s in wb.sheetnames}

    # Step 5: Define function to find end row
    def find_end_row(sheet):
        for row in range(sheet.max_row, 0, -1):
            if any(sheet.cell(row=row, column=col).value is not None for col in range(1, sheet.max_column + 1)):
                return row
        return sheet.max_row

    # Step 6: Find end row in Sachaufwand
    end_row = find_end_row(sach_sheet)
    print(f"✅ Detected end row in 'Sachaufwand': {end_row}")

    # Step 7: Loop through each visible row
    for row in range(5, end_row + 1):
        if sach_sheet.row_dimensions[row].hidden:
            continue

        ref_value = sach_sheet.cell(row=row, column=1).value
        if not ref_value or not isinstance(ref_value, str):
            continue

        ref_key = ref_value.strip().lower()
        matched_sheet_name = sheet_map.get(ref_key)
        if not matched_sheet_name:
            continue

        matched_sheet = wb[matched_sheet_name]

        # Step 8: Find bold 'Summe' row
        summe_row = None
        for r in range(5, matched_sheet.max_row + 1):
            cell = matched_sheet.cell(row=r, column=1)
            if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
                summe_row = r
                break

        if not summe_row:
            continue

        # Step 9: Identify visible source columns
        visible_source_cols = [
            col for col in range(2, matched_sheet.max_column + 1)
            if not matched_sheet.column_dimensions[get_column_letter(col)].hidden
        ]

        if not visible_source_cols:
            continue

        # Step 10: Collect values from Summe row
        data_to_copy = []
        for col in visible_source_cols:
            val = matched_sheet.cell(row=summe_row, column=col).value
            data_to_copy.append((col, val))

        # Step 11: Identify visible target columns in Sachaufwand
        visible_target_cols = [
            col for col in range(2, sach_sheet.max_column + 1)
            if not sach_sheet.column_dimensions[get_column_letter(col)].hidden
        ]

        # Step 12: Paste values into target row
        for i, col in enumerate(visible_target_cols):
            if i < len(data_to_copy):
                value = data_to_copy[i][1]
                sach_sheet.cell(row=row, column=col).value = value

    # Step 13: Apply Veränderung formulas to Sachaufwand sheet
    veraenderung_cols = find_merged_veraenderung_columns(sach_sheet)
    if veraenderung_cols:
        vera_start_col, vera_end_col = veraenderung_cols
        print(f"🛠️ Applying Veränderung formulas to 'Sachaufwand': Columns {get_column_letter(vera_start_col)} to {get_column_letter(vera_end_col)}")

        # Identify IST and PLAN columns (2 left of vera_start_col)
        ist_col = vera_start_col - 2
        plan_col = vera_start_col - 1

        # Check if these columns are within bounds
        if ist_col >= 1 and plan_col >= 1:
            apply_veraenderung_formulas(
                sach_sheet,
                ist_col=ist_col,
                plan_col=plan_col,
                vera_start_col=vera_start_col,
                end_row=end_row
            )
        else:
            print("⚠️ Not enough columns to determine IST and PLAN in 'Sachaufwand'. Skipping.")
    else:
        print("⚠️ Veränderung columns not found in 'Sachaufwand'. Skipping formula application.")








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

from openpyxl.cell.cell import MergedCell

from openpyxl.cell.cell import MergedCell

def apply_veraenderung_formulas(ws, ist_col, plan_col, vera_start_col, end_row):
    diff_col = vera_start_col
    perc_col = vera_start_col + 1

    print(f"💡 IST column: {get_column_letter(ist_col)} ({ist_col})")
    print(f"💡 PLAN column: {get_column_letter(plan_col)} ({plan_col})")
    print(f"💡 Veränderung DIFF column: {get_column_letter(diff_col)} ({diff_col})")
    print(f"💡 Veränderung % column: {get_column_letter(perc_col)} ({perc_col})")

    for row in range(5, end_row + 1):
        plan_letter = get_column_letter(plan_col)
        ist_letter = get_column_letter(ist_col)
        diff_letter = get_column_letter(diff_col)

        # Debug input cell values
        ist_val = ws.cell(row=row, column=ist_col).value
        plan_val = ws.cell(row=row, column=plan_col).value
        print(f"🔍 Row {row}: {plan_letter}{row}={plan_val}, {ist_letter}{row}={ist_val}")

        if isinstance(ws.cell(row=row, column=diff_col), MergedCell):
            print(f"⚠️ Skipping row {row} DIFF - merged cell at {diff_letter}{row}")
            continue
        if isinstance(ws.cell(row=row, column=perc_col), MergedCell):
            print(f"⚠️ Skipping row {row} % - merged cell at {get_column_letter(perc_col)}{row}")
            continue

        formula1 = f"={plan_letter}{row}-{ist_letter}{row}"
        formula2 = f"=IF({ist_letter}{row}=0,0,({diff_letter}{row}/{ist_letter}{row})*100)"
        print(f"🧾 Writing to Row {row}: {get_column_letter(diff_col)} → {formula1}, {get_column_letter(perc_col)} → {formula2}")

        ws.cell(row=row, column=diff_col).value = formula1
        ws.cell(row=row, column=perc_col).value = formula2



def delete_columns_B_and_C(ws):
    ws.delete_cols(2, 2)

def final_sum_pass(directory):
    print("\n📘 Starting final sum pass (apply_final_sums at very end)...")
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
            end_row = find_end_row(ws, sheet_name)
            apply_final_sums(ws, end_row)

        wb.save(file_path)
        print(f"✅ Final sum pass completed for: {file}")
import copy
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

            reference_col = vera_start_col + 2  # or use any reliable column index
            for row in range(5, end_row + 1):
                ref_cell = ws.cell(row=row, column=reference_col)
                if ref_cell.has_style:
                    ist_cell = ws.cell(row=row, column=insert_col)
                    plan_cell = ws.cell(row=row, column=insert_col + 1)

                    ist_cell.font = copy.copy(ref_cell.font)
                    ist_cell.alignment = copy.copy(ref_cell.alignment)
                    ist_cell.border = copy.copy(ref_cell.border)
                    ist_cell.fill = copy.copy(ref_cell.fill)
                    ist_cell.number_format = ref_cell.number_format

                    plan_cell.font = copy.copy(ref_cell.font)
                    plan_cell.alignment = copy.copy(ref_cell.alignment)
                    plan_cell.border = copy.copy(ref_cell.border)
                    plan_cell.fill = copy.copy(ref_cell.fill)
                    plan_cell.number_format = ref_cell.number_format


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
            #apply_final_sums(ws, end_row)
        wb.save(file_path)
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


from openpyxl.utils import get_column_letter
from datetime import datetime

import copy





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
        final_sum_pass(selected_directory)

        for file in os.listdir(selected_directory):
            if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
                continue
            file_path = os.path.join(selected_directory, file)
            wb = openpyxl.load_workbook(file_path)
            process_sachaufwand_links(wb, file_path) 
            for sheet_name in wb.sheetnames:
                if sheet_name.lower() == 'sachaufwand':
                    ws = wb[sheet_name]
                    end_row = find_end_row(ws, sheet_name)
                    apply_final_sums(ws, end_row)
                    #apply_number_format_to_ist_and_plan_columns(ws, end_row)
                    print(f"📘 Zwischensumme and Summe logic applied to 'Sachaufwand' in file: {file}")
            wb.save(file_path)
            print(f"💾 Final update (Sachaufwand) saved in file: {file}")
        apply_light_grey_fill_final(selected_directory)

if __name__ == "__main__":
    main()

