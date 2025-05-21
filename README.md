from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

def find_merged_veraenderung_columns(ws):
    for row in [3, 4]:
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == row and merged_range.max_row == row:
                cell_value = ws.cell(row=row, column=merged_range.min_col).value
                if cell_value and "veränderung" in str(cell_value).lower():
                    return (merged_range.min_col, merged_range.max_col)
        # Check unmerged adjacent cells
        for col in range(1, ws.max_column):
            val1 = ws.cell(row=row, column=col).value
            val2 = ws.cell(row=row, column=col + 1).value
            if (val1 and "veränderung" in str(val1).lower()) and \
               (val2 and "veränderung" in str(val2).lower()):
                return (col, col + 1)
    return None

def apply_veraenderung_formulas(ws, ist_col, plan_col, vera_start_col, end_row):
    diff_col = vera_start_col         # First Veränderung column (DIFFERENCE)
    perc_col = vera_start_col + 1     # Second Veränderung column (% CHANGE)

    print(f"\n🧮 Applying Veränderung formulas in '{ws.title}'")
    print(f"   IST: {get_column_letter(ist_col)} ({ist_col}), PLAN: {get_column_letter(plan_col)} ({plan_col})")
    print(f"   Veränderung cols: {get_column_letter(diff_col)} ({diff_col}), {get_column_letter(perc_col)} ({perc_col})")

    for row in range(5, end_row + 1):
        plan_letter = get_column_letter(plan_col)
        ist_letter = get_column_letter(ist_col)
        diff_letter = get_column_letter(diff_col)

        diff_cell = ws.cell(row=row, column=diff_col)
        perc_cell = ws.cell(row=row, column=perc_col)

        if isinstance(diff_cell, MergedCell) or isinstance(perc_cell, MergedCell):
            print(f"⚠️ Row {row}: Skipped (merged cell)")
            continue

        diff_formula = f"={plan_letter}{row}-{ist_letter}{row}"
        perc_formula = f"=IF({ist_letter}{row}=0,0,({diff_letter}{row}/{ist_letter}{row})*100)"

        diff_cell.value = diff_formula
        perc_cell.value = perc_formula

        print(f"✅ Row {row}: {get_column_letter(diff_col)}{row} = {diff_formula}")
        print(f"            {get_column_letter(perc_col)}{row} = {perc_formula}")

def process_sachaufwand_links(wb, file_path):
    wb_with_formulas = openpyxl.load_workbook(file_path, data_only=False)

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

    cleared_cells = 0
    max_row = sach_sheet_formula.max_row
    max_col = sach_sheet_formula.max_column

    for row in range(5, max_row + 1):
        for col in range(2, max_col + 1):
            cell_formula = sach_sheet_formula.cell(row=row, column=col)
            cell_target = sach_sheet.cell(row=row, column=col)
            if isinstance(cell_formula.value, str) and cell_formula.value.strip().startswith("="):
                cell_target.value = None
                cleared_cells += 1
            elif cell_target.value is not None:
                cell_target.value = None
                cleared_cells += 1

    print(f"🧹 Cleared {cleared_cells} cells from 'Sachaufwand' (excluding headers and column A).")

    sheet_map = {s.lower(): s for s in wb.sheetnames}

    def find_end_row(sheet):
        for row in range(sheet.max_row, 0, -1):
            if any(sheet.cell(row=row, column=col).value is not None for col in range(1, sheet.max_column + 1)):
                return row
        return sheet.max_row

    end_row = find_end_row(sach_sheet)
    print(f"✅ Detected end row in 'Sachaufwand': {end_row}")

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

        summe_row = None
        for r in range(5, matched_sheet.max_row + 1):
            cell = matched_sheet.cell(row=r, column=1)
            if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
                summe_row = r
                break

        if not summe_row:
            continue

        visible_source_cols = [
            col for col in range(2, matched_sheet.max_column + 1)
            if not matched_sheet.column_dimensions[get_column_letter(col)].hidden
        ]

        if not visible_source_cols:
            continue

        data_to_copy = []
        for col in visible_source_cols:
            val = matched_sheet.cell(row=summe_row, column=col).value
            data_to_copy.append((col, val))

        visible_target_cols = [
            col for col in range(2, sach_sheet.max_column + 1)
            if not sach_sheet.column_dimensions[get_column_letter(col)].hidden
        ]

        for i, col in enumerate(visible_target_cols):
            if i < len(data_to_copy):
                value = data_to_copy[i][1]
                sach_sheet.cell(row=row, column=col).value = value

    veraenderung_cols = find_merged_veraenderung_columns(sach_sheet)
    if veraenderung_cols:
        vera_start_col = veraenderung_cols[0]
        ist_col_sach = vera_start_col - 2
        plan_col_sach = vera_start_col - 1

        apply_veraenderung_formulas(
            sach_sheet,
            ist_col=ist_col_sach,
            plan_col=plan_col_sach,
            vera_start_col=vera_start_col,
            end_row=end_row
        )
        print(f"🧮 Applied Veränderung formulas to 'Sachaufwand' columns {get_column_letter(vera_start_col)} and {get_column_letter(vera_start_col + 1)}")
    else:
        print("⚠️ Veränderung columns not found in 'Sachaufwand'.")
