from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

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

    # Step 13: Apply Veränderung formulas to the detected Veränderung columns
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
