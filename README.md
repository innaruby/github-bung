from openpyxl.utils import get_column_letter
import openpyxl

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

    # Step 5: Find end row in Sachaufwand
    try:
        end_row = find_end_row(sach_sheet, "Sachaufwand")
        print(f"✅ Detected end row in 'Sachaufwand': {end_row}")
    except Exception as e:
        print(f"❌ Error determining end row in 'Sachaufwand': {e}")
        return

    # Step 6: Loop through each visible row
    for row in range(5, end_row + 1):
        if sach_sheet.row_dimensions[row].hidden:
            print(f"🚫 Row {row} in 'Sachaufwand' is hidden. Skipping.")
            continue

        ref_value = sach_sheet.cell(row=row, column=1).value
        if not ref_value or not isinstance(ref_value, str):
            print(f"⚠️ Row {row}: Empty or invalid value in column A.")
            continue

        ref_key = ref_value.strip().lower()
        matched_sheet_name = sheet_map.get(ref_key)
        if not matched_sheet_name:
            print(f"❌ Row {row}: No matching sheet found for '{ref_value}'.")
            continue

        matched_sheet = wb[matched_sheet_name]
        print(f"\n🔗 Row {row}: Linking to sheet '{matched_sheet_name}'...")

        # Step 7: Find bold 'Summe' row
        summe_row = None
        for r in range(5, matched_sheet.max_row + 1):
            cell = matched_sheet.cell(row=r, column=1)
            if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
                summe_row = r
                break

        if not summe_row:
            print(f"⚠️ Row {row}: No bold 'Summe' row found in '{matched_sheet_name}'.")
            continue
        print(f"✅ Found 'Summe' at row {summe_row} in '{matched_sheet_name}'.")

        # Step 8: Identify visible source columns
        visible_source_cols = []
        for col in range(2, matched_sheet.max_column + 1):
            col_letter = get_column_letter(col)
            if not matched_sheet.column_dimensions[col_letter].hidden:
                visible_source_cols.append(col)

        if not visible_source_cols:
            print(f"⚠️ Row {row}: No visible columns in source sheet '{matched_sheet_name}'.")
            continue

        print(f"✅ Last visible column in '{matched_sheet_name}': {get_column_letter(visible_source_cols[-1])}")
        print(f"👁️ Visible source columns: {[get_column_letter(c) for c in visible_source_cols]}")

        # Step 9: Collect values from Summe row
        data_to_copy = []
        for col in visible_source_cols:
            val = matched_sheet.cell(row=summe_row, column=col).value
            data_to_copy.append((col, val))
        print(f"📦 Collected {len(data_to_copy)} values from 'Summe' row.")

        # Step 10: Identify visible target columns in Sachaufwand
        visible_target_cols = []
        for col in range(2, sach_sheet.max_column + 1):
            col_letter = get_column_letter(col)
            if not sach_sheet.column_dimensions[col_letter].hidden:
                visible_target_cols.append(col)

        if not visible_target_cols:
            print(f"❌ No visible target columns in 'Sachaufwand'. Skipping paste.")
            continue

        print(f"✅ Last visible column in 'Sachaufwand': {get_column_letter(visible_target_cols[-1])}")
        print(f"👁️ Visible target columns: {[get_column_letter(c) for c in visible_target_cols]}")

        # Step 11: Paste values into target row
        pasted_count = 0
        copy_index = 0
        for col in visible_target_cols:
            col_letter = get_column_letter(col)
            if copy_index < len(data_to_copy):
                value = data_to_copy[copy_index][1]
                sach_sheet.cell(row=row, column=col).value = value
                pasted_count += 1
                print(f"✅ Pasted '{value}' to {col_letter}{row}")
                copy_index += 1
            else:
                print(f"⚠️ No more values to paste at {col_letter}{row}.")

        print(f"✅ Completed pasting {pasted_count} values into row {row} of 'Sachaufwand'.")
