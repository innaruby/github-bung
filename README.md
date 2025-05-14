from openpyxl.utils import get_column_letter

def process_sachaufwand_links(wb):
    # Step 1: Locate 'Sachaufwand' sheet
    sach_sheet = None
    for sheet in wb.sheetnames:
        if sheet.lower() == "sachaufwand":
            sach_sheet = wb[sheet]
            break

    if not sach_sheet:
        print("❌ No 'Sachaufwand' sheet found.")
        return

    print("\n🔍 Starting process_sachaufwand_links for 'Sachaufwand'...")

    # Step 2: Clear all formulas in Sachaufwand
    formula_count = 0
    for row in sach_sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().startswith("="):
                cell.value = None
                formula_count += 1
    print(f"🧹 Cleared {formula_count} formulas from 'Sachaufwand'.")

    # Step 3: Prepare lowercase sheet name map
    sheet_map = {s.lower(): s for s in wb.sheetnames}

    # Step 4: Find end row in Sachaufwand
    try:
        end_row = find_end_row(sach_sheet, "Sachaufwand")
        print(f"✅ Detected end row in 'Sachaufwand': {end_row}")
    except Exception as e:
        print(f"❌ Error determining end row in 'Sachaufwand': {e}")
        return

    # Step 5: Loop through visible rows in Sachaufwand
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

        # Step 6: Find bold 'Summe' row
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

        # Step 7: Identify last visible column in matched sheet
        last_visible_col = None
        visible_source_cols = []
        for col in range(2, matched_sheet.max_column + 1):
            col_letter = get_column_letter(col)
            if not matched_sheet.column_dimensions[col_letter].hidden:
                visible_source_cols.append(col)
                last_visible_col = col

        if not visible_source_cols:
            print(f"⚠️ Row {row}: No visible columns in source sheet '{matched_sheet_name}'.")
            continue

        print(f"✅ Last visible column in '{matched_sheet_name}' is {last_visible_col} ({get_column_letter(last_visible_col)})")
        print(f"👁️ Visible source columns: {[get_column_letter(c) for c in visible_source_cols]}")

        # Step 8: Collect values from Summe row in visible columns
        data_to_copy = []
        for col in visible_source_cols:
            val = matched_sheet.cell(row=summe_row, column=col).value
            data_to_copy.append((col, val))
        print(f"📦 Collected {len(data_to_copy)} values from 'Summe' row.")

        # Step 9: Debug visible columns in Sachaufwand before pasting
        print(f"👁️ Visible columns in 'Sachaufwand' from B:")
        for i in range(2, 2 + len(data_to_copy)):
            col_letter = get_column_letter(i)
            hidden = sach_sheet.column_dimensions[col_letter].hidden
            print(f"   {col_letter} → {'HIDDEN' if hidden else 'VISIBLE'}")

        # Step 10: Paste into Sachaufwand row only into visible columns
        paste_col_idx = 2
        pasted_count = 0
        for _, val in data_to_copy:
            col_letter = get_column_letter(paste_col_idx)
            is_hidden = sach_sheet.column_dimensions[col_letter].hidden
            if not is_hidden:
                sach_sheet.cell(row=row, column=paste_col_idx).value = val
                pasted_count += 1
                print(f"✅ Pasted '{val}' into 'Sachaufwand'! Cell: {col_letter}{row}")
            else:
                print(f"🚫 Skipped hidden column {col_letter}{row}")
            paste_col_idx += 1

        print(f"✅ Pasted {pasted_count} values into row {row} of 'Sachaufwand'.")
