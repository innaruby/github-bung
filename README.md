def process_sachaufwand_links(wb):
    # Step 1: Locate Sachaufwand sheet (case-insensitive)
    sach_sheet = None
    for sheet in wb.sheetnames:
        if sheet.lower() == "sachaufwand":
            sach_sheet = wb[sheet]
            break

    if not sach_sheet:
        print("❌ No 'Sachaufwand' sheet found.")
        return

    print("🔍 Starting process_sachaufwand_links for 'Sachaufwand'...")

    # Step 2: Clear all formulas from the Sachaufwand sheet
    formula_count = 0
    for row in sach_sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().startswith("="):
                cell.value = None
                formula_count += 1
    print(f"🧹 Cleared {formula_count} formula cells in 'Sachaufwand'.")

    # Step 3: Prepare sheet map for case-insensitive lookup
    sheet_map = {s.lower(): s for s in wb.sheetnames}

    # Step 4: Detect end row in Sachaufwand sheet
    try:
        end_row = find_end_row(sach_sheet, "Sachaufwand")
        print(f"✅ Determined end row in 'Sachaufwand': {end_row}")
    except Exception as e:
        print(f"❌ Error determining end row in 'Sachaufwand': {e}")
        return

    # Step 5: Loop through visible rows only from row 5 to end_row
    for row in range(5, end_row + 1):
        if sach_sheet.row_dimensions[row].hidden:
            print(f"🚫 Skipping hidden row {row} in 'Sachaufwand'.")
            continue

        ref_value = sach_sheet.cell(row=row, column=1).value
        if not ref_value or not isinstance(ref_value, str):
            print(f"⚠️ Skipping row {row}: Empty or invalid value in column A.")
            continue

        ref_key = ref_value.strip().lower()
        matched_sheet_name = sheet_map.get(ref_key)
        if not matched_sheet_name:
            print(f"❌ No matching sheet for '{ref_value}' (row {row}).")
            continue

        try:
            matched_sheet = wb[matched_sheet_name]
            print(f"\n🔗 Row {row}: Linking to sheet '{matched_sheet_name}'...")
        except Exception as e:
            print(f"❌ Error loading sheet '{matched_sheet_name}': {e}")
            continue

        # Step 6: Find 'Summe' row in matched sheet
        summe_row = None
        try:
            for r in range(5, matched_sheet.max_row + 1):
                cell = matched_sheet.cell(row=r, column=1)
                if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
                    summe_row = r
                    break
        except Exception as e:
            print(f"❌ Error searching for 'Summe' in sheet '{matched_sheet_name}': {e}")
            continue

        if not summe_row:
            print(f"⚠️ 'Summe' row not found in sheet '{matched_sheet_name}'.")
            continue
        print(f"✅ Found 'Summe' in row {summe_row} of sheet '{matched_sheet_name}'.")

        # Step 7: Find second 'Veränderung' column
        veraenderung_cols = []
        try:
            for r in [3, 4]:
                for c in range(2, matched_sheet.max_column + 1):
                    val = matched_sheet.cell(row=r, column=c).value
                    if val and "veränderung" in str(val).lower():
                        veraenderung_cols.append(c)
        except Exception as e:
            print(f"❌ Error detecting 'Veränderung' columns in sheet '{matched_sheet_name}': {e}")
            continue

        if len(veraenderung_cols) < 2:
            print(f"⚠️ Only {len(veraenderung_cols)} 'Veränderung' column(s) found in sheet '{matched_sheet_name}'.")
            continue

        vera_limit_col = veraenderung_cols[1]
        print(f"✅ Using 2nd 'Veränderung' column index: {vera_limit_col} ({get_column_letter(vera_limit_col)})")

        # Step 8: Collect values from visible columns in matched sheet's Summe row
        data_to_copy = []
        for col in range(2, vera_limit_col + 1):
            col_letter = get_column_letter(col)
            if not matched_sheet.column_dimensions[col_letter].hidden:
                val = matched_sheet.cell(row=summe_row, column=col).value
                data_to_copy.append((col, val))

        print(f"📦 Collected {len(data_to_copy)} values from '{matched_sheet_name}' to paste into 'Sachaufwand'.")

        # Step 9: Paste into visible columns in 'Sachaufwand' starting at column B
        paste_col_idx = 2
        pasted_count = 0
        for col_idx, val in data_to_copy:
            col_letter = get_column_letter(paste_col_idx)
            if not sach_sheet.column_dimensions[col_letter].hidden:
                sach_sheet.cell(row=row, column=paste_col_idx).value = val
                pasted_count += 1
            paste_col_idx += 1

        print(f"✅ Pasted {pasted_count} values into row {row} of 'Sachaufwand'.")
