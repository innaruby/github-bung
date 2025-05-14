from openpyxl.utils import get_column_letter

def apply_final_sums(ws, end_row):
    sheet_name = ws.title
    print(f"\n🧮 Starting apply_final_sums for sheet: {sheet_name}")

    # Step 1: Locate "Veränderung" column and visible columns before it
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

    # Step 2: Find the "Summe" row (bold text in column A from row 5 to end_row)
    summe_row = None
    for row in range(5, end_row + 1):
        cell = ws.cell(row=row, column=1)
        val = str(cell.value).strip().lower() if cell.value else ""
        is_bold = cell.font.bold
        print(f"🔎 [Sheet: {sheet_name}] Row {row}, A: '{val}', Bold: {is_bold}")
        if val == "summe" and is_bold:
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

    # Step 4: Sum values in each visible column and write to Summe row
    for col in visible_cols:
        col_letter = get_column_letter(col)
        total = 0
        for row in visible_rows:
            cell = ws.cell(row=row, column=col)
            val = cell.value
            if val is None or val == "":
                val = 0
            elif isinstance(val, (int, float)):
                pass  # use as is
            else:
                try:
                    val = float(str(val).strip())
                except:
                    print(f"⚠️ [Sheet: {sheet_name}] Non-numeric value ignored at {col_letter}{row}: {val}")
                    continue
            total += val

        # Clear formula or value and write new total
        target_cell = ws.cell(row=summe_row, column=col)
        target_cell.value = None  # Clear any formula
        target_cell.value = total  # Write sum
        print(f"🟢 [Sheet: {sheet_name}] Wrote sum {total} to {col_letter}{summe_row}")
