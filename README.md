def apply_final_sums(ws, end_row):
    sheet_name = ws.title
    print(f"\n🧮 Starting apply_final_sums for sheet: {sheet_name}")

    # Step 1: Find Veränderung column and visible columns before it
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

    # Step 2: Find bold 'Zwischensumme' and 'Summe' in visible rows
    special_rows = []
    for row in range(5, end_row + 1):
        if ws.row_dimensions[row].hidden:
            continue
        cell = ws.cell(row=row, column=1)
        val = str(cell.value).strip().lower() if cell.value else ""
        if cell.font.bold and ("zwischensumme" in val or "summe" in val):
            special_rows.append((row, val))

    # Step 3: Collect visible rows that are not special rows
    visible_data_rows = [row for row in range(5, end_row + 1)
                         if not ws.row_dimensions[row].hidden and
                         all(row != r for r, _ in special_rows)]

    all_rows = visible_data_rows + [r for r, _ in special_rows]
    all_rows.sort()

    # Step 4: Sum segments up to each special row
    segment_start = 0
    for idx, (special_row, label) in enumerate(special_rows):
        segment_rows = [r for r in all_rows if segment_start < r < special_row]
        print(f"🔍 Processing '{label}' at row {special_row}, summing rows: {segment_rows}")

        for col in visible_cols:
            total = 0
            for r in segment_rows:
                val = ws.cell(row=r, column=col).value
                parsed_val = 0
                if val is None or val == "":
                    continue
                elif isinstance(val, (int, float)):
                    parsed_val = val
                elif isinstance(val, str) and val.strip().startswith("="):
                    try:
                        parsed_val = eval(val.strip().lstrip("="))
                    except:
                        continue
                else:
                    try:
                        parsed_val = float(str(val).strip())
                    except:
                        continue
                total += parsed_val

            # Clear formula/value in the target cell before writing
            target_cell = ws.cell(row=special_row, column=col)
            target_cell.value = None
            target_cell.value = total
            print(f"🟢 [Sheet: {sheet_name}] Wrote sum {total} to {get_column_letter(col)}{special_row}")

        segment_start = special_row
