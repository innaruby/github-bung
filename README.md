from openpyxl.utils import get_column_letter

def apply_final_sums_with_zwischensumme(ws, end_row):
    sheet_name = ws.title
    print(f"\n🧮 Starting enhanced sum logic for sheet: {sheet_name}")

    # Step 1: Identify the start of the 'Veränderung' column to get visible columns before it
    veraenderung_cols = find_merged_veraenderung_columns(ws)
    if not veraenderung_cols:
        print(f"❌ [Sheet: {sheet_name}] Veränderung columns not found.")
        return

    vera_col_start = veraenderung_cols[0]
    visible_cols = [col for col in range(2, vera_col_start)
                    if not ws.column_dimensions[get_column_letter(col)].hidden]

    if not visible_cols:
        print(f"⚠️ [Sheet: {sheet_name}] No visible columns found before Veränderung.")
        return

    # Step 2: Identify 'Zwischensumme' and 'Summe' rows in visible rows only
    zwischensumme_row = None
    summe_row = None
    visible_rows = []

    for row in range(5, end_row + 1):
        hidden = ws.row_dimensions[row].hidden
        cell = ws.cell(row=row, column=1)
        val = str(cell.value).strip().lower() if cell.value else ""
        is_bold = cell.font.bold

        if not hidden:
            visible_rows.append(row)
            if val == "zwischensumme" and is_bold:
                zwischensumme_row = row
            elif "summe" in val and is_bold:
                summe_row = row

    if not zwischensumme_row or not summe_row:
        print(f"❌ [Sheet: {sheet_name}] Could not find both 'Zwischensumme' and 'Summe' rows.")
        return

    print(f"✅ [Sheet: {sheet_name}] Found 'Zwischensumme' in row {zwischensumme_row}, 'Summe' in row {summe_row}")

    # Step 3: Categorize rows before and after 'Zwischensumme' up to 'Summe'
    pre_zwischen_rows = [r for r in visible_rows if r < zwischensumme_row]
    post_zwischen_rows = [r for r in visible_rows if zwischensumme_row < r < summe_row]

    # Step 4: Process each visible column
    for col in visible_cols:
        col_letter = get_column_letter(col)
        sum_pre = 0
        sum_post = 0

        # Sum values before Zwischensumme
        for row in pre_zwischen_rows:
            cell = ws.cell(row=row, column=col)
            val = cell.value
            try:
                sum_pre += float(val) if val not in [None, ""] else 0
            except:
                print(f"⚠️ [Sheet: {sheet_name}] Skipped non-numeric at {col_letter}{row}: {val}")

        # Write to Zwischensumme row
        ws.cell(row=zwischensumme_row, column=col).value = sum_pre

        # Sum values after Zwischensumme up to Summe
        for row in post_zwischen_rows:
            cell = ws.cell(row=row, column=col)
            val = cell.value
            try:
                sum_post += float(val) if val not in [None, ""] else 0
            except:
                print(f"⚠️ [Sheet: {sheet_name}] Skipped non-numeric at {col_letter}{row}: {val}")

        total_sum = sum_pre + sum_post

        # Write to Summe row
        ws.cell(row=summe_row, column=col).value = total_sum
        print(f"🟢 [Sheet: {sheet_name}] Wrote {sum_pre} to {col_letter}{zwischensumme_row} and {total_sum} to {col_letter}{summe_row}")
