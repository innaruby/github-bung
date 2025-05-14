def apply_final_sums(ws, end_row):
    from openpyxl.utils import column_index_from_string

    # Step 1: Find columns: from B to the column before "Veränderung", only visible ones
    veraenderung_cols = find_merged_veraenderung_columns(ws)
    if not veraenderung_cols:
        return
    vera_col_start = veraenderung_cols[0]
    visible_cols = [
        col for col in range(2, vera_col_start)
        if not ws.column_dimensions[get_column_letter(col)].hidden
    ]

    # Step 2: Find the "Summe" row in column A (bold), starting from row 5
    summe_row = None
    for row in range(5, end_row + 1):
        cell = ws.cell(row=row, column=1)
        if str(cell.value).strip().lower() == "summe" and cell.font.bold:
            summe_row = row
            break
    if not summe_row:
        return

    # Step 3: Identify visible rows (5 to end_row, excluding the Summe row)
    visible_rows = [
        row for row in range(5, end_row + 1)
        if row != summe_row and not ws.row_dimensions[row].hidden
    ]

    # Step 4: Sum up values in each visible column and write into the Summe row
    for col in visible_cols:
        col_letter = get_column_letter(col)
        total = 0
        for row in visible_rows:
            cell = ws.cell(row=row, column=col)
            val = cell.value
            if isinstance(val, (int, float)):
                total += val
            elif val is None or val == "":
                total += 0

        # Remove formula before writing
        target_cell = ws.cell(row=summe_row, column=col)
        target_cell.value = total
