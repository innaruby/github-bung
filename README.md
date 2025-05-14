def sum_visible_columns(ws, end_row):
    # Find the column range to process
    for row in [3, 4]:
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == row and merged_range.max_row == row:
                cell_value = ws.cell(row=row, column=merged_range.min_col).value
                if cell_value and "veränderung" in str(cell_value).lower():
                    vera_start_col = merged_range.min_col
                    vera_end_col = merged_range.max_col
                    break
    else:
        # If "Veränderung" column is not found, return
        return

    # Identify "Summe" row
    summe_row = None
    for row in range(7, end_row + 1):
        if ws[f"A{row}"].value and "summe" in str(ws[f"A{row}"].value).lower() and ws[f"A{row}"].font.bold:
            summe_row = row
            break

    if not summe_row:
        print("Summe row not found.")
        return

    # Loop through each column from B to the one before "Veränderung"
    for col in range(2, vera_start_col):
        column_letter = get_column_letter(col)
        column_sum = 0

        # Loop through visible rows from 5 to the end row
        for row in range(5, end_row + 1):
            if not ws.row_dimensions[row].hidden:  # Check if row is visible
                cell_value = ws.cell(row=row, column=col).value
                if cell_value is None:
                    cell_value = 0
                column_sum += cell_value

        # Before writing the sum, remove formulas in the "Summe" row for columns other than A and the "Veränderung" columns
        if col != 1 and col < vera_start_col:
            cell = ws.cell(row=summe_row, column=col)
            if cell.has_formula:
                cell.value = None  # Remove formula

        # Write the sum to the "Summe" row
        ws.cell(row=summe_row, column=col).value = column_sum
        print(f"Sum written to {column_letter}{summe_row}: {column_sum}")
