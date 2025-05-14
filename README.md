def sum_visible_columns(ws, end_row):
    # Find the column of "Veränderung"
    vera_cols = find_merged_veraenderung_columns(ws)
    if vera_cols is None:
        return
    vera_start_col, vera_end_col = vera_cols

    # Find the row where "Summe" is located in column A and is bold
    summe_row = None
    for row in range(7, end_row + 1):
        cell = ws[f"A{row}"]
        if cell.value and isinstance(cell.value, str) and cell.value.strip().lower() == "summe" and cell.font.bold:
            summe_row = row
            break

    if not summe_row:
        print("Summe row not found.")
        return

    # Loop over columns B to the column before "Veränderung" and sum the visible rows
    for col in range(2, vera_start_col):  # From column B to the column before "Veränderung"
        column_sum = 0
        for row in range(5, end_row + 1):
            cell = ws.cell(row=row, column=col)
            if not ws.row_dimensions[row].hidden and not ws.column_dimensions[get_column_letter(col)].hidden:
                # Treat blank cells as 0
                column_sum += cell.value if cell.value else 0
        
        # Remove formulas and write the sum in the Summe row
        summe_cell = ws.cell(row=summe_row, column=col)
        if summe_cell.has_formula:
            summe_cell.value = column_sum  # Replace formula with sum
        else:
            summe_cell.value = column_sum

        print(f"Sum for column {get_column_letter(col)}: {column_sum}")
    
    # Specifically for "Veränderung" columns, skip removing formulas
    for col in range(vera_start_col, vera_end_col + 1):
        summe_cell = ws.cell(row=summe_row, column=col)
        if summe_cell.has_formula:
            summe_cell.value = summe_cell.formula  # Keep the existing formula if any
        print(f"Skipping formula removal for column {get_column_letter(col)}")

