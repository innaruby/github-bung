def sum_visible_columns(ws, end_row):
    # Find the row where the "Summe" keyword is located in column A
    summe_row = None
    for row in range(5, end_row + 1):
        cell = ws[f"A{row}"]
        if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
            summe_row = row
            print(f"Found 'Summe' keyword in row {summe_row}")  # Debug statement
            break

    if not summe_row:
        print(" 'Summe' not found in column A.")
        return

    # Find the columns to be summed (from B until the one before "Veränderung" column)
    vera_cols = find_merged_veraenderung_columns(ws)
    if not vera_cols:
        print(" 'Veränderung' column not found.")
        return
    vera_start_col, vera_end_col = vera_cols
    print(f"Summing columns from B to {get_column_letter(vera_start_col - 1)}")  # Debug statement

    for col in range(2, vera_start_col):  # from column B to the one before "Veränderung"
        column_sum = 0
        print(f"Summing column {get_column_letter(col)}")  # Debug statement
        
        for row in range(5, end_row + 1):
            if ws.row_dimensions[row].hidden:  # Skip hidden rows
                continue
            cell = ws.cell(row=row, column=col)
            cell_value = cell.value if cell.value is not None else 0
            print(f"Row {row}, Column {get_column_letter(col)}: Adding value {cell_value}")  # Debug statement
            column_sum += cell_value

        # Directly remove any formula in the "Summe" row and write the sum value
        if col != 1 and col < vera_start_col:  # If the column is not A or one of the "Veränderung" columns
            sum_cell = ws.cell(row=summe_row, column=col)
            sum_cell.value = column_sum  # Directly assign the summed value, clearing any existing formula

        print(f" Sum for column {get_column_letter(col)}: {column_sum}")  # Debug statement
