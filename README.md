from openpyxl.utils import get_column_letter

def replace_blanks_with_zero(ws, end_row):
    """
    Replaces all blank (None or "") cells in visible columns and rows 
    from row 5 to the specified end_row with 0.
    """
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        if ws.column_dimensions[col_letter].hidden:
            continue  # Skip hidden columns

        for row in range(5, end_row + 1):
            if ws.row_dimensions[row].hidden:
                continue  # Skip hidden rows

            cell = ws.cell(row=row, column=col)
            if cell.value in (None, ""):
                cell.value = 0

    print(f"Blank spaces replaced with 0 in visible cells from row 5 to {end_row}.")
