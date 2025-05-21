from openpyxl.utils import get_column_letter

def fill_empty_cells_with_zero(ws, end_row):
    visible_columns = [col for col in range(1, ws.max_column + 1)
                       if not ws.column_dimensions[get_column_letter(col)].hidden]
    
    for row in range(5, end_row + 1):
        if ws.row_dimensions[row].hidden:
            continue
        for col in visible_columns:
            cell = ws.cell(row=row, column=col)
            if cell.value in (None, ""):
                cell.value = 0
fill_empty_cells_with_zero(ws, end_row)
