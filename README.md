from openpyxl.utils import get_column_letter

def set_final_column_widths(ws, width=20):
    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        if not ws.column_dimensions[col_letter].hidden:
            ws.column_dimensions[col_letter].width = width
set_final_column_widths(ws)
