from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

def fill_empty_cells_with_zero(ws, end_row):
    merged_cells_set = set()
    for merged_range in ws.merged_cells.ranges:
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                merged_cells_set.add((row, col))

    for row in range(5, end_row + 1):
        if ws.row_dimensions[row].hidden:
            continue
        for col in range(1, ws.max_column + 1):
            if ws.column_dimensions[get_column_letter(col)].hidden:
                continue

            if (row, col) in merged_cells_set and (row, col) != (min(r for r, c in merged_cells_set if c == col), col):
                continue  # skip non-top-left merged cells

            cell = ws.cell(row=row, column=col)
            val = cell.value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                if not isinstance(cell, MergedCell):
                    cell.value = 0
