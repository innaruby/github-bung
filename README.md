for row in range(5, end_row + 1):
    ws.cell(row=row, column=insert_col).value = None
    ws.cell(row=row, column=insert_col + 1).value = None
    style_cell(ws.cell(row=row, column=insert_col))
    style_cell(ws.cell(row=row, column=insert_col + 1))

reference_col = vera_start_col + 2  # or use any reliable column index
for row in range(5, end_row + 1):
    ref_cell = ws.cell(row=row, column=reference_col)
    if ref_cell.has_style:
        ist_cell = ws.cell(row=row, column=insert_col)
        plan_cell = ws.cell(row=row, column=insert_col + 1)

        ist_cell.font = copy.copy(ref_cell.font)
        ist_cell.alignment = copy.copy(ref_cell.alignment)
        ist_cell.border = copy.copy(ref_cell.border)
        ist_cell.fill = copy.copy(ref_cell.fill)
        ist_cell.number_format = ref_cell.number_format

        plan_cell.font = copy.copy(ref_cell.font)
        plan_cell.alignment = copy.copy(ref_cell.alignment)
        plan_cell.border = copy.copy(ref_cell.border)
        plan_cell.fill = copy.copy(ref_cell.fill)
        plan_cell.number_format = ref_cell.number_format
