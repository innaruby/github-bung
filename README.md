from openpyxl.cell.cell import MergedCell

def apply_veraenderung_formulas(ws, ist_col, plan_col, vera_start_col, end_row):
    diff_col = vera_start_col + 2
    perc_col = vera_start_col + 3

    for row in range(5, end_row + 1):
        plan_letter = get_column_letter(plan_col)
        ist_letter = get_column_letter(ist_col)
        diff_letter = get_column_letter(diff_col)

        # Skip if target cell is a merged cell
        if isinstance(ws.cell(row=row, column=diff_col), MergedCell):
            continue
        if isinstance(ws.cell(row=row, column=perc_col), MergedCell):
            continue

        ws.cell(row=row, column=diff_col).value = f"={plan_letter}{row}-{ist_letter}{row}"
        ws.cell(row=row, column=perc_col).value = f"=IF({ist_letter}{row}=0,0,({diff_letter}{row}/{ist_letter}{row})*100)"
