from openpyxl.cell.cell import MergedCell

def apply_veraenderung_formulas(ws, ist_col, plan_col, vera_start_col, end_row):
    diff_col = vera_start_col
    perc_col = vera_start_col + 1

    print(f"💡 IST column: {get_column_letter(ist_col)} ({ist_col})")
    print(f"💡 PLAN column: {get_column_letter(plan_col)} ({plan_col})")
    print(f"💡 Veränderung DIFF column: {get_column_letter(diff_col)} ({diff_col})")
    print(f"💡 Veränderung % column: {get_column_letter(perc_col)} ({perc_col})")

    for row in range(5, end_row + 1):
        plan_letter = get_column_letter(plan_col)
        ist_letter = get_column_letter(ist_col)
        diff_letter = get_column_letter(diff_col)

        # Debug input cell values
        ist_val = ws.cell(row=row, column=ist_col).value
        plan_val = ws.cell(row=row, column=plan_col).value
        print(f"🔍 Row {row}: {plan_letter}{row}={plan_val}, {ist_letter}{row}={ist_val}")

        if isinstance(ws.cell(row=row, column=diff_col), MergedCell):
            print(f"⚠️ Skipping row {row} DIFF - merged cell at {diff_letter}{row}")
            continue
        if isinstance(ws.cell(row=row, column=perc_col), MergedCell):
            print(f"⚠️ Skipping row {row} % - merged cell at {get_column_letter(perc_col)}{row}")
            continue

        formula1 = f"={plan_letter}{row}-{ist_letter}{row}"
        formula2 = f"=IF({ist_letter}{row}=0,0,({diff_letter}{row}/{ist_letter}{row})*100)"
        print(f"🧾 Writing to Row {row}: {get_column_letter(diff_col)} → {formula1}, {get_column_letter(perc_col)} → {formula2}")

        ws.cell(row=row, column=diff_col).value = formula1
        ws.cell(row=row, column=perc_col).value = formula2
