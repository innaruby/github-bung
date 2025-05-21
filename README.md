in apply_veraenderung_formulas
    ws.cell(row=row, column=diff_col).value = f"={plan_letter}{row}-{ist_letter}{row}"
AttributeError: 'MergedCell' object attribute 'value' is read-only
