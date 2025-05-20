
    process_sachaufwand_links(wb, file_path)
  File "Zum testen\2505.py", line 279, in process_sachaufwand_links
    apply_veraenderung_formulas(
  File "Zum testen\2505.py", line 324, in apply_veraenderung_formulas
    ws.cell(row=row, column=diff_col).value = f"={plan_letter}{row}-{ist_letter}{row}"
AttributeError: 'MergedCell' object attribute 'value' is read-only
