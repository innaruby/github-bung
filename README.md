def apply_number_format_to_ist_and_plan_columns(ws, end_row):
    """
    Converts and formats values in visible IST and PLAN columns from 420000 to 420.000.
    """
    current_year = datetime.now().year
    ist_col = None
    plan_col = None

    for col in range(1, ws.max_column + 1):
        if ws.column_dimensions[get_column_letter(col)].hidden:
            continue
        header_3 = str(ws.cell(row=3, column=col).value or "").strip().upper()
        header_4 = str(ws.cell(row=4, column=col).value or "").strip().replace("e", "")
        if header_3 == "IST" and header_4 == str(current_year):
            ist_col = col
        elif header_3 == "PLAN" and header_4 == str(current_year + 1):
            plan_col = col

    # Use 3 decimal places
    number_format = '#,##0.000'

    for target_col in [ist_col, plan_col]:
        if not target_col:
            continue
        for row in range(5, end_row + 1):
            cell = ws.cell(row=row, column=target_col)
            val = cell.value
            if isinstance(val, (int, float)):
                new_val = val / 1000
                cell.value = new_val
                cell.number_format = number_format
