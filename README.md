def apply_number_format_to_ist_and_plan_columns(ws, end_row):
    """
    Finds visible IST and PLAN columns and formats their numeric values:
    - Divides values >= 1000 by 1000
    - Applies German-style number formatting
    """
    current_year = datetime.now().year
    ist_col = None
    plan_col = None

    for col in range(1, ws.max_column + 1):
        if ws.column_dimensions[get_column_letter(col)].hidden:
            continue
        header_3 = str(ws.cell(row=3, column=col).value).strip().upper()
        header_4 = str(ws.cell(row=4, column=col).value).strip().replace("e", "")
        if header_3 == "IST" and header_4 == str(current_year):
            ist_col = col
        elif header_3 == "PLAN" and header_4 == str(current_year + 1):
            plan_col = col

    number_format = '#,##0.000'  # German-style decimal format

    for target_col in [ist_col, plan_col]:
        if not target_col:
            continue
        for row in range(5, end_row + 1):
            cell = ws.cell(row=row, column=target_col)
            if isinstance(cell.value, (int, float)):
                value = cell.value
                if value >= 1000:
                    cell.value = value / 1000
            cell.number_format = number_format
