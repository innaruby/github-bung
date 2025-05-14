def perform_custom_vlookup(current_ws, kosten_ws, end_row, current_year, sheet_name):
    print(f"\n Processing VLOOKUP for sheet: {sheet_name}")
    ist_prev_year_col = ist_current_year_col = plan_next_year_col = None

    # Identify target columns
    for col in range(1, current_ws.max_column + 1):
        header = current_ws.cell(row=3, column=col).value
        year = str(current_ws.cell(row=4, column=col).value).replace("e", "").strip()

        if header and header.strip().upper() == "IST":
            if year == str(current_year - 1):
                ist_prev_year_col = col
                print(f" Found IST column for year {current_year - 1} → Column {get_column_letter(col)}")
            elif year == str(current_year):
                ist_current_year_col = col
                print(f" Found IST column for year {current_year} → Column {get_column_letter(col)}")

        elif header and header.strip().upper() == "PLAN" and year == str(current_year + 1):
            plan_next_year_col = col
            print(f" Found PLAN column for year {current_year + 1} → Column {get_column_letter(col)}")

    if not ist_prev_year_col:
        print(" IST column with previous year not found.")
        return

    for row in range(5, end_row + 1):
        ab_value = str(current_ws.cell(row=row, column=28).value)
        if not ab_value.strip():
            continue

        print(f"\n🖎 Row {row}, AB value: {ab_value}")
        tokens = extract_valid_tokens(ab_value)
        print(f" Tokens extracted: {tokens}")

        def get_kosten_value(token, column_index):
            for kosten_row in range(2, kosten_ws.max_row + 1):
                key = str(kosten_ws.cell(row=kosten_row, column=1).value)
                if token == key or token in key:
                    value = kosten_ws.cell(row=kosten_row, column=column_index).value
                    return float(str(value).replace(",", ".")) if value else 0
            return 0

        # Processing and writing for each target column
        for target_col, kosten_col in [(ist_prev_year_col, 3), (ist_current_year_col, 8), (plan_next_year_col, 9)]:
            if target_col is None:
                continue

            expr = ""
            for token in tokens:
                if token in ['+', '-']:
                    expr += f" {token} "
                    continue

                match_value = get_kosten_value(token, kosten_col)
                expr += str(match_value)

            if not expr.strip():
                print(f" No valid tokens to evaluate at row {row} for column {get_column_letter(target_col)} — skipping.")
                continue

            try:
                result = eval(expr)
                if result >= 1000:
                    final_value = round(result / 1000, 3)
                else:
                    final_value = round(result)
                print(f" Expression for {get_column_letter(target_col)}{row}: {expr} = {final_value}")
            except Exception as e:
                print(f" Error evaluating expression '{expr}': {e}")
                final_value = 0

            # Writing final rounded value
            cell = current_ws.cell(row=row, column=target_col)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                print(f" Cannot write to merged cell at {get_column_letter(target_col)}{row} — skipping.")
                continue

            cell.value = final_value
            print(f" Value {final_value} written to {get_column_letter(target_col)}{row}")
