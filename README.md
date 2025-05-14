def perform_custom_vlookup(current_ws, kosten_ws, end_row, current_year, sheet_name):
    print(f"\n Processing VLOOKUP for sheet: {sheet_name}")

    def find_column(header_1, header_2):
        for col in range(1, current_ws.max_column + 1):
            h1 = str(current_ws.cell(row=3, column=col).value).strip().upper()
            h2 = str(current_ws.cell(row=4, column=col).value).replace("e", "").strip()
            if h1 == header_1 and h2 == header_2:
                print(f" Found column → {header_1} {header_2} → {get_column_letter(col)} (Index {col})")
                return col
        print(f" Column not found → {header_1} {header_2}")
        return None

    # Find target columns
    ist_prev_col = find_column("IST", str(current_year - 1))
    ist_curr_col = find_column("IST", str(current_year))
    plan_next_col = find_column("PLAN", str(current_year + 1))

    if not ist_prev_col:
        print(" IST column with previous year not found.")
        return

    for row in range(5, end_row + 1):
        ab_value = str(current_ws.cell(row=row, column=28).value)
        if not ab_value.strip():
            continue
        print(f"\n🖎 Row {row}, AB value: {ab_value}")
        tokens = extract_valid_tokens(ab_value)
        print(f" Tokens extracted: {tokens}")

        expr_c = ""  # Column C
        expr_h = ""  # Column H
        expr_i = ""  # Column I

        for token in tokens:
            if token in ['+', '-']:
                expr_c += f" {token} "
                expr_h += f" {token} "
                expr_i += f" {token} "
                continue

            val_c = val_h = val_i = 0
            for kosten_row in range(2, kosten_ws.max_row + 1):
                key = str(kosten_ws.cell(row=kosten_row, column=1).value)
                if token == key or token in key:
                    val_c = kosten_ws.cell(row=kosten_row, column=3).value or 0
                    val_h = kosten_ws.cell(row=kosten_row, column=8).value or 0
                    val_i = kosten_ws.cell(row=kosten_row, column=9).value or 0
                    print(f"   Matched '{token}' in row {kosten_row} → C: {val_c}, H: {val_h}, I: {val_i}")
                    break

            expr_c += str(int(val_c))
            expr_h += str(int(val_h))
            expr_i += str(int(val_i))

        # Evaluate and write results
        def evaluate_and_write(expr, col_index, label):
            if not expr.strip() or not col_index:
                return
            try:
                result = eval(expr)
                result = round(result)
                print(f" Final Expression ({label}): {expr} = {result}")
                cell = current_ws.cell(row=row, column=col_index)
                if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                    cell.value = result
                    print(f"  → Value {result} written to {get_column_letter(col_index)}{row}")
                else:
                    print(f"  → Cannot write to merged cell at {get_column_letter(col_index)}{row}")
            except Exception as e:
                print(f"  → Error evaluating expression for {label}: {e}")

        evaluate_and_write(expr_c, ist_prev_col, f"IST {current_year - 1}")
        evaluate_and_write(expr_h, ist_curr_col, f"IST {current_year}")
        evaluate_and_write(expr_i, plan_next_col, f"PLAN {current_year + 1}")
