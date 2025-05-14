i would like to modify the code without affecting its current logic such that , during the V-look up currently the value from column index D is being copied to the column index which contain the keyword IST in row 3 and current year -1 value in row 4 . Here instead of column index D , i want to copy the values from column index C , to the  column index which contain the keyword IST in row 3 and current year -1 value in row 4. Likewise i would like to copy the values from column index H in kostenselle file to the column index  which contain the keyword IST in row 3 and current year  value in row 4  and additionally  i would like to copy the values from column index I in kostenselle file to the column index  which contain the keyword PLAN in row 3 and current year + 1 value in row 4 . the final values written to the target columns shouldnt be like 1619,532
202,6
9142,954
113,01
1051,982
  but instead it should be  like 1620, 203 , 9143 , 113 and 1052 please update the following code 
def perform_custom_vlookup(current_ws, kosten_ws, end_row, current_year, sheet_name):
    print(f"\n Processing VLOOKUP for sheet: {sheet_name}")
    ist_col_index = None
    for col in range(1, current_ws.max_column + 1):
        h1 = current_ws.cell(row=3, column=col).value
        h2 = str(current_ws.cell(row=4, column=col).value).replace("e", "").strip()
        if h1 and h1.strip().upper() == "IST" and h2 == str(current_year - 1):
            ist_col_index = col
            print(f" Found IST column for year {current_year - 1} → Column {get_column_letter(col)} (Index {col})")
            break
    if ist_col_index is None:
        print(" IST column with previous year not found.")
        return

    for row in range(5, end_row + 1):
        ab_value = str(current_ws.cell(row=row, column=28).value)
        if not ab_value.strip():
            continue
        print(f"\n🖎 Row {row}, AB value: {ab_value}")
        tokens = extract_valid_tokens(ab_value)
        print(f" Tokens extracted: {tokens}")

        expr = ""
        for token in tokens:
            if token in ['+', '-']:
                expr += f" {token} "
                continue

            match_value = None
            for kosten_row in range(2, kosten_ws.max_row + 1):
                key = str(kosten_ws.cell(row=kosten_row, column=1).value)
                if token == key or token in key:
                    match_value = kosten_ws.cell(row=kosten_row, column=4).value
                    if match_value is None:
                        print(f"   Matched '{token}' but D is None → using 0")
                        match_value = 0
                    print(f"   Matched '{token}' in row {kosten_row} → D: {match_value}")
                    break
            if match_value is None:
                print(f"   No match found for '{token}', using 0")
                match_value = 0

            expr += str(int(match_value))

        if not expr.strip():
            print(f" No valid tokens to evaluate at row {row} — skipping.")
            continue

        try:
            result = eval(expr)
            print(f" Final Expression: {expr} = {result}")
        except Exception as e:
            print(f" Error evaluating expression '{expr}': {e}")
            result = 0

        #  Write to the IST column for current_year - 1, if not merged
        cell = current_ws.cell(row=row, column=ist_col_index)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            print(f" Cannot write to merged cell at {get_column_letter(ist_col_index)}{row} — skipping.")
        else:
            if result >= 1000:
                cell.value = round(result / 1000, 3)  # Write the value in thousands with three decimal places
            else:
                cell.value = round(result)  # Write the value as is if less than 1000
            print(f" Value {cell.value} written to {get_column_letter(ist_col_index)}{row}")
