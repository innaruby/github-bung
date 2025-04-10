def find_column(ws, keyword, row=3):
    keyword = keyword.strip().lower()
    for col_idx in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=row, column=col_idx).value
        if cell_value:
            cell_text = str(cell_value).strip().lower()
            if keyword in cell_text:
                print(f"Found '{keyword}' in column {col_idx} with header '{cell_value}'")
                return col_idx
    print(f"Keyword '{keyword}' not found in row {row}")
    return None
