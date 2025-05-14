def sum_visible_columns(ws, end_row):
    # Find the column range to process
    for row in [3, 4]:
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == row and merged_range.max_row == row:
                cell_value = ws.cell(row=row, column=merged_range.min_col).value
                if cell_value and "veränderung" in str(cell_value).lower():
                    vera_start_col = merged_range.min_col
                    vera_end_col = merged_range.max_col
                    break
    else:
        # If "Veränderung" column is not found, return
        return

    # Identify "Summe" row
    summe_row = None
    for row in range(7, end_row + 1):
        if ws[f"A{row}"].value and "summe" in str(ws[f"A{row}"].value).lower() and ws[f"A{row}"].font.bold:
            summe_row = row
            break

    if not summe_row:
        print("Summe row not found.")
        return

    # Loop through each column from B to the one before "Veränderung"
    for col in range(2, vera_start_col):
        column_letter = get_column_letter(col)
        column_sum = 0

        # Loop through visible rows from 5 to the end row
        for row in range(5, end_row + 1):
            if not ws.row_dimensions[row].hidden:  # Check if row is visible
                cell_value = ws.cell(row=row, column=col).value
                if cell_value is None:
                    cell_value = 0
                column_sum += cell_value

        # Write the sum to the "Summe" row
        ws.cell(row=summe_row, column=col).value = column_sum
        print(f"Sum written to {column_letter}{summe_row}: {column_sum}")



def process_excel_files(directory):
    current_year = datetime.now().year

    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue
        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)
        sheet_colors = get_sheet_tab_colors(file_path)

        for sheet_name in wb.sheetnames:
            tab_color = sheet_colors.get(sheet_name, "")
            if "green" not in tab_color.lower():
                continue

            ws = wb[sheet_name]
            delete_columns_B_and_C(ws)
            end_row = find_end_row(ws, sheet_name)
            vera_cols = find_merged_veraenderung_columns(ws)
            if vera_cols is None:
                continue

            vera_start_col, vera_end_col = vera_cols
            insert_col = vera_start_col

            merged_to_restore = []
            for merged_range in list(ws.merged_cells.ranges):
                if merged_range.min_row == 3 and merged_range.max_row == 3:
                    if merged_range.min_col == vera_start_col and merged_range.max_col == vera_end_col:
                        merged_to_restore.append(merged_range)
                        ws.unmerge_cells(str(merged_range))

            # Your existing operations here ...

            sum_visible_columns(ws, end_row)  # Add this line to perform the summing operation

            wb.save(file_path)
