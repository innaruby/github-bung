def sum_visible_columns(ws, end_row):
    # Find the row where the "Summe" keyword is located in column A
    summe_row = None
    for row in range(5, end_row + 1):
        cell = ws[f"A{row}"]
        if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
            summe_row = row
            break

    if not summe_row:
        print(" 'Summe' not found in column A.")
        return

    # Find the columns to be summed (from B until the one before "Veränderung" column)
    vera_cols = find_merged_veraenderung_columns(ws)
    if not vera_cols:
        print(" 'Veränderung' column not found.")
        return
    vera_start_col, vera_end_col = vera_cols

    for col in range(2, vera_start_col):  # from column B to the one before "Veränderung"
        column_sum = 0
        for row in range(5, end_row + 1):
            cell = ws.cell(row=row, column=col)
            if cell.row_dimensions.hidden:  # Skip hidden rows
                continue
            column_sum += (cell.value if cell.value is not None else 0)

        # Directly remove any formula in the "Summe" row and write the sum value
        if col != 1 and col < vera_start_col:  # If the column is not A or one of the "Veränderung" columns
            sum_cell = ws.cell(row=summe_row, column=col)
            sum_cell.value = column_sum  # Directly assign the summed value, clearing any existing formula

        print(f" Sum for column {get_column_letter(col)}: {column_sum}")

def process_excel_files_with_sum(directory):
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
            end_row = find_end_row(ws, sheet_name)
            sum_visible_columns(ws, end_row)  # Call the sum function
            wb.save(file_path)

def main():
    root = Tk()
    root.withdraw()
    selected_directory = filedialog.askdirectory(title="Select Directory with Excel Files")
    if selected_directory:
        process_excel_files_with_sum(selected_directory)

if __name__ == "__main__":
    main()
