def sum_visible_columns(ws, end_row):
    # Find the row with the bold "Summe" in column A
    summe_row = None
    for row in range(5, end_row + 1):
        cell = ws[f"A{row}"]
        if cell.value and str(cell.value).lower() == "summe" and cell.font.bold:
            summe_row = row
            break

    if summe_row is None:
        print("No bold 'Summe' row found.")
        return

    # Find the column index for "Veränderung"
    vera_cols = find_merged_veraenderung_columns(ws)
    if vera_cols is None:
        print("No 'Veränderung' column found.")
        return
    vera_start_col, vera_end_col = vera_cols

    # Loop through all columns from B to the column before "Veränderung"
    for col in range(2, vera_start_col):
        # Check if the column is visible
        col_letter = get_column_letter(col)
        if ws.column_dimensions[col_letter].hidden:
            continue  # Skip hidden columns

        # Clear any existing formula in the cell before writing the sum (except for "Summe" row and "Veränderung" columns)
        summe_cell = ws.cell(row=summe_row, column=col)
        if summe_cell.has_formula:
            summe_cell.value = None  # Remove formula

        total_sum = 0
        # Sum the visible rows from 5 to the end row
        for row in range(5, end_row + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value is None or cell.value == "":
                total_sum += 0  # Treat blank cells as zero
            else:
                total_sum += cell.value

        # Write the total sum in the "Summe" row
        ws.cell(row=summe_row, column=col).value = total_sum
        print(f"Sum written in {get_column_letter(col)}{summe_row} → Total: {total_sum}")

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

            # Perform sum for visible columns before "Veränderung"
            sum_visible_columns(ws, end_row)

        wb.save(file_path)

# Modify the main function to call the new processing function
def main():
    root = Tk()
    root.withdraw()
    selected_directory = filedialog.askdirectory(title="Select Directory with Excel Files")
    if selected_directory:
        process_excel_files_with_sum(selected_directory)
        post_processing_with_vlookup(selected_directory)

if __name__ == "__main__":
    main()
