import openpyxl

def get_column_a_colors(file_path, sheet_name='Sheet1'):
    # Load the workbook and select the sheet
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb[sheet_name]

    # Iterate over each cell in column A
    for row in sheet.iter_rows(min_col=1, max_col=1, min_row=1):
        cell = row[0]
        color_in_hex = cell.fill.start_color.index

        # Convert hex to RGB
        if color_in_hex != '00000000':  # Check if the color is not default
            rgb_color = tuple(int(color_in_hex[i:i+2], 16) for i in (0, 2, 4))
            print(f"Cell {cell.coordinate} - HEX: {color_in_hex}, RGB: {rgb_color}")
        else:
            print(f"Cell {cell.coordinate} - No color")

# Example usage
get_column_a_colors('your_file.xlsx', 'Sheet1')
