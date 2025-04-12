

import openpyxl
from tkinter import filedialog, Tk

# Common Excel tab colors and their friendly names
COLOR_MAP = {
    '#FFFFFF': 'White',
    '#FF0000': 'Red',
    '#00B050': 'Green',
    '#92D050': 'Light Green',
    '#0070C0': 'Blue',
    '#00B0F0': 'Light Blue',
    '#FFFF00': 'Yellow',
    '#FFC000': 'Orange',
    '#7030A0': 'Purple',
    '#D9D9D9': 'Gray',
    '#000000': 'Black',
    '#ED7D31': 'Dark Orange',
    '#A9D08E': 'Pale Green',
    '#F4B084': 'Peach',
    '#FFD966': 'Pale Yellow'
}

def rgb_to_hex_name(rgb):
    """Convert openpyxl RGB color to HEX and friendly name."""
    if rgb is None:
        return "No Color"

    if rgb.type == "rgb":
        hex_color = f"#{rgb.rgb[2:]}"  # Skip "FF" alpha part
        color_name = COLOR_MAP.get(hex_color.upper(), "Custom Color")
        return f"{hex_color} ({color_name})"
    elif rgb.type == "theme":
        return f"Theme Color {rgb.theme} (Tint {rgb.tint})"
    return "Unknown Format"

def get_sheet_tab_colors(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_colors = {}

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        color = sheet.sheet_properties.tabColor
        sheet_colors[sheet_name] = rgb_to_hex_name(color)

    return sheet_colors

# GUI for file selection
def main():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )

    if not file_path:
        print("No file selected.")
        return

    colors = get_sheet_tab_colors(file_path)
    for sheet, color in colors.items():
        print(f"Sheet: {sheet}, Tab Color: {color}")

if __name__ == "__main__":
    main()





Sheet: Töchter Fachbereiche, Tab Color: #92D050------------------------------------------------------------------
import openpyxl
from tkinter import filedialog, Tk

def rgb_to_hex(rgb):
    """Convert RGB to HEX format."""
    if rgb is None:
        return "No Color"
    if rgb.type == "rgb":
        return f"#{rgb.rgb[2:]}"  # openpyxl stores color as 'FFxxxxxx' (ARGB)
    elif rgb.type == "theme":
        return f"Theme Color {rgb.theme} (Tint {rgb.tint})"
    return "Unknown Format"

def get_sheet_tab_colors(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet_colors = {}

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        color = sheet.sheet_properties.tabColor
        sheet_colors[sheet_name] = rgb_to_hex(color)

    return sheet_colors

# GUI for file selection
def main():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )
    
    if not file_path:
        print("No file selected.")
        return

    colors = get_sheet_tab_colors(file_path)
    for sheet, color in colors.items():
        print(f"Sheet: {sheet}, Tab Color: {color}")

if __name__ == "__main__":
    main()

-------------------------------------------------------------------------------------------------------------------------------------------------------

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def get_sheet_tab_colors(file_path):
    wb = load_workbook(file_path, data_only=True)

    sheet_colors = {}
    for sheet in wb.worksheets:
        color = sheet.sheet_properties.tabColor
        if color is not None:
            if color.type == 'rgb':
                sheet_colors[sheet.title] = color.rgb
            elif color.type == 'theme':
                sheet_colors[sheet.title] = f"Theme color index: {color.theme}"
            else:
                sheet_colors[sheet.title] = "Unknown color format"
        else:
            sheet_colors[sheet.title] = "No tab color set"

    return sheet_colors

# Example usage:
file_path = 'your_excel_file.xlsx'  # replace with your file path
sheet_tab_colors = get_sheet_tab_colors(file_path)

for sheet, color in sheet_tab_colors.items():
    print(f"Sheet: {sheet}, Tab Color: {color}")
