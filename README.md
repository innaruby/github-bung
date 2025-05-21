# Your full script with the new logic to apply Zwischensumme and Summe logic to 'Sachaufwand'

import os
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from tkinter import Tk, filedialog

# ... (rest of your code remains unchanged)

# Modify the `main` function to include final sum logic for 'Sachaufwand'
def main():
    root = Tk()
    root.withdraw()
    selected_directory = filedialog.askdirectory(title="Select Directory with Excel Files")
    if selected_directory:
        process_excel_files(selected_directory)
        post_processing_with_vlookup(selected_directory)
        final_sum_pass(selected_directory)

        for file in os.listdir(selected_directory):
            if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
                continue
            file_path = os.path.join(selected_directory, file)
            wb = openpyxl.load_workbook(file_path)
            process_sachaufwand_links(wb, file_path) 

            # NEW: Apply Zwischensumme and Summe logic to Sachaufwand if present
            for sheet_name in wb.sheetnames:
                if sheet_name.lower() == 'sachaufwand':
                    ws = wb[sheet_name]
                    end_row = find_end_row(ws, sheet_name)
                    apply_final_sums(ws, end_row)
                    print(f"📘 Zwischensumme and Summe logic applied to 'Sachaufwand' in file: {file}")

            wb.save(file_path)
            print(f"💾 Final update (Sachaufwand) saved in file: {file}")

if __name__ == "__main__":
    main()
