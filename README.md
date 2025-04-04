import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def select_file(title):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])

def process_files(input_path, kostenstelle_path):
    # Step 1: Create a copy of the input file in the same directory
    input_dir = os.path.dirname(input_path)
    input_filename = os.path.basename(input_path)
    copy_path = os.path.join(input_dir, "copy_of_" + input_filename)
    shutil.copy(input_path, copy_path)

    # Step 2: Load workbooks
    original_wb = load_workbook(input_path)
    copy_wb = load_workbook(copy_path)
    kostenstelle_wb = load_workbook(kostenstelle_path)

    orig_ws = original_wb.active
    copy_ws = copy_wb.active
    kosten_ws = kostenstelle_wb.active

    # Step 3: Delete all rows from row 16 onwards in copy file
    max_row = copy_ws.max_row
    for row in range(16, max_row + 1):
        for col in range(1, copy_ws.max_column + 1):
            copy_ws.cell(row=row, column=col).value = None

    # Step 4: Find last non-empty row in column C (index 3) of original file
    end_row = 16
    while orig_ws.cell(row=end_row, column=3).value:
        end_row += 1
    end_row -= 1

    # Step 5: Copy columns C,D,E,F,H from original to copy
    for row in range(16, end_row + 1):
        for col in [3, 4, 5, 6, 8]:
            copy_ws.cell(row=row, column=col).value = orig_ws.cell(row=row, column=col).value

    # Step 6: VLOOKUP to fill column B in copy file
    kosten_dict = {kosten_ws.cell(row=i, column=1).value: kosten_ws.cell(row=i, column=5).value
                   for i in range(2, kosten_ws.max_row + 1)}

    for row in range(16, end_row + 1):
        key = copy_ws.cell(row=row, column=8).value
        copy_ws.cell(row=row, column=2).value = kosten_dict.get(key, None)

    # Step 7: Write in column G based on conditions
    for row in range(16, end_row + 1):
        c_val = str(copy_ws.cell(row=row, column=3).value)
        b_val = copy_ws.cell(row=row, column=2).value
        if c_val.startswith(('705', '706', '707')) and b_val == 1001:
            copy_ws.cell(row=row, column=7).value = "V0"
        elif c_val.startswith(('705', '706', '707')) and b_val == 1002:
            copy_ws.cell(row=row, column=7).value = "U0"
        elif c_val.startswith('704') and b_val == 1001:
            copy_ws.cell(row=row, column=7).value = "A0"
        elif c_val.startswith(('705', '706', '707')) and b_val == 1002:
            copy_ws.cell(row=row, column=7).value = "D0"

    # Step 8: Column L text length and color
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    for row in range(16, end_row + 1):
        d_val = str(copy_ws.cell(row=row, column=4).value or "")
        text_length = len(d_val.replace(" ", ""))
        cell = copy_ws.cell(row=row, column=12)
        cell.value = text_length
        cell.fill = red_fill if text_length >= 50 else green_fill

    # Step 9: Advanced VLOOKUP for column M
    green_cells = [kosten_ws.cell(row=r, column=1).value
                   for r in range(2, kosten_ws.max_row + 1)
                   if kosten_ws.cell(row=r, column=1).fill.start_color.rgb == "FF90EE90"]

    for row in range(16, end_row + 1):
        h_val = copy_ws.cell(row=row, column=8).value
        for r in range(2, kosten_ws.max_row + 1):
            if kosten_ws.cell(row=r, column=1).value == h_val:
                status = str(kosten_ws.cell(row=r, column=6).value).lower()
                if status == 'aktiv':
                    copy_ws.cell(row=row, column=13).value = h_val
                elif status == 'inaktiv':
                    i_val = kosten_ws.cell(row=r, column=9).value
                    copy_ws.cell(row=row, column=13).value = i_val if i_val in green_cells else i_val
                break

    # Step 10: Formula in column K if G is A0 or D0, and delete H
    for row in range(16, end_row + 1):
        g_val = copy_ws.cell(row=row, column=7).value
        if g_val in ["A0", "D0"]:
            h_val = copy_ws.cell(row=row, column=8).value
            if h_val is not None:
                h_str = str(h_val).zfill(3)[-3:]
                k_val = int("100000" + h_str)
                copy_ws.cell(row=row, column=11).value = k_val
                copy_ws.cell(row=row, column=8).value = None

    # Save the copy file
    copy_wb.save(copy_path)
    messagebox.showinfo("Success", f"Processing completed. File saved at: {copy_path}")


# GUI Setup
root = tk.Tk()
root.title("Excel File Processor")

input_file = ""
kosten_file = ""

def select_input():
    global input_file
    input_file = select_file("Select Input File")

def select_kosten():
    global kosten_file
    kosten_file = select_file("Select Kostenstelle File")

def run_process():
    if not input_file or not kosten_file:
        messagebox.showerror("Error", "Please select both files before proceeding.")
    else:
        process_files(input_file, kosten_file)

btn1 = tk.Button(root, text="Select Input File", command=select_input)
btn1.pack(pady=5)

btn2 = tk.Button(root, text="Select Kostenstelle File", command=select_kosten)
btn2.pack(pady=5)

btn3 = tk.Button(root, text="Process Files", command=run_process)
btn3.pack(pady=10)

root.mainloop()
