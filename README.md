import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from openpyxl.styles import PatternFill

# Define color fills
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

def select_file(title):
    return filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])

def process_files(input_path, kostenstelle_path):
    input_dir = os.path.dirname(input_path)
    input_wb = openpyxl.load_workbook(input_path)
    input_ws = input_wb.active

    # Create a copy
    copy_path = os.path.join(input_dir, "copy_of_input.xlsx")
    input_wb.save(copy_path)

    copy_wb = openpyxl.load_workbook(copy_path)
    copy_ws = copy_wb.active

    # Find end row in original
    row = 16
    while input_ws[f"C{row}"].value:
        row += 1
    end_row = row - 1

    # Clear column B values in copy file from row 16 to end_row
    for r in range(16, end_row + 1):
        copy_ws[f"B{r}"].value = None

    # Copy columns C,D,E,F,H to copy file
    for r in range(16, end_row + 1):
        for col in ["C", "D", "E", "F", "H"]:
            copy_ws[f"{col}{r}"].value = input_ws[f"{col}{r}"].value

    # Load kostenstelle
    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path)
    kostenstelle_ws = kostenstelle_wb.active
    kostenstelle_data = {
        row[0].value: {
            "E": row[4].value,
            "F": row[5].value,
            "I": row[8].value,
            "A_fill": row[0].fill.start_color.rgb if isinstance(row[0].fill, PatternFill) else None
        }
        for row in kostenstelle_ws.iter_rows(min_row=2)
    }

    for r in range(16, end_row + 1):
        h_val = copy_ws[f"H{r}"].value
        if h_val in kostenstelle_data:
            k_data = kostenstelle_data[h_val]
            copy_ws[f"B{r}"].value = k_data["E"]
            b_val = k_data["E"]

            # G value logic
            c_val = str(copy_ws[f"C{r}"].value)
            if c_val.startswith(("705", "706", "707")) and b_val == 1001:
                copy_ws[f"G{r}"].value = "V0"
            elif c_val.startswith(("705", "706", "707")) and b_val == 1002:
                copy_ws[f"G{r}"].value = "U0"
            elif c_val.startswith("704") and b_val == 1001:
                copy_ws[f"G{r}"].value = "A0"
            elif c_val.startswith(("705", "706", "707")) and b_val == 1002:
                copy_ws[f"G{r}"].value = "D0"

            # Column L: Text length of D (spaces excluded)
            d_val = copy_ws[f"D{r}"].value
            if d_val:
                length = len(str(d_val).replace(" ", ""))
                copy_ws[f"L{r}"].value = length
                copy_ws[f"L{r}"].fill = red_fill if length >= 50 else green_fill

            # Column M logic
            f_val = k_data["F"]
            if isinstance(f_val, str) and f_val.lower() == "aktiv":
                copy_ws[f"M{r}"].value = "okay"
            elif isinstance(f_val, str) and f_val.lower() == "inaktiv":
                i_val = k_data["I"]
                copy_ws[f"M{r}"].value = i_val

            # Column K formula logic
            g_val = copy_ws[f"G{r}"].value
            if g_val in ["A0", "D0"]:
                h_cell_val = copy_ws[f"H{r}"].value
                if h_cell_val:
                    copy_ws[f"K{r}"].value = int("100000" + str(h_cell_val)[-3:])
                    copy_ws[f"H{r}"].value = None

    copy_wb.save(copy_path)
    messagebox.showinfo("Success", f"File processed and saved as {copy_path}")

def main():
    root = tk.Tk()
    root.withdraw()

    input_path = select_file("Select Input File")
    if not input_path:
        return

    kostenstelle_path = select_file("Select Kostenstelle File")
    if not kostenstelle_path:
        return

    process_files(input_path, kostenstelle_path)

if __name__ == "__main__":
    main()
