import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from openpyxl.styles import PatternFill

# Define color fills
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill = PatternFill(start_color="FF90EE90", end_color="FF90EE90", fill_type="solid")

# GUI callback to browse and store file paths
def browse_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

# Function to process Column M with debug statements
def process_column_m(row_num, copy_ws, kostenstelle_data, green_a_values, green_i_map):
    print(f"\n[DEBUG] Processing Column M for row {row_num}")
    h_val = copy_ws[f"H{row_num}"].value
    print(f"[DEBUG] H{row_num} value: {h_val}")

    if h_val not in kostenstelle_data:
        print(f"[DEBUG] No matching H value found in kostenstelle_data for H{row_num}")
        return

    k_data = kostenstelle_data[h_val]
    print(f"[DEBUG] Retrieved kostenstelle data: {k_data}")

    f_val = k_data.get("F")
    i_val = k_data.get("I")
    print(f"[DEBUG] F = {f_val}, I = {i_val}")

    if isinstance(f_val, str):
        f_val_lower = f_val.lower()
        print(f"[DEBUG] Interpreted F (lower): {f_val_lower}")
        if f_val_lower == "aktiv":
            copy_ws[f"M{row_num}"].value = "okay"
            print(f"[DEBUG] Writing 'okay' to M{row_num}")
        elif f_val_lower == "inaktiv":
            if i_val in green_a_values:
                matched_value = green_i_map.get(i_val)
                copy_ws[f"M{row_num}"].value = matched_value
                print(f"[DEBUG] I value '{i_val}' found in green A set. Writing matched I: '{matched_value}' to M{row_num}")
            else:
                copy_ws[f"M{row_num}"].value = i_val
                print(f"[DEBUG] I value '{i_val}' not found in green A set. Writing raw I: '{i_val}' to M{row_num}")
    else:
        print(f"[DEBUG] F value is not a string or is missing for H{row_num}")

# Main file processing function
def process_files(input_path, kostenstelle_path):
    input_dir = os.path.dirname(input_path)
    input_wb = openpyxl.load_workbook(input_path)
    input_ws = input_wb.active

    # Create a copy of the input
    copy_path = os.path.join(input_dir, "copy_of_input.xlsx")
    input_wb.save(copy_path)

    copy_wb = openpyxl.load_workbook(copy_path)
    copy_ws = copy_wb.active

    # Find end row
    row = 16
    while input_ws[f"C{row}"].value:
        row += 1
    end_row = row - 1
    print(f"Detected end row: {end_row}")

    # Clear column B values
    for r in range(16, end_row + 1):
        copy_ws[f"B{r}"].value = None

    # Copy C,D,E,F,H from input to copy
    for r in range(16, end_row + 1):
        for col in ["C", "D", "E", "F", "H"]:
            copy_ws[f"{col}{r}"].value = input_ws[f"{col}{r}"].value

    # Load kostenstelle with data_only=True for first pass processing
    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=True)
    kostenstelle_ws = kostenstelle_wb.active

    # Prepare lookup for kostenstelle_data and green cell properties
    kostenstelle_data = {}
    green_a_values = set()
    green_i_map = {}
    column_a_cell_properties = {}

    # Store cell properties (fill color) of cells in column A
    for i, row in enumerate(kostenstelle_ws.iter_rows(min_row=2), start=2):
        a_val = row[0].value
        fill = row[0].fill
        fill_color = fill.start_color.rgb if isinstance(fill, PatternFill) and fill.fill_type == "solid" else None
        column_a_cell_properties[a_val] = fill_color  # Store cell color in column A
        print(f"[DEBUG] A{a_val} fill color: {fill_color}")

        if fill_color == "FF90EE90":  # Green fill
            green_a_values.add(a_val)
            green_i_map[a_val] = row[8].value
        kostenstelle_data[a_val] = {
            "E": row[4].value,
            "F": row[5].value,
            "I": row[8].value,
        }

    # First pass (process B, G, L, K)
    for r in range(16, end_row + 1):
        h_val = copy_ws[f"H{r}"].value
        if h_val in kostenstelle_data:
            k_data = kostenstelle_data[h_val]
            b_val = k_data["E"]
            copy_ws[f"B{r}"].value = b_val

            # G logic
            c_val = str(copy_ws[f"C{r}"].value)
            if c_val.startswith(("705", "706", "707")) and b_val == 1001:
                copy_ws[f"G{r}"].value = "V0"
            elif c_val.startswith(("705", "706", "707")) and b_val == 1002:
                copy_ws[f"G{r}"].value = "U0"
            elif c_val.startswith("704") and b_val == 1001:
                copy_ws[f"G{r}"].value = "A0"
            elif c_val.startswith(("705", "706", "707")) and b_val == 1002:
                copy_ws[f"G{r}"].value = "D0"

            # L logic
            d_val = copy_ws[f"D{r}"].value
            if d_val:
                length = len(str(d_val).replace(" ", ""))
                copy_ws[f"L{r}"].value = length
                copy_ws[f"L{r}"].fill = red_fill if length >= 50 else green_fill

            # K logic
            g_val = copy_ws[f"G{r}"].value
            if g_val in ["A0", "D0"]:
                h_val = copy_ws[f"H{r}"].value
                if h_val:
                    copy_ws[f"K{r}"].value = int("100000" + str(h_val)[-3:])
                    copy_ws[f"H{r}"].value = None

    # Save and close before processing column M
    copy_wb.save(copy_path)
    input_wb.close()
    copy_wb.close()
    kostenstelle_wb.close()

    # Open kostenstelle file again, this time with data_only=False, to read cell properties
    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=False)
    kostenstelle_ws = kostenstelle_wb.active

    # Rebuild the kostenstelle_data and the green cells properties with data_only=False
    kostenstelle_data = {}
    green_a_values = set()
    green_i_map = {}

    # Store the cell properties in column A
    for i, row in enumerate(kostenstelle_ws.iter_rows(min_row=2), start=2):
        a_val = row[0].value
        fill = row[0].fill
        fill_color = fill.start_color.rgb if isinstance(fill, PatternFill) and fill.fill_type == "solid" else None
        column_a_cell_properties[a_val] = fill_color  # Store cell color in column A
        print(f"[DEBUG] A{a_val} (second pass) fill color: {fill_color}")

        if fill_color == "FF90EE90":  # Green fill
            green_a_values.add(a_val)
            green_i_map[a_val] = row[8].value
        kostenstelle_data[a_val] = {
            "E": row[4].value,
            "F": row[5].value,
            "I": row[8].value,
        }

    # Open copy file with data_only=True for the final processing of column M
    copy_wb = openpyxl.load_workbook(copy_path, data_only=True)
    copy_ws = copy_wb.active

    # Final pass: process column M
    for r in range(16, end_row + 1):
        process_column_m(r, copy_ws, kostenstelle_data, green_a_values, green_i_map)

    # Final save and close
    copy_wb.save(copy_path)
    copy_wb.close()
    kostenstelle_wb.close()

    messagebox.showinfo("Success", f"File processed and saved as {copy_path}")

# GUI setup
def main():
    root = tk.Tk()
    root.title("Excel Processing GUI")

    tk.Label(root, text="Input File (original)").grid(row=0, column=0, padx=10, pady=10)
    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1)
    tk.Button(root, text="Browse", command=lambda: browse_file(input_entry)).grid(row=0, column=2)

    tk.Label(root, text="Kostenstelle File").grid(row=1, column=0, padx=10, pady=10)
    kostenstelle_entry = tk.Entry(root, width=50)
    kostenstelle_entry.grid(row=1, column=1)
    tk.Button(root, text="Browse", command=lambda: browse_file(kostenstelle_entry)).grid(row=1, column=2)

    def run_process():
        input_path = input_entry.get()
        kostenstelle_path = kostenstelle_entry.get()
        if not input_path or not kostenstelle_path:
            messagebox.showerror("Error", "Both file paths are required!")
            return
        process_files(input_path, kostenstelle_path)

    tk.Button(root, text="Process Files", command=run_process, bg="lightblue").grid(row=2, column=1, pady=20)
    root.mainloop()

if __name__ == "__main__":
    main()
