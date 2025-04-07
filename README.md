import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from openpyxl.styles import PatternFill

# Define color fills
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

def browse_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def process_files(input_path, kostenstelle_path):
    input_dir = os.path.dirname(input_path)
    
    # Open both files with data_only=True
    input_wb = openpyxl.load_workbook(input_path, data_only=True)
    input_ws = input_wb.active

    # Create a copy
    copy_path = os.path.join(input_dir, "copy_of_input.xlsx")
    input_wb.save(copy_path)
    input_wb.close()

    copy_wb = openpyxl.load_workbook(copy_path, data_only=True)
    copy_ws = copy_wb.active

    # Load kostenstelle with data_only=True
    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=True)
    kostenstelle_ws = kostenstelle_wb.active

    # Find end row in original
    row = 16
    while input_ws[f"C{row}"].value:
        row += 1
    end_row = row - 1
    print(f"Detected end row: {end_row}")

    # Clear column B values in copy file from row 16 to end_row
    for r in range(16, end_row + 1):
        copy_ws[f"B{r}"].value = None

    # Copy columns C,D,E,F,H to copy file
    for r in range(16, end_row + 1):
        for col in ["C", "D", "E", "F", "H"]:
            copy_ws[f"{col}{r}"].value = input_ws[f"{col}{r}"].value

    # Create lookup and identify green cell values in column A
    kostenstelle_data = {}
    green_a_values = set()
    green_i_map = {}
    for i, row in enumerate(kostenstelle_ws.iter_rows(min_row=2), start=2):
        a_val = row[0].value
        fill = row[0].fill
        fill_color = fill.start_color.rgb if fill.fill_type == "solid" else None
        print(f"Row {i} in Kostenstelle: A={a_val}, Fill={fill_color}")
        if fill_color == "FF90EE90":
            green_a_values.add(a_val)
            green_i_map[a_val] = row[8].value  # Map green A to its corresponding I value
            print(f"  -> Registered green A cell: {a_val} with I={row[8].value}")
        kostenstelle_data[a_val] = {
            "E": row[4].value,
            "F": row[5].value,
            "I": row[8].value,
        }

    for r in range(16, end_row + 1):
        h_val = copy_ws[f"H{r}"].value
        print(f"Row {r}: H = {h_val}")
        if h_val in kostenstelle_data:
            k_data = kostenstelle_data[h_val]
            copy_ws[f"B{r}"].value = k_data["E"]
            b_val = k_data["E"]
            print(f"  Match found in Kostenstelle - B = {b_val}, F = {k_data['F']}, I = {k_data['I']}")

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

    # Save and close the workbooks
    copy_wb.save(copy_path)
    copy_wb.close()
    kostenstelle_wb.close()

    # Reopen both files with data_only=False for processing column M
    copy_wb = openpyxl.load_workbook(copy_path, data_only=False)
    copy_ws = copy_wb.active
    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=False)
    kostenstelle_ws = kostenstelle_wb.active

    # Process column M
    for r in range(16, end_row + 1):
        h_val = copy_ws[f"H{r}"].value
        if h_val in kostenstelle_data:
            k_data = kostenstelle_data[h_val]

            # Column M logic
            f_val = k_data["F"]
            if isinstance(f_val, str):
                if f_val.lower() == "aktiv":
                    copy_ws[f"M{r}"].value = "okay"
                    print(f"  Writing 'okay' to M{r}")
                elif f_val.lower() == "inaktiv":
                    i_val = k_data["I"]
                    print(f"  Inaktiv detected. I = {i_val}. Green A values = {green_a_values}")
                    if i_val in green_a_values:
                        matched_value = green_i_map.get(i_val)
                        copy_ws[f"M{r}"].value = matched_value
                        print(f"  Writing '{matched_value}' to M{r} (from green A match)")
                    else:
                        copy_ws[f"M{r}"].value = i_val
                        print(f"  Writing '{i_val}' to M{r} from column I")

    # Save the final changes
    copy_wb.save(copy_path)
    messagebox.showinfo("Success", f"File processed and saved as {copy_path}")

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
