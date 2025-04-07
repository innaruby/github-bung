import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from openpyxl.styles import PatternFill

# Define color fills
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill = PatternFill(start_color="FF90EE90", end_color="FF90EE90", fill_type="solid")

def browse_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

# Process column M logic
def process_column_m(row_num, copy_ws, kostenstelle_data, kostenstelle_a_styles):
    print(f"\n[DEBUG] Processing Column M for row {row_num}")
    h_val = copy_ws[f"H{row_num}"].value
    print(f"[DEBUG] H{row_num} value: {h_val}")

    if h_val not in kostenstelle_data:
        print(f"[DEBUG] No matching H value found in kostenstelle_data for H{row_num}")
        return

    k_data = kostenstelle_data[h_val]
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
            is_green = kostenstelle_a_styles.get(i_val, False)
            if is_green:
                matched_value = k_data["I"]
                copy_ws[f"M{row_num}"].value = matched_value
                print(f"[DEBUG] Green fill detected for I='{i_val}'. Writing: '{matched_value}' to M{row_num}")
            else:
                copy_ws[f"M{row_num}"].value = i_val
                print(f"[DEBUG] No green fill for I='{i_val}'. Writing: '{i_val}' to M{row_num}")
    else:
        print(f"[DEBUG] F value is not a string or is missing for H{row_num}")

def process_files(input_path, kostenstelle_path):
    input_dir = os.path.dirname(input_path)
    input_wb = openpyxl.load_workbook(input_path)
    input_ws = input_wb.active

    copy_path = os.path.join(input_dir, "copy_of_input.xlsx")
    input_wb.save(copy_path)

    copy_wb = openpyxl.load_workbook(copy_path)
    copy_ws = copy_wb.active

    row = 16
    while input_ws[f"C{row}"].value:
        row += 1
    end_row = row - 1
    print(f"Detected end row: {end_row}")

    for r in range(16, end_row + 1):
        copy_ws[f"B{r}"].value = None
        for col in ["C", "D", "E", "F", "H"]:
            copy_ws[f"{col}{r}"].value = input_ws[f"{col}{r}"].value

    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=True)
    kostenstelle_ws = kostenstelle_wb.active

    kostenstelle_data = {}
    for row in kostenstelle_ws.iter_rows(min_row=2):
        a_val = row[0].value
        kostenstelle_data[a_val] = {
            "E": row[4].value,
            "F": row[5].value,
            "I": row[8].value
        }

    for r in range(16, end_row + 1):
        h_val = copy_ws[f"H{r}"].value
        if h_val in kostenstelle_data:
            k_data = kostenstelle_data[h_val]
            b_val = k_data["E"]
            copy_ws[f"B{r}"].value = b_val

            c_val = str(copy_ws[f"C{r}"].value)
            if c_val.startswith(("705", "706", "707")) and b_val == 1001:
                copy_ws[f"G{r}"].value = "V0"
            elif c_val.startswith(("705", "706", "707")) and b_val == 1002:
                copy_ws[f"G{r}"].value = "U0"
            elif c_val.startswith("704") and b_val == 1001:
                copy_ws[f"G{r}"].value = "A0"
            elif c_val.startswith(("705", "706", "707")) and b_val == 1002:
                copy_ws[f"G{r}"].value = "D0"

            d_val = copy_ws[f"D{r}"].value
            if d_val:
                length = len(str(d_val).replace(" ", ""))
                copy_ws[f"L{r}"].value = length
                copy_ws[f"L{r}"].fill = red_fill if length >= 50 else green_fill

            g_val = copy_ws[f"G{r}"].value
            if g_val in ["A0", "D0"]:
                h_val = copy_ws[f"H{r}"].value
                if h_val:
                    copy_ws[f"K{r}"].value = int("100000" + str(h_val)[-3:])
                    copy_ws[f"H{r}"].value = None

    copy_wb.save(copy_path)
    input_wb.close()
    copy_wb.close()
    kostenstelle_wb.close()

    # Step 2: capture fill color from column A (data_only=False)
    print("\n[DEBUG] Loading styles from column A in 'Kostenstelle'...")

    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=False)
    kostenstelle_ws = kostenstelle_wb.active
    kostenstelle_a_styles = {}

    for row_idx, row in enumerate(kostenstelle_ws.iter_rows(min_row=2), start=2):
        a_cell = row[0]
        a_val = a_cell.value
        fill = a_cell.fill
        is_green = (
            isinstance(fill, PatternFill)
            and fill.fill_type == "solid"
            and fill.start_color.rgb == "FF90EE90"
        )
        kostenstelle_a_styles[a_val] = is_green
        print(f"[DEBUG] Row {row_idx}: A={a_val}, Green Fill={is_green}")

    kostenstelle_wb.close()

    # Step 3: reopen for final processing with data_only=True
    copy_wb = openpyxl.load_workbook(copy_path, data_only=True)
    copy_ws = copy_wb.active
    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=True)
    kostenstelle_ws = kostenstelle_wb.active

    kostenstelle_data = {}
    for row in kostenstelle_ws.iter_rows(min_row=2):
        a_val = row[0].value
        kostenstelle_data[a_val] = {
            "E": row[4].value,
            "F": row[5].value,
            "I": row[8].value
        }

    for r in range(16, end_row + 1):
        process_column_m(r, copy_ws, kostenstelle_data, kostenstelle_a_styles)

    copy_wb.save(copy_path)
    copy_wb.close()
    kostenstelle_wb.close()

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
