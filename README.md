import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from openpyxl.styles import PatternFill
import win32com.client

# Define color fills
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill = PatternFill(start_color="FF90EE90", end_color="FF90EE90", fill_type="solid")

# GUI callback to browse and store file paths
def browse_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def get_green_cells_using_win32(kostenstelle_path):
    green_cells = {}
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(kostenstelle_path))
    ws = wb.Sheets(1)

    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
    for row in range(2, last_row + 1):
        cell = ws.Cells(row, 1)
        color = cell.Interior.Color  # BGR int
        blue = color // 65536
        green = (color % 65536) // 256
        red = color % 256
        rgb = (red, green, blue)
        print(f"Row {row}: A = {cell.Value}, Color = {rgb}")
        if rgb == (144, 238, 144):  # Light green
            green_cells[cell.Value] = ws.Cells(row, 9).Value  # Column I
            print(f"  -> win32 detected green: A{row}={cell.Value}, I={ws.Cells(row, 9).Value}")

    wb.Close(SaveChanges=0)
    excel.Quit()
    return green_cells

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
    print(f"Detected end row: {end_row}")

    # Clear column B values in copy file from row 16 to end_row
    for r in range(16, end_row + 1):
        copy_ws[f"B{r}"].value = None

    # Copy columns C,D,E,F,H to copy file
    for r in range(16, end_row + 1):
        for col in ["C", "D", "E", "F", "H"]:
            copy_ws[f"{col}{r}"].value = input_ws[f"{col}{r}"].value

    # Load kostenstelle
    kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=True)
    kostenstelle_ws = kostenstelle_wb.active

    # Get green cells using win32
    green_i_map = get_green_cells_using_win32(kostenstelle_path)
    green_a_values = set(green_i_map.keys())

    # Create kostenstelle data
    kostenstelle_data = {}
    for row in kostenstelle_ws.iter_rows(min_row=2):
        a_val = row[0].value
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
