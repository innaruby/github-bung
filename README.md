import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from openpyxl.styles import PatternFill

# Acceptable green fill variations
GREEN_HEX_CODES = {"FF90EE90", "FF92D050", "FF00FF00"}

# Define color fills
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill = PatternFill(start_color="FF90EE90", end_color="FF90EE90", fill_type="solid")


def browse_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def get_column_a_colors(file_path, sheet_name='Sheet1'):
    wb = openpyxl.load_workbook(file_path, data_only=False)
    sheet = wb[sheet_name]

    column_a_colors = {}
    green_i_lookup = {}
    unique_colors = set()

    for row in sheet.iter_rows(min_row=2):
        a_cell = row[0]
        i_cell = row[8]
        a_val = a_cell.value
        i_val = i_cell.value
        fill = a_cell.fill

        if fill and fill.fill_type == "solid":
            color = fill.start_color
            color_hex = color.rgb if color.type == 'rgb' else color.index
        else:
            color_hex = None

        column_a_colors[a_val] = color_hex

        if color_hex in GREEN_HEX_CODES:
            green_i_lookup[a_val] = i_val

        if color_hex:
            unique_colors.add(color_hex)
            print(f"[DEBUG] Cell {a_cell.coordinate}: A={a_val}, HEX={color_hex}")
        else:
            print(f"[DEBUG] Cell {a_cell.coordinate}: A={a_val}, No fill")

    wb.close()
    print(f"\n[DEBUG] Unique fill colors in column A: {unique_colors}")
    return column_a_colors, green_i_lookup


def process_column_m(row_num, copy_ws, kostenstelle_data, column_a_colors, green_i_lookup):
    print(f"\n[DEBUG] Processing Column M for row {row_num}")
    h_val = copy_ws[f"H{row_num}"].value
    print(f"[DEBUG] H{row_num} value: {h_val}")

    if h_val not in kostenstelle_data:
        print(f"[DEBUG] No match in kostenstelle_data for H{row_num}")
        return

    k_data = kostenstelle_data[h_val]
    f_val = k_data.get("F")
    i_val = k_data.get("I")
    b_val = k_data.get("E")
    print(f"[DEBUG] F = {f_val}, I = {i_val}, B = {b_val}")

    vertriebsberichte = {
        "Vertriebsbericht Nürnberg", "Vertriebsbericht Regensburg", "Vertriebsbericht Sondervolumen Markt SüdD",
        "Vertriebsbericht Würzburg", "Vertriebsbericht München", "Vertriebsbericht Ulm",
        "Vertriebsbericht Heilbronn", "Vertriebsbericht Stuttgart", "Vertriebsbericht Augsburg"
    }

    if isinstance(f_val, str):
        f_val_lower = f_val.lower()
        if f_val_lower == "aktiv":
            copy_ws[f"M{row_num}"].value = "okay"
            print(f"[DEBUG] F is 'aktiv' → Writing 'okay' to M{row_num}")
        elif f_val_lower == "inaktiv":
            fill_color = column_a_colors.get(i_val)
            print(f"[DEBUG] Fill color for A={i_val}: {fill_color}")

            if i_val in green_i_lookup:
                matched_i = green_i_lookup[i_val]
                copy_ws[f"M{row_num}"].value = matched_i
                print(f"[DEBUG] Green A match found → Writing matched I='{matched_i}' to M{row_num}")
            elif b_val in vertriebsberichte:
                copy_ws[f"M{row_num}"].value = i_val
                print(f"[DEBUG] B is Vertriebsbericht match → Writing I='{i_val}' to M{row_num}")
            else:
                copy_ws[f"M{row_num}"].value = i_val
                print(f"[DEBUG] No green A match or B override → Writing I='{i_val}' to M{row_num}")
    else:
        print(f"[DEBUG] Invalid or missing F for row {row_num}")


def process_files(input_path, kostenstelle_path):
    input_dir = os.path.dirname(input_path)
    input_wb = openpyxl.load_workbook(input_path)
    input_ws = input_wb.active

    copy_path = os.path.join(input_dir, "Kopie.xlsx")
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
            elif c_val.startswith("704") and b_val == 1002:
                copy_ws[f"G{r}"].value = "D0"

            d_val = copy_ws[f"D{r}"].value
            if d_val:
                length = len(str(d_val).replace(" ", ""))
                copy_ws[f"L{r}"].value = length
                if length >= 50:
                    copy_ws[f"L{r}"].fill = red_fill  # Only fill red if >= 50

            g_val = copy_ws[f"G{r}"].value
            if g_val in ["A0", "D0"]:
                h_cell_val = copy_ws[f"H{r}"].value
                if h_cell_val:
                    copy_ws[f"K{r}"].value = int("100000" + str(h_cell_val)[-3:])
                    copy_ws[f"H{r}"].value = None

    input_wb.close()
    copy_wb.save(copy_path)
    copy_wb.close()
    kostenstelle_wb.close()

    # Phase 2: Process column M
    print("\n[DEBUG] Reading fill colors from column A...")
    column_a_colors, green_i_lookup = get_column_a_colors(kostenstelle_path)

    copy_wb = openpyxl.load_workbook(copy_path)
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
        process_column_m(r, copy_ws, kostenstelle_data, column_a_colors, green_i_lookup)

    # Add final values and sums
    sum_e = sum(copy_ws[f"E{r}"].value for r in range(16, end_row + 1) if copy_ws[f"E{r}"].value is not None)
    sum_f = sum(copy_ws[f"F{r}"].value for r in range(16, end_row + 1) if copy_ws[f"F{r}"].value is not None)

    copy_ws.cell(row=12, column=5).value = sum_e
    copy_ws.cell(row=12, column=6).value = sum_f
    copy_ws.cell(row=12, column=4).value = '=E12-F12'

    copy_ws.cell(row=15, column=13).value = "Kontrolle alte KST"
    copy_ws.cell(row=15, column=12).value = "Kontrolle Länge Positionstext"

    # Apply fill to M if value is numeric
    for r in range(16, end_row + 1):
        m_cell = copy_ws[f"M{r}"]
        try:
            if isinstance(m_cell.value, (int, float)):
                m_cell.fill = green_fill
                print(f"[DEBUG] M{r} is numeric → Applied green fill")
        except:
            continue

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
