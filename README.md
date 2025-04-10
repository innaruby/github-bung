import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime

current_year = datetime.now().year
next_year = current_year + 1
previous_year = current_year - 1
two_years_back = current_year - 2

def is_yellow_or_green(cell):
    if cell.fill.start_color.type == "rgb":
        rgb = cell.fill.start_color.rgb
        if rgb:
            print(f"Cell color RGB: {rgb}")
            return rgb.startswith("FFFF00") or rgb.startswith("FF00")
    return False

def extract_sheet_color(sheet):
    color_check = is_yellow_or_green(sheet["A1"])
    print(f"Sheet '{sheet.title}' color check: {color_check}")
    return color_check

def find_end_row(ws, sheet_name):
    print(f"Finding end row for sheet: {sheet_name}")
    for row in range(7, ws.max_row + 1):
        val = ws[f"A{row}"].value
        if val and "summe" in str(val).lower() and ws[f"A{row}"].font.bold:
            print(f"Found 'Summe' in bold at row {row}")
            return row

    for row in range(7, ws.max_row + 1):
        val = ws[f"A{row}"].value
        if val and str(val).strip().lower() == sheet_name.strip().lower():
            print(f"Found sheet name match at row {row}")
            return row

    for row in range(7, ws.max_row + 1):
        if ws[f"A{row}"].value is None:
            print(f"Found empty cell at row {row}, using {row - 1} as end row")
            return row - 1

    print("No end row logic matched, using last row")
    return ws.max_row

def find_column(ws, keyword, row=3):
    for col in ws.iter_cols(min_row=row, max_row=row):
        if col[0].value and str(col[0].value).strip().lower() == keyword.lower():
            print(f"Found '{keyword}' at column {col[0].column}")
            return col[0].column
    print(f"'{keyword}' not found in row {row}")
    return None

def lookup_and_sum(values, kosten_dict):
    total_c, total_d = 0, 0
    for v in values:
        matched = False
        for key, (c_val, d_val) in kosten_dict.items():
            if v == key or f"/{v}" in key:
                matched = True
                print(f"Matched '{v}' with '{key}': C={c_val}, D={d_val}")
                total_c += c_val
                total_d += d_val
        if not matched:
            print(f"No match found for '{v}'")
    return total_c, total_d

def process_file(file_path, kostenstelle_dict):
    print(f"\n📂 Processing file: {file_path}")
    wb = load_workbook(file_path)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"\n➡️ Sheet: {sheet_name}")
        if not extract_sheet_color(ws):
            print("Skipped: Sheet color not yellow/green.")
            continue

        end_row = find_end_row(ws, sheet_name)
        veraenderung_col = find_column(ws, "Veränderung")
        if veraenderung_col is None:
            print("Skipped: 'Veränderung' column not found.")
            continue

        plan_col = veraenderung_col - 1
        new_ist_col = veraenderung_col + 1

        ws.insert_cols(veraenderung_col, 2)
        print(f"Inserted 2 columns at index {veraenderung_col}")

        ws.cell(row=3, column=plan_col, value="Plan").font = Font(bold=True)
        ws.cell(row=4, column=plan_col, value=str(next_year)).font = Font(bold=True)
        ws.cell(row=3, column=plan_col).alignment = Alignment(horizontal="center")
        ws.cell(row=4, column=plan_col).alignment = Alignment(horizontal="center")
        print(f"Plan column set at {plan_col}")

        ws.cell(row=3, column=new_ist_col, value="IST").font = Font(bold=True)
        ws.cell(row=4, column=new_ist_col, value=f"{current_year}e").font = Font(bold=True)
        ws.cell(row=3, column=new_ist_col).alignment = Alignment(horizontal="center")
        ws.cell(row=4, column=new_ist_col).alignment = Alignment(horizontal="center")
        print(f"IST column set at {new_ist_col}")

        for row in range(5, end_row + 1):
            raw_val = ws[f"AB{row}"].value
            if not raw_val:
                continue
            ids = [x.strip() for x in str(raw_val).replace(",", " ").split()]
            print(f"Row {row} lookup values: {ids}")
            total_c, total_d = lookup_and_sum(ids, kostenstelle_dict)

            ws.cell(row=row, column=plan_col, value=total_d)
            ws.cell(row=row, column=new_ist_col, value=total_c)
            print(f"Row {row}: Plan={total_d}, IST={total_c}")

        print("Hiding unnecessary columns...")
        for col in range(1, ws.max_column + 1):
            hide = True
            if col in [1, plan_col, new_ist_col, veraenderung_col]:
                hide = False
            cell_3 = ws.cell(row=3, column=col).value
            cell_4 = ws.cell(row=4, column=col).value
            if cell_3 == "Plan" and str(cell_4) == str(next_year):
                hide = False
            elif cell_3 == "IST" and (str(cell_4).startswith(str(previous_year)) or str(cell_4).startswith(str(two_years_back))):
                hide = False

            col_letter = ws.cell(row=1, column=col).column_letter
            ws.column_dimensions[col_letter].hidden = hide
            if hide:
                print(f"Hiding column {col_letter}")
            else:
                print(f"Keeping column {col_letter} visible")

    wb.save(file_path)
    print(f"✅ Saved changes to {file_path}")

def load_kostenstelle_data(path):
    print(f"\n🔍 Loading Kostenstelle file: {path}")
    kosten_data = {}
    wb = load_workbook(path)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        key = str(row[0].value)
        val_c = row[2].value if row[2].value else 0
        val_d = row[3].value if row[3].value else 0
        kosten_data[key] = (val_c, val_d)
    print(f"Loaded {len(kosten_data)} entries from Kostenstelle")
    return kosten_data

# -------- MAIN SCRIPT --------
directory = os.getcwd()
kosten_file = next((f for f in os.listdir(directory) if f.startswith("Kostenstelle")), None)
if not kosten_file:
    raise FileNotFoundError("Kostenstelle file not found")

kosten_data = load_kostenstelle_data(os.path.join(directory, kosten_file))

for file in os.listdir(directory):
    if file.endswith(".xlsx") and not file.startswith("Kostenstelle"):
        process_file(os.path.join(directory, file), kosten_data)

print("\n🎉 All files processed successfully!")
