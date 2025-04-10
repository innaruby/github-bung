import os
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# SETTINGS
CURRENT_YEAR = datetime.now().year
NEXT_YEAR = CURRENT_YEAR + 1
PREVIOUS_YEAR = CURRENT_YEAR - 1
TWO_YEARS_AGO = CURRENT_YEAR - 2

def is_yellow_or_green(rgb):
    if rgb is None:
        return False
    rgb = rgb.replace("FF", "")  # Remove alpha
    # Check for yellows and greens (very basic ranges)
    return (
        rgb.startswith("FF") or rgb.startswith("FFFF") or rgb.startswith("FFFF00") or  # Yellows
        rgb.startswith("00FF00") or rgb.startswith("008000") or rgb.startswith("ADFF2F")  # Greens
    )

def find_end_row(sheet, sheet_name):
    for row in range(7, sheet.max_row + 1):
        cell = sheet[f"A{row}"]
        if cell.value and isinstance(cell.value, str) and "summe" in cell.value.lower() and cell.font and cell.font.bold:
            return row
    for row in range(7, sheet.max_row + 1):
        if str(sheet[f"A{row}"].value).strip().lower() == sheet_name.lower():
            return row
    for row in range(7, sheet.max_row + 1):
        if sheet[f"A{row}"].value is None:
            return row - 1
    return sheet.max_row

def fuzzy_lookup(values, kostenstelle_df, column):
    total = 0
    for val in values:
        matched = kostenstelle_df[kostenstelle_df['A'].str.contains(val, na=False, case=False)]
        if not matched.empty:
            total += matched.iloc[0][column]
    return total

def get_kostenstelle_df(kosten_file):
    df = pd.read_excel(kosten_file, header=None)
    df.columns = ['A', 'B', 'C', 'D']
    df['A'] = df['A'].astype(str)
    return df

def main():
    directory = os.getcwd()
    files = [f for f in os.listdir(directory) if f.endswith(".xlsx") and not f.startswith("~$")]
    kosten_file = next((f for f in files if f.startswith("Kostenstelle")), None)

    if not kosten_file:
        print("Kostenstelle file not found.")
        return

    kosten_path = os.path.join(directory, kosten_file)
    kosten_df = get_kostenstelle_df(kosten_path)

    for file in files:
        if file.startswith("Kostenstelle"):
            continue

        filepath = os.path.join(directory, file)
        wb = load_workbook(filepath)
        for sheetname in wb.sheetnames:
            sheet = wb[sheetname]

            sheet_name_cell = sheet["A1"]
            if not is_yellow_or_green(sheet_name_cell.fill.start_color.rgb):
                continue

            end_row = find_end_row(sheet, sheetname)
            var_col = None
            for col in range(1, sheet.max_column + 1):
                val = str(sheet.cell(row=3, column=col).value).lower()
                if "veränderung" in val:
                    var_col = col
                    break
            if not var_col:
                continue

            # Insert two columns to the left
            sheet.insert_cols(var_col, amount=2)
            left_col1 = var_col
            left_col2 = var_col + 1
            center_align = Alignment(horizontal='center')

            # Write "Plan" and next year
            sheet.cell(row=3, column=left_col1).value = "Plan"
            sheet.cell(row=3, column=left_col1).font = Font(bold=True)
            sheet.cell(row=3, column=left_col1).alignment = center_align
            sheet.cell(row=4, column=left_col1).value = NEXT_YEAR
            sheet.cell(row=4, column=left_col1).font = Font(bold=True)
            sheet.cell(row=4, column=left_col1).alignment = center_align

            # Write "IST" and current year + "e"
            sheet.cell(row=3, column=left_col2).value = "IST"
            sheet.cell(row=3, column=left_col2).font = Font(bold=True)
            sheet.cell(row=3, column=left_col2).alignment = center_align
            sheet.cell(row=4, column=left_col2).value = f"{CURRENT_YEAR}e"
            sheet.cell(row=4, column=left_col2).font = Font(bold=True)
            sheet.cell(row=4, column=left_col2).alignment = center_align

            for row in range(5, end_row + 1):
                ab_val = sheet.cell(row=row, column=28).value  # AB column is index 28
                if not ab_val:
                    continue
                ids = re.findall(r'\w+', str(ab_val))

                # Lookup for PLAN (left_col1) from D column in Kostenstelle
                val_d = fuzzy_lookup(ids, kosten_df, "D")
                sheet.cell(row=row, column=left_col1).value = val_d

                # Lookup for IST (left_col2) from C column in Kostenstelle
                val_c = fuzzy_lookup(ids, kosten_df, "C")
                sheet.cell(row=row, column=left_col2).value = val_c

            keep_cols = {1, left_col1, left_col2, var_col + 2}  # var_col+2 because two were inserted before

            # Unhide Plan+NextYear, IST+CurrentYear, IST+PreviousYear, IST+2YearsAgo
            for col in range(1, sheet.max_column + 1):
                r3 = str(sheet.cell(row=3, column=col).value).lower()
                r4 = str(sheet.cell(row=4, column=col).value).lower()
                if ("plan" in r3 and str(NEXT_YEAR) in r4) or \
                   ("ist" in r3 and str(CURRENT_YEAR) in r4) or \
                   ("ist" in r3 and str(PREVIOUS_YEAR) in r4) or \
                   ("ist" in r3 and str(TWO_YEARS_AGO) in r4):
                    keep_cols.add(col)

            # Hide other columns
            for col in range(1, sheet.max_column + 1):
                if col not in keep_cols:
                    sheet.column_dimensions[get_column_letter(col)].hidden = True

        wb.save(filepath)
        print(f"Processed: {file}")

if __name__ == "__main__":
    main()
