import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime

def is_yellow_or_green(cell):
    fill = cell.fill.start_color.rgb
    return fill and (fill.startswith('FFFFFF00') or fill.startswith('FF00FF00') or fill.startswith('FFFFFF99') or fill.startswith('FFFFE135'))

def get_kostenstelle_data(kosten_file_path):
    wb = load_workbook(kosten_file_path, data_only=True)
    ws = wb.active
    lookup_dict_c = {}
    lookup_dict_d = {}
    for row in ws.iter_rows(min_row=2, values_only=True):  # assuming header row
        key = str(row[0]).strip() if row[0] else ''
        if key:
            lookup_dict_c[key] = row[2]  # Column C
            lookup_dict_d[key] = row[3]  # Column D
    return lookup_dict_c, lookup_dict_d

def fuzzy_match(key, lookup_keys):
    for ref_key in lookup_keys:
        if key == ref_key or key in ref_key or ref_key in key:
            return ref_key
    return None

def process_workbook(file_path, kosten_data_c, kosten_data_d, current_year):
    wb = load_workbook(file_path)
    for sheet in wb.worksheets:
        cell = sheet['A1']
        if not is_yellow_or_green(cell):
            continue

        sheet_name = sheet.title.strip().lower()
        end_row = None
        for row in range(7, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=1)
            if cell.value and str(cell.value).strip().lower() == 'summe' and cell.font.bold:
                end_row = row
                break
        if not end_row:
            for row in range(7, sheet.max_row + 1):
                if str(sheet.cell(row=row, column=1).value).strip().lower() == sheet_name:
                    end_row = row
                    break
        if not end_row:
            for row in range(7, sheet.max_row + 1):
                if sheet.cell(row=row, column=1).value is None:
                    end_row = row - 1
                    break
        if not end_row:
            end_row = sheet.max_row

        # Step 2: Find “Veränderung”
        ver_col = None
        for col in range(1, sheet.max_column + 1):
            if str(sheet.cell(row=3, column=col).value).strip().lower() == 'veränderung':
                ver_col = col
                break
            if str(sheet.cell(row=4, column=col).value).strip().lower() == 'veränderung':
                ver_col = col
                break
        if not ver_col:
            continue

        sheet.insert_cols(ver_col)
        sheet.insert_cols(ver_col)

        # Header formatting
        plan_col = ver_col
        ist_col = ver_col + 1
        plan_cell = sheet.cell(row=3, column=plan_col)
        ist_cell = sheet.cell(row=3, column=ist_col)
        plan_cell.value = 'Plan'
        plan_cell.font = Font(bold=True)
        plan_cell.alignment = Alignment(horizontal='center')

        ist_cell.value = 'IST'
        ist_cell.font = Font(bold=True)
        ist_cell.alignment = Alignment(horizontal='center')

        sheet.cell(row=4, column=plan_col).value = str(current_year + 1)
        sheet.cell(row=4, column=plan_col).font = Font(bold=True)
        sheet.cell(row=4, column=plan_col).alignment = Alignment(horizontal='center')

        sheet.cell(row=4, column=ist_col).value = str(current_year) + 'e'
        sheet.cell(row=4, column=ist_col).font = Font(bold=True)
        sheet.cell(row=4, column=ist_col).alignment = Alignment(horizontal='center')

        # Step 3: VLOOKUP logic
        for row in range(5, end_row + 1):
            ab_val = sheet.cell(row=row, column=28).value  # AB is col 28
            if not ab_val:
                continue
            parts = [part.strip() for part in str(ab_val).replace(',', ' ').split()]
            sum_c = sum_d = 0
            for part in parts:
                match_key = fuzzy_match(part, kosten_data_c.keys())
                if match_key:
                    sum_c += kosten_data_c[match_key] if kosten_data_c[match_key] else 0
                    sum_d += kosten_data_d[match_key] if kosten_data_d[match_key] else 0
            sheet.cell(row=row, column=plan_col).value = sum_d
            sheet.cell(row=row, column=ist_col).value = sum_c

        # Step 4: Hide unnecessary columns
        keep_cols = set([1, plan_col, ist_col, ver_col])
        for col in range(1, sheet.max_column + 1):
            hdr1 = str(sheet.cell(row=3, column=col).value).strip().lower() if sheet.cell(row=3, column=col).value else ''
            hdr2 = str(sheet.cell(row=4, column=col).value).strip().lower() if sheet.cell(row=4, column=col).value else ''
            if (
                (hdr1 == 'ist' and (hdr2 == str(current_year) or hdr2 == f'{current_year}e' or
                                    hdr2 == str(current_year - 1) or hdr2 == f'{current_year - 1}e'))
                or col in keep_cols
            ):
                sheet.column_dimensions[sheet.cell(row=1, column=col).column_letter].hidden = False
            else:
                sheet.column_dimensions[sheet.cell(row=1, column=col).column_letter].hidden = True

    wb.save(file_path)

def main():
    current_year = datetime.now().year
    directory = '.'  # current working directory
    kosten_file = None
    for file in os.listdir(directory):
        if file.lower().startswith('kostenstelle') and file.endswith('.xlsx'):
            kosten_file = os.path.join(directory, file)
            break

    if not kosten_file:
        print("Kostenstelle file not found!")
        return

    kosten_data_c, kosten_data_d = get_kostenstelle_data(kosten_file)

    for file in os.listdir(directory):
        if file.lower().startswith('kostenstelle') or not file.endswith('.xlsx'):
            continue
        file_path = os.path.join(directory, file)
        process_workbook(file_path, kosten_data_c, kosten_data_d, current_year)

if __name__ == "__main__":
    main()
