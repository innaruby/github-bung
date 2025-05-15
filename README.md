from openpyxl.utils import get_column_letter
import openpyxl

def process_sachaufwand_links(wb, file_path):
    # Step 1: Reload workbook WITHOUT data_only to access formulas
    wb_with_formulas = openpyxl.load_workbook(file_path, data_only=False)

    # Step 2: Get 'Sachaufwand' sheet (case-insensitive) from both workbooks
    sach_sheet = None
    sach_sheet_formula = None
    for sheet in wb.sheetnames:
        if sheet.lower() == "sachaufwand":
            sach_sheet = wb[sheet]
            break
    for sheet in wb_with_formulas.sheetnames:
        if sheet.lower() == "sachaufwand":
            sach_sheet_formula = wb_with_formulas[sheet]
            break

    if not sach_sheet or not sach_sheet_formula:
        print("❌ 'Sachaufwand' sheet not found in one or both workbooks.")
        return

    print("\n🔍 Starting process_sachaufwand_links for 'Sachaufwand'...")

    # Step 3: Clear values and formulas from row 5 onward, columns B+
    cleared_cells = 0
    max_row = sach_sheet_formula.max_row
    max_col = sach_sheet_formula.max_column

    for row in range(5, max_row + 1):
        for col in range(2, max_col + 1):  # Start from column B
            cell_formula = sach_sheet_formula.cell(row=row, column=col)
            cell_target = sach_sheet.cell(row=row, column=col)
            if isinstance(cell_formula.value, str) and cell_formula.value.strip().startswith("="):
                cell_target.value = None
                cleared_cells += 1
            elif cell_target.value is not None:
                cell_target.value = None
                cleared_cells += 1

    print(f"🧹 Cleared {cleared_cells} cells from 'Sachaufwand' (excluding headers and column A).")

    # Step 4: Prepare lowercase sheet name map
    sheet_map = {s.lower(): s for s in wb.sheetnames}

    # Step 5: Define function to find end row
    def find_end_row(sheet, sheet_name):
        for row in range(sheet.max_row, 0, -1):
            if any(sheet.cell(row=row, column=col).value is not None for col in range(1, sheet.max_column + 1)):
                return row
        raise ValueError(f"End row not found in sheet {sheet_name}")

    # Step 6: Find end row in Sachaufwand
    try:
        end_row = find_end_row(sach_sheet, "Sachaufwand")
        print(f"✅ Detected end row in 'Sachaufwand': {end_row}")
    except Exception as e:
        print(f"❌ Error determining end row in 'Sachaufwand': {e}")
        return

    # Step 7: Loop through each visible row
    for row in range(5, end_row + 1):
        if sach_sheet.row_dimensions[row].hidden:
            continue

        ref_value = sach_sheet.cell(row=row, column=1).value
        if not ref_value or not isinstance(ref_value, str):
            continue

        ref_key = ref_value.strip().lower()
        matched_sheet_name = sheet_map.get(ref_key)
        if not matched_sheet_name:
            continue

        matched_sheet = wb[matched_sheet_name]

        # Step 8: Find bold 'Summe' row
        summe_row = None
        for r in range(5, matched_sheet.max_row + 1):
            cell = matched_sheet.cell(row=r, column=1)
            if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
                summe_row = r
                break

        if not summe_row:
            continue

        # Step 9: Identify visible source columns
        visible_source_cols = [
            col for col in range(2, matched_sheet.max_column + 1)
            if not matched_sheet.column_dimensions[get_column_letter(col)].hidden
        ]

        if not visible_source_cols:
            continue

        # Step 10: Collect values from Summe row
        data_to_copy = []
        for col in visible_source_cols:
            val = matched_sheet.cell(row=summe_row, column=col).value
            data_to_copy.append((col, val))

        # Step 11: Identify visible target columns in Sachaufwand
        visible_target_cols = [
            col for col in range(2, sach_sheet.max_column + 1)
            if not sach_sheet.column_dimensions[get_column_letter(col)].hidden
        ]

        # Step 12: Paste values into target row
        for i, col in enumerate(visible_target_cols):
            if i < len(data_to_copy):
                value = data_to_copy[i][1]
                sach_sheet.cell(row=row, column=col).value = value

    # -----------------------
    # ✨ Additional Step: Zwischensumme and Summe Aggregation
    # -----------------------
    zwischensumme_row = None
    final_summe_row = None

    for row in range(5, end_row + 1):
        cell = sach_sheet.cell(row=row, column=1)
        if cell.value and isinstance(cell.value, str):
            text = str(cell.value).lower()
            if "zwischensumme" in text and cell.font.bold and not zwischensumme_row:
                zwischensumme_row = row
            elif "summe" in text and cell.font.bold and not final_summe_row:
                final_summe_row = row

    # Get visible columns B+
    visible_cols = [
        col for col in range(2, sach_sheet.max_column + 1)
        if not sach_sheet.column_dimensions[get_column_letter(col)].hidden
    ]

    def sum_visible_rows(start_row, end_row):
        col_sums = {col: 0 for col in visible_cols}
        for row in range(start_row, end_row):
            if sach_sheet.row_dimensions[row].hidden:
                continue
            for col in visible_cols:
                val = sach_sheet.cell(row=row, column=col).value
                if isinstance(val, (int, float)):
                    col_sums[col] += val
        return col_sums

    if zwischensumme_row:
        zw_sum = sum_visible_rows(5, zwischensumme_row)
        for col, value in zw_sum.items():
            sach_sheet.cell(row=zwischensumme_row, column=col).value = value
        print(f"✅ Wrote Zwischensumme totals at row {zwischensumme_row}.")

    if final_summe_row and zwischensumme_row:
        summe_sum = sum_visible_rows(zwischensumme_row, final_summe_row)
        for col, value in summe_sum.items():
            sach_sheet.cell(row=final_summe_row, column=col).value = value
        print(f"✅ Wrote Summe totals at row {final_summe_row}.")
