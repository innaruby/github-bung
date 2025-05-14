from openpyxl.utils import get_column_letter

def process_sachaufwand_links(wb):
    sach_sheet = None
    for sheet in wb.sheetnames:
        if sheet.lower() == "sachaufwand":
            sach_sheet = wb[sheet]
            break

    if not sach_sheet:
        print("❌ No 'Sachaufwand' sheet found.")
        return

    print("🔍 Processing Sachaufwand links...")

    # Prepare lowercase sheet name map for case-insensitive lookup
    sheet_map = {s.lower(): s for s in wb.sheetnames}

    for row in range(5, sach_sheet.max_row + 1):
        # ✅ Skip hidden rows
        if sach_sheet.row_dimensions[row].hidden:
            continue

        ref_sheet_name = sach_sheet.cell(row=row, column=1).value
        if not ref_sheet_name or not isinstance(ref_sheet_name, str):
            continue

        matched_sheet_name = sheet_map.get(ref_sheet_name.strip().lower())
        if not matched_sheet_name:
            print(f"⚠️ Sheet '{ref_sheet_name}' not found.")
            continue

        matched_sheet = wb[matched_sheet_name]

        # 🔎 Find the "Summe" row (bold + contains "summe")
        summe_row = None
        for r in range(5, matched_sheet.max_row + 1):
            cell = matched_sheet.cell(row=r, column=1)
            if cell.value and "summe" in str(cell.value).lower() and cell.font.bold:
                summe_row = r
                break
        if not summe_row:
            print(f"⚠️ 'Summe' row not found in sheet '{matched_sheet_name}'.")
            continue

        # 🔎 Find the second "Veränderung" column
        veraenderung_cols = []
        for r in [3, 4]:
            for c in range(2, matched_sheet.max_column + 1):
                val = matched_sheet.cell(row=r, column=c).value
                if val and "veränderung" in str(val).lower():
                    veraenderung_cols.append(c)

        if len(veraenderung_cols) < 2:
            print(f"⚠️ Less than 2 'Veränderung' columns found in sheet '{matched_sheet_name}'.")
            continue

        vera_limit_col = veraenderung_cols[1]  # include this column

        # 📥 Copy visible values from matched sheet's Summe row (columns B to 2nd Veränderung col)
        data_to_copy = []
        for col in range(2, vera_limit_col + 1):
            col_letter = get_column_letter(col)
            if not matched_sheet.column_dimensions[col_letter].hidden:
                val = matched_sheet.cell(row=summe_row, column=col).value
                data_to_copy.append((col, val))

        # 📤 Paste into Sachaufwand at same row starting from column B in visible columns only
        paste_col_idx = 2
        for col_idx, val in data_to_copy:
            paste_col_letter = get_column_letter(paste_col_idx)
            if not sach_sheet.column_dimensions[paste_col_letter].hidden:
                sach_sheet.cell(row=row, column=paste_col_idx).value = val
            paste_col_idx += 1

        print(f"✅ Copied from '{matched_sheet_name}' (Summe row) → Sachaufwand row {row}")
