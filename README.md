import copy
...

# Insert this function at the very end, after the main execution

def apply_number_format_to_ist_plan_columns(directory):
    from openpyxl.styles import numbers

    current_year = datetime.now().year
    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue

        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)
        print(f"\n🔍 Processing file: {file}")

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            tab_color = rgb_to_hex_name(ws.sheet_properties.tabColor)
            if ws.sheet_properties.tabColor is None or "green" not in tab_color.lower():
                print(f"  ⏭️ Skipping sheet '{sheet_name}' (non-green tab)")
                continue

            end_row = find_end_row(ws, sheet_name)
            print(f"  📄 Sheet: {sheet_name}, End Row: {end_row}")

            visible_cols = [col for col in range(1, ws.max_column + 1)
                            if not ws.column_dimensions[get_column_letter(col)].hidden]

            ist_col = None
            plan_col = None

            for col in visible_cols:
                val_row3 = str(ws.cell(row=3, column=col).value).strip().upper()
                val_row4 = str(ws.cell(row=4, column=col).value).replace("e", "").strip()

                print(f"    🔎 Column {get_column_letter(col)}: Row 3 = '{val_row3}', Row 4 = '{val_row4}'")

                if val_row3 == "IST" and val_row4 == str(current_year):
                    ist_col = col
                if val_row3 == "PLAN" and val_row4 == str(current_year + 1):
                    plan_col = col

            print(f"    📌 Identified IST column: {get_column_letter(ist_col) if ist_col else 'None'}")
            print(f"    📌 Identified PLAN column: {get_column_letter(plan_col) if plan_col else 'None'}")

            if ist_col is None or plan_col is None:
                continue

            insert_col = ist_col

            ws.cell(row=3, column=insert_col).value = "IST"
            ws.cell(row=4, column=insert_col).value = f"{current_year}e"
            style_cell(ws.cell(row=3, column=insert_col))
            style_cell(ws.cell(row=4, column=insert_col))

            ws.cell(row=3, column=insert_col + 1).value = "PLAN"
            ws.cell(row=4, column=insert_col + 1).value = current_year + 1
            style_cell(ws.cell(row=3, column=insert_col + 1))
            style_cell(ws.cell(row=4, column=insert_col + 1))

            reference_col = None
            for c in range(insert_col - 1, 1, -1):
                if not ws.column_dimensions[get_column_letter(c)].hidden:
                    reference_col = c
                    break

            for row in range(5, end_row + 1):
                ist_cell = ws.cell(row=row, column=insert_col)
                plan_cell = ws.cell(row=row, column=insert_col + 1)

                ist_cell.value = None
                plan_cell.value = None

                if reference_col:
                    ref_cell = ws.cell(row=row, column=reference_col)
                    if ref_cell.has_style:
                        ist_cell.font = copy.copy(ref_cell.font)
                        ist_cell.alignment = copy.copy(ref_cell.alignment)
                        ist_cell.border = copy.copy(ref_cell.border)
                        ist_cell.fill = copy.copy(ref_cell.fill)
                        ist_cell.number_format = ref_cell.number_format

                        plan_cell.font = copy.copy(ref_cell.font)
                        plan_cell.alignment = copy.copy(ref_cell.alignment)
                        plan_cell.border = copy.copy(ref_cell.border)
                        plan_cell.fill = copy.copy(ref_cell.fill)
                        plan_cell.number_format = ref_cell.number_format
                else:
                    style_cell(ist_cell)
                    style_cell(plan_cell)

                for col in [insert_col, insert_col + 1]:
                    cell = ws.cell(row=row, column=col)
                    if isinstance(cell.value, (int, float)):
                        original_val = cell.value
                        rounded_val = round(float(original_val), 3)
                        cell.value = rounded_val
                        cell.number_format = '#,##0.000'
                        print(f"      ✅ Formatted {get_column_letter(col)}{row}: {original_val} → {rounded_val}")
                    elif isinstance(cell.value, str) and cell.value.strip().startswith("="):
                        cell.number_format = '#,##0.000'
                        print(f"      🧮 Formula cell {get_column_letter(col)}{row} formatted: {cell.value}")
                    else:
                        print(f"      ❌ Skipped {get_column_letter(col)}{row} (non-numeric or empty): {cell.value}")

        wb.save(file_path)
        print(f"🧾 Number formatting applied to IST and PLAN columns in file: {file}")
