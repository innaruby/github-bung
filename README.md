def process_excel_files(directory):
    current_year = datetime.now().year

    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue
        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)
        sheet_colors = get_sheet_tab_colors(file_path)

        for sheet_name in wb.sheetnames:
            tab_color = sheet_colors.get(sheet_name, "")
            if "green" not in tab_color.lower():
                continue

            ws = wb[sheet_name]
            delete_columns_B_and_C(ws)
            end_row = find_end_row(ws, sheet_name)
            vera_cols = find_merged_veraenderung_columns(ws)
            if vera_cols is None:
                continue

            vera_start_col, vera_end_col = vera_cols
            insert_col = vera_start_col

            merged_to_restore = []
            for merged_range in list(ws.merged_cells.ranges):
                if merged_range.min_row == 3 and merged_range.max_row == 3:
                    if merged_range.min_col == vera_start_col and merged_range.max_col == vera_end_col:
                        merged_to_restore.append(merged_range)
                        ws.unmerge_cells(str(merged_range))

            existing_plan = ws.cell(row=3, column=vera_start_col - 2).value
            existing_ist = ws.cell(row=3, column=vera_start_col - 1).value
            if str(existing_plan).strip().lower() == "plan" and str(existing_ist).strip().lower() == "ist":
                print(f"Skipping insertion in sheet '{sheet_name}' of file '{file}' as columns already exist.")
                continue

            ws.insert_cols(insert_col, 2)

            for merged_range in merged_to_restore:
                new_start = merged_range.min_col + 2
                new_end = merged_range.max_col + 2
                ws.merge_cells(start_row=3, start_column=new_start, end_row=3, end_column=new_end)

            ws.cell(row=3, column=insert_col).value = "IST"
            ws.cell(row=4, column=insert_col).value = f"{current_year}e"
            style_cell(ws.cell(row=3, column=insert_col))
            style_cell(ws.cell(row=4, column=insert_col))

            ws.cell(row=3, column=insert_col + 1).value = "PLAN"
            ws.cell(row=4, column=insert_col + 1).value = current_year + 1
            style_cell(ws.cell(row=3, column=insert_col + 1))
            style_cell(ws.cell(row=4, column=insert_col + 1))

            for row in range(5, end_row + 1):
                ws.cell(row=row, column=insert_col).value = None
                ws.cell(row=row, column=insert_col + 1).value = None
                style_cell(ws.cell(row=row, column=insert_col))
                style_cell(ws.cell(row=row, column=insert_col + 1))

            apply_veraenderung_formulas(ws, ist_col=insert_col, plan_col=insert_col + 1,
                                        vera_start_col=vera_start_col, end_row=end_row)

            unhide_cols = {1, insert_col, insert_col + 1}
            unhide_cols.update(range(vera_start_col + 2, vera_end_col + 4))

            for col in range(1, ws.max_column + 1):
                header3 = ws.cell(row=3, column=col).value
                header4 = str(ws.cell(row=4, column=col).value)
                if (header3 == "PLAN" and header4.replace("e", "").strip() == str(current_year)) or \
                   (header3 == "IST" and header4.replace("e", "").strip() in [str(current_year), str(current_year - 1), str(current_year - 2)]):
                    unhide_cols.add(col)
                    if col != insert_col and col != insert_col + 1:
                        ws.cell(row=4, column=col).value = header4.replace("e", "").strip()

            for col in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col)
                ws.column_dimensions[col_letter].hidden = col not in unhide_cols

            for col in unhide_cols:
                if col != 1:
                    col_letter = get_column_letter(col)
                    ws.column_dimensions[col_letter].width = 18

        # Save the green sheet updates
        wb.save(file_path)

        # ✅ Apply Veränderung logic to Sachaufwand if it exists
        sach_sheet = None
        for sname in wb.sheetnames:
            if sname.lower() == "sachaufwand":
                sach_sheet = wb[sname]
                break

        if sach_sheet:
            sach_end_row = find_end_row(sach_sheet, sach_sheet.title)
            vera_cols = find_merged_veraenderung_columns(sach_sheet)
            if vera_cols:
                vera_start_col, _ = vera_cols
                apply_veraenderung_formulas(
                    sach_sheet,
                    ist_col=vera_start_col - 2,
                    plan_col=vera_start_col - 1,
                    vera_start_col=vera_start_col,
                    end_row=sach_end_row
                )
                apply_final_sums(sach_sheet, sach_end_row)
                print(f"✅ Veränderung formulas and final sums applied to 'Sachaufwand' sheet.")
