def apply_final_adjustments_clean(file_path):
    wb = openpyxl.load_workbook(file_path)
    for ws in wb.worksheets:
        veraenderung_cols = find_merged_veraenderung_columns(ws)
        if not veraenderung_cols:
            continue

        vera_start_col = veraenderung_cols[0]
        new_ist_col = vera_start_col
        new_plan_col = vera_start_col + 1

        for col in [new_ist_col, new_plan_col]:
            for row in [3, 4]:
                cell = ws.cell(row=row, column=col)
                cell.border = Border(
                    top=Side(style=None),
                    bottom=Side(style=None),
                    left=Side(style=cell.border.left.style if cell.border.left else None),
                    right=Side(style=cell.border.right.style if cell.border.right else None)
                )

        reference_col = None
        for col in range(new_ist_col - 1, 1, -1):
            if not ws.column_dimensions[get_column_letter(col)].hidden:
                reference_col = col
                break

        if reference_col:
            for row in [3, 4]:
                ref_cell = ws.cell(row=row, column=reference_col)
                for target_col in [new_ist_col, new_plan_col]:
                    target_cell = ws.cell(row=row, column=target_col)
                    target_cell.font = Font(
                        name=ref_cell.font.name,
                        size=ref_cell.font.size,
                        bold=ref_cell.font.bold,
                        italic=ref_cell.font.italic,
                        color=ref_cell.font.color
                    )

    wb.save(file_path)
