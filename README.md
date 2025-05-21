 ws.cell(row=3, column=insert_col).value = "IST"
            ws.cell(row=4, column=insert_col).value = f"{current_year}e"
            style_cell(ws.cell(row=3, column=insert_col))
            style_cell(ws.cell(row=4, column=insert_col))

            ws.cell(row=3, column=insert_col + 1).value = "PLAN"
            ws.cell(row=4, column=insert_col + 1).value = current_year + 1
            style_cell(ws.cell(row=3, column=insert_col + 1))
            style_cell(ws.cell(row=4, column=insert_col + 1))

            # Identify a visible reference column before the insertion point
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
                        ist_cell.number_format = ref_cell.number_format  # This is a string, safe to assign directly

                        plan_cell.font = copy.copy(ref_cell.font)
                        plan_cell.alignment = copy.copy(ref_cell.alignment)
                        plan_cell.border = copy.copy(ref_cell.border)
                        plan_cell.fill = copy.copy(ref_cell.fill)
                        plan_cell.number_format = ref_cell.number_format
                else:
                    # fallback to your existing styling
                    style_cell(ist_cell)
                    style_cell(plan_cell)
