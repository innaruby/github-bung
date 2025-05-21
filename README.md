def apply_final_sums(ws, end_row):
    from openpyxl.cell.cell import MergedCell

    sheet_name = ws.title
    print(f"\n🧮 Starting apply_final_sums for sheet: {sheet_name}")

    veraenderung_cols = find_merged_veraenderung_columns(ws)
    if not veraenderung_cols:
        print(f"❌ [Sheet: {sheet_name}] Veränderung columns not found.")
        return

    vera_col_start = veraenderung_cols[0]
    visible_cols = [col for col in range(2, vera_col_start)
                    if not ws.column_dimensions[get_column_letter(col)].hidden]

    if not visible_cols:
        print(f"⚠️ [Sheet: {sheet_name}] No visible columns before Veränderung.")
        return

    # Identify visible rows and special rows
    visible_rows = [r for r in range(5, end_row + 1) if not ws.row_dimensions[r].hidden]
    zwischensumme_rows = []
    summe_rows = []

    for row in visible_rows:
        cell_val = str(ws.cell(row=row, column=1).value or "").strip().lower()
        if ws.cell(row=row, column=1).font.bold:
            if "zwischensumme" in cell_val:
                zwischensumme_rows.append(row)
            elif "summe" in cell_val:
                summe_rows.append(row)

    # Pass 1: Handle Zwischensumme
    zwischensumme_values = {}  # {row: {col: value}}
    previous = 4
    for z_row in zwischensumme_rows:
        rows_to_sum = [r for r in visible_rows if previous < r < z_row]
        zwischensumme_values[z_row] = {}
        print(f"🧩 Zwischensumme row {z_row} → summing rows {rows_to_sum}")
        for col in visible_cols:
            total = 0
            for r in rows_to_sum:
                val = ws.cell(row=r, column=col).value
                parsed = parse_numeric(val)
                total += parsed
            cell = ws.cell(row=z_row, column=col)
            cell.value = None
            cell.value = total
            zwischensumme_values[z_row][col] = total
            print(f"   🟢 Wrote {total} to {get_column_letter(col)}{z_row}")
        previous = z_row

    # Pass 2: Handle Summe
    for s_row in summe_rows:
        print(f"🧮 Processing Summe row {s_row}")
        prev_z_row = max([z for z in zwischensumme_rows if z < s_row], default=None)
        rows_to_sum = []

        if prev_z_row:
            rows_to_sum = [r for r in visible_rows if prev_z_row < r < s_row]
        else:
            rows_to_sum = [r for r in visible_rows if r < s_row]

        for col in visible_cols:
            total = 0

            if prev_z_row:
                total += zwischensumme_values.get(prev_z_row, {}).get(col, 0)

            for r in rows_to_sum:
                val = ws.cell(row=r, column=col).value
                parsed = parse_numeric(val)
                total += parsed

            cell = ws.cell(row=s_row, column=col)
            cell.value = None
            cell.value = total
            print(f"   ✅ Wrote total {total} to {get_column_letter(col)}{s_row}")

def parse_numeric(val):
    if val is None or val == "":
        return 0
    elif isinstance(val, (int, float)):
        return val
    elif isinstance(val, str) and val.strip().startswith("="):
        try:
            return eval(val.strip().lstrip("="))
        except:
            return 0
    try:
        return float(str(val).strip())
    except:
        return 0
