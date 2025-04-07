# Load kostenstelle without data_only so you can read styles
kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=False)
kostenstelle_ws = kostenstelle_wb.active

# Create lookup and identify green cell values in column A
kostenstelle_data = {}
green_a_values = set()
green_i_map = {}
for i, row in enumerate(kostenstelle_ws.iter_rows(min_row=2), start=2):
    a_val = row[0].value
    fill = row[0].fill
    fill_color = fill.start_color.rgb if fill and fill.fill_type == "solid" else None
    print(f"Row {i} in Kostenstelle: A={a_val}, Fill={fill_color}")
    if fill_color == "FF90EE90":
        green_a_values.add(a_val)
        green_i_map[a_val] = row[8].value
        print(f"  -> Registered green A cell: {a_val} with I={row[8].value}")
    kostenstelle_data[a_val] = {
        "E": row[4].value,
        "F": row[5].value,
        "I": row[8].value,
    }
