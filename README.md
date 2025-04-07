for i, row in enumerate(kostenstelle_ws.iter_rows(min_row=2), start=2):
    a_val = row[0].value
    fill = row[0].fill

    print(f"Row {i} in Kostenstelle: A={a_val}")
    if isinstance(fill, PatternFill):
        print(f"  Fill type: {fill.fill_type}")
        print(f"  fgColor: {fill.fgColor}, start_color: {fill.start_color}")
        print(f"  rgb: {fill.start_color.rgb}, indexed: {fill.start_color.indexed}, theme: {fill.start_color.theme}")

    fill_color = fill.start_color.rgb if isinstance(fill, PatternFill) and fill.fill_type == "solid" else None

    # Match only explicitly defined FF90EE90 colors (not theme-based)
    if fill_color == "FF90EE90":
        green_a_values.add(a_val)
        green_i_map[a_val] = row[8].value
        print(f"  -> Registered green A cell: {a_val} with I={row[8].value}")

    kostenstelle_data[a_val] = {
        "E": row[4].value,
        "F": row[5].value,
        "I": row[8].value,
    }
