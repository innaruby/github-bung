for i, (row_val, row_style) in enumerate(zip(kostenstelle_ws.iter_rows(min_row=2),
                                              kostenstelle_ws_styles.iter_rows(min_row=2)), start=2):
    a_val = row_val[0].value
    fill = row_style[0].fill
    fill_color = None
    if fill and fill.fill_type == "solid":
        if fill.start_color.type == "rgb":
            fill_color = fill.start_color.rgb
        elif fill.start_color.type == "theme":
            fill_color = fill.start_color.theme  # fallback for theme color

    print(f"Row {i} in Kostenstelle: A={a_val}, Fill={fill_color}")

    # Accept both RGB and theme index (3 is common for light green in themes)
    if fill_color in ("FF90EE90", "0090EE90", 3):
        green_a_values.add(a_val)
        green_i_map[a_val] = row_val[8].value
        print(f"  -> Registered green A cell: {a_val} with I={row_val[8].value}")

    kostenstelle_data[a_val] = {
        "E": row_val[4].value,
        "F": row_val[5].value,
        "I": row_val[8].value,
    }
