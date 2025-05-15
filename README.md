if c_val.startswith("9"):
    copy_ws[f"C{r}"].value = input_ws[f"C{r}"].value
    copy_ws[f"D{r}"].value = input_ws[f"D{r}"].value

    # Perform VLOOKUP using column H and write result to column B
    h_val = input_ws[f"H{r}"].value
    if h_val in kostenstelle_data:
        copy_ws[f"B{r}"].value = kostenstelle_data[h_val]["E"]

    d_val = copy_ws[f"D{r}"].value
    if d_val:
        length = len(str(d_val))  # include spaces
        copy_ws[f"L{r}"].value = length
        if length > 50:
            copy_ws[f"L{r}"].fill = orange_fill

    process_column_m(r, copy_ws, kostenstelle_data, column_a_colors, green_i_lookup, kostenstelle_ws)
    continue
