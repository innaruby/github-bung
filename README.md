for r in range(start_row, end_row + 1):
    c_val = str(input_ws[f"C{r}"].value)
    if c_val.startswith("9"):
        continue

    h_val = copy_ws[f"H{r}"].value
    if h_val in kostenstelle_data:
        k_data = kostenstelle_data[h_val]
        b_val = k_data["E"]
        copy_ws[f"B{r}"].value = b_val

        if c_val.startswith(("705", "706", "707", "5")):
            copy_ws[f"G{r}"].value = "V0" if b_val == 1001 else "U0" if b_val == 1002 else None
        elif c_val.startswith(("704", "6")):
            copy_ws[f"G{r}"].value = "A0" if b_val == 1001 else "D0" if b_val == 1002 else None
        elif c_val.startswith("4"):
            # Clear columns G and H
            copy_ws[f"G{r}"].value = None
            copy_ws[f"H{r}"].value = None

    d_val = copy_ws[f"D{r}"].value
    if d_val:
        length = len(str(d_val))  # include spaces
        copy_ws[f"L{r}"].value = length
        if length > 50:
            copy_ws[f"L{r}"].fill = orange_fill
