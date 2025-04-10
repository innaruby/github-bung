def get_sheet_tab_color(file_path, sheet_name):
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    try:
        sht = wb.sheets[sheet_name]
        color = sht.api.Tab.Color  # RGB integer or None
        print(f"Sheet '{sheet_name}' tab color: {color}")
    finally:
        wb.close()
        app.quit()

    if color is None:
        return False

    # If color is a single RGB int, extract R, G, B
    try:
        b = color & 255
        g = (color >> 8) & 255
        r = (color >> 16) & 255
        print(f"Extracted RGB: R={r}, G={g}, B={b}")
    except Exception as e:
        print(f"Error decoding tab color: {e}")
        return False

    # Define green/yellow logic
    is_yellow = r > 200 and g > 200 and b < 150
    is_green = g > 150 and r < 200 and b < 150

    return is_yellow or is_green
