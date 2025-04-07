# Load kostenstelle for values only
kostenstelle_wb = openpyxl.load_workbook(kostenstelle_path, data_only=True)
kostenstelle_ws = kostenstelle_wb.active

# Load kostenstelle again for styles (fills)
kostenstelle_wb_styles = openpyxl.load_workbook(kostenstelle_path, data_only=False)
kostenstelle_ws_styles = kostenstelle_wb_styles.active
