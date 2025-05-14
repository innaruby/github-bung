def final_sum_pass(directory):
    print("\n📘 Starting final sum pass (apply_final_sums at very end)...")
    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue

        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)
        sheet_colors = get_sheet_tab_colors(file_path)

        for sheet_name in wb.sheetnames:
            tab_color = sheet_colors.get(sheet_name, "")
            if "green" not in tab_color.lower():
                continue

            ws = wb[sheet_name]
            end_row = find_end_row(ws, sheet_name)
            apply_final_sums(ws, end_row)

        wb.save(file_path)
        print(f"✅ Final sum pass completed for: {file}")


def main():
    root = Tk()
    root.withdraw()
    selected_directory = filedialog.askdirectory(title="Select Directory with Excel Files")
    if selected_directory:
        process_excel_files(selected_directory)
        post_processing_with_vlookup(selected_directory)
        final_sum_pass(selected_directory)  # ✅ Run apply_final_sums at very end
