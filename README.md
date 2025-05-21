...

# Insert this function at the very end, after the main execution

def apply_number_format_to_ist_plan_columns(directory):
    from openpyxl.styles import numbers

    current_year = datetime.now().year
    for file in os.listdir(directory):
        if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
            continue

        file_path = os.path.join(directory, file)
        wb = openpyxl.load_workbook(file_path)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            tab_color = rgb_to_hex_name(ws.sheet_properties.tabColor)
            if ws.sheet_properties.tabColor is None or "green" not in tab_color.lower():
                continue

            end_row = find_end_row(ws, sheet_name)

            visible_cols = [col for col in range(1, ws.max_column + 1)
                            if not ws.column_dimensions[get_column_letter(col)].hidden]

            ist_col = None
            plan_col = None

            for col in visible_cols:
                val_row3 = str(ws.cell(row=3, column=col).value).strip().upper()
                val_row4 = str(ws.cell(row=4, column=col).value).replace("e", "").strip()

                if val_row3 == "IST" and val_row4 == str(current_year):
                    ist_col = col
                if val_row3 == "PLAN" and val_row4 == str(current_year + 1):
                    plan_col = col

            for col in [ist_col, plan_col]:
                if col is not None:
                    for row in range(5, end_row + 1):
                        cell = ws.cell(row=row, column=col)
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0.000'  # Match formatting logic used elsewhere

        wb.save(file_path)
        print(f"🧾 Number formatting applied to IST and PLAN columns in file: {file}")

# Update main() to include this at the end

def main():
    root = Tk()
    root.withdraw()
    selected_directory = filedialog.askdirectory(title="Select Directory with Excel Files")
    if selected_directory:
        process_excel_files(selected_directory)
        post_processing_with_vlookup(selected_directory)
        final_sum_pass(selected_directory)

        for file in os.listdir(selected_directory):
            if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
                continue
            file_path = os.path.join(selected_directory, file)
            wb = openpyxl.load_workbook(file_path)
            process_sachaufwand_links(wb, file_path) 
            for sheet_name in wb.sheetnames:
                if sheet_name.lower() == 'sachaufwand':
                    ws = wb[sheet_name]
                    end_row = find_end_row(ws, sheet_name)
                    apply_final_sums(ws, end_row)
                    print(f"📘 Zwischensumme and Summe logic applied to 'Sachaufwand' in file: {file}")
            wb.save(file_path)
            print(f"💾 Final update (Sachaufwand) saved in file: {file}")

        # 🆕 Apply formatting to IST and PLAN columns
        apply_number_format_to_ist_plan_columns(selected_directory)

if __name__ == "__main__":
    main()
