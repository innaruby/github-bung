def execute_script2():
    try:
        # Directory and file identification
        date = date_entry.get()
        month = int(date.split('.')[0])
        if 1 <= month <= 3:
            Quartal = "1. Quartal"
        elif 4 <= month <= 6:
            Quartal = "2. Quartal"
        elif 7 <= month <= 9:
            Quartal = "3. Quartal"
        elif 10 <= month <= 12:
            Quartal = "4. Quartal"
        else:
            Quartal = "Invalid month"
        year = date.split('.')[1]

        base_path = rf'U:\\Skript zum Aufrufen'
        file_name = next((file for file in os.listdir(base_path) if file.startswith('KONBUCH') and file.endswith('.xlsx')), None)
        if not file_name:
            raise FileNotFoundError("No file starting with 'KONBUCH' found in the directory.")
        file_path = os.path.join(base_path, file_name)

        sheets = pd.read_excel(file_path, sheet_name=None, header=None)
        df_rlbooe = sheets.get('RLBOOE')
        df_rlbooe_mit_teilkonzerne = sheets.get('RLB mit Teilkonzerne')
        df_rlz_mapping = sheets.get('RLZ Mapping I7 und N7').iloc[11:]
        df_kifi = sheets.get('KIFI')
        df_kifi_mit_teilkonzerne = sheets.get('Kifi mit Teilkonzerne')

        # Move rows from RLBOOE to KIFI, considering from row 4 onward
        values_to_move = {'ILI', 'ILG', 'KIFI'}
        rlbooe_data = df_rlbooe.iloc[3:].copy()
        rows_to_move = rlbooe_data[rlbooe_data[0].isin(values_to_move)]
        remaining_rlbooe = rlbooe_data[~rlbooe_data[0].isin(values_to_move)]
        df_rlbooe = pd.concat([df_rlbooe.iloc[:3], remaining_rlbooe], ignore_index=True)

        df_kifi_data = df_kifi.iloc[3:].copy()
        df_kifi_header = df_kifi.iloc[:3]
        df_kifi_updated = pd.concat([df_kifi_data, rows_to_move], ignore_index=True)
        df_kifi = pd.concat([df_kifi_header, df_kifi_updated], ignore_index=True)

        date_format = '%d.%b.%y'
        df_rlbooe.iloc[3:, 34] = pd.to_datetime(df_rlbooe.iloc[3:, 34], errors='coerce').dt.strftime(date_format)
        df_rlbooe.iloc[3:, 40] = pd.to_datetime(df_rlbooe.iloc[3:, 40], errors='coerce').dt.strftime(date_format)

        def copy_data_to_sections(source_df, target_df):
            source_data = source_df.iloc[3:]
            target_data = target_df.iloc[3:].reindex(range(len(source_data)), fill_value=None)
            target_data.iloc[:, :source_data.shape[1]] = source_data.values
            target_data = target_data.reindex(columns=range(max(66, target_data.shape[1])), fill_value=None)
            target_data.iloc[:, 43:56] = source_data.iloc[:, :13].values
            target_data.iloc[:, 57:59] = source_data.iloc[:, 13:15].values
            target_data.iloc[:, 60:66] = source_data.iloc[:, 15:21].values
            return pd.concat([target_df.iloc[:3], target_data], ignore_index=True)

        def perform_vlookup_and_update(target_df, mapping_df):
            lookup_dict = mapping_df.set_index(0)[2].to_dict()
            target_df.iloc[3:, 52] = target_df.iloc[3:, 9].map(lookup_dict).fillna("kein Mapping Vorhanden")
            return target_df

        def check_and_update_bd_column(target_df, mapping_df):
            lookup_dict = mapping_df.set_index(2)[3].to_dict()
            target_df.iloc[3:, 55] = target_df.iloc[3:, 52].map(lookup_dict).fillna("kein Mapping Vorhanden")
            return target_df

        def update_bn_column(target_df):
            target_df.iloc[3:, 65] = target_df.iloc[3:].apply(lambda row: row[20] if row[55] != "kein Mapping Vorhanden" else "kein Mapping Vorhanden", axis=1)
            return target_df

        def update_bb_column(target_df, specific_value):
            target_df.iloc[3:, 53] = specific_value
            return target_df

        def update_be_column(target_df):
            target_df.iloc[3:, 56] = target_df.iloc[3:, 57].apply(lambda x: "H" if x in [0, 0.00, 0.0000, "0,00"] else "S")
            return target_df

        def cut_and_paste_rows_by_values(source_df, target_df, values):
            source_data = source_df.iloc[3:].copy()
            target_data = target_df.iloc[3:].copy()
            rows_to_move = source_data[source_data[0].isin(values)]
            source_data = source_data[~source_data[0].isin(values)]
            target_data = pd.concat([target_data, rows_to_move], ignore_index=True)
            new_source_df = pd.concat([source_df.iloc[:3], source_data], ignore_index=True)
            new_target_df = pd.concat([target_df.iloc[:3], target_data], ignore_index=True)
            return new_source_df, new_target_df

        def filter_rows_based_on_ag(df):
            df.iloc[3:, 32] = pd.to_numeric(df.iloc[3:, 32], errors='coerce')
            filtered_data = df.iloc[3:][df.iloc[3:, 32] == 5]
            return pd.concat([df.iloc[:3], filtered_data], ignore_index=True)

        def convert_to_european_format(value):
            if isinstance(value, str):
                value = value.replace('.', '').replace(',', '.')
                try:
                    return float(value)
                except ValueError:
                    return value
            return value

        def apply_european_formatting_to_columns(df, columns):
            for column in columns:
                df.iloc[3:, column] = df.iloc[3:, column].apply(convert_to_european_format)
            return df

        def filter_rows(df):
            filtered_data = df.iloc[3:][(df.iloc[3:, 3] != 'SK') | (df.iloc[3:, 4] != '99')]
            return pd.concat([df.iloc[:3], filtered_data], ignore_index=True)

        # Process RLBOOE mit Teilkonzerne
        df_rlbooe_mit_teilkonzerne = copy_data_to_sections(df_rlbooe, df_rlbooe_mit_teilkonzerne)
        df_rlbooe_mit_teilkonzerne = perform_vlookup_and_update(df_rlbooe_mit_teilkonzerne, df_rlz_mapping)
        df_rlbooe_mit_teilkonzerne = check_and_update_bd_column(df_rlbooe_mit_teilkonzerne, df_rlz_mapping)
        df_rlbooe_mit_teilkonzerne = update_bn_column(df_rlbooe_mit_teilkonzerne)
        df_rlbooe_mit_teilkonzerne = update_bb_column(df_rlbooe_mit_teilkonzerne, "B100")
        df_rlbooe_mit_teilkonzerne = update_be_column(df_rlbooe_mit_teilkonzerne)

        values_rlbooe = {'KIFI', 'ILG', 'ILI'}
        df_rlbooe_mit_teilkonzerne, df_kifi_mit_teilkonzerne = cut_and_paste_rows_by_values(
            df_rlbooe_mit_teilkonzerne, df_kifi_mit_teilkonzerne, values_rlbooe
        )

        df_rlbooe_mit_teilkonzerne = filter_rows_based_on_ag(df_rlbooe_mit_teilkonzerne)
        df_rlbooe_mit_teilkonzerne = apply_european_formatting_to_columns(df_rlbooe_mit_teilkonzerne, [13, 14, 15, 57, 58, 60])
        df_rlbooe_mit_teilkonzerne = filter_rows(df_rlbooe_mit_teilkonzerne)

        # Process KIFI mit Teilkonzerne
        df_kifi_mit_teilkonzerne = copy_data_to_sections(df_kifi, df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = perform_vlookup_and_update(df_kifi_mit_teilkonzerne, df_rlz_mapping)
        df_kifi_mit_teilkonzerne = check_and_update_bd_column(df_kifi_mit_teilkonzerne, df_rlz_mapping)
        df_kifi_mit_teilkonzerne = update_bn_column(df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = update_bb_column(df_kifi_mit_teilkonzerne, "B100")
        df_kifi_mit_teilkonzerne = update_be_column(df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = filter_rows_based_on_ag(df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = apply_european_formatting_to_columns(df_kifi_mit_teilkonzerne, [13, 14, 15, 57, 58, 60])
        df_kifi_mit_teilkonzerne = filter_rows(df_kifi_mit_teilkonzerne)

        workbook = load_workbook(file_path)

        def write_data_to_sheet(sheet_name, df):
            print(f"Writing to sheet: {sheet_name}")
            sheet = workbook[sheet_name]
            for row_idx, row in enumerate(df.itertuples(index=False, name=None), start=4):
                for col_idx, value in enumerate(row, start=1):
                    cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                    if col_idx in [14, 15, 16, 58, 59, 61]:
                        cell.number_format = '#,##0.00'

        write_data_to_sheet('RLB mit Teilkonzerne', df_rlbooe_mit_teilkonzerne)
        write_data_to_sheet('KIFI', df_kifi)
        write_data_to_sheet('Kifi mit Teilkonzerne', df_kifi_mit_teilkonzerne)

        def auto_adjust_column_width(sheet):
            columns_to_adjust = ['M', 'N', 'O', 'P', 'U', 'V', 'BA', 'BF', 'BG', 'BN', 'BB', 'BD', 'BI']
            for column_letter in columns_to_adjust:
                max_length = 0
                for cell in sheet[column_letter]:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

        workbook['RLB mit Teilkonzerne']['BO4'] = 'B01 B100 SK 08'
        workbook['Kifi mit Teilkonzerne']['BO4'] = 'B01 B99 SK 08'

        auto_adjust_column_width(workbook['RLB mit Teilkonzerne'])
        auto_adjust_column_width(workbook['Kifi mit Teilkonzerne'])

        # 🔻 Delete row 4 and 5 (Excel) from specific sheets
        for sheet_name in ['KIFI', 'RLB mit Teilkonzerne', 'Kifi mit Teilkonzerne']:
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                sheet.delete_rows(4, 2)

        workbook.save(file_path)

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Erfolgreich")
        root.destroy()

    except Exception as e:
        print(f"An error occurred during execution: {e}")
