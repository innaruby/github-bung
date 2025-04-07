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
            # Extract the year from the date entered in MM.YYYY format
        year = date.split('.')[1]
        base_path = rf'U:\\Skript zum Aufrufen'
        file_name = next((file for file in os.listdir(base_path) if file.startswith('KONBUCH') and file.endswith('.xlsx')), None)

        if not file_name:
            raise FileNotFoundError("No file starting with 'KONBUCH' found in the directory.")

        # Load all sheets into a dictionary of DataFrames
        file_path = os.path.join(base_path, file_name)
        sheets = pd.read_excel(file_path, sheet_name=None, header=None)

        # Extract specific sheets
        df_rlbooe = sheets.get('RLBOOE')
        df_rlbooe_mit_teilkonzerne = sheets.get('RLB mit Teilkonzerne')
        df_rlz_mapping = sheets.get('RLZ Mapping I7 und N7').iloc[11:]
        df_kifi = sheets.get('KIFI')
        df_kifi_mit_teilkonzerne = sheets.get('Kifi mit Teilkonzerne')

        # Step to filter and move rows from RLBOOE to KIFI
        values_to_move = {'ILI', 'ILG', 'KIFI'}
        rows_to_move = df_rlbooe[df_rlbooe[0].isin(values_to_move)]
        df_rlbooe = df_rlbooe[~df_rlbooe[0].isin(values_to_move)]
        df_kifi = pd.concat([df_kifi, rows_to_move], ignore_index=True)

        # Convert date columns to strings in the desired format
        date_format = '%d.%b.%y'
        df_rlbooe[34] = pd.to_datetime(df_rlbooe[34], errors='coerce').dt.strftime(date_format)
        df_rlbooe[40] = pd.to_datetime(df_rlbooe[40], errors='coerce').dt.strftime(date_format)

        # Function to copy data to sections
        def copy_data_to_sections(source_df, target_df):
            target_df = target_df.reindex(range(len(source_df)), fill_value=None)
            target_df.iloc[:, :source_df.shape[1]] = source_df.values

            specific_data_to_copy = source_df.iloc[:, :21]
            target_df = target_df.reindex(columns=range(max(66, target_df.shape[1])), fill_value=None)

            target_df.iloc[:, 43:56] = specific_data_to_copy.iloc[:, :13].values
            target_df.iloc[:, 57:59] = specific_data_to_copy.iloc[:, 13:15].values
            target_df.iloc[:, 60:66] = specific_data_to_copy.iloc[:, 15:21].values

            return target_df

        # Function to perform the VLOOKUP-like operation
        def perform_vlookup_and_update(target_df, mapping_df):
            lookup_dict = mapping_df.set_index(0)[2].to_dict()
            target_df[52] = target_df[9].map(lookup_dict).fillna("kein Mapping Vorhanden")
            return target_df

        # Function to perform the additional check and update
        def check_and_update_bd_column(target_df, mapping_df):
            lookup_dict = mapping_df.set_index(2)[3].to_dict()
            target_df[55] = target_df[52].map(lookup_dict).fillna("kein Mapping Vorhanden")
            return target_df

        # Function to update column BN based on column BD
        def update_bn_column(target_df):
            target_df[65] = target_df.apply(lambda row: row[20] if row[55] != "kein Mapping Vorhanden" else "kein Mapping Vorhanden", axis=1)
            return target_df

        # Function to update column BB based on column K and Umgliederung sheet
        def update_bb_column(target_df, specific_value):
            target_df[53] = specific_value
            return target_df

        # Function to update column BE based on column BF
        def update_be_column(target_df):
            target_df[56] = target_df[57].apply(lambda x: "H" if x in [0, 0.00, 0.0000, "0,00"] else "S")
            return target_df

        # Apply the function to the relevant sheet
        df_rlbooe_mit_teilkonzerne = copy_data_to_sections(df_rlbooe, df_rlbooe_mit_teilkonzerne)

        # Perform VLOOKUP and update
        df_rlbooe_mit_teilkonzerne = perform_vlookup_and_update(df_rlbooe_mit_teilkonzerne, df_rlz_mapping)

        # Perform the additional check and update
        df_rlbooe_mit_teilkonzerne = check_and_update_bd_column(df_rlbooe_mit_teilkonzerne, df_rlz_mapping)

        df_rlbooe_mit_teilkonzerne = update_bn_column(df_rlbooe_mit_teilkonzerne)

        # Update column BB based on column K and Umgliederung sheet
        df_rlbooe_mit_teilkonzerne = update_bb_column(df_rlbooe_mit_teilkonzerne, "B100")

        # Update column BE based on column BF
        df_rlbooe_mit_teilkonzerne = update_be_column(df_rlbooe_mit_teilkonzerne)

        # Function to cut and paste rows based on specific values in a column
        def cut_and_paste_rows_by_values(source_df, target_df, values):
            rows_to_move = source_df[source_df[0].isin(values)]
            source_df = source_df[~source_df[0].isin(values)]
            target_df = pd.concat([target_df, rows_to_move], ignore_index=True)
            return source_df, target_df

        # Define the values to cut and paste for the sheet
        values_rlbooe = {'KIFI', 'ILG', 'ILI'}

        # Cut and paste rows in the DataFrame
        df_rlbooe_mit_teilkonzerne, df_kifi_mit_teilkonzerne = cut_and_paste_rows_by_values(df_rlbooe_mit_teilkonzerne, df_kifi_mit_teilkonzerne, values_rlbooe)

        # Function to filter rows based on column AG (index 32) being 5
        def filter_rows_based_on_ag(df):
            df[32] = pd.to_numeric(df[32], errors='coerce')
            return df[df[32] == 5]

        # Apply the filtering to the DataFrame
        df_rlbooe_mit_teilkonzerne = filter_rows_based_on_ag(df_rlbooe_mit_teilkonzerne)

        # Convert specified columns to European format
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
                df[column] = df[column].apply(convert_to_european_format)
            return df

        columns_to_convert = [13, 14, 15, 57, 58, 60]
        df_rlbooe_mit_teilkonzerne = apply_european_formatting_to_columns(df_rlbooe_mit_teilkonzerne, columns_to_convert)

        # Apply the same processing steps to the KIFI sheet
        df_kifi_mit_teilkonzerne = copy_data_to_sections(df_kifi, df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = perform_vlookup_and_update(df_kifi_mit_teilkonzerne, df_rlz_mapping)
        df_kifi_mit_teilkonzerne = check_and_update_bd_column(df_kifi_mit_teilkonzerne, df_rlz_mapping)
        df_kifi_mit_teilkonzerne = update_bn_column(df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = update_bb_column(df_kifi_mit_teilkonzerne, "B100")
        df_kifi_mit_teilkonzerne = update_be_column(df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = filter_rows_based_on_ag(df_kifi_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = apply_european_formatting_to_columns(df_kifi_mit_teilkonzerne, columns_to_convert)

        # Function to filter rows based on column D and E values
        def filter_rows(df):
            return df[(df[3] != 'SK') | (df[4] != '99')]

        # Apply the filter to both sheets
        df_rlbooe_mit_teilkonzerne = filter_rows(df_rlbooe_mit_teilkonzerne)
        df_kifi_mit_teilkonzerne = filter_rows(df_kifi_mit_teilkonzerne)

        # Load the workbook to preserve formatting
        workbook = load_workbook(file_path)

        # Function to write data back to the sheet
        def write_data_to_sheet(sheet_name, df):
            print(f"Writing to sheet: {sheet_name}")
            sheet = workbook[sheet_name]
            for row_idx, row in enumerate(df.itertuples(index=False, name=None), start=4):  # Start from row 4
                for col_idx, value in enumerate(row, start=1):
                    cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                    # Apply European number format to specific columns
                    if col_idx in [14, 15, 16, 58, 59, 61]:  # N, O, BF, BG
                        cell.number_format = '#,##0.00'

        # Write the modified data back to the sheets
        write_data_to_sheet('RLB mit Teilkonzerne', df_rlbooe_mit_teilkonzerne)
        write_data_to_sheet('KIFI', df_kifi)
        write_data_to_sheet('Kifi mit Teilkonzerne', df_kifi_mit_teilkonzerne)

        # Function to auto-adjust column width
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

        # Adjust column widths for better readability
        auto_adjust_column_width(workbook['RLB mit Teilkonzerne'])
        auto_adjust_column_width(workbook['Kifi mit Teilkonzerne'])

        # Save the workbook
        workbook.save(file_path)

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Erfolgreich")
        root.destroy()

    except Exception as e:
        print(f"An error occurred during execution: {e}")           in this code , how is it copying data from RLBOOE to KIFI sheet , means from which row of the rlbooe to which row of the kifi sheet ?
