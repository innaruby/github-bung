structured = extract_aligned_table(ocr_data)
df = pd.DataFrame(structured)

# Replace empty strings with NaN, drop empty columns, convert pd.NA to None
df = df.replace('', pd.NA).dropna(how='all', axis=1)
df = df.where(pd.notnull(df), None)  # ✅ fix for openpyxl write issue

sheet_name = f"Page_{i+1}"
ws = wb.create_sheet(title=sheet_name)
for row in dataframe_to_rows(df, index=False, header=False):
    ws.append(row)
