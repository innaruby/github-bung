from openpyxl.utils import get_column_letter
from datetime import datetime

def apply_number_format_to_ist_and_plan_columns(ws, end_row):
    """
    This function identifies visible columns with:
    - 'IST' in row 3 and current year in row 4
    - 'PLAN' in row 3 and current year + 1 in row 4

    It then applies number formatting:
    - If value >= 1000 → format as '#,##0.000'
    - Otherwise → format as '0'
    """
    print(f"\n🎯 Running final formatting for sheet: {ws.title}")
    current_year = datetime.now().year

    def is_column_visible(col):
        return not ws.column_dimensions[get_column_letter(col)].hidden

    ist_col = None
    plan_col = None

    for col in range(2, ws.max_column + 1):
        if not is_column_visible(col):
            continue
        header3 = str(ws.cell(row=3, column=col).value or "").strip().upper()
        header4 = str(ws.cell(row=4, column=col).value or "").replace("e", "").strip()
        if header3 == "IST" and header4 == str(current_year):
            ist_col = col
        if header3 == "PLAN" and header4 == str(current_year + 1):
            plan_col = col

    if not ist_col and not plan_col:
        print("❌ No matching IST or PLAN columns found with expected year headers.")
        return

    for row in range(5, end_row + 1):
        for target_col in [ist_col, plan_col]:
            if target_col:
                cell = ws.cell(row=row, column=target_col)
                try:
                    val = float(cell.value)
                    if val >= 1000:
                        cell.number_format = "#,##0.000"
                    else:
                        cell.number_format = "0"
                except:
                    pass
    print("✅ Final number formatting applied to IST and PLAN columns.")
