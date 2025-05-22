from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

def remove_vertical_border_between_ist_and_plan(ws):
    current_year = datetime.now().year
    ist_col = None
    plan_col = None

    # Identify IST and PLAN columns based on row 3 & 4 headers
    for col in range(1, ws.max_column + 1):
        val3 = str(ws.cell(row=3, column=col).value).strip().upper()
        val4 = str(ws.cell(row=4, column=col).value).replace("e", "").strip()
        if val3 == "IST" and val4 == str(current_year):
            ist_col = col
        if val3 == "PLAN" and val4 == str(current_year + 1):
            plan_col = col

    if ist_col is None or plan_col is None:
        print(f"❌ IST or PLAN columns not found (Year {current_year})")
        return

    print(f"🎯 Found IST column: {get_column_letter(ist_col)} | PLAN column: {get_column_letter(plan_col)}")

    # Remove vertical border between IST and PLAN in rows 3 and 4
    thin_border_none = Border(
        left=Side(style=None), right=Side(style=None),
        top=Side(style=None), bottom=Side(style=None)
    )

    # Row 3
    ist_cell_r3 = ws.cell(row=3, column=ist_col)
    plan_cell_r3 = ws.cell(row=3, column=plan_col)

    ist_cell_r3.border = Border(
        left=ist_cell_r3.border.left,
        right=Side(style=None),  # Remove right border
        top=ist_cell_r3.border.top,
        bottom=ist_cell_r3.border.bottom
    )

    plan_cell_r3.border = Border(
        left=Side(style=None),  # Remove left border
        right=plan_cell_r3.border.right,
        top=plan_cell_r3.border.top,
        bottom=plan_cell_r3.border.bottom
    )

    # Row 4
    ist_cell_r4 = ws.cell(row=4, column=ist_col)
    plan_cell_r4 = ws.cell(row=4, column=plan_col)

    ist_cell_r4.border = Border(
        left=ist_cell_r4.border.left,
        right=Side(style=None),
        top=ist_cell_r4.border.top,
        bottom=ist_cell_r4.border.bottom
    )

    plan_cell_r4.border = Border(
        left=Side(style=None),
        right=plan_cell_r4.border.right,
        top=plan_cell_r4.border.top,
        bottom=plan_cell_r4.border.bottom
    )

    print(f"✅ Removed vertical border between IST and PLAN headers.")


-----------------------------------------------------------------------------------------------------------


from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side

def remove_border_between_ist_plan(ws, current_year):
    ist_col = None
    plan_col = None

    # Find IST and PLAN columns from visible columns
    visible_cols = [
        col for col in range(1, ws.max_column + 1)
        if not ws.column_dimensions[get_column_letter(col)].hidden
    ]

    for col in visible_cols:
        header_3 = str(ws.cell(row=3, column=col).value or "").strip().upper()
        header_4 = str(ws.cell(row=4, column=col).value or "").replace("e", "").strip()
        if header_3 == "IST" and header_4 == str(current_year):
            ist_col = col
        elif header_3 == "PLAN" and header_4 == str(current_year + 1):
            plan_col = col

    if ist_col is None or plan_col is None:
        print("❌ Could not identify both IST and PLAN columns.")
        return

    print(f"🎯 Found IST column: {get_column_letter(ist_col)}, PLAN column: {get_column_letter(plan_col)}")

    no_border = Side(style=None)

    # Remove only the border between row 3 and 4, for IST and PLAN columns
    # Remove right border of IST (row 3 and 4)
    for row in [3, 4]:
        ist_cell = ws.cell(row=row, column=ist_col)
        ist_cell.border = Border(
            left=ist_cell.border.left,
            right=no_border,
            top=ist_cell.border.top,
            bottom=ist_cell.border.bottom
        )

    # Remove left border of PLAN (row 3 and 4)
    for row in [3, 4]:
        plan_cell = ws.cell(row=row, column=plan_col)
        plan_cell.border = Border(
            left=no_border,
            right=plan_cell.border.right,
            top=plan_cell.border.top,
            bottom=plan_cell.border.bottom
        )

remove_vertical_border_between_ist_and_plan(ws)

----------------------------------------------------------------------------------------------------------

from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side

def remove_border_between_ist_plan(ws, current_year):
    ist_col = None
    plan_col = None

    # Find IST and PLAN columns from visible columns
    visible_cols = [
        col for col in range(1, ws.max_column + 1)
        if not ws.column_dimensions[get_column_letter(col)].hidden
    ]

    for col in visible_cols:
        header_3 = str(ws.cell(row=3, column=col).value or "").strip().upper()
        header_4 = str(ws.cell(row=4, column=col).value or "").replace("e", "").strip()
        if header_3 == "IST" and header_4 == str(current_year):
            ist_col = col
        elif header_3 == "PLAN" and header_4 == str(current_year + 1):
            plan_col = col

    if ist_col is None or plan_col is None:
        print("❌ Could not identify both IST and PLAN columns.")
        return

    print(f"🎯 Found IST column: {get_column_letter(ist_col)}, PLAN column: {get_column_letter(plan_col)}")

    no_border = Side(style=None)

    # Remove only the border between row 3 and 4, for IST and PLAN columns
    # Remove right border of IST (row 3 and 4)
    for row in [3, 4]:
        ist_cell = ws.cell(row=row, column=ist_col)
        ist_cell.border = Border(
            left=ist_cell.border.left,
            right=no_border,
            top=ist_cell.border.top,
            bottom=ist_cell.border.bottom
        )

    # Remove left border of PLAN (row 3 and 4)
    for row in [3, 4]:
        plan_cell = ws.cell(row=row, column=plan_col)
        plan_cell.border = Border(
            left=no_border,
            right=plan_cell.border.right,
            top=plan_cell.border.top,
            bottom=plan_cell.border.bottom
        )

remove_border_between_ist_plan(ws, current_year)


