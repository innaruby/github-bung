import re
from openpyxl.utils import column_index_from_string

def parse_numeric(val, ws, current_row=None, current_col=None):
    if val is None or val == "":
        return 0
    if isinstance(val, (int, float)):
        return val

    if isinstance(val, str):
        original_expr = val
        val = val.strip()
        if val.startswith("="):
            val = val[1:]

        # Normalize decimal commas
        val = val.replace(",", ".")

        # Extract all Excel-style references (e.g., T9, AB10)
        references = re.findall(r'\b([A-Za-z]{1,3}\d{1,4})\b', val)

        for ref in references:
            try:
                col_letters = re.match(r'[A-Za-z]+', ref).group()
                row_digits = re.match(r'[A-Za-z]+(\d+)', ref).group(1)
                col_idx = column_index_from_string(col_letters)
                row_idx = int(row_digits)

                # Safety check: avoid self-reference
                if current_row and current_col:
                    if row_idx == current_row and col_idx == current_col:
                        val = val.replace(ref, "0")
                        continue

                ref_cell = ws.cell(row=row_idx, column=col_idx)
                ref_val = parse_numeric(ref_cell.value, ws)  # Recursive parse
                val = val.replace(ref, str(ref_val))
            except Exception as e:
                print(f"⚠️ Could not resolve reference '{ref}' in expression '{original_expr}': {e}")
                val = val.replace(ref, "0")

        # Remove any remaining bad characters
        val = re.sub(r'[^\d\.\+\-\*/\(\)\s]', '', val)

        try:
            result = eval(val)
            return float(result)
        except Exception as e:
            print(f"⚠️ Skipping bad expression after resolving refs: '{val}' → {e}")
            return 0

    try:
        return float(val)
    except:
        return 0
