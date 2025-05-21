import re
from openpyxl.utils import get_column_letter

def parse_numeric(val, ws=None, row=None, col=None, ignored_tokens_log=None):
    if val is None or val == "":
        return 0
    if isinstance(val, (int, float)):
        return val

    original_val = val
    if isinstance(val, str):
        val = val.strip()
        if val.startswith("="):
            val = val[1:]

        val = val.replace(",", ".")
        ignored_refs = re.findall(r'\b([A-Za-z]{1,3}\d{1,4})\b', val)

        for ref in ignored_refs:
            val = val.replace(ref, "")

        val = re.sub(r'[^\d\.\+\-\*/\(\)\s]', '', val)

        try:
            result = eval(val)
            if ignored_refs and ignored_tokens_log is not None and ws and row and col:
                sheet_name = ws.title
                cell_ref = f"{get_column_letter(col)}{row}"
                ignored_tokens_log.append((sheet_name, cell_ref, original_val, ignored_refs))
            return float(result)
        except Exception as e:
            print(f"⚠️ Failed to evaluate expression '{original_val}' → {e}")
            return 0

    try:
        return float(val)
    except:
        return 0

def report_ignored_tokens(ignored_tokens_log):
    if not ignored_tokens_log:
        print("\n✅ No ignored tokens found.")
        return

    print("\n⚠️ Ignored Tokens Summary:")
    for sheet, cell, original_expr, tokens in ignored_tokens_log:
        print(f"  🔎 Sheet '{sheet}', Cell {cell}: Ignored tokens {tokens} in expression '{original_expr}'")

report_ignored_tokens(ignored_tokens_log)
parsed = parse_numeric(cell.value, ws, r, col, ignored_tokens_log)

ignored_tokens_log = []
