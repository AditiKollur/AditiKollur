```
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ---------------- Example DataFrame (replace with your actual one) ----------------
df = pd.DataFrame({
    "custid": [101, 102, 103],
    "gho": ["A", "B", "C"],
    "att": ["X", "Y", "Z"],
    "prd": ["P1", "P2", "P3"],
    "values": [100, 200, 300]
})
# -------------------------------------------------------------------------------

# ---------------- File and sheet names ----------------
input_file = "your_file.xlsx"
output_file = "your_file_refreshed.xlsx"
sheet_name = "Gb"
# -------------------------------------------------------------------------------

# Load workbook and sheet
wb = load_workbook(input_file)
ws = wb[sheet_name]

start_row = 18
nrows = len(df)
last_data_row = start_row + nrows - 1 if nrows > 0 else start_row - 1

# Columns to replace with df
col_map = {"custid": "A", "gho": "C", "att": "D", "prd": "E", "values": "F"}

# Formula columns to drag
formula_cols = ["B", "G", "H", "I", "J", "K", "L", "M"]

# Regex to detect A1-style references (e.g. A18, $A$18, A$18)
cell_ref_re = re.compile(r'(\$?[A-Za-z]{1,3})(\$?)(\d+)')

def shift_formula_rows(formula: str, row_offset: int) -> str:
    """
    Shift only A1-style cell references by `row_offset`.
    Absolute rows (with $) remain unchanged.
    """
    def _repl(m):
        col_part = m.group(1)      # e.g. A or $A
        row_dollar = m.group(2)    # '' or '$'
        rownum = int(m.group(3))
        if row_dollar == '$':      # absolute row -> unchanged
            new_row = rownum
        else:
            new_row = rownum + row_offset
        return f"{col_part}{row_dollar}{new_row}"
    return cell_ref_re.sub(_repl, formula)


# 1️⃣ Save template formulas from row 18 before clearing
templates = {}
for col in formula_cols:
    templates[col] = ws[f"{col}{start_row}"].value

# 2️⃣ Clear old data
max_row = ws.max_row
for r in range(start_row, max_row + 1):
    for col_idx in range(1, 14):  # A–M
        col_letter = get_column_letter(col_idx)
        if col_letter in ["A", "C", "D", "E", "F"]:
            ws.cell(row=r, column=col_idx).value = None
        elif col_letter in formula_cols and r > start_row:
            ws.cell(row=r, column=col_idx).value = None

# 3️⃣ Write DataFrame values into A, C, D, E, F
for i, row in df.iterrows():
    excel_row = start_row + i
    for col_name, col_letter in col_map.items():
        ws[f"{col_letter}{excel_row}"] = row[col_name]

# 4️⃣ Ensure template formulas remain in row 18
for col in formula_cols:
    tpl = templates.get(col)
    if tpl is not None:
        ws[f"{col}{start_row}"].value = tpl

# 5️⃣ Drag formulas down to last_data_row
for col in formula_cols:
    tpl = templates.get(col)
    if isinstance(tpl, str) and tpl.startswith("=") and nrows > 0:
        for r in range(start_row + 1, last_data_row + 1):
            offset = r - start_row
            ws[f"{col}{r}"].value = shift_formula_rows(tpl, offset)

# 6️⃣ Clear rows below last_data_row (A–M only)
if last_data_row < max_row:
    for r in range(last_data_row + 1, max_row + 1):
        for col_idx in range(1, 14):
            ws.cell(row=r, column=col_idx).value = None

# Save refreshed workbook
wb.save(output_file)
print(f"✅ Sheet refreshed and saved to {output_file}")
```
