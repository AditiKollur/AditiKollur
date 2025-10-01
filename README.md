```
import re
import pandas as pd
from openpyxl import load_workbook

# Example dataframe (replace with your actual one)
df = pd.DataFrame({
    "custid": [101, 102, 103],
    "gho": ["A", "B", "C"],
    "att": ["X", "Y", "Z"],
    "prd": ["P1", "P2", "P3"],
    "values": [100, 200, 300]
})

# Load workbook and target sheet
wb = load_workbook("your_file.xlsx")
ws = wb["Gb"]

# Define mapping: DataFrame column → Excel column letter
col_map = {
    "custid": "A",
    "gho": "C",
    "att": "D",
    "prd": "E",
    "values": "F"
}

start_row = 18
last_data_row = start_row + len(df) - 1

# 1️⃣ Clear old data in A–M from row 18 downward (full wipe)
for row in range(start_row, ws.max_row + 1):
    for col in range(1, 14):  # A (1) to M (13)
        ws.cell(row=row, column=col).value = None

# 2️⃣ Write DataFrame values into A, C, D, E, F
for i, row in df.iterrows():
    excel_row = start_row + i
    for col_name, col_letter in col_map.items():
        ws[f"{col_letter}{excel_row}"] = row[col_name]

# 3️⃣ Drag formulas from row 18 for columns B, G–M (relative references)
formula_cols = ["B", "G", "H", "I", "J", "K", "L", "M"]

for col in formula_cols:
    template_formula = ws[f"{col}{start_row}"].value
    if template_formula and isinstance(template_formula, str) and template_formula.startswith("="):

        for r in range(start_row, last_data_row + 1):
            if r == start_row:
                # Keep original formula in row 18
                continue  

            # Adjust row references inside formula to simulate Excel drag
            new_formula = re.sub(
                r"(\d+)", 
                lambda m: str(int(m.group(1)) + (r - start_row)), 
                template_formula
            )

            ws[f"{col}{r}"].value = new_formula

# 4️⃣ (Optional) Clear any rows below last_data_row in A–M so sheet ends at new data
for row in range(last_data_row + 1, ws.max_row + 1):
    for col in range(1, 14):  # A–M
        ws.cell(row=row, column=col).value = None

# Save refreshed file
wb.save("your_file_refreshed.xlsx")

```
