```
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string

# ================= CONFIG =================
WORKBOOK_PATH = "input_workbook.xlsx"
REFERENCE_PATH = "reference.xlsx"
REFERENCE_SHEET = "Sheet1"

# Tab colors
AMBER_COLOR = "FFC000"
BLUE_COLOR = "00B0F0"

# ================= LOAD =================
ref_df = pd.read_excel(REFERENCE_PATH, sheet_name=REFERENCE_SHEET)
wb = load_workbook(WORKBOOK_PATH, data_only=True)

# ================= FUNCTIONS =================
def get_column_sum_by_letter(ws, col_letter):
    """
    Reads numeric values from a column using Excel column letter.
    Does NOT alter formulas or formatting.
    """
    col_idx = column_index_from_string(col_letter)
    total = 0

    for row in range(2, ws.max_row + 1):  # assuming row 1 is header
        val = ws.cell(row=row, column=col_idx).value
        if isinstance(val, (int, float)):
            total += val

    return total

# ================= PROCESS =================
for _, row in ref_df.iterrows():
    tab_name = row["Tab name"]
    check_flag = row["Check column"]
    ead_col_letter = str(row["EAD Column"]).strip()
    rwa_col_letter = str(row["RWA Column"]).strip()

    if check_flag != "Y":
        continue

    if tab_name not in wb.sheetnames:
        continue

    ws = wb[tab_name]

    ead_value = get_column_sum_by_letter(ws, ead_col_letter)
    rwa_value = get_column_sum_by_letter(ws, rwa_col_letter)

    # Apply coloring rules
    if ead_value == 0 and rwa_value != 0:
        ws.sheet_properties.tabColor = AMBER_COLOR
    elif ead_value != 0 and rwa_value == 0:
        ws.sheet_properties.tabColor = BLUE_COLOR

# ================= SAVE =================
wb.save("output_workbook.xlsx")
