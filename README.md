```
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.hyperlink import Hyperlink

# ================= CONFIG =================
WORKBOOK_PATH = "input.xlsx"
EXCEPTION_SHEET = "EAD_RWA exception"

list_ead = ["EAD", "Exposure"]
list_rwa = ["RWA", "Risk"]

# list_rep dataframe assumed already loaded
# columns: Sheet_Name, Check, EAD, RWA

# ================= LOAD WORKBOOK =================
wb = load_workbook(WORKBOOK_PATH)

# ================= CREATE / RESET EXCEPTION SHEET =================
if EXCEPTION_SHEET in wb.sheetnames:
    del wb[EXCEPTION_SHEET]

exc_ws = wb.create_sheet(EXCEPTION_SHEET)

headers = [
    "EADzero_RWAnonzero", "Amount_RWA", "", "", "",
    "EADnonzero_RWAzero", "Amount_EAD"
]
exc_ws.append(headers)

exc_row = 2

# ================= PROCESS SHEETS =================
for _, rep in list_rep.iterrows():
    if str(rep["Check"]).upper() != "Y":
        continue

    sheet_name = rep["Sheet_Name"]
    if sheet_name not in wb.sheetnames:
        continue

    ws = wb[sheet_name]

    ead_col_letter = rep["EAD"]
    rwa_col_letter = rep["RWA"]

    ead_col = column_index_from_string(ead_col_letter)
    rwa_col = column_index_from_string(rwa_col_letter)

    start_row = None

    # ---- Find first row where both strings exist ----
    for r in range(1, ws.max_row + 1):
        ead_val = str(ws.cell(r, ead_col).value or "")
        rwa_val = str(ws.cell(r, rwa_col).value or "")

        if any(x in ead_val for x in list_ead) and any(x in rwa_val for x in list_rwa):
            start_row = r + 1
            break

    if not start_row:
        continue

    ead_numbers = []
    rwa_numbers = []
    ead_cell = None
    rwa_cell = None

    # ---- Scan numeric values ----
    for r in range(start_row, ws.max_row + 1):
        ead_v = ws.cell(r, ead_col).value
        rwa_v = ws.cell(r, rwa_col).value

        if isinstance(ead_v, (int, float)):
            ead_numbers.append(ead_v)
            if ead_v != 0 and not ead_cell:
                ead_cell = ws.cell(r, ead_col)

        if isinstance(rwa_v, (int, float)):
            rwa_numbers.append(rwa_v)
            if rwa_v != 0 and not rwa_cell:
                rwa_cell = ws.cell(r, rwa_col)

    # ================= CONDITIONS =================
    ead_all_zero = ead_numbers and all(v == 0 for v in ead_numbers)
    rwa_all_zero = rwa_numbers and all(v == 0 for v in rwa_numbers)

    # ---- Case 1 ----
    if ead_all_zero and any(v != 0 for v in rwa_numbers):
        link = f"#{sheet_name}!{ead_col_letter}{ead_cell.row}"
        exc_ws.cell(exc_row, 1).value = "Link"
        exc_ws.cell(exc_row, 1).hyperlink = link
        exc_ws.cell(exc_row, 2).value = rwa_cell.value
        exc_row += 1

    # ---- Case 2 ----
    if rwa_all_zero and any(v != 0 for v in ead_numbers):
        link = f"#{sheet_name}!{rwa_col_letter}{rwa_cell.row}"
        exc_ws.cell(exc_row, 6).value = "Link"
        exc_ws.cell(exc_row, 6).hyperlink = link
        exc_ws.cell(exc_row, 7).value = ead_cell.value
        exc_row += 1

# ================= SAVE =================
wb.save(WORKBOOK_PATH)
