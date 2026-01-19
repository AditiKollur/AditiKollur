```
import pandas as pd
import msoffcrypto
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from tempfile import NamedTemporaryFile
import os

# ================= CONFIG =================
INPUT_FILE = "protected.xlsx"
PASSWORD = "your_password"

EXCEPTION_SHEET = "EAD_RWA exception"

list_ead = ["EAD", "Exposure"]
list_rwa = ["RWA", "Risk"]

# list_rep dataframe must exist
# columns: Sheet_Name, Check, EAD, RWA

# ================= DECRYPT FILE =================
tmp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
tmp_file.close()

with open(INPUT_FILE, "rb") as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=PASSWORD)
    with open(tmp_file.name, "wb") as decrypted:
        office_file.decrypt(decrypted)

# ================= LOAD WORKBOOK =================
wb = load_workbook(tmp_file.name)

# ================= CREATE / RESET EXCEPTION SHEET =================
if EXCEPTION_SHEET in wb.sheetnames:
    del wb[EXCEPTION_SHEET]

exc_ws = wb.create_sheet(EXCEPTION_SHEET)

# Header with 3 empty columns after every 2 columns
exc_ws.append([
    "EADzero_RWAnonzero", "Amount_RWA", "", "", "",
    "EADnonzero_RWAzero", "Amount_EAD"
])

exc_row = 2

# ================= PROCESS EACH SHEET =================
for _, row in list_rep.iterrows():

    if str(row["Check"]).upper() != "Y":
        continue

    sheet_name = row["Sheet_Name"]
    if sheet_name not in wb.sheetnames:
        continue

    ws = wb[sheet_name]

    ead_col_letter = row["EAD"]
    rwa_col_letter = row["RWA"]

    ead_col = column_index_from_string(ead_col_letter)
    rwa_col = column_index_from_string(rwa_col_letter)

    start_row = None

    # ---- Find first row where both EAD & RWA strings exist ----
    for r in range(1, ws.max_row + 1):
        ead_txt = str(ws.cell(r, ead_col).value or "")
        rwa_txt = str(ws.cell(r, rwa_col).value or "")

        if any(x in ead_txt for x in list_ead) and any(x in rwa_txt for x in list_rwa):
            start_row = r + 1
            break

    if not start_row:
        continue

    ead_vals, rwa_vals = [], []
    ead_cell, rwa_cell = None, None

    # ---- Scan numeric values below header ----
    for r in range(start_row, ws.max_row + 1):
        ead_v = ws.cell(r, ead_col).value
        rwa_v = ws.cell(r, rwa_col).value

        if isinstance(ead_v, (int, float)):
            ead_vals.append(ead_v)
            if ead_v != 0 and not ead_cell:
                ead_cell = ws.cell(r, ead_col)

        if isinstance(rwa_v, (int, float)):
            rwa_vals.append(rwa_v)
            if rwa_v != 0 and not rwa_cell:
                rwa_cell = ws.cell(r, rwa_col)

    if not ead_vals or not rwa_vals:
        continue

    ead_all_zero = all(v == 0 for v in ead_vals)
    rwa_all_zero = all(v == 0 for v in rwa_vals)

    # ================= CONDITIONS =================

    # ---- Case 1: EAD zero, RWA non-zero ----
    if ead_all_zero and any(v != 0 for v in rwa_vals):
        link = f"#{sheet_name}!{ead_col_letter}{ead_cell.row}"
        exc_ws.cell(exc_row, 1).value = "Link"
        exc_ws.cell(exc_row, 1).hyperlink = link
        exc_ws.cell(exc_row, 2).value = rwa_cell.value
        exc_row += 1

    # ---- Case 2: RWA zero, EAD non-zero ----
    if rwa_all_zero and any(v != 0 for v in ead_vals):
        link = f"#{sheet_name}!{rwa_col_letter}{rwa_cell.row}"
        exc_ws.cell(exc_row, 6).value = "Link"
        exc_ws.cell(exc_row, 6).hyperlink = link
        exc_ws.cell(exc_row, 7).value = ead_cell.value
        exc_row += 1

# ================= SAVE RESULT =================
wb.save(tmp_file.name)

# ================= REPLACE ORIGINAL FILE =================
os.replace(tmp_file.name, "processed_decrypted.xlsx")

print("Processing completed successfully.")
