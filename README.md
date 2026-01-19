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
OUTPUT_FILE = "processed_decrypted.xlsx"

EXCEPTION_SHEET = "EAD_RWA exception"

list_ead = ["EAD", "Exposure"]
list_rwa = ["RWA", "Risk"]

# list_rep dataframe must already exist
# columns: Sheet_Name, Check, EAD, RWA

# ================= DECRYPT =================
tmp = NamedTemporaryFile(delete=False, suffix=".xlsx")
tmp.close()

with open(INPUT_FILE, "rb") as f:
    office = msoffcrypto.OfficeFile(f)
    office.load_key(password=PASSWORD)
    with open(tmp.name, "wb") as d:
        office.decrypt(d)

# ================= LOAD =================
wb = load_workbook(tmp.name)

# ================= EXCEPTION SHEET =================
if EXCEPTION_SHEET in wb.sheetnames:
    del wb[EXCEPTION_SHEET]

exc_ws = wb.create_sheet(EXCEPTION_SHEET)
exc_ws.append([
    "EADzero_RWAnonzero", "Amount_RWA", "", "", "",
    "EADnonzero_RWAzero", "Amount_EAD"
])

exc_row = 2

# ================= PROCESS =================
for _, rep in list_rep.iterrows():

    if str(rep["Check"]).upper() != "Y":
        continue

    sheet = rep["Sheet_Name"]
    if sheet not in wb.sheetnames:
        continue

    ws = wb[sheet]

    ead_col_letter = rep["EAD"]
    rwa_col_letter = rep["RWA"]

    ead_col = column_index_from_string(ead_col_letter)
    rwa_col = column_index_from_string(rwa_col_letter)

    start_row = None

    # ---- find header row ----
    for r in range(1, ws.max_row + 1):
        if (
            any(x in str(ws.cell(r, ead_col).value or "") for x in list_ead)
            and any(x in str(ws.cell(r, rwa_col).value or "") for x in list_rwa)
        ):
            start_row = r + 1
            break

    if not start_row:
        continue

    ead_vals, rwa_vals = [], []
    ead_nonzero = None
    rwa_nonzero = None

    # ---- scan numeric values ----
    for r in range(start_row, ws.max_row + 1):
        ead_v = ws.cell(r, ead_col).value
        rwa_v = ws.cell(r, rwa_col).value

        if isinstance(ead_v, (int, float)):
            ead_vals.append(ead_v)
            if ead_v != 0 and not ead_nonzero:
                ead_nonzero = (r, ead_v)

        if isinstance(rwa_v, (int, float)):
            rwa_vals.append(rwa_v)
            if rwa_v != 0 and not rwa_nonzero:
                rwa_nonzero = (r, rwa_v)

    if not ead_vals or not rwa_vals:
        continue

    ead_all_zero = all(v == 0 for v in ead_vals)
    rwa_all_zero = all(v == 0 for v in rwa_vals)

    # ================= CONDITIONS =================

    # ---- Case 1: EAD zero, RWA non-zero ----
    if ead_all_zero and rwa_nonzero:
        r, amount = rwa_nonzero
        link_cell = f"#{sheet}!{ead_col_letter}{start_row}"

        exc_ws.cell(exc_row, 1).value = "Link"
        exc_ws.cell(exc_row, 1).hyperlink = link_cell
        exc_ws.cell(exc_row, 2).value = amount
        exc_row += 1

    # ---- Case 2: RWA zero, EAD non-zero ----
    if rwa_all_zero and ead_nonzero:
        r, amount = ead_nonzero
        link_cell = f"#{sheet}!{rwa_col_letter}{start_row}"

        exc_ws.cell(exc_row, 6).value = "Link"
        exc_ws.cell(exc_row, 6).hyperlink = link_cell
        exc_ws.cell(exc_row, 7).value = amount
        exc_row += 1

# ================= SAVE =================
wb.save(tmp.name)
os.replace(tmp.name, OUTPUT_FILE)

print("Hyperlinks created successfully.")

