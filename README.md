```
import xlwings as xw
import pandas as pd
import os

# ================= USER CONFIG =================
INPUT_FILE = r"C:\path\to\protected.xlsx"
OUTPUT_FILE = r"C:\path\to\protected_UPDATED.xlsx"
PASSWORD = "your_password"

EXCEPTION_SHEET = "EAD_RWA exception"

list_ead = ["EAD", "Exposure"]
list_rwa = ["RWA", "Risk"]

# list_rep dataframe must exist
# Required columns: Sheet_Name, Check, EAD, RWA

# ================= OPEN EXCEL =================
app = xw.App(visible=False)
app.screen_updating = False
app.display_alerts = False
app.enable_events = False
app.calculation = 'manual'

wb = app.books.open(INPUT_FILE, password=PASSWORD)

# ================= CREATE EXCEPTION SHEET FIRST =================
try:
    wb.sheets[EXCEPTION_SHEET].delete()
except:
    pass

exc_ws = wb.sheets.add(EXCEPTION_SHEET, before=wb.sheets[0])

exc_ws.range("A1").value = [
    "EADzero_RWAnonzero", "Amount_RWA", "", "", "",
    "EADnonzero_RWAzero", "Amount_EAD"
]

exc_row = 2

# ================= PROCESS =================
sheet_names = {s.name for s in wb.sheets}

for _, rep in list_rep.iterrows():

    if str(rep["Check"]).upper() != "Y":
        continue

    sheet_name = rep["Sheet_Name"]
    if sheet_name not in sheet_names:
        continue

    ws = wb.sheets[sheet_name]

    ead_col = rep["EAD"]
    rwa_col = rep["RWA"]

    last_row = ws.used_range.last_cell.row

    # BULK READ
    ead_vals = ws.range(f"{ead_col}1:{ead_col}{last_row}").value
    rwa_vals = ws.range(f"{rwa_col}1:{rwa_col}{last_row}").value

    if not ead_vals or not rwa_vals:
        continue

    # Find header row
    start_row = None
    for i, (e, r) in enumerate(zip(ead_vals, rwa_vals)):
        if (
            isinstance(e, str) and any(x in e for x in list_ead)
            and isinstance(r, str) and any(x in r for x in list_rwa)
        ):
            start_row = i + 2
            break

    if not start_row:
        continue

    ead_nums, rwa_nums = [], []
    ead_nonzero, rwa_nonzero = None, None

    for i in range(start_row - 1, len(ead_vals)):
        e, r = ead_vals[i], rwa_vals[i]

        if isinstance(e, (int, float)):
            ead_nums.append(e)
            if e != 0 and ead_nonzero is None:
                ead_nonzero = e

        if isinstance(r, (int, float)):
            rwa_nums.append(r)
            if r != 0 and rwa_nonzero is None:
                rwa_nonzero = r

    if not ead_nums or not rwa_nums:
        continue

    ead_all_zero = all(v == 0 for v in ead_nums)
    rwa_all_zero = all(v == 0 for v in rwa_nums)

    # WRITE RESULTS
    if ead_all_zero and rwa_nonzero is not None:
        exc_ws.range(f"A{exc_row}").add_hyperlink(
            f"'{sheet_name}'!{ead_col}{start_row}",
            text_to_display="Link"
        )
        exc_ws.range(f"B{exc_row}").value = rwa_nonzero
        exc_row += 1

    if rwa_all_zero and ead_nonzero is not None:
        exc_ws.range(f"F{exc_row}").add_hyperlink(
            f"'{sheet_name}'!{rwa_col}{start_row}",
            text_to_display="Link"
        )
        exc_ws.range(f"G{exc_row}").value = ead_nonzero
        exc_row += 1

# ================= FORCE SAVE =================
wb.api.SaveAs(OUTPUT_FILE)

wb.close()
app.calculation = 'automatic'
app.quit()

print(f"Workbook successfully updated: {OUTPUT_FILE}")

