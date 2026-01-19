```
import xlwings as xw
import pandas as pd

# ================= CONFIG =================
FILE_PATH = r"C:\path\to\protected.xlsx"
PASSWORD = "your_password"

EXCEPTION_SHEET = "EAD_RWA exception"

list_ead = ["EAD", "Exposure"]
list_rwa = ["RWA", "Risk"]

# list_rep dataframe must exist
# columns: Sheet_Name, Check, EAD, RWA

# ================= OPEN EXCEL =================
app = xw.App(visible=False)
wb = app.books.open(FILE_PATH, password=PASSWORD)

# ================= CREATE EXCEPTION SHEET AS FIRST =================
try:
    wb.sheets[EXCEPTION_SHEET].delete()
except:
    pass

exc_ws = wb.sheets.add(EXCEPTION_SHEET, before=wb.sheets[0])

# Header with 3 empty columns after every 2 columns
exc_ws.range("A1").value = [
    "EADzero_RWAnonzero", "Amount_RWA", "", "", "",
    "EADnonzero_RWAzero", "Amount_EAD"
]

exc_row = 2

# ================= PROCESS EACH SHEET =================
for _, rep in list_rep.iterrows():

    if str(rep["Check"]).upper() != "Y":
        continue

    sheet_name = rep["Sheet_Name"]
    if sheet_name not in [s.name for s in wb.sheets]:
        continue

    ws = wb.sheets[sheet_name]

    ead_col = rep["EAD"]
    rwa_col = rep["RWA"]

    used = ws.used_range
    last_row = used.last_cell.row

    start_row = None

    # ---- Find header row ----
    for r in range(1, last_row + 1):
        ead_val = str(ws.range(f"{ead_col}{r}").value or "")
        rwa_val = str(ws.range(f"{rwa_col}{r}").value or "")

        if any(x in ead_val for x in list_ead) and any(x in rwa_val for x in list_rwa):
            start_row = r + 1
            break

    if not start_row:
        continue

    ead_vals, rwa_vals = [], []
    ead_nonzero = None
    rwa_nonzero = None

    # ---- Scan numeric values ----
    for r in range(start_row, last_row + 1):
        ead_v = ws.range(f"{ead_col}{r}").value
        rwa_v = ws.range(f"{rwa_col}{r}").value

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
        target = f"'{sheet_name}'!{ead_col}{start_row}"

        exc_ws.range(f"A{exc_row}").add_hyperlink(
            target,
            text_to_display="Link"
        )
        exc_ws.range(f"B{exc_row}").value = amount
        exc_row += 1

    # ---- Case 2: RWA zero, EAD non-zero ----
    if rwa_all_zero and ead_nonzero:
        r, amount = ead_nonzero
        target = f"'{sheet_name}'!{rwa_col}{start_row}"

        exc_ws.range(f"F{exc_row}").add_hyperlink(
            target,
            text_to_display="Link"
        )
        exc_ws.range(f"G{exc_row}").value = amount
        exc_row += 1

# ================= SAVE & CLOSE =================
wb.save()
wb.close()
app.quit()

print("Hyperlinks successfully created using xlwings.")

