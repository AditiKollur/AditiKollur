```
import xlwings as xw
import pandas as pd

# ================= CONFIG =================
INPUT_FILE = r"C:\path\to\protected.xlsx"
OUTPUT_FILE = r"C:\path\to\protected_UPDATED.xlsx"
PASSWORD = "your_password"

EXCEPTION_SHEET = "EAD_RWA exception"

list_ead = ["EAD", "Exposure"]
list_rwa = ["RWA", "Risk"]

# list_rep dataframe columns:
# Sheet_Name, Check, EAD, RWA, Row_to_scan_values

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
    "Exception_Type",
    "Link",
    "Amount",
    "",
    "",
    ""
]

exc_row = 2  # append pointer

# ================= MAIN LOOP =================
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
    row_scan_start = rep.get("Row_to_scan_values")

    last_row = ws.used_range.last_cell.row

    # ---- BULK READ ----
    ead_vals = ws.range(f"{ead_col}1:{ead_col}{last_row}").value
    rwa_vals = ws.range(f"{rwa_col}1:{rwa_col}{last_row}").value

    if not ead_vals or not rwa_vals:
        continue

    # ============================================================
    # 1️⃣ PRIMARY HEADER SEARCH
    # ============================================================
    start_row = None

    for i, (e, r) in enumerate(zip(ead_vals, rwa_vals)):
        if (
            isinstance(e, str) and any(x in e for x in list_ead)
            and isinstance(r, str) and any(x in r for x in list_rwa)
        ):
            start_row = i + 1   # 🔥 FIX: start from same row
            break

    # ============================================================
    # 2️⃣ FALLBACK SEARCH (MERGED HEADERS)
    # ============================================================
    if start_row is None and pd.notna(row_scan_start):
        row_scan_start = int(row_scan_start)

        for i in range(row_scan_start - 1, len(ead_vals)):
            e = ead_vals[i]
            r = rwa_vals[i]

            if (
                isinstance(e, str) and any(x in e for x in list_ead)
                and isinstance(r, str) and any(x in r for x in list_rwa)
            ):
                start_row = i + 1   # 🔥 FIX
                break

    if start_row is None:
        continue

    # ============================================================
    # 3️⃣ NUMERIC SCAN (UNCHANGED)
    # ============================================================
    for idx in range(start_row - 1, len(ead_vals)):
        e = ead_vals[idx]
        r = rwa_vals[idx]

        if not isinstance(e, (int, float)) or not isinstance(r, (int, float)):
            continue

        excel_row = idx + 1

        # CASE 1: EAD zero, RWA non-zero
        if e == 0 and r != 0:
            link = f"#'{sheet_name}'!{ead_col}{excel_row}"
            text = f"{sheet_name}_{ead_col}{excel_row}"

            exc_ws.range(f"A{exc_row}").value = "EAD zero / RWA non-zero"
            exc_ws.range(f"B{exc_row}").add_hyperlink(link, text_to_display=text)
            exc_ws.range(f"C{exc_row}").value = r
            exc_row += 1

        # CASE 2: RWA zero, EAD non-zero
        elif r == 0 and e != 0:
            link = f"#'{sheet_name}'!{rwa_col}{excel_row}"
            text = f"{sheet_name}_{rwa_col}{excel_row}"

            exc_ws.range(f"A{exc_row}").value = "RWA zero / EAD non-zero"
            exc_ws.range(f"B{exc_row}").add_hyperlink(link, text_to_display=text)
            exc_ws.range(f"C{exc_row}").value = e
            exc_row += 1

# ================= SAVE =================
wb.api.SaveAs(OUTPUT_FILE)
wb.close()

app.calculation = 'automatic'
app.quit()

print("Numeric scan now starts from the specified row itself.")
