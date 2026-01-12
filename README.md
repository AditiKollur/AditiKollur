```
import pandas as pd
import win32com.client as win32
from win32com.client import constants
import shutil
import os

# ================= DATA =================
df = pd.DataFrame({
    "Site": ["Plant1", "Plant1", "Plant2", "Plant2", "Plant1"],
    "Product": ["A", "B", "A", "B", "A"],
    "Sales": [100, 150, 200, 180, 120],
    "Active_Flag": ["Y", "N", "N", "Y", "N"],
    "Valid_Flag": ["N", "N", "Y", "N", "N"]
})

# ================= PATHS =================
TEMPLATE_FILE = os.path.abspath("pivot_template.xlsx")
OUTPUT_FILE = os.path.abspath("final_output.xlsx")

DATA_SHEET = "Data"

# ================= COPY TEMPLATE =================
shutil.copy(TEMPLATE_FILE, OUTPUT_FILE)

# ================= OPEN EXCEL =================
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(OUTPUT_FILE)
ws_data = wb.Worksheets(DATA_SHEET)

# ================= CLEAR OLD DATA =================
ws_data.Cells.Clear()

# ================= WRITE NEW DATA =================
for col_idx, col_name in enumerate(df.columns, start=1):
    ws_data.Cells(1, col_idx).Value = col_name

for row_idx, row in enumerate(df.itertuples(index=False), start=2):
    for col_idx, value in enumerate(row, start=1):
        ws_data.Cells(row_idx, col_idx).Value = value

# ================= REFRESH ALL PIVOTS =================
wb.RefreshAll()

# ================= SAVE & CLOSE =================
wb.Save()
wb.Close()
excel.Quit()

print("✅ Data updated and pivot refreshed successfully")

