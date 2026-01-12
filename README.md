```
import pandas as pd
import win32com.client as win32
from win32com.client import constants
import os
import time

# ================= DATA =================
df = pd.DataFrame({
    "Site": ["Plant1", "Plant1", "Plant2", "Plant2", "Plant1"],
    "Product": ["A", "B", "A", "B", "A"],
    "Sales": [100, 150, 200, 180, 120],
    "Active_Flag": ["Y", "N", "N", "Y", None],   # includes blank
    "Valid_Flag": ["N", "N", "Y", "N", "N"],
    "Year": [2023, 2023, 2023, 2023, 2024]
})

# ================= CONFIG =================
FILE_PATH = os.path.abspath("product_by_site_pivot.xlsx")

DATA_SHEET = "Data"
PIVOT_SHEET = "Pivot"

FILTER_COLUMNS = ["Active_Flag", "Valid_Flag"]
FILTER_VALUE = "N"

ROW_FIELDS = ["Site"]
COLUMN_FIELDS = ["Product"]
VALUE_FIELD = "Sales"

# ================= WRITE DATA =================
df.to_excel(FILE_PATH, sheet_name=DATA_SHEET, index=False)

# ================= OPEN EXCEL =================
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(FILE_PATH)
ws_data = wb.Worksheets(DATA_SHEET)
ws_pivot = wb.Worksheets.Add()
ws_pivot.Name = PIVOT_SHEET

# ================= SOURCE RANGE =================
last_row = ws_data.Cells(ws_data.Rows.Count, 1).End(constants.xlUp).Row
last_col = ws_data.Cells(1, ws_data.Columns.Count).End(constants.xlToLeft).Column

source_range = ws_data.Range(
    ws_data.Cells(1, 1),
    ws_data.Cells(last_row, last_col)
)

# ================= CREATE PIVOT CACHE =================
pivot_cache = wb.PivotCaches().Create(
    SourceType=constants.xlDatabase,
    SourceData=source_range
)

# ================= CREATE PIVOT =================
pivot = pivot_cache.CreatePivotTable(
    TableDestination=ws_pivot.Cells(3, 1),
    TableName="Product_By_Site"
)

# 🔴 Force pivot materialisation
pivot.RefreshTable()
time.sleep(0.5)

# ================= APPLY FILTERS (SAFE WAY) =================
for field in FILTER_COLUMNS:
    pf = pivot.PivotFields(field)
    pf.Orientation = constants.xlPageField
    pf.ClearAllFilters()

    # Collect available items
    items = [item.Name for item in pf.PivotItems()]

    # SAFE assignment
    if FILTER_VALUE in items:
        pf.CurrentPage = FILTER_VALUE
    else:
        # Fallback: first non-blank item
        for itm in items:
            if itm not in ("(blank)", ""):
                pf.CurrentPage = itm
                break

# ================= ROWS =================
for i, field in enumerate(ROW_FIELDS, start=1):
    pf = pivot.PivotFields(field)
    pf.Orientation = constants.xlRowField
    pf.Position = i

# ================= COLUMNS =================
for i, field in enumerate(COLUMN_FIELDS, start=1):
    pf = pivot.PivotFields(field)
    pf.Orientation = constants.xlColumnField
    pf.Position = i

# ================= VALUES =================
pivot.AddDataField(
    pivot.PivotFields(VALUE_FIELD),
    f"Sum of {VALUE_FIELD}",
    constants.xlSum
)

# ================= FORMAT =================
pivot.RowAxisLayout(constants.xlTabularRow)
pivot.TableStyle2 = "PivotStyleMedium9"
ws_pivot.Columns.AutoFit()

# ================= SAVE & CLOSE =================
wb.Save()
wb.Close()
excel.Quit()

print("✅ Pivot created safely with validated filters")


