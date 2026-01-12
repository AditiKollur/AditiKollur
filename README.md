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
    "Active_Flag": ["Y", "N", "N", "Y", "N"],
    "Year": [2023, 2023, 2023, 2023, 2024]
})

# ================= CONFIG =================
FILE_PATH = os.path.abspath("product_by_site_pivot.xlsx")

DATA_SHEET = "Data"
PIVOT_SHEET = "Pivot"

FILTER_FIELDS = ["Active_Flag"]
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
excel.Calculation = constants.xlCalculationManual

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

# ================= CREATE PIVOT CACHE (THIS IS PivotCaches) =================
pivot_cache = wb.PivotCaches().Create(
    SourceType=constants.xlDatabase,
    SourceData=source_range
)

# ================= CREATE PIVOT TABLE =================
pivot_table = pivot_cache.CreatePivotTable(
    TableDestination=ws_pivot.Cells(3, 1),
    TableName="Product_By_Site"
)

# 🔴 Force Excel to fully initialise the pivot
_ = pivot_table.PivotFields().Count
time.sleep(0.3)

# ================= PAGE FILTERS =================
for field in FILTER_FIELDS:
    pf = pivot_table.PivotFields(field)
    pf.Orientation = constants.xlPageField
    pf.ClearAllFilters()
    pf.CurrentPage = FILTER_VALUE   # SAFE & REQUIRED

# ================= ROW FIELDS =================
for pos, field in enumerate(ROW_FIELDS, start=1):
    pf = pivot_table.PivotFields(field)
    pf.Orientation = constants.xlRowField
    pf.Position = pos

# ================= COLUMN FIELDS =================
for pos, field in enumerate(COLUMN_FIELDS, start=1):
    pf = pivot_table.PivotFields(field)
    pf.Orientation = constants.xlColumnField
    pf.Position = pos

# ================= VALUE FIELD =================
pivot_table.AddDataField(
    pivot_table.PivotFields(VALUE_FIELD),
    f"Sum of {VALUE_FIELD}",
    constants.xlSum
)

# ================= FORMAT =================
pivot_table.RowAxisLayout(constants.xlTabularRow)
pivot_table.TableStyle2 = "PivotStyleMedium9"
ws_pivot.Columns.AutoFit()

# ================= SAVE & CLOSE =================
wb.Save()
wb.Close()

excel.Calculation = constants.xlCalculationAutomatic
excel.Quit()

print("✅ Pivot table created using Excel PivotCaches successfully.")

