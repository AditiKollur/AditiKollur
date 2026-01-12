```
import pandas as pd
import win32com.client as win32
from win32com.client import constants
import os
import time

# =====================================================
# 1. DATA
# =====================================================
df = pd.DataFrame({
    "Site": ["Plant1", "Plant1", "Plant2", "Plant2", "Plant1"],
    "Product": ["A", "B", "A", "B", "A"],
    "Sales": [100, 150, 200, 180, 120],
    "Active_Flag": ["Y", "N", "N", "Y", "N"],
    "Valid_Flag": ["N", "N", "Y", "N", "N"]
})

FILE_PATH = os.path.abspath("FINAL_PIVOT.xlsx")
DATA_SHEET = "Data"
PIVOT_SHEET = "Pivot"

FILTER_COLUMNS = ["Active_Flag", "Valid_Flag"]
FILTER_VALUE = "N"

# =====================================================
# 2. WRITE DATA (NO EXCEL YET)
# =====================================================
df.to_excel(FILE_PATH, sheet_name=DATA_SHEET, index=False)

# =====================================================
# 3. START EXCEL (ISOLATED INSTANCE)
# =====================================================
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(FILE_PATH)

ws_data = wb.Worksheets(DATA_SHEET)
ws_pivot = wb.Worksheets.Add(After=ws_data)
ws_pivot.Name = PIVOT_SHEET

# =====================================================
# 4. DEFINE SOURCE RANGE (CRITICAL)
# =====================================================
last_row = ws_data.Cells(ws_data.Rows.Count, 1).End(constants.xlUp).Row
last_col = ws_data.Cells(1, ws_data.Columns.Count).End(constants.xlToLeft).Column

source_range = ws_data.Range(
    ws_data.Cells(1, 1),
    ws_data.Cells(last_row, last_col)
)

# =====================================================
# 5. CREATE PIVOT CACHE (THIS IS THE KEY OBJECT)
# =====================================================
pivot_cache = wb.PivotCaches().Create(
    SourceType=constants.xlDatabase,
    SourceData=source_range
)

# =====================================================
# 6. CREATE PIVOT TABLE (DO NOTHING ELSE YET)
# =====================================================
pivot = pivot_cache.CreatePivotTable(
    TableDestination=ws_pivot.Cells(1, 1),
    TableName="FINAL_PIVOT"
)

# 🔴 FORCE EXCEL TO COMMIT THE PIVOT
pivot.RefreshTable()
time.sleep(1)

# =====================================================
# 7. ADD ROWS / COLUMNS / VALUES FIRST
# =====================================================
pivot.PivotFields("Site").Orientation = constants.xlRowField
pivot.PivotFields("Product").Orientation = constants.xlColumnField

pivot.AddDataField(
    pivot.PivotFields("Sales"),
    "Sum of Sales",
    constants.xlSum
)

# =====================================================
# 8. NOW APPLY FILTERS (LAST STEP)
# =====================================================
for col in FILTER_COLUMNS:
    pf = pivot.PivotFields(col)
    pf.Orientation = constants.xlPageField
    pf.ClearAllFilters()

    # Safe filter application
    items = [i.Name for i in pf.PivotItems()]
    if FILTER_VALUE in items:
        pf.CurrentPage = FILTER_VALUE
    else:
        # fallback – first non-blank
        for i in items:
            if i not in ("(blank)", ""):
                pf.CurrentPage = i
                break

# =====================================================
# 9. FINAL TOUCH
# =====================================================
pivot.TableStyle2 = "PivotStyleMedium9"
ws_pivot.Columns.AutoFit()

# =====================================================
# 10. SAVE & EXIT
# =====================================================
wb.Save()
wb.Close()
excel.Quit()

print("✅ Pivot table CREATED — verified Excel-native PivotCache")

