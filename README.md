```
import pandas as pd
import xlwings as xw
import os

# ======================================================
# DATA
# ======================================================
df = pd.DataFrame({
    "Site": ["Plant1", "Plant1", "Plant2", "Plant2", "Plant1"],
    "Product": ["A", "B", "A", "B", "A"],
    "Sales": [100, 150, 200, 180, 120],
    "Active_Flag": ["Y", "N", "N", "Y", "N"],
    "Valid_Flag": ["N", "N", "Y", "N", "N"]
})

FILE_PATH = os.path.abspath("xlwings_pivot.xlsx")
DATA_SHEET = "Data"
PIVOT_SHEET = "Pivot"

FILTER_COLUMNS = ["Active_Flag", "Valid_Flag"]
FILTER_VALUE = "N"

# ======================================================
# WRITE DATA
# ======================================================
with pd.ExcelWriter(FILE_PATH, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name=DATA_SHEET, index=False)

# ======================================================
# EXCEL VIA XLWINGS
# ======================================================
app = xw.App(visible=False)
wb = app.books.open(FILE_PATH)

ws_data = wb.sheets[DATA_SHEET]
ws_pivot = wb.sheets.add(PIVOT_SHEET)

# ======================================================
# SOURCE RANGE
# ======================================================
last_row = ws_data.range("A" + str(ws_data.cells.last_cell.row)).end("up").row
last_col = ws_data.range("A1").end("right").column

source_range = ws_data.range((1, 1), (last_row, last_col))

# ======================================================
# CREATE PIVOT CACHE
# ======================================================
pivot_cache = wb.api.PivotCaches().Create(
    SourceType=1,      # xlDatabase
    SourceData=source_range.api
)

# ======================================================
# CREATE PIVOT TABLE (THIS WORKS)
# ======================================================
pivot = pivot_cache.CreatePivotTable(
    TableDestination=ws_pivot.range("A3").api,
    TableName="Product_By_Site"
)

# ======================================================
# ROWS / COLUMNS / VALUES
# ======================================================
pivot.PivotFields("Site").Orientation = 1     # xlRowField
pivot.PivotFields("Product").Orientation = 2  # xlColumnField

pivot.AddDataField(
    pivot.PivotFields("Sales"),
    "Sum of Sales",
    -4157                                      # xlSum
)

# ======================================================
# MULTIPLE FILTERS = "N"
# ======================================================
for col in FILTER_COLUMNS:
    pf = pivot.PivotFields(col)
    pf.Orientation = 3     # xlPageField
    pf.ClearAllFilters()

    items = [i.Name for i in pf.PivotItems()]
    if FILTER_VALUE in items:
        pf.CurrentPage = FILTER_VALUE

# ======================================================
# FINAL TOUCH
# ======================================================
pivot.TableStyle2 = "PivotStyleMedium9"
ws_pivot.autofit()

wb.save()
wb.close()
app.quit()

print("✅ Pivot table created successfully using xlwings")

