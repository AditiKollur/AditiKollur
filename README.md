```
import pandas as pd
import win32com.client as win32
import os

# ================= SAMPLE DATA =================
df = pd.DataFrame({
    "Site": ["Plant1", "Plant1", "Plant2", "Plant2", "Plant1"],
    "Product": ["A", "B", "A", "B", "A"],
    "Sales": [100, 150, 200, 180, 120],
    "Active_Flag": ["Y", "N", "N", "Y", "N"],
    "Year": [2023, 2023, 2023, 2023, 2024]
})

# ================= CONFIG =================
output_file = os.path.abspath("product_by_site_pivot.xlsx")

data_sheet = "Data"
pivot_sheet = "Pivot"

filt_pt = ["Active_Flag"]   # Page filter
rows_pt = ["Site"]
columns_pt = ["Product"]
values_pt = "Sales"

FILTER_VALUE = "N"

# ================= WRITE DATA =================
df.to_excel(output_file, sheet_name=data_sheet, index=False)

# ================= OPEN EXCEL =================
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(output_file)
ws_data = wb.Worksheets(data_sheet)
ws_pivot = wb.Worksheets.Add()
ws_pivot.Name = pivot_sheet

# ================= DATA RANGE =================
last_row = ws_data.Cells(ws_data.Rows.Count, 1).End(-4162).Row   # xlUp
last_col = ws_data.Cells(1, ws_data.Columns.Count).End(-4159).Column  # xlToLeft

source_range = ws_data.Range(
    ws_data.Cells(1, 1),
    ws_data.Cells(last_row, last_col)
)

# ================= CREATE PIVOT CACHE =================
pivot_cache = wb.PivotCaches().Create(
    SourceType=1,        # xlDatabase
    SourceData=source_range
)

# ================= CREATE PIVOT TABLE =================
pivot_table = pivot_cache.CreatePivotTable(
    TableDestination=ws_pivot.Cells(3, 1),
    TableName="Product_By_Site"
)

# ================= FILTERS (SAFE WAY) =================
for field in filt_pt:
    pf = pivot_table.PivotFields(field)
    pf.Orientation = 3      # xlPageField
    pf.ClearAllFilters()
    pf.CurrentPage = FILTER_VALUE   # ✅ ONLY SAFE METHOD

# ================= ROWS =================
for i, field in enumerate(rows_pt, start=1):
    pf = pivot_table.PivotFields(field)
    pf.Orientation = 1      # xlRowField
    pf.Position = i

# ================= COLUMNS =================
for i, field in enumerate(columns_pt, start=1):
    pf = pivot_table.PivotFields(field)
    pf.Orientation = 2      # xlColumnField
    pf.Position = i

# ================= VALUES =================
pivot_table.AddDataField(
    pivot_table.PivotFields(values_pt),
    f"Sum of {values_pt}",
    -4157                  # xlSum
)

# ================= FORMATTING =================
pivot_table.RowAxisLayout(1)   # xlTabular
pivot_table.TableStyle2 = "PivotStyleMedium9"
ws_pivot.Columns.AutoFit()

# ================= SAVE & CLOSE =================
wb.Save()
wb.Close()
excel.Quit()

print("✅ Pivot created successfully with filter set to 'N'")

