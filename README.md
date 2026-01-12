```
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table
from openpyxl.pivot.table import PivotTable, PivotField
from openpyxl.pivot.cache import CacheDefinition, CacheSource, WorksheetSource

# ================= SAMPLE DATA =================
df = pd.DataFrame({
    "Region": ["Asia", "Asia", "Europe", "Europe", "Asia"],
    "Country": ["India", "China", "France", "Germany", "India"],
    "Product": ["A", "A", "B", "A", "B"],
    "Sales": [100, 150, 200, 180, 120],
    "Year": [2023, 2023, 2023, 2023, 2024]
})

# ================= CONFIG =================
output_file = "excel_native_pivot.xlsx"

data_sheet = "Data"
pivot_sheet = "Pivot"

filt_pt = ["Year"]
rows_pt = ["Region", "Country"]
columns_pt = ["Product"]
values_pt = "Sales"

# ================= WRITE DATA =================
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name=data_sheet, index=False)

# ================= LOAD WORKBOOK =================
wb = load_workbook(output_file)
ws_data = wb[data_sheet]
ws_pivot = wb.create_sheet(pivot_sheet)

max_row = ws_data.max_row
max_col = ws_data.max_column
data_ref = f"{data_sheet}!A1:{chr(64+max_col)}{max_row}"

# ================= CREATE CACHE =================
cache_source = CacheSource(
    type="worksheet",
    worksheetSource=WorksheetSource(
        sheet=data_sheet,
        ref=f"A1:{chr(64+max_col)}{max_row}"
    )
)

cache_def = CacheDefinition(cacheSource=cache_source)
cache = wb._add_pivot_cache(cache_def)

# ================= CREATE PIVOT TABLE =================
pivot = PivotTable(
    cache=cache,
    ref="A3",
    name="DynamicPivot"
)

headers = [cell.value for cell in ws_data[1]]

# Filters
for f in filt_pt:
    idx = headers.index(f)
    pf = PivotField(index=idx)
    pivot.pageFields.append(pf)

# Rows
for r in rows_pt:
    idx = headers.index(r)
    pf = PivotField(index=idx)
    pivot.rowFields.append(pf)

# Columns
for c in columns_pt:
    idx = headers.index(c)
    pf = PivotField(index=idx)
    pivot.colFields.append(pf)

# Values
val_idx = headers.index(values_pt)
pivot.dataFields.append(
    PivotField(index=val_idx, name=f"Sum of {values_pt}")
)

ws_pivot.add_pivot(pivot)

# ================= SAVE =================
wb.save(output_file)

print("Excel-native modifiable Pivot Table created successfully.")

