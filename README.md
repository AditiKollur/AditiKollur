```
import pandas as pd

# ================= SAMPLE DATA =================
df = pd.DataFrame({
    "Site": ["Plant1", "Plant1", "Plant2", "Plant2", "Plant1"],
    "Product": ["A", "B", "A", "B", "A"],
    "Sales": [100, 150, 200, 180, 120],
    "Year": [2023, 2023, 2023, 2023, 2024]
})

# ================= CONFIG =================
output_file = "product_by_site_pivot.xlsx"
data_sheet = "Data"
pivot_sheet = "Pivot"

filt_pt = ["Year"]
rows_pt = ["Site"]
columns_pt = ["Product"]
values_pt = "Sales"

# ================= WRITE WITH XLSXWRITER =================
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name=data_sheet, index=False)

    workbook = writer.book
    worksheet = writer.sheets[data_sheet]

    pivot_ws = workbook.add_worksheet(pivot_sheet)

    # Define source range
    max_row, max_col = df.shape
    source_range = f"{data_sheet}!A1:{chr(65+max_col-1)}{max_row+1}"

    # Create Pivot Table
    pivot_ws.add_pivot_table({
        "data": source_range,
        "rows": rows_pt,
        "columns": columns_pt,
        "filters": filt_pt,
        "values": [
            {
                "field": values_pt,
                "function": "sum",
                "name": f"Sum of {values_pt}"
            }
        ],
        "row_headers": True,
        "column_headers": True
    })

    pivot_ws.set_column("A:Z", 15)

print("Excel-native modifiable Pivot Table created successfully.")

