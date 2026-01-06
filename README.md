```
from openpyxl import load_workbook
from openpyxl import Workbook
import shutil
import os

# ---------------- CONFIG ----------------
MASTER_TEMPLATE_PATH = "Master_Template.xlsx"
FIELD_MAPPER_PATH = "Field_mapper.xlsx"
COST_WALK_PATH = "Cost_walk_summary.xlsx"
OUTPUT_DIR = "output_templates"

FIELD_MAPPER_SHEET = "Booking Country"
BOOKING_COUNTRY_COLUMN_HEADER = "Booking Country"

TEMPLATE_SHEET_NAME = "Template"
TARGET_CELL = "C4"

COST_WALK_START_ROW = 28
COST_WALK_VALUE_COL = "A"
COST_WALK_COUNTRY_COL = "D"
# ----------------------------------------

os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------- Step 1: Read Booking Countries --------
fm_wb = load_workbook(FIELD_MAPPER_PATH, data_only=True)
fm_ws = fm_wb[FIELD_MAPPER_SHEET]

headers = {cell.value: idx + 1 for idx, cell in enumerate(fm_ws[1])}
country_col_idx = headers[BOOKING_COUNTRY_COLUMN_HEADER]

booking_countries = set()
for row in fm_ws.iter_rows(min_row=2, values_only=True):
    if row[country_col_idx - 1]:
        booking_countries.add(str(row[country_col_idx - 1]).strip())

# -------- Step 2: Read Cost Walk Summary --------
cw_wb = load_workbook(COST_WALK_PATH, data_only=True)
cw_ws = cw_wb.active

country_to_values = {}

for row in range(COST_WALK_START_ROW, cw_ws.max_row + 1):
    country = cw_ws[f"{COST_WALK_COUNTRY_COL}{row}"].value
    value = cw_ws[f"{COST_WALK_VALUE_COL}{row}"].value

    if country:
        country = str(country).strip()
        country_to_values.setdefault(country, []).append(str(value))

# -------- Step 3: Create Templates --------
for country in booking_countries:
    output_path = os.path.join(
        OUTPUT_DIR, f"{country}_template.xlsx"
    )

    # Copy master template
    shutil.copy(MASTER_TEMPLATE_PATH, output_path)

    wb = load_workbook(output_path)
    ws = wb[TEMPLATE_SHEET_NAME]

    values = country_to_values.get(country, [])
    ws[TARGET_CELL] = ", ".join(values) if values else ""

    wb.save(output_path)

print("✅ Templates created successfully.")

