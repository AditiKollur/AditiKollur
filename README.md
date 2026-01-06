```
from openpyxl import load_workbook
from collections import defaultdict
import shutil
import os

# ================= CONFIG =================
MASTER_TEMPLATE_PATH = "Master_Template.xlsx"
FIELD_MAPPER_PATH = "Field_mapper.xlsx"
COST_WALK_PATH = "Cost_walk_summary.xlsx"
OUTPUT_DIR = "output_templates"

# Field mapper sheets
BOOKING_COUNTRY_SHEET = "Booking Country"
FIELD_FILL_MAP_SHEET = "Field fill map"

# Template
TEMPLATE_SHEET = "Template"
C4_CELL = "C4"

# Cost walk
COST_WALK_SHEET = "TM1"
HEADER_ROW = 27
DATA_START_ROW = 28
BOOKING_COUNTRY_COL = "D"
COL_A = "A"
# =========================================

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------- STEP 1: READ BOOKING COUNTRIES ----------
fm_wb = load_workbook(FIELD_MAPPER_PATH, data_only=True)
bc_ws = fm_wb[BOOKING_COUNTRY_SHEET]

bc_headers = {cell.value: idx + 1 for idx, cell in enumerate(bc_ws[1])}
bc_col_idx = bc_headers["Booking Country"]

booking_countries = []
seen = set()
for row in bc_ws.iter_rows(min_row=2, values_only=True):
    if row[bc_col_idx - 1]:
        country = str(row[bc_col_idx - 1]).strip()
        if country not in seen:
            booking_countries.append(country)
            seen.add(country)

# ---------- STEP 2: READ FIELD FILL MAP ----------
ff_ws = fm_wb[FIELD_FILL_MAP_SHEET]
ff_headers = {cell.value: idx + 1 for idx, cell in enumerate(ff_ws[1])}

tm1_col = ff_headers["TM1"]
field_name_col = ff_headers["Field name"]

tm1_to_field = []
for row in ff_ws.iter_rows(min_row=2, values_only=True):
    if row[tm1_col - 1] and row[field_name_col - 1]:
        tm1_to_field.append((
            str(row[tm1_col - 1]).strip(),
            str(row[field_name_col - 1]).strip()
        ))

# ---------- STEP 3: READ COST WALK SUMMARY ----------
cw_wb = load_workbook(COST_WALK_PATH, data_only=True)
cw_ws = cw_wb[COST_WALK_SHEET]

# TM1 headers from row 27
tm1_headers = {
    cw_ws.cell(row=HEADER_ROW, column=col).value: col
    for col in range(1, cw_ws.max_column + 1)
    if cw_ws.cell(row=HEADER_ROW, column=col).value
}

# Booking country → rows
country_rows = {}
for row in range(DATA_START_ROW, cw_ws.max_row + 1):
    country = cw_ws[f"{BOOKING_COUNTRY_COL}{row}"].value
    if country:
        country_rows.setdefault(str(country).strip(), []).append(row)

# ---------- STEP 4: CREATE TEMPLATES ----------
for country in booking_countries:
    output_path = os.path.join(OUTPUT_DIR, f"{country}_template.xlsx")
    shutil.copy(MASTER_TEMPLATE_PATH, output_path)

    tpl_wb = load_workbook(output_path)
    tpl_ws = tpl_wb[TEMPLATE_SHEET]

    rows = country_rows.get(country, [])

    # ---- 4A: Fill C4 with Column A values ----
    col_a_values = []
    for r in rows:
        val = cw_ws[f"{COL_A}{r}"].value
        if val not in [None, ""]:
            col_a_values.append(str(val))
    tpl_ws[C4_CELL] = ", ".join(col_a_values)

    # ---- 4B: Build Template text → ordered cell list (CRITICAL FIX) ----
    template_text_cells = defaultdict(list)

    for row in tpl_ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                template_text_cells[cell.value.strip()].append(cell)

    # ---- 4C: TM1 → Field → Template Mapping (FIRST MATCH ONLY) ----
    for tm1_header, field_name in tm1_to_field:
        if tm1_header not in tm1_headers:
            continue
        if field_name not in template_text_cells:
            continue
        if not rows:
            continue

        col_idx = tm1_headers[tm1_header]

        # First NON-NULL value from Cost Walk
        value = next(
            (
                cw_ws.cell(r, col_idx).value
                for r in rows
                if cw_ws.cell(r, col_idx).value not in [None, ""]
            ),
            None
        )

        if value is None:
            continue

        # FIRST occurrence in Template
        field_cell = template_text_cells[field_name][0]

        target_cell = tpl_ws.cell(
            row=field_cell.row,
            column=field_cell.column + 1
        )

        # Write only once
        if target_cell.value in [None, ""]:
            target_cell.value = value

    tpl_wb.save(output_path)

print("✅ Templates created. First text match populated correctly.")

