```
from openpyxl import load_workbook
from collections import defaultdict
import pandas as pd
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
OMNIA_ROWS_SHEET = "Omnia_rows"
OMNIA_COL_SHEET = "Omnia_col"

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

# ---------- LOAD FIELD MAPPER ----------
fm_wb = load_workbook(FIELD_MAPPER_PATH, data_only=True)

# Booking Countries
bc_ws = fm_wb[BOOKING_COUNTRY_SHEET]
bc_headers = {c.value: i + 1 for i, c in enumerate(bc_ws[1])}
bc_col_idx = bc_headers["Booking Country"]

booking_countries = []
seen = set()
for r in bc_ws.iter_rows(min_row=2, values_only=True):
    if r[bc_col_idx - 1]:
        c = str(r[bc_col_idx - 1]).strip()
        if c not in seen:
            booking_countries.append(c)
            seen.add(c)

# TM1 → Field mapping
ff_ws = fm_wb[FIELD_FILL_MAP_SHEET]
ff_headers = {c.value: i + 1 for i, c in enumerate(ff_ws[1])}

tm1_col = ff_headers["TM1"]
field_col = ff_headers["Field name"]

tm1_to_field = []
for r in ff_ws.iter_rows(min_row=2, values_only=True):
    if r[tm1_col - 1] and r[field_col - 1]:
        tm1_to_field.append((
            str(r[tm1_col - 1]).strip(),
            str(r[field_col - 1]).strip()
        ))

# Omnia metrics (PT COLUMNS)
omnia_metrics = [
    str(r[0]).strip()
    for r in fm_wb[OMNIA_ROWS_SHEET].iter_rows(min_row=2, values_only=True)
    if r[0]
]

# Omnia business areas (PT ROW FILTER)
omnia_business_areas = [
    str(r[0]).strip()
    for r in fm_wb[OMNIA_COL_SHEET].iter_rows(min_row=2, values_only=True)
    if r[0]
]

# ---------- LOAD COST WALK ----------
cw_wb = load_workbook(COST_WALK_PATH, data_only=True)
cw_ws = cw_wb[COST_WALK_SHEET]

tm1_headers = {
    cw_ws.cell(row=HEADER_ROW, column=c).value: c
    for c in range(1, cw_ws.max_column + 1)
    if cw_ws.cell(row=HEADER_ROW, column=c).value
}

country_rows = {}
for r in range(DATA_START_ROW, cw_ws.max_row + 1):
    country = cw_ws[f"{BOOKING_COUNTRY_COL}{r}"].value
    if country:
        country_rows.setdefault(str(country).strip(), []).append(r)

# ---------- PT DATAFRAME (ALREADY IN MEMORY) ----------
# pt_df must already exist
# Required columns:
# Booking Country, Business Area, + omnia_metrics

pt_df["Booking Country"] = pt_df["Booking Country"].astype(str).str.strip()
pt_df["Business Area"] = pt_df["Business Area"].astype(str).str.strip()

# ---------- CREATE TEMPLATES ----------
for country in booking_countries:
    output_path = os.path.join(OUTPUT_DIR, f"{country}_template.xlsx")
    shutil.copy(MASTER_TEMPLATE_PATH, output_path)

    tpl_wb = load_workbook(output_path)
    tpl_ws = tpl_wb[TEMPLATE_SHEET]

    rows = country_rows.get(country, [])

    # ---- Fill C4 ----
    tpl_ws[C4_CELL] = ", ".join(
        str(cw_ws[f"{COL_A}{r}"].value)
        for r in rows
        if cw_ws[f"{COL_A}{r}"].value not in [None, ""]
    )

    # ---- Build FIRST OCCURRENCE text map ----
    template_text_cells = defaultdict(list)
    for row in tpl_ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                template_text_cells[cell.value.strip()].append(cell)

    # ---- TM1 → Field fill ----
    for tm1_header, field_name in tm1_to_field:
        if tm1_header not in tm1_headers or field_name not in template_text_cells:
            continue
        if not rows:
            continue

        col_idx = tm1_headers[tm1_header]
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

        anchor = template_text_cells[field_name][0]
        target = tpl_ws.cell(anchor.row, anchor.column + 1)
        if target.value in [None, ""]:
            target.value = value

    # ---- OMNIA MATRIX FILL (CORRECTED) ----
    for business_area in omnia_business_areas:

        # Column anchor in template
        if business_area not in template_text_cells:
            continue
        col_anchor = template_text_cells[business_area][0]
        col_idx = col_anchor.column

        # Filter PT for country + business area
        pt_slice = pt_df[
            (pt_df["Booking Country"] == country) &
            (pt_df["Business Area"] == business_area)
        ]

        if pt_slice.empty:
            continue

        pt_row = pt_slice.iloc[0]

        for metric in omnia_metrics:

            # Metric must be PT column
            if metric not in pt_row:
                continue
            if metric not in template_text_cells:
                continue

            row_anchor = template_text_cells[metric][0]
            value = pt_row[metric]

            if value in [None, ""]:
                continue

            target = tpl_ws.cell(row_anchor.row, col_idx)
            if target.value in [None, ""]:
                target.value = value

    tpl_wb.save(output_path)

print("✅ Templates populated correctly using Omnia row/column logic.")
