````
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
TEMPLATE_COL_HEADER_ROW = 26
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

# ---- Booking Countries ----
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

# ---- TM1 → Template Field mapping (RESTORED) ----
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

# ---- Omnia_col: Template column ↔ PT row filter ----
omnia_col_ws = fm_wb[OMNIA_COL_SHEET]
omnia_col_headers = {c.value: i for i, c in enumerate(omnia_col_ws[1])}

omnia_col_map = []
for r in omnia_col_ws.iter_rows(min_row=2, values_only=True):
    if r[omnia_col_headers["temp_col"]] and r[omnia_col_headers["O_rows"]]:
        omnia_col_map.append({
            "temp_col": str(r[omnia_col_headers["temp_col"]]).strip(),
            "pt_row": str(r[omnia_col_headers["O_rows"]]).strip()
        })

# ---- Omnia_rows: Template row ↔ PT column ----
omnia_rows_ws = fm_wb[OMNIA_ROWS_SHEET]
omnia_rows_headers = {c.value: i for i, c in enumerate(omnia_rows_ws[1])}

omnia_row_map = []
for r in omnia_rows_ws.iter_rows(min_row=2, values_only=True):
    if r[omnia_rows_headers["temp_rows"]] and r[omnia_rows_headers["O_col"]]:
        omnia_row_map.append({
            "temp_row": str(r[omnia_rows_headers["temp_rows"]]).strip(),
            "pt_col": str(r[omnia_rows_headers["O_col"]]).strip()
        })

# ---------- LOAD COST WALK ----------
cw_wb = load_workbook(COST_WALK_PATH, data_only=True)
cw_ws = cw_wb[COST_WALK_SHEET]

# TM1 headers from row 27
tm1_headers = {
    cw_ws.cell(row=HEADER_ROW, column=c).value: c
    for c in range(1, cw_ws.max_column + 1)
    if cw_ws.cell(row=HEADER_ROW, column=c).value
}

# Booking Country → data rows
country_rows = {}
for r in range(DATA_START_ROW, cw_ws.max_row + 1):
    country = cw_ws[f"{BOOKING_COUNTRY_COL}{r}"].value
    if country:
        country_rows.setdefault(str(country).strip(), []).append(r)

# ---------- PT DATAFRAME (ALREADY EXISTS) ----------
# Required columns:
# Booking Country, Business Area, + metric columns
pt_df["Booking Country"] = pt_df["Booking Country"].astype(str).str.strip()
pt_df["Business Area"] = pt_df["Business Area"].astype(str).str.strip()

# ---------- CREATE TEMPLATES ----------
for country in booking_countries:
    output_path = os.path.join(OUTPUT_DIR, f"{country}_template.xlsx")
    shutil.copy(MASTER_TEMPLATE_PATH, output_path)

    tpl_wb = load_workbook(output_path)
    tpl_ws = tpl_wb[TEMPLATE_SHEET]

    rows = country_rows.get(country, [])

    # ======================================================
    # 1️⃣ C4 population from Cost Walk (Column A)
    # ======================================================
    tpl_ws[C4_CELL] = ", ".join(
        str(cw_ws[f"{COL_A}{r}"].value)
        for r in rows
        if cw_ws[f"{COL_A}{r}"].value not in [None, ""]
    )

    # ======================================================
    # 2️⃣ Build Template row / column index (FIRST occurrence)
    # ======================================================
    template_row_index = {}
    for row in tpl_ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                key = cell.value.strip()
                if key not in template_row_index:
                    template_row_index[key] = cell.row

    template_col_index = {}
    for cell in tpl_ws[TEMPLATE_COL_HEADER_ROW]:
        if isinstance(cell.value, str):
            key = cell.value.strip()
            if key not in template_col_index:
                template_col_index[key] = cell.column

    # ======================================================
    # 3️⃣ COST WALK TM1 → TEMPLATE FIELD POPULATION (RESTORED)
    # ======================================================
    for tm1_header, field_name in tm1_to_field:
        if tm1_header not in tm1_headers:
            continue
        if field_name not in template_row_index:
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

        row_idx = template_row_index[field_name]
        target_cell = tpl_ws.cell(row=row_idx, column=template_col_index.get(field_name, 0) + 1)

        if target_cell.value in [None, ""]:
            target_cell.value = value

    # ======================================================
    # 4️⃣ OMNIA MATRIX POPULATION (CORRECT MAPPING)
    # ======================================================
    for col_map in omnia_col_map:
        temp_col = col_map["temp_col"]
        pt_row_value = col_map["pt_row"]

        if temp_col not in template_col_index:
            continue

        col_idx = template_col_index[temp_col]

        pt_slice = pt_df[
            (pt_df["Booking Country"] == country) &
            (pt_df["Business Area"] == pt_row_value)
        ]

        if pt_slice.empty:
            continue

        pt_row = pt_slice.iloc[0]

        for row_map in omnia_row_map:
            temp_row = row_map["temp_row"]
            pt_col = row_map["pt_col"]

            if temp_row not in template_row_index:
                continue
            if pt_col not in pt_row:
                continue

            value = pt_row[pt_col]
            if value in [None, ""]:
                continue

            row_idx = template_row_index[temp_row]
            target_cell = tpl_ws.cell(row=row_idx, column=col_idx)

            if target_cell.value in [None, ""]:
                target_cell.value = value

    tpl_wb.save(output_path)

print("✅ Templates populated with Cost Walk + Omnia mappings successfully.")
