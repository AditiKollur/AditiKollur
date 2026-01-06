```
import os
import shutil
from datetime import datetime
from openpyxl import load_workbook

# --------------------------------------------------
# FILE PATHS (UPDATE IF NEEDED)
# --------------------------------------------------
desktop = os.path.join(os.path.expanduser("~"), "Desktop")

master_template_path = os.path.join(desktop, "Master_Template.xlsx")
field_mapper_path = os.path.join(desktop, "Field_Mapper.xlsx")
cost_template_path = os.path.join(desktop, "Cost_Template.xlsx")

output_folder = os.path.join(desktop, "Cost Templates")
os.makedirs(output_folder, exist_ok=True)

# --------------------------------------------------
# LOAD FIELD MAPPER
# --------------------------------------------------
fm_wb = load_workbook(field_mapper_path, data_only=True)

booking_sheet = fm_wb["Booking country"]
field_map_sheet = fm_wb["Field fill map"]

# Get Booking Country list (skip header)
booking_countries = [
    row[0].value
    for row in booking_sheet.iter_rows(min_row=2, max_col=1)
    if row[0].value
]

# Load Field Fill Map
field_map = {}
for row in field_map_sheet.iter_rows(min_row=2, values_only=True):
    field_name, tm1_col = row
    if field_name and tm1_col:
        field_map[field_name] = tm1_col

# --------------------------------------------------
# LOAD COST TEMPLATE
# --------------------------------------------------
cost_wb = load_workbook(cost_template_path, data_only=True)
tm1_sheet = cost_wb["TM1 YTD2025"]

# --------------------------------------------------
# DATE FORMAT FOR FILE NAME
# --------------------------------------------------
today = datetime.today()
mmm = today.strftime("%b")
yyyy = today.strftime("%Y")

# --------------------------------------------------
# PROCESS EACH BOOKING COUNTRY
# --------------------------------------------------
for country in booking_countries:

    # Create master template copy
    output_file = f"{country}_Cost_recon TM vs Om - {mmm} YTD {yyyy}.xlsx"
    output_path = os.path.join(output_folder, output_file)

    shutil.copy(master_template_path, output_path)

    master_wb = load_workbook(output_path)
    template_sheet = master_wb["Template"]

    # ----------------------------------------------
    # FIND MATCHING ROW IN COST TEMPLATE (COLUMN D)
    # ----------------------------------------------
    matched_row = None
    for row in tm1_sheet.iter_rows(min_row=2):
        if row[3].value == country:  # Column D
            matched_row = row
            break

    if not matched_row:
        print(f"Booking Country not found in Cost Template: {country}")
        continue

    # ----------------------------------------------
    # 1. FILL Template!C4 FROM COLUMN A
    # ----------------------------------------------
    template_sheet["C4"].value = matched_row[0].value

    # ----------------------------------------------
    # 2. FIELD FILL MAP (ROW 27)
    # ----------------------------------------------
    master_cell_map = {
        "B10": "B10",
        "B13": "B13",
        "B14": "B14",
        "B15": "B15",
        "B16": "B16",
    }

    for master_cell, field_name in master_cell_map.items():
        if field_name in field_map:
            tm1_column_letter = field_map[field_name]

            tm1_col_index = ord(tm1_column_letter.upper()) - 65
            value = tm1_sheet.cell(row=27, column=tm1_col_index + 1).value

            template_sheet[master_cell].value = value

    # ----------------------------------------------
    # SAVE MASTER TEMPLATE
    # ----------------------------------------------
    master_wb.save(output_path)
    print(f"Created: {output_file}")

print("✅ All Cost Templates created successfully.")

