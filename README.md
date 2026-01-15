```
import os
import pandas as pd
from openpyxl import load_workbook

# ================= CONFIG =================
FOLDER_PATH = "input_files"      # folder containing excel files
OUTPUT_FILE = "final_output.xlsx"

START_ROW_DEF = 27
HEADER_ROW_DEF = 26

# ================= STORAGE =================
bc_data = {}
def_data = {}

common_b_col = None
common_b_col_def = None

# ================= PROCESS FILES =================
for file in os.listdir(FOLDER_PATH):
    if not file.endswith(".xlsx"):
        continue

    file_path = os.path.join(FOLDER_PATH, file)
    file_key = file.split("_")[0]   # first part of filename

    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    # ---------- B & C extraction ----------
    b_vals = []
    c_vals = []

    for row in range(1, ws.max_row + 1):
        b = ws[f"B{row}"].value
        c = ws[f"C{row}"].value
        if b is not None:
            b_vals.append(b)
            c_vals.append(c)

    if common_b_col is None:
        common_b_col = b_vals

    bc_data[file_key] = c_vals

    # ---------- D/E/F extraction ----------
    headers = [
        ws[f"D{HEADER_ROW_DEF}"].value,
        ws[f"E{HEADER_ROW_DEF}"].value,
        ws[f"F{HEADER_ROW_DEF}"].value,
    ]

    b_def = []
    d_vals, e_vals, f_vals = [], [], []

    for row in range(START_ROW_DEF, ws.max_row + 1):
        b = ws[f"B{row}"].value
        if b is None:
            continue

        b_def.append(b)
        d_vals.append(ws[f"D{row}"].value)
        e_vals.append(ws[f"E{row}"].value)
        f_vals.append(ws[f"F{row}"].value)

    if common_b_col_def is None:
        common_b_col_def = b_def

    def_data[f"{file_key}_{headers[0]}"] = d_vals
    def_data[f"{file_key}_{headers[1]}"] = e_vals
    def_data[f"{file_key}_{headers[2]}"] = f_vals

# ================= BUILD DATAFRAMES =================
df_bc = pd.DataFrame(bc_data)
df_bc.insert(0, "Label", common_b_col)

df_def = pd.DataFrame(def_data)
df_def.insert(0, "Label", common_b_col_def)

# ================= WRITE OUTPUT =================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_bc.to_excel(writer, sheet_name="BC_Extract", index=False)
    df_def.to_excel(writer, sheet_name="DEF_Extract", index=False)

print("Extraction completed successfully!")
