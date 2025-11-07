```
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from copy import copy
import pandas as pd

# ------------------ CONFIG ------------------
TAB_COLORS = ["FFC0CB", "90EE90", "87CEEB", "FFFF99", "C0C0C0", "FFD580"]  # pink, green, blue, yellow, grey, orange
SINGLE_INSTANCE_SHEETS = {"ReadMe", "Taxonomy Dropdowns"}
SKIP_SHEETS = {"Sample"}
START_ROW = 10  # consolidation starts from this row (10th row as header)

# ------------------ FILE SELECTION ------------------
def select_files():
    return filedialog.askopenfilenames(
        title="Select Excel files to combine",
        filetypes=[("Excel Files", "*.xlsx *.xlsm")]
    )

# ------------------ SHEET COPY (WITH FORMATTING) ------------------
def copy_sheet(source_ws, target_ws):
    for r_idx, row in enumerate(source_ws.iter_rows(), start=1):
        for c_idx, cell in enumerate(row, start=1):
            new_cell = target_ws.cell(row=r_idx, column=c_idx, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    # copy merged cells
    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

# ------------------ MAIN FUNCTION ------------------
def combine_excels(files, output_path=None):
    if not files:
        raise ValueError("No files selected")

    combined_wb = Workbook()
    combined_wb.remove(combined_wb.active)  # remove default empty sheet

    color_index = 0
    added_single_instance = set()
    consolidated_dfs = []

    for file_path in files:
        wb = load_workbook(file_path, data_only=True)
        color = TAB_COLORS[color_index % len(TAB_COLORS)]
        color_index += 1

        for sheet_name in wb.sheetnames:
            if sheet_name in SKIP_SHEETS:
                continue

            # handle single-instance sheets only once
            if sheet_name in SINGLE_INSTANCE_SHEETS:
                if sheet_name in added_single_instance:
                    continue
                added_single_instance.add(sheet_name)

            src_ws = wb[sheet_name]
            new_name = sheet_name
            suffix = 1
            while new_name in combined_wb.sheetnames:
                new_name = f"{sheet_name}_{suffix}"
                suffix += 1

            tgt_ws = combined_wb.create_sheet(title=new_name)
            tgt_ws.sheet_properties.tabColor = color
            copy_sheet(src_ws, tgt_ws)

            # --- For consolidation ---
            if sheet_name not in SINGLE_INSTANCE_SHEETS:
                try:
                    # Read sheet starting row 10 as header
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=START_ROW - 1)

                    # Skip if empty or less than 3 columns
                    if df.shape[1] < 3:
                        continue

                    # Drop last column (skip reading last column)
                    df = df.iloc[:, :-1]

                    # Keep only rows where first 3 columns have values
                    df = df.dropna(subset=df.columns[:3], how="any")

                    # Add metadata columns
                    df["Source_File"] = os.path.basename(file_path)
                    df["Source_Sheet"] = sheet_name

                    consolidated_dfs.append(df)
                except Exception as e:
                    print(f"Skipping {sheet_name} in {os.path.basename(file_path)} due to error: {e}")
                    continue

    # ------------------ CONSOLIDATED SHEET ------------------
    if consolidated_dfs:
        final_df = pd.concat(consolidated_dfs, ignore_index=True)

        # Create consolidated sheet as first sheet
        cons_ws = combined_wb.create_sheet("Consolidated", 0)

        # Write header
        for c_idx, col_name in enumerate(final_df.columns, start=1):
            cons_ws.cell(row=1, column=c_idx, value=col_name)

        # Write data
        for r_idx, row in enumerate(final_df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                cons_ws.cell(row=r_idx, column=c_idx, value=value)

    # ------------------ SAVE ------------------
    if not output_path:
        output_path = filedialog.asksaveasfilename(
            title="Save Combined File",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not output_path:
            return

    combined_wb.save(output_path)
    return output_path

# ------------------ RUN GUI ------------------
def run_gui():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Excel Combiner", "Select Excel files to combine")
    files = select_files()
    if not files:
        messagebox.showinfo("Cancelled", "No files selected")
        return
    try:
        out = combine_excels(files)
        if out:
            messagebox.showinfo("Success", f"Combined workbook saved:\n{out}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    run_gui()
