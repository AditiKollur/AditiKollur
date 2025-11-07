```
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.comments import Comment
from copy import copy

# --- SETTINGS ---
TAB_COLORS = ["FFC0CB", "90EE90", "87CEEB", "FFFF99", "C0C0C0", "FFD580"]  # pink, green, blue, yellow, grey, orange
SKIP_SHEETS = {"Sample"}
SINGLE_INSTANCE_SHEETS = {"ReadMe", "Taxonomy Dropdowns"}
START_ROW = 10  # data starts here in each sheet

# --- CORE HELPERS ---

def select_files_dialog():
    return filedialog.askopenfilenames(
        title="Select Excel files to combine",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )

def copy_sheet_contents(src_ws, tgt_ws):
    """Copy full sheet contents + formatting (for non-consolidation sheets)."""
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            tgt_ws.column_dimensions[col_letter].width = dim.width
    for row_idx, dim in src_ws.row_dimensions.items():
        if dim.height:
            tgt_ws.row_dimensions[row_idx].height = dim.height

    try:
        tgt_ws.page_setup = copy(src_ws.page_setup)
        tgt_ws.print_options = copy(src_ws.print_options)
        tgt_ws.page_margins = copy(src_ws.page_margins)
    except Exception:
        pass

    if src_ws.freeze_panes:
        tgt_ws.freeze_panes = src_ws.freeze_panes

    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = tgt_ws.cell(row=cell.row, column=cell.col_idx, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
            if cell.comment:
                new_cell.comment = Comment(cell.comment.text, cell.comment.author)
            if cell.hyperlink:
                new_cell.hyperlink = cell.hyperlink.target if cell.hyperlink.target else None

    for merged in src_ws.merged_cells.ranges:
        tgt_ws.merge_cells(str(merged))

def copy_range_with_formatting(src_ws, tgt_ws, src_start_row, tgt_start_row):
    """Copy rows from src_start_row to end with formatting."""
    max_row = src_ws.max_row
    max_col = src_ws.max_column
    tgt_row = tgt_start_row
    for r in range(src_start_row, max_row + 1):
        for c in range(1, max_col + 1):
            src_cell = src_ws.cell(row=r, column=c)
            tgt_cell = tgt_ws.cell(row=tgt_row, column=c, value=src_cell.value)
            if src_cell.has_style:
                tgt_cell.font = copy(src_cell.font)
                tgt_cell.fill = copy(src_cell.fill)
                tgt_cell.border = copy(src_cell.border)
                tgt_cell.alignment = copy(src_cell.alignment)
                tgt_cell.number_format = src_cell.number_format
        tgt_row += 1
    return tgt_row

# --- MAIN FUNCTION ---

def combine_excels(files, output_path=None):
    combined_wb = Workbook()
    if "Sheet" in combined_wb.sheetnames:
        combined_wb.remove(combined_wb["Sheet"])

    used_single_instance = set()
    color_index = 0
    sheets_for_consolidation = []

    # Step 1: Copy all sheets from selected files
    for file_path in files:
        if not os.path.isfile(file_path):
            continue
        try:
            wb = load_workbook(file_path, data_only=False)
        except Exception as e:
            print(f"Error loading {file_path}: {e}")
            continue

        tab_color = TAB_COLORS[color_index % len(TAB_COLORS)]
        color_index += 1

        for sheet_name in wb.sheetnames:
            if sheet_name in SKIP_SHEETS:
                continue
            if sheet_name in SINGLE_INSTANCE_SHEETS and sheet_name in used_single_instance:
                continue

            src_ws = wb[sheet_name]
            tgt_name = sheet_name
            count = 1
            while tgt_name in combined_wb.sheetnames:
                tgt_name = f"{sheet_name}_{count}"
                count += 1
            tgt_ws = combined_wb.create_sheet(tgt_name)
            copy_sheet_contents(src_ws, tgt_ws)
            tgt_ws.sheet_properties.tabColor = tab_color

            if sheet_name in SINGLE_INSTANCE_SHEETS:
                used_single_instance.add(sheet_name)
            else:
                sheets_for_consolidation.append(tgt_ws)

    # Step 2: Create consolidated sheet
    cons_ws = combined_wb.create_sheet("Consolidated")
    cons_ws.sheet_properties.tabColor = "ADD8E6"

    tgt_row = 1
    header_copied = False

    for ws in sheets_for_consolidation:
        if ws.max_row < START_ROW:
            continue  # skip empty or too short sheets

        if not header_copied:
            # Copy header (row 10)
            tgt_row = copy_range_with_formatting(ws, cons_ws, START_ROW, tgt_row)
            header_copied = True
            tgt_row += 1  # leave a blank line
            # Copy data from row 11 onward
            tgt_row = copy_range_with_formatting(ws, cons_ws, START_ROW + 1, tgt_row)
        else:
            # Only copy data (skip header)
            tgt_row = copy_range_with_formatting(ws, cons_ws, START_ROW + 1, tgt_row)
        tgt_row += 1  # blank line between sheets

    # Step 3: Save file
    if not output_path:
        output_path = filedialog.asksaveasfilename(
            title="Save Combined Excel As",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not output_path:
            return None

    combined_wb.save(output_path)
    return output_path

# --- GUI RUNNER ---

def run_gui():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Excel Combiner", "Select Excel files to combine.")
    files = select_files_dialog()
    if not files:
        messagebox.showinfo("Cancelled", "No files selected.")
        return
    try:
        saved = combine_excels(files)
        if saved:
            messagebox.showinfo("Success", f"Combined workbook saved at:\n{saved}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    run_gui()
