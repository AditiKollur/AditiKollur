```
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.comments import Comment
from copy import copy

# --- SETTINGS ---
TAB_COLORS = ["FFC0CB", "90EE90", "87CEEB", "FFFF99", "C0C0C0", "FFD580"]  # pink, light green, sky blue, yellow, grey, light orange
SKIP_SHEETS = {"Sample"}
SINGLE_INSTANCE_SHEETS = {"ReadMe", "Taxonomy Dropdowns"}
START_ROW = 10  # treat row 10 as header in each sheet

# --- HELPERS ---

def select_files_dialog():
    return filedialog.askopenfilenames(
        title="Select Excel files to combine",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )

def safe_copy_cell(src_cell, tgt_ws, tgt_row, tgt_col):
    """Create target cell at (tgt_row, tgt_col) copying value & style from src_cell."""
    tgt_cell = tgt_ws.cell(row=tgt_row, column=tgt_col, value=src_cell.value)
    try:
        if src_cell.has_style:
            tgt_cell.font = copy(src_cell.font)
            tgt_cell.fill = copy(src_cell.fill)
            tgt_cell.border = copy(src_cell.border)
            tgt_cell.alignment = copy(src_cell.alignment)
            tgt_cell.number_format = src_cell.number_format
            tgt_cell.protection = copy(src_cell.protection)
    except Exception:
        pass
    if src_cell.comment:
        try:
            tgt_cell.comment = Comment(src_cell.comment.text, src_cell.comment.author)
        except Exception:
            pass
    if src_cell.hyperlink:
        try:
            # try to copy hyperlink object
            tgt_cell._hyperlink = copy(src_cell.hyperlink)
        except Exception:
            try:
                tgt_cell.hyperlink = src_cell.hyperlink.target
            except Exception:
                pass
    return tgt_cell

def copy_sheet_contents(src_ws, tgt_ws):
    """
    Copy full sheet contents + formatting.
    Uses enumerate(...) to avoid accessing .col_idx on MergedCell.
    """
    # Column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        try:
            if dim.width is not None:
                tgt_ws.column_dimensions[col_letter].width = dim.width
        except Exception:
            pass

    # Row heights
    for row_idx, dim in src_ws.row_dimensions.items():
        try:
            if dim.height is not None:
                tgt_ws.row_dimensions[row_idx].height = dim.height
        except Exception:
            pass

    # Page setup / margins / print options (best-effort)
    try:
        tgt_ws.page_setup = copy(src_ws.page_setup)
        tgt_ws.print_options = copy(src_ws.print_options)
        tgt_ws.page_margins = copy(src_ws.page_margins)
    except Exception:
        pass

    # Freeze panes
    try:
        if src_ws.freeze_panes:
            tgt_ws.freeze_panes = src_ws.freeze_panes
    except Exception:
        pass

    # Copy all cells using enumerate to compute column index
    for r_idx, row in enumerate(src_ws.iter_rows(), start=1):
        for c_idx, src_cell in enumerate(row, start=1):
            safe_copy_cell(src_cell, tgt_ws, r_idx, c_idx)

    # Copy merged cells AFTER values copied
    try:
        for merged in src_ws.merged_cells.ranges:
            try:
                tgt_ws.merge_cells(str(merged))
            except Exception:
                pass
    except Exception:
        pass

def copy_range_with_formatting(src_ws, tgt_ws, src_start_row, tgt_start_row):
    """
    Copy rows from src_start_row..end from src_ws into tgt_ws starting at tgt_start_row.
    Returns the next free row in target.
    """
    max_row = src_ws.max_row
    max_col = src_ws.max_column
    tgt_row = tgt_start_row

    for r in range(src_start_row, max_row + 1):
        for c in range(1, max_col + 1):
            src_cell = src_ws.cell(row=r, column=c)
            safe_copy_cell(src_cell, tgt_ws, tgt_row, c)
        tgt_row += 1
    return tgt_row

# --- MAIN COMBINE FUNCTION ---

def combine_excels(files, output_path=None):
    if not files:
        raise ValueError("No files provided")

    combined_wb = Workbook()
    # remove default if present
    if combined_wb.active and combined_wb.active.title == "Sheet":
        combined_wb.remove(combined_wb.active)

    used_single_instance = set()
    color_index = 0
    sheets_for_consolidation = []

    # Copy sheets from each file into combined workbook
    for file_path in files:
        if not os.path.isfile(file_path):
            continue
        try:
            wb = load_workbook(file_path, data_only=False)
        except Exception as e:
            print(f"Warning: skipping {file_path} (load error: {e})")
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
            suffix = 1
            while tgt_name in combined_wb.sheetnames:
                tgt_name = f"{sheet_name}_{suffix}"
                suffix += 1

            tgt_ws = combined_wb.create_sheet(title=tgt_name)
            copy_sheet_contents(src_ws, tgt_ws)

            # Set per-file tab color
            try:
                tgt_ws.sheet_properties.tabColor = tab_color
            except Exception:
                pass

            if sheet_name in SINGLE_INSTANCE_SHEETS:
                used_single_instance.add(sheet_name)
            else:
                # save reference to target sheet (in combined_wb) for consolidation
                sheets_for_consolidation.append(tgt_ws)

    # Build Consolidated sheet
    cons_ws = combined_wb.create_sheet(title="Consolidated")
    try:
        cons_ws.sheet_properties.tabColor = "ADD8E6"
    except Exception:
        pass

    tgt_row = 1
    header_written = False

    for ws in sheets_for_consolidation:
        # ensure sheet has at least START_ROW rows
        if ws.max_row < START_ROW:
            continue

        # If header not yet written: copy row START_ROW as header + style
        if not header_written:
            # Copy header row only (START_ROW) -> place at tgt_row
            max_col = ws.max_column
            for c in range(1, max_col + 1):
                src_cell = ws.cell(row=START_ROW, column=c)
                safe_copy_cell(src_cell, cons_ws, tgt_row, c)
            header_written = True
            tgt_row += 1  # move below header

            # Copy data rows (START_ROW+1 .. end)
            tgt_row = copy_range_with_formatting(ws, cons_ws, START_ROW + 1, tgt_row)
        else:
            # Only append data rows (skip header)
            tgt_row = copy_range_with_formatting(ws, cons_ws, START_ROW + 1, tgt_row)

        # blank row between blocks for readability
        tgt_row += 1

    # If no header was written (no eligible sheets), leave Consolidated empty or remove it
    if not header_written:
        # remove the empty consolidated sheet
        try:
            combined_wb.remove(cons_ws)
        except Exception:
            pass

    # Save file (ask if not specified)
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
    messagebox.showinfo("Excel Combiner", "Select Excel files to combine (Ctrl/Cmd+Click to multi-select).")
    files = select_files_dialog()
    if not files:
        messagebox.showinfo("Cancelled", "No files selected.")
        return
    try:
        saved = combine_excels(files)
        if saved:
            messagebox.showinfo("Success", f"Combined workbook saved:\n{saved}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    run_gui()
