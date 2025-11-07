```
"""
Fixed version: handles MergedCell objects (no .col_idx) by using enumerate()
Tkinter GUI -> select multiple Excel files -> combine their sheets into one workbook
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from copy import copy

# Tab colors (hex without '#') - pink, light green, sky blue, yellow, grey, light orange
TAB_COLORS = ["FFC0CB", "90EE90", "87CEEB", "FFFF99", "C0C0C0", "FFD580"]

SKIP_SHEETS = {"Sample"}
SINGLE_INSTANCE_SHEETS = {"ReadMe", "Taxonomy Dropdowns"}

def select_files_dialog():
    files = filedialog.askopenfilenames(
        title="Select Excel files to combine",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )
    return list(files)

def copy_sheet_contents(src_ws, tgt_ws):
    """
    Copy contents and most formatting from src_ws (Worksheet) to tgt_ws (Worksheet).
    - Uses enumerate(...) for column index to avoid MergedCell .col_idx issues.
    - Copies: column widths, row heights, cell values, number formats, fonts, fills, borders,
      alignment, protection, comments, hyperlinks, and merged cell ranges (applied after values).
    - Does NOT copy images/charts.
    """

    # Copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        # skip default/empty entries
        try:
            if dim.width is not None:
                tgt_ws.column_dimensions[col_letter].width = dim.width
        except Exception:
            pass

    # Copy row heights
    for row_idx, row_dim in src_ws.row_dimensions.items():
        try:
            if row_dim.height is not None:
                tgt_ws.row_dimensions[row_idx].height = row_dim.height
        except Exception:
            pass

    # Copy print/page setup (safe best-effort)
    try:
        tgt_ws.page_setup = copy(src_ws.page_setup)
        tgt_ws.print_options = copy(src_ws.print_options)
        tgt_ws.page_margins = copy(src_ws.page_margins)
    except Exception:
        pass

    # Copy freeze panes
    try:
        if src_ws.freeze_panes:
            tgt_ws.freeze_panes = src_ws.freeze_panes
    except Exception:
        pass

    # --- Copy cell values & styles ---
    # Use enumerate to get column index (works with MergedCell placeholders)
    for row_idx, row in enumerate(src_ws.iter_rows(), start=1):
        for col_idx, cell in enumerate(row, start=1):
            # create cell in target
            new_cell = tgt_ws.cell(row=row_idx, column=col_idx, value=cell.value)

            # Copy number format and basic properties
            try:
                new_cell.number_format = cell.number_format
            except Exception:
                pass

            # Copy style components individually (safe copies)
            try:
                if cell.has_style:
                    if cell.font:
                        new_cell.font = copy(cell.font)
                    if cell.fill:
                        new_cell.fill = copy(cell.fill)
                    if cell.border:
                        new_cell.border = copy(cell.border)
                    if cell.alignment:
                        new_cell.alignment = copy(cell.alignment)
                    if cell.protection:
                        new_cell.protection = copy(cell.protection)
            except Exception:
                # ignore any weird style copy errors
                pass

            # Copy comment
            try:
                if cell.comment:
                    new_cell.comment = Comment(cell.comment.text, cell.comment.author)
            except Exception:
                pass

            # Copy hyperlink
            try:
                if cell.hyperlink:
                    # try to copy hyperlink object or fallback to target
                    try:
                        new_cell._hyperlink = copy(cell.hyperlink)
                    except Exception:
                        new_cell.hyperlink = cell.hyperlink.target if hasattr(cell.hyperlink, "target") else cell.hyperlink
            except Exception:
                pass

    # --- Copy merged cells (after values are set) ---
    try:
        for merged_range in src_ws.merged_cells.ranges:
            # merged_range is a MultiCellRange or CellRange; write as string and apply
            try:
                tgt_ws.merge_cells(str(merged_range))
            except Exception:
                pass
    except Exception:
        pass

    # Try to copy sheet view (best-effort)
    try:
        tgt_ws.sheet_view = copy(src_ws.sheet_view)
    except Exception:
        pass

def combine_excels(files, output_path=None):
    if not files:
        raise ValueError("No files provided")

    combined_wb = Workbook()
    # remove default sheet if present
    if combined_wb.active and combined_wb.active.title == "Sheet":
        combined_wb.remove(combined_wb.active)

    used_single_instance = set()
    color_index = 0

    for file_path in files:
        if not os.path.isfile(file_path):
            continue

        try:
            wb = load_workbook(file_path, data_only=False, keep_vba=False)
        except Exception as e:
            print(f"Warning: Skipping file {file_path} due to load error: {e}")
            continue

        tab_color = TAB_COLORS[color_index % len(TAB_COLORS)]
        color_index += 1

        for sheet_name in wb.sheetnames:
            # Skip "Sample" sheets entirely
            if sheet_name in SKIP_SHEETS:
                continue

            # Skip single-instance sheets if already used
            if sheet_name in SINGLE_INSTANCE_SHEETS and sheet_name in used_single_instance:
                continue

            src_ws = wb[sheet_name]

            # Create a unique sheet name in combined workbook
            tgt_name = sheet_name
            counter = 1
            while tgt_name in combined_wb.sheetnames:
                tgt_name = f"{sheet_name}_{counter}"
                counter += 1

            tgt_ws = combined_wb.create_sheet(title=tgt_name)

            # Copy contents and styles
            copy_sheet_contents(src_ws, tgt_ws)

            # Set tab color for the sheet (per-file color)
            try:
                tgt_ws.sheet_properties.tabColor = tab_color
            except Exception:
                pass

            # Track single-instance sheets
            if sheet_name in SINGLE_INSTANCE_SHEETS:
                used_single_instance.add(sheet_name)

    # Ensure at least one sheet exists
    if not combined_wb.sheetnames:
        combined_wb.create_sheet("Combined")

    # Save
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

def run_gui():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("Select files", "Select the Excel files you want to combine (ctrl/cmd+click to multi-select).")
    files = select_files_dialog()
    if not files:
        messagebox.showinfo("Cancelled", "No files selected. Exiting.")
        return

    try:
        saved = combine_excels(files)
        if saved:
            messagebox.showinfo("Success", f"Combined workbook saved:\n{saved}")
        else:
            messagebox.showinfo("Cancelled", "Save cancelled.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    run_gui()
