```
"""
Tkinter GUI -> select multiple Excel files -> combine their sheets into one workbook
- Each source file's sheets will be assigned a tab color (colors cycle if more files than colors)
- Does not include any "Sample" sheet from any workbook
- Includes "ReadMe" and "Taxonomy Dropdowns" at most once (first occurrence)
- Tries to preserve formatting (values, fonts, fills, borders, number formats, alignment,
  column widths, row heights, merged cells, comments, hyperlinks, frozen pane)
- Does NOT copy images/charts/embedded objects (openpyxl doesn't reliably support cross-workbook image/chart copying)
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection, NamedStyle
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
    This is a best-effort copy for cells, styles, widths, heights, merged cells, comments, hyperlinks.
    Does NOT copy images/charts.
    """

    # Copy column widths
    for col_dim in src_ws.column_dimensions.items():
        col_letter, dim = col_dim
        tgt_ws.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for row_idx, row_dim in src_ws.row_dimensions.items():
        if row_dim.height is not None:
            tgt_ws.row_dimensions[row_idx].height = row_dim.height

    # Copy freeze/pane
    try:
        if src_ws.sheet_view.selection:
            # frozen panes are stored as src_ws.freeze_panes attribute as well.
            if src_ws.freeze_panes:
                tgt_ws.freeze_panes = src_ws.freeze_panes
    except Exception:
        # ignore if weird sheet_view structure
        pass

    # Copy print/page setup (basic)
    try:
        tgt_ws.page_setup = copy(src_ws.page_setup)
        tgt_ws.print_options = copy(src_ws.print_options)
        tgt_ws.page_margins = copy(src_ws.page_margins)
    except Exception:
        pass

    # Copy merged cells
    for merged in src_ws.merged_cells.ranges:
        tgt_ws.merge_cells(str(merged))

    # Copy cells
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = tgt_ws.cell(row=cell.row, column=cell.col_idx, value=cell.value)

            # Copy number format and basic properties
            new_cell.number_format = cell.number_format
            new_cell.data_type = cell.data_type

            # Copy style components individually (safe)
            if cell.has_style:
                try:
                    if cell.font:
                        new_cell.font = copy(cell.font)
                except Exception:
                    pass
                try:
                    if cell.fill:
                        new_cell.fill = copy(cell.fill)
                except Exception:
                    pass
                try:
                    if cell.border:
                        new_cell.border = copy(cell.border)
                except Exception:
                    pass
                try:
                    if cell.alignment:
                        new_cell.alignment = copy(cell.alignment)
                except Exception:
                    pass
                try:
                    if cell.protection:
                        new_cell.protection = copy(cell.protection)
                except Exception:
                    pass

            # Copy comment
            if cell.comment:
                # cell.comment is an openpyxl.comments.Comment
                try:
                    new_cell.comment = Comment(cell.comment.text, cell.comment.author)
                except Exception:
                    pass

            # Copy hyperlink
            if cell.hyperlink:
                try:
                    new_cell._hyperlink = copy(cell.hyperlink)
                except Exception:
                    # fallback
                    try:
                        new_cell.hyperlink = cell.hyperlink.target
                    except Exception:
                        pass

    # Copy sheet-level properties
    try:
        tgt_ws.sheet_view = copy(src_ws.sheet_view)
    except Exception:
        pass

    # Copy tab color handled outside (on target sheet)

def combine_excels(files, output_path=None):
    if not files:
        raise ValueError("No files provided")

    combined_wb = Workbook()
    # remove default sheet if we will create sheets; keep if needed
    default_sheet = combined_wb.active
    combined_wb.remove(default_sheet)

    used_single_instance = set()
    color_index = 0

    for file_path in files:
        if not os.path.isfile(file_path):
            continue

        try:
            # load workbook with styles preserved; keep formulas (data_only=False)
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

            # Create a new sheet in combined wb with a unique name (avoid name conflicts)
            tgt_name = sheet_name
            counter = 1
            while tgt_name in combined_wb.sheetnames:
                tgt_name = f"{sheet_name}_{counter}"
                counter += 1

            tgt_ws = combined_wb.create_sheet(title=tgt_name)

            # Copy content & styles
            copy_sheet_contents(src_ws, tgt_ws)

            # Set tab color for the sheet (per-file color)
            try:
                tgt_ws.sheet_properties.tabColor = tab_color
            except Exception:
                pass

            # Keep track if single-instance sheet was added
            if sheet_name in SINGLE_INSTANCE_SHEETS:
                used_single_instance.add(sheet_name)

    # If no sheets were added, add a blank sheet to save
    if not combined_wb.sheetnames:
        combined_wb.create_sheet("Combined")

    # Save
    if not output_path:
        # ask user where to save
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
