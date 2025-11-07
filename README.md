```
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.comments import Comment
from copy import copy

# Tab colors - pink, light green, sky blue, yellow, grey, light orange
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
    """Copy all contents and formatting from one sheet to another."""
    # Copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            tgt_ws.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for row_idx, row_dim in src_ws.row_dimensions.items():
        if row_dim.height is not None:
            tgt_ws.row_dimensions[row_idx].height = row_dim.height

    # Copy sheet formatting (margins, setup, panes)
    try:
        tgt_ws.page_setup = copy(src_ws.page_setup)
        tgt_ws.print_options = copy(src_ws.print_options)
        tgt_ws.page_margins = copy(src_ws.page_margins)
    except Exception:
        pass

    if src_ws.freeze_panes:
        tgt_ws.freeze_panes = src_ws.freeze_panes

    # Copy cells and styles
    for row_idx, row in enumerate(src_ws.iter_rows(), start=1):
        for col_idx, cell in enumerate(row, start=1):
            new_cell = tgt_ws.cell(row=row_idx, column=col_idx, value=cell.value)
            try:
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.fill = copy(cell.fill)
                    new_cell.border = copy(cell.border)
                    new_cell.alignment = copy(cell.alignment)
                    new_cell.number_format = cell.number_format
            except Exception:
                pass

            if cell.comment:
                try:
                    new_cell.comment = Comment(cell.comment.text, cell.comment.author)
                except Exception:
                    pass

            if cell.hyperlink:
                try:
                    new_cell._hyperlink = copy(cell.hyperlink)
                except Exception:
                    try:
                        new_cell.hyperlink = cell.hyperlink.target
                    except Exception:
                        pass

    # Copy merged cells
    for merged_range in src_ws.merged_cells.ranges:
        try:
            tgt_ws.merge_cells(str(merged_range))
        except Exception:
            pass

def copy_data_with_formatting(src_ws, tgt_ws, src_start_row, tgt_start_row):
    """
    Copy all data (with formatting) from src_start_row to end of src_ws into tgt_ws starting at tgt_start_row.
    Returns the next available target row.
    """
    max_row = src_ws.max_row
    max_col = src_ws.max_column
    tgt_row = tgt_start_row

    for row_idx in range(src_start_row, max_row + 1):
        for col_idx in range(1, max_col + 1):
            src_cell = src_ws.cell(row=row_idx, column=col_idx)
            tgt_cell = tgt_ws.cell(row=tgt_row, column=col_idx, value=src_cell.value)
            try:
                if src_cell.has_style:
                    tgt_cell.font = copy(src_cell.font)
                    tgt_cell.fill = copy(src_cell.fill)
                    tgt_cell.border = copy(src_cell.border)
                    tgt_cell.alignment = copy(src_cell.alignment)
                    tgt_cell.number_format = src_cell.number_format
            except Exception:
                pass
        tgt_row += 1

    return tgt_row

def combine_excels(files, output_path=None):
    """Combine multiple Excel files and add a 'Consolidated' sheet."""
    if not files:
        raise ValueError("No files provided")

    combined_wb = Workbook()
    if combined_wb.active and combined_wb.active.title == "Sheet":
        combined_wb.remove(combined_wb.active)

    used_single_instance = set()
    color_index = 0
    sheets_for_consolidation = []

    # Step 1: Combine all files, color tabs, skip certain sheets
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
            if sheet_name in SKIP_SHEETS:
                continue
            if sheet_name in SINGLE_INSTANCE_SHEETS and sheet_name in used_single_instance:
                continue

            src_ws = wb[sheet_name]
            tgt_name = sheet_name
            counter = 1
            while tgt_name in combined_wb.sheetnames:
                tgt_name = f"{sheet_name}_{counter}"
                counter += 1

            tgt_ws = combined_wb.create_sheet(title=tgt_name)
            copy_sheet_contents(src_ws, tgt_ws)
            tgt_ws.sheet_properties.tabColor = tab_color

            # Track single-instance and consolidation sheets
            if sheet_name in SINGLE_INSTANCE_SHEETS:
                used_single_instance.add(sheet_name)
            else:
                sheets_for_consolidation.append(tgt_ws)

    # Step 2: Create Consolidated sheet
    cons_ws = combined_wb.create_sheet(title="Consolidated")
    cons_ws.sheet_properties.tabColor = "ADD8E6"  # light blue
    tgt_row = 1

    for ws in sheets_for_consolidation:
        # Add data from row 10 onward
        tgt_row = copy_data_with_formatting(ws, cons_ws, src_start_row=10, tgt_start_row=tgt_row)
        tgt_row += 1  # blank line between sheets

    # Step 3: Save final file
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
    messagebox.showinfo("Excel Combiner", "Select the Excel files you want to combine (Ctrl/Cmd+Click for multiple).")
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
