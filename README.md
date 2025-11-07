```
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.comments import Comment
from copy import copy

# ---------- CONFIG ----------
TAB_COLORS = ["FFC0CB", "90EE90", "87CEEB", "FFFF99", "C0C0C0", "FFD580"]  # pink, light green, sky blue, yellow, grey, light orange
SKIP_SHEETS = {"Sample"}
SINGLE_INSTANCE_SHEETS = {"ReadMe", "Taxonomy Dropdowns"}
START_ROW = 10  # use row 10 as header; consolidate rows 11..end

# ---------- HELPERS ----------
def select_files_dialog():
    return filedialog.askopenfilenames(
        title="Select Excel files to combine",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm")]
    )

def safe_copy_cell(src_cell, tgt_ws, tgt_row, tgt_col):
    """Copy value + common styles from src_cell into tgt_ws[tgt_row, tgt_col]."""
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
            tgt_cell._hyperlink = copy(src_cell.hyperlink)
        except Exception:
            try:
                tgt_cell.hyperlink = src_cell.hyperlink.target
            except Exception:
                pass
    return tgt_cell

def copy_full_sheet(src_ws, tgt_ws):
    """Copy entire sheet contents and formatting from src_ws to tgt_ws (used for per-sheet copy)."""
    # copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        try:
            if dim.width is not None:
                tgt_ws.column_dimensions[col_letter].width = dim.width
        except Exception:
            pass

    # copy row heights
    for r_idx, dim in src_ws.row_dimensions.items():
        try:
            if getattr(dim, "height", None) is not None:
                tgt_ws.row_dimensions[r_idx].height = dim.height
        except Exception:
            pass

    # copy simple page setup/margins best-effort
    try:
        tgt_ws.page_setup = copy(src_ws.page_setup)
        tgt_ws.print_options = copy(src_ws.print_options)
        tgt_ws.page_margins = copy(src_ws.page_margins)
    except Exception:
        pass

    # freeze panes
    try:
        if src_ws.freeze_panes:
            tgt_ws.freeze_panes = src_ws.freeze_panes
    except Exception:
        pass

    # copy cells (use enumerate to avoid MergedCell.col_idx)
    for r_idx, row in enumerate(src_ws.iter_rows(), start=1):
        for c_idx, src_cell in enumerate(row, start=1):
            safe_copy_cell(src_cell, tgt_ws, r_idx, c_idx)

    # copy merged cell ranges after values
    try:
        for merged in src_ws.merged_cells.ranges:
            try:
                tgt_ws.merge_cells(str(merged))
            except Exception:
                pass
    except Exception:
        pass

def copy_block_to_consolidated(src_ws, cons_ws, src_start_row, cons_start_row):
    """
    Copy rows src_start_row..src_end (src_end = src_ws.max_row) into cons_ws starting at cons_start_row.
    Returns next free row in consolidated (after the appended block).
    Also copies merged ranges inside the copied block (adjusted to consolidated coordinates).
    """
    max_row = src_ws.max_row
    max_col = src_ws.max_column
    tgt_row = cons_start_row

    # Copy row-by-row
    for r in range(src_start_row, max_row + 1):
        for c in range(1, max_col + 1):
            src_cell = src_ws.cell(row=r, column=c)
            safe_copy_cell(src_cell, cons_ws, tgt_row, c)
        tgt_row += 1

    # Copy merged ranges that lie (partially or fully) within src_start_row..max_row
    try:
        for merged in src_ws.merged_cells.ranges:
            # merged is a CellRange; get its bounds
            min_col, min_row, max_col, max_row = merged.bounds  # bounds -> (min_col, min_row, max_col, max_row)
            # if the merged range intersects the copied area
            if max_row >= src_start_row and min_row <= src_ws.max_row:
                # determine overlap portion to map relative to consolidated block
                # Only copy the portion that sits at/after src_start_row.
                # We will map the top row of the source copy area to cons_start_row
                # offset = cons_start_row - src_start_row
                offset = cons_start_row - src_start_row
                new_min_row = min_row + offset
                new_max_row = max_row + offset
                try:
                    cons_ws.merge_cells(start_row=new_min_row, start_column=min_col,
                                        end_row=new_max_row, end_column=max_col)
                except Exception:
                    pass
    except Exception:
        pass

    return tgt_row

# ---------- MAIN FUNCTION ----------
def combine_excels(files, output_path=None):
    if not files:
        raise ValueError("No files provided")

    combined_wb = Workbook()
    # remove default sheet if present
    if combined_wb.active and combined_wb.active.title == "Sheet":
        combined_wb.remove(combined_wb.active)

    used_single_instance = set()
    color_index = 0

    # Create consolidated sheet first (so we can append to it as we go)
    cons_ws = combined_wb.create_sheet(title="Consolidated")
    try:
        cons_ws.sheet_properties.tabColor = "ADD8E6"
    except Exception:
        pass
    cons_next_row = 1
    header_written = False

    for file_path in files:
        if not os.path.isfile(file_path):
            continue
        try:
            wb = load_workbook(file_path, data_only=False)
        except Exception as e:
            print(f"Warning: skipping {file_path} due to load error: {e}")
            continue

        tab_color = TAB_COLORS[color_index % len(TAB_COLORS)]
        color_index += 1

        for sheet_name in wb.sheetnames:
            # Skip always
            if sheet_name in SKIP_SHEETS:
                continue
            # Skip single-instance if already added once
            if sheet_name in SINGLE_INSTANCE_SHEETS and sheet_name in used_single_instance:
                continue

            src_ws = wb[sheet_name]

            # --- 1) copy full sheet into combined workbook (preserve formatting) ---
            tgt_name = sheet_name
            suffix = 1
            while tgt_name in combined_wb.sheetnames:
                tgt_name = f"{sheet_name}_{suffix}"
                suffix += 1
            tgt_ws = combined_wb.create_sheet(title=tgt_name)
            copy_full_sheet(src_ws, tgt_ws)
            try:
                tgt_ws.sheet_properties.tabColor = tab_color
            except Exception:
                pass

            if sheet_name in SINGLE_INSTANCE_SHEETS:
                used_single_instance.add(sheet_name)
                # Do not include these sheets in consolidated
                continue

            # --- 2) append rows START_ROW..end to Consolidated (treat START_ROW as header) ---
            if src_ws.max_row < START_ROW:
                continue  # nothing to copy from this sheet

            # If header not written yet, copy header row (START_ROW) into consolidated once
            if not header_written:
                # copy header row (START_ROW)
                for c in range(1, src_ws.max_column + 1):
                    src_cell = src_ws.cell(row=START_ROW, column=c)
                    safe_copy_cell(src_cell, cons_ws, cons_next_row, c)
                header_written = True
                cons_next_row += 1
                # copy data rows START_ROW+1 .. end
                cons_next_row = copy_block_to_consolidated(src_ws, cons_ws, START_ROW + 1, cons_next_row)
            else:
                # header already written, only append data rows START_ROW+1 .. end
                cons_next_row = copy_block_to_consolidated(src_ws, cons_ws, START_ROW + 1, cons_next_row)

            # add one blank row between appended blocks for readability
            cons_next_row += 1

    # If header never written (no eligible sheets), remove consolidated sheet
    if not header_written:
        try:
            combined_wb.remove(cons_ws)
        except Exception:
            pass

    # Save final workbook
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

# ---------- GUI ----------
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
