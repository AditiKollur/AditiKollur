```
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import coordinate_from_string, column_index_from_string
from copy import copy
from pathlib import Path

NORMAL_SHEETS = ["gbm", "cmb", "tm1gbm", "tm1cmb"]
ALL_SHEETS = NORMAL_SHEETS + ["siteprod_snapshot"]

def copy_sheet_with_formatting(src_ws, tgt_ws):
    """Copy cells, styles, merged ranges, column widths and row heights."""
    # Copy sheet-level properties (where sensible)
    try:
        tgt_ws.sheet_properties = copy(src_ws.sheet_properties)
    except Exception:
        pass

    # Copy merged cells first
    for merged_range in src_ws.merged_cells.ranges:
        tgt_ws.merge_cells(str(merged_range))

    # Copy cells (value/formula + style + comment + hyperlink)
    for row in src_ws.iter_rows():
        for cell in row:
            # robust column index extraction (works for MergedCell too)
            col_letter, _ = coordinate_from_string(cell.coordinate)
            col_idx = column_index_from_string(col_letter)
            new_cell = tgt_ws.cell(row=cell.row, column=col_idx, value=cell.value)

            # copy styles where available
            if hasattr(cell, "font") and cell.font is not None:
                new_cell.font = copy(cell.font)
            if hasattr(cell, "border") and cell.border is not None:
                new_cell.border = copy(cell.border)
            if hasattr(cell, "fill") and cell.fill is not None:
                new_cell.fill = copy(cell.fill)
            if hasattr(cell, "number_format") and cell.number_format is not None:
                new_cell.number_format = copy(cell.number_format)
            if hasattr(cell, "alignment") and cell.alignment is not None:
                new_cell.alignment = copy(cell.alignment)
            if hasattr(cell, "protection") and cell.protection is not None:
                new_cell.protection = copy(cell.protection)

            # hyperlinks & comments
            if getattr(cell, "hyperlink", None):
                try:
                    new_cell._hyperlink = copy(cell.hyperlink)
                except Exception:
                    pass
            if getattr(cell, "comment", None):
                new_cell.comment = copy(cell.comment)

    # Copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        try:
            tgt_ws.column_dimensions[col_letter].width = dim.width
        except Exception:
            pass

    # Copy row heights
    for idx, dim in src_ws.row_dimensions.items():
        try:
            tgt_ws.row_dimensions[idx].height = dim.height
        except Exception:
            pass

def find_header_row(ws, look_for=("Booking Country", "Product Scope"), max_scan_rows=10):
    """Scan the first `max_scan_rows` rows to find the header row containing the required headers."""
    for r in range(1, max_scan_rows + 1):
        headers = [c.value for c in ws[r]]
        if all(h in headers for h in look_for):
            return r, headers
    return None, None

def filter_rows_keep_only(ws, country):
    """
    Delete rows that do NOT satisfy:
      Booking Country == country AND Product Scope == 'Y'
    Assumes header row can be anywhere in first 10 rows (auto-detected).
    """
    header_row, headers = find_header_row(ws)
    if header_row is None:
        # header not found; nothing to filter
        return

    # determine column indexes (1-based)
    try:
        col_country = headers.index("Booking Country") + 1
        col_scope = headers.index("Product Scope") + 1
    except ValueError:
        return

    # delete rows bottom-up to avoid shifting problems
    for row_idx in range(ws.max_row, header_row, -1):
        val_country = ws.cell(row=row_idx, column=col_country).value
        val_scope = ws.cell(row=row_idx, column=col_scope).value
        # normalize scope to string for robust comparison
        scope_ok = str(val_scope).strip() == "Y" if val_scope is not None else False
        country_ok = (val_country == country)
        if not (country_ok and scope_ok):
            ws.delete_rows(row_idx, 1)

def generate_country_files(input_file, countries, output_dir="output_files"):
    input_file = Path(input_file)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    src_wb = load_workbook(input_file, data_only=False)  # keep formulas

    for country in countries:
        print(f"Processing: {country}")
        tgt_wb = openpyxl.Workbook()
        # remove default sheet
        default = tgt_wb.active
        tgt_wb.remove(default)

        # Create sheets in the same order as ALL_SHEETS
        for sheet_name in ALL_SHEETS:
            if sheet_name not in src_wb.sheetnames:
                print(f"Warning: {sheet_name} not found in source workbook; skipping.")
                continue
            src_ws = src_wb[sheet_name]
            tgt_ws = tgt_wb.create_sheet(title=sheet_name)
            copy_sheet_with_formatting(src_ws, tgt_ws)

            if sheet_name in NORMAL_SHEETS:
                filter_rows_keep_only(tgt_ws, country)

            if sheet_name == "siteprod_snapshot":
                # set B2 to country (preserve formatting)
                tgt_ws["B2"].value = country

        out_path = output_dir / f"{country}.xlsx"
        tgt_wb.save(out_path)
        print(f"Saved: {out_path}")

    print("All done.")

# Example usage
if __name__ == "__main__":
    input_xlsx = "input.xlsx"
    countries_list = ["Hong Kong", "India", "Singapore"]
    generate_country_files(input_xlsx, countries_list, output_dir="country_files")

```
