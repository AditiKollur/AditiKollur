```
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from copy import copy
from pathlib import Path

NORMAL_SHEETS = ["gbm", "cmb", "tm1gbm", "tm1cmb"]
ALL_SHEETS = NORMAL_SHEETS + ["siteprod_snapshot"]

def copy_styles_only(src_cell, tgt_cell):
    """Copy common style attributes (but not formula/value decision)."""
    if getattr(src_cell, "font", None) is not None:
        tgt_cell.font = copy(src_cell.font)
    if getattr(src_cell, "border", None) is not None:
        tgt_cell.border = copy(src_cell.border)
    if getattr(src_cell, "fill", None) is not None:
        tgt_cell.fill = copy(src_cell.fill)
    if getattr(src_cell, "number_format", None) is not None:
        tgt_cell.number_format = copy(src_cell.number_format)
    if getattr(src_cell, "alignment", None) is not None:
        tgt_cell.alignment = copy(src_cell.alignment)
    if getattr(src_cell, "protection", None) is not None:
        tgt_cell.protection = copy(src_cell.protection)
    if getattr(src_cell, "comment", None) is not None:
        try:
            tgt_cell.comment = copy(src_cell.comment)
        except Exception:
            pass

def copy_sheet_with_formatting(src_ws, tgt_ws):
    """Full copy: values (or formulas) + formatting + merged ranges + dims."""
    # copy merged ranges
    for merged_range in src_ws.merged_cells.ranges:
        try:
            tgt_ws.merge_cells(str(merged_range))
        except Exception:
            pass

    for row in src_ws.iter_rows():
        for cell in row:
            col_letter, _ = coordinate_from_string(cell.coordinate)
            col_idx = column_index_from_string(col_letter)
            tgt = tgt_ws.cell(row=cell.row, column=col_idx, value=cell.value)
            copy_styles_only(cell, tgt)

    # column widths & row heights
    for col_letter, dim in src_ws.column_dimensions.items():
        try:
            tgt_ws.column_dimensions[col_letter].width = dim.width
        except Exception:
            pass
    for idx, dim in src_ws.row_dimensions.items():
        try:
            tgt_ws.row_dimensions[idx].height = dim.height
        except Exception:
            pass

def find_header_row(ws, look_for=("Booking Country", "Product Scope"), max_scan_rows=20):
    """Return (header_row_index, headers_list) or (None, None)."""
    maxr = min(max_scan_rows, ws.max_row)
    for r in range(1, maxr + 1):
        headers = [c.value for c in ws[r]]
        if headers and all(h in headers for h in look_for):
            return r, headers
    return None, None

def copy_filtered_data(src_ws_vals, src_ws_fmt, tgt_ws, country):
    """
    Copy rows that should be kept:
      - Keep header rows (found automatically)
      - Keep data rows where Booking Country == country AND Product Scope == 'Y'
    Values come from src_ws_vals (data_only=True workbook)
    Formatting comes from src_ws_fmt (data_only=False workbook)
    """
    # merged ranges
    for merged_range in src_ws_fmt.merged_cells.ranges:
        try:
            tgt_ws.merge_cells(str(merged_range))
        except Exception:
            pass

    header_row, headers = find_header_row(src_ws_fmt)
    if header_row is None:
        # no header found -> copy whole sheet values+formatting
        # fallback behavior: copy everything as values + formatting
        tgt_row = 1
        for r in range(1, src_ws_fmt.max_row + 1):
            for cell_fmt in src_ws_fmt[r]:
                col_letter, _ = coordinate_from_string(cell_fmt.coordinate)
                col_idx = column_index_from_string(col_letter)
                # get value from data-only sheet if possible
                try:
                    value = src_ws_vals.cell(row=r, column=col_idx).value
                except Exception:
                    value = cell_fmt.value
                tgt = tgt_ws.cell(row=tgt_row, column=col_idx, value=value)
                copy_styles_only(cell_fmt, tgt)
            tgt_row += 1
    else:
        col_country = headers.index("Booking Country") + 1
        col_scope = headers.index("Product Scope") + 1

        tgt_row = 1
        for r in range(1, src_ws_fmt.max_row + 1):
            # determine whether to keep this row
            is_header = (r <= header_row)
            val_country = src_ws_vals.cell(row=r, column=col_country).value
            val_scope = src_ws_vals.cell(row=r, column=col_scope).value
            scope_ok = (str(val_scope).strip() == "Y") if val_scope is not None else False
            country_ok = (val_country == country)
            keep = is_header or (country_ok and scope_ok)
            if keep:
                for cell_fmt in src_ws_fmt[r]:
                    col_letter, _ = coordinate_from_string(cell_fmt.coordinate)
                    col_idx = column_index_from_string(col_letter)
                    # value from evaluated workbook (data_only)
                    try:
                        value = src_ws_vals.cell(row=r, column=col_idx).value
                    except Exception:
                        value = cell_fmt.value
                    tgt = tgt_ws.cell(row=tgt_row, column=col_idx, value=value)
                    copy_styles_only(cell_fmt, tgt)
                tgt_row += 1

    # preserve dims from formatting sheet
    for col_letter, dim in src_ws_fmt.column_dimensions.items():
        try:
            tgt_ws.column_dimensions[col_letter].width = dim.width
        except Exception:
            pass
    for idx, dim in src_ws_fmt.row_dimensions.items():
        try:
            tgt_ws.row_dimensions[idx].height = dim.height
        except Exception:
            pass

def generate_country_files(input_file, countries, output_dir="output_files"):
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input not found: {input_file}")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # load two workbooks: one with formulas, one with evaluated values
    src_wb_formula = load_workbook(input_path, data_only=False)
    src_wb_values = load_workbook(input_path, data_only=True)

    for country in countries:
        print(f"Processing -> {country}")
        tgt_wb = openpyxl.Workbook()
        default_ws = tgt_wb.active
        created = 0

        for sheet_name in ALL_SHEETS:
            if sheet_name not in src_wb_formula.sheetnames:
                print(f"  Warning: source missing sheet '{sheet_name}', skipping.")
                continue

            src_ws_fmt = src_wb_formula[sheet_name]
            src_ws_vals = src_wb_values[sheet_name]  # exists because sheet exists in formula WB

            # reuse the default sheet for the first created sheet to avoid empty-wb issues
            if created == 0:
                tgt_ws = default_ws
                tgt_ws.title = sheet_name
            else:
                tgt_ws = tgt_wb.create_sheet(title=sheet_name)

            if sheet_name in NORMAL_SHEETS:
                copy_filtered_data(src_ws_vals, src_ws_fmt, tgt_ws, country)
            else:  # siteprod_snapshot -> full copy (formulas preserved)
                copy_sheet_with_formatting(src_ws_fmt, tgt_ws)
                # update B2 into the target (preserve style)
                tgt_ws["B2"].value = country

            # ensure this sheet is visible
            try:
                tgt_ws.sheet_state = 'visible'
            except Exception:
                pass

            created += 1

        # If nothing got created (e.g., all expected sheets missing), fallback: copy first source sheet
        if created == 0:
            fallback_src = src_wb_formula[src_wb_formula.sheetnames[0]]
            tgt_ws = default_ws
            tgt_ws.title = fallback_src.title
            copy_sheet_with_formatting(fallback_src, tgt_ws)
            print(f"  No expected sheets found; copied fallback sheet '{fallback_src.title}' into output.")
            # attempt to set B2 if present
            try:
                tgt_ws["B2"].value = country
            except Exception:
                pass

        # final safety: ensure at least one visible sheet
        visible_exists = any(getattr(s, "sheet_state", "visible") == "visible" for s in tgt_wb.worksheets)
        if not visible_exists:
            # force first worksheet visible
            try:
                tgt_wb.worksheets[0].sheet_state = 'visible'
            except Exception:
                pass

        out_path = output_dir / f"{country}.xlsx"
        try:
            tgt_wb.save(out_path)
            print(f"  Saved: {out_path}")
        except Exception as e:
            print(f"  ERROR saving {out_path}: {e}")
            # last-resort fallback: try saving to a different name
            fallback_path = output_dir / f"{country}_fallback.xlsx"
            tgt_wb.save(fallback_path)
            print(f"  Saved fallback to {fallback_path}")

    print("Done generating files.")

# Example usage:
if __name__ == "__main__":
    input_xlsx = "input.xlsx"
    countries_list = ["Hong Kong", "India", "Singapore"]
    generate_country_files(input_xlsx, countries_list, output_dir="country_files")

```
