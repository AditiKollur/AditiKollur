```
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from copy import copy
from pathlib import Path

NORMAL_SHEETS = ["gbm", "cmb", "tm1gbm", "tm1cmb"]
ALL_SHEETS = NORMAL_SHEETS + ["siteprod_snapshot"]

def copy_styles_only(src_cell, tgt_cell):
    """Copy style/formatting but not formula"""
    if hasattr(src_cell, "font") and src_cell.font is not None:
        tgt_cell.font = copy(src_cell.font)
    if hasattr(src_cell, "border") and src_cell.border is not None:
        tgt_cell.border = copy(src_cell.border)
    if hasattr(src_cell, "fill") and src_cell.fill is not None:
        tgt_cell.fill = copy(src_cell.fill)
    if hasattr(src_cell, "number_format") and src_cell.number_format is not None:
        tgt_cell.number_format = copy(src_cell.number_format)
    if hasattr(src_cell, "alignment") and src_cell.alignment is not None:
        tgt_cell.alignment = copy(src_cell.alignment)
    if hasattr(src_cell, "protection") and src_cell.protection is not None:
        tgt_cell.protection = copy(src_cell.protection)

def copy_sheet_with_formatting(src_ws, tgt_ws):
    """Full copy: values + formulas + formatting"""
    # Copy merged cells first
    for merged_range in src_ws.merged_cells.ranges:
        tgt_ws.merge_cells(str(merged_range))

    for row in src_ws.iter_rows():
        for cell in row:
            col_letter, _ = coordinate_from_string(cell.coordinate)
            col_idx = column_index_from_string(col_letter)
            new_cell = tgt_ws.cell(row=cell.row, column=col_idx, value=cell.value)
            copy_styles_only(cell, new_cell)

    # Copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        tgt_ws.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for idx, dim in src_ws.row_dimensions.items():
        tgt_ws.row_dimensions[idx].height = dim.height

def copy_filtered_data(src_ws, tgt_ws, country):
    """Copy only filtered rows: values (no formulas) + formatting"""
    # Find header row
    header_row, headers = None, None
    for r in range(1, 11):
        row_values = [c.value for c in src_ws[r]]
        if all(h in row_values for h in ("Booking Country", "Product Scope")):
            header_row, headers = r, row_values
            break
    if header_row is None:
        return

    col_country = headers.index("Booking Country") + 1
    col_scope = headers.index("Product Scope") + 1

    tgt_row = 1
    for r in range(1, src_ws.max_row + 1):
        val_country = src_ws.cell(row=r, column=col_country).value
        val_scope = src_ws.cell(row=r, column=col_scope).value

        keep = (r <= header_row) or (
            val_country == country and str(val_scope).strip() == "Y"
        )
        if keep:
            for cell in src_ws[r]:
                col_letter, _ = coordinate_from_string(cell.coordinate)
                col_idx = column_index_from_string(col_letter)
                new_cell = tgt_ws.cell(row=tgt_row, column=col_idx, value=cell.value)
                copy_styles_only(cell, new_cell)
            tgt_row += 1

    # Copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        tgt_ws.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for idx, dim in src_ws.row_dimensions.items():
        tgt_ws.row_dimensions[idx].height = dim.height

def generate_country_files(input_file, countries, output_dir="output_files"):
    input_file = Path(input_file)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    src_wb = load_workbook(input_file, data_only=False)  # keep formulas

    for country in countries:
        print(f"Processing: {country}")
        tgt_wb = openpyxl.Workbook()
        tgt_wb.remove(tgt_wb.active)

        for sheet_name in ALL_SHEETS:
            if sheet_name not in src_wb.sheetnames:
                continue
            src_ws = src_wb[sheet_name]
            tgt_ws = tgt_wb.create_sheet(title=sheet_name)

            if sheet_name in NORMAL_SHEETS:
                copy_filtered_data(src_ws, tgt_ws, country)
            else:  # siteprod_snapshot
                copy_sheet_with_formatting(src_ws, tgt_ws)
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
