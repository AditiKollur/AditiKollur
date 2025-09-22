```
import openpyxl
from openpyxl import load_workbook
from pathlib import Path

def copy_sheet_with_formatting(source_ws, target_ws):
    """
    Copy content + formatting + merged cells + column widths
    from source_ws to target_ws
    """
    for row in source_ws.iter_rows():
        for cell in row:
            new_cell = target_ws.cell(row=cell.row, column=cell.col_idx, value=cell.value)
            if cell.has_style:
                new_cell._style = cell._style

    # Copy column widths
    for col_letter, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for row_idx, dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[row_idx].height = dim.height

    # Copy merged cells
    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

def filter_rows(ws, country):
    """
    Delete rows that do not match:
    Booking Country == country AND Product Scope == 'Y'
    """
    # Find column indexes
    headers = [cell.value for cell in ws[1]]
    try:
        col_country = headers.index("Booking Country") + 1
        col_scope = headers.index("Product Scope") + 1
    except ValueError:
        return  # headers not found

    # Iterate from bottom up (to safely delete rows)
    for row in range(ws.max_row, 1, -1):
        val_country = ws.cell(row=row, column=col_country).value
        val_scope = ws.cell(row=row, column=col_scope).value
        if not (val_country == country and val_scope == "Y"):
            ws.delete_rows(row, 1)

def generate_country_files(input_file, countries, output_dir="output_files"):
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Load the input workbook
    wb = load_workbook(input_file, data_only=False)

    sheets_to_process = ["gbm", "cmb", "tm1gbm", "tm1cmb", "siteprod_snapshot"]

    for country in countries:
        wb_new = openpyxl.Workbook()
        wb_new.remove(wb_new.active)  # remove default empty sheet

        for sheet_name in sheets_to_process:
            source_ws = wb[sheet_name]
            target_ws = wb_new.create_sheet(sheet_name)
            copy_sheet_with_formatting(source_ws, target_ws)

            if sheet_name in ["gbm", "cmb", "tm1gbm", "tm1cmb"]:
                filter_rows(target_ws, country)

            if sheet_name == "siteprod_snapshot":
                target_ws["B2"] = country

        # Save output
        output_path = Path(output_dir) / f"{country}.xlsx"
        wb_new.save(output_path)

    print("✅ All country files generated successfully!")

# Example usage
input_file = "input.xlsx"
countries = ["Hong Kong", "India", "Singapore"]
generate_country_files(input_file, countries)


```
