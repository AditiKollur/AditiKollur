```
import shutil
import pandas as pd
from openpyxl import load_workbook

def update_cv_template(findf, template_path="cvtemp.xlsx", output_path="cvtemp_copy.xlsx"):
    """
    Copies the template file and writes findf values (cust, prod, she)
    into CustName, ProdName, OtgeropIncome columns of the copied file.
    Keeps all formatting/macros intact.
    """
    # Step 1: Copy template file
    shutil.copy(template_path, output_path)

    # Step 2: Load copied file with openpyxl
    wb = load_workbook(output_path)
    ws = wb.active  # use first sheet (or specify by name)

    # Step 3: Find the header row and column positions
    header = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}  # assumes headers in row 1

    # Ensure required columns exist in template
    required_map = {
        "cust": "CustName",
        "prod": "ProdName",
        "she": "OtgeropIncome"
    }
    for src, tgt in required_map.items():
        if tgt not in header:
            raise KeyError(f"Column '{tgt}' not found in template.")

    # Step 4: Write data row by row
    for i, row in findf.iterrows():
        excel_row = i + 2  # +2 because Excel rows start at 1 and row 1 is header
        ws.cell(row=excel_row, column=header["CustName"], value=row["cust"])
        ws.cell(row=excel_row, column=header["ProdName"], value=row["prod"])
        ws.cell(row=excel_row, column=header["OtgeropIncome"], value=row["she"])

    # Step 5: Save updated file
    wb.save(output_path)
    print(f"✅ Updated file saved at: {output_path}")

```
