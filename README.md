```
import shutil
import pandas as pd
from openpyxl import load_workbook

def create_templates_by_site(findf, template_path="cvtemp.xlsx", output_prefix="cvtemp_"):
    """
    For each unique sitecode in findf:
      1. Copy the template file
      2. Write cust, prod, she columns into CustName, ProdName, OtgeropIncome
      3. Save as cvtemp_<sitecode>.xlsx
    """
    required_map = {
        "cust": "CustName",
        "prod": "ProdName",
        "she": "OtgeropIncome"
    }

    # Loop over each unique sitecode
    for site in findf["sitecode"].unique():
        site_df = findf[findf["sitecode"] == site].reset_index(drop=True)

        output_path = f"{output_prefix}{site}.xlsx"
        shutil.copy(template_path, output_path)

        # Load workbook
        wb = load_workbook(output_path)
        ws = wb.active  # adjust if multiple sheets

        # Find header positions (assume header in row 1)
        header = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}

        # Check that required columns exist
        for tgt in required_map.values():
            if tgt not in header:
                raise KeyError(f"Column '{tgt}' not found in template.")

        # Write data row by row
        for i, row in site_df.iterrows():
            excel_row = i + 2  # row 1 is header
            ws.cell(row=excel_row, column=header["CustName"], value=row["cust"])
            ws.cell(row=excel_row, column=header["ProdName"], value=row["prod"])
            ws.cell(row=excel_row, column=header["OtgeropIncome"], value=row["she"])

        # Save output
        wb.save(output_path)
        print(f"✅ Created template for site {site}: {output_path}")

```
