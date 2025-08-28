```
import pandas as pd
import shutil

def update_cv_template(findf, template_path="cvtemp.xlsx", output_path="cvtemp_copy.xlsx"):
    """
    Copies cvtemp.xlsx and fills CustName, ProdName, OtgeropIncome
    columns with values from findf (cust, prod, she).
    """
    # Step 1: Copy template
    shutil.copy(template_path, output_path)

    # Step 2: Load copied template
    df_template = pd.read_excel(output_path)

    # Step 3: Ensure target columns exist
    for col in ["CustName", "ProdName", "OtgeropIncome"]:
        if col not in df_template.columns:
            df_template[col] = None  # create if missing

    # Step 4: Find how many rows to update
    rows_to_fill = min(len(df_template), len(findf))
    if rows_to_fill > 0:
        df_template.loc[df_template.index[:rows_to_fill], "CustName"] = findf.loc[findf.index[:rows_to_fill], "cust"].values
        df_template.loc[df_template.index[:rows_to_fill], "ProdName"] = findf.loc[findf.index[:rows_to_fill], "prod"].values
        df_template.loc[df_template.index[:rows_to_fill], "OtgeropIncome"] = findf.loc[findf.index[:rows_to_fill], "she"].values

    # Step 5: Save updated copy
    df_template.to_excel(output_path, index=False)

    print(f"✅ Updated file saved at: {output_path}")

```
