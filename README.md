```
import pandas as pd
import shutil

def update_cv_template(findf, template_path="cvtemp.xlsx", output_path="cvtemp_copy.xlsx"):
    """
    Copies the cvtemp.xlsx template and fills columns from findf
    into specified columns of the copied template.
    
    findf: DataFrame with columns ['cust', 'prod', 'she']
    template_path: path to the original template Excel file
    output_path: path where the updated copy will be saved
    """
    
    # Step 1: Copy template
    shutil.copy(template_path, output_path)

    # Step 2: Load copied template into pandas
    df_template = pd.read_excel(output_path)

    # Step 3: Map data from findf into template
    # Ensure we don’t overflow if template is larger than data
    rows_to_fill = min(len(df_template), len(findf))
    df_template.loc[:rows_to_fill-1, "CustName"] = findf.loc[:rows_to_fill-1, "cust"].values
    df_template.loc[:rows_to_fill-1, "ProdName"] = findf.loc[:rows_to_fill-1, "prod"].values
    df_template.loc[:rows_to_fill-1, "OtgeropIncome"] = findf.loc[:rows_to_fill-1, "she"].values

    # Step 4: Save back into the copied file
    df_template.to_excel(output_path, index=False)

    print(f"Updated file saved at: {output_path}")
```
