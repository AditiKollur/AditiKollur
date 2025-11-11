```
import pandas as pd

def consolidate_excel(file_path, exclude_sheets):
    xls = pd.ExcelFile(file_path)
    consolidated = []

    for sheet in xls.sheet_names:
        if sheet in exclude_sheets:
            continue

        # Read entire sheet first
        df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)

        # Find header row (the row containing "Band")
        header_row = df_raw.apply(lambda row: row.astype(str).str.contains("Band", case=False, na=False)).any(axis=1)
        if not header_row.any():
            continue  # Skip if no header found
        header_idx = header_row.idxmax()

        # Read again using that row as header
        df = pd.read_excel(file_path, sheet_name=sheet, header=header_idx)

        # Keep only rows having values in 'Level 1' or 'Level 0'
        df = df[df['Level 1'].notna() | df['Level 0'].notna()]

        # Add sheet name
        df['SheetName'] = sheet

        consolidated.append(df)

    # Combine all sheets
    if consolidated:
        final_df = pd.concat(consolidated, ignore_index=True)
    else:
        final_df = pd.DataFrame()

    return final_df
