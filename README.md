```
import pandas as pd

def calculate_mtd_lag(df):
    # Get all YTD columns
    ytd_cols = [col for col in df.columns if col.endswith("_YTD")]

    # Sort YTD columns chronologically
    ytd_sorted = sorted(ytd_cols, key=lambda x: pd.to_datetime(x.replace("_YTD", ""), format="%b%y"))

    mtd_data = {}

    # Loop from the 1st index (compare with previous)
    for i in range(1, len(ytd_sorted)):
        curr_col = ytd_sorted[i]      # current YTD
        prev_col = ytd_sorted[i-1]    # previous YTD

        # The MTD should be named after NEXT month of prev_col
        prev_date = pd.to_datetime(prev_col.replace("_YTD", ""), format="%b%y")
        next_month = (prev_date + pd.offsets.MonthEnd(1)).strftime("%b%y")
        mtd_col = f"{next_month}_MTD"

        # Values = current YTD - previous YTD
        mtd_data[mtd_col] = df[curr_col] - df[prev_col]

    # Add MTD columns to df
    for col, values in mtd_data.items():
        df[col] = values

    return df
```
