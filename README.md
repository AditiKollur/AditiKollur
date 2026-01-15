```
import pandas as pd

# example
l = ["Region", "Country", "Date"]
prod = ["Prod", "Product", "SKU"]

# keep columns
cols_to_keep = [
    col for col in df.columns
    if col in l or any(p.lower() in col.lower() for p in prod)
]

df_filtered = df[cols_to_keep]
