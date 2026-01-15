```
import re

l = ["Region", "Country"]
prod = ["ghy", "GFS"]

# regex: starts with prod, then only digits or special characters
pattern = r'^(' + '|'.join(p.lower() for p in prod) + r')[^a-zA-Z]*$'

cols_to_keep = [
    col for col in df.columns
    if col in l or re.fullmatch(pattern, col.lower())
]

df_filtered = df[cols_to_keep]
