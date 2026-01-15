import re

l = ["Region", "Country"]
prod = ["Prod", "Product"]

# build regex pattern
prod_pattern = r'^([^\w]*(' + '|'.join(prod) + r')[^\w\d]*[\d\W]*)$'

cols_to_keep = [
    col for col in df.columns
    if col in l or re.fullmatch(prod_pattern, col)
]

df_filtered = df[cols_to_keep]
