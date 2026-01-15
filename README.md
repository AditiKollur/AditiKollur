```
top_column = None

for col in df.columns:
    if df[col].astype(str).str.contains("Top", case=False, na=False).any():
        top_column = col
        break

print("Column containing 'Top':", top_column)
