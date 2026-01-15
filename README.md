```
import pandas as pd
import re

# ================= KEEP ONLY REQUIRED COLUMNS =================
# Strictly allow only GPS/GTS/IB with .1 .2 .3
pattern = re.compile(r"^(GPS|GTS|IB)\.(1|2|3)$")

cols_to_keep = [
    c for c in cost5.columns
    if c == "CIB" or pattern.fullmatch(c)
]

cost5 = cost5[cols_to_keep]

# ================= TRANSFORMATION =================

# Convert columns to MultiIndex
new_cols = []
for c in cost5.columns:
    if c == "CIB":
        new_cols.append(("CIB", ""))
    else:
        metric, period = c.split(".")
        new_cols.append((metric, period))

cost5.columns = pd.MultiIndex.from_tuples(new_cols)

# Reshape to required format
out = (
    cost5
    .set_index(("CIB", ""))
    .stack(level=0)
    .reset_index()
)

# Rename final columns
out.columns = ["CIB", "Metric", "1", "2", "3"]

# ================= RESULT =================
print(out)
