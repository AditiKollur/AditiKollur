```
import pandas as pd
import re

# ===================== SAMPLE DATA =====================
cost5 = pd.DataFrame(columns=[
    "CIB",
    "Gross Revenue",
    "GPS",
    "GPS.1",
    "IB",
    "Net Revenue",
    "GTS",
    "GPS.2",
    "Random Column"
])

# List of anchor (item) columns
l = ["Gross Revenue", "Net Revenue"]

# Product column pattern
prod_pattern = re.compile(r"^(GPS|GTS|IB)(\.\d+)?$")

# ===================== RENAMING LOGIC =====================
cols = list(cost5.columns)
new_cols = cols.copy()

# Find positions of anchor columns
anchor_positions = [i for i, c in enumerate(cols) if c in l]
anchor_positions.append(len(cols))  # sentinel

for idx in range(len(anchor_positions) - 1):
    start = anchor_positions[idx]
    end = anchor_positions[idx + 1]
    anchor_name = cols[start]

    # Rename prod columns between anchors
    for j in range(start + 1, end):
        col = cols[j]

        if prod_pattern.fullmatch(col):
            prod = col.split(".")[0]  # GPS.1 → GPS
            new_cols[j] = f"{prod}_{anchor_name}"

# Apply renamed columns
cost5.columns = new_cols

# ===================== RESULT =====================
print(cost5.columns.tolist())
