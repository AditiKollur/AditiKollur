```
import pandas as pd

# ================= SAMPLE INPUT =================
df = pd.DataFrame({
    "Booking Country": ["India", "USA"],
    "GPS.1": [10, 20],
    "GTS.1": [5, 7],
    "IB.1": [2, 3],
    "GPS.2": [11, 21],
    "GTS.2": [6, 8],
    "IB.2": [3, 4],
    "GPS.3": [12, 22],
    "GTS.3": [7, 9],
    "IB.3": [4, 5],
})

# ================= TRANSFORMATION =================

# Step 1: Convert columns into MultiIndex
new_cols = []
for c in df.columns:
    if c == "Booking Country":
        new_cols.append(("Booking Country", ""))
    else:
        metric, period = c.split(".")
        new_cols.append((metric, period))

df.columns = pd.MultiIndex.from_tuples(new_cols)

# Step 2: Reshape
out = (
    df
    .set_index(("Booking Country", ""))
    .stack(level=0)
    .reset_index()
)

# Step 3: Rename columns
out.columns = ["Booking Country", "Metric", "1", "2", "3"]

# ================= RESULT =================
print(out)
