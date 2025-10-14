```
from docx import Document
from docx.shared import RGBColor
import pandas as pd

# Example DataFrame
data = {
    "Business Line": ["BL1", "BL2", "BL3"],
    "Product": ["P1", "P2", "P3"],
    "YoY Value": [12, 10, -5],
    "YoY %": [8, 5, -3],
    "Total Relationship Income": [120, 100, 50],
}

df = pd.DataFrame(data)

# Sort by Total Relationship Income to get top and bottom performers
df_sorted = df.sort_values(by="Total Relationship Income", ascending=False).reset_index(drop=True)

top1 = df_sorted.iloc[0]
top2 = df_sorted.iloc[1]
bottom = df_sorted.iloc[-1]

# Example products contributing to top1 business line
# In real logic, you might filter products by that BL and sort them by contribution
top_products = df[df["Business Line"] == top1["Business Line"]].sort_values(by="Total Relationship Income", ascending=False)["Product"].head(2).tolist()

# Create Word document
doc = Document()

# Heading
doc.add_heading("Total Relationship Income", level=1)

# Helper function to add colored run
def add_colored_run(paragraph, text, value):
    run = paragraph.add_run(text)
    if value > 0:
        run.font.color.rgb = RGBColor(0, 128, 0)  # Green
    else:
        run.font.color.rgb = RGBColor(255, 0, 0)  # Red
    return run

# --- Commentary 1: Top BL ---
p1 = doc.add_paragraph("Products – Strong ")
p1.add_run(top1["Business Line"]).bold = True
p1.add_run(" performance (")

add_colored_run(p1, f"{top1['YoY Value']} / {top1['YoY %']}%)", top1["YoY Value"])

p1.add_run(" driven by ")
p1.add_run(top_products[0]).italic = True
if len(top_products) > 1:
    p1.add_run(" and ")
    p1.add_run(top_products[1]).italic = True

# --- Commentary 2: Second BL ---
p2 = doc.add_paragraph()
p2.add_run(top2["Business Line"]).bold = True
p2.add_run(" (")
add_colored_run(p2, f"{top2['YoY Value']} / {top2['YoY %']}%)", top2["YoY Value"])
p2.add_run(" saw strong growth.")

# --- Commentary 3: Last BL ---
p3 = doc.add_paragraph("Conversely, ")
p3.add_run(bottom["Business Line"]).bold = True
p3.add_run(" was down (")
add_colored_run(p3, f"{bottom['YoY Value']} / {bottom['YoY %']}%)", bottom["YoY Value"])
p3.add_run(" despite strong Balance Sheet growth.")

# Save Word file
doc.save("Total_Relationship_Income_Commentary.docx")
print("✅ Word commentary created: Total_Relationship_Income_Commentary.docx")
