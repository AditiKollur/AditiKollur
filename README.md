```
import pandas as pd
from docx import Document
from docx.shared import Pt
from datetime import datetime
import numpy as np

# ------------------------------------------------------------
# Sample data (replace with your own DataFrame)
# ------------------------------------------------------------
np.random.seed(42)
df = pd.DataFrame({
    'Region': np.random.choice(['North', 'South', 'East', 'West'], 100),
    'Country': np.random.choice(['India', 'USA', 'UK', 'Germany'], 100),
    'Product': np.random.choice(['ProdA', 'ProdB', 'ProdC'], 100),
    'Segment': np.random.choice(['Retail', 'Wholesale'], 100),
    'TRI': np.random.randint(80, 150, 100)
})


# ------------------------------------------------------------
# Helper: Generate commentary text (rule-based)
# ------------------------------------------------------------
def generate_tri_commentary(df):
    commentary = {}

    # 1️⃣ Region, Country, Product contribution
    region_sum = df.groupby('Region')['TRI'].sum().sort_values(ascending=False)
    top_region = region_sum.idxmax()
    top_region_val = region_sum.max()
    country_sum = df.groupby('Country')['TRI'].sum().sort_values(ascending=False)
    top_country = country_sum.idxmax()
    product_sum = df.groupby('Product')['TRI'].sum().sort_values(ascending=False)
    top_product = product_sum.idxmax()

    commentary['geo'] = (
        f"The {top_region} region recorded the highest TRI value of {top_region_val:.0f}, "
        f"indicating strong performance in this geography. "
        f"{top_country} contributed the most among all countries, reflecting its market maturity. "
        f"Across all geographies, {top_product} was the leading product in terms of TRI contribution."
    )

    # 2️⃣ Product snapshot by country
    top_by_country = (
        df.groupby(['Country', 'Product'])['TRI']
        .sum()
        .reset_index()
        .sort_values(['Country', 'TRI'], ascending=[True, False])
    )
    lines = []
    for c in top_by_country['Country'].unique():
        top_entry = top_by_country[top_by_country['Country'] == c].iloc[0]
        lines.append(f"In {c}, {top_entry['Product']} led with a TRI of {top_entry['TRI']:.0f}.")
    commentary['product_country'] = " ".join(lines)

    # 3️⃣ Segment-Region-Product contribution
    seg_summary = (
        df.groupby(['Segment', 'Region'])['TRI']
        .sum()
        .reset_index()
        .sort_values('TRI', ascending=False)
    )
    top_seg = seg_summary.iloc[0]
    commentary['segment_region'] = (
        f"The {top_seg['Segment']} segment in the {top_seg['Region']} region "
        f"showed the highest TRI of {top_seg['TRI']:.0f}. "
        "Retail segments generally outperform Wholesale across most regions, "
        "driven by broader customer reach and consistent sales volumes."
    )

    return commentary


# ------------------------------------------------------------
# Generate the Word report
# ------------------------------------------------------------
def generate_tri_report(df, out_path="tri_commentary_auto.docx"):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

    doc.add_heading("TRI Commentary Report (Automated – No LLMs)", level=1)
    doc.add_paragraph(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph("---")

    commentary = generate_tri_commentary(df)

    doc.add_heading("1. Geographical, Country & Product Contribution", level=2)
    doc.add_paragraph(commentary['geo'])
    doc.add_paragraph(" ")

    doc.add_heading("2. Product Snapshot by Country", level=2)
    doc.add_paragraph(commentary['product_country'])
    doc.add_paragraph(" ")

    doc.add_heading("3. Segment & Region-wise Product Contribution", level=2)
    doc.add_paragraph(commentary['segment_region'])
    doc.add_paragraph(" ")

    doc.add_paragraph("This commentary is auto-generated using rule-based logic (no LLMs).")
    doc.save(out_path)
    print(f"✅ TRI commentary report saved to: {out_path}")


# ------------------------------------------------------------
# Run the report
# ------------------------------------------------------------
generate_tri_report(df)
```
