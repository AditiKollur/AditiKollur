```
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Pt
from datetime import datetime
from openai import OpenAI

# Initialize your OpenAI client (set your API key)
client = OpenAI(api_key="YOUR_API_KEY_HERE")

# --------------------------
# Sample data (replace with your own)
# --------------------------
np.random.seed(42)
df = pd.DataFrame({
    'Region': np.random.choice(['North', 'South', 'East', 'West'], 200),
    'Country': np.random.choice(['India', 'USA', 'UK', 'Germany'], 200),
    'Product': np.random.choice(['ProdA', 'ProdB', 'ProdC'], 200),
    'Segment': np.random.choice(['Retail', 'Wholesale'], 200),
    'TRI': np.random.randint(80, 150, 200)
})

# --------------------------
# Helper to query LLM
# --------------------------
def ask_llm(prompt: str) -> str:
    """Send prompt to LLM and return response text."""
    completion = client.chat.completions.create(
        model="gpt-4o-mini",  # or gpt-5 when available
        messages=[
            {"role": "system", "content": "You are a financial analyst writing concise, insightful commentary."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7,
        max_tokens=600
    )
    return completion.choices[0].message.content.strip()

# --------------------------
# Generate TRI Commentary Report using LLM
# --------------------------
def generate_llm_tri_report(df, out_path="tri_llm_commentary.docx"):
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)
    doc.add_heading("LLM-Generated TRI Commentary Report", level=1)
    doc.add_paragraph(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph('---')

    # 1️⃣ Geographical Region & Country Contribution
    doc.add_heading("1. Geographical & Product Contribution", level=2)
    region_summary = (
        df.groupby(['Region', 'Country', 'Product'])['TRI']
        .sum()
        .reset_index()
        .sort_values('TRI', ascending=False)
    )
    prompt_geo = (
        "Analyze the following TRI data grouped by region, country, and product.\n"
        "Write a 2-3 paragraph analytical commentary describing:\n"
        "• Which regions contribute the most to TRI\n"
        "• Which countries and products dominate each region\n"
        "• Any patterns or outliers you observe\n\n"
        f"{region_summary.head(20).to_string(index=False)}"
    )
    geo_text = ask_llm(prompt_geo)
    doc.add_paragraph(geo_text)

    doc.add_paragraph('---')

    # 2️⃣ Product Snapshot by Country
    doc.add_heading("2. Product Snapshot by Country", level=2)
    prod_summary = (
        df.groupby(['Country', 'Product'])['TRI']
        .sum()
        .reset_index()
        .sort_values(['Country', 'TRI'], ascending=[True, False])
    )
    prompt_prod = (
        "Using the TRI summary below, write country-level product commentary.\n"
        "Highlight top products per country, comparative performance, and insights.\n\n"
        f"{prod_summary.head(30).to_string(index=False)}"
    )
    prod_text = ask_llm(prompt_prod)
    doc.add_paragraph(prod_text)

    doc.add_paragraph('---')

    # 3️⃣ Segment & Region-wise Product Contribution
    doc.add_heading("3. Segment & Region-wise Product Contribution", level=2)
    seg_summary = (
        df.groupby(['Segment', 'Region', 'Product'])['TRI']
        .sum()
        .reset_index()
        .sort_values('TRI', ascending=False)
    )
    prompt_seg = (
        "Analyze the following TRI data by segment, region, and product.\n"
        "Write a paragraph describing how segments and regions differ in TRI contribution,\n"
        "highlighting the top and underperforming combinations.\n\n"
        f"{seg_summary.head(25).to_string(index=False)}"
    )
    seg_text = ask_llm(prompt_seg)
    doc.add_paragraph(seg_text)

    doc.add_paragraph('---')
    doc.add_paragraph("This report was generated using an LLM to provide automated narrative insights from structured TRI data.")

    # Save file
    doc.save(out_path)
    print(f"✅ LLM-generated TRI commentary report saved to: {out_path}")

# --------------------------
# Run it
# --------------------------
generate_llm_tri_report(df)

```
