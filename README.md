```
import pandas as pd
import numpy as np
from itertools import combinations
from docx import Document
from docx.shared import Pt
from datetime import datetime

# --------------------------
# SAMPLE DATA (replace with your own)
# --------------------------
np.random.seed(42)
products = ['ProdA', 'ProdB', 'ProdC']
subproducts = ['Sub1', 'Sub2']
regions = ['North', 'South', 'East']
segments = ['Retail', 'Wholesale']

rows = []
for _ in range(200):
    rows.append({
        'Product': np.random.choice(products),
        'Subproduct': np.random.choice(subproducts),
        'Region': np.random.choice(regions),
        'Segment': np.random.choice(segments),
        'TRI': round(np.random.normal(100, 20), 2)  # numeric metric
    })

df = pd.DataFrame(rows)

# --------------------------
# FUNCTION TO GENERATE COMMENTARY REPORT
# --------------------------
def generate_tri_commentary_report(
    df,
    group_cols=['Product', 'Subproduct', 'Region', 'Segment'],
    metric_col='TRI',
    out_path='tri_commentary.docx'
):
    """Generates a Word report summarizing TRI commentaries for all group combinations."""

    # Ensure required columns exist
    missing = [c for c in group_cols + [metric_col] if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in DataFrame: {missing}")

    # Overall statistics
    overall_mean = df[metric_col].mean()
    overall_std = df[metric_col].std(ddof=0) or 1.0

    # Word document setup
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

    doc.add_heading('TRI Commentary Report', level=1)
    doc.add_paragraph(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"Overall {metric_col} Mean = {overall_mean:.2f} | Std Dev = {overall_std:.2f}")
    doc.add_paragraph('---')

    # Generate all possible combinations of grouping columns
    for r in range(1, len(group_cols) + 1):
        for combo in combinations(group_cols, r):
            combo_name = ', '.join(combo)
            doc.add_heading(f"Group by: {combo_name}", level=2)

            # Aggregate metrics
            grouped = (
                df.groupby(list(combo))[metric_col]
                .agg(['count', 'mean', 'median', 'std'])
                .reset_index()
            )
            grouped['std'] = grouped['std'].fillna(0)

            commentaries = []

            for _, row in grouped.iterrows():
                group_vals = ', '.join([f"{col}={row[col]}" for col in combo])
                cnt, mean_val, med_val, std_val = int(row['count']), row['mean'], row['median'], row['std']

                # Compute difference vs overall
                pct_diff = (mean_val - overall_mean) / overall_mean * 100
                zscore = (mean_val - overall_mean) / overall_std

                # Qualitative interpretation
                if zscore >= 1.5:
                    comparison = "much higher than"
                elif zscore >= 0.5:
                    comparison = "higher than"
                elif zscore > -0.5:
                    comparison = "around the same as"
                elif zscore > -1.5:
                    comparison = "lower than"
                else:
                    comparison = "much lower than"

                text = (
                    f"{group_vals}: n={cnt}, mean={mean_val:.2f}, median={med_val:.2f}, std={std_val:.2f}. "
                    f"This is {comparison} the overall mean ({overall_mean:.2f}), "
                    f"a {pct_diff:+.1f}% difference (z={zscore:.2f})."
                )

                if cnt <= 3:
                    text += " (Small sample size — interpret cautiously.)"

                commentaries.append((mean_val, text))

            # Sort by mean TRI (descending)
            commentaries.sort(key=lambda x: -x[0])

            # Add top & bottom highlights
            doc.add_paragraph("Top Highlights:", style='List Bullet')
            for mean_val, text in commentaries[:5]:
                doc.add_paragraph(text, style='List Bullet')

            doc.add_paragraph("Bottom Highlights:", style='List Bullet')
            for mean_val, text in commentaries[-5:]:
                doc.add_paragraph(text, style='List Number')

            doc.add_page_break()

    doc.save(out_path)
    print(f"✅ TRI Commentary report saved to: {out_path}")


# --------------------------
# RUN FUNCTION
# --------------------------
generate_tri_commentary_report(df)

```
