```
import pandas as pd
import numpy as np

def gen_region_comm_hierarchical_conditional_offset(df_cy, df_py, metric_col, Level1, Level2, Level3,
                                                    top_n_level2=1, top_n_level3=3):
    """
    Hierarchical commentary:
      - For each Level1: explain by top Level2 (lead) and an offset Level2.
      - If the offset Level2 is negative (i.e. the worst Level2 Change < 0),
        use the bottom `top_n_level2` (most negative) to explain the offset.
      - For each Level2 mentioned: explain by its top Level3 products.

    Returns:
      para1, final_comm (strings)
    """
    # defensive copies
    cy = df_cy.copy()
    py = df_py.copy()

    # ---------- Aggregate Level1 ----------
    cy_l1 = cy.groupby([Level1])[[metric_col]].sum().reset_index()
    py_l1 = py.groupby([Level1])[[metric_col]].sum().reset_index()
    merged = cy_l1.merge(py_l1, how='left', on=Level1, suffixes=('_cy', '_py')).fillna(0)
    merged['Change'] = merged[f'{metric_col}_cy'] - merged[f'{metric_col}_py']
    merged['YoY%'] = np.where(merged[f'{metric_col}_py'] != 0,
                              (merged['Change'] / merged[f'{metric_col}_py']) * 100,
                              np.nan)
    merged.sort_values(by='Change', ascending=False, inplace=True)

    # ---------- Aggregate Level1+Level2 ----------
    cy_l2 = cy.groupby([Level1, Level2])[[metric_col]].sum().reset_index()
    py_l2 = py.groupby([Level1, Level2])[[metric_col]].sum().reset_index()
    merged_l2 = cy_l2.merge(py_l2, how='left', on=[Level1, Level2], suffixes=('_cy', '_py')).fillna(0)
    merged_l2['Change'] = merged_l2[f'{metric_col}_cy'] - merged_l2[f'{metric_col}_py']
    merged_l2['YoY%'] = np.where(merged_l2[f'{metric_col}_py'] != 0,
                                 (merged_l2['Change'] / merged_l2[f'{metric_col}_py']) * 100,
                                 np.nan)
    # sort Level2 within Level1 by Change desc
    merged_l2.sort_values(by=[Level1, 'Change'], ascending=[True, False], inplace=True)

    # ---------- Aggregate Level1+Level2+Level3 ----------
    cy_l3 = cy.groupby([Level1, Level2, Level3])[[metric_col]].sum().reset_index()
    py_l3 = py.groupby([Level1, Level2, Level3])[[metric_col]].sum().reset_index()
    merged_l3 = cy_l3.merge(py_l3, how='left', on=[Level1, Level2, Level3], suffixes=('_cy', '_py')).fillna(0)
    merged_l3['Change'] = merged_l3[f'{metric_col}_cy'] - merged_l3[f'{metric_col}_py']
    merged_l3['YoY%'] = np.where(merged_l3[f'{metric_col}_py'] != 0,
                                 (merged_l3['Change'] / merged_l3[f'{metric_col}_py']) * 100,
                                 np.nan)
    merged_l3.sort_values(by=[Level1, Level2, 'Change'], ascending=[True, True, False], inplace=True)

    # ---------- Compose commentary ----------
    positive_segments = merged[merged['Change'] > 0][Level1].tolist()
    sectors = ", ".join(map(str, positive_segments)) if positive_segments else "N/A"
    para1 = f"TRI growth across {sectors}."

    lines = [para1]

    # Helper to format Level3 items (used inside build_l2_with_l3)
    def format_l3_rows(l3_df):
        items = []
        for _, row in l3_df.iterrows():
            l3name = row[Level3]
            ch = row['Change']
            yoy = row['YoY%']
            yoy_str = f", {yoy:.1f}%" if not pd.isna(yoy) else ""
            items.append(f"{l3name} (+{ch:.1f}{yoy_str})")
        return ", ".join(items)

    # For each Level1 (ordered by Change), produce a hierarchical sentence
    for _, row_l1 in merged.iterrows():
        seg = row_l1[Level1]
        seg_change = row_l1['Change']
        seg_yoy = row_l1['YoY%']

        sign_word = "up" if seg_change > 0 else "down" if seg_change < 0 else "flat"
        l1_phrase = f"{seg} TRI {sign_word} by {seg_change:.1f}"
        if not pd.isna(seg_yoy):
            l1_phrase += f" ({seg_yoy:.1f}% YoY)"
        l1_phrase += "."

        # Level2 subset for this Level1
        l2_subset = merged_l2[merged_l2[Level1] == seg]
        if l2_subset.empty:
            lines.append(l1_phrase)
            continue

        # Lead Level2: top N by Change (desc)
        lead_l2 = l2_subset.sort_values('Change', ascending=False).head(top_n_level2)

        # Offset Level2 selection (conditional):
        # - If the worst Level2 (most negative) for this Level1 is negative, pick bottom N (most negative).
        # - Otherwise, pick top N (largest positive) as the offset (fallback).
        bottom_candidate = l2_subset.sort_values('Change', ascending=True).head(top_n_level2)
        use_bottom_for_offset = False
        if not bottom_candidate.empty:
            # check the worst change value
            worst_change = bottom_candidate['Change'].iloc[0]
            if worst_change < 0:
                use_bottom_for_offset = True

        if use_bottom_for_offset:
            offset_l2 = l2_subset.sort_values('Change', ascending=True).head(top_n_level2)
        else:
            # fallback: use top_n_level2 by Change descending (so 'offset' will be the lesser positives)
            offset_l2 = l2_subset.sort_values('Change', ascending=False).tail(top_n_level2)

        # Build descriptive phrase for Level2 including their Level3 drivers
        def build_l2_with_l3(df_l2_rows, descriptor):
            parts = []
            for _, r in df_l2_rows.iterrows():
                l2name = r[Level2]
                l2chg = r['Change']
                l2y = r['YoY%']
                l2_yoy_str = f" ({l2y:.1f}% YoY)" if not pd.isna(l2y) else ""
                # get top Level3 under (seg, l2name)
                l3_subset = merged_l3[
                    (merged_l3[Level1] == seg) &
                    (merged_l3[Level2] == l2name)
                ].sort_values('Change', ascending=False).head(top_n_level3)

                if not l3_subset.empty:
                    l3_str = format_l3_rows(l3_subset)
                    parts.append(f"{l2name} (+{l2chg:.1f}{l2_yoy_str}) driven by {l3_str}")
                else:
                    parts.append(f"{l2name} (+{l2chg:.1f}{l2_yoy_str})")
            return f"{descriptor} " + " & ".join(parts) + "."

        lead_phrase = build_l2_with_l3(lead_l2, "growth was led by")
        offset_phrase = build_l2_with_l3(offset_l2, "but was partially offset by")

        combined = f"{l1_phrase} {lead_phrase} {offset_phrase}"
        lines.append(combined)

    final_comm = "\n".join(lines)
    return para1, final_comm


# -------------------------
# Quick demo / test block
# -------------------------
if __name__ == "__main__":
    # sample data
    data_cy = [
        {"Region":"Corporate", "Sector":"Manufacturing", "Product":"Overdraft", "TRI":50},
        {"Region":"Corporate", "Sector":"Manufacturing", "Product":"FX", "TRI":30},
        {"Region":"Corporate", "Sector":"Energy",        "Product":"Trade Finance", "TRI":10},
        {"Region":"Business Banking", "Sector":"Healthcare","Product":"Term Loan", "TRI":20},
        {"Region":"Business Banking", "Sector":"Trading","Product":"CMS", "TRI":5},
    ]
    data_py = [
        {"Region":"Corporate", "Sector":"Manufacturing", "Product":"Overdraft", "TRI":35},
        {"Region":"Corporate", "Sector":"Manufacturing", "Product":"FX", "TRI":22},
        {"Region":"Corporate", "Sector":"Energy",        "Product":"Trade Finance", "TRI":22},
        {"Region":"Business Banking", "Sector":"Healthcare","Product":"Term Loan", "TRI":10},
        {"Region":"Business Banking", "Sector":"Trading","Product":"CMS", "TRI":15},
    ]
    df_cy = pd.DataFrame(data_cy).rename(columns={"TRI":"Total Relationship Income ($M)"})
    df_py = pd.DataFrame(data_py).rename(columns={"TRI":"Total Relationship Income ($M)"})

    para1, commentary = gen_region_comm_hierarchical_conditional_offset(
        df_cy, df_py,
        metric_col="Total Relationship Income ($M)",
        Level1="Region", Level2="Sector", Level3="Product",
        top_n_level2=2, top_n_level3=3
    )

    print("PARA1:")
    print(para1)
    print("\nFINAL COMMENTARY:")
    print(commentary)
