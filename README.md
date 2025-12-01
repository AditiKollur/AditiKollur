```
import pandas as pd
import numpy as np

def gen_region_comm(df_cy, df_py, metric_col, Level1, Level2, Level3):
    """
    Generates commentary across a 3-level hierarchy (Level1 > Level2 > Level3).
    Returns: (para1, final_comm) where final_comm is the full commentary string.

    Inputs:
      - df_cy: current-period dataframe
      - df_py: prior-period dataframe
      - metric_col: metric column name (string) e.g. "Total Relationship Income ($M)"
      - Level1, Level2, Level3: hierarchy column names (strings)
    """

    # --- Defensive copy ---
    cy = df_cy.copy()
    py = df_py.copy()

    # --- 1) Aggregate at Level1 ---------------------------------------------
    cy_l1 = cy.groupby([Level1])[[metric_col]].sum().reset_index()
    py_l1 = py.groupby([Level1])[[metric_col]].sum().reset_index()

    merged = cy_l1.merge(py_l1, how='left', on=Level1, suffixes=('_cy', '_py'))
    merged = merged.fillna(0)
    merged['Change'] = merged[f'{metric_col}_cy'] - merged[f'{metric_col}_py']
    # safe YoY% (if py is zero -> set NaN)
    merged['YoY%'] = np.where(merged[f'{metric_col}_py'] != 0,
                               (merged['Change'] / merged[f'{metric_col}_py']) * 100,
                               np.nan)
    merged.sort_values(by='Change', ascending=False, inplace=True)

    # segments order string for para1
    segments_order = merged[merged['Change'] > 0][Level1].tolist()
    sectors = ", ".join(map(str, segments_order)) if segments_order else "N/A"
    para1 = f"TRI growth across {sectors}."

    # --- 2) Aggregate at Level1 + Level2 -----------------------------------
    cy_seg = cy.groupby([Level1, Level2])[[metric_col]].sum().reset_index()
    py_seg = py.groupby([Level1, Level2])[[metric_col]].sum().reset_index()

    merged_seg = cy_seg.merge(py_seg, how='left', on=[Level1, Level2], suffixes=('_cy', '_py'))
    merged_seg = merged_seg.fillna(0)
    merged_seg['Change'] = merged_seg[f'{metric_col}_cy'] - merged_seg[f'{metric_col}_py']
    merged_seg['YoY%'] = np.where(merged_seg[f'{metric_col}_py'] != 0,
                                  (merged_seg['Change'] / merged_seg[f'{metric_col}_py']) * 100,
                                  np.nan)
    merged_seg.sort_values(by=[Level1, 'Change'], ascending=[True, False], inplace=True)

    # --- 3) Aggregate at Level1 + Level2 + Level3 ---------------------------
    cy_seg_biz = cy.groupby([Level1, Level2, Level3])[[metric_col]].sum().reset_index()
    py_seg_biz = py.groupby([Level1, Level2, Level3])[[metric_col]].sum().reset_index()

    merged_seg_biz = cy_seg_biz.merge(py_seg_biz, how='left',
                                      on=[Level1, Level2, Level3],
                                      suffixes=('_cy', '_py'))
    merged_seg_biz = merged_seg_biz.fillna(0)
    merged_seg_biz['Change'] = merged_seg_biz[f'{metric_col}_cy'] - merged_seg_biz[f'{metric_col}_py']
    merged_seg_biz['YoY%'] = np.where(merged_seg_biz[f'{metric_col}_py'] != 0,
                                      (merged_seg_biz['Change'] / merged_seg_biz[f'{metric_col}_py']) * 100,
                                      np.nan)
    merged_seg_biz.sort_values(by=[Level1, Level2, 'Change'], ascending=[True, True, False], inplace=True)

    # --- Build commentary pieces --------------------------------------------
    final_parts = []
    final_parts.append(para1)

    # Level1 top & bottom
    if not merged.empty:
        top_all = merged.sort_values('Change', ascending=False).head(1)
        bot_all = merged.sort_values('Change', ascending=True).head(1)

        if not top_all.empty:
            t = top_all.iloc[0]
            final_parts.append(f"{t[Level1]} TRI up by {t['Change']:.1f} ({t['YoY%']:.1f}% YoY).")
        if not bot_all.empty:
            b = bot_all.iloc[0]
            final_parts.append(f"{b[Level1]} TRI down by {b['Change']:.1f} ({b['YoY%']:.1f}% YoY).")

    # Level2 commentary per Level1 (lead & offset)
    lvl1_list = merged[merged['Change'] != 0][Level1].tolist()
    for seg in lvl1_list:
        sub = merged_seg[merged_seg[Level1] == seg]
        if sub.empty:
            continue
        top2 = sub.sort_values('Change', ascending=False).head(1)
        bot2 = sub.sort_values('Change', ascending=True).head(1)
        if not top2.empty and not bot2.empty:
            t2 = top2.iloc[0]
            b2 = bot2.iloc[0]
            final_parts.append(
                f"In {seg}, growth was led by {t2[Level2]} (+{t2['Change']:.1f}) but was partially offset by {b2[Level2]} ({b2['Change']:.1f})."
            )

    # Overall TRI change
    total_cy = merged[f"{metric_col}_cy"].sum()
    total_py = merged[f"{metric_col}_py"].sum()
    total_change = total_cy - total_py
    total_yoy = (total_change / total_py) * 100 if total_py != 0 else np.nan
    final_parts.append(f"Overall TRI change {total_change:.1f} ({total_yoy:.1f}% YoY).")

    # Level3 Top products (with Level2 in brackets)
    top3 = merged_seg_biz.sort_values('Change', ascending=False).head(3)
    if not top3.empty:
        top_items = [
            f"{row[Level3]} ({row[Level2]}) +{row['Change']:.1f} / {row['YoY%']:.1f}%"
            for _, row in top3.iterrows()
        ]
        final_parts.append("Top products: " + ", ".join(top_items) + ".")

    # Level3 Weak products
    bot3 = merged_seg_biz.sort_values('Change', ascending=True).head(3)
    if not bot3.empty:
        bot_items = [
            f"{row[Level3]} ({row[Level2]}) {row['Change']:.1f} / {row['YoY%']:.1f}%"
            for _, row in bot3.iterrows()
        ]
        final_parts.append("Weak products: " + ", ".join(bot_items) + ".")

    # Level3 commentary within the top Level2 of the top Level1
    if not merged.empty:
        top_lvl1 = merged.sort_values('Change', ascending=False).head(1)[Level1].iloc[0]
        sub_lvl2 = merged_seg[merged_seg[Level1] == top_lvl1]
        if not sub_lvl2.empty:
            top_lvl2 = sub_lvl2.sort_values('Change', ascending=False).head(1)[Level2].iloc[0]
            sub_lvl3 = merged_seg_biz[
                (merged_seg_biz[Level1] == top_lvl1) &
                (merged_seg_biz[Level2] == top_lvl2)
            ].sort_values('Change', ascending=False).head(3)

            if not sub_lvl3.empty:
                lvl3_items = [
                    f"{r[Level3]} (+{r['Change']:.1f}, {r['YoY%']:.1f}%)"
                    for _, r in sub_lvl3.iterrows()
                ]
                final_parts.append(
                    f"In {top_lvl1}, {top_lvl2} was led by products: " + ", ".join(lvl3_items) + "."
                )

    final_comm = "\n".join(final_parts)
    return para1, final_comm


# ---------------------------
# Example usage (small demo)
# ---------------------------
if __name__ == "__main__":
    # small sample dataset to test the function
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

    para1, commentary = gen_region_comm(
        df_cy, df_py,
        metric_col="Total Relationship Income ($M)",
        Level1="Region", Level2="Sector", Level3="Product"
    )

    print("PARA1:")
    print(para1)
    print("\nFINAL COMMENTARY:")
    print(commentary)
