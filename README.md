```
def gen_region_comm(df_cy, df_py, metric_col, Level1, Level2, Level3):
    """
    Generate commentary across hierarchical levels.
    - df_cy: current year dataframe (or current period)
    - df_py: prior year dataframe (or comparable period)
    - metric_col: name of metric column (e.g. 'Total Relationship Income ($M)')
    - Level1, Level2, Level3: column names for hierarchy (strings)
    Returns: paragraph header + final commentary string
    """

    # 1) Aggregate at Level1 (region/segment)
    cy = df_cy.groupby([Level1])[[metric_col]].sum().reset_index()
    py = df_py.groupby([Level1])[[metric_col]].sum().reset_index()

    merged = cy.merge(py, how='left', on=Level1, suffixes=('_cy', '_py'))
    merged['Change'] = merged[f'{metric_col}_cy'] - merged[f'{metric_col}_py']
    # avoid divide by zero; if py is zero, set YoY% = NaN or some handle
    merged['YoY%'] = (merged['Change'] / merged[f'{metric_col}_py'].replace({0: pd.NA})) * 100
    merged.sort_values(by='Change', ascending=False, inplace=True)

    # order of segments with positive change
    segments_order = merged[merged['Change'] > 0][Level1].tolist()
    a = str(segments_order).replace("'", "")
    sectors = a.replace('[', '').replace(']', '')

    para1 = f"TRI growth across {sectors}."

    # 2) Aggregation at Level1 + Level2
    cy_seg = df_cy.groupby([Level1, Level2])[[metric_col]].sum().reset_index()
    py_seg = df_py.groupby([Level1, Level2])[[metric_col]].sum().reset_index()
    merged_seg = cy_seg.merge(py_seg, how='left', on=[Level1, Level2], suffixes=('_cy', '_py'))
    merged_seg['Change'] = merged_seg[f'{metric_col}_cy'] - merged_seg[f'{metric_col}_py']
    merged_seg['YoY%'] = (merged_seg['Change'] / merged_seg[f'{metric_col}_py'].replace({0: pd.NA})) * 100
    # sort within each Level1 by Change descending
    merged_seg.sort_values(by=[Level1, 'Change'], ascending=[True, False], inplace=True)

    # 3) Aggregation at Level1 + Level2 + Level3 (biz/products)
    cy_seg_biz = df_cy.groupby([Level1, Level2, Level3])[[metric_col]].sum().reset_index()
    py_seg_biz = df_py.groupby([Level1, Level2, Level3])[[metric_col]].sum().reset_index()
    merged_seg_biz = cy_seg_biz.merge(py_seg_biz, how='left', on=[Level1, Level2, Level3], suffixes=('_cy', '_py'))
    merged_seg_biz['Change'] = merged_seg_biz[f'{metric_col}_cy'] - merged_seg_biz[f'{metric_col}_py']
    merged_seg_biz['YoY%'] = (merged_seg_biz['Change'] / merged_seg_biz[f'{metric_col}_py'].replace({0: pd.NA})) * 100
    merged_seg_biz.sort_values(by=[Level1, Level2, 'Change'], ascending=[True, True, False], inplace=True)

    # Helper to get top/bottom rows for a given Level1 value and optional Level2 filter
    def top_bottom_for_segment(df, level1_val, level2_val=None, top_n=2):
        q = df[df[Level1] == level1_val]
        if level2_val is not None:
            q = q[q[Level2] == level2_val]
        if q.empty:
            return pd.DataFrame()
        q_sorted = q.sort_values(by='Change', ascending=False)
        top = q_sorted.head(top_n)
        bot = q_sorted.tail(top_n).sort_values(by='Change')  # bottom (most negative)
        return top, bot

    # Build st1: top/bottom at Level1
    st1 = ""
    if not merged.empty:
        top_all = merged.head(1)
        bot_all = merged.tail(1)
        if not top_all.empty:
            top_name = top_all[Level1].iloc[0]
            top_change = top_all['Change'].iloc[0]
            top_yoy = top_all['YoY%'].iloc[0]
            st1 += f"{top_name} TRI up by {top_change:.1f} ({top_yoy:.1f}% YoY). "
        if not bot_all.empty:
            bot_name = bot_all[Level1].iloc[0]
            bot_change = bot_all['Change'].iloc[0]
            bot_yoy = bot_all['YoY%'].iloc[0]
            st1 += f"{bot_name} TRI down by {bot_change:.1f} ({bot_yoy:.1f}% YoY). "

    # Build st2: Level2 commentary for each Level1 where relevant (partially offset wording preserved)
    st2 = ""
    for seg in merged[merged['Change'] != 0][Level1].tolist():
        # pick top positive Level2 within this Level1 and bottom negative Level2 within this Level1
        sub = merged_seg[merged_seg[Level1] == seg]
        if sub.empty:
            continue
        top2 = sub.sort_values(by='Change', ascending=False).head(1)
        bot2 = sub.sort_values(by='Change', ascending=True).head(1)
        if not top2.empty and not bot2.empty:
            top2_name = top2[Level2].iloc[0]
            top2_change = top2['Change'].iloc[0]
            bot2_name = bot2[Level2].iloc[0]
            bot2_change = bot2['Change'].iloc[0]
            st2 += (f"In {seg}, growth was led by {top2_name} (+{top2_change:.1f}) but was partially offset by "
                    f"{bot2_name} ({bot2_change:.1f}). ")

    # Build st3: overall growth numbers (example: bps/growth percent)
    # (here we keep simple overall growth)
    total_grow_cy = merged[f'{metric_col}_cy'].sum()
    total_grow_py = merged[f'{metric_col}_py'].sum()
    total_change = total_grow_cy - total_grow_py
    total_grow_pct = (total_change / total_grow_py.replace({0: pd.NA})) * 100 if total_grow_py != 0 else pd.NA
    st3 = f"Overall TRI change {total_change:.1f} ({total_grow_pct:.1f}% YoY). "

    # Build st4: Level3 commentary but include Level2 context -> this is the requested change
    st4 = ""
    # We'll identify top/bottom products within each Level1 (optionally focusing on the top Level2(s) already identified)
    # For readability, gather top product(s) across the whole merged_seg_biz
    if not merged_seg_biz.empty:
        # get top N products overall
        overall_top = merged_seg_biz.sort_values(by='Change', ascending=False).head(3)
        overall_bot = merged_seg_biz.sort_values(by='Change', ascending=True).head(3)

        if not overall_top.empty:
            # create a sentence that lists product (Level3) and its Level2 for context
            top_parts = []
            for _, r in overall_top.iterrows():
                l3 = r[Level3]
                l2 = r[Level2]
                ch = r['Change']
                yoy = r['YoY%']
                top_parts.append(f"{l3} ({l2}) +{ch:.1f} / {yoy:.1f}%")
            st4 += "Top products: " + ", ".join(top_parts) + ". "

        if not overall_bot.empty:
            bot_parts = []
            for _, r in overall_bot.iterrows():
                l3 = r[Level3]
                l2 = r[Level2]
                ch = r['Change']
                yoy = r['YoY%']
                bot_parts.append(f"{l3} ({l2}) {ch:.1f} / {yoy:.1f}%")
            st4 += "Weak products: " + ", ".join(bot_parts) + ". "

    # If you specifically want Level3 commentary *within* the top Level2 for a given Level1, you can generate like this:
    # Example: if top Level1 is X, get its top Level2 and then list Level3 items inside that Level2
    if not merged.empty:
        top_level1 = merged.sort_values('Change', ascending=False).head(1)[Level1].iloc[0]
        # find top Level2 within that Level1
        sub2 = merged_seg[merged_seg[Level1] == top_level1]
        if not sub2.empty:
            top_level2_name = sub2.sort_values('Change', ascending=False).head(1)[Level2].iloc[0]
            # now get top Level3 items inside (top_level1, top_level2_name)
            top3_df = merged_seg_biz[(merged_seg_biz[Level1] == top_level1) & (merged_seg_biz[Level2] == top_level2_name)]
            if not top3_df.empty:
                top3_list = []
                for _, r in top3_df.sort_values('Change', ascending=False).head(3).iterrows():
                    top3_list.append(f"{r[Level3]} (+{r['Change']:.1f}, {r['YoY%']:.1f}%)")
                st4 += (f"In {top_level1}, {top_level2_name} was led by products: " +
                        ", ".join(top3_list) + ". ")

    # combine everything
    final_comm = para1 + "\n" + st1 + st2 + st3 + st4
    return para1, final_comm
