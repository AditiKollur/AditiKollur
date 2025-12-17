```
def global_commentary(
    df_cy,
    df_py,
    metric_col,
    segment_col,
    region_col,
    country_col,
    biz_col,
    product_col,
    top_n=2
):
    """
    Global-level commentary (all regions combined)
    Mirrors Segments / Markets / Business Lines / Products logic
    """

    lines = []

    # ======================================================
    # 1. GLOBAL SUMMARY
    # ======================================================
    total_cy = df_cy[metric_col].sum()
    total_py = df_py[metric_col].sum()
    change = total_cy - total_py
    yoy = (change / total_py * 100) if total_py != 0 else None

    lines.append(
        f"Managed Total Relationship income globally of "
        f"{fmt_mn_bn(to_mn(total_cy, detect_unit(metric_col)))}, "
        f"{fmt_change_yoy(change, yoy, metric_col)}"
    )

    # ======================================================
    # 2. SEGMENTS (GLOBAL)
    # ======================================================
    seg_agg = drop_noise(
        compute_agg(df_cy, df_py, [segment_col], metric_col)
    )

    pos = seg_agg[seg_agg["Change"] > 0].sort_values("Change", ascending=False).head(top_n)
    neg = seg_agg[seg_agg["Change"] < 0].sort_values("Change").head(top_n)

    seg_text = (
        f"Segments - Globally, growth was led by "
        f"{join_items(pos, segment_col, metric_col)}"
    )
    if not neg.empty:
        seg_text += f". Partially offset by {join_items(neg, segment_col, metric_col)}."

    lines.append(seg_text)

    # ======================================================
    # 3. MARKETS (GLOBAL = REGIONS)
    # ======================================================
    reg_agg = drop_noise(
        compute_agg(df_cy, df_py, [region_col], metric_col)
    )

    pos_r = reg_agg[reg_agg["Change"] > 0].sort_values("Change", ascending=False).head(top_n)
    neg_r = reg_agg[reg_agg["Change"] < 0].sort_values("Change").head(top_n)

    def describe_region(r):
        region = r[region_col]
        base = f"{region} ({fmt_change_yoy(r['Change'], r['YoY%'], metric_col)})"

        c_agg = drop_noise(
            compute_agg(
                df_cy[df_cy[region_col] == region],
                df_py[df_py[region_col] == region],
                [country_col],
                metric_col
            )
        )

        if r["Change"] > 0:
            c_sel = c_agg[c_agg["Change"] > 0].sort_values("Change", ascending=False).head(top_n)
        else:
            c_sel = c_agg[c_agg["Change"] < 0].sort_values("Change").head(top_n)

        return f"{base} driven by {join_items(c_sel, country_col, metric_col)}" if not c_sel.empty else base

    main_regs = " and ".join(describe_region(r) for _, r in pos_r.iterrows())

    markets_text = (
        f"Markets - Globally, growth was led by {main_regs}"
    )

    if not neg_r.empty:
        offset_regs = " and ".join(describe_region(r) for _, r in neg_r.iterrows())
        markets_text += f", partially offset by {offset_regs}."

    lines.append(markets_text)

    # ======================================================
    # 4. BUSINESS LINES (GLOBAL)
    # ======================================================
    biz_agg = drop_noise(
        compute_agg(df_cy, df_py, [biz_col], metric_col)
    )

    pos_b = biz_agg[biz_agg["Change"] > 0].sort_values("Change", ascending=False).head(top_n)
    neg_b = biz_agg[biz_agg["Change"] < 0].sort_values("Change").head(top_n)

    biz_text = (
        f"Business Lines - Globally, growth was led by "
        f"{join_items(pos_b, biz_col, metric_col)}"
    )
    if not neg_b.empty:
        biz_text += f". Partially offset by {join_items(neg_b, biz_col, metric_col)}."

    lines.append(biz_text)

    # ======================================================
    # 5. PRODUCTS (GLOBAL)
    # ======================================================
    prod_agg = drop_noise(
        compute_agg(df_cy, df_py, [product_col], metric_col)
    )

    pos_p = prod_agg[prod_agg["Change"] > 0].sort_values("Change", ascending=False).head(top_n)
    neg_p = prod_agg[prod_agg["Change"] < 0].sort_values("Change").head(top_n)

    prod_text = (
        f"Products - Globally, growth was led by "
        f"{join_items(pos_p, product_col, metric_col)}"
    )
    if not neg_p.empty:
        prod_text += f". Partially offset by {join_items(neg_p, product_col, metric_col)}."

    lines.append(prod_text)

    return "\n".join(lines)


global_text = global_commentary(
    df_cy,
    df_py,
    metric_col="Total Relationship Income ($M)",
    segment_col="CIB Segment",
    region_col="Managed Region",
    country_col="Managed country",
    biz_col="Business Line",
    product_col="Product",
    top_n=2
)

print(global_text)
