```
def markets_with_biz_inline(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    country_col: str, biz_col: str,
    metric_col: str,
    top_n: int = 2
):
    """
    Enhanced Markets drilldown that stays SINGLE SENTENCE.
    For each selected country, adds:
        <Country> (change/YoY) driven by <top1/bottom1 bizline (change/YoY)>
    """

    # Full region sign
    _, region_sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    # Region filter
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    country_agg = _compute_aggregates(cy_r, py_r, [country_col], metric_col)
    if country_agg.empty:
        return f"Markets - No data for {region_value}."

    # Top and bottom countries
    top_countries = country_agg.sort_values("Change", ascending=False).head(top_n)
    bottom_countries = country_agg.sort_values("Change", ascending=True).head(top_n)

    # Decide lead & offset sets
    if region_sign > 0:
        verb_main = "growth was led by"
        main_set = top_countries
        offset_set = bottom_countries
    elif region_sign < 0:
        verb_main = "contraction was led by"
        main_set = bottom_countries
        offset_set = top_countries
    else:
        verb_main = "movement was led by"
        main_set = top_countries
        offset_set = bottom_countries

    def describe_country_with_bizline(row):
        """
        Returns:
           "<Country> (+/-X / YoY%) driven by <BizLine> (+/-X / YoY%)"
        """

        country = row[country_col]

        # Filter for that country
        cy_c = cy_r[cy_r[country_col] == country]
        py_c = py_r[py_r[country_col] == country]

        biz_agg = _compute_aggregates(cy_c, py_c, [biz_col], metric_col)
        if biz_agg.empty:
            return f"{country} ({_fmt_change_yoy(row['Change'], row['YoY%'])})"

        # Top1 and Bottom1 biz lines
        top1 = biz_agg.sort_values("Change", ascending=False).head(1)
        bottom1 = biz_agg.sort_values("Change", ascending=True).head(1)

        # Country-specific sign
        if row["Change"] > 0:
            lead_biz = top1
        else:
            lead_biz = bottom1

        biz_part = _join_items(lead_biz, biz_col)

        return f"{country} ({_fmt_change_yoy(row['Change'], row['YoY%'])}) driven by {biz_part}"

    # Build main countries portion
    main_parts = [describe_country_with_bizline(r) for _, r in main_set.iterrows()]
    main_text = " and ".join(main_parts) if len(main_parts) == 2 else ", ".join(main_parts)

    # Build offset countries portion
    offset_parts = [describe_country_with_bizline(r) for _, r in offset_set.iterrows()]
    offset_text = " and ".join(offset_parts) if len(offset_parts) == 2 else ", ".join(offset_parts)

    # Final single-sentence paragraph
    return (
        f"Markets - In the {region_value} region, {verb_main} {main_text}, "
        f"partially offset by {offset_text}."
    )
