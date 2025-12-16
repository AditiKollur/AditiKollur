```
def all_regions_commentary(
    df_cy,
    df_py,
    region_col,
    metric_col,
    segment_col,
    country_col,
    biz_col,
    return_type="dict"
):
    """
    Generates commentary for ALL regions in region_col.

    return_type:
        - "dict" → {region: commentary_text}
        - "df"   → pandas DataFrame (one row per region)
    """

    regions = (
        pd.Index(df_cy[region_col].dropna().unique())
        .union(pd.Index(df_py[region_col].dropna().unique()))
        .sort_values()
    )

    output = {}

    for region in regions:
        output[region] = region_commentary(
            df_cy=df_cy,
            df_py=df_py,
            region_col=region_col,
            region=region,
            metric_col=metric_col,
            segment_col=segment_col,
            country_col=country_col,
            biz_col=biz_col
        )

    if return_type == "df":
        rows = []
        for region, text in output.items():
            lines = text.split("\n")
            rows.append({
                region_col: region,
                "Summary": lines[0] if len(lines) > 0 else "",
                "Segments": lines[1] if len(lines) > 1 else "",
                "Markets": lines[2] if len(lines) > 2 else "",
                "Business Lines": lines[3] if len(lines) > 3 else "",
                "Full Commentary": text
            })
        return pd.DataFrame(rows)

    return output
all_out = all_regions_commentary(
    df_cy=df_cy,
    df_py=df_py,
    region_col="Managed Region",
    metric_col="Total Relationship Income ($M)",
    segment_col="CIB Segment",
    country_col="Managed country",
    biz_col="Business Line",
    return_type="dict"
)

for region, text in all_out.items():
    print(f"\n===== REGION: {region} =====\n")
    print(text)


df_out = all_regions_commentary(
    df_cy=df_cy,
    df_py=df_py,
    region_col="Managed Region",
    metric_col="Total Relationship Income ($M)",
    segment_col="CIB Segment",
    country_col="Managed country",
    biz_col="Business Line",
    return_type="df"
)

df_out

