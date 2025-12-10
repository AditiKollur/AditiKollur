```
output = all_regions_commentary(
    df_cy,
    df_py,
    region_col="Managed Region",
    metric_col="Total Relationship Income ($M)",
    segment_col="CIB Segment",
    market_col="Managed country",
    top_n=2
)

for region, text in output.items():
    print("====", region, "====")
    print(text)
    print()
