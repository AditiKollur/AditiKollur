```
import pandas as pd

# ---- Merge CY & PY ----
df = pd.merge(
    df_cy,
    df_py,
    on=["Region", "Managed Country", "Biz_new_com", "abc"],
    suffixes=("_cy", "_py")
)

# Compute CY-PY changes
df["Change"] = df["Total Relationship Income_cy"] - df["Total Relationship Income_py"]
df["YoY"] = (df["Change"] / df["Total Relationship Income_py"]) * 100

# ---- Function to build commentary ----
def generate_region_commentary(region_df, region_name):

    # REGION LEVEL SUMMARY
    total_tri = region_df["Total Relationship Income_cy"].sum()
    total_change = region_df["Change"].sum()

    # Top 2 countries by change
    top_countries = (
        region_df.groupby("Managed Country")[["Change", "YoY"]]
        .sum()
        .sort_values("Change", ascending=False)
        .head(2)
        .reset_index()
    )

    country_parts = [
        f"{row['Managed Country']} ({row['Change']:.0f} / {row['YoY']:.1f}%)"
        for _, row in top_countries.iterrows()
    ]
    country_text = ", ".join(country_parts)

    # REGION-LEVEL PRODUCT GROWTH (top 2)
    top_products = (
        region_df.groupby("Biz_new_com")[["Change", "YoY"]]
        .sum()
        .sort_values("Change", ascending=False)
        .head(2)
        .reset_index()
    )

    prod_parts = [
        f"{row['Biz_new_com']} ({row['Change']:.0f} / {row['YoY']:.1f}%)"
        for _, row in top_products.iterrows()
    ]
    prod_text = ", ".join(prod_parts)

    commentary = []
    commentary.append(
        f"{region_name} TRI of {total_tri:.0f}, growth of {total_change:.0f}, mainly from markets {country_text}."
    )

    # ---- COUNTRY BREAKDOWN ----
    for country, c_df in region_df.groupby("Managed Country"):

        tri_c = c_df["Total Relationship Income_cy"].sum()
        change_c = c_df["Change"].sum()

        commentary.append(f"{country} TRI of {tri_c:.0f}, growth of {change_c:.0f}.")

        # Top products under the country
        country_prod = (
            c_df.groupby("Biz_new_com")[["Change", "YoY"]]
            .sum()
            .sort_values("Change", ascending=False)
            .head(2)
            .reset_index()
        )

        prod_summary = [
            f"{row['Biz_new_com']} ({row['Change']:.0f} / {row['YoY']:.1f}%)"
            for _, row in country_prod.iterrows()
        ]
        commentary.append("Top products: " + ", ".join(prod_summary) + ".")

        # ---- PRODUCT-LEVEL with ABC ----
        for product, p_df in c_df.groupby("Biz_new_com"):

            # ABC breakdown
            abc_rows = (
                p_df.groupby("abc")[["Change", "YoY"]]
                .sum()
                .reset_index()
            )

            abc_parts = [
                f"{row['abc'].upper()} ({row['Change']:.0f} / {row['YoY']:.1f}%)"
                for _, row in abc_rows.iterrows()
            ]

            commentary.append(
                f"ABC Split under {product}: " + ", ".join(abc_parts) + "."
            )

    # REGION-LEVEL PRODUCT SUMMARY (final line)
    commentary.append("Overall Product growth mainly from " + prod_text + ".")

    return "\n".join(commentary)


# ---- Generate commentary for all regions ----
final_output = []
for region, r_df in df.groupby("Region"):
    final_output.append(generate_region_commentary(r_df, region))

print("\n\n".join(final_output))
