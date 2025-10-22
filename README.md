```
import pandas as pd

# Example dataframe
df = pd.DataFrame({
    "Region": ["Asia", "Asia", "Asia", "Europe", "Europe"],
    "Managed Country": ["India", "China", "Japan", "UK", "Germany"],
    "Biz_new_com": ["Lending", "Deposits", "Treasury", "Lending", "Deposits"],
    "Total Relationship Income_cy": [200, 180, 120, 250, 180],
    "Total Relationship Income_py": [150, 160, 100, 200, 150],
    "Change": [50, 20, 20, 50, 30],
    "YoY": [33.3, 12.5, 20.0, 25.0, 20.0]
})

# Generate commentary for each Region
commentary_list = []

for region, group in df.groupby("Region"):
    total_tri = group["Total Relationship Income_cy"].sum()
    total_change = group["Change"].sum()

    # Top 2 managed countries by Change
    top_countries = group.sort_values("Change", ascending=False).head(2)
    countries_text = ", ".join(
        f"{row['Managed Country']} ({row['Change']} / {row['YoY']}%)"
        for _, row in top_countries.iterrows()
    )

    # Top 2 Biz_new_com by Change
    top_biz = group.groupby("Biz_new_com", as_index=False)["Change", "YoY"].sum()
    top_biz = top_biz.sort_values("Change", ascending=False).head(2)
    biz_text = ", ".join(
        f"{row['Biz_new_com']} ({row['Change']} / {row['YoY']}%)"
        for _, row in top_biz.iterrows()
    )

    # Final commentary
    text = (
        f"{region} TRI of {total_tri}, growth of {total_change}, mainly from markets "
        f"{countries_text}. Product growth in {biz_text}."
    )
    commentary_list.append(text)

# Combine all commentaries
final_output = "\n".join(commentary_list)
print(final_output)
