```
import pandas as pd
from datetime import datetime, timedelta
from docx import Document

def excel_serial_to_date(serial):
    """Convert Excel serial date number to datetime"""
    return datetime(1899, 12, 30) + timedelta(days=int(serial))

def generate_tri_commentary_yoy(df):
    # --- Convert Excel serials to datetime ---
    df["Date"] = df["Year-Month"].apply(excel_serial_to_date)
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month

    # --- Combine UK RFB + NRFB + Europe into single 'Europe' region ---
    df["Managed region"] = df["Managed region"].replace(
        {"United Kingdom - RFB": "Europe", "United Kingdom - NRFB": "Europe"}
    )

    # --- Detect current and previous year/month ---
    latest_date = df["Date"].max()
    current_month = latest_date.month
    current_year = latest_date.year
    prev_year = current_year - 1

    # --- Filter current & previous year same month ---
    current_data = df[(df["Month"] == current_month) & (df["Year"] == current_year)]
    prev_year_data = df[(df["Month"] == current_month) & (df["Year"] == prev_year)]

    # --- Managed TRI total ---
    current_tri = current_data["Total Relationship Income ($M)"].sum()
    prev_year_tri = prev_year_data["Total Relationship Income ($M)"].sum()

    yoy_change_num = current_tri - prev_year_tri
    yoy_change_pct = ((yoy_change_num / prev_year_tri) * 100) if prev_year_tri != 0 else 0

    # --- Segment-level comparison ---
    seg_current = current_data.groupby("CIB SME Segment", dropna=False)["Total Relationship Income ($M)"].sum()
    seg_prev = prev_year_data.groupby("CIB SME Segment", dropna=False)["Total Relationship Income ($M)"].sum()

    seg_df = pd.concat([seg_current, seg_prev], axis=1, keys=["Current", "Prev_Year"]).fillna(0)
    seg_df["YoY_Change"] = seg_df["Current"] - seg_df["Prev_Year"]
    seg_df["YoY%"] = seg_df.apply(
        lambda x: ((x["YoY_Change"] / x["Prev_Year"]) * 100) if x["Prev_Year"] != 0 else 0, axis=1
    )

    # --- Top segment by YoY absolute growth ---
    top_segment = seg_df["YoY_Change"].idxmax()
    top_segment_yoy_num = seg_df.loc[top_segment, "YoY_Change"]
    top_segment_yoy_pct = seg_df.loc[top_segment, "YoY%"]

    # --- Business Line-level comparison for top segment ---
    bl_current = current_data[current_data["CIB SME Segment"] == top_segment]
    bl_prev = prev_year_data[prev_year_data["CIB SME Segment"] == top_segment]

    bl_curr_agg = bl_current.groupby("Business Line", dropna=False)["Total Relationship Income ($M)"].sum()
    bl_prev_agg = bl_prev.groupby("Business Line", dropna=False)["Total Relationship Income ($M)"].sum()

    bl_df = pd.concat([bl_curr_agg, bl_prev_agg], axis=1, keys=["Current", "Prev_Year"]).fillna(0)
    bl_df["YoY_Change"] = bl_df["Current"] - bl_df["Prev_Year"]
    bl_df["YoY%"] = bl_df.apply(
        lambda x: ((x["YoY_Change"] / x["Prev_Year"]) * 100) if x["Prev_Year"] != 0 else 0, axis=1
    )

    # --- Top 2 Business Lines ---
    top_bl = bl_df.sort_values("YoY_Change", ascending=False).head(2).reset_index()

    # --- Region-level analysis (after merging Europe+UKs) ---
    region_current = current_data.groupby("Managed region", dropna=False)["Total Relationship Income ($M)"].sum()
    region_prev = prev_year_data.groupby("Managed region", dropna=False)["Total Relationship Income ($M)"].sum()

    region_df = pd.concat([region_current, region_prev], axis=1, keys=["Current", "Prev_Year"]).fillna(0)
    region_df["YoY_Change"] = region_df["Current"] - region_df["Prev_Year"]
    region_df["YoY%"] = region_df.apply(
        lambda x: ((x["YoY_Change"] / x["Prev_Year"]) * 100) if x["Prev_Year"] != 0 else 0, axis=1
    )
    region_df = region_df.sort_values("YoY_Change", ascending=False).reset_index()

    # --- Top 5 regions ---
    top5_regions = region_df.head(5)

    # --- Top 2 countries within top 2 regions ---
    region_country_current = current_data.groupby(["Managed region", "Managed Country"], dropna=False)["Total Relationship Income ($M)"].sum()
    region_country_prev = prev_year_data.groupby(["Managed region", "Managed Country"], dropna=False)["Total Relationship Income ($M)"].sum()

    region_country_df = pd.concat([region_country_current, region_country_prev], axis=1, keys=["Current", "Prev_Year"]).fillna(0)
    region_country_df["YoY_Change"] = region_country_df["Current"] - region_country_df["Prev_Year"]
    region_country_df["YoY%"] = region_country_df.apply(
        lambda x: ((x["YoY_Change"] / x["Prev_Year"]) * 100) if x["Prev_Year"] != 0 else 0, axis=1
    )
    region_country_df = region_country_df.reset_index()

    def top_countries(region_name):
        sub = region_country_df[region_country_df["Managed region"] == region_name]
        return sub.sort_values("YoY_Change", ascending=False).head(2)[["Managed Country", "YoY_Change", "YoY%"]]

    # --- Round numbers to 0 decimals everywhere ---
    round_cols = ["Current", "Prev_Year", "YoY_Change", "YoY%"]
    for df_ in [seg_df, bl_df, region_df, region_country_df]:
        df_[round_cols] = df_[round_cols].round(0)

    # --- Build commentary text ---
    doc = Document()
    month_name = datetime(1900, current_month, 1).strftime("%B")

    para1 = (
        f"Managed TRI of ${current_tri:.0f}M in {month_name} {current_year}, "
        f"YoY change of ${yoy_change_num:+.0f}M ({yoy_change_pct:+.0f}%) from {month_name} {prev_year}."
    )

    para2 = (
        f"Segments – Growth/Fall across all client segments. "
        f"Top-performing CIB SME segment: '{top_segment}' with YoY change of "
        f"${top_segment_yoy_num:+.0f}M ({top_segment_yoy_pct:+.0f}%). "
        f"Top contributing Business Lines: "
        f"{top_bl.loc[0, 'Business Line']} (${top_bl.loc[0, 'Current']:.0f}M, "
        f"YoY change ${top_bl.loc[0, 'YoY_Change']:+.0f}M, {top_bl.loc[0, 'YoY%']:+.0f}%), "
        f"and {top_bl.loc[1, 'Business Line']} (${top_bl.loc[1, 'Current']:.0f}M, "
        f"YoY change ${top_bl.loc[1, 'YoY_Change']:+.0f}M, {top_bl.loc[1, 'YoY%']:+.0f}%)."
    )

    # --- Third commentary (Region growth summary) ---
    top_regions = top5_regions["Managed region"].tolist()
    top1, top2, top3, top4, top5 = top_regions[:5]
    c1 = top_countries(top1)
    c2 = top_countries(top2)

    para3 = (
        f"Regions – Strong growth in {top1} (${top5_regions.loc[0, 'YoY_Change']:+.0f}M, "
        f"{top5_regions.loc[0, 'YoY%']:+.0f}%) led by {c1.iloc[0, 0]} and {c1.iloc[1, 0]}, "
        f"followed by {top2} (${top5_regions.loc[1, 'YoY_Change']:+.0f}M, "
        f"{top5_regions.loc[1, 'YoY%']:+.0f}%) driven by {c2.iloc[0, 0]} and {c2.iloc[1, 0]}. "
        f"Accompanied by steady growth in {top3} (${top5_regions.loc[2, 'YoY_Change']:+.0f}M, "
        f"{top5_regions.loc[2, 'YoY%']:+.0f}%), {top4} (${top5_regions.loc[3, 'YoY_Change']:+.0f}M, "
        f"{top5_regions.loc[3, 'YoY%']:+.0f}%), and {top5} (${top5_regions.loc[4, 'YoY_Change']:+.0f}M, "
        f"{top5_regions.loc[4, 'YoY%']:+.0f}%)."
    )

    # --- Add paragraphs ---
    doc.add_paragraph(para1)
    doc.add_paragraph(para2)
    doc.add_paragraph(para3)

    # --- Save the Word file ---
    file_path = "TRI_Commentary_YoY_Final.docx"
    doc.save(file_path)
    print(f"✅ Commentary generated and saved as {file_path}")

    return para1, para2, para3, seg_df, bl_df, region_df
