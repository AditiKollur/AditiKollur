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

    # --- Product-level comparison for top segment ---
    prod_current = current_data[current_data["CIB SME Segment"] == top_segment]
    prod_prev = prev_year_data[prev_year_data["CIB SME Segment"] == top_segment]

    prod_curr_agg = prod_current.groupby("Product", dropna=False)["Total Relationship Income ($M)"].sum()
    prod_prev_agg = prod_prev.groupby("Product", dropna=False)["Total Relationship Income ($M)"].sum()

    prod_df = pd.concat([prod_curr_agg, prod_prev_agg], axis=1, keys=["Current", "Prev_Year"]).fillna(0)
    prod_df["YoY_Change"] = prod_df["Current"] - prod_df["Prev_Year"]
    prod_df["YoY%"] = prod_df.apply(
        lambda x: ((x["YoY_Change"] / x["Prev_Year"]) * 100) if x["Prev_Year"] != 0 else 0, axis=1
    )

    # --- Top 2 products ---
    top_products = prod_df.sort_values("YoY_Change", ascending=False).head(2).reset_index()

    # --- Build commentary ---
    doc = Document()
    month_name = datetime(1900, current_month, 1).strftime("%B")

    para1 = (
        f"Managed TRI of ${current_tri:.2f}M in {month_name} {current_year}, "
        f"YoY change of ${yoy_change_num:+.2f}M ({yoy_change_pct:+.1f}%) from {month_name} {prev_year}."
    )

    para2 = (
        f"Segments – Growth/Fall across all client segments. "
        f"Top-performing CIB SME segment: '{top_segment}' with YoY change of "
        f"${top_segment_yoy_num:+.2f}M ({top_segment_yoy_pct:+.1f}%). "
        f"Top contributing products: "
        f"{top_products.loc[0, 'Product']} (${top_products.loc[0, 'Current']:.2f}M, "
        f"YoY change ${top_products.loc[0, 'YoY_Change']:+.2f}M, {top_products.loc[0, 'YoY%']:+.1f}%), "
        f"and {top_products.loc[1, 'Product']} (${top_products.loc[1, 'Current']:.2f}M, "
        f"YoY change ${top_products.loc[1, 'YoY_Change']:+.2f}M, {top_products.loc[1, 'YoY%']:+.1f}%)."
    )

    doc.add_paragraph(para1)
    doc.add_paragraph(para2)

    # --- Save commentary to Word ---
    file_path = "TRI_Commentary_YoY_with_Percentage.docx"
    doc.save(file_path)
    print(f"✅ Commentary generated and saved as {file_path}")

    return para1, para2, seg_df, prod_df
