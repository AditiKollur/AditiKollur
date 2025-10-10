```
import pandas as pd
from datetime import datetime, timedelta
from docx import Document

def excel_serial_to_date(serial):
    """Convert Excel serial date number to datetime"""
    return datetime(1899, 12, 30) + timedelta(days=int(serial))

def generate_tri_commentary(df):
    # Convert Excel serial to datetime
    df["Date"] = df["Year-Month"].apply(excel_serial_to_date)
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month

    # Identify current and previous months
    current_month = df["Month"].max()
    current_year = df.loc[df["Month"] == current_month, "Year"].max()

    prev_year = current_year - 1
    prev_month = current_month - 1 if current_month > 1 else 12
    prev_month_year = current_year if current_month > 1 else current_year - 1

    # Filter data
    current_data = df[(df["Month"] == current_month) & (df["Year"] == current_year)]
    prev_year_data = df[(df["Month"] == current_month) & (df["Year"] == prev_year)]
    prev_month_data = df[(df["Month"] == prev_month) & (df["Year"] == prev_month_year)]

    # Managed TRI totals
    current_tri = current_data["Total Relationship Income ($M)"].sum()
    prev_year_tri = prev_year_data["Total Relationship Income ($M)"].sum()
    prev_month_tri = prev_month_data["Total Relationship Income ($M)"].sum()

    yoy_change = ((current_tri - prev_year_tri) / prev_year_tri * 100) if prev_year_tri != 0 else 0
    mom_change = current_tri - prev_month_tri  # absolute change, not %

    # Segment-level analysis
    seg_current = current_data.groupby("CIB SME Segment")["Total Relationship Income ($M)"].sum()
    seg_prev_year = prev_year_data.groupby("CIB SME Segment")["Total Relationship Income ($M)"].sum()
    seg_prev_month = prev_month_data.groupby("CIB SME Segment")["Total Relationship Income ($M)"].sum()

    seg_df = pd.DataFrame({
        "Current": seg_current,
        "Prev_Year": seg_prev_year,
        "Prev_Month": seg_prev_month
    }).fillna(0)

    seg_df["YoY%"] = ((seg_df["Current"] - seg_df["Prev_Year"]) / seg_df["Prev_Year"].replace(0, pd.NA)) * 100
    seg_df["MoM_Change($M)"] = seg_df["Current"] - seg_df["Prev_Month"]

    # Pick top segment based on higher relative growth (YoY% or MoM change)
    seg_df["Relative_Score"] = (seg_df["YoY%"].fillna(0) + seg_df["MoM_Change($M)"].fillna(0))
    top_segment = seg_df["Relative_Score"].idxmax()
    top_segment_yoy = seg_df.loc[top_segment, "YoY%"]
    top_segment_mom = seg_df.loc[top_segment, "MoM_Change($M)"]

    # Product-level within top segment
    prod_current = current_data[current_data["CIB SME Segment"] == top_segment]
    prod_prev_year = prev_year_data[prev_year_data["CIB SME Segment"] == top_segment]
    prod_prev_month = prev_month_data[prev_month_data["CIB SME Segment"] == top_segment]

    prod_curr_agg = prod_current.groupby("Product")["Total Relationship Income ($M)"].sum()
    prod_prev_year_agg = prod_prev_year.groupby("Product")["Total Relationship Income ($M)"].sum()
    prod_prev_month_agg = prod_prev_month.groupby("Product")["Total Relationship Income ($M)"].sum()

    prod_df = pd.DataFrame({
        "Current": prod_curr_agg,
        "Prev_Year": prod_prev_year_agg,
        "Prev_Month": prod_prev_month_agg
    }).fillna(0)

    prod_df["YoY%"] = ((prod_df["Current"] - prod_df["Prev_Year"]) / prod_df["Prev_Year"].replace(0, pd.NA)) * 100
    prod_df["MoM_Change($M)"] = prod_df["Current"] - prod_df["Prev_Month"]
    prod_df["Relative_Score"] = prod_df["YoY%"].fillna(0) + prod_df["MoM_Change($M)"].fillna(0)

    top_products = prod_df.sort_values("Relative_Score", ascending=False).head(2).reset_index()

    # Create commentary text
    doc = Document()
    month_name = datetime(1900, current_month, 1).strftime("%B")

    para1 = (
        f"Managed TRI of ${current_tri:.2f}M in {month_name} {current_year}, "
        f"YoY change: {yoy_change:+.1f}%, MoM change: {mom_change:+.2f}M "
        f"vs {month_name} {prev_year} and {datetime(1900, prev_month, 1).strftime('%B')} {prev_month_year} respectively."
    )

    para2 = (
        f"Segments – Growth/Fall across all client segments. "
        f"Top-performing CIB SME segment: '{top_segment}' "
        f"with YoY change {top_segment_yoy:+.1f}% and MoM change {top_segment_mom:+.2f}M. "
        f"Top contributing products: "
        f"{top_products.loc[0, 'Product']} (${top_products.loc[0, 'Current']:.2f}M, "
        f"YoY {top_products.loc[0, 'YoY%']:+.1f}%, MoM {top_products.loc[0, 'MoM_Change($M)']:+.2f}M), and "
        f"{top_products.loc[1, 'Product']} (${top_products.loc[1, 'Current']:.2f}M, "
        f"YoY {top_products.loc[1, 'YoY%']:+.1f}%, MoM {top_products.loc[1, 'MoM_Change($M)']:+.2f}M)."
    )

    doc.add_paragraph(para1)
    doc.add_paragraph(para2)
    file_path = "TRI_Commentary.docx"
    doc.save(file_path)
    print(f"✅ Commentary generated and saved as {file_path}")

    return para1, para2, seg_df, prod_df


# Example:
# df = pd.read_excel("your_file.xlsx")
# generate_tri_commentary(df)



```
