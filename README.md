```
import pandas as pd
from datetime import datetime, timedelta
from docx import Document

def excel_serial_to_date(serial):
    """Convert Excel serial date number to datetime"""
    return datetime(1899, 12, 30) + timedelta(days=int(serial))

def generate_tri_commentary_yoy(df):
    # Convert Excel serial to datetime
    df["Date"] = df["Year-Month"].apply(excel_serial_to_date)
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month

    # Identify current and previous years/months
    current_month = df["Month"].max()
    current_year = df.loc[df["Month"] == current_month, "Year"].max()
    prev_year = current_year - 1

    # Filter data for comparison
    current_data = df[(df["Month"] == current_month) & (df["Year"] == current_year)]
    prev_year_data = df[(df["Month"] == current_month) & (df["Year"] == prev_year)]

    # Managed TRI totals
    current_tri = current_data["Total Relationship Income ($M)"].sum()
    prev_year_tri = prev_year_data["Total Relationship Income ($M)"].sum()
    yoy_change = ((current_tri - prev_year_tri) / prev_year_tri * 100) if prev_year_tri != 0 else 0

    # Segment-level performance
    seg_current = current_data.groupby("CIB SME Segment")["Total Relationship Income ($M)"].sum()
    seg_prev_year = prev_year_data.groupby("CIB SME Segment")["Total Relationship Income ($M)"].sum()

    seg_df = pd.DataFrame({
        "Current": seg_current,
        "Prev_Year": seg_prev_year
    }).fillna(0)
    seg_df["YoY%"] = ((seg_df["Current"] - seg_df["Prev_Year"]) / seg_df["Prev_Year"].replace(0, pd.NA)) * 100

    # Pick top segment by YoY%
    top_segment = seg_df["YoY%"].idxmax()
    top_segment_yoy = seg_df.loc[top_segment, "YoY%"]

    # Product-level within top segment
    prod_current = current_data[current_data["CIB SME Segment"] == top_segment]
    prod_prev_year = prev_year_data[prev_year_data["CIB SME Segment"] == top_segment]

    prod_curr_agg = prod_current.groupby("Product")["Total Relationship Income ($M)"].sum()
    prod_prev_year_agg = prod_prev_year.groupby("Product")["Total Relationship Income ($M)"].sum()

    prod_df = pd.DataFrame({
        "Current": prod_curr_agg,
        "Prev_Year": prod_prev_year_agg
    }).fillna(0)
    prod_df["YoY%"] = ((prod_df["Current"] - prod_df["Prev_Year"]) / prod_df["Prev_Year"].replace(0, pd.NA)) * 100

    # Top 2 products for the top segment
    top_products = prod_df.sort_values("YoY%", ascending=False).head(2).reset_index()

    # --- Create commentary text ---
    doc = Document()
    month_name = datetime(1900, current_month, 1).strftime("%B")

    para1 = (
        f"Managed TRI of ${current_tri:.2f}M in {month_name} {current_year}, "
        f"YoY change: {yoy_change:+.1f}% vs {month_name} {prev_year}."
    )

    para2 = (
        f"Segments – Growth/Fall across all client segments. "
        f"Top-performing CIB SME segment: '{top_segment}' "
        f"with YoY change {top_segment_yoy:+.1f}%. "
        f"Top contributing products: "
        f"{top_products.loc[0, 'Product']} (${top_products.loc[0, 'Current']:.2f}M, YoY {top_products.loc[0, 'YoY%']:+.1f}%), "
        f"and {top_products.loc[1, 'Product']} (${top_products.loc[1, 'Current']:.2f}M, YoY {top_products.loc[1, 'YoY%']:+.1f}%)."
    )

    doc.add_paragraph(para1)
    doc.add_paragraph(para2)

    # Save Word file
    file_path = "TRI_Commentary_YoY.docx"
    doc.save(file_path)
    print(f"✅ Commentary generated and saved as {file_path}")

    return para1, para2, seg_df, prod_df


# Example Usage
# df = pd.read_excel("your_file.xlsx")
# generate_tri_commentary_yoy(df)
```
