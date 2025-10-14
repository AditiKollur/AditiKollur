```
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from docx import Document

# --------------------------
# Function: Generate Commentary
# --------------------------

def generate_tri_commentary(df):
    # Convert Excel-style serial date to datetime
    df['Date'] = pd.to_datetime('1899-12-30') + pd.to_timedelta(df['Year-Month'], unit='D')
    df['Year'] = df['Date'].dt.year
    df['Month'] = df['Date'].dt.month
    df['MonthName'] = df['Date'].dt.strftime('%B')

    # Combine United Kingdom regions into Europe
    df['Managed region'] = df['Managed region'].replace(
        {'United Kingdom - RFB': 'Europe', 'United Kingdom - NRFB': 'Europe'}
    )

    # Detect current and previous year
    current_year = df['Year'].max()
    previous_year = current_year - 1

    # Detect latest month
    current_month = df.loc[df['Year'] == current_year, 'Month'].max()
    current_month_name = df.loc[df['Month'] == current_month, 'MonthName'].iloc[0]

    # Current and previous year-month data
    current_df = df[(df['Year'] == current_year) & (df['Month'] == current_month)]
    prev_df = df[(df['Year'] == previous_year) & (df['Month'] == current_month)]

    # ------------------------------
    # 1️⃣ Managed TRI (Overall)
    # ------------------------------
    total_current = current_df['Total Relationship Income ($M)'].sum()
    total_prev = prev_df['Total Relationship Income ($M)'].sum()
    yoy_change = total_current - total_prev
    yoy_percent = (yoy_change / total_prev) * 100 if total_prev != 0 else np.nan

    # ------------------------------
    # 2️⃣ Segment Commentary
    # ------------------------------
    seg_curr = current_df.groupby('CIB SME Segment')['Total Relationship Income ($M)'].sum().reset_index()
    seg_prev = prev_df.groupby('CIB SME Segment')['Total Relationship Income ($M)'].sum().reset_index()
    seg_merge = pd.merge(seg_curr, seg_prev, on='CIB SME Segment', suffixes=('_curr', '_prev'))
    seg_merge['YoY_change'] = seg_merge['Total Relationship Income ($M)_curr'] - seg_merge['Total Relationship Income ($M)_prev']
    seg_merge['YoY_%'] = (seg_merge['YoY_change'] / seg_merge['Total Relationship Income ($M)_prev']) * 100
    top_seg = seg_merge.sort_values('YoY_change', ascending=False).iloc[0]

    # Top 2 Business Lines contributing to that segment
    prod_curr = current_df[current_df['CIB SME Segment'] == top_seg['CIB SME Segment']]
    prod_prev = prev_df[prev_df['CIB SME Segment'] == top_seg['CIB SME Segment']]
    prod_curr_sum = prod_curr.groupby('Business Line')['Total Relationship Income ($M)'].sum().reset_index()
    prod_prev_sum = prod_prev.groupby('Business Line')['Total Relationship Income ($M)'].sum().reset_index()
    prod_merge = pd.merge(prod_curr_sum, prod_prev_sum, on='Business Line', suffixes=('_curr', '_prev'))
    prod_merge['YoY_change'] = prod_merge['Total Relationship Income ($M)_curr'] - prod_merge['Total Relationship Income ($M)_prev']
    prod_merge = prod_merge.sort_values('YoY_change', ascending=False).head(2)
    top_products = ', '.join(prod_merge['Business Line'].tolist())

    # ------------------------------
    # 3️⃣ Regional Commentary
    # ------------------------------
    reg_curr = current_df.groupby('Managed region')['Total Relationship Income ($M)'].sum().reset_index()
    reg_prev = prev_df.groupby('Managed region')['Total Relationship Income ($M)'].sum().reset_index()
    reg_merge = pd.merge(reg_curr, reg_prev, on='Managed region', suffixes=('_curr', '_prev'))
    reg_merge['YoY_change'] = reg_merge['Total Relationship Income ($M)_curr'] - reg_merge['Total Relationship Income ($M)_prev']
    reg_merge['YoY_%'] = (reg_merge['YoY_change'] / reg_merge['Total Relationship Income ($M)_prev']) * 100
    reg_merge = reg_merge.sort_values('YoY_change', ascending=False).reset_index(drop=True)

    # Get top 5 regions
    top_regions = reg_merge.head(5)

    def get_top_countries(region_name):
        # Get top 2 countries within the given region
        reg_countries_curr = current_df[current_df['Managed region'] == region_name].groupby('Managed Country')['Total Relationship Income ($M)'].sum().reset_index()
        reg_countries_prev = prev_df[prev_df['Managed region'] == region_name].groupby('Managed Country')['Total Relationship Income ($M)'].sum().reset_index()
        merge = pd.merge(reg_countries_curr, reg_countries_prev, on='Managed Country', suffixes=('_curr', '_prev'))
        merge['YoY_change'] = merge['Total Relationship Income ($M)_curr'] - merge['Total Relationship Income ($M)_prev']
        merge = merge.sort_values('YoY_change', ascending=False).head(2)
        return ', '.join(merge['Managed Country'].tolist())

    # Region details
    top1 = top_regions.iloc[0]
    top2 = top_regions.iloc[1]
    others = top_regions.iloc[2:5]

    top1_countries = get_top_countries(top1['Managed region'])
    top2_countries = get_top_countries(top2['Managed region'])

    # ------------------------------
    # Write to Word File
    # ------------------------------
    doc = Document()

    doc.add_heading(f"TRI Commentary - {current_month_name} {current_year}", level=1)

    # 1️⃣ Managed TRI Commentary
    doc.add_paragraph(
        f"1. Managed TRI of ${total_current:.2f}M in {current_month_name} {current_year}, "
        f"representing a change of ${yoy_change:.2f}M ({yoy_percent:.2f}%) from last year."
    )

    # 2️⃣ Segment Commentary
    doc.add_paragraph(
        f"2. Segments - Growth observed across all client segments, primarily in {top_seg['CIB SME Segment']} "
        f"(YoY change ${top_seg['YoY_change']:.2f}M, {top_seg['YoY_%']:.2f}%), "
        f"driven by {top_products}."
    )

    # 3️⃣ Regional Commentary
    reg_text = (
        f"3. Regions - Strong growth in {top1['Managed region']} "
        f"(YoY change ${top1['YoY_change']:.2f}M, {top1['YoY_%']:.2f}%) led by {top1_countries}, "
        f"and {top2['Managed region']} (YoY change ${top2['YoY_change']:.2f}M, {top2['YoY_%']:.2f}%) led by {top2_countries}, "
        f"accompanied by steady growth in "
    )
    steady_parts = [
        f"{row['Managed region']} (YoY change ${row['YoY_change']:.2f}M, {row['YoY_%']:.2f}%)"
        for _, row in others.iterrows()
    ]
    reg_text += ', '.join(steady_parts) + '.'

    doc.add_paragraph(reg_text)

    # Save Word file
    output_path = "tri_commentary.docx"
    doc.save(output_path)
    print(f"✅ Commentary file generated: {output_path}")


# --------------------------
# Example Usage
# --------------------------
# df = pd.read_excel("your_data.xlsx")
# generate_tri_commentary(df)
