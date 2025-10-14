```
import pandas as pd
import numpy as np
from datetime import datetime
from docx import Document

def generate_tri_commentary(df):
    # Convert Excel serial date to datetime
    df['Year-Month'] = pd.to_datetime(df['Year-Month'], unit='d', origin='1899-12-30')
    df['Year'] = df['Year-Month'].dt.year
    df['Month'] = df['Year-Month'].dt.month_name().str[:3]

    # Combine United Kingdom regions into Europe
    df['Managed Region'] = df['Managed Region'].replace({
        'United Kingdom - RFB': 'Europe',
        'United Kingdom - NRFB': 'Europe'
    })

    # Identify current and previous year from latest month
    current_year = df['Year'].max()
    current_month = df.loc[df['Year'] == current_year, 'Year-Month'].max().month
    prev_year = current_year - 1
    month_name = datetime(1900, current_month, 1).strftime('%B')

    # Filter data for current and previous years, current month
    df_curr = df[(df['Year'] == current_year) & (df['Year-Month'].dt.month == current_month)]
    df_prev = df[(df['Year'] == prev_year) & (df['Year-Month'].dt.month == current_month)]

    # --- 1️⃣ Managed TRI Commentary ---
    tri_curr = df_curr['Total Relationship Income ($M)'].sum()
    tri_prev = df_prev['Total Relationship Income ($M)'].sum()
    yoy_change = tri_curr - tri_prev
    yoy_pct = (yoy_change / tri_prev * 100) if tri_prev != 0 else np.nan

    # --- 2️⃣ Segment Commentary ---
    seg_curr = df_curr.groupby('CIB SME Segment')['Total Relationship Income ($M)'].sum().reset_index()
    seg_prev = df_prev.groupby('CIB SME Segment')['Total Relationship Income ($M)'].sum().reset_index()
    seg = pd.merge(seg_curr, seg_prev, on='CIB SME Segment', suffixes=('_curr', '_prev'), how='outer').fillna(0)
    seg['YoY_Change'] = seg['Total Relationship Income ($M)_curr'] - seg['Total Relationship Income ($M)_prev']
    seg['YoY%'] = seg.apply(
        lambda x: (x['YoY_Change'] / x['Total Relationship Income ($M)_prev'] * 100)
        if x['Total Relationship Income ($M)_prev'] != 0 else np.nan, axis=1
    )
    top_seg = seg.sort_values('YoY_Change', ascending=False).iloc[0]

    # Get top 2 Business Lines contributing to top segment
    seg_curr_top = df_curr[df_curr['CIB SME Segment'] == top_seg['CIB SME Segment']]
    bl = seg_curr_top.groupby('Business Line')['Total Relationship Income ($M)'].sum().nlargest(2).index.tolist()

    # --- 3️⃣ Region Commentary ---
    reg_curr = df_curr.groupby('Managed Region')['Total Relationship Income ($M)'].sum().reset_index()
    reg_prev = df_prev.groupby('Managed Region')['Total Relationship Income ($M)'].sum().reset_index()
    reg = pd.merge(reg_curr, reg_prev, on='Managed Region', suffixes=('_curr', '_prev'), how='outer').fillna(0)
    reg['YoY_Change'] = reg['Total Relationship Income ($M)_curr'] - reg['Total Relationship Income ($M)_prev']
    reg['YoY%'] = reg.apply(
        lambda x: (x['YoY_Change'] / x['Total Relationship Income ($M)_prev'] * 100)
        if x['Total Relationship Income ($M)_prev'] != 0 else np.nan, axis=1
    )
    reg = reg.sort_values('YoY%', ascending=False).reset_index(drop=True)

    # Get Top 5 Regions
    top_regions = reg.head(5)

    # For each region, get top 2 countries
    region_country_data = []
    for region in top_regions['Managed Region']:
        df_region_curr = df_curr[df_curr['Managed Region'] == region]
        df_region_prev = df_prev[df_prev['Managed Region'] == region]
        country_curr = df_region_curr.groupby('Managed Country')['Total Relationship Income ($M)'].sum().reset_index()
        country_prev = df_region_prev.groupby('Managed Country')['Total Relationship Income ($M)'].sum().reset_index()
        country = pd.merge(country_curr, country_prev, on='Managed Country', suffixes=('_curr', '_prev'), how='outer').fillna(0)
        country['YoY_Change'] = country['Total Relationship Income ($M)_curr'] - country['Total Relationship Income ($M)_prev']
        country = country.sort_values('YoY_Change', ascending=False).head(2)
        top_countries = ', '.join(country['Managed Country'])
        region_country_data.append((region, top_countries))

    # --- 4️⃣ Product Commentary ---
    bl_curr = df_curr.groupby('Business Line')['Total Relationship Income ($M)'].sum().reset_index()
    bl_prev = df_prev.groupby('Business Line')['Total Relationship Income ($M)'].sum().reset_index()
    bl_data = pd.merge(bl_curr, bl_prev, on='Business Line', suffixes=('_curr', '_prev'), how='outer').fillna(0)
    bl_data['YoY_Change'] = bl_data['Total Relationship Income ($M)_curr'] - bl_data['Total Relationship Income ($M)_prev']
    bl_data['YoY%'] = bl_data.apply(
        lambda x: (x['YoY_Change'] / x['Total Relationship Income ($M)_prev'] * 100)
        if x['Total Relationship Income ($M)_prev'] != 0 else np.nan, axis=1
    )
    bl_data = bl_data.sort_values('YoY_Change', ascending=False).reset_index(drop=True)

    # Top 2 and Bottom 1 Business Lines
    top1 = bl_data.iloc[0]
    top2 = bl_data.iloc[1] if len(bl_data) > 1 else None
    bottom = bl_data.iloc[-1]

    # Get top 2 products for each of these Business Lines
    def get_top_products(df_curr, df_prev, bl_name):
        df_bl_curr = df_curr[df_curr['Business Line'] == bl_name]
        df_bl_prev = df_prev[df_prev['Business Line'] == bl_name]
        prod_curr = df_bl_curr.groupby('Product')['Total Relationship Income ($M)'].sum().reset_index()
        prod_prev = df_bl_prev.groupby('Product')['Total Relationship Income ($M)'].sum().reset_index()
        prod = pd.merge(prod_curr, prod_prev, on='Product', suffixes=('_curr', '_prev'), how='outer').fillna(0)
        prod['YoY_Change'] = prod['Total Relationship Income ($M)_curr'] - prod['Total Relationship Income ($M)_prev']
        prod = prod.sort_values('YoY_Change', ascending=False).head(2)
        return ', '.join(prod['Product'].tolist())

    top1_prods = get_top_products(df_curr, df_prev, top1['Business Line'])
    top2_prods = get_top_products(df_curr, df_prev, top2['Business Line']) if top2 is not None else ''
    bottom_prods = get_top_products(df_curr, df_prev, bottom['Business Line'])

    # -------------------------
    # Generate commentary text (rounded to 0 decimals)
    # -------------------------
    para1 = (
        f"Managed TRI of ${tri_curr:.0f}M in {month_name} {current_year}, "
        f"a change of ${yoy_change:.0f}M ({yoy_pct:.0f}%) from last year."
    )

    para2 = (
        f"Segments – Growth observed across client segments, primarily in {top_seg['CIB SME Segment']} "
        f"(${top_seg['YoY_Change']:.0f}M, {top_seg['YoY%']:.0f}%) driven by "
        f"{bl[0]} and {bl[1]} business lines."
    )

    region_parts = []
    for i, (region, countries) in enumerate(region_country_data):
        row = top_regions.iloc[i]
        if i == 0:
            region_parts.append(f"Strong Growth in {region} (${row['YoY_Change']:.0f}M, {row['YoY%']:.0f}%) ({countries})")
        elif i == 1:
            region_parts.append(f"and {region} (${row['YoY_Change']:.0f}M, {row['YoY%']:.0f}%) ({countries})")
        else:
            region_parts.append(f"accompanied by steady growth in {region} (${row['YoY_Change']:.0f}M, {row['YoY%']:.0f}%)")
    para3 = "Regions – " + ", ".join(region_parts) + "."

    para4 = (
        f"Products – Strong {top1['Business Line']} (${top1['YoY_Change']:.0f}M, {top1['YoY%']:.0f}%) driven by "
        f"{top1_prods}. {top2['Business Line']} (${top2['YoY_Change']:.0f}M, {top2['YoY%']:.0f}%) driven by "
        f"{top2_prods}. Conversely, {bottom['Business Line']} was down "
        f"${abs(bottom['YoY_Change']):.0f}M ({abs(bottom['YoY%']):.0f}%) from external factors."
    )

    # -------------------------
    # Write to Word File
    # -------------------------
    doc = Document()
    doc.add_heading(f"TRI Commentary - {month_name} {current_year}", level=1)
    doc.add_paragraph(para1)
    doc.add_paragraph(para2)
    doc.add_paragraph(para3)
    doc.add_paragraph(para4)
    doc.save("TRI_Commentary.docx")

    print("✅ TRI Commentary Word file generated successfully: TRI_Commentary.docx")

# --------------------------
# Example usage
# --------------------------
# df = pd.read_excel("your_data.xlsx")  # Load your actual data
# generate_tri_commentary(df)
```
