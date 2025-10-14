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

    # -------------------------
    # Generate commentary text (rounded to 0 decimals)
    # -------------------------
    para1 = (
