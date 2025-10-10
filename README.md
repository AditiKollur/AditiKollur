```
import pandas as pd
from datetime import datetime, timedelta
from docx import Document

def excel_serial_to_date(serial):
    """
    Convert Excel serial date (e.g., 45504) to datetime.
    Excel's day 1 = 1899-12-30 (for Windows default date system).
    """
    base_date = datetime(1899, 12, 30)
    return base_date + timedelta(days=int(serial))

def generate_tri_commentary(df, output_path='TRI_Commentary.docx'):
    # Convert Excel serials to datetime
    df['_date'] = df['Year-Month'].apply(excel_serial_to_date)

    # Extract Year and Month
    df['Year'] = df['_date'].dt.year
    df['Month'] = df['_date'].dt.month

    # Determine current and previous year based on latest date
    curr_date = df['_date'].max()
    curr_year, curr_month = curr_date.year, curr_date.month
    prev_year = curr_year - 1

    # Define YTD filters
    curr_ytd = (df['Year'] == curr_year) & (df['Month'] <= curr_month)
    prev_ytd = (df['Year'] == prev_year) & (df['Month'] <= curr_month)

    # --- (1) Managed TRI Total + YoY change ---
    tri_col = 'Total Relationship Income ($M)'
    curr_tri = df.loc[curr_ytd, tri_col].sum()
    prev_tri = df.loc[prev_ytd, tri_col].sum()
    yoy = ((curr_tri - prev_tri) / prev_tri * 100) if prev_tri else None

    # --- (2) Segment-level growth/fall ---
    seg_curr = df.loc[curr_ytd].groupby('CIB SME Segment')[tri_col].sum()
    seg_prev = df.loc[prev_ytd].groupby('CIB SME Segment')[tri_col].sum()

    seg_df = pd.concat([seg_curr, seg_prev], axis=1, keys=['Current', 'Previous']).fillna(0)
    seg_df['YoY%'] = seg_df.apply(
        lambda r: ((r['Current'] - r['Previous']) / r['Previous'] * 100) if r['Previous'] else None, axis=1
    )
    seg_df['ChangeType'] = seg_df.apply(
        lambda r: 'Growth' if r['Current'] > r['Previous'] else 'Fall' if r['Current'] < r['Previous'] else 'No Change', axis=1
    )

    # Identify top segment
    seg_sorted = seg_df[seg_df['YoY%'].notna()].sort_values('YoY%', ascending=False)
    if not seg_sorted.empty:
        top_segment = seg_sorted.index[0]
        top_seg_yoy = seg_sorted.iloc[0]['YoY%']
    else:
        top_segment, top_seg_yoy = None, None

    # --- (3) Top 2 Products contributing to top segment ---
    if top_segment:
        curr_top_seg = df.loc[curr_ytd & (df['CIB SME Segment'] == top_segment)]
        prod_curr = curr_top_seg.groupby('Product')[tri_col].sum().sort_values(ascending=False)
        top_products = prod_curr.head(2).index.tolist()

        prod_prev = df.loc[prev_ytd & (df['CIB SME Segment'] == top_segment)]
        prod_prev = prod_prev.groupby('Product')[tri_col].sum()

        product_contrib = []
        for p in top_products:
            curr_val = prod_curr.get(p, 0)
            prev_val = prod_prev.get(p, 0)
            yoy_prod = ((curr_val - prev_val) / prev_val * 100) if prev_val else None
            product_contrib.append((p, curr_val, yoy_prod))
    else:
        product_contrib = []

    # --- Generate commentary text ---
    month_label = curr_date.strftime('%B %Y')

    # 1️⃣ Managed TRI
    para1 = (
        f"Managed TRI of ${curr_tri:,.2f}M in {month_label} YTD, "
        f"{'change from last year: N/A' if yoy is None else f'change from last year, {yoy:+.1f}% YoY.'}"
    )

    # 2️⃣ Segment summary
    seg_summary = "; ".join([
        f"{seg}: {row.ChangeType} ({'N/A' if pd.isna(row['YoY%']) else f'{row['YoY%']:+.1f}%'} YoY)"
        for seg, row in seg_df.iterrows()
    ])

    if top_segment:
        top_seg_text = (
            f"Primarily the '{top_segment}' segment showing {top_seg_yoy:+.1f}% YoY growth, "
            f"driven by products "
            f"{' and '.join([f'{p} ({y:+.1f}% YoY)' if y is not None else f'{p} (N/A)' for p,_,y in product_contrib])}."
        )
    else:
        top_seg_text = "No segment growth data available."

    para2 = f"Segments - Growth/Fall across all client segments: {seg_summary}. {top_seg_text}"

    # --- Write to Word file ---
    doc = Document()
    doc.add_heading(f"TRI Commentary — {month_label}", level=1)
    doc.add_paragraph(para1)
    doc.add_paragraph(para2)
    doc.save(output_path)

    print(f"✅ Commentary file created: {output_path}")
    print("\n--- Paragraph 1 ---\n", para1)
    print("\n--- Paragraph 2 ---\n", para2)


# Example usage:
# ------------------------------------------
data = [
    [45504,'Asia','India',False,'CBR1','BL1','ProductA','SME Large',12.5],
    [45504,'United States','USA',False,'CBR1','BL1','ProductB','SME Small',8.0],
    [45869,'Asia','India',False,'CBR1','BL1','ProductA','SME Large',15.0],
    [45869,'United States','USA',False,'CBR1','BL1','ProductB','SME Small',9.5],
    [45869,'Europe','UK',False,'CBR2','BL2','ProductD','SME Large',7.0],
    [45535,'Asia','India',False,'CBR1','BL1','ProductA','SME Large',10.0],
    [45535,'United States','USA',False,'CBR1','BL1','ProductB','SME Small',9.0],
    [45900,'Asia','India',False,'CBR1','BL1','ProductA','SME Large',11.0],
    [45900,'United States','USA',False,'CBR1','BL1','ProductB','SME Small',8.5],
    [45900,'Europe','UK',False,'CBR2','BL2','ProductD','SME Large',6.0],
]

df = pd.DataFrame(data, columns=[
    'Year-Month','Managed region','Managed Country','Multi Jurisdiction','CBR',
    'Business Line','Product','CIB SME Segment','Total Relationship Income ($M)'
])

generate_tri_commentary(df)
```
