```
from pptx import Presentation
import pandas as pd

def extract_tables_from_ppt(pptx_path):
    prs = Presentation(pptx_path)
    tables = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                data = []
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        row_data.append(cell.text)
                    data.append(row_data)
                df = pd.DataFrame(data)
                tables.append(df)
    return tables

def extract_charts_from_ppt(pptx_path):
    prs = Presentation(pptx_path)
    charts_data = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_chart:
                chart = shape.chart
                chart_info = {
                    "chart_type": chart.chart_type,
                    "plots": []
                }
                for plot in chart.plots:
                    categories = [cat.label for cat in plot.categories]
                    series_data = []
                    for series in plot.series:
                        series_data.append({
                            "name": series.name,
                            "values": list(series.values)
                        })
                    chart_info["plots"].append({
                        "categories": categories,
                        "series": series_data
                    })
                charts_data.append(chart_info)
    return charts_data

# Replace 'your_presentation.pptx' with your file path
pptx_file = "your_presentation.pptx"

# Extract and display tables
tables = extract_tables_from_ppt(pptx_file)
print("Extracted Tables:")
for i, table_df in enumerate(tables):
    print(f"Table {i+1}:")
    print(table_df)
    print()

# Extract and display charts
charts = extract_charts_from_ppt(pptx_file)
print("Extracted Charts Data:")
for i, chart in enumerate(charts):
    print(f"Chart {i+1}: Type - {chart['chart_type']}")
    for plot in chart["plots"]:
        print("Categories:", plot["categories"])
        for series in plot["series"]:
            print(f"Series '{series['name']}':", series["values"])
    print()
```
