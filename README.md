```
import pandas as pd
import numpy as np
from typing import Tuple

def _compute_aggregates(df_cy: pd.DataFrame, df_py: pd.DataFrame, group_cols, metric_col: str) -> pd.DataFrame:
    """
    Helper: aggregate CY and PY by group_cols and compute Change and YoY%.
    """
    cy = df_cy.groupby(group_cols)[[metric_col]].sum().reset_index().rename(columns={metric_col: f"{metric_col}_cy"})
    py = df_py.groupby(group_cols)[[metric_col]].sum().reset_index().rename(columns={metric_col: f"{metric_col}_py"})
    merged = cy.merge(py, how="left", on=group_cols).fillna(0)
    merged["Change"] = merged[f"{metric_col}_cy"] - merged[f"{metric_col}_py"]
    # safe YoY%
    merged["YoY%"] = np.where(merged[f"{metric_col}_py"] != 0,
                               merged["Change"] / merged[f"{metric_col}_py"] * 100,
                               np.nan)
    return merged

def _fmt_money(x: float) -> str:
    """Format money with commas and no decimals if integer-ish, else one decimal. Keep sign."""
    if pd.isna(x):
        return "N/A"
    # include sign
    sign = "+" if x > 0 else "-" if x < 0 else ""
    val = abs(x)
    if float(val).is_integer():
        s = f"{int(val):,}"
    else:
        s = f"{val:,.1f}"
    return f"{sign}{s}"

def _fmt_change_yoy(change: float, yoy: float) -> str:
    """Return string like +1,234 / 12.3% (handle NaN YoY)"""
    ch = _fmt_money(change)
    if pd.isna(yoy):
        return f"{ch} / N/A"
    return f"{ch} / {yoy:.1f}%"

# ----------------------------
# Function 1: summary line
# ----------------------------
def summary_line(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                 region_col: str, region_value: str,
                 metric_col: str) -> Tuple[str, float]:
    """
    Build the first summary line for a specific region.

    Returns:
      - summary string
      - region_change_sign (positive if growth, negative if contraction, 0 if flat)
    Example:
      "Managed Total Relationship income of Europe of 1,234, +123 / 11.1%"
    """
    # filter region
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    # aggregate totals
    total_cy = cy_r[metric_col].sum()
    total_py = py_r[metric_col].sum()

    change = total_cy - total_py
    yoy = (change / total_py * 100) if total_py != 0 else np.nan

    # format numbers
    total_str = f"{total_cy:,.0f}" if float(total_cy).is_integer() else f"{total_cy:,.1f}"
    change_yoy_str = _fmt_change_yoy(change, yoy)

    sign = 1 if change > 0 else (-1 if change < 0 else 0)
    summary = (f"Managed Total Relationship income of {region_value} of {total_str}, "
               f"{change_yoy_str}")
    return summary, sign

# ----------------------------
# Function 2: drilldown (Segments/Markets)
# ----------------------------
def drilldown_top_or_bottom(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                            region_col: str, region_value: str,
                            group_col: str, metric_col: str,
                            overall_sign: int, n: int = 2,
                            label_name: str = None) -> str:
    """
    Build a line that lists top n (if overall_sign>0) OR bottom n (if overall_sign<0)
    grouped by group_col within the specified region. Returns a formatted line.

    Example output for group_col='CIB Segment' and label_name='Segments':
      "Segments - In the Europe region, growth was led by Corporate (+123 / 12.3%) and SMB (+45 / 5.6%)"

    Parameters:
      - overall_sign: 1 (growth) / -1 (contraction) / 0 (flat) — decides top or bottom selection
      - n: how many items to pick
      - label_name: human label like "Segments" or "Markets" (if None, uses group_col)
    """
    label = label_name or group_col
    # restrict to region
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    # aggregate by group_col
    agg = _compute_aggregates(cy_r, py_r, [group_col], metric_col)
    if agg.empty:
        return f"{label} - No data for {region_value}."

    # Choose top or bottom n
    if overall_sign > 0:
        # growth -> top n by Change
        sel = agg.sort_values("Change", ascending=False).head(n)
        verb = "growth was led by"
    elif overall_sign < 0:
        # contraction -> bottom n by Change (most negative)
        sel = agg.sort_values("Change", ascending=True).head(n)
        verb = "contraction was led by"
    else:
        # flat -> pick top n (neutral wording)
        sel = agg.sort_values("Change", ascending=False).head(n)
        verb = "movement was led by"

    # format each selected row
    parts = []
    for _, r in sel.iterrows():
        name = str(r[group_col])
        change = r["Change"]
        yoy = r["YoY%"]
        parts.append(f"{name} ({_fmt_change_yoy(change, yoy)})")

    # join with ' and '
    if len(parts) == 1:
        joined = parts[0]
    elif len(parts) == 2:
        joined = " and ".join(parts)
    else:
        joined = ", ".join(parts[:-1]) + ", and " + parts[-1]

    # final sentence
    sentence = f"{label} - In the {region_value} region, {verb} {joined}"
    return sentence

# ----------------------------
# Wrapper: produce full commentary (2 lines)
# ----------------------------
def region_commentary(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                      region_col: str, region_value: str,
                      metric_col: str,
                      segment_col: str = "CIB Segment",
                      market_col: str = "Managed country",
                      top_n: int = 2) -> str:
    """
    Produces the full commentary with:
      1) summary_line(...)
      2) segments line (using drilldown_top_or_bottom)
      3) markets line (using drilldown_top_or_bottom)

    Returns the complete commentary string (multi-line).
    """
    # first line
    summary, sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    # second lines
    segments_line = drilldown_top_or_bottom(df_cy, df_py, region_col, region_value,
                                            segment_col, metric_col, overall_sign=sign,
                                            n=top_n, label_name="Segments")
    markets_line = drilldown_top_or_bottom(df_cy, df_py, region_col, region_value,
                                           market_col, metric_col, overall_sign=sign,
                                           n=top_n, label_name="Markets")

    # combine
    return "\n".join([summary, segments_line, markets_line])

# ----------------------------
# Example usage / demo
# ----------------------------
if __name__ == "__main__":
    # sample data using the columns you mentioned
    data_cy = [
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France",
         "Business Line":"Loans", "MJ":1, "CBR":2, "TRI":500},
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"Germany",
         "Business Line":"FX", "MJ":1, "CBR":2, "TRI":300},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Italy",
         "Business Line":"Deposits", "MJ":1, "CBR":2, "TRI":200},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Spain",
         "Business Line":"Cash", "MJ":1, "CBR":2, "TRI":150},
    ]
    data_py = [
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France",
         "Business Line":"Loans", "MJ":1, "CBR":2, "TRI":430},
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"Germany",
         "Business Line":"FX", "MJ":1, "CBR":2, "TRI":280},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Italy",
         "Business Line":"Deposits", "MJ":1, "CBR":2, "TRI":210},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Spain",
         "Business Line":"Cash", "MJ":1, "CBR":2, "TRI":160},
    ]

    df_cy = pd.DataFrame(data_cy).rename(columns={"TRI":"Total Relationship Income ($M)"})
    df_py = pd.DataFrame(data_py).rename(columns={"TRI":"Total Relationship Income ($M)"})

    commentary = region_commentary(
        df_cy, df_py,
        region_col="Managed Region", region_value="Europe",
        metric_col="Total Relationship Income ($M)",
        segment_col="CIB Segment",
        market_col="Managed country",
        top_n=2
    )

    print(commentary)
