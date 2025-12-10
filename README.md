```
import pandas as pd
import numpy as np
from typing import Tuple, Dict

# ============================================================
# Helper Functions
# ============================================================

def _compute_aggregates(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                        group_cols, metric_col: str) -> pd.DataFrame:
    """
    Aggregates CY & PY, computes Change and YoY%.
    """
    cy = (
        df_cy.groupby(group_cols)[[metric_col]]
        .sum()
        .reset_index()
        .rename(columns={metric_col: f"{metric_col}_cy"})
    )

    py = (
        df_py.groupby(group_cols)[[metric_col]]
        .sum()
        .reset_index()
        .rename(columns={metric_col: f"{metric_col}_py"})
    )

    merged = cy.merge(py, how="left", on=group_cols).fillna(0)
    merged["Change"] = merged[f"{metric_col}_cy"] - merged[f"{metric_col}_py"]
    merged["YoY%"] = np.where(
        merged[f"{metric_col}_py"] != 0,
        merged["Change"] / merged[f"{metric_col}_py"] * 100,
        np.nan,
    )
    return merged


def _fmt_money(x: float) -> str:
    """
    Formats money with sign and comma separators.
    """
    if pd.isna(x):
        return "N/A"
    sign = "+" if x > 0 else "-" if x < 0 else ""
    val = abs(x)
    if float(val).is_integer():
        return f"{sign}{int(val):,}"
    return f"{sign}{val:,.1f}"


def _fmt_change_yoy(change: float, yoy: float) -> str:
    """
    Returns string in: +1,234 / 11.2%
    """
    ch = _fmt_money(change)
    if pd.isna(yoy):
        return f"{ch} / N/A"
    return f"{ch} / {yoy:.1f}%"


# ============================================================
# 1 — SUMMARY LINE function
# ============================================================

def summary_line(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                 region_col: str, region_value: str,
                 metric_col: str) -> Tuple[str, int]:
    """
    Builds Line 1:
    "Managed Total Relationship income of Europe of 1,234, +123 / 12.3%"

    Returns:
      - sentence string
      - sign of change (1 growth, -1 contraction, 0 flat)
    """

    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    total_cy = cy_r[metric_col].sum()
    total_py = py_r[metric_col].sum()

    change = total_cy - total_py
    yoy = (change / total_py * 100) if total_py != 0 else np.nan

    total_str = f"{int(total_cy):,}" if float(total_cy).is_integer() else f"{total_cy:,.1f}"
    change_yoy_str = _fmt_change_yoy(change, yoy)

    sign = 1 if change > 0 else -1 if change < 0 else 0

    sentence = (
        f"Managed Total Relationship income of {region_value} of {total_str}, "
        f"{change_yoy_str}"
    )

    return sentence, sign


# ============================================================
# 2 — DRILLDOWN LINE (Segments OR Markets)
# ============================================================

def drilldown_top_or_bottom(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    group_col: str, metric_col: str,
    overall_sign: int, n: int = 2,
    label_name: str = None
) -> str:
    """
    Builds Lines 2 & 3.
    Selects:
        - Top n if growth
        - Bottom n if contraction
    """

    label = label_name or group_col

    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    agg = _compute_aggregates(cy_r, py_r, [group_col], metric_col)

    if agg.empty:
        return f"{label} - No data for {region_value}."

    if overall_sign > 0:
        sel = agg.sort_values("Change", ascending=False).head(n)
        verb = "growth was led by"
    elif overall_sign < 0:
        sel = agg.sort_values("Change", ascending=True).head(n)
        verb = "contraction was led by"
    else:
        sel = agg.sort_values("Change", ascending=False).head(n)
        verb = "movement was led by"

    parts = []
    for _, r in sel.iterrows():
        parts.append(f"{r[group_col]} ({_fmt_change_yoy(r['Change'], r['YoY%'])})")

    if len(parts) == 1:
        joined = parts[0]
    elif len(parts) == 2:
        joined = " and ".join(parts)
    else:
        joined = ", ".join(parts[:-1]) + ", and " + parts[-1]

    return f"{label} - In the {region_value} region, {verb} {joined}"


# ============================================================
# 3 — FULL COMMENTARY FOR ONE REGION
# ============================================================

def region_commentary(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    metric_col: str,
    segment_col: str = "CIB Segment",
    market_col: str = "Managed country",
    top_n: int = 2
) -> str:

    summary, sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    segments_line = drilldown_top_or_bottom(
        df_cy, df_py, region_col, region_value,
        segment_col, metric_col, overall_sign=sign, n=top_n,
        label_name="Segments"
    )

    markets_line = drilldown_top_or_bottom(
        df_cy, df_py, region_col, region_value,
        market_col, metric_col, overall_sign=sign, n=top_n,
        label_name="Markets"
    )

    return "\n".join([summary, segments_line, markets_line])


# ============================================================
# 4 — FULL COMMENTARY FOR ALL REGIONS
# ============================================================

def all_regions_commentary(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str,
    metric_col: str,
    segment_col: str = "CIB Segment",
    market_col: str = "Managed country",
    top_n: int = 2,
    return_type: str = "dict"
):
    """
    Computes commentary for ALL region values in the region column.
    """

    regions = (
        pd.Index(df_cy[region_col].dropna().unique())
        .union(df_py[region_col].dropna().unique())
    )

    out = {}

    for region in regions:
        out[region] = region_commentary(
            df_cy, df_py,
            region_col, region,
            metric_col,
            segment_col=segment_col,
            market_col=market_col,
            top_n=top_n
        )

    if return_type == "df":
        rows = []
        for region, comm in out.items():
            lines = comm.split("\n")
            rows.append({
                region_col: region,
                "summary": lines[0],
                "segments": lines[1] if len(lines) > 1 else "",
                "markets": lines[2] if len(lines) > 2 else "",
                "full_commentary": comm
            })
        return pd.DataFrame(rows)

    return out
