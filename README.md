```
import pandas as pd
import numpy as np
from typing import Tuple, Dict

# ============================================================
# Helper Functions
# ============================================================

def _compute_aggregates(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                        group_cols, metric_col: str) -> pd.DataFrame:
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
    if pd.isna(x):
        return "N/A"
    sign = "+" if x > 0 else "-" if x < 0 else ""
    val = abs(x)
    if float(val).is_integer():
        return f"{sign}{int(val):,}"
    return f"{sign}{val:,.1f}"


def _fmt_change_yoy(change: float, yoy: float) -> str:
    ch = _fmt_money(change)
    if pd.isna(yoy):
        return f"{ch} / N/A"
    return f"{ch} / {yoy:.1f}%"


def _join_items(rows, name_col, metric_change_col="Change", metric_yoy_col="YoY%"):
    parts = []
    for _, r in rows.iterrows():
        parts.append(f"{r[name_col]} ({_fmt_change_yoy(r[metric_change_col], r[metric_yoy_col])})")
    if len(parts) == 0:
        return ""
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return " and ".join(parts)
    return ", ".join(parts[:-1]) + ", and " + parts[-1]


# ============================================================
# 1 — SUMMARY LINE function
# ============================================================

def summary_line(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                 region_col: str, region_value: str,
                 metric_col: str) -> Tuple[str, int]:
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    total_cy = cy_r[metric_col].sum()
    total_py = py_r[metric_col].sum()

    change = total_cy - total_py
    yoy = (change / total_py * 100) if total_py != 0 else np.nan

    total_str = f"{int(total_cy):,}" if float(total_cy).is_integer() else f"{total_cy:,.1f}"
    change_yoy_str = _fmt_change_yoy(change, yoy)
    sign = 1 if change > 0 else -1 if change < 0 else 0

    sentence = f"Managed Total Relationship income of {region_value} of {total_str}, {change_yoy_str}"
    return sentence, sign


# ============================================================
# 2 — DRILLDOWN LINE (Segments OR Markets) — UPDATED
# ============================================================

def drilldown_with_offsets(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    group_col: str, metric_col: str,
    n: int = 2,
    label_name: str = None
) -> str:
    """
    Builds a line that contains both:
      - leaders (top n by Change)
      - bottomers (bottom n by Change) presented as 'partially offset by'

    Wording adapts:
      - If overall region change > 0: "growth was led by <top n>. Partially offset by <bottom n>."
      - If overall region change < 0: "contraction was led by <bottom n>. Partially offset by <top n>."
      - If flat: "movement was led by <top n>. Partially offset by <bottom n>."
    """
    label = label_name or group_col

    # region totals sign to decide wording
    _, sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    # filter region
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    agg = _compute_aggregates(cy_r, py_r, [group_col], metric_col)
    if agg.empty:
        return f"{label} - No data for {region_value}."

    tops = agg.sort_values("Change", ascending=False).head(n)
    bots = agg.sort_values("Change", ascending=True).head(n)

    if sign > 0:
        verb_main = "growth was led by"
        main_part = _join_items(tops, group_col)
        offset_part = _join_items(bots, group_col)
    elif sign < 0:
        verb_main = "contraction was led by"
        main_part = _join_items(bots, group_col)
        offset_part = _join_items(tops, group_col)
    else:
        verb_main = "movement was led by"
        main_part = _join_items(tops, group_col)
        offset_part = _join_items(bots, group_col)

    main_str = f"{verb_main} {main_part}" if main_part else f"{verb_main} N/A"
    offset_str = f"Partially offset by {offset_part}" if offset_part else "Partially offset by N/A"

    return f"{label} - In the {region_value} region, {main_str}. {offset_str}."


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
    summary, _ = summary_line(df_cy, df_py, region_col, region_value, metric_col)
    segments_line = drilldown_with_offsets(
        df_cy, df_py, region_col, region_value, segment_col, metric_col, n=top_n, label_name="Segments"
    )
    markets_line = drilldown_with_offsets(
        df_cy, df_py, region_col, region_value, market_col, metric_col, n=top_n, label_name="Markets"
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
    regions = (
        pd.Index(df_cy[region_col].dropna().unique())
        .union(df_py[region_col].dropna().unique())
    )
    out = {}
    for region in regions:
        out[region] = region_commentary(
            df_cy, df_py, region_col, region, metric_col,
            segment_col=segment_col, market_col=market_col, top_n=top_n
        )
    if return_type == "df":
        rows = []
        for region, comm in out.items():
            lines = comm.split("\n")
            rows.append({
                region_col: region,
                "summary": lines[0] if len(lines) > 0 else "",
                "segments": lines[1] if len(lines) > 1 else "",
                "markets": lines[2] if len(lines) > 2 else "",
                "full_commentary": comm
            })
        return pd.DataFrame(rows)
    return out


# ============================================================
# Demo (run as script)
# ============================================================
if __name__ == "__main__":
    data_cy = [
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France", "Business Line":"Loans", "TRI":500},
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"Germany", "Business Line":"FX", "TRI":300},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Italy", "Business Line":"Deposits", "TRI":200},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Spain", "Business Line":"Cash", "TRI":150},
        {"Managed Region":"Asia",   "CIB Segment":"Corporate", "Managed country":"India", "Business Line":"Loans", "TRI":600},
        {"Managed Region":"Asia",   "CIB Segment":"SME", "Managed country":"China", "Business Line":"Deposits", "TRI":100},
    ]
    data_py = [
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France", "Business Line":"Loans", "TRI":430},
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"Germany", "Business Line":"FX", "TRI":280},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Italy", "Business Line":"Deposits", "TRI":210},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Spain", "Business Line":"Cash", "TRI":160},
        {"Managed Region":"Asia",   "CIB Segment":"Corporate", "Managed country":"India", "Business Line":"Loans", "TRI":650},
        {"Managed Region":"Asia",   "CIB Segment":"SME", "Managed country":"China", "Business Line":"Deposits", "TRI":120},
    ]

    df_cy = pd.DataFrame(data_cy).rename(columns={"TRI":"Total Relationship Income ($M)"})
    df_py = pd.DataFrame(data_py).rename(columns={"TRI":"Total Relationship Income ($M)"})

    all_comm = all_regions_commentary(
        df_cy, df_py,
        region_col="Managed Region",
        metric_col="Total Relationship Income ($M)",
        segment_col="CIB Segment",
        market_col="Managed country",
        top_n=2,
        return_type="dict"
    )

    for region, comm in all_comm.items():
        print("=== REGION:", region, "===\n")
        print(comm)
        print("\n")
