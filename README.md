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
# 1 — SUMMARY LINE
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
# 2 — SEGMENTS & BUSINESS LINES DRILLDOWN (same as before)
# ============================================================

def drilldown_with_offsets(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    group_col: str, metric_col: str,
    n: int = 2,
    label_name: str = None
):
    label = label_name or group_col

    _, region_sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    agg = _compute_aggregates(cy_r, py_r, [group_col], metric_col)
    if agg.empty:
        return f"{label} - No data for {region_value}."

    tops = agg.sort_values("Change", ascending=False).head(n)
    bottoms = agg.sort_values("Change", ascending=True).head(n)

    if region_sign > 0:
        verb = "growth was led by"
        main_set = tops
        offset_set = bottoms
    elif region_sign < 0:
        verb = "contraction was led by"
        main_set = bottoms
        offset_set = tops
    else:
        verb = "movement was led by"
        main_set = tops
        offset_set = bottoms

    main_text = _join_items(main_set, group_col)
    offset_text = _join_items(offset_set, group_col)

    return (
        f"{label} - In the {region_value} region, {verb} {main_text}. "
        f"Partially offset by {offset_text}."
    )


# ============================================================
# 3 — **MARKETS DRILLDOWN WITH INLINE BIZ LINE**
# ============================================================

def markets_with_biz_inline(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    country_col: str, biz_col: str,
    metric_col: str,
    top_n: int = 2
):
    _, region_sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    country_agg = _compute_aggregates(cy_r, py_r, [country_col], metric_col)
    if country_agg.empty:
        return f"Markets - No data for {region_value}."

    top_countries = country_agg.sort_values("Change", ascending=False).head(top_n)
    bottom_countries = country_agg.sort_values("Change", ascending=True).head(top_n)

    if region_sign > 0:
        verb = "growth was led by"
        main_set = top_countries
        offset_set = bottom_countries
    elif region_sign < 0:
        verb = "contraction was led by"
        main_set = bottom_countries
        offset_set = top_countries
    else:
        verb = "movement was led by"
        main_set = top_countries
        offset_set = bottom_countries

    def describe_country_with_biz(row):
        country = row[country_col]

        cy_c = cy_r[cy_r[country_col] == country]
        py_c = py_r[py_r[country_col] == country]

        biz_agg = _compute_aggregates(cy_c, py_c, [biz_col], metric_col)

        if biz_agg.empty:
            return f"{country} ({_fmt_change_yoy(row['Change'], row['YoY%'])})"

        top1 = biz_agg.sort_values("Change", ascending=False).head(1)
        bottom1 = biz_agg.sort_values("Change", ascending=True).head(1)

        if row["Change"] > 0:
            biz_part = _join_items(top1, biz_col)
        else:
            biz_part = _join_items(bottom1, biz_col)

        return f"{country} ({_fmt_change_yoy(row['Change'], row['YoY%'])}) driven by {biz_part}"

    main_parts = [describe_country_with_biz(r) for _, r in main_set.iterrows()]
    offset_parts = [describe_country_with_biz(r) for _, r in offset_set.iterrows()]

    main_text = " and ".join(main_parts) if len(main_parts) == 2 else ", ".join(main_parts)
    offset_text = " and ".join(offset_parts) if len(offset_parts) == 2 else ", ".join(offset_parts)

    return (
        f"Markets - In the {region_value} region, {verb} {main_text}, "
        f"partially offset by {offset_text}."
    )


# ============================================================
# 4 — REGION COMMENTARY (FINAL)
# ============================================================

def region_commentary(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    metric_col: str,
    segment_col: str = "CIB Segment",
    market_col: str = "Managed country",
    biz_col: str = "Business Line",
    top_n: int = 2
):
    summary, _ = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    segments_line = drilldown_with_offsets(
        df_cy, df_py, region_col, region_value,
        segment_col, metric_col, n=top_n, label_name="Segments"
    )

    markets_line = markets_with_biz_inline(
        df_cy, df_py, region_col, region_value,
        country_col=market_col,
        biz_col=biz_col,
        metric_col=metric_col,
        top_n=top_n
    )

    business_line = drilldown_with_offsets(
        df_cy, df_py, region_col, region_value,
        biz_col, metric_col, n=top_n, label_name="Business Lines"
    )

    return "\n".join([summary, segments_line, markets_line, business_line])


# ============================================================
# 5 — ALL REGIONS COMMENTARY
# ============================================================

def all_regions_commentary(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str,
    metric_col: str,
    segment_col: str = "CIB Segment",
    market_col: str = "Managed country",
    biz_col: str = "Business Line",
    top_n: int = 2,
    return_type: str = "dict"
):
    regions = (
        pd.Index(df_cy[region_col].dropna().unique())
        .union(pd.Index(df_py[region_col].dropna().unique()))
    )

    out = {}
    for region in regions:
        out[region] = region_commentary(
            df_cy, df_py,
            region_col, region,
            metric_col,
            segment_col=segment_col,
            market_col=market_col,
            biz_col=biz_col,
            top_n=top_n
        )

    if return_type == "df":
        rows = []
        for region, comm in out.items():
            lines = comm.split("\n")
            rows.append({
                region_col: region,
                "summary": lines[0],
                "segments": lines[1],
                "markets": lines[2],
                "business_lines": lines[3],
                "full_commentary": comm
            })
        return pd.DataFrame(rows)

    return out


output = all_regions_commentary(
    df_cy, df_py,
    region_col="Managed Region",
    metric_col="Total Relationship Income ($M)",
    segment_col="CIB Segment",
    market_col="Managed country",
    biz_col="Business Line",
    top_n=2
)

for region, text in output.items():
    print("=== REGION:", region, "===\n")
    print(text)
    print("\n")
