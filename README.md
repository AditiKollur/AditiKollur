```
import pandas as pd
import numpy as np
from typing import Tuple, Dict

# ============================================================
# Helper Functions (updated: skip-zero logic + unit formatting)
# ============================================================

def _detect_unit(metric_col: str):
    """
    Detect unit from metric_col name.
    Returns: 'M' if millions, 'K' if thousands, else None (assume raw units).
    """
    if metric_col is None:
        return None
    text = metric_col.upper()
    if "$M" in text or " M)" in text or "M)" in text:
        return "M"
    if "$K" in text or " K)" in text or "K)" in text:
        return "K"
    return None


def _to_mn(value: float, unit: str):
    """
    Convert raw value to 'millions' for internal consistent display:
    - If unit == 'M' (values already in millions) -> return value (mn)
    - If unit == 'K' (values in thousands) -> convert to millions: value / 1000
    - If unit is None -> treat as raw numbers (no conversion) and still treat as mn-esque
      (we'll display with mn/bn logic based on value magnitude)
    """
    if pd.isna(value):
        return value
    if unit == "K":
        return value / 1000.0
    # unit == 'M' or None -> interpret as already in millions for formatting
    return value


def _fmt_magnitude_mn_bn(mn_value: float) -> str:
    """
    Format a value expressed in millions to 'xxx mn' or 'x.y bn'
    - if abs(mn_value) >= 1000 -> display in bn with one decimal (value/1000) + 'bn'
    - else display in mn as integer if whole, else one decimal + 'mn'
    Keep sign.
    """
    if pd.isna(mn_value):
        return "N/A"
    sign = "+" if mn_value > 0 else "-" if mn_value < 0 else ""
    abs_mn = abs(mn_value)
    if abs_mn >= 1000:
        # show in billions with one decimal unless integer
        bn = abs_mn / 1000.0
        if float(bn).is_integer():
            s = f"{int(bn):,}bn"
        else:
            s = f"{bn:,.1f}bn"
    else:
        # show in millions
        if float(abs_mn).is_integer():
            s = f"{int(abs_mn):,}mn"
        else:
            s = f"{abs_mn:,.1f}mn"
    return f"{sign}{s}"


def _fmt_change_yoy_scaled(change: float, yoy: float, metric_col: str) -> str:
    """
    Format change (scaled to mn/bn) and YoY% string.
    """
    unit = _detect_unit(metric_col)
    mn = _to_mn(change, unit)
    ch_str = _fmt_magnitude_mn_bn(mn)
    if pd.isna(yoy):
        return f"{ch_str} / N/A"
    return f"{ch_str} / {yoy:.1f}%"


def _fmt_total_scaled(total: float, metric_col: str) -> str:
    """
    Format total (same scaling/units as change).
    """
    unit = _detect_unit(metric_col)
    mn = _to_mn(total, unit)
    return _fmt_magnitude_mn_bn(mn)


def _compute_aggregates(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                        group_cols, metric_col: str) -> pd.DataFrame:
    """
    Aggregate CY & PY and compute Change and YoY%. (unchanged)
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


def _row_is_skippable(r):
    """
    Returns True if a row should be skipped based on:
      - Change == 0 and YoY% == 0
      - Change == 0 and YoY% is NaN
    i.e. when there's no meaningful movement / not informative.
    """
    ch = r.get("Change", None)
    yoy = r.get("YoY%", None)
    # treat nan properly
    if pd.isna(ch):
        return True
    if ch == 0:
        if pd.isna(yoy):
            return True
        if yoy == 0:
            return True
    return False


def _filter_non_skippable(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a dataframe filtered to remove skippable rows.
    """
    if df is None or df.empty:
        return df
    mask = df.apply(lambda r: not _row_is_skippable(r), axis=1)
    return df[mask].copy()


def _join_items(rows, name_col, metric_col, metric_change_col="Change", metric_yoy_col="YoY%"):
    """
    Build joined string of items, excluding skippable rows.
    rows: DataFrame
    name_col: column to display as name
    metric_col: original metric column name (for formatting)
    """
    if rows is None or rows.empty:
        return ""
    # filter skippable rows first
    rows = _filter_non_skippable(rows)
    if rows.empty:
        return ""
    parts = []
    for _, r in rows.iterrows():
        name = r[name_col]
        ch = r[metric_change_col]
        yoy = r[metric_yoy_col]
        parts.append(f"{name} ({_fmt_change_yoy_scaled(ch, yoy, metric_col)})")
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return " and ".join(parts)
    return ", ".join(parts[:-1]) + ", and " + parts[-1]


# ============================================================
# 1 — SUMMARY LINE (updated scaling)
# ============================================================

def summary_line(df_cy: pd.DataFrame, df_py: pd.DataFrame,
                 region_col: str, region_value: str,
                 metric_col: str) -> Tuple[str, int]:
    """
    Build summary line and sign. Total is scaled/formatted according to metric unit.
    """
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    total_cy = cy_r[metric_col].sum()
    total_py = py_r[metric_col].sum()

    change = total_cy - total_py
    yoy = (change / total_py * 100) if total_py != 0 else np.nan

    total_str = _fmt_total_scaled(total_cy, metric_col)
    change_yoy_str = _fmt_change_yoy_scaled(change, yoy, metric_col)

    sign = 1 if change > 0 else -1 if change < 0 else 0

    sentence = f"Managed Total Relationship income of {region_value} of {total_str}, {change_yoy_str}"
    return sentence, sign


# ============================================================
# 2 — SEGMENTS & BUSINESS LINES DRILLDOWN (skip zeros)
# ============================================================

def drilldown_with_offsets(
    df_cy: pd.DataFrame, df_py: pd.DataFrame,
    region_col: str, region_value: str,
    group_col: str, metric_col: str,
    n: int = 2,
    label_name: str = None
):
    label = label_name or group_col

    # region sign
    _, region_sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    # filter region
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    agg = _compute_aggregates(cy_r, py_r, [group_col], metric_col)
    if agg.empty:
        return f"{label} - No data for {region_value}."

    # remove skippable rows before choosing top/bottom
    agg_filtered = _filter_non_skippable(agg)

    if agg_filtered.empty:
        return f"{label} - No meaningful movement for {region_value}."

    tops = agg_filtered.sort_values("Change", ascending=False).head(n)
    bottoms = agg_filtered.sort_values("Change", ascending=True).head(n)

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

    main_text = _join_items(main_set, group_col, metric_col)
    offset_text = _join_items(offset_set, group_col, metric_col)

    if not main_text:
        main_text = "N/A"
    if not offset_text:
        offset_text = "N/A"

    return (
        f"{label} - In the {region_value} region, {verb} {main_text}. "
        f"Partially offset by {offset_text}."
    )


# ============================================================
# 3 — MARKETS DRILLDOWN WITH INLINE BIZ LINE (skip zeros)
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
    # filter skippable
    country_agg_filtered = _filter_non_skippable(country_agg)
    if country_agg_filtered.empty:
        return f"Markets - No meaningful movement for {region_value}."

    top_countries = country_agg_filtered.sort_values("Change", ascending=False).head(top_n)
    bottom_countries = country_agg_filtered.sort_values("Change", ascending=True).head(top_n)

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
        biz_agg_filtered = _filter_non_skippable(biz_agg)

        # if no biz line has movement, just return country (if country row not skippable)
        country_part = f"{country} ({_fmt_change_yoy_scaled(row['Change'], row['YoY%'], metric_col)})"

        if biz_agg_filtered.empty:
            return country_part

        # pick top1 & bottom1 from biz_agg_filtered
        top1 = biz_agg_filtered.sort_values("Change", ascending=False).head(1)
        bottom1 = biz_agg_filtered.sort_values("Change", ascending=True).head(1)

        # choose biz part depending on country change sign
        if row["Change"] > 0:
            biz_part = _join_items(top1, biz_col, metric_col)
        elif row["Change"] < 0:
            biz_part = _join_items(bottom1, biz_col, metric_col)
        else:
            biz_part = _join_items(top1, biz_col, metric_col)

        if biz_part:
            return f"{country_part} driven by {biz_part}"
        return country_part

    main_parts = [describe_country_with_biz(r) for _, r in main_set.iterrows()]
    offset_parts = [describe_country_with_biz(r) for _, r in offset_set.iterrows()]

    # remove any empty (in case some country rows got filtered out)
    main_parts = [p for p in main_parts if p]
    offset_parts = [p for p in offset_parts if p]

    if not main_parts:
        main_text = "N/A"
    else:
        main_text = " and ".join(main_parts) if len(main_parts) == 2 else ", ".join(main_parts)

    if not offset_parts:
        offset_text = "N/A"
    else:
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
                "summary": lines[0] if len(lines) > 0 else "",
                "segments": lines[1] if len(lines) > 1 else "",
                "markets": lines[2] if len(lines) > 2 else "",
                "business_lines": lines[3] if len(lines) > 3 else "",
                "full_commentary": comm
            })
        return pd.DataFrame(rows)

    return out


# ============================================================
# Example / quick test (small demo)
# ============================================================
if __name__ == "__main__":
    data_cy = [
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France", "Business Line":"Loans", "TRI":5000},
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France", "Business Line":"FX", "TRI":200},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Germany", "Business Line":"Deposits", "TRI":800},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Italy", "Business Line":"Cash", "TRI":0},
        {"Managed Region":"Asia",   "CIB Segment":"Corporate", "Managed country":"India", "Business Line":"Loans", "TRI":600000},
        {"Managed Region":"Asia",   "CIB Segment":"SME", "Managed country":"China", "Business Line":"Deposits", "TRI":1000},
    ]
    data_py = [
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France", "Business Line":"Loans", "TRI":4300},
        {"Managed Region":"Europe", "CIB Segment":"Corporate", "Managed country":"France", "Business Line":"FX", "TRI":180},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Germany", "Business Line":"Deposits", "TRI":820},
        {"Managed Region":"Europe", "CIB Segment":"SME", "Managed country":"Italy", "Business Line":"Cash", "TRI":0},
        {"Managed Region":"Asia",   "CIB Segment":"Corporate", "Managed country":"India", "Business Line":"Loans", "TRI":650000},
        {"Managed Region":"Asia",   "CIB Segment":"SME", "Managed country":"China", "Business Line":"Deposits", "TRI":1200},
    ]

    # Assume metric_col is in $K or $M depending on your data; for this demo, we'll label as $M
    df_cy = pd.DataFrame(data_cy).rename(columns={"TRI":"Total Relationship Income ($M)"})
    df_py = pd.DataFrame(data_py).rename(columns={"TRI":"Total Relationship Income ($M)"})

    out = all_regions_commentary(
        df_cy, df_py,
        region_col="Managed Region",
        metric_col="Total Relationship Income ($M)",
        segment_col="CIB Segment",
        market_col="Managed country",
        biz_col="Business Line",
        top_n=2,
        return_type="dict"
    )

    for region, text in out.items():
        print("=== REGION:", region, "===\n")
        print(text)
        print("\n")
