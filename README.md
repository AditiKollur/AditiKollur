```
import pandas as pd
import numpy as np
from typing import Tuple, Dict

# ============================================================
# UNIT DETECTION & SCALING HELPERS
# ============================================================

def _detect_unit(metric_col: str):
    text = metric_col.upper()
    if "$M" in text:
        return "M"
    if "$K" in text:
        return "K"
    return None

def _to_mn(value: float, unit: str):
    if pd.isna(value): return value
    if unit == "K": return value / 1000.0
    return value  # unit M or None assumed already mn-like for formatting

def _fmt_magnitude_mn_bn(mn_value: float):
    if pd.isna(mn_value): return "N/A"
    sign = "+" if mn_value > 0 else "-" if mn_value < 0 else ""
    abs_val = abs(mn_value)
    if abs_val >= 1000:
        bn = abs_val / 1000.0
        s = f"{int(bn):,}bn" if bn.is_integer() else f"{bn:,.1f}bn"
    else:
        s = f"{int(abs_val):,}mn" if abs_val.is_integer() else f"{abs_val:,.1f}mn"
    return f"{sign}{s}"

def _fmt_change_yoy_scaled(change: float, yoy: float, metric_col: str):
    unit = _detect_unit(metric_col)
    mn = _to_mn(change, unit)
    ch_str = _fmt_magnitude_mn_bn(mn)
    if pd.isna(yoy): return f"{ch_str} / N/A"
    return f"{ch_str} / {yoy:.1f}%"

def _fmt_total_scaled(total: float, metric_col: str):
    unit = _detect_unit(metric_col)
    mn = _to_mn(total, unit)
    return _fmt_magnitude_mn_bn(mn)


# ============================================================
# AGGREGATION HELPERS
# ============================================================

def _compute_aggregates(df_cy, df_py, group_cols, metric_col):
    cy = df_cy.groupby(group_cols)[[metric_col]].sum().reset_index()
    cy.rename(columns={metric_col: f"{metric_col}_cy"}, inplace=True)

    py = df_py.groupby(group_cols)[[metric_col]].sum().reset_index()
    py.rename(columns={metric_col: f"{metric_col}_py"}, inplace=True)

    merged = cy.merge(py, on=group_cols, how="left").fillna(0)
    merged["Change"] = merged[f"{metric_col}_cy"] - merged[f"{metric_col}_py"]
    merged["YoY%"] = np.where(
        merged[f"{metric_col}_py"] != 0,
        merged["Change"] / merged[f"{metric_col}_py"] * 100,
        np.nan,
    )
    return merged

def _row_is_skippable(r):
    ch, yoy = r["Change"], r["YoY%"]
    if ch == 0 and (pd.isna(yoy) or yoy == 0): return True
    return False

def _filter_non_skippable(df):
    if df.empty: return df
    mask = df.apply(lambda r: not _row_is_skippable(r), axis=1)
    return df[mask].copy()

def _join_items(rows, name_col, metric_col):
    rows = _filter_non_skippable(rows)
    if rows.empty: return ""
    parts = []
    for _, r in rows.iterrows():
        parts.append(
            f"{r[name_col]} ({_fmt_change_yoy_scaled(r['Change'], r['YoY%'], metric_col)})"
        )
    return " and ".join(parts) if len(parts) == 2 else ", ".join(parts)


# ============================================================
# SUMMARY LINE
# ============================================================

def summary_line(df_cy, df_py, region_col, region_value, metric_col):
    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    total_cy = cy_r[metric_col].sum()
    total_py = py_r[metric_col].sum()

    change = total_cy - total_py
    yoy = (change / total_py * 100) if total_py != 0 else np.nan

    summary = (
        f"Managed Total Relationship income of {region_value} of "
        f"{_fmt_total_scaled(total_cy, metric_col)}, "
        f"{_fmt_change_yoy_scaled(change, yoy, metric_col)}"
    )

    sign = 1 if change > 0 else -1 if change < 0 else 0

    return summary, sign


# ============================================================
# SEGMENTS + BUSINESS LINE DRILLDOWN WITH NEGATIVE-BOTTOM RULE
# ============================================================

def drilldown_with_offsets(df_cy, df_py, region_col, region_value,
                           group_col, metric_col, n=2, label_name=None):

    label = label_name or group_col
    summary, region_sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    agg = _filter_non_skippable(_compute_aggregates(cy_r, py_r, [group_col], metric_col))
    if agg.empty:
        return f"{label} - No meaningful movement in {region_value}."

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

    main_text = _join_items(main_set, group_col, metric_col)
    offset_text = _join_items(offset_set, group_col, metric_col)

    # NEW RULE: offset only if any bottom contributor has negative change
    any_negative = (offset_set["Change"] < 0).any()

    if any_negative and offset_text:
        return (
            f"{label} - In the {region_value} region, {verb} {main_text}. "
            f"Partially offset by {offset_text}."
        )
    else:
        return f"{label} - In the {region_value} region, {verb} {main_text}."


# ============================================================
# MARKETS DRILLDOWN WITH INLINE BIZLINE + NEGATIVE RULE
# ============================================================

def markets_with_biz_inline(df_cy, df_py, region_col, region_value,
                            country_col, biz_col, metric_col, top_n=2):

    summary, region_sign = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    cy_r = df_cy[df_cy[region_col] == region_value]
    py_r = df_py[df_py[region_col] == region_value]

    country_agg = _filter_non_skippable(
        _compute_aggregates(cy_r, py_r, [country_col], metric_col)
    )

    if country_agg.empty:
        return f"Markets - No meaningful movement in {region_value}."

    tops = country_agg.sort_values("Change", ascending=False).head(top_n)
    bottoms = country_agg.sort_values("Change", ascending=True).head(top_n)

    # determine verb direction
    if region_sign > 0:
        verb = "growth was led by"
        main_set, offset_set = tops, bottoms
    elif region_sign < 0:
        verb = "contraction was led by"
        main_set, offset_set = bottoms, tops
    else:
        verb = "movement was led by"
        main_set, offset_set = tops, bottoms

    def describe_country(row):
        country = row[country_col]
        country_txt = f"{country} ({_fmt_change_yoy_scaled(row['Change'], row['YoY%'], metric_col)})"

        # now find bizline top/bottom within country
        cy_c = cy_r[cy_r[country_col] == country]
        py_c = py_r[py_r[country_col] == country]

        biz_agg = _filter_non_skippable(
            _compute_aggregates(cy_c, py_c, [biz_col], metric_col)
        )

        if biz_agg.empty:
            return country_txt

        # choose biz contributor depending on country sign
        if row["Change"] > 0:
            chosen = biz_agg.sort_values("Change", ascending=False).head(1)
        else:
            chosen = biz_agg.sort_values("Change", ascending=True).head(1)

        biz_part = _join_items(chosen, biz_col, metric_col)
        return f"{country_txt} driven by {biz_part}"

    # build main side
    main_parts = [describe_country(r) for _, r in main_set.iterrows()]
    main_text = " and ".join(main_parts) if len(main_parts) == 2 else ", ".join(main_parts)

    # build offset side
    offset_parts = [describe_country(r) for _, r in offset_set.iterrows()]
    offset_text = " and ".join(offset_parts) if len(offset_parts) == 2 else ", ".join(offset_parts)

    # NEW RULE OPTION 1:
    # offset shown if ANY bottom country OR chosen biz contributor is negative
    bottom_country_negative = (offset_set["Change"] < 0).any()

    # bizline negative logic:
    biz_negative = False
    for _, r in offset_set.iterrows():
        country = r[country_col]
        cy_c = cy_r[cy_r[country_col] == country]
        py_c = py_r[py_r[country_col] == country]
        biz_agg = _filter_non_skippable(_compute_aggregates(cy_c, py_c, [biz_col], metric_col))
        if biz_agg.empty:
            continue
        if r["Change"] > 0:
            chosen = biz_agg.sort_values("Change", ascending=False).head(1)
        else:
            chosen = biz_agg.sort_values("Change", ascending=True).head(1)
        if (chosen["Change"] < 0).any():
            biz_negative = True

    show_offset = bottom_country_negative or biz_negative

    if show_offset and offset_text:
        return (
            f"Markets - In the {region_value} region, {verb} {main_text}, "
            f"partially offset by {offset_text}."
        )
    else:
        return f"Markets - In the {region_value} region, {verb} {main_text}."


# ============================================================
# REGION COMMENTARY
# ============================================================

def region_commentary(df_cy, df_py, region_col, region_value, metric_col,
                      segment_col="CIB Segment", market_col="Managed country",
                      biz_col="Business Line", top_n=2):

    summary, _ = summary_line(df_cy, df_py, region_col, region_value, metric_col)

    segments = drilldown_with_offsets(
        df_cy, df_py, region_col, region_value, segment_col, metric_col, top_n, "Segments"
    )

    markets = markets_with_biz_inline(
        df_cy, df_py, region_col, region_value, market_col, biz_col, metric_col, top_n
    )

    bizlines = drilldown_with_offsets(
        df_cy, df_py, region_col, region_value, biz_col, metric_col, top_n, "Business Lines"
    )

    return "\n".join([summary, segments, markets, bizlines])


# ============================================================
# ALL REGIONS COMMENTARY
# ============================================================

def all_regions_commentary(df_cy, df_py, region_col, metric_col,
                           segment_col="CIB Segment", market_col="Managed country",
                           biz_col="Business Line", top_n=2,
                           return_type="dict"):

    regions = pd.Index(df_cy[region_col].dropna().unique()).union(
              pd.Index(df_py[region_col].dropna().unique()))

    out = {}
    for region in regions:
        out[region] = region_commentary(
            df_cy, df_py, region_col, region, metric_col,
            segment_col, market_col, biz_col, top_n
        )

    if return_type == "df":
        df_out = []
        for region, txt in out.items():
            lines = txt.split("\n")
            df_out.append({
                region_col: region,
                "summary": lines[0],
                "segments": lines[1],
                "markets": lines[2],
                "business_lines": lines[3],
                "full": txt
            })
        return pd.DataFrame(df_out)

    return out
