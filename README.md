```
import pandas as pd
import numpy as np
import re
from docx import Document
from docx.shared import RGBColor


# ============================================================
# UNIT + FORMATTING HELPERS
# ============================================================

def detect_unit(metric_col):
    t = metric_col.upper()
    if "$K" in t:
        return "K"
    if "$M" in t:
        return "M"
    return None

def to_mn(value, unit):
    if pd.isna(value):
        return value
    return value / 1000 if unit == "K" else value

def fmt_mn_bn(mn):
    if pd.isna(mn):
        return "N/A"
    sign = "+" if mn > 0 else "-" if mn < 0 else ""
    mn = abs(mn)
    if mn >= 1000:
        return f"{sign}{mn/1000:.1f}bn"
    return f"{sign}{mn:.0f}mn"

def fmt_change_yoy(change, yoy, metric_col):
    unit = detect_unit(metric_col)
    mn = to_mn(change, unit)
    ch = fmt_mn_bn(mn)
    if pd.isna(yoy):
        return f"{ch} / N/A"
    return f"{ch} / {yoy:.1f}%"


# ============================================================
# AGGREGATION + FILTERING
# ============================================================

def compute_agg(df_cy, df_py, group_cols, metric_col):
    cy = df_cy.groupby(group_cols)[metric_col].sum().reset_index(name="CY")
    py = df_py.groupby(group_cols)[metric_col].sum().reset_index(name="PY")
    m = cy.merge(py, on=group_cols, how="left").fillna(0)
    m["Change"] = m["CY"] - m["PY"]
    m["YoY%"] = np.where(m["PY"] != 0, m["Change"] / m["PY"] * 100, np.nan)
    return m

def drop_noise(df):
    return df[
        ~((df["Change"] == 0) & ((df["YoY%"] == 0) | df["YoY%"].isna()))
    ]


# ============================================================
# 90% COVERAGE SELECTION
# ============================================================

def select_by_coverage(df, metric_col, pct=0.9, min_mn=5):
    if df.empty:
        return df

    unit = detect_unit(metric_col)
    df = df.copy()
    df["_abs_mn"] = df["Change"].abs().apply(lambda x: to_mn(x, unit))
    df = df[df["_abs_mn"] >= min_mn]

    if df.empty:
        return df

    total = df["_abs_mn"].sum()
    cutoff = total * pct

    rows, run = [], 0
    for _, r in df.iterrows():
        rows.append(r)
        run += r["_abs_mn"]
        if run >= cutoff:
            break

    return pd.DataFrame(rows).drop(columns="_abs_mn")

def join_items(df, name_col, metric_col):
    return " and ".join(
        f"{r[name_col]} ({fmt_change_yoy(r['Change'], r['YoY%'], metric_col)})"
        for _, r in df.iterrows()
    )


# ============================================================
# SUMMARY
# ============================================================

def summary_line(df_cy, df_py, region_col, region, metric_col):
    cy = df_cy[df_cy[region_col] == region][metric_col].sum()
    py = df_py[df_py[region_col] == region][metric_col].sum()
    change = cy - py
    yoy = (change / py * 100) if py != 0 else np.nan
    unit = detect_unit(metric_col)

    return (
        f"Managed Total Relationship income of {region} of "
        f"{fmt_mn_bn(to_mn(cy, unit))}, "
        f"{fmt_change_yoy(change, yoy, metric_col)}"
    )


# ============================================================
# SEGMENTS / BUSINESS LINES DRILLDOWN
# ============================================================

def drilldown_coverage(df_cy, df_py, region_col, region,
                       group_col, metric_col, label):

    agg = drop_noise(
        compute_agg(
            df_cy[df_cy[region_col] == region],
            df_py[df_py[region_col] == region],
            [group_col], metric_col
        )
    )

    pos = agg[agg["Change"] > 0].sort_values("Change", ascending=False)
    neg = agg[agg["Change"] < 0].sort_values("Change")

    top = select_by_coverage(pos, metric_col)
    bottom = select_by_coverage(neg, metric_col)

    text = (
        f"{label} - In the {region} region, growth was led by "
        f"{join_items(top, group_col, metric_col)}"
    )

    if not bottom.empty:
        text += f". Partially offset by {join_items(bottom, group_col, metric_col)}."

    return text


# ============================================================
# MARKETS WITH INLINE BUSINESS LINE
# ============================================================

def markets_with_biz(df_cy, df_py, region_col, region,
                     country_col, biz_col, metric_col):

    agg = drop_noise(
        compute_agg(
            df_cy[df_cy[region_col] == region],
            df_py[df_py[region_col] == region],
            [country_col], metric_col
        )
    )

    pos = agg[agg["Change"] > 0].sort_values("Change", ascending=False)
    neg = agg[agg["Change"] < 0].sort_values("Change")

    top = select_by_coverage(pos, metric_col)
    bottom = select_by_coverage(neg, metric_col)

    def describe_country(r):
        country = r[country_col]
        base = f"{country} ({fmt_change_yoy(r['Change'], r['YoY%'], metric_col)})"

        biz_agg = drop_noise(
            compute_agg(
                df_cy[(df_cy[region_col] == region) & (df_cy[country_col] == country)],
                df_py[(df_py[region_col] == region) & (df_py[country_col] == country)],
                [biz_col], metric_col
            )
        )

        if r["Change"] > 0:
            biz = select_by_coverage(
                biz_agg[biz_agg["Change"] > 0].sort_values("Change", ascending=False),
                metric_col
            )
        else:
            biz = select_by_coverage(
                biz_agg[biz_agg["Change"] < 0].sort_values("Change"),
                metric_col
            )

        return f"{base} driven by {join_items(biz, biz_col, metric_col)}" if not biz.empty else base

    main = " and ".join(describe_country(r) for _, r in top.iterrows())

    text = f"Markets - In the {region} region, growth was led by {main}"

    if not bottom.empty:
        offset = " and ".join(describe_country(r) for _, r in bottom.iterrows())
        text += f", partially offset by {offset}."

    return text


# ============================================================
# REGION & ALL-REGIONS COMMENTARY
# ============================================================

def region_commentary(df_cy, df_py, region_col, region, metric_col,
                      segment_col, country_col, biz_col):

    return "\n".join([
        summary_line(df_cy, df_py, region_col, region, metric_col),
        drilldown_coverage(df_cy, df_py, region_col, region, segment_col, metric_col, "Segments"),
        markets_with_biz(df_cy, df_py, region_col, region, country_col, biz_col, metric_col),
        drilldown_coverage(df_cy, df_py, region_col, region, biz_col, metric_col, "Business Lines")
    ])

def all_regions_commentary(df_cy, df_py, region_col, metric_col,
                           segment_col, country_col, biz_col):

    regions = (
        pd.Index(df_cy[region_col].dropna().unique())
        .union(pd.Index(df_py[region_col].dropna().unique()))
        .sort_values()
    )

    return {
        region: region_commentary(
            df_cy, df_py,
            region_col, region,
            metric_col,
            segment_col, country_col, biz_col
        )
        for region in regions
    }


# ============================================================
# WORD WRITER WITH COLOURS
# ============================================================

def write_commentary_to_word_colored(commentary_dict, output_file):
    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

    pattern = re.compile(r"\+[0-9]+(\.[0-9]+)?(mn|bn)|\-[0-9]+(\.[0-9]+)?(mn|bn)")

    for region, text in commentary_dict.items():
        doc.add_heading(region, level=2)

        for line in text.split("\n"):
            p = doc.add_paragraph()
            idx = 0
            matches = list(pattern.finditer(line))

            if not matches:
                p.add_run(line)
                continue

            for m in matches:
                start, end = m.span()

                if start > idx:
                    p.add_run(line[idx:start])

                run = p.add_run(line[start:end])

                if line[start] == "+":
                    run.font.color.rgb = RGBColor(0, 176, 80)   # Green
                else:
                    run.font.color.rgb = RGBColor(192, 0, 0)   # Red

                idx = end

            if idx < len(line):
                p.add_run(line[idx:])

    doc.save(output_file)
    print(f"Word commentary written to {output_file}")





commentary = all_regions_commentary(
    df_cy,
    df_py,
    region_col="Managed Region",
    metric_col="Total Relationship Income ($M)",
    segment_col="CIB Segment",
    country_col="Managed country",
    biz_col="Business Line"
)

write_commentary_to_word_colored(
    commentary,
    output_file="TRI_Commentary_Final.docx"
)
