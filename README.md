```
import pandas as pd

def write_commentary_to_excel_colored(
    df_cy,
    df_py,
    region_col,
    metric_col,
    segment_col,
    country_col,
    biz_col,
    output_file="TRI_Commentary.xlsx"
):
    # Generate commentary
    commentary_dict = all_regions_commentary(
        df_cy=df_cy,
        df_py=df_py,
        region_col=region_col,
        metric_col=metric_col,
        segment_col=segment_col,
        country_col=country_col,
        biz_col=biz_col
    )

    rows = []
    for region, text in commentary_dict.items():
        lines = text.split("\n")
        rows.append({
            region_col: region,
            "Summary": lines[0] if len(lines) > 0 else "",
            "Segments": lines[1] if len(lines) > 1 else "",
            "Markets": lines[2] if len(lines) > 2 else "",
            "Business Lines": lines[3] if len(lines) > 3 else "",
            "Full Commentary": text
        })

    df_out = pd.DataFrame(rows)

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Commentary")

        workbook = writer.book
        worksheet = writer.sheets["Commentary"]

        # -------- FORMATS --------
        wrap = workbook.add_format({"text_wrap": True, "valign": "top"})
        green = workbook.add_format({"font_color": "green"})
        red = workbook.add_format({"font_color": "red"})

        # Column sizing + wrap
        worksheet.set_column("A:A", 18)
        worksheet.set_column("B:F", 80, wrap)

        last_row = len(df_out) + 1
        last_col_letter = "F"

        # -------- CONDITIONAL FORMATTING --------
        # Positive numbers: +xxxmn or +x.xbn
        worksheet.conditional_format(
            f"B2:{last_col_letter}{last_row}",
            {
                "type": "formula",
                "criteria": '=REGEXMATCH(B2,"\\+[0-9]+(\\.[0-9]+)?(mn|bn)")',
                "format": green,
            },
        )

        # Negative numbers: -xxxmn or -x.xbn
        worksheet.conditional_format(
            f"B2:{last_col_letter}{last_row}",
            {
                "type": "formula",
                "criteria": '=REGEXMATCH(B2,"\\-[0-9]+(\\.[0-9]+)?(mn|bn)")',
                "format": red,
            },
        )

    print(f"✅ Commentary written to {output_file} with color formatting")



