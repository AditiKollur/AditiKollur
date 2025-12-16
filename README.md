```
from docx import Document
from docx.shared import RGBColor
import re

def write_commentary_to_word_colored(
    commentary_dict,
    output_file,
    country_names,
    bizline_names
):
    """
    Word formatting rules:
    - Side headers -> BOLD
    - Country names -> BOLD (except 'ats')
    - Business line names -> BOLD
    - Positive numbers & % -> GREEN
    - Negative numbers & % -> RED
    """

    # ---- Exclusions ----
    country_names = [c for c in country_names if c.lower() != "ats"]

    # Sort longest first to avoid partial matches
    country_names = sorted(country_names, key=len, reverse=True)
    bizline_names = sorted(bizline_names, key=len, reverse=True)

    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

    # Numbers with optional YoY %
    num_pattern = re.compile(
        r"(\+[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)|"
        r"(\-[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)"
    )

    headers = ["Segments -", "Markets -", "Business Lines -", "Products -"]

    for region, text in commentary_dict.items():
        doc.add_heading(region, level=2)

        for line in text.split("\n"):
            p = doc.add_paragraph()

            # ---- Bold side headers ----
            for h in headers:
                if line.startswith(h):
                    r = p.add_run(h)
                    r.bold = True
                    line = line[len(h):]
                    break

            i = 0
            while i < len(line):

                # ---- Country bolding ----
                matched = None
                for c in country_names:
                    if line[i:i+len(c)] == c:
                        matched = ("country", c)
                        break

                # ---- Business line bolding ----
                if not matched:
                    for b in bizline_names:
                        if line[i:i+len(b)] == b:
                            matched = ("biz", b)
                            break

                if matched:
                    run = p.add_run(matched[1])
                    run.bold = True
                    i += len(matched[1])
                    continue

                # ---- Numbers coloring ----
                m = num_pattern.match(line, i)
                if m:
                    token = m.group()
                    run = p.add_run(token)
                    run.font.color.rgb = (
                        RGBColor(0, 176, 80) if token.startswith("+")
                        else RGBColor(192, 0, 0)
                    )
                    i += len(token)
                    continue

                # ---- Normal text ----
                p.add_run(line[i])
                i += 1

    doc.save(output_file)
    print(f"Word commentary written to {output_file}")


country_list = sorted(
    set(df_cy["Managed country"].dropna().unique())
    | set(df_py["Managed country"].dropna().unique())
)

bizline_list = sorted(
    set(df_cy["Business Line"].dropna().unique())
    | set(df_py["Business Line"].dropna().unique())
)

write_commentary_to_word_colored(
    commentary_dict=commentary,
    output_file="TRI_Commentary_Final.docx",
    country_names=country_list,
    bizline_names=bizline_list
)



