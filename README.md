```
from docx import Document
from docx.shared import RGBColor
import re

def write_commentary_to_word_colored(
    commentary_dict,
    output_file,
    country_names
):
    """
    Word output rules:
    - Side headers -> BOLD
    - Country names -> BOLD (EXCEPT 'ats')
    - Positive numbers & % -> GREEN
    - Negative numbers & % -> RED
    """

    # -------- REMOVE ats FROM COUNTRY BOLDING --------
    country_names = [
        c for c in country_names
        if c.lower() != "ats"
    ]

    # Sort longest first to avoid partial matches
    country_names = sorted(country_names, key=len, reverse=True)

    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

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

                # ---- Country bolding (except ats) ----
                matched_country = None
                for c in country_names:
                    if line[i:i+len(c)] == c:
                        matched_country = c
                        break

                if matched_country:
                    run = p.add_run(matched_country)
                    run.bold = True
                    i += len(matched_country)
                    continue

                # ---- Numbers + percentages coloring ----
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

