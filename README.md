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
    - Side headers (Segments -, Markets -, Business Lines -, Products -) -> BOLD
    - Country names -> BOLD
    - Positive numbers & % -> GREEN
    - Negative numbers & % -> RED
    """

    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

    # Match numbers WITH optional YoY %
    num_pattern = re.compile(
        r"(\+[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)|"
        r"(\-[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)"
    )

    headers = ["Segments -", "Markets -", "Business Lines -", "Products -"]

    # Sort country names by length (important to avoid partial matches)
    country_names = sorted(country_names, key=len, reverse=True)

    for region, text in commentary_dict.items():
        doc.add_heading(region, level=2)

        for line in text.split("\n"):
            p = doc.add_paragraph()
            idx = 0

            # ---- Bold side headers ----
            for h in headers:
                if line.startswith(h):
                    r = p.add_run(h)
                    r.bold = True
                    line = line[len(h):]
                    break

            # Find all numeric tokens
            matches = list(num_pattern.finditer(line))

            # Helper: find country match at a given index
            def match_country(s, start):
                for c in country_names:
                    if s[start:start+len(c)] == c:
                        return c
                return None

            i = 0
            while i < len(line):
                # --- Country names (bold) ---
                ctry = match_country(line, i)
                if ctry:
                    run = p.add_run(ctry)
                    run.bold = True
                    i += len(ctry)
                    continue

                # --- Numbers (colored) ---
                m = next((m for m in matches if m.start() == i), None)
                if m:
                    token = m.group()
                    run = p.add_run(token)
                    run.font.color.rgb = (
                        RGBColor(0, 176, 80) if token.startswith("+")
                        else RGBColor(192, 0, 0)
                    )
                    i = m.end()
                    continue

                # --- Normal text ---
                p.add_run(line[i])
                i += 1

    doc.save(output_file)
    print(f"Word commentary written to {output_file}")

