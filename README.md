```
from docx import Document
from docx.shared import RGBColor
import re

def write_commentary_to_word_colored(
    commentary_dict,
    output_file,
    country_names,
    bizline_names,
    product_names
):
    """
    Word formatting rules:
    - Side headers -> BOLD
    - Country names -> BOLD (Markets only, except 'ats')
    - Business Line names -> BOLD ONLY in Business Lines section
    - Products section:
        * If product = "<BizLine> - <Product>", bold ONLY <Product>
        * Do NOT bold BizLine prefix
    - Positive numbers & % -> GREEN
    - Negative numbers & % -> RED
    """

    # ---- exclusions ----
    country_names = [c for c in country_names if c.lower() != "ats"]

    country_names = sorted(country_names, key=len, reverse=True)
    bizline_names = sorted(bizline_names, key=len, reverse=True)
    product_names = sorted(product_names, key=len, reverse=True)

    num_pattern = re.compile(
        r"(\+[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)|"
        r"(\-[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)"
    )

    headers = ["Segments -", "Markets -", "Business Lines -", "Products -"]

    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

    for region, text in commentary_dict.items():
        doc.add_heading(region, level=2)
        current_section = None

        for line in text.split("\n"):
            p = doc.add_paragraph()

            # ---- section detection + header bold ----
            for h in headers:
                if line.startswith(h):
                    run = p.add_run(h)
                    run.bold = True
                    line = line[len(h):]
                    current_section = h.replace(" -", "")
                    break

            i = 0
            while i < len(line):

                # ---- country bolding (Markets only) ----
                if current_section == "Markets":
                    for c in country_names:
                        if line[i:i+len(c)] == c:
                            run = p.add_run(c)
                            run.bold = True
                            i += len(c)
                            break
                    else:
                        pass
                    if i < len(line) and line[i-1:i] == c:
                        continue

                # ---- business line bolding (Business Lines only) ----
                if current_section == "Business Lines":
                    for b in bizline_names:
                        if line[i:i+len(b)] == b:
                            run = p.add_run(b)
                            run.bold = True
                            i += len(b)
                            break
                    else:
                        pass
                    if i < len(line) and line[i-1:i] == b:
                        continue

                # ---- product bolding (Products only, suffix only) ----
                if current_section == "Products":
                    for pdt in product_names:
                        if line[i:i+len(pdt)] == pdt:
                            # check if product has bizline prefix
                            matched_prefix = None
                            for b in bizline_names:
                                prefix = f"{b} - "
                                if pdt.startswith(prefix):
                                    matched_prefix = prefix
                                    break

                            if matched_prefix:
                                # write prefix as normal
                                p.add_run(matched_prefix)
                                # bold only suffix
                                suffix = pdt[len(matched_prefix):]
                                run = p.add_run(suffix)
                                run.bold = True
                            else:
                                run = p.add_run(pdt)
                                run.bold = True

                            i += len(pdt)
                            break
                    else:
                        pass
                    if i < len(line) and line[i-1:i] == pdt:
                        continue

                # ---- numbers coloring ----
                m = num_pattern.match(line, i)
                if m:
                    token = m.group()
                    run = p.add_run(token)
                    run.font.color.rgb = (
                        RGBColor(0, 176, 80)
                        if token.startswith("+")
                        else RGBColor(192, 0, 0)
                    )
                    i += len(token)
                    continue

                # ---- default ----
                p.add_run(line[i])
                i += 1

    doc.save(output_file)
    print(f"Word commentary written to {output_file}")
