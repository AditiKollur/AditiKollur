```
from docx import Document
from docx.shared import RGBColor
import re

def write_commentary_to_word_colored(commentary_dict, output_file):
    """
    Writes region-wise commentary to a Word file.
    - Positive numbers AND percentages -> Green
    - Negative numbers AND percentages -> Red
    """

    doc = Document()
    doc.add_heading("Total Relationship Income Commentary", level=1)

    # Match:
    # +120mn
    # +120mn / 8.4%
    # -45mn
    # -45mn / -6.2%
    pattern = re.compile(
        r"(\+[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)|"
        r"(\-[0-9]+(\.[0-9]+)?(mn|bn)(\s*/\s*-?[0-9]+(\.[0-9]+)?%)?)"
    )

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

                # Normal text before match
                if start > idx:
                    p.add_run(line[idx:start])

                token = line[start:end]
                run = p.add_run(token)

                # Color logic based on sign of first char
                if token.strip().startswith("+"):
                    run.font.color.rgb = RGBColor(0, 176, 80)   # Green
                else:
                    run.font.color.rgb = RGBColor(192, 0, 0)   # Red

                idx = end

            # Remaining text
            if idx < len(line):
                p.add_run(line[idx:])

    doc.save(output_file)
    print(f"Word commentary written to {output_file}")
