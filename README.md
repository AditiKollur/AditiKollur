```
import re

EXCEPTIONS = {"ab", "aps", "ats", "amg"}

def proper_case_except(text, exceptions=EXCEPTIONS):
    if not isinstance(text, str):
        return text

    # Step 1: normalize everything to lowercase
    text = text.lower()

    # Step 2: convert to Proper Case (title case)
    text = text.title()

    # Step 3: force exceptions back to lowercase using word boundaries
    for exc in exceptions:
        # replace whole word only (case-insensitive safe)
        text = re.sub(
            rf"\b{exc}\b",
            exc,
            text,
            flags=re.IGNORECASE
        )

    return text
