```
import re

EXCEPTIONS = {"ab", "aps", "ats", "amg"}

def proper_case_except(text, exceptions=EXCEPTIONS):
    if not isinstance(text, str):
        return text

    def fix_word(word):
        lw = word.lower()
        if lw in exceptions:
            return lw  # keep exactly lowercase
        return word.capitalize()

    # Split while keeping punctuation
    tokens = re.findall(r"[A-Za-z]+|[^A-Za-z]+", text)
    return "".join(fix_word(tok) if tok.isalpha() else tok for tok in tokens)



    df["commentary"] = df["commentary"].apply(proper_case_except)
