```
import pandas as pd
import re

def validate_con(df, col="con"):
    # Remove spaces from column
    df[col] = df[col].astype(str).str.replace(" ", "", regex=False)
    
    # Create comments column
    df["comments"] = ""

    # Regex for valid pattern: 3 digits - 6 digits
    pattern = re.compile(r"^\d{3}-\d{6}$")

    def check_value(val):
        if pattern.match(val):
            return ""  # valid, no comment
        # check for specific issues
        if "-" not in val:
            return "Missing hyphen"
        before, after = val.split("-", 1)
        if not before.isdigit() or len(before) != 3:
            return "Invalid: before hyphen not 3 digits"
        if not after.isdigit() or len(after) != 6:
            return "Invalid: after hyphen not 6 digits"
        return "Invalid format"

    df["comments"] = df[col].apply(check_value)
    return df


# Example usage
data = {
    "con": ["123-456789", "12-456789", "123-45678", "123456789", "abc-123456"]
}
df = pd.DataFrame(data)

result = validate_con(df)
print(result)
```
