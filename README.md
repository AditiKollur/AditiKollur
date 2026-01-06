```
import pandas as pd
import difflib
from openpyxl import load_workbook

# ================= CONFIG =================
FIELD_MAPPER_PATH = "Field_mapper.xlsx"
COST_WALK_PATH = "Cost_walk_summary.xlsx"
MASTER_TEMPLATE_PATH = "Master_Template.xlsx"

FIELD_FILL_MAP_SHEET = "Field fill map"
COST_WALK_SHEET = "TM1"
TEMPLATE_SHEET = "Template"

HEADER_ROW = 27
# =========================================


# ---------- NORMALIZATION ----------
def normalize(text):
    if text is None:
        return None
    return " ".join(
        str(text)
        .replace("\u00a0", " ")
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
        .lower()
        .split()
    )


# ---------- LOAD FIELD FILL MAP ----------
ff_wb = load_workbook(FIELD_MAPPER_PATH, data_only=True)
ff_ws = ff_wb[FIELD_FILL_MAP_SHEET]

ff_headers = {cell.value: idx + 1 for idx, cell in enumerate(ff_ws[1])}
tm1_col = ff_headers["TM1"]
field_col = ff_headers["Field name"]

field_map_rows = []
for row in ff_ws.iter_rows(min_row=2, values_only=True):
    field_map_rows.append({
        "tm1_raw": row[tm1_col - 1],
        "field_name_raw": row[field_col - 1]
    })


# ---------- LOAD COST WALK HEADERS ----------
cw_wb = load_workbook(COST_WALK_PATH, data_only=True)
cw_ws = cw_wb[COST_WALK_SHEET]

cost_headers_raw = []
for col in range(1, cw_ws.max_column + 1):
    val = cw_ws.cell(row=HEADER_ROW, column=col).value
    if val:
        cost_headers_raw.append(val)

cost_headers_norm = {normalize(h): h for h in cost_headers_raw}


# ---------- LOAD TEMPLATE CELL TEXT ----------
tpl_wb = load_workbook(MASTER_TEMPLATE_PATH, data_only=True)
tpl_ws = tpl_wb[TEMPLATE_SHEET]

template_text_raw = []
for row in tpl_ws.iter_rows():
    for cell in row:
        if isinstance(cell.value, str):
            template_text_raw.append(cell.value)

template_text_norm = {normalize(t): t for t in template_text_raw}


# ---------- BUILD DEBUG DATAFRAME ----------
records = []

for r in field_map_rows:
    tm1_raw = r["tm1_raw"]
    field_raw = r["field_name_raw"]

    tm1_norm = normalize(tm1_raw)
    field_norm = normalize(field_raw)

    tm1_found = tm1_norm in cost_headers_norm
    field_found = field_norm in template_text_norm

    tm1_similar = difflib.get_close_matches(
        tm1_norm or "",
        cost_headers_norm.keys(),
        n=3,
        cutoff=0.6
    )

    field_similar = difflib.get_close_matches(
        field_norm or "",
        template_text_norm.keys(),
        n=3,
        cutoff=0.6
    )

    if tm1_found and field_found:
        reason = "OK"
    elif not tm1_found and not field_found:
        reason = "TM1 header NOT found in Cost Walk AND Field name NOT found in Template"
    elif not tm1_found:
        reason = "TM1 header NOT found in Cost Walk (spacing / unicode / rename)"
    else:
        reason = "Field name NOT found in Template (textbox / merged cell / formatting)"

    records.append({
        "tm1_raw": tm1_raw,
        "tm1_normalized": tm1_norm,
        "found_in_cost_walk": tm1_found,
        "cost_walk_similar_headers": ", ".join(tm1_similar),
        "field_name_raw": field_raw,
        "field_name_normalized": field_norm,
        "found_in_template": field_found,
        "template_similar_texts": ", ".join(field_similar),
        "failure_reason": reason
    })


df_debug = pd.DataFrame(records)

# ---------- VIEW RESULTS ----------
print("\n=== ONLY FAILURES ===")
print(df_debug[df_debug["failure_reason"] != "OK"])

# Optional export
df_debug.to_excel("mapping_debug.xlsx", index=False)

print("\n✅ Debug file created: mapping_debug.xlsx")

