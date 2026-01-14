```
# ======================================================
# ADD-ON: PRODUCT_CIR – TM1 PRODUCT BLOCK FILL (DEBUGGED)
# ======================================================

def norm(x):
    return " ".join(str(x).strip().lower().split()) if x not in [None, ""] else None


PROD_SHEET = "Product_CIR"

if PROD_SHEET in tpl_wb.sheetnames:

    prod_tpl_ws = tpl_wb[PROD_SHEET]

    # ---------- Field Mapper → Prod_CIR ----------
    prod_map_ws = fm_wb["Prod_CIR"]
    prod_headers = {c.value: i for i, c in enumerate(prod_map_ws[1])}

    prod_mappings = []
    for r in prod_map_ws.iter_rows(min_row=2, values_only=True):
        if r[prod_headers["temp_col"]] and r[prod_headers["Cost_walk"]]:
            prod_mappings.append({
                "temp_col": norm(r[prod_headers["temp_col"]]),
                "cw_anchor": norm(r[prod_headers["Cost_walk"]])
            })

    print(f"\n[DEBUG] Prod_CIR mappings loaded: {prod_mappings}")

    # ---------- Template: row 2 headers ----------
    tpl_col_index = {}
    for cell in prod_tpl_ws[2]:
        if isinstance(cell.value, str):
            tpl_col_index[norm(cell.value)] = cell.column

    print(f"[DEBUG] Product_CIR row-2 headers found: {list(tpl_col_index.keys())}")

    # ---------- Template: Product names in Column B ----------
    product_row_index = {}
    for r in range(3, prod_tpl_ws.max_row + 1):
        v = prod_tpl_ws.cell(row=r, column=2).value
        if isinstance(v, str):
            product_row_index[norm(v)] = r

    print(f"[DEBUG] Products found in Product_CIR column B: {list(product_row_index.keys())}")

    # ---------- Cost Walk TM1 headers (row 26) ----------
    cw_headers = []
    for col in range(1, cw_ws.max_column + 1):
        v = cw_ws.cell(row=26, column=col).value
        if isinstance(v, str):
            cw_headers.append((norm(v), col))

    cw_header_positions = {h: c for h, c in cw_headers}

    print(f"[DEBUG] TM1 row-26 headers: {[h for h, _ in cw_headers]}")

    # ---------- Cost Walk row for Booking Country ----------
    cw_rows = country_rows.get(country, [])
    if not cw_rows:
        print(f"[DEBUG] No Cost Walk rows for country: {country}")
    else:
        cw_row = cw_rows[0]

        # ---------- Apply mappings ----------
        for idx, m in enumerate(prod_mappings):

            temp_col = m["temp_col"]
            anchor = m["cw_anchor"]

            if temp_col not in tpl_col_index:
                print(f"[DEBUG] Template column not found for temp_col: {temp_col}")
                continue

            if anchor not in cw_header_positions:
                print(f"[DEBUG] Cost Walk anchor not found: {anchor}")
                continue

            tpl_col = tpl_col_index[temp_col]
            start_col = cw_header_positions[anchor] + 1

            # Determine end column (next anchor or end)
            if idx + 1 < len(prod_mappings):
                next_anchor = prod_mappings[idx + 1]["cw_anchor"]
                end_col = cw_header_positions.get(next_anchor, cw_ws.max_column + 1)
            else:
                end_col = cw_ws.max_column + 1

            print(
                f"[DEBUG] Filling for anchor '{anchor}' → template col '{temp_col}', "
                f"TM1 cols {start_col} to {end_col - 1}"
            )

            for col in range(start_col, end_col):
                prod_name_raw = cw_ws.cell(row=26, column=col).value
                prod_name = norm(prod_name_raw)

                if prod_name not in product_row_index:
                    continue

                value = cw_ws.cell(row=cw_row, column=col).value
                if value in [None, ""]:
                    continue

                row_idx = product_row_index[prod_name]
                target = prod_tpl_ws.cell(row=row_idx, column=tpl_col)

                if target.value in [None, ""]:
                    target.value = value
                    print(
                        f"[DEBUG] WROTE value {value} "
                        f"for product '{prod_name_raw}' "
                        f"at row {row_idx}, col {tpl_col}"
                    )
