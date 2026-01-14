```
# ======================================================
# ADD-ON: PRODUCT_CIR – PRODUCT VALUES FROM COST WALK TM1
# ======================================================

PROD_SHEET = "Product_CIR"

if PROD_SHEET in tpl_wb.sheetnames:

    prod_tpl_ws = tpl_wb[PROD_SHEET]

    # ---------- 1. Read Field Mapper → Prod_CIR ----------
    prod_map_ws = fm_wb["Prod_CIR"]
    prod_headers = {c.value: i for i, c in enumerate(prod_map_ws[1])}

    prod_mappings = []
    for r in prod_map_ws.iter_rows(min_row=2, values_only=True):
        if r[prod_headers["temp_col"]] and r[prod_headers["Cost_walk"]]:
            prod_mappings.append({
                "temp_col": str(r[prod_headers["temp_col"]]).strip(),
                "cw_anchor": str(r[prod_headers["Cost_walk"]]).strip()
            })

    # ---------- 2. Product_CIR: temp_col → column index (ROW 2) ----------
    tpl_col_index = {}
    for cell in prod_tpl_ws[2]:
        if isinstance(cell.value, str):
            key = cell.value.strip()
            if key not in tpl_col_index:
                tpl_col_index[key] = cell.column

    # ---------- 3. Product_CIR: Product name → row index (COLUMN B) ----------
    product_row_index = {}
    for r in range(3, prod_tpl_ws.max_row + 1):
        val = prod_tpl_ws.cell(row=r, column=2).value  # Column B
        if isinstance(val, str):
            key = val.strip()
            if key not in product_row_index:
                product_row_index[key] = r

    # ---------- 4. Cost Walk TM1 headers (ROW 26) ----------
    cw_header_index = {}
    for col in range(1, cw_ws.max_column + 1):
        v = cw_ws.cell(row=26, column=col).value
        if isinstance(v, str):
            cw_header_index[v.strip()] = col

    # ---------- 5. Cost Walk row for Booking Country ----------
    cw_rows = country_rows.get(country, [])
    if not cw_rows:
        pass
    else:
        cw_row = cw_rows[0]  # first occurrence only

        # ---------- 6. Apply mappings ----------
        for m in prod_mappings:
            temp_col = m["temp_col"]
            cw_anchor = m["cw_anchor"]

            if temp_col not in tpl_col_index:
                continue
            if cw_anchor not in cw_header_index:
                continue

            tpl_col = tpl_col_index[temp_col]
            cw_start_col = cw_header_index[cw_anchor] + 1

            # Scan product columns AFTER anchor
            for col in range(cw_start_col, cw_ws.max_column + 1):
                product_name = cw_ws.cell(row=26, column=col).value
                if not isinstance(product_name, str):
                    continue

                product_name = product_name.strip()
                if product_name not in product_row_index:
                    continue

                value = cw_ws.cell(row=cw_row, column=col).value
                if value in [None, ""]:
                    continue

                row_idx = product_row_index[product_name]
                target_cell = prod_tpl_ws.cell(row=row_idx, column=tpl_col)

                # First occurrence only
                if target_cell.value in [None, ""]:
                    target_cell.value = value
