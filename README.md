```
# ======================================================
# 5️⃣ RELATIONSHIP MAINTENANCE – SEGMENT → GiD FILL
# ======================================================

REL_SHEET = "Relationship Maintenance"

if REL_SHEET in tpl_wb.sheetnames:

    rel_ws = tpl_wb[REL_SHEET]

    # Normalize PT_LE dataframe once
    pt_le["Booking Country"] = pt_le["Booking Country"].astype(str).str.strip()
    pt_le["Segment"] = pt_le["Segment"].astype(str).str.strip()

    # Filter PT_LE for current Booking Country
    pt_le_country = pt_le[pt_le["Booking Country"] == country]

    if not pt_le_country.empty:

        # Build Segment → GiD map (first occurrence only)
        segment_gid_map = {}
        for _, r in pt_le_country.iterrows():
            seg = r["Segment"]
            gid = r["GiD"]
            if seg not in segment_gid_map and gid not in [None, ""]:
                segment_gid_map[seg] = gid

        # Scan Relationship Maintenance sheet
        for row in rel_ws.iter_rows(min_row=1):

            seg_cell = row[2]   # Column C (0=A,1=B,2=C)
            target_cell = row[4]  # Column E

            if not isinstance(seg_cell.value, str):
                continue

            segment = seg_cell.value.strip()

            if segment in segment_gid_map:
                # First occurrence only
                if target_cell.value in [None, ""]:
                    target_cell.value = segment_gid_map[segment]
