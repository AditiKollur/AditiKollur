```
import pandas as pd
import numpy as np

def compare_mg_inscope_mastergroup_strict(
    mg_p,
    mg_c,
    cust_gbm_p,
    cust_cmb_p,
    cust_gbm_c,
    cust_cmb_c,
    inscope_col_mg,
    financial_col='Total Operating Income (HORIS YTD Financials)',
    mastergroup_col='Mastergroup name',
    out_file='mg_exceptions_mastergroup_strict.xlsx'
):
    """
    Compare mg_p and mg_c considering ONLY columns in inscope_col_mg.
    Primary key = inscope_col_mg[0]. Detect Additions / Deletions / Changes.
    Attach financial_col strictly by matching on mastergroup_col (must exist in mg and cust).
    Writes three sheets (Additions, Deletions, Changes) and returns dict of DataFrames.
    """
    # Defensive copies
    mg_p = mg_p.copy()
    mg_c = mg_c.copy()
    cust_gbm_p = cust_gbm_p.copy()
    cust_cmb_p = cust_cmb_p.copy()
    cust_gbm_c = cust_gbm_c.copy()
    cust_cmb_c = cust_cmb_c.copy()
    inscope = list(inscope_col_mg)

    if len(inscope) == 0:
        raise ValueError("inscope_col_mg must contain at least one column name.")

    primary_key = inscope[0]
    compare_cols = inscope[:]  # only these columns matter

    # helper: find actual column name in a list of dataframes using case-insensitive match
    def resolve_column(name, dfs):
        normalized = lambda s: "" if s is None else "".join(s.lower().split())
        target = normalized(name)
        for df in dfs:
            if df is None:
                continue
            for col in df.columns:
                if normalized(col) == target:
                    return col
        return None

    # Resolve primary_key presence exactly in mg_p and mg_c
    if primary_key not in mg_p.columns or primary_key not in mg_c.columns:
        # try fuzzy resolve primary key (rare but safe)
        resolved_pk = resolve_column(primary_key, [mg_p, mg_c])
        if resolved_pk is None:
            raise ValueError(f"Primary key '{primary_key}' must exist in both mg_p and mg_c (no fuzzy match found).")
        else:
            # if found, rename both mg frames to use the requested primary_key name for rest of logic
            if resolved_pk != primary_key:
                mg_p = mg_p.rename(columns={resolved_pk: primary_key}) if resolved_pk in mg_p.columns else mg_p
                mg_c = mg_c.rename(columns={resolved_pk: primary_key}) if resolved_pk in mg_c.columns else mg_c

    # Resolve mastergroup column across all provided dataframes (mg and cust frames)
    resolved_master = resolve_column(mastergroup_col, [mg_c, mg_p, cust_gbm_c, cust_cmb_c, cust_gbm_p, cust_cmb_p])
    if resolved_master is None:
        raise ValueError(f"Mastergroup column '{mastergroup_col}' not found in any of the provided frames (case-insensitive search failed).")
    # if resolved name is different, let user know and use resolved name
    if resolved_master != mastergroup_col:
        print(f"Resolved mastergroup column '{mastergroup_col}' -> '{resolved_master}'. Using '{resolved_master}' going forward.")
    mastergroup_col = resolved_master

    # Reduce to inscope columns for comparison logic (validate they exist)
    missing_compare_cols = [c for c in compare_cols if c not in mg_p.columns and c not in mg_c.columns]
    if missing_compare_cols:
        raise ValueError(f"The following inscope columns are not present in either mg_p or mg_c: {missing_compare_cols}")

    mg_p_inscope = mg_p[[c for c in compare_cols if c in mg_p.columns]].drop_duplicates().copy()
    mg_c_inscope = mg_c[[c for c in compare_cols if c in mg_c.columns]].drop_duplicates().copy()

    # ---------- Additions / Deletions ----------
    keys_p = set(mg_p_inscope[primary_key].dropna().unique())
    keys_c = set(mg_c_inscope[primary_key].dropna().unique())

    added_keys = keys_c - keys_p
    deleted_keys = keys_p - keys_c
    common_keys = keys_p & keys_c

    additions = mg_c[mg_c[primary_key].isin(added_keys)].copy().reset_index(drop=True)
    deletions = mg_p[mg_p[primary_key].isin(deleted_keys)].copy().reset_index(drop=True)

    # ---------- Changes (only inscope cols considered; ignore other columns) ----------
    if common_keys:
        # ensure both inscope frames have primary_key column (they should)
        merged = mg_p_inscope.merge(mg_c_inscope, on=primary_key, how='inner', suffixes=('_p', '_c'))

        def row_changed(row):
            # do NOT count primary_key as a change
            for col in compare_cols:
                if col == primary_key:
                    continue
                a = row.get(f"{col}_p", np.nan)
                b = row.get(f"{col}_c", np.nan)
                if pd.isna(a) and pd.isna(b):
                    continue
                try:
                    if (a != b) and not (pd.isna(a) and pd.isna(b)):
                        return True
                except Exception:
                    if str(a) != str(b):
                        return True
            return False

        merged['is_changed'] = merged.apply(row_changed, axis=1)
        changed_keys = merged.loc[merged['is_changed'], primary_key].unique()

        # get original full rows for changed keys
        changes_p_rows = mg_p[mg_p[primary_key].isin(changed_keys)].copy().reset_index(drop=True)
        changes_c_rows = mg_c[mg_c[primary_key].isin(changed_keys)].copy().reset_index(drop=True)

        # Build mapping of changed columns per key
        changed_cols_map = {}
        for _, r in merged[merged['is_changed']].iterrows():
            key = r[primary_key]
            changed = []
            for col in compare_cols:
                if col == primary_key:
                    continue
                try:
                    left = r.get(f"{col}_p", np.nan)
                    right = r.get(f"{col}_c", np.nan)
                    if not (pd.isna(left) and pd.isna(right)) and (left != right):
                        changed.append(col)
                except Exception:
                    if str(r.get(f"{col}_p", '')) != str(r.get(f"{col}_c", '')):
                        changed.append(col)
            changed_cols_map[key] = changed

        # create side-by-side inscope columns with _p/_c suffixes
        p_side = changes_p_rows[[c for c in compare_cols if c in changes_p_rows.columns]].copy().add_suffix('_p')
        c_side = changes_c_rows[[c for c in compare_cols if c in changes_c_rows.columns]].copy().add_suffix('_c')

        # Merge sides on primary key suffix
        changes = p_side.merge(c_side, left_on=f"{primary_key}_p", right_on=f"{primary_key}_c", how='outer', suffixes=('_p','_c'))
        # unify primary key
        changes[primary_key] = changes[f"{primary_key}_c"].combine_first(changes[f"{primary_key}_p"])
        # add changed_columns list
        changes['changed_columns'] = changes[primary_key].map(lambda k: changed_cols_map.get(k, []))
        # drop duplicate key suffix cols
        changes = changes.drop(columns=[f"{primary_key}_p", f"{primary_key}_c"], errors='ignore')

        # Now attach Mastergroup name into changes (prefer current mg_c value)
        # Build series mapping from available mg frames (prefer mg_c, fallback mg_p)
        mg_c_key_master = pd.Series(dtype=object)
        mg_p_key_master = pd.Series(dtype=object)
        if primary_key in mg_c.columns and mastergroup_col in mg_c.columns:
            tmp = mg_c[[primary_key, mastergroup_col]].drop_duplicates()
            if not tmp.empty:
                mg_c_key_master = tmp.set_index(primary_key)[mastergroup_col]
        if primary_key in mg_p.columns and mastergroup_col in mg_p.columns:
            tmp = mg_p[[primary_key, mastergroup_col]].drop_duplicates()
            if not tmp.empty:
                mg_p_key_master = tmp.set_index(primary_key)[mastergroup_col]

        def get_master_for_key(k):
            if (not mg_c_key_master.empty) and (k in mg_c_key_master.index):
                return mg_c_key_master.loc[k]
            elif (not mg_p_key_master.empty) and (k in mg_p_key_master.index):
                return mg_p_key_master.loc[k]
            else:
                return np.nan

        changes[mastergroup_col] = changes[primary_key].map(get_master_for_key)
    else:
        # empty changes frame with sensible columns
        cols = [primary_key] + [f"{c}_p" for c in compare_cols if c != primary_key] + [f"{c}_c" for c in compare_cols if c != primary_key] + ['changed_columns', mastergroup_col]
        changes = pd.DataFrame(columns=cols)

    # For additions and deletions, ensure mastergroup_col exists in those rows (it should in mg)
    if mastergroup_col not in additions.columns:
        if primary_key in mg_c.columns and mastergroup_col in mg_c.columns:
            mg_c_master = mg_c[[primary_key, mastergroup_col]].drop_duplicates()
            additions = additions.merge(mg_c_master, on=primary_key, how='left')
        else:
            additions[mastergroup_col] = np.nan

    if mastergroup_col not in deletions.columns:
        if primary_key in mg_p.columns and mastergroup_col in mg_p.columns:
            mg_p_master = mg_p[[primary_key, mastergroup_col]].drop_duplicates()
            deletions = deletions.merge(mg_p_master, on=primary_key, how='left')
        else:
            deletions[mastergroup_col] = np.nan

    # ---------- Attach financials strictly by Mastergroup name ----------
    def attach_fin_by_mastergroup_strict(df_rows, cust_gbm_source, cust_cmb_source):
        """
        Merge on mastergroup_col (strict). Prefer values from cust_gbm_source; if NaN, use cust_cmb_source.
        """
        if df_rows.empty:
            df_rows[financial_col] = np.nan
            return df_rows

        if mastergroup_col not in df_rows.columns:
            # can't join strictly
            df_rows[financial_col] = np.nan
            return df_rows

        df = df_rows.copy()

        # Prepare cust subsets (only mastergroup_col + financial_col) if they exist
        gbm_sub = pd.DataFrame(columns=[mastergroup_col, financial_col])
        cmb_sub = pd.DataFrame(columns=[mastergroup_col, financial_col])

        if isinstance(cust_gbm_source, pd.DataFrame) and mastergroup_col in cust_gbm_source.columns and financial_col in cust_gbm_source.columns:
            gbm_sub = cust_gbm_source[[mastergroup_col, financial_col]].drop_duplicates()
        if isinstance(cust_cmb_source, pd.DataFrame) and mastergroup_col in cust_cmb_source.columns and financial_col in cust_cmb_source.columns:
            cmb_sub = cust_cmb_source[[mastergroup_col, financial_col]].drop_duplicates()

        # Merge gbm
        tmp = df.merge(gbm_sub, on=mastergroup_col, how='left')

        # If missing, try cmb for those mastergroups
        if financial_col not in tmp.columns:
            tmp[financial_col] = np.nan

        na_mask = tmp[financial_col].isna()
        if na_mask.any() and not cmb_sub.empty:
            missing_series = tmp.loc[na_mask, mastergroup_col].drop_duplicates()
            if not missing_series.empty:
                missing_df = pd.DataFrame({mastergroup_col: missing_series.values})
                merged_cmb = missing_df.merge(cmb_sub, on=mastergroup_col, how='left')
                tmp = tmp.merge(merged_cmb, on=mastergroup_col, how='left', suffixes=('','_cmbtmp'))
                if financial_col + '_cmbtmp' in tmp.columns:
                    tmp[financial_col] = tmp[financial_col].fillna(tmp[financial_col + '_cmbtmp'])
                    tmp = tmp.drop(columns=[financial_col + '_cmbtmp'])

        tmp[financial_col] = pd.to_numeric(tmp[financial_col], errors='coerce')
        return tmp

    # Attach: additions & changes from current cust frames; deletions from prior cust frames
    additions = attach_fin_by_mastergroup_strict(additions, cust_gbm_c, cust_cmb_c)
    changes = attach_fin_by_mastergroup_strict(changes, cust_gbm_c, cust_cmb_c)
    deletions = attach_fin_by_mastergroup_strict(deletions, cust_gbm_p, cust_cmb_p)

    # Ensure financial_col exists numeric for sorting
    for df_ in (additions, changes, deletions):
        if financial_col not in df_.columns:
            df_[financial_col] = np.nan
        df_[financial_col] = pd.to_numeric(df_[financial_col], errors='coerce')

    # Sort descending by financial_col
    additions = additions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    changes = changes.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    deletions = deletions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)

    # Write to Excel (each sheet contains original rows + financial column)
    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        additions.to_excel(writer, sheet_name='Additions', index=False)
        deletions.to_excel(writer, sheet_name='Deletions', index=False)
        changes.to_excel(writer, sheet_name='Changes', index=False)

    return {
        'additions': additions,
        'deletions': deletions,
        'changes': changes,
        'out_file': out_file
    }

# Quick debug helpers (run if you still see a KeyError):
# print("mg_p columns:", mg_p.columns.tolist())
# print("mg_c columns:", mg_c.columns.tolist())
# print("cust_gbm_c columns:", cust_gbm_c.columns.tolist())
# print("cust_cmb_c columns:", cust_cmb_c.columns.tolist())
