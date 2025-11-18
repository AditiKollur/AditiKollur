import pandas as pd
import numpy as np

def compare_mg_inscope_only(
    mg_p,
    mg_c,
    cust_gbm_p,
    cust_cmb_p,
    cust_gbm_c,
    cust_cmb_c,
    inscope_col_mg,
    financial_col='Total Operating Income (HORIS YTD Financials)',
    out_file='mg_exceptions_inscope.xlsx'
):
    """
    Compare mg_p (prior) and mg_c (current) considering ONLY columns in inscope_col_mg.
    The first column in inscope_col_mg is treated as the primary key for alignment.
    Writes Additions / Deletions / Changes to Excel and returns them as DataFrames.
    """

    # defensive copies
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
    compare_cols = inscope[:]  # all inscope columns are relevant
    # We will align on primary_key; changes mean other inscope columns (or entire inscope set) differ.

    # Ensure primary_key exists
    if primary_key not in mg_p.columns or primary_key not in mg_c.columns:
        raise ValueError(f"Primary key '{primary_key}' must exist in both mg_p and mg_c.")

    # Reduce to only inscope cols for deterministic comparison & to keep original data requested
    mg_p_inscope = mg_p[compare_cols].drop_duplicates().copy()
    mg_c_inscope = mg_c[compare_cols].drop_duplicates().copy()

    # ---------- Additions and Deletions based on primary key ----------
    keys_p = set(mg_p_inscope[primary_key].dropna().unique())
    keys_c = set(mg_c_inscope[primary_key].dropna().unique())

    added_keys = keys_c - keys_p
    deleted_keys = keys_p - keys_c
    common_keys = keys_p & keys_c

    additions = mg_c[mg_c[primary_key].isin(added_keys)].copy().reset_index(drop=True)
    deletions = mg_p[mg_p[primary_key].isin(deleted_keys)].copy().reset_index(drop=True)

    # ---------- Changes: same primary_key but differences in any inscope column ----------
    if common_keys:
        # inner merge on primary_key to compare inscope columns. Keep suffixes to distinguish.
        merged = mg_p_inscope.merge(mg_c_inscope, on=primary_key, how='inner', suffixes=('_p', '_c'))
        def row_changed(row):
            for col in compare_cols:
                col_p = f"{col}_p"
                col_c = f"{col}_c"
                # don't count primary key itself as a change for 'Changes' detection;
                if col == primary_key:
                    continue
                a = row.get(col_p, np.nan)
                b = row.get(col_c, np.nan)
                if pd.isna(a) and pd.isna(b):
                    continue
                # robust compare: treat numeric vs string carefully
                try:
                    if (a != b) and not (pd.isna(a) and pd.isna(b)):
                        return True
                except Exception:
                    if str(a) != str(b):
                        return True
            return False

        merged['is_changed'] = merged.apply(row_changed, axis=1)
        changed_keys = merged.loc[merged['is_changed'], primary_key].unique()
        # Extract original full rows from mg_p and mg_c for changed keys and join side-by-side if desired.
        changes_p_rows = mg_p[mg_p[primary_key].isin(changed_keys)].copy().reset_index(drop=True)
        changes_c_rows = mg_c[mg_c[primary_key].isin(changed_keys)].copy().reset_index(drop=True)

        # For user-friendliness, create a single Changes frame that contains:
        # primary_key, all inscope cols from prior (suffix _p), all inscope cols from current (suffix _c),
        # and a column changed_columns listing which inscope columns changed.
        # Build mapping of changed columns per key:
        changed_cols_map = {}
        for _, r in merged[merged['is_changed']].iterrows():
            key = r[primary_key]
            changed = []
            for col in compare_cols:
                if col == primary_key:
                    continue
                if not (pd.isna(r[f"{col}_p"]) and pd.isna(r[f"{col}_c"])):
                    try:
                        if r[f"{col}_p"] != r[f"{col}_c"]:
                            changed.append(col)
                    except Exception:
                        if str(r[f"{col}_p"]) != str(r[f"{col}_c"]):
                            changed.append(col)
            changed_cols_map[key] = changed

        # prepare side-by-side df
        p_side = changes_p_rows[compare_cols].copy()
        p_side = p_side.add_suffix('_p')
        c_side = changes_c_rows[compare_cols].copy()
        c_side = c_side.add_suffix('_c')

        # merge p_side and c_side on primary key columns (primary key suffix exists in both)
        changes = p_side.merge(c_side, left_on=f"{primary_key}_p", right_on=f"{primary_key}_c", how='outer', suffixes=('_p','_c'))
        # unify primary key column name
        changes[primary_key] = changes[f"{primary_key}_c"].combine_first(changes[f"{primary_key}_p"])
        # add changed_columns list
        changes['changed_columns'] = changes[primary_key].map(lambda k: changed_cols_map.get(k, []))
        # reorder to put primary key first
        # drop duplicate key columns
        changes = changes.drop(columns=[f"{primary_key}_p", f"{primary_key}_c"], errors='ignore')
        cols_order = [primary_key] + [c for c in changes.columns if c != primary_key]
        changes = changes[cols_order].reset_index(drop=True)
    else:
        changes = pd.DataFrame(columns=[primary_key] + [f"{c}_p" for c in compare_cols if c!=primary_key] + [f"{c}_c" for c in compare_cols if c!=primary_key] + ['changed_columns'])

    # ---------- Attach financials ----------
    def attach_fin(df_rows, cust_gbm_source, cust_cmb_source, join_on_cols):
        """Prefer cust_gbm then cust_cmb. join_on_cols can be a list (preferably primary_key)."""
        if df_rows.empty:
            df_rows[financial_col] = np.nan
            return df_rows
        df = df_rows.copy()
        # If join_on_cols are not all present in cust_gbm, fall back to cust_cmb or just primary_key
        join_cols = [c for c in join_on_cols if c in cust_gbm_source.columns]
        use_gbm = cust_gbm_source
        use_cmb = cust_cmb_source
        if not join_cols:
            # try cust_cmb
            join_cols = [c for c in join_on_cols if c in cust_cmb_source.columns]
            use_gbm = cust_cmb_source
            use_cmb = pd.DataFrame(columns=[])  # nothing left
        if not join_cols:
            # fallback to primary_key if present there
            join_cols = [c for c in join_on_cols if c in use_gbm.columns]
        if not join_cols:
            # can't join -> set NaN
            df[financial_col] = np.nan
            return df

        # merge from gbm
        tmp = df.merge(use_gbm[[*join_cols, financial_col]].drop_duplicates(), on=join_cols, how='left')
        if financial_col not in tmp.columns:
            tmp[financial_col] = np.nan

        # for missing financials, try cmb
        na_mask = tmp[financial_col].isna()
        if na_mask.any() and not use_cmb.empty:
            missing = tmp.loc[na_mask, join_cols].drop_duplicates()
            if not missing.empty:
                merged_cmb = missing.merge(use_cmb[[*join_cols, financial_col]].drop_duplicates(), on=join_cols, how='left')
                # align and fill
                tmp = tmp.merge(merged_cmb, on=join_cols, how='left', suffixes=('','_cmb_tmp'))
                # prefer existing financial, else fill with _cmb_tmp
                if financial_col + '_cmb_tmp' in tmp.columns:
                    tmp[financial_col] = tmp[financial_col].fillna(tmp[financial_col + '_cmb_tmp'])
                    tmp = tmp.drop(columns=[financial_col + '_cmb_tmp'])
        # coerce numeric
        tmp[financial_col] = pd.to_numeric(tmp[financial_col], errors='coerce')
        return tmp

    # For additions and changes attach from current month cust sources, for deletions use prior month cust sources
    # For joining financials prefer full inscope columns if present; otherwise primary key
    join_cols_for_fin = [c for c in compare_cols if c in cust_gbm_c.columns or c in cust_cmb_c.columns]
    if not join_cols_for_fin:
        join_cols_for_fin = [primary_key]

    additions = attach_fin(additions, cust_gbm_c, cust_cmb_c, join_cols_for_fin)
    changes = attach_fin(changes, cust_gbm_c, cust_cmb_c, join_cols_for_fin)
    deletions = attach_fin(deletions, cust_gbm_p, cust_cmb_p, join_cols_for_fin)

    # Ensure financial column exists numeric for sorting
    for df_ in (additions, changes, deletions):
        if financial_col not in df_.columns:
            df_[financial_col] = np.nan
        df_[financial_col] = pd.to_numeric(df_[financial_col], errors='coerce')

    # Sort descending by financial_col (NaNs last)
    additions = additions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    changes = changes.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    deletions = deletions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)

    # Write to Excel
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

