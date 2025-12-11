```
import pandas as pd
import numpy as np
import re

def compare_mg_inscope_mastergroup_strict(
    mg_p,
    mg_c,
    cust_gbm_p,
    cust_cmb_p,
    cust_gbm_c,
    cust_cmb_c,
    inscope_col_mg,
    financial_col='Total Operating Income (HORIS YTD Financials)',
    mastergroup_col='Mastergroup Name',
    out_file='mg_exceptions_mastergroup_strict.xlsx'
):
    """
    Robust compare function (drop-in).
    - Normalizes/flattens column names across frames
    - Resolves primary key and inscope columns (whitespace/case-insensitive)
    - Avoids KeyError and never uses set_index() on possibly non-unique keys
    - Attaches financials strictly by mastergroup (prefer cust_gbm then cust_cmb)
    - Writes Additions/Deletions/Changes sheets and returns dict of DataFrames.
    """

    # -------------------- helpers --------------------
    def flatten_multiindex_cols(df):
        if getattr(df.columns, "nlevels", 1) > 1:
            df = df.copy()
            df.columns = [" ".join([("" if v is None else str(v)).strip() for v in col]).strip() for col in df.columns]
        return df

    def normalize_cols_list(cols):
        return re.sub(r'\s+', '', str(cols)).lower()

    def normalize_and_make_unique(df):
        """Return cleaned df and map(normalized_key -> actual_column_name)"""
        df = df.copy()
        df = flatten_multiindex_cols(df)

        new_cols = []
        for c in df.columns:
            s = "" if c is None else str(c)
            s = s.replace('\xa0', ' ')
            s = re.sub(r'\s+', ' ', s)
            s = s.strip()
            new_cols.append(s)

        # make duplicates unique
        seen = {}
        unique_cols = []
        for s in new_cols:
            if s in seen:
                seen[s] += 1
                new = f"{s}__dup{seen[s]}"
            else:
                seen[s] = 0
                new = s
            unique_cols.append(new)

        df.columns = unique_cols

        norm_map = {}
        for actual in unique_cols:
            nk = normalize_cols_list(actual)
            if nk not in norm_map:
                norm_map[nk] = actual

        return df, norm_map

    def build_combined_map(maps_in_order):
        combined = {}
        for m in maps_in_order:
            if not isinstance(m, dict):
                continue
            for k, v in m.items():
                if k not in combined:
                    combined[k] = v
        return combined

    def resolve_name(requested, combined_map):
        nk = normalize_cols_list(requested)
        return combined_map.get(nk)

    def safe_select(df, cols):
        out = pd.DataFrame(index=df.index)
        for c in cols:
            if c in df.columns:
                out[c] = df[c]
            else:
                out[c] = np.nan
        return out

    def build_safe_mapping(df, key_col, val_col):
        """
        Build safe mapping dict from key_col -> val_col.
        - If duplicates exist, first non-null wins.
        - If columns missing, returns {}.
        """
        mapping = {}
        if not isinstance(df, pd.DataFrame):
            return mapping
        if key_col not in df.columns or val_col not in df.columns:
            return mapping

        for _, r in df[[key_col, val_col]].dropna(subset=[key_col]).iterrows():
            k = r[key_col]
            v = r[val_col]
            if k not in mapping:
                mapping[k] = v
        return mapping

    # -------------------- normalize all frames --------------------
    frames = {
        'mg_p': mg_p,
        'mg_c': mg_c,
        'cust_gbm_p': cust_gbm_p,
        'cust_cmb_p': cust_cmb_p,
        'cust_gbm_c': cust_gbm_c,
        'cust_cmb_c': cust_cmb_c
    }

    norm_maps = {}
    for k, df in frames.items():
        if isinstance(df, pd.DataFrame):
            df_clean, map_clean = normalize_and_make_unique(df)
            frames[k] = df_clean
            norm_maps[k] = map_clean
        else:
            frames[k] = pd.DataFrame()
            norm_maps[k] = {}

    mg_p = frames['mg_p']
    mg_c = frames['mg_c']
    cust_gbm_p = frames['cust_gbm_p']
    cust_cmb_p = frames['cust_cmb_p']
    cust_gbm_c = frames['cust_gbm_c']
    cust_cmb_c = frames['cust_cmb_c']

    combined_map = build_combined_map([
        norm_maps.get('mg_c', {}), norm_maps.get('mg_p', {}),
        norm_maps.get('cust_gbm_c', {}), norm_maps.get('cust_cmb_c', {}),
        norm_maps.get('cust_gbm_p', {}), norm_maps.get('cust_cmb_p', {})
    ])

    # -------------------- resolve primary key and inscope --------------------
    inscope = list(inscope_col_mg)
    if len(inscope) == 0:
        raise ValueError("inscope_col_mg must contain at least one column name.")

    requested_pk = inscope[0]
    resolved_pk = resolve_name(requested_pk, combined_map)
    if resolved_pk is None:
        raise ValueError(f"Primary key '{requested_pk}' not found after normalization. Available keys: {list(combined_map.keys())}")

    resolved_compare_cols = []
    for c in inscope:
        rc = resolve_name(c, combined_map)
        if rc is None:
            raise ValueError(f"Inscope column '{c}' not found after normalization. Available keys: {list(combined_map.keys())}")
        resolved_compare_cols.append(rc)

    primary_key = resolved_pk
    compare_cols = resolved_compare_cols

    # -------------------- resolve mastergroup --------------------
    resolved_master = resolve_name(mastergroup_col, combined_map)
    if resolved_master is None:
        raise ValueError(f"Mastergroup column '{mastergroup_col}' not found after normalization. Available keys: {list(combined_map.keys())}")
    mastergroup_col = resolved_master

    # -------------------- prepare inscope subsets --------------------
    mg_p_inscope = safe_select(mg_p, compare_cols).drop_duplicates().copy()
    mg_c_inscope = safe_select(mg_c, compare_cols).drop_duplicates().copy()

    # -------------------- additions / deletions --------------------
    keys_p = set(mg_p_inscope[primary_key].dropna().unique())
    keys_c = set(mg_c_inscope[primary_key].dropna().unique())

    added_keys = keys_c - keys_p
    deleted_keys = keys_p - keys_c
    common_keys = keys_p & keys_c

    additions = mg_c[mg_c[primary_key].isin(added_keys)].copy().reset_index(drop=True) if not mg_c.empty and primary_key in mg_c.columns else pd.DataFrame(columns=mg_c.columns)
    deletions = mg_p[mg_p[primary_key].isin(deleted_keys)].copy().reset_index(drop=True) if not mg_p.empty and primary_key in mg_p.columns else pd.DataFrame(columns=mg_p.columns)

    # -------------------- changes --------------------
    if common_keys:
        merged = mg_p_inscope.merge(mg_c_inscope, on=primary_key, how='inner', suffixes=('_p', '_c'))

        def row_changed(row):
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

        changes_p_rows = mg_p[mg_p[primary_key].isin(changed_keys)].copy().reset_index(drop=True) if primary_key in mg_p.columns else pd.DataFrame()
        changes_c_rows = mg_c[mg_c[primary_key].isin(changed_keys)].copy().reset_index(drop=True) if primary_key in mg_c.columns else pd.DataFrame()

        # Build changed columns map
        changed_cols_map = {}
        for _, r in merged[merged['is_changed']].iterrows():
            key = r[primary_key]
            changed = []
            for col in compare_cols:
                if col == primary_key:
                    continue
                left = r.get(f"{col}_p", np.nan)
                right = r.get(f"{col}_c", np.nan)
                try:
                    if not (pd.isna(left) and pd.isna(right)) and (left != right):
                        changed.append(col)
                except Exception:
                    if str(left) != str(right):
                        changed.append(col)
            changed_cols_map[key] = changed

        # side-by-side
        p_side = changes_p_rows[[c for c in compare_cols if c in changes_p_rows.columns]].copy().add_suffix('_p')
        c_side = changes_c_rows[[c for c in compare_cols if c in changes_c_rows.columns]].copy().add_suffix('_c')

        left_key = f"{primary_key}_p"
        right_key = f"{primary_key}_c"
        changes = p_side.merge(c_side, left_on=left_key, right_on=right_key, how='outer', suffixes=('_p','_c'))

        if right_key in changes.columns or left_key in changes.columns:
            changes[primary_key] = changes[right_key].combine_first(changes[left_key])
        else:
            changes[primary_key] = np.nan

        changes['changed_columns'] = changes[primary_key].map(lambda k: changed_cols_map.get(k, []))
        changes = changes.drop(columns=[left_key, right_key], errors='ignore')

        # Attach mastergroup using safe mappings (no set_index)
        mg_c_key_master_map = build_safe_mapping(mg_c, primary_key, mastergroup_col)
        mg_p_key_master_map = build_safe_mapping(mg_p, primary_key, mastergroup_col)

        def get_master_for_key(k):
            if k in mg_c_key_master_map:
                return mg_c_key_master_map[k]
            if k in mg_p_key_master_map:
                return mg_p_key_master_map[k]
            return np.nan

        changes[mastergroup_col] = changes[primary_key].map(get_master_for_key)
    else:
        cols = [primary_key] + [f"{c}_p" for c in compare_cols if c != primary_key] + [f"{c}_c" for c in compare_cols if c != primary_key] + ['changed_columns', mastergroup_col]
        changes = pd.DataFrame(columns=cols)

    # -------------------- ensure mastergroup exists for additions/deletions --------------------
    if mastergroup_col not in additions.columns:
        if primary_key in mg_c.columns and mastergroup_col in mg_c.columns and not mg_c[[primary_key, mastergroup_col]].drop_duplicates().empty and not additions.empty:
            additions = additions.merge(mg_c[[primary_key, mastergroup_col]].drop_duplicates(), on=primary_key, how='left')
        else:
            additions[mastergroup_col] = np.nan

    if mastergroup_col not in deletions.columns:
        if primary_key in mg_p.columns and mastergroup_col in mg_p.columns and not mg_p[[primary_key, mastergroup_col]].drop_duplicates().empty and not deletions.empty:
            deletions = deletions.merge(mg_p[[primary_key, mastergroup_col]].drop_duplicates(), on=primary_key, how='left')
        else:
            deletions[mastergroup_col] = np.nan

    # -------------------- attach financials strictly by mastergroup --------------------
    def attach_fin_by_mastergroup_strict(df_rows, cust_gbm_source, cust_cmb_source):
        if df_rows.empty:
            df_rows[financial_col] = np.nan
            return df_rows
        if mastergroup_col not in df_rows.columns:
            df_rows[financial_col] = np.nan
            return df_rows

        df = df_rows.copy()

        gbm_sub = pd.DataFrame(columns=[mastergroup_col, financial_col])
        cmb_sub = pd.DataFrame(columns=[mastergroup_col, financial_col])

        if isinstance(cust_gbm_source, pd.DataFrame) and mastergroup_col in cust_gbm_source.columns and financial_col in cust_gbm_source.columns:
            gbm_sub = cust_gbm_source[[mastergroup_col, financial_col]].drop_duplicates()
        if isinstance(cust_cmb_source, pd.DataFrame) and mastergroup_col in cust_cmb_source.columns and financial_col in cust_cmb_source.columns:
            cmb_sub = cust_cmb_source[[mastergroup_col, financial_col]].drop_duplicates()

        tmp = df.merge(gbm_sub, on=mastergroup_col, how='left')

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

    additions = attach_fin_by_mastergroup_strict(additions, cust_gbm_c, cust_cmb_c)
    changes = attach_fin_by_mastergroup_strict(changes, cust_gbm_c, cust_cmb_c)
    deletions = attach_fin_by_mastergroup_strict(deletions, cust_gbm_p, cust_cmb_p)

    # -------------------- ensure financial_col numeric and sort --------------------
    for df_ in (additions, changes, deletions):
        if financial_col not in df_.columns:
            df_[financial_col] = np.nan
        df_[financial_col] = pd.to_numeric(df_[financial_col], errors='coerce')

    additions = additions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    changes = changes.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    deletions = deletions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)

    # -------------------- write excel and return --------------------
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
