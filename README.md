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
    Full robust implementation that:
    - normalizes & flattens column names (handles NBSP, MultiIndex, trailing spaces),
    - resolves primary key + inscope cols + mastergroup via whitespace/case-insensitive lookup,
    - avoids set_index() and builds safe mappings (handles duplicates, unhashable keys),
    - uses shape-safe assignments to avoid "unshapable series" errors,
    - attaches financials strictly by mastergroup (prefers cust_gbm, then cust_cmb),
    - writes Additions/Deletions/Changes sheets and returns dict of DataFrames.
    """

    # -------------------- helpers --------------------
    def flatten_multiindex_cols(df):
        if getattr(df.columns, "nlevels", 1) > 1:
            df = df.copy()
            df.columns = [" ".join([("" if v is None else str(v)).strip() for v in col]).strip() for col in df.columns]
        return df

    def normalize_cols_list(colname):
        """Normalized key for comparing column names: remove whitespace and lowercase."""
        return re.sub(r'\s+', '', str(colname)).lower()

    def normalize_and_make_unique(df):
        """
        Return cleaned df and a map: normalized_key -> actual_clean_column_name.
        - cleans NBSPs, collapses whitespaces, strips
        - makes duplicate names unique by appending __dupN
        """
        df = df.copy()
        df = flatten_multiindex_cols(df)

        cleaned = []
        for c in df.columns:
            s = "" if c is None else str(c)
            s = s.replace('\xa0', ' ')
            s = re.sub(r'\s+', ' ', s)
            s = s.strip()
            cleaned.append(s)

        seen = {}
        unique_cols = []
        for s in cleaned:
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

    # ---- robust key normalization / hashing for mapping ----
    def _normalize_key_for_mapping(k):
        """Return a stable, hashable representation for mapping lookups or None for invalid keys."""
        # skip None / NaN
        try:
            if k is None:
                return None
            if isinstance(k, float) and np.isnan(k):
                return None
        except Exception:
            pass

        # if hashable, return as-is
        try:
            hash(k)
            return k
        except TypeError:
            # convert common unhashable types
            if isinstance(k, (list, tuple)):
                return tuple(k)
            # fallback to string representation
            try:
                return str(k)
            except Exception:
                return repr(k)

    def build_safe_mapping(df, key_col, val_col):
        """
        Build mapping normalized_key -> value (first non-null wins).
        Skips NaN keys and converts unhashable keys to stable forms.
        """
        mapping = {}
        if not isinstance(df, pd.DataFrame):
            return mapping
        if key_col not in df.columns or val_col not in df.columns:
            return mapping

        for _, r in df[[key_col, val_col]].iterrows():
            raw_k = r[key_col]
            # skip None/NaN
            try:
                if raw_k is None or (isinstance(raw_k, float) and np.isnan(raw_k)):
                    continue
            except Exception:
                pass

            nk = _normalize_key_for_mapping(raw_k)
            if nk is None:
                continue
            if nk not in mapping:
                mapping[nk] = r[val_col]
        return mapping

    def _lookup_in_mapping(mapping, raw_key):
        nk = _normalize_key_for_mapping(raw_key)
        if nk is None:
            return np.nan
        return mapping.get(nk, np.nan)

    # -------------------- normalize all incoming frames --------------------
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

    # -------------------- resolve inscope + primary key --------------------
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
    mg_p_inscope = safe_select(mg_p, compare_cols).drop_duplicates().reset_index(drop=True).copy()
    mg_c_inscope = safe_select(mg_c, compare_cols).drop_duplicates().reset_index(drop=True).copy()

    # -------------------- additions / deletions --------------------
    keys_p = set(mg_p_inscope[primary_key].dropna().unique())
    keys_c = set(mg_c_inscope[primary_key].dropna().unique())

    added_keys = keys_c - keys_p
    deleted_keys = keys_p - keys_c
    common_keys = keys_p & keys_c

    additions = (mg_c[mg_c[primary_key].isin(added_keys)].copy().reset_index(drop=True)
                 if not mg_c.empty and primary_key in mg_c.columns else pd.DataFrame(columns=mg_c.columns))
    deletions = (mg_p[mg_p[primary_key].isin(deleted_keys)].copy().reset_index(drop=True)
                 if not mg_p.empty and primary_key in mg_p.columns else pd.DataFrame(columns=mg_p.columns))

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

        changes_p_rows = (mg_p[mg_p[primary_key].isin(changed_keys)].copy().reset_index(drop=True)
                          if primary_key in mg_p.columns else pd.DataFrame())
        changes_c_rows = (mg_c[mg_c[primary_key].isin(changed_keys)].copy().reset_index(drop=True)
                          if primary_key in mg_c.columns else pd.DataFrame())

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

        # side-by-side (reset indices to ensure shape-safe assignments)
        p_side = changes_p_rows[[c for c in compare_cols if c in changes_p_rows.columns]].copy().add_suffix('_p').reset_index(drop=True)
        c_side = changes_c_rows[[c for c in compare_cols if c in changes_c_rows.columns]].copy().add_suffix('_c').reset_index(drop=True)

        left_key = f"{primary_key}_p"
        right_key = f"{primary_key}_c"
        changes = p_side.merge(c_side, left_on=left_key, right_on=right_key, how='outer', suffixes=('_p','_c')).reset_index(drop=True)

        # unify primary key as shape-safe list assignment
        pk_values = []
        for i in range(len(changes)):
            v_right = changes.at[i, right_key] if right_key in changes.columns else np.nan
            v_left = changes.at[i, left_key] if left_key in changes.columns else np.nan
            pk_values.append(v_right if pd.notna(v_right) else v_left)
        changes[primary_key] = pk_values

        # changed_columns as shape-safe list
        changes['changed_columns'] = [changed_cols_map.get(k, []) for k in changes[primary_key].tolist()]

        # drop duplicate key suffix cols
        changes = changes.drop(columns=[left_key, right_key], errors='ignore')

        # Attach mastergroup using safe mapping and shape-safe assignment
        mg_c_key_master_map = build_safe_mapping(mg_c, primary_key, mastergroup_col)
        mg_p_key_master_map = build_safe_mapping(mg_p, primary_key, mastergroup_col)

        def get_master_for_key(k):
            v = _lookup_in_mapping(mg_c_key_master_map, k)
            if pd.notna(v):
                return v
            v = _lookup_in_mapping(mg_p_key_master_map, k)
            if pd.notna(v):
                return v
            return np.nan

        changes[mastergroup_col] = [get_master_for_key(k) for k in changes[primary_key].tolist()]

    else:
        cols = [primary_key] + [f"{c}_p" for c in compare_cols if c != primary_key] + [f"{c}_c" for c in compare_cols if c != primary_key] + ['changed_columns', mastergroup_col]
        changes = pd.DataFrame(columns=cols)

    # -------------------- ensure mastergroup exists for additions/deletions --------------------
    if mastergroup_col not in additions.columns:
        if primary_key in mg_c.columns and mastergroup_col in mg_c.columns and not mg_c[[primary_key, mastergroup_col]].drop_duplicates().empty and not additions.empty:
            additions = additions.merge(mg_c[[primary_key, mastergroup_col]].drop_duplicates(), on=primary_key, how='left').reset_index(drop=True)
        else:
            additions[mastergroup_col] = np.nan

    if mastergroup_col not in deletions.columns:
        if primary_key in mg_p.columns and mastergroup_col in mg_p.columns and not mg_p[[primary_key, mastergroup_col]].drop_duplicates().empty and not deletions.empty:
            deletions = deletions.merge(mg_p[[primary_key, mastergroup_col]].drop_duplicates(), on=primary_key, how='left').reset_index(drop=True)
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

        df = df_rows.copy().reset_index(drop=True)

        gbm_sub = pd.DataFrame(columns=[mastergroup_col, financial_col])
        cmb_sub = pd.DataFrame(columns=[mastergroup_col, financial_col])

        if isinstance(cust_gbm_source, pd.DataFrame) and mastergroup_col in cust_gbm_source.columns and financial_col in cust_gbm_source.columns:
            gbm_sub = cust_gbm_source[[mastergroup_col, financial_col]].drop_duplicates().reset_index(drop=True)
        if isinstance(cust_cmb_source, pd.DataFrame) and mastergroup_col in cust_cmb_source.columns and financial_col in cust_cmb_source.columns:
            cmb_sub = cust_cmb_source[[mastergroup_col, financial_col]].drop_duplicates().reset_index(drop=True)

        tmp = df.merge(gbm_sub, on=mastergroup_col, how='left').reset_index(drop=True)

        if financial_col not in tmp.columns:
            tmp[financial_col] = np.nan

        na_mask = tmp[financial_col].isna()
        if na_mask.any() and not cmb_sub.empty:
            missing_series = tmp.loc[na_mask, mastergroup_col].drop_duplicates()
            if not missing_series.empty:
                missing_df = pd.DataFrame({mastergroup_col: missing_series.values})
                merged_cmb = missing_df.merge(cmb_sub, on=mastergroup_col, how='left')
                tmp = tmp.merge(merged_cmb, on=mastergroup_col, how='left', suffixes=('','_cmbtmp')).reset_index(drop=True)
                if financial_col + '_cmbtmp' in tmp.columns:
                    tmp[financial_col] = tmp[financial_col].fillna(tmp[financial_col + '_cmbtmp'])
                    tmp = tmp.drop(columns=[financial_col + '_cmbtmp'])

        tmp[financial_col] = pd.to_numeric(tmp[financial_col], errors='coerce')
        return tmp.reset_index(drop=True)

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
