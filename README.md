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

    # Basic validation
    if primary_key not in mg_p.columns or primary_key not in mg_c.columns:
        raise ValueError(f"Primary key '{primary_key}' must exist in both mg_p and mg_c.")
    if mastergroup_col not in mg_p.columns and mastergroup_col not in mg_c.columns:
        raise ValueError(f"Mastergroup column '{mastergroup_col}' must exist in mg_p or mg_c.")
    if mastergroup_col not in cust_gbm_p.columns and mastergroup_col not in cust_cmb_p.columns \
       and mastergroup_col not in cust_gbm_c.columns and mastergroup_col not in cust_cmb_c.columns:
        raise ValueError(f"Mastergroup column '{mastergroup_col}' must exist in at least one cust frame.")

    # Reduce to inscope columns for comparison logic
    mg_p_inscope = mg_p[compare_cols].drop_duplicates().copy()
    mg_c_inscope = mg_c[compare_cols].drop_duplicates().copy()

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
                    if not (pd.isna(r[f"{col}_p"]) and pd.isna(r[f"{col}_c"])) and (r[f"{col}_p"] != r[f"{col}_c"]):
                        changed.append(col)
                except Exception:
                    if str(r[f"{col}_p"]) != str(r[f"{col}_c"]):
                        changed.append(col)
            changed_cols_map[key] = changed

        # create side-by-side inscope columns with _p/_c suffixes
        p_side = changes_p_rows[compare_cols].copy().add_suffix('_p')
        c_side = changes_c_rows[compare_cols].copy().add_suffix('_c')

        # Merge sides on primary key suffix
        changes = p_side.merge(c_side, left_on=f"{primary_key}_p", right_on=f"{primary_key}_c", how='outer', suffixes=('_p','_c'))
        # unify primary key
        changes[primary_key] = changes[f"{primary_key}_c"].combine_first(changes[f"{primary_key}_p"])
        # add changed_columns list
        changes['changed_columns'] = changes[primary_key].map(lambda k: changed_cols_map.get(k, []))
        # drop duplicate key suffix cols
        changes = changes.drop(columns=[f"{primary_key}_p", f"{primary_key}_c"], errors='ignore')

        # Now attach Mastergroup name into changes (prefer current mg_c value)
        # build helper df: key -> mastergroup (prefer mg_c)
        mg_c_key_master = mg_c[[primary_key, mastergroup_col]].drop_duplicates().set_index(primary_key)
        mg_p_key_master = mg_p[[primary_key, mastergroup_col]].drop_duplicates().set_index(primary_key)
        def get_master_for_key(k):
            if k in mg_c_key_master.index:
                return mg_c_key_master.loc[k, mastergroup_col]
            elif k in mg_p_key_master.index:
                return mg_p_key_master.loc[k, mastergroup_col]
            else:
                return np.nan
        changes[mastergroup_col] = changes[primary_key].map(get_master_for_key)
    else:
        # empty changes frame with sensible columns
        cols = [primary_key] + [f"{c}_p" for c in compare_cols if c != primary_key] + [f"{c}_c" for c in compare_cols if c != primary_key] + ['changed_columns', mastergroup_col]
        changes = pd.DataFrame(columns=cols)

    # For additions and deletions, ensure mastergroup_col exists in those rows (it should in mg)
    if mastergroup_col not in additions.columns:
        # try to bring mastergroup from mg_c subset
        mg_c_master = mg_c[[primary_key, mastergroup_col]].drop_duplicates()
        additions = additions.merge(mg_c_master, on=primary_key, how='left')
    if mastergroup_col not in deletions.columns:
        mg_p_master = mg_p[[primary_key, mastergroup_col]].drop_duplicates()
        deletions = deletions.merge(mg_p_master, on=primary_key, how='left')

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

        # Prepare cust subsets (only mastergroup_col + financial_col); if financial_col missing, empty df
        gbm_sub = cust_gbm_source[[mastergroup_col, financial_col]].drop_duplicates() if financial_col in cust_gbm_source.columns and mastergroup_col in cust_gbm_source.columns else pd.DataFrame(columns=[mastergroup_col, financial_col])
        cmb_sub = cust_cmb_source[[mastergroup_col, financial_col]].drop_duplicates() if financial_col in cust_cmb_source.columns and mastergroup_col in cust_cmb_source.columns else pd.DataFrame(columns=[mastergroup_col, financial_col])

        # Merge gbm
        tmp = df.merge(gbm_sub, on=mastergroup_col, how='left')

        # If missing, try cmb for those mastergroups
        if financial_col not in tmp.columns:
            tmp[financial_col] = np.nan

        na_mask = tmp[financial_col].isna()
        if na_mask.any() and not cmb_sub.empty:
            missing = tmp.loc[na_mask, mastergroup_col].drop_duplicates()
            if not missing.empty:
                merged_cmb = missing.merge(cmb_sub, on=mastergroup_col, how='left')
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

res = compare_mg_inscope_mastergroup_strict(
    mg_p=mg_prior_df,
    mg_c=mg_current_df,
    cust_gbm_p=cust_gbm_prior_df,
    cust_cmb_p=cust_cmb_prior_df,
    cust_gbm_c=cust_gbm_current_df,
    cust_cmb_c=cust_cmb_current_df,
    inscope_col_mg=['Mastergroup ID', 'Mastergroup name', 'SomeOtherInScopeCol']  # example
)
print("Wrote:", res['out_file'])
res['additions'].head()




import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="MG Exception Storytelling", layout="wide")

# ---------------------------------------------------------
# UTILITY FUNCTIONS
# ---------------------------------------------------------

def kpi_card(label, value, color="#2266cc"):
    st.markdown(
        f"""
        <div style="padding:15px;border-radius:10px;background:{color}20;border-left:6px solid {color};margin-bottom:10px">
            <h4 style="margin:0;color:{color}">{label}</h4>
            <h2 style="margin:0;color:#000">{value}</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

def download_excel(df, filename="data.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def waterfall_chart(added, deleted, final):
    fig = go.Figure(go.Waterfall(
        name="Revenue Movement",
        orientation="v",
        measure=["relative", "relative", "total"],
        x=["Additions", "Deletions", "Net Change"],
        y=[added, -deleted, final],
        connector={"line": {"color": "grey"}}
    ))
    fig.update_layout(title="Revenue Waterfall", height=350)
    return fig

# ---------------------------------------------------------
# TITLE
# ---------------------------------------------------------

st.title("📊 MG Exceptions — Storytelling Dashboard")
st.markdown("Upload your Additions, Deletions & Changes outputs to visualize the revenue story.")

# ---------------------------------------------------------
# FILE UPLOAD SECTION
# ---------------------------------------------------------

st.header("📁 Upload MG Exception Outputs")

add_file = st.file_uploader("Upload Additions Excel", type=["xlsx"])
del_file = st.file_uploader("Upload Deletions Excel", type=["xlsx"])
chg_file = st.file_uploader("Upload Changes Excel", type=["xlsx"])

if add_file and del_file:
    additions = pd.read_excel(add_file)
    deletions = pd.read_excel(del_file)
    changes = pd.read_excel(chg_file) if chg_file else pd.DataFrame()

    st.success("Files loaded successfully!")

    # Mastergroup + Revenue should exist
    master_col = "Mastergroup name"
    fin_col = "Total Operating Income (HORIS YTD Financials)"

    # ---------------------------------------------------------
    # FEATURE SELECTION
    # ---------------------------------------------------------
    st.header("🎛 Select up to 2 features for storytelling")

    feature_candidates = [c for c in additions.columns if c not in [master_col, fin_col] and additions[c].nunique() < 30]
    selected_features = st.multiselect("Choose features", feature_candidates, max_selections=2)

    if selected_features:
        st.info(f"Storytelling based on: **{', '.join(selected_features)}**")

        # ---------------------------------------------------------
        # FILTERED DATA
        # ---------------------------------------------------------
        df_add = additions[[master_col, fin_col] + selected_features].copy()
        df_del = deletions[[master_col, fin_col] + selected_features].copy()

        revenue_added = df_add[fin_col].sum()
        revenue_deleted = df_del[fin_col].sum()
        net_change = revenue_added - revenue_deleted

        # ---------------------------------------------------------
        # KPI SECTION
        # ---------------------------------------------------------
        st.header("📌 Key Portfolio Metrics")
        col1, col2, col3, col4 = st.columns(4)
        with col1: kpi_card("Total Additions", len(df_add))
        with col2: kpi_card("Revenue Added", f"₹{revenue_added:,.2f}")
        with col3: kpi_card("Total Deletions", len(df_del), color="#cc2222")
        with col4: kpi_card("Revenue Lost", f"₹{revenue_deleted:,.2f}", color="#cc2222")

        st.subheader("Net Impact")
        kpi_card("Net Revenue Change", f"₹{net_change:,.2f}", "#0a8a0a")

        # ---------------------------------------------------------
        # STORY NARRATIVE SECTION
        # ---------------------------------------------------------
        st.header("📝 Executive Summary Story")

        feature_text = " & ".join(selected_features) if len(selected_features) == 2 else selected_features[0]

        story = f"""
### **📌 Portfolio Revenue Story — Based on {feature_text}**

**1. Additions increased revenue by ₹{revenue_added:,.2f}.**  
These accounts entered the portfolio this month and strengthened the revenue base.  
The largest positive contributors were from the **top-performing clusters** based on {feature_text}.

**2. Deletions decreased revenue by ₹{revenue_deleted:,.2f}.**  
These accounts exited or became inactive. Several deletions appear in segments  
with lower engagement or migration to competitors.

**3. Net Revenue Change is ₹{net_change:,.2f}.**  
Overall, the portfolio shows a **{'positive' if net_change>0 else 'negative'} uplift**,  
highlighting the importance of sustaining high-value additions and re-engaging  
key accounts at risk.

"""

        st.markdown(story)

        # ---------------------------------------------------------
        # CHARTS SECTION
        # ---------------------------------------------------------
        st.header("📈 Visual Analysis")

        colA, colB = st.columns(2)

        with colA:
            st.subheader("Additions — Revenue Contribution")
            fig_add = px.bar(df_add.sort_values(fin_col, ascending=False).head(20),
                             x=master_col, y=fin_col, title="Top Additions")
            st.plotly_chart(fig_add, use_container_width=True)

        with colB:
            st.subheader("Deletions — Revenue Loss")
            fig_del = px.bar(df_del.sort_values(fin_col, ascending=False).head(20),
                             x=master_col, y=fin_col, title="Top Deletions")
            st.plotly_chart(fig_del, use_container_width=True)

        st.subheader("Revenue Waterfall")
        fig_water = waterfall_chart(revenue_added, revenue_deleted, net_change)
        st.plotly_chart(fig_water, use_container_width=True)

        # ---------------------------------------------------------
        # TABLE SECTION WITH TABS
        # ---------------------------------------------------------
        st.header("📋 Detailed Data")

        tab1, tab2, tab3 = st.tabs(["Additions", "Deletions", "Changes"])

        with tab1:
            st.dataframe(df_add)
            st.download_button("⬇ Download Additions", 
                               data=download_excel(df_add),
                               file_name="additions_filtered.xlsx")

        with tab2:
            st.dataframe(df_del)
            st.download_button("⬇ Download Deletions", 
                               data=download_excel(df_del),
                               file_name="deletions_filtered.xlsx")

        with tab3:
            if not changes.empty:
                st.dataframe(changes)
                st.download_button("⬇ Download Changes", 
                                   data=download_excel(changes),
                                   file_name="changes.xlsx")
            else:
                st.info("No changes file uploaded.")

        # ---------------------------------------------------------
        # ACTION BUTTONS
        # ---------------------------------------------------------
        st.header("⚙ Actions")

        colx1, colx2, colx3 = st.columns(3)

        with colx1:
            st.button("🔁 Refresh Story")

        with colx2:
            st.button("📤 Export Story as PDF (Coming Soon)")

        with colx3:
            st.button("📈 Compare with Last 6 Months (Future Feature)")

else:
    st.info("Upload at least Additions and Deletions to continue.")






