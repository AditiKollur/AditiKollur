```
# app.py
import base64
import io
import json
import pandas as pd
import numpy as np
from dash import Dash, html, dcc, Input, Output, State, dash_table
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------
# Config / Defaults
# ---------------------------
MASTER_COL_DEFAULT = "Mastergroup name"
FIN_COL_DEFAULT = "Total Operating Income (HORIS YTD Financials)"

# ---------------------------
# Helpers
# ---------------------------
def parse_contents(contents, filename):
    """
    Parse uploaded file contents (base64) into a pandas DataFrame.
    Supports XLSX/XLS and CSV.
    """
    if contents is None:
        return None
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    try:
        if filename.lower().endswith(('.xls', '.xlsx')):
            return pd.read_excel(io.BytesIO(decoded))
        elif filename.lower().endswith('.csv'):
            return pd.read_csv(io.StringIO(decoded.decode('utf-8')))
        else:
            return None
    except Exception as e:
        print("Error parsing file:", e)
        return None

def compare_mg_strict(mg_p, mg_c, cust_gbm_p, cust_cmb_p, cust_gbm_c, cust_cmb_c,
                      inscope_col_mg, mastergroup_col=MASTER_COL_DEFAULT, financial_col=FIN_COL_DEFAULT):
    """
    Strict comparison: use only inscope cols to detect additions/deletions/changes.
    Primary key = first inscope_col_mg element.
    Attach financials strictly by mastergroup_col.
    Returns additions_df, deletions_df, changes_df
    """
    inscope = list(inscope_col_mg)
    if len(inscope) == 0:
        raise ValueError("inscope_col_mg must contain at least one column name.")
    primary_key = inscope[0]

    # Validate
    for df, name in [(mg_p, 'mg_p'), (mg_c, 'mg_c')]:
        if df is None:
            raise ValueError(f"{name} is required")
        if primary_key not in df.columns:
            raise ValueError(f"Primary key '{primary_key}' must exist in {name}")

    # Reduce to unique rows on inscope columns
    mg_p_ins = mg_p[inscope].drop_duplicates().copy()
    mg_c_ins = mg_c[inscope].drop_duplicates().copy()

    keys_p = set(mg_p_ins[primary_key].dropna().unique())
    keys_c = set(mg_c_ins[primary_key].dropna().unique())

    added_keys = keys_c - keys_p
    deleted_keys = keys_p - keys_c
    common_keys = keys_p & keys_c

    additions = mg_c[mg_c[primary_key].isin(added_keys)].copy().reset_index(drop=True)
    deletions = mg_p[mg_p[primary_key].isin(deleted_keys)].copy().reset_index(drop=True)

    # Changes: compare only inscope cols excluding primary
    changed_rows = []
    if common_keys:
        merged = mg_p_ins.merge(mg_c_ins, on=primary_key, how='inner', suffixes=('_p','_c'))
        def is_row_changed(row):
            for col in inscope:
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
        merged['is_changed'] = merged.apply(is_row_changed, axis=1)
        changed_keys = merged.loc[merged['is_changed'], primary_key].unique().tolist()
        # build changes side by side (inscope only)
        if changed_keys:
            p_side = mg_p[mg_p[primary_key].isin(changed_keys)][inscope].copy().add_suffix('_p')
            c_side = mg_c[mg_c[primary_key].isin(changed_keys)][inscope].copy().add_suffix('_c')
            changes = p_side.merge(c_side, left_on=f"{primary_key}_p", right_on=f"{primary_key}_c", how='outer', suffixes=('',''))
            # unify primary key column
            changes[primary_key] = changes[f"{primary_key}_c"].combine_first(changes[f"{primary_key}_p"])
            # compute changed_columns
            changed_map = {}
            for _, row in merged[merged['is_changed']].iterrows():
                key = row[primary_key]
                cols_changed = []
                for col in inscope:
                    if col == primary_key:
                        continue
                    try:
                        if not (pd.isna(row[f"{col}_p"]) and pd.isna(row[f"{col}_c"])) and (row[f"{col}_p"] != row[f"{col}_c"]):
                            cols_changed.append(col)
                    except Exception:
                        if str(row[f"{col}_p"]) != str(row[f"{col}_c"]):
                            cols_changed.append(col)
                changed_map[key] = cols_changed
            changes['changed_columns'] = changes[primary_key].map(lambda k: changed_map.get(k, []))
            # attach mastergroup (prefer current)
            mgc_map = mg_c.set_index(primary_key)[mastergroup_col].to_dict() if mastergroup_col in mg_c.columns else {}
            mgp_map = mg_p.set_index(primary_key)[mastergroup_col].to_dict() if mastergroup_col in mg_p.columns else {}
            def choose_master(k):
                return mgc_map.get(k, mgp_map.get(k, None))
            changes[mastergroup_col] = changes[primary_key].map(choose_master)
            # reorder columns
            cols = [primary_key, mastergroup_col] + [c for c in changes.columns if c not in [primary_key, mastergroup_col]]
            changes = changes[cols]
        else:
            changes = pd.DataFrame(columns=[primary_key, mastergroup_col, 'changed_columns'])
    else:
        changes = pd.DataFrame(columns=[primary_key, mastergroup_col, 'changed_columns'])

    # Attach financials strictly by mastergroup_col
    def attach_fin(df_rows, cust_gbm, cust_cmb, current=True):
        if df_rows is None or df_rows.empty:
            return pd.DataFrame() if isinstance(df_rows, pd.DataFrame) else df_rows
        df = df_rows.copy()
        if mastergroup_col not in df.columns:
            # for additions this should be present in mg_c, for deletions in mg_p
            # attempt to bring it from mg frames if missing (best-effort)
            df[mastergroup_col] = None
        # build lookup from gbm then cmb
        gbm_map = {}
        cmb_map = {}
        if cust_gbm is not None and mastergroup_col in cust_gbm.columns and financial_col in cust_gbm.columns:
            gbm_map = cust_gbm.set_index(mastergroup_col)[financial_col].to_dict()
        if cust_cmb is not None and mastergroup_col in cust_cmb.columns and financial_col in cust_cmb.columns:
            cmb_map = cust_cmb.set_index(mastergroup_col)[financial_col].to_dict()
        def lookup(mgname):
            v = gbm_map.get(mgname)
            if v is None or (isinstance(v, float) and np.isnan(v)):
                v = cmb_map.get(mgname)
            try:
                return float(v) if v is not None else np.nan
            except Exception:
                return np.nan
        df[financial_col] = df[mastergroup_col].map(lookup)
        # coerce numeric
        df[financial_col] = pd.to_numeric(df[financial_col], errors='coerce')
        return df

    additions = additions if not additions.empty else pd.DataFrame(columns=mg_c.columns)
    deletions = deletions if not deletions.empty else pd.DataFrame(columns=mg_p.columns)

    additions = attach_fin(additions, cust_gbm_c, cust_cmb_c)
    changes = attach_fin(changes, cust_gbm_c, cust_cmb_c) if not changes.empty else changes
    deletions = attach_fin(deletions, cust_gbm_p, cust_cmb_p)

    # ensure financial col numeric
    for df in (additions, changes, deletions):
        if isinstance(df, pd.DataFrame):
            if financial_col not in df.columns:
                df[financial_col] = np.nan
            df[financial_col] = pd.to_numeric(df[financial_col], errors='coerce')

    # Sort descending by financial_col
    additions = additions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    deletions = deletions.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)
    if isinstance(changes, pd.DataFrame) and not changes.empty:
        changes = changes.sort_values(by=financial_col, ascending=False, na_position='last').reset_index(drop=True)

    return additions, deletions, changes

def make_waterfall(added, deleted, net):
    fig = go.Figure(go.Waterfall(
        name = "Revenue Movement",
        orientation = "v",
        measure = ["relative", "relative", "total"],
        x = ["Additions", "Deletions", "Net Change"],
        y = [added, -deleted, net],
        connector = {"line":{"color":"grey"}}
    ))
    fig.update_layout(title_text="Revenue Waterfall", height=400, margin=dict(l=40,r=20,t=50,b=20))
    return fig

# ---------------------------
# Dash App Layout
# ---------------------------
app = Dash(__name__)
app.title = "MG Exceptions — Storytelling (Dash)"

app.layout = html.Div([
    html.H2("MG Exceptions — Storytelling Dashboard (Dash)"),
    html.Div("Upload required files: mg_p, mg_c, cust_gbm_p, cust_cmb_p, cust_gbm_c, cust_cmb_c"),
    html.Div([
        html.Div([
            html.Label("mg_p (prior month)"),
            dcc.Upload(id='upload-mg-p', children=html.Button("Upload mg_p"), multiple=False),
            html.Div(id='mg-p-fname', style={'fontSize':12, 'color':'gray'})
        ], style={'display':'inline-block','marginRight':20}),
        html.Div([
            html.Label("mg_c (current month)"),
            dcc.Upload(id='upload-mg-c', children=html.Button("Upload mg_c"), multiple=False),
            html.Div(id='mg-c-fname', style={'fontSize':12, 'color':'gray'})
        ], style={'display':'inline-block','marginRight':20}),
    ]),
    html.Br(),
    html.Div([
        html.Div([
            html.Label("cust_gbm_p"),
            dcc.Upload(id='upload-cust-gbm-p', children=html.Button("Upload cust_gbm_p"), multiple=False),
            html.Div(id='cust-gbm-p-fname', style={'fontSize':12, 'color':'gray'})
        ], style={'display':'inline-block','marginRight':20}),
        html.Div([
            html.Label("cust_cmb_p"),
            dcc.Upload(id='upload-cust-cmb-p', children=html.Button("Upload cust_cmb_p"), multiple=False),
            html.Div(id='cust-cmb-p-fname', style={'fontSize':12, 'color':'gray'})
        ], style={'display':'inline-block','marginRight':20}),
    ]),
    html.Br(),
    html.Div([
        html.Div([
            html.Label("cust_gbm_c"),
            dcc.Upload(id='upload-cust-gbm-c', children=html.Button("Upload cust_gbm_c"), multiple=False),
            html.Div(id='cust-gbm-c-fname', style={'fontSize':12, 'color':'gray'})
        ], style={'display':'inline-block','marginRight':20}),
        html.Div([
            html.Label("cust_cmb_c"),
            dcc.Upload(id='upload-cust-cmb-c', children=html.Button("Upload cust_cmb_c"), multiple=False),
            html.Div(id='cust-cmb-c-fname', style={'fontSize':12, 'color':'gray'})
        ], style={'display':'inline-block','marginRight':20}),
    ]),
    html.Hr(),
    html.Div([
        html.Label("Inscope columns (comma-separated). First = primary key"),
        dcc.Input(id='inscope-input', type='text', style={'width':'70%'}, value="PS ID,Mastergroup name,Segment"),
    ]),
    html.Br(),
    html.Div([
        html.Label("Mastergroup column name"),
        dcc.Input(id='master-col', type='text', value=MASTER_COL_DEFAULT),
        html.Label(" Financial column name", style={'marginLeft':20}),
        dcc.Input(id='fin-col', type='text', value=FIN_COL_DEFAULT, style={'width':'40%'}),
    ], style={'marginTop':10}),
    html.Br(),
    html.Button("Run Comparison", id='run-button'),
    html.Div(id='run-error', style={'color':'red'}),
    html.Hr(),
    html.Div(id='summary-area'),
    html.Hr(),
    html.Div([
        html.Div([
            html.Label("Select up to 2 features for storytelling (will be populated after uploads)"),
            dcc.Dropdown(id='feature-dropdown', multi=True, value=[], placeholder="Select features"),
        ], style={'width':'50%','display':'inline-block'}),
        html.Div(id='feature-warning', style={'color':'red','marginLeft':20})
    ]),
    html.Br(),
    html.Div(id='kpi-cards', style={'display':'flex','gap':'10px'}),
    html.Br(),
    html.Div([
        dcc.Graph(id='additions-bar', style={'width':'48%','display':'inline-block'}),
        dcc.Graph(id='deletions-bar', style={'width':'48%','display':'inline-block'})
    ]),
    html.Br(),
    dcc.Graph(id='waterfall'),
    html.Hr(),
    html.H4("Detailed Tables"),
    dcc.Tabs(id='tabs', value='tab-add', children=[
        dcc.Tab(label='Additions', value='tab-add'),
        dcc.Tab(label='Deletions', value='tab-del'),
        dcc.Tab(label='Changes', value='tab-chg'),
    ]),
    html.Div(id='tab-content'),
    html.Br(),
    html.Div([
        html.Button("Download Additions (xlsx)", id='dl-add'),
        html.Button("Download Deletions (xlsx)", id='dl-del'),
        html.Button("Download Changes (xlsx)", id='dl-chg'),
        html.Div(id='download-link', style={'display':'inline-block','marginLeft':20})
    ]),
    # hidden store to keep dataframes between callbacks
    dcc.Store(id='store-mg-p'),
    dcc.Store(id='store-mg-c'),
    dcc.Store(id='store-cust-gbm-p'),
    dcc.Store(id='store-cust-cmb-p'),
    dcc.Store(id='store-cust-gbm-c'),
    dcc.Store(id='store-cust-cmb-c'),
    dcc.Store(id='store-additions'),
    dcc.Store(id='store-deletions'),
    dcc.Store(id='store-changes'),
], style={'margin':20, 'fontFamily':'Arial, sans-serif'})

# ---------------------------
# Callbacks: upload handlers to store the dataframes in json
# ---------------------------
def upload_to_store(contents, filename):
    if contents is None:
        return None
    df = parse_contents(contents, filename)
    if df is None:
        return None
    return df.to_json(date_format='iso', orient='records')

@app.callback(Output('mg-p-fname','children'),
              Output('store-mg-p','data'),
              Input('upload-mg-p','contents'),
              State('upload-mg-p','filename'))
def handle_mg_p(contents, filename):
    if contents:
        df_json = upload_to_store(contents, filename)
        return f"{filename}", df_json
    return "", None

@app.callback(Output('mg-c-fname','children'),
              Output('store-mg-c','data'),
              Input('upload-mg-c','contents'),
              State('upload-mg-c','filename'))
def handle_mg_c(contents, filename):
    if contents:
        df_json = upload_to_store(contents, filename)
        return f"{filename}", df_json
    return "", None

@app.callback(Output('cust-gbm-p-fname','children'),
              Output('store-cust-gbm-p','data'),
              Input('upload-cust-gbm-p','contents'),
              State('upload-cust-gbm-p','filename'))
def handle_cust_gbm_p(contents, filename):
    if contents:
        df_json = upload_to_store(contents, filename)
        return f"{filename}", df_json
    return "", None

@app.callback(Output('cust-cmb-p-fname','children'),
              Output('store-cust-cmb-p','data'),
              Input('upload-cust-cmb-p','contents'),
              State('upload-cust-cmb-p','filename'))
def handle_cust_cmb_p(contents, filename):
    if contents:
        df_json = upload_to_store(contents, filename)
        return f"{filename}", df_json
    return "", None

@app.callback(Output('cust-gbm-c-fname','children'),
              Output('store-cust-gbm-c','data'),
              Input('upload-cust-gbm-c','contents'),
              State('upload-cust-gbm-c','filename'))
def handle_cust_gbm_c(contents, filename):
    if contents:
        df_json = upload_to_store(contents, filename)
        return f"{filename}", df_json
    return "", None

@app.callback(Output('cust-cmb-c-fname','children'),
              Output('store-cust-cmb-c','data'),
              Input('upload-cust-cmb-c','contents'),
              State('upload-cust-cmb-c','filename'))
def handle_cust_cmb_c(contents, filename):
    if contents:
        df_json = upload_to_store(contents, filename)
        return f"{filename}", df_json
    return "", None

# ---------------------------
# Run Comparison callback
# ---------------------------
@app.callback(
    Output('run-error','children'),
    Output('store-additions','data'),
    Output('store-deletions','data'),
    Output('store-changes','data'),
    Output('feature-dropdown','options'),
    Input('run-button','n_clicks'),
    State('store-mg-p','data'),
    State('store-mg-c','data'),
    State('store-cust-gbm-p','data'),
    State('store-cust-cmb-p','data'),
    State('store-cust-gbm-c','data'),
    State('store-cust-cmb-c','data'),
    State('inscope-input','value'),
    State('master-col','value'),
    State('fin-col','value')
)
def run_comparison(nclicks, mgp_json, mgc_json, gtmp_json, ctpm_json, gtc_json, ctcm_json, inscope_txt, master_col, fin_col):
    if not nclicks:
        return "", None, None, None, []
    try:
        # load dfs
        mg_p = pd.read_json(mgp_json, orient='records') if mgp_json else None
        mg_c = pd.read_json(mgc_json, orient='records') if mgc_json else None
        cust_gbm_p = pd.read_json(gtmp_json, orient='records') if gtmp_json else pd.DataFrame()
        cust_cmb_p = pd.read_json(ctpm_json, orient='records') if ctpm_json else pd.DataFrame()
        cust_gbm_c = pd.read_json(gtc_json, orient='records') if gtc_json else pd.DataFrame()
        cust_cmb_c = pd.read_json(ctcm_json, orient='records') if ctcm_json else pd.DataFrame()

        inscope_cols = [s.strip() for s in inscope_txt.split(',') if s.strip()]

        adds, dels, chgs = compare_mg_strict(mg_p, mg_c, cust_gbm_p, cust_cmb_p, cust_gbm_c, cust_cmb_c,
                                             inscope_cols, mastergroup_col=master_col, financial_col=fin_col)

        # prepare feature dropdown candidates: pick columns from additions/deletions excluding master and fin col,
        # but only include columns with reasonable cardinality
        samp = adds if not adds.empty else (dels if not dels.empty else (chgs if isinstance(chgs,pd.DataFrame) and not chgs.empty else None))
        options = []
        if samp is not None and not samp.empty:
            candidate_cols = [c for c in samp.columns if c not in [master_col, fin_col]]
            for c in candidate_cols:
                # include columns with < 100 unique values to keep UI manageable
                if samp[c].nunique() < 200:
                    options.append({'label': c, 'value': c})

        # store results as json
        adds_json = adds.to_json(orient='records', date_format='iso')
        dels_json = dels.to_json(orient='records', date_format='iso')
        chgs_json = chgs.to_json(orient='records', date_format='iso') if isinstance(chgs,pd.DataFrame) else json.dumps(chgs)

        return "", adds_json, dels_json, chgs_json, options
    except Exception as e:
        return f"Error: {str(e)}", None, None, None, []

# ---------------------------
# KPI & Charts callbacks
# ---------------------------
@app.callback(
    Output('kpi-cards','children'),
    Output('additions-bar','figure'),
    Output('deletions-bar','figure'),
    Output('waterfall','figure'),
    Input('store-additions','data'),
    Input('store-deletions','data'),
    Input('store-changes','data'),
    Input('feature-dropdown','value'),
    State('master-col','value'),
    State('fin-col','value')
)
def update_visuals(add_json, del_json, chg_json, selected_features, master_col, fin_col):
    adds = pd.read_json(add_json, orient='records') if add_json else pd.DataFrame()
    dels = pd.read_json(del_json, orient='records') if del_json else pd.DataFrame()
    chgs = pd.read_json(chg_json, orient='records') if chg_json else pd.DataFrame()

    # basic KPIs
    rev_added = float(adds[fin_col].sum()) if not adds.empty and fin_col in adds.columns else 0.0
    rev_deleted = f
