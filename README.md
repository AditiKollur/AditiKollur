```

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
    rev_deleted = float(dels[fin_col].sum()) if not dels.empty and fin_col in dels.columns else 0.0
    net = rev_added - rev_deleted

    kpi_nodes = []
    def make_card(title, value, subtitle=""):
        return html.Div([
            html.Div(title, style={'fontSize':12, 'color':'#666'}),
            html.Div(f"{value}", style={'fontSize':20, 'fontWeight':'bold'}),
            html.Div(subtitle, style={'fontSize':11, 'color':'#999'})
        ], style={'padding':10, 'border':'1px solid #eee', 'borderRadius':6, 'width':220})

    kpi_nodes.append(make_card("Total Additions", len(adds), "count"))
    kpi_nodes.append(make_card("Revenue Added", f"₹{rev_added:,.2f}", "sum of financials"))
    kpi_nodes.append(make_card("Total Deletions", len(dels), "count"))
    kpi_nodes.append(make_card("Revenue Lost", f"₹{rev_deleted:,.2f}", "sum of financials"))
    kpi_nodes.append(make_card("Net Change", f"₹{net:,.2f}", ""))

    # Top additions bar
    if not adds.empty and fin_col in adds.columns:
        top_add = adds.sort_values(by=fin_col, ascending=False).head(20)
        fig_add = px.bar(top_add, x=master_col, y=fin_col, title="Top Additions by Revenue")
    else:
        fig_add = go.Figure()
        fig_add.update_layout(title="Top Additions by Revenue")

    # Top deletions bar
    if not dels.empty and fin_col in dels.columns:
        top_del = dels.sort_values(by=fin_col, ascending=False).head(20)
        fig_del = px.bar(top_del, x=master_col, y=fin_col, title="Top Deletions by Revenue")
    else:
        fig_del = go.Figure()
        fig_del.update_layout(title="Top Deletions by Revenue")

    # Waterfall
    fig_water = make_waterfall(rev_added, rev_deleted, net)

    return kpi_nodes, fig_add, fig_del, fig_water

# ---------------------------
# Tab content & table callbacks
# ---------------------------
@app.callback(
    Output('tab-content','children'),
    Input('tabs','value'),
    State('store-additions','data'),
    State('store-deletions','data'),
    State('store-changes','data'),
    State('master-col','value'),
    State('fin-col','value')
)
def render_tab(tab, add_json, del_json, chg_json, master_col, fin_col):
    adds = pd.read_json(add_json, orient='records') if add_json else pd.DataFrame()
    dels = pd.read_json(del_json, orient='records') if del_json else pd.DataFrame()
    chgs = pd.read_json(chg_json, orient='records') if chg_json else pd.DataFrame()

    if tab == 'tab-add':
        if adds.empty:
            return html.Div("No additions found.")
        return dash_table.DataTable(
            data=adds.to_dict('records'),
            columns=[{"name":c,"id":c} for c in adds.columns],
            page_size=15,
            style_table={'overflowX':'auto'},
        )
    elif tab == 'tab-del':
        if dels.empty:
            return html.Div("No deletions found.")
        return dash_table.DataTable(
            data=dels.to_dict('records'),
            columns=[{"name":c,"id":c} for c in dels.columns],
            page_size=15,
            style_table={'overflowX':'auto'},
        )
    else:
        if chgs.empty:
            return html.Div("No changes found.")
        # keep important columns and changed_columns
        cols = [c for c in chgs.columns]
        return dash_table.DataTable(
            data=chgs.to_dict('records'),
            columns=[{"name":c,"id":c} for c in cols],
            page_size=15,
            style_table={'overflowX':'auto'},
        )

# ---------------------------
# Download buttons (basic implementation)
# ---------------------------
def df_to_xlsx_bytes(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return out.getvalue()

@app.callback(
    Output('download-link','children'),
    Input('dl-add','n_clicks'),
    Input('dl-del','n_clicks'),
    Input('dl-chg','n_clicks'),
    State('store-additions','data'),
    State('store-deletions','data'),
    State('store-changes','data'),
    State('fin-col','value'),
    prevent_initial_call=True
)
def handle_download(n1, n2, n3, add_json, del_json, chg_json, fin_col):
    ctx = dash.callback_context
    if not ctx.triggered:
        return ""
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    if button_id == 'dl-add':
        df = pd.read_json(add_json, orient='records') if add_json else pd.DataFrame()
        b = df_to_xlsx_bytes(df)
        return html.A("Download Additions", href="data:application/octet-stream;base64," + base64.b64encode(b).decode(), download="additions.xlsx")
    elif button_id == 'dl-del':
        df = pd.read_json(del_json, orient='records') if del_json else pd.DataFrame()
        b = df_to_xlsx_bytes(df)
        return html.A("Download Deletions", href="data:application/octet-stream;base64," + base64.b64encode(b).decode(), download="deletions.xlsx")
    else:
        # changes
        df = pd.read_json(chg_json, orient='records') if chg_json else pd.DataFrame()
        b = df_to_xlsx_bytes(df)
        return html.A("Download Changes", href="data:application/octet-stream;base64," + base64.b64encode(b).decode(), download="changes.xlsx")

# ---------------------------
# Run
# ---------------------------
if __name__ == "__main__":
    app.run_server(debug=True, port=8050)

How it works (quick)

1. Upload the six Excel files using the
