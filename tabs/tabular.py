import dash
from dash import dcc, html, dash_table, Input, Output, State
from dash.dash_table.Format import Format
import pandas as pd
from app import app
from data import FY_DATA, initial_summary_tt, get_district_summary, categorize
layout = html.Div([
    html.H3(id='tabular-title', children="📋 Summary Table (Financial Year 2024-2025)"),
    dcc.Download(id="download-summary-table"),
    html.Div([
        html.Button("Download Excel", id="btn-download-summary-table",
                    style={'backgroundColor': '#0d6efd', 'color': 'white', 'border': 'none', 'padding': '8px 16px',
                           'borderRadius': '4px', 'cursor': 'pointer', 'marginBottom': '12px', 'fontWeight': '600'})
    ]),
    dash_table.DataTable(
        id="summary-table",
        columns=[
            {"name": "District", "id": "District_Eng"},
            {"name": "Office", "id": "Office_Eng"},
            {"name": "Service", "id": "Service_Eng"},
            {"name": "Late %", "id": "Late_Disposed_%", "type": "numeric", "format": Format(precision=2, scheme="f")},
            {"name": "Category", "id": "Category"},
            {"name": "Received", "id": "Received"},
        ],
        data=initial_summary_tt.to_dict("records"),
        style_table={"overflowX": "auto"},
        style_cell={"textAlign": "center"},
        style_header={"fontWeight": "bold"},
        row_selectable="single"
    ),
    html.Div("👉 Click a row to expand/collapse", style={"marginTop": "10px"})
], style={"padding": "20px"})

@app.callback(
    Output("summary-table", "data"),
    Output("tabular-title", "children"),   # ← new output
    Input("summary-table", "active_cell"),
    Input("district-filter", "value"),
    Input("office-filter", "value"),
    Input("fy-store", "data"),             # ← new input
    State("summary-table", "data"),
)
def update_summary_table(active_cell, selected_districts, selected_offices, fy, current_data):
    df_tt = FY_DATA[fy]['df_tt']
    title = f"📋 Summary Table ({FY_DATA[fy]['label']})"
    filtered_df = df_tt.copy()
    if selected_districts: filtered_df = filtered_df[filtered_df["District_Eng"].isin(selected_districts)]
    if selected_offices: filtered_df = filtered_df[filtered_df["Office_Eng"].isin(selected_offices)]

    base_summary = get_district_summary(filtered_df)
    if not active_cell or current_data is None: return base_summary.to_dict("records"), title

    current = list(current_data)
    row_idx = active_cell["row"]
    if row_idx >= len(current): return base_summary.to_dict("records"), title
    clicked_row = current[row_idx]

    def is_child_of(row, parent):
        return (row["District_Eng"] == parent["District_Eng"] and (parent["Office_Eng"] == "" or row["Office_Eng"] == parent["Office_Eng"]) and (parent["Service_Eng"] == "" or row["Service_Eng"] == parent["Service_Eng"]) and row != parent)

    children =[r for r in current if is_child_of(r, clicked_row)]
    if children: return [r for r in current if r not in children], title

    if clicked_row["Office_Eng"] == "" and clicked_row["Service_Eng"] == "":
        d = clicked_row["District_Eng"]
        offices = filtered_df[filtered_df["District_Eng"] == d].groupby("Office_Eng", as_index=False).agg({"Disposed_Out": "sum", "Disposed": "sum", "Received": "sum"})
        offices["District_Eng"], offices["Service_Eng"] = d, ""
        offices["Late_Disposed_%"] = (offices["Disposed_Out"] / offices["Disposed"].replace(0, pd.NA) * 100).fillna(0)
        offices["Category"] = offices["Late_Disposed_%"].apply(categorize)
        for i, r in enumerate(offices.to_dict("records")): current.insert(row_idx + 1 + i, r)

    elif clicked_row["Service_Eng"] == "":
        d, o = clicked_row["District_Eng"], clicked_row["Office_Eng"]
        services = filtered_df[(filtered_df["District_Eng"] == d) & (filtered_df["Office_Eng"] == o)].copy()[["District_Eng", "Office_Eng", "Service_Eng", "Disposed_Out", "Disposed", "Received"]]
        services["Late_Disposed_%"] = (services["Disposed_Out"] / services["Disposed"].replace(0, pd.NA) * 100).fillna(0)
        services["Category"] = services["Late_Disposed_%"].apply(categorize)
        for i, r in enumerate(services.to_dict("records")): current.insert(row_idx + 1 + i, r)

    return current, title

@app.callback(
    Output("download-summary-table", "data"),
    Input("btn-download-summary-table", "n_clicks"),
    State("summary-table", "data"),
    prevent_initial_call=True
)
def download_summary_table(n, data):
    if n and data: return dcc.send_data_frame(pd.DataFrame(data).to_excel, "Summary_Table_FY2024_2025.xlsx", sheet_name="Summary", index=False)