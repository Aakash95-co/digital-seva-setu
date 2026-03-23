import dash
from dash import dcc, html, dash_table, Input, Output, State
from dash.dash_table.Format import Format
import pandas as pd
from app import app
from data import df

layout = html.Div([
    # Downloads (Metadata specific)
    dcc.Download(id="download-services"),
    dcc.Download(id="download-services-outoftime"),
    dcc.Download(id="download-office-received"),
    dcc.Download(id="download-office-outoftime"),
    dcc.Download(id="download-district-received"),
    dcc.Download(id="download-district-outoftime"),

    html.Div([
        html.H3("Metadata Summary"),
        html.Div([
            # Left Column
            html.Div([
                html.Div([
                    html.H4("1) Total Services - Count & % (desc)", style={'display': 'inline-block', 'margin-right': '10px'}),
                    html.Button("📥", id="btn-download-services", style={'background': 'none', 'border': 'none', 'cursor': 'pointer', 'font-size': '20px'}, title="Download Excel"),
                ], style={'display': 'flex', 'align-items': 'center'}),
                dash_table.DataTable(
                    id="meta-total-services",
                    columns=[
                        {"name": "Service", "id": "Service_Eng"},
                        {"name": "Count", "id": "Count"},
                        {"name": "Percent", "id": "Percent"},
                    ],
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center"},
                    page_size=10
                ),
                html.Div([
                    html.H4("3) Total Services by Office (Received)", style={'display': 'inline-block', 'margin-right': '10px'}),
                    html.Button("📥", id="btn-download-office-received", style={'background': 'none', 'border': 'none', 'cursor': 'pointer', 'font-size': '20px'}, title="Download Excel"),
                ], style={'display': 'flex', 'align-items': 'center'}),
                dash_table.DataTable(
                    id="meta-total-services-by-office",
                    columns=[
                        {"name": "Office", "id": "Office_Eng"},
                        {"name": "Received", "id": "Received"},
                        {"name": "Percent", "id": "Percent", "type": "numeric", "format": Format(precision=2, scheme="f", symbol_suffix="%")},
                    ],
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center"},
                    page_size=10
                ),
                html.Div([
                    html.H4("5) Total Services by District (Received) - desc", style={'display': 'inline-block', 'margin-right': '10px'}),
                    html.Button("📥", id="btn-download-district-received", style={'background': 'none', 'border': 'none', 'cursor': 'pointer', 'font-size': '20px'}, title="Download Excel"),
                ], style={'display': 'flex', 'align-items': 'center'}),
                dash_table.DataTable(
                    id="meta-total-services-by-district",
                    columns=[
                        {"name": "District", "id": "District_Eng"},
                        {"name": "Received", "id": "Received"},
                        {"name": "Percent", "id": "Percent", "type": "numeric", "format": Format(precision=2, scheme="f", symbol_suffix="%")},
                    ],
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center"},
                    page_size=10
                ),
            ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"}),

            # Right Column
            html.Div([
                html.Div([
                    html.H4("2) Total Services Out of Time - Count & % (desc)", style={'display': 'inline-block', 'margin-right': '10px'}),
                    html.Button("📥", id="btn-download-services-outoftime", style={'background': 'none', 'border': 'none', 'cursor': 'pointer', 'font-size': '20px'}, title="Download Excel"),
                ], style={'display': 'flex', 'align-items': 'center'}),
                dash_table.DataTable(
                    id="meta-total-services-outoftime",
                    columns=[
                        {"name": "Service", "id": "Service_Eng"},
                        {"name": "OutOfTime_Count", "id": "OutOfTime_Count"},
                        {"name": "Percent", "id": "Percent", "type": "numeric", "format": Format(precision=2, scheme="f", symbol_suffix="%")},
                    ],
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center"},
                    page_size=10
                ),
                html.Div([
                    html.H4("4) Total Out of Time by Office", style={'display': 'inline-block', 'margin-right': '10px'}),
                    html.Button("📥", id="btn-download-office-outoftime", style={'background': 'none', 'border': 'none', 'cursor': 'pointer', 'font-size': '20px'}, title="Download Excel"),
                ], style={'display': 'flex', 'align-items': 'center'}),
                dash_table.DataTable(
                    id="meta-total-outoftime-by-office",
                    columns=[
                        {"name": "Office", "id": "Office_Eng"},
                        {"name": "OutOfTime_Count", "id": "OutOfTime_Count"},
                        {"name": "Percent", "id": "Percent", "type": "numeric", "format": Format(precision=2, scheme="f", symbol_suffix="%")},
                    ],
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center"},
                    page_size=10
                ),
                html.Div([
                    html.H4("6) Total Out of Time by District - desc", style={'display': 'inline-block', 'margin-right': '10px'}),
                    html.Button("📥", id="btn-download-district-outoftime", style={'background': 'none', 'border': 'none', 'cursor': 'pointer', 'font-size': '20px'}, title="Download Excel"),
                ], style={'display': 'flex', 'align-items': 'center'}),
                dash_table.DataTable(
                    id="meta-total-outoftime-by-district",
                    columns=[
                        {"name": "District", "id": "District_Eng"},
                        {"name": "OutOfTime_Count", "id": "OutOfTime_Count"},
                        {"name": "Percent", "id": "Percent", "type": "numeric", "format": Format(precision=2, scheme="f", symbol_suffix="%")},
                    ],
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center"},
                    page_size=10
                ),
            ], style={"width": "48%", "display": "inline-block", "marginLeft": "2%", "verticalAlign": "top"}),
        ], style={"display": "flex", "justifyContent": "space-between"}),
    ], style={"padding": "20px"})
])

@app.callback(
    [Output(f"meta-{table_id}", "data") for table_id in[
        "total-services", "total-services-outoftime", "total-services-by-office",
        "total-outoftime-by-office", "total-services-by-district", "total-outoftime-by-district"]],
    Input("main-tabs", "value")
)
def update_metadata_tables(_):
    srv = df.groupby("Service_Eng")[["Received"]].sum().reset_index()
    srv.columns = ["Service_Eng", "Count"]
    srv = srv.sort_values("Count", ascending=False).head(10)
    srv["Percent"] = (srv["Count"] / srv["Count"].sum() * 100).round(2)

    srv_oot = df.groupby("Service_Eng")[["Disposed_Out"]].sum().reset_index()
    srv_oot.columns = ["Service_Eng", "OutOfTime_Count"]
    srv_oot = srv_oot.sort_values("OutOfTime_Count", ascending=False).head(10)
    srv_oot["Percent"] = (srv_oot["OutOfTime_Count"] / srv_oot["OutOfTime_Count"].sum() * 100).round(2)

    office_rcv = df.groupby("Office_Eng")[["Received"]].sum().reset_index()
    office_rcv_total = office_rcv["Received"].sum()
    office_rcv = office_rcv.sort_values("Received", ascending=False).head(10)
    office_rcv["Percent"] = (office_rcv["Received"] / office_rcv_total * 100).round(2)

    office_oot = df.groupby("Office_Eng")[["Disposed_Out"]].sum().reset_index()
    office_oot.columns =["Office_Eng", "OutOfTime_Count"]
    office_oot_total = office_oot["OutOfTime_Count"].sum()
    office_oot = office_oot.sort_values("OutOfTime_Count", ascending=False).head(10)
    office_oot["Percent"] = (office_oot["OutOfTime_Count"] / office_oot_total * 100).round(2)

    district_rcv = df.groupby("District_Eng")[["Received"]].sum().reset_index()
    district_rcv_total = district_rcv["Received"].sum()
    district_rcv = district_rcv.sort_values("Received", ascending=False).head(10)
    district_rcv["Percent"] = (district_rcv["Received"] / district_rcv_total * 100).round(2)

    district_oot = df.groupby("District_Eng")[["Disposed_Out"]].sum().reset_index()
    district_oot.columns = ["District_Eng", "OutOfTime_Count"]
    district_oot_total = district_oot["OutOfTime_Count"].sum()
    district_oot = district_oot.sort_values("OutOfTime_Count", ascending=False).head(10)
    district_oot["Percent"] = (district_oot["OutOfTime_Count"] / district_oot_total * 100).round(2)

    return[
        srv.to_dict('records'), srv_oot.to_dict('records'),
        office_rcv.to_dict('records'), office_oot.to_dict('records'),
        district_rcv.to_dict('records'), district_oot.to_dict('records')
    ]

# Download Callbacks
@app.callback(
    Output("download-services", "data"), Input("btn-download-services", "n_clicks"), State("meta-total-services", "data"), prevent_initial_call=True
)
def download_services(n_clicks, data):
    if data: return dcc.send_data_frame(pd.DataFrame(data).to_excel, "Total_Services.xlsx", sheet_name="Services", index=False)

@app.callback(
    Output("download-services-outoftime", "data"), Input("btn-download-services-outoftime", "n_clicks"), State("meta-total-services-outoftime", "data"), prevent_initial_call=True
)
def download_services_outoftime(n_clicks, data):
    if data: return dcc.send_data_frame(pd.DataFrame(data).to_excel, "Services_OutOfTime.xlsx", sheet_name="Services_OutOfTime", index=False)

@app.callback(
    Output("download-office-received", "data"), Input("btn-download-office-received", "n_clicks"), State("meta-total-services-by-office", "data"), prevent_initial_call=True
)
def download_office_received(n_clicks, data):
    if data: return dcc.send_data_frame(pd.DataFrame(data).to_excel, "Office_Received.xlsx", sheet_name="Office_Received", index=False)

@app.callback(
    Output("download-office-outoftime", "data"), Input("btn-download-office-outoftime", "n_clicks"), State("meta-total-outoftime-by-office", "data"), prevent_initial_call=True
)
def download_office_outoftime(n_clicks, data):
    if data: return dcc.send_data_frame(pd.DataFrame(data).to_excel, "Office_OutOfTime.xlsx", sheet_name="Office_OutOfTime", index=False)

@app.callback(
    Output("download-district-received", "data"), Input("btn-download-district-received", "n_clicks"), State("meta-total-services-by-district", "data"), prevent_initial_call=True
)
def download_district_received(n_clicks, data):
    if data: return dcc.send_data_frame(pd.DataFrame(data).to_excel, "District_Received.xlsx", sheet_name="District_Received", index=False)

@app.callback(
    Output("download-district-outoftime", "data"), Input("btn-download-district-outoftime", "n_clicks"), State("meta-total-outoftime-by-district", "data"), prevent_initial_call=True
)
def download_district_outoftime(n_clicks, data):
    if data: return dcc.send_data_frame(pd.DataFrame(data).to_excel, "District_OutOfTime.xlsx", sheet_name="District_OutOfTime", index=False)