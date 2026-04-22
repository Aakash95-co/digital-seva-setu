import dash
from dash import dcc, html, Input, Output
import plotly.express as px
from app import app
from data import FY_DATA, hierarchies, district_order_2425

layout = html.Div([
    html.H2("📊 Treemap Drill-down"),
    html.Div([
        html.Div([
            html.Label("Select Hierarchy:"),
            dcc.Dropdown(
                id="hierarchy-choice",
                options=[{"label": k, "value": k} for k in hierarchies.keys()],
                value="District → Office → Service", clearable=False
            )
        ], style={"width": "30%", "margin-right": "20px"}),
        html.Div([
            html.Label("Filter by District:"),
            dcc.Dropdown(id="district-filter", options=[{"label": d, "value": d} for d in district_order_2425], multi=True,
                         placeholder="Select District(s)"),
            html.Div([
                html.Button("Top 3 Districts", id="top3-btn", n_clicks=0, style={"margin-right": "10px"}),
                html.Button("Bottom 3 Districts", id="bottom3-btn", n_clicks=0)
            ], style={"margin-top": "5px"})
        ], style={"width": "35%", "margin-right": "20px"}),
        html.Div([
            html.Label("Filter by Office:"),
            dcc.Dropdown(id="office-filter", multi=True, placeholder="Select Office(s)")
        ], style={"width": "35%"})
    ], style={"display": "flex", "align-items": "flex-start", "margin-bottom": "20px"}),
    dcc.Graph(id="treemap", style={"height": "90vh"}),
], style={"padding": "20px"})


@app.callback(Output("office-filter", "options"), Input("district-filter", "value"), Input("fy-store", "data"))
def update_office_options(selected_districts, fy):
    df_tt = FY_DATA[fy]['df_tt']
    filtered_df = df_tt.copy()
    if selected_districts: filtered_df = filtered_df[filtered_df["District_Eng"].isin(selected_districts)]
    office_order = filtered_df.groupby("Office_Eng")["Disposed_Out"].sum().sort_values(ascending=False).index.tolist()
    return [{"label": o, "value": o} for o in office_order]


@app.callback(
    Output("treemap", "figure"),
    Input("district-filter", "value"), Input("office-filter", "value"),
    Input("hierarchy-choice", "value"), Input("top3-btn", "n_clicks"), Input("bottom3-btn", "n_clicks"), Input("fy-store", "data")
)
def update_treemap(selected_districts, selected_offices, hierarchy_choice, top3_clicks, bottom3_clicks, fy):
    df_tt = FY_DATA[fy]['df_tt']
    filtered_df = df_tt.copy()
    ctx = dash.callback_context
    if ctx.triggered:
        trigger_id = ctx.triggered[0]["prop_id"].split(".")[0]
        if trigger_id == "top3-btn":
            filtered_df = filtered_df[filtered_df["District_Eng"].isin(
                df_tt.groupby("District_Eng")["Disposed_Out"].sum().nlargest(3).index.tolist())]
        elif trigger_id == "bottom3-btn":
            filtered_df = filtered_df[filtered_df["District_Eng"].isin(
                df_tt.groupby("District_Eng")["Disposed_Out"].sum().nsmallest(3).index.tolist())]
        elif selected_districts:
            filtered_df = filtered_df[filtered_df["District_Eng"].isin(selected_districts)]
    elif selected_districts:
        filtered_df = filtered_df[filtered_df["District_Eng"].isin(selected_districts)]

    if selected_offices: filtered_df = filtered_df[filtered_df["Office_Eng"].isin(selected_offices)]

    path = hierarchies[hierarchy_choice]
    custom_cols = ["Disposed_Out", "Disposed", "Total", "Pending", "Received", "Late_Disposed_%"]

    fig = px.treemap(
        filtered_df, path=path, values="Disposed_Out", color="Late_Disposed_%",
        color_continuous_scale="Reds", custom_data=filtered_df[custom_cols]
    )
    fig.update_traces(
        texttemplate="%{label}<br>Out of Time=%{customdata[0]}<br>Total=%{customdata[2]}<br>Pending=%{customdata[3]}<br>Received=%{customdata[4]}<br>Out of Time %=%{customdata[5]:.2f}%",
        textinfo="label+value"
    )
    return fig


@app.callback(
    Output("district-filter", "options"),
    Output("district-filter", "value"),
    Input("fy-store", "data")
)
def update_district_options(fy):
    opts = [{"label": d, "value": d} for d in FY_DATA[fy]['district_order']]
    return opts, None