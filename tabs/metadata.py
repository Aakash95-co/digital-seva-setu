import pandas as pd
from dash import html, dcc, Input, Output
import dash_bootstrap_components as dbc
from app import app
from data import FY_DATA

def get_top10_table(df, group_col, value_col):
    agg_df = df.groupby(group_col)[value_col].sum().reset_index()
    agg_df = agg_df.sort_values(by=value_col, ascending=False).head(10)
    total_val = agg_df[value_col].sum() if agg_df[value_col].sum() > 0 else 1
    agg_df['Percentage'] = (agg_df[value_col] / total_val * 100).round(2).astype(str) + '%'
    display_name = group_col.replace('_Eng', '').capitalize()
    agg_df.rename(columns={group_col: display_name, value_col: 'Count'}, inplace=True)
    return agg_df

def create_dash_table(title, dataframe):
    return html.Div([
        html.H5(title, style={'marginTop': '20px', 'color': '#2d6a9f', 'fontSize': '16px'}),
        dbc.Table.from_dataframe(dataframe, striped=True, bordered=True, hover=True, size="sm")
    ])

def _oot_tile(df):
    oot = int(df['Disposed_Out'].sum()) if 'Disposed_Out' in df.columns else 0
    disp = int(df['Disposed'].sum()) if 'Disposed' in df.columns else 0
    pct = f" ({oot/disp*100:.1f}%)" if disp > 0 else ""
    return f"{oot:,}{pct}"

layout = html.Div([
    html.H2("📋 Dataset Overview & Metadata", style={
        'background': 'linear-gradient(90deg, #1a3c5e 0%, #2d6a9f 100%)',
        'color': 'white', 'padding': '18px 28px', 'borderRadius': '10px', 'marginBottom': '20px'
    }),
    html.Div(id='metadata-dynamic-content'),
], style={"padding": "20px"})

@app.callback(Output('metadata-dynamic-content', 'children'), Input('fy-store', 'data'))
def update_metadata(fy):
    df = FY_DATA[fy]['df']
    fy_label = FY_DATA[fy]['label']

    top_srv_rec = get_top10_table(df, 'Service_Eng', 'Received')
    top_srv_out = get_top10_table(df, 'Service_Eng', 'Disposed_Out')
    top_off_rec = get_top10_table(df, 'Office_Eng', 'Received')
    top_off_out = get_top10_table(df, 'Office_Eng', 'Disposed_Out')
    top_dist_rec = get_top10_table(df, 'District_Eng', 'Received')
    top_dist_out = get_top10_table(df, 'District_Eng', 'Disposed_Out')

    return [
        html.H5(f"Showing data for: {fy_label}", style={'color': '#2d6a9f', 'marginBottom': '12px'}),
        dbc.Row([
            dbc.Col(dbc.Card(dbc.CardBody([html.H5("Districts"),            html.H3(df["District_Eng"].nunique() if "District_Eng" in df.columns else "—")])), md=2),
            dbc.Col(dbc.Card(dbc.CardBody([html.H5("Offices"),              html.H3(df["Office_Eng"].nunique()   if "Office_Eng"   in df.columns else "—")])), md=2),
            dbc.Col(dbc.Card(dbc.CardBody([html.H5("Services"),             html.H3(df["Service_Eng"].nunique()  if "Service_Eng"  in df.columns else "—")])), md=2),
            dbc.Col(dbc.Card(dbc.CardBody([html.H5("Total Received"),       html.H3(f"{int(df['Received'].sum()):,}"   if 'Received'    in df.columns else "—")])), md=2),
            dbc.Col(dbc.Card(dbc.CardBody([html.H5("Total Disposed"),       html.H3(f"{int(df['Disposed'].sum()):,}"   if 'Disposed'    in df.columns else "—")])), md=2),
            dbc.Col(dbc.Card(dbc.CardBody([html.H5("Disposed Out of Time"), html.H3(_oot_tile(df))])), md=2),
        ], className="mb-4"),
        html.H4(f"📊 Top 10 Insights — {fy_label}", style={'marginTop': '30px'}),
        html.Hr(),
        dbc.Row([
            dbc.Col(create_dash_table("1. Top 10 Services (Total Received)", top_srv_rec), md=6),
            dbc.Col(create_dash_table("2. Top 10 Services (Out of Time)", top_srv_out), md=6),
        ], className="mb-4"),
        dbc.Row([
            dbc.Col(create_dash_table("3. Top 10 Offices (Total Received)", top_off_rec), md=6),
            dbc.Col(create_dash_table("4. Top 10 Offices (Out of Time)", top_off_out), md=6),
        ], className="mb-4"),
        dbc.Row([
            dbc.Col(create_dash_table("5. Top 10 Districts (Total Received)", top_dist_rec), md=6),
            dbc.Col(create_dash_table("6. Top 10 Districts (Out of Time)", top_dist_out), md=6),
        ], className="mb-5"),
    ]