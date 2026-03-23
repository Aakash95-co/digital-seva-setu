# from dash import html
#
# layout = html.Div([
#     html.H2("Advanced Analytics and Report", style={'textAlign': 'center', 'marginTop': '40px'}),
#     html.P("This section is currently under development.", style={'textAlign': 'center', 'color': '#6c757d'})
# ], style={"padding": "20px"})


import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression

from app import app
from data import df_adv

# Prepare District Dropdown Options
district_options = [{'label': d, 'value': d} for d in
                    sorted(df_adv['District_name'].dropna().unique())] if not df_adv.empty else []

layout = dbc.Container([
    dbc.Row(dbc.Col(html.H2("🧠 Advanced Analytics & Predictions", className="text-center text-primary my-4"))),

    # Controls Section
    dbc.Card(dbc.CardBody([
        html.H5("Prediction Configuration", className="card-title"),
        dbc.Row([
            dbc.Col([
                html.Label("Select District", className="fw-bold"),
                dcc.Dropdown(id='adv-district-dropdown', options=district_options, placeholder="Select a District...",
                             clearable=False)
            ], md=4),
            dbc.Col([
                html.Label("Target Year", className="fw-bold"),
                dbc.Input(id='adv-year-input', type="number", placeholder="e.g. 2025", min=2022, max=2030, step=1)
            ], md=3),
            dbc.Col([
                html.Label("Target Month", className="fw-bold"),
                dbc.Input(id='adv-month-input', type="number", placeholder="1 to 12", min=1, max=12, step=1)
            ], md=3),
            dbc.Col([
                html.Button("Predict & Analyze", id="adv-predict-btn", n_clicks=0, className="btn btn-primary w-100",
                            style={'marginTop': '24px'})
            ], md=2)
        ])
    ]), className="mb-4 shadow-sm"),

    # Prediction Output
    html.Div(id='adv-prediction-output', className="mb-4"),

    # Insights Cards (Best/Worst)
    dcc.Loading(html.Div(id='adv-insights-output')),

    # Consistency Chart
    dbc.Card(dbc.CardBody([
        html.H5("Office Consistency Analysis (Out-of-Time Fluctuations)", className="card-title"),
        html.P(
            "Shows the variance (Standard Deviation) of out-of-time disposals over months. Lower bars indicate more consistent performance.",
            className="text-muted"),
        dcc.Loading(dcc.Graph(id='adv-consistency-chart'))
    ]), className="mt-4 shadow-sm")

], fluid=True, className="p-4")


@app.callback([Output('adv-prediction-output', 'children'),
               Output('adv-insights-output', 'children'),
               Output('adv-consistency-chart', 'figure')], [Input('adv-predict-btn', 'n_clicks')],
              [State('adv-district-dropdown', 'value'),
               State('adv-year-input', 'value'),
               State('adv-month-input', 'value')]
              )
def run_advanced_analytics(n_clicks, district, year, month):
    if n_clicks == 0 or not district:
        # Default empty state
        return html.Div(), html.Div(), {}

    # 1. Filter Data for selected district
    dist_data = df_adv[df_adv['District_name'] == district].copy()
    if dist_data.empty:
        return dbc.Alert("No data available for the selected district.", color="danger"), html.Div(), {}

    # ==========================================
    # ML PREDICTION (Linear Regression on Time)
    # ==========================================
    ts_data = dist_data.groupby(['Year', 'Month'])['application_Disposed_Out_of_time'].sum().reset_index()
    prediction_alert = dbc.Alert("Provide both Year and Month to view predictions.", color="info")

    if year and month:
        # Create a continuous time index for regression
        ts_data['TimeIndex'] = ts_data['Year'] * 12 + ts_data['Month']
        X = ts_data[['TimeIndex']]
        y = ts_data['application_Disposed_Out_of_time']

        if len(X) >= 3:
            model = LinearRegression()
            model.fit(X, y)
            target_index = year * 12 + month
            pred_val = model.predict([[target_index]])[0]
            pred_val = max(0, int(round(pred_val)))  # Can't have negative applications

            prediction_alert = dbc.Alert([
                html.H4("🔮 Prediction Result", className="alert-heading"),
                html.P(f"Predicted Out-of-Time Disposals for {district} in {month}/{year}: "),
                html.H2(f"{pred_val:,}", className="text-center text-danger fw-bold")
            ], color="light", className="shadow-sm border-danger")
        else:
            prediction_alert = dbc.Alert(
                "Not enough historical data in this district to make a prediction (needs at least 3 months).",
                color="warning")

    # ==========================================
    # BEST / WORST INSIGHTS (Based on Rates)
    # ==========================================
    # We only consider entities with at least 5 total disposals to avoid skewed percentages
    def get_best_worst(df, group_col):
        stats = df.groupby(group_col)[['application_Disposed_Out_of_time', 'application_Disposed']].sum()
        valid_stats = stats[stats['application_Disposed'] >= 5].copy()
        if valid_stats.empty:
            return "N/A", 0, "N/A", 0

        valid_stats['OOT_Rate'] = (valid_stats['application_Disposed_Out_of_time'] / valid_stats[
            'application_Disposed']) * 100

        best_idx = valid_stats['OOT_Rate'].idxmin()
        best_val = valid_stats.loc[best_idx, 'OOT_Rate']

        worst_idx = valid_stats['OOT_Rate'].idxmax()
        worst_val = valid_stats.loc[worst_idx, 'OOT_Rate']

        return best_idx, best_val, worst_idx, worst_val

    best_srv, best_srv_val, worst_srv, worst_srv_val = get_best_worst(dist_data, 'Service_name')
    best_off, best_off_val, worst_off, worst_off_val = get_best_worst(dist_data, 'Office_name')

    def make_card(title, entity, rate, is_best):
        color = "success" if is_best else "danger"
        icon = "🏆" if is_best else "⚠️"
        return dbc.Card([
            dbc.CardHeader(f"{icon} {title}", className=f"text-white bg-{color} fw-bold"),
            dbc.CardBody([
                html.H6(entity, className="card-title text-truncate", title=entity),
                html.P(f"Out-of-Time Rate: {rate:.1f}%", className="card-text fw-bold")
            ])
        ], className="shadow-sm h-100")

    insights_ui = dbc.Row([
        dbc.Col(make_card("Best Service", best_srv, best_srv_val, True), md=3),
        dbc.Col(make_card("Worst Service", worst_srv, worst_srv_val, False), md=3),
        dbc.Col(make_card("Best Office", best_off, best_off_val, True), md=3),
        dbc.Col(make_card("Worst Office", worst_off, worst_off_val, False), md=3),
    ], className="mb-4")

    # ==========================================
    # OFFICE CONSISTENCY (Standard Deviation)
    # ==========================================
    # Group by Office, Year, Month
    monthly_off = dist_data.groupby(['Office_name', 'Year', 'Month'])[
        'application_Disposed_Out_of_time'].sum().reset_index()
    # Calculate Mean and Std Deviation per office
    consistency = monthly_off.groupby('Office_name')['application_Disposed_Out_of_time'].agg(
        ['std', 'sum', 'count']).reset_index()
    # Filter to offices that have at least 3 months of data and some meaningful volume
    consistency = consistency[(consistency['count'] >= 3) & (consistency['sum'] > 5)].fillna(0)

    if consistency.empty:
        fig = {}
    else:
        # Sort by worst standard deviation (highest fluctuation) to display the top 15 most unstable offices
        consistency = consistency.sort_values('std', ascending=False).head(15)

        fig = px.bar(
            consistency,
            x='Office_name',
            y='std',
            color='std',
            color_continuous_scale="Reds",
            title=f"Top 15 Most Fluctuating Offices in {district}",
            labels={'std': 'Standard Deviation (Volatility)', 'Office_name': 'Office Name'},
            template="plotly_white"
        )
        fig.update_layout(xaxis_tickangle=-45, margin=dict(b=100))

    return prediction_alert, insights_ui, fig