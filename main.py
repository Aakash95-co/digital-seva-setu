import dash
from dash import dcc, html, Input, Output
from app import app
from tabs import metadata, monthly_trends, tabular, treemap, advanced_analytics, findings, oot_drilldown

app.layout = html.Div([
    html.H1("📊 Performance Analytics Dashboard", style={'textAlign': 'center', 'marginBottom': '20px'}),

    # FY Selector (shared across tabs)
    dcc.Store(id='fy-store', data='2425'),
    html.Div([
        html.Label("Select Financial Year:", style={'fontWeight': 'bold', 'marginRight': '10px'}),
        dcc.RadioItems(
            id='fy-selector',
            options=[
                {'label': ' FY 2024-25', 'value': '2425'},
                {'label': ' FY 2025-26', 'value': '2526'},
            ],
            value='2425',
            inline=True,
            inputStyle={"marginRight": "5px"},
            labelStyle={"marginRight": "20px"},
        )
    ], style={'textAlign': 'center', 'marginBottom': '20px', 'padding': '10px',
              'background': '#f0f4f8', 'borderRadius': '8px'}),

    dcc.Tabs(id="main-tabs", value="metadata-tab", children=[
        dcc.Tab(label="🗂️ Metadata", value="metadata-tab", children=[metadata.layout]),
        dcc.Tab(label="📅 Monthly Trends", value="trends-tab", children=[monthly_trends.layout]),
        dcc.Tab(label="📋 Tabular Data", value="table-tab", children=[tabular.layout]),
        dcc.Tab(label="📈 Treemap View", value="treemap-tab", children=[treemap.layout]),
        dcc.Tab(label="🧠 Advanced Analytics and Report", value="advanced-analytics-tab", children=[advanced_analytics.layout]),
        dcc.Tab(label="🔍 Findings", value="findings-tab", children=[findings.layout]),          # ← add
        dcc.Tab(label="📌 OOT Drilldown", value="oot-drilldown-tab", children=[oot_drilldown.layout]),  # ← add
    ])
])

@app.callback(Output('fy-store', 'data'), Input('fy-selector', 'value'))
def sync_fy_store(fy): return fy

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8052, debug=True)