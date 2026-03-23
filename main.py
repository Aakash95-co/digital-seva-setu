import dash
from dash import dcc, html
from app import app

# Import all independent layouts/callbacks
from tabs import metadata, monthly_trends, tabular, treemap, advanced_analytics

app.layout = html.Div([
    html.H1("📊 Performance Analytics Dashboard", style={'textAlign': 'center', 'marginBottom': '30px'}),

    dcc.Tabs(id="main-tabs", value="metadata-tab", children=[
        dcc.Tab(label="🗂️ Metadata", value="metadata-tab", children=[metadata.layout]),
        dcc.Tab(label="📅 Monthly Trends", value="trends-tab", children=[monthly_trends.layout]),
        dcc.Tab(label="📋 Tabular Data", value="table-tab", children=[tabular.layout]),
        dcc.Tab(label="📈 Treemap View", value="treemap-tab", children=[treemap.layout]),
        dcc.Tab(label="🧠 Advanced Analytics and Report", value="advanced-analytics-tab", children=[advanced_analytics.layout]),
    ])
])

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8052, debug=True)