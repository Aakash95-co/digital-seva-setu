import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from app import app
from data import FY_DATA, all_months_2425, month_options_2425, COLOR_PALETTE
layout = dbc.Container([
    dbc.Row(dbc.Col(html.H1("Service Performance Analytics Dashboard", className="text-center text-primary my-4"))),
    dbc.Card(dbc.CardBody([
        dbc.Row([
            dbc.Col([
                html.Label("1. Select Analysis Mode", className="fw-bold"),
                dbc.RadioItems(
                    id='analysis-mode-selector',
                    options=[{'label': 'Single Entity', 'value': 'single'},
                             {'label': 'Comparison', 'value': 'comparison'}],
                    value='single', inline=True
                )
            ], md=6),
            dbc.Col([
                html.Label("2. Select Time Period", className="fw-bold"),
                dcc.Dropdown(id='month-dropdown', options=month_options_2425,
             value=[all_months_2425[-1]] if all_months_2425 else [], multi=True)
            ], md=6)
        ], className="mb-3"),
        html.Hr(),
        dbc.Row([
            dbc.Col([
                html.Label("3. Select Primary Level", className="fw-bold"),
                dcc.Dropdown(id='primary-level',
                             options=[{'label': 'District', 'value': 'district'},
                                      {'label': 'Service', 'value': 'service'},
                                      {'label': 'Office', 'value': 'office'}],
                             value='district', clearable=False)
            ], md=4),
            dbc.Col([
                html.Div(id='single-entity-wrapper', style={'display': 'none'}, children=[
                    html.Label("Select Entity to Analyze", className="fw-bold"),
                    dcc.Dropdown(id='single-entity-dropdown', placeholder="Select Entity...")
                ]),
                html.Div(id='comparison-entity-wrapper', style={'display': 'block'}, children=[
                    dbc.Row([
                        dbc.Col(html.Label("Select First Entity", className="fw-bold")),
                        dbc.Col(html.Label("Select Second Entity", className="fw-bold"))
                    ]),
                    dbc.Row([
                        dbc.Col(dcc.Dropdown(id='entity1-dropdown', placeholder="First Entity...")),
                        dbc.Col(dcc.Dropdown(id='entity2-dropdown', placeholder="Second Entity..."))
                    ])
                ])
            ], md=8),
        ], className="mb-3"),
        html.Hr(),
        dbc.Row([
            dbc.Col([
                html.Label("4. (Optional) Drill-Down", className="fw-bold"),
                dbc.Checklist(id='drill-down-levels', value=[], inline=True),
                html.Div(id='drilldown-entity-dropdown-wrapper', children=[
                    dcc.Dropdown(id='drilldown-service-dropdown', options=[], value=None, multi=False,
                                 placeholder="Select Service...", style={'display': 'none'}, className="mt-2"),
                    dcc.Dropdown(id='drilldown-office-dropdown', options=[], value=None, multi=False,
                                 placeholder="Select Office...", style={'display': 'none'}, className="mt-2"),
                    dcc.Dropdown(id='drilldown-district-dropdown', options=[], value=None, multi=False,
                                 placeholder="Select District...", style={'display': 'none'}, className="mt-2"),
                ], className="mt-2")
            ], md=12),
        ]),
    ]), className="mb-4"),

    html.Div(id='analysis-status'),
    html.Div(id='performance-summary', className="mb-4"),
    dbc.Card(dbc.CardBody(dcc.Loading(html.Div(id='main-visualizations'))), className="mb-4"),
    dbc.Card(dbc.CardBody(dcc.Loading(html.Div(id='detailed-breakdown')))),
], fluid=True, className="bg-light p-4")


@app.callback(Output('month-dropdown', 'value'), Input('month-dropdown', 'value'))
def select_all_months(selected):
    if selected is None: return [all_months[-1]] if all_months else []
    if 'ALL_MONTHS' in selected: return all_months
    return selected


@app.callback([Output('single-entity-wrapper', 'style'), Output('comparison-entity-wrapper', 'style')],
              Input('analysis-mode-selector', 'value')
              )
def toggle_entity_selectors(mode):
    if mode == 'single': return {'display': 'block'}, {'display': 'none'}
    return {'display': 'none'}, {'display': 'block'}


@app.callback([Output('single-entity-dropdown', 'options'), Output('entity1-dropdown', 'options'),
               Output('entity2-dropdown', 'options')],
              [Input('primary-level', 'value'), Input('month-dropdown', 'value'), Input('fy-store', 'data')]
              )
def update_entity_options(primary_level, selected_months, fy):
    df_mt = FY_DATA[fy]['df_mt']
    if not selected_months: return [], [], []
    col_map = {'district': 'District_name', 'service': 'Service_name', 'office': 'Office_name'}
    entities = sorted(df_mt[df_mt['Month_Year'].isin(selected_months)][col_map[primary_level]].unique())
    options = [{'label': e, 'value': e} for e in entities]
    return options, options, options


@app.callback(Output('drill-down-levels', 'options'), Input('primary-level', 'value'))
def update_drill_down_options(primary_level):
    if primary_level == 'district': return [{'label': 'By Service', 'value': 'service'},
                                            {'label': 'By Office', 'value': 'office'}]
    if primary_level == 'service': return [{'label': 'By District', 'value': 'district'},
                                           {'label': 'By Office', 'value': 'office'}]
    if primary_level == 'office': return [{'label': 'By District', 'value': 'district'},
                                          {'label': 'By Service', 'value': 'service'}]
    return []


@app.callback(Output('drill-down-levels', 'value'), Input('primary-level', 'value'))
def reset_drill_levels_on_primary_change(_): return []


@app.callback([Output('drilldown-service-dropdown', 'options'), Output('drilldown-service-dropdown', 'value'),
               Output('drilldown-service-dropdown', 'style'),
               Output('drilldown-office-dropdown', 'options'), Output('drilldown-office-dropdown', 'value'),
               Output('drilldown-office-dropdown', 'style'),
               Output('drilldown-district-dropdown', 'options'), Output('drilldown-district-dropdown', 'value'),
               Output('drilldown-district-dropdown', 'style')],
              [Input('drill-down-levels', 'value'), Input('primary-level', 'value'), Input('month-dropdown', 'value'),
               Input('analysis-mode-selector', 'value'), Input('single-entity-dropdown', 'value'),
               Input('entity1-dropdown', 'value'), Input('entity2-dropdown', 'value'),
               Input('fy-store', 'data')],
              [State('drilldown-service-dropdown', 'value'), State('drilldown-office-dropdown', 'value'),
               State('drilldown-district-dropdown', 'value')]
              )
def populate_drilldown_dropdowns(drill_levels, primary_level, months, mode, single_entity, entity1, entity2,
                                 fy,
                                 current_service, current_office, current_district):
    hide = {'display': 'none'}
    if not months: return (dash.no_update,) * 9
    df_mt_light = FY_DATA[fy]['df_mt_light']
    col_map = {'district': 'District_name', 'service': 'Service_name', 'office': 'Office_name'}
    primary_col = col_map[primary_level]
    df_base = df_mt_light[df_mt_light['Month_Year'].isin(months)]
    if mode == 'single' and single_entity:
        df_base = df_base[df_base[primary_col] == single_entity]
    elif mode == 'comparison' and entity1 and entity2:
        df_base = df_base[df_base[primary_col].isin([entity1, entity2])]

    show_service, show_office, show_district = 'service' in (drill_levels or []), 'office' in (
                drill_levels or []), 'district' in (drill_levels or [])
    specific_service, specific_office, specific_district = bool(
        current_service and not str(current_service).startswith('ALL_')), bool(
        current_office and not str(current_office).startswith('ALL_')), bool(
        current_district and not str(current_district).startswith('ALL_'))

    def oot_counts_sorted(dataframe, col_name):
        if col_name not in dataframe.columns: return pd.Series(dtype='int64')
        s = dataframe.groupby(col_name)['application_Disposed_Out_of_time'].sum().sort_values(ascending=False)
        return s[s.index.map(lambda x: isinstance(x, str) and x.strip() != '')]

    def build_opts(counts_series, all_label, include_counts=False):
        opts = [{'label': f'All {all_label}', 'value': f'ALL_{all_label.upper().replace(" ", "_")}'}]
        for v, c in counts_series.items():
            opts.append({'label': f"{v} ({int(c):,})" if include_counts else v, 'value': v})
        return opts

    service_opts = build_opts(oot_counts_sorted(df_base, 'Service_name'), 'Service')
    office_opts = build_opts(oot_counts_sorted(df_base, 'Office_name'), 'Office')
    district_opts = build_opts(oot_counts_sorted(df_base, 'District_name'), 'District')

    trigger_ids = [t['prop_id'].split('.')[0] for t in dash.callback_context.triggered]
    reset = any(t in ('primary-level', 'month-dropdown', 'analysis-mode-selector') for t in trigger_ids)

    def resolve(val, valid):
        return val if val and (val in set(valid) or str(val).startswith('ALL_')) else None

    svc_val = None if reset else resolve(current_service, oot_counts_sorted(df_base, 'Service_name').index)
    off_val = None if reset else resolve(current_office, oot_counts_sorted(df_base, 'Office_name').index)
    dist_val = None if reset else resolve(current_district, oot_counts_sorted(df_base, 'District_name').index)

    return (
        service_opts if show_service else dash.no_update,
        svc_val if show_service else (None if reset else dash.no_update), {} if show_service else hide,
        office_opts if show_office else dash.no_update, off_val if show_office else (None if reset else dash.no_update),
        {} if show_office else hide,
        district_opts if show_district else dash.no_update,
        dist_val if show_district else (None if reset else dash.no_update), {} if show_district else hide
    )


@app.callback([Output('analysis-status', 'children'), Output('performance-summary', 'children'),
               Output('main-visualizations', 'children'), Output('detailed-breakdown', 'children')],
              [Input('analysis-mode-selector', 'value'), Input('month-dropdown', 'value'),
               Input('single-entity-dropdown', 'value'), Input('entity1-dropdown', 'value'),
               Input('entity2-dropdown', 'value'),
               Input('drill-down-levels', 'value'), Input('drilldown-service-dropdown', 'value'),
               Input('drilldown-office-dropdown', 'value'), Input('drilldown-district-dropdown', 'value')],
              [State('primary-level', 'value'), State('fy-store', 'data')]   # ← fy as State
              )
def update_dashboard(mode, months, single_entity, entity1, entity2, drill_levels, service_filter, office_filter,
                     district_filter, primary_level, fy):   # ← add fy
    df_mt = FY_DATA[fy]['df_mt']   # ← add this as first line
    if not months: return dbc.Alert("Please select a time period.", color="warning"), None, None, None
    col_map = {'district': 'District_name', 'service': 'Service_name', 'office': 'Office_name'}
    primary_col = col_map[primary_level]

    def filter_df(df_target):
        if 'service' in (drill_levels or []) and service_filter and not str(service_filter).startswith(
            'ALL_'): df_target = df_target[df_target['Service_name'] == service_filter]
        if 'office' in (drill_levels or []) and office_filter and not str(office_filter).startswith('ALL_'): df_target = \
        df_target[df_target['Office_name'] == office_filter]
        if 'district' in (drill_levels or []) and district_filter and not str(district_filter).startswith(
            'ALL_'): df_target = df_target[df_target['District_name'] == district_filter]
        return df_target

    if mode == 'single':
        if not single_entity: return dbc.Alert("Please select an entity to analyze.", color="info"), None, None, None
        df1 = filter_df(df_mt[(df_mt['Month_Year'].isin(months)) & (df_mt[primary_col] == single_entity)])
        if df1.empty: return dbc.Alert(f"Displaying analysis for: {single_entity}", color="success"), dbc.Alert(
            "No data available for this selection.", color="warning"), None, None
        return dbc.Alert(f"Displaying analysis for: {single_entity}", color="success"), generate_summary_cards(df1,
                                                                                                               None,
                                                                                                               single_entity,
                                                                                                               None,
                                                                                                               mode), generate_main_visuals(
            df1, None, single_entity, None, months, mode, primary_level), html.P(
            "Drill-down selections act as filters only.", className="text-muted")
    else:
        if not entity1 or not entity2: return dbc.Alert("Please select two entities to compare.",
                                                        color="info"), None, None, None
        if entity1 == entity2: return dbc.Alert("Please select different entities for comparison.",
                                                color="warning"), None, None, None
        df1, df2 = filter_df(df_mt[(df_mt['Month_Year'].isin(months)) & (df_mt[primary_col] == entity1)]), filter_df(
            df_mt[(df_mt['Month_Year'].isin(months)) & (df_mt[primary_col] == entity2)])
        if df1.empty and df2.empty: return dbc.Alert(f"Comparing {entity1} vs {entity2}", color="success"), dbc.Alert(
            "No data for this selection.", color="warning"), None, None
        return dbc.Alert(f"Comparing {entity1} vs {entity2}", color="success"), generate_summary_cards(df1, df2,
                                                                                                       entity1, entity2,
                                                                                                       mode), generate_main_visuals(
            df1, df2, entity1, entity2, months, mode, primary_level), html.P(
            "Drill-down selections act as filters only.", className="text-muted")


def generate_summary_cards(df1, df2, entity1, entity2, mode):
    data_map = [(entity1, df1)]
    if mode == 'comparison': data_map.append((entity2, df2))
    cards = []
    for i, (name, data) in enumerate(data_map):
        card = dbc.Col(dbc.Card([
            dbc.CardHeader(name, className="text-white", style={
                'backgroundColor': COLOR_PALETTE['primary'] if i == 0 else COLOR_PALETTE['secondary']}),
            dbc.CardBody([
                html.P(f"Received: {data['application_Received'].sum():,}"),
                html.P(f"Disposed: {data['application_Disposed'].sum():,}"),
                html.P(f"Out-of-Time: {data['application_Disposed_Out_of_time'].sum():,}"),
                html.P(f"Pending: {data['Pending_Applications'].sum():,}"),
                html.H5(f"Efficiency: {data['Efficiency_Percentage'].mean() if not data.empty else 0:.2f}%")
            ])
        ]), md=6 if mode == 'comparison' else 12)
        cards.append(card)
    return dbc.Row(cards, justify="center")


def generate_main_visuals(df1, df2, entity1, entity2, months, mode, primary_level):
    def build_trend_figs(dfX, entity_label):
        if dfX.empty: return html.P(f"No data available for {entity_label}.", className="text-muted")
        metrics = [('application_Received', 'Applications Received', '#1f77b4'),
                   ('application_Disposed', 'Applications Disposed', '#ff7f0e'),
                   ('application_Disposed_with_in_time', 'On-Time Disposal', '#2ca02c'),
                   ('application_Disposed_Out_of_time', 'Out-of-Time Disposal', '#8c564b'),
                   ('Pending_Applications', 'Pending Applications', '#d62728')]
        trend_data = dfX.groupby('Month_Year')[
            ['application_Received', 'application_Disposed', 'application_Disposed_with_in_time',
             'application_Disposed_Out_of_time', 'Pending_Applications', 'Efficiency_Percentage']].agg(
            {'application_Received': 'sum', 'application_Disposed': 'sum', 'application_Disposed_with_in_time': 'sum',
             'application_Disposed_Out_of_time': 'sum', 'Pending_Applications': 'sum',
             'Efficiency_Percentage': 'mean'}).reset_index()
        trend_data['Date'] = pd.to_datetime(trend_data['Month_Year'], format='%b-%Y', errors='coerce')
        trend_data = trend_data.sort_values('Date')

        fig_scaler = go.Figure()
        for col, label, color in metrics: fig_scaler.add_trace(
            go.Scatter(x=trend_data['Month_Year'], y=trend_data[col], mode='lines+markers', name=label,
                       line=dict(color=color)))
        fig_scaler.update_layout(title=f"{entity_label} - Scaler Values", template="plotly_white", xaxis_title="Month",
                                 yaxis_title="Value")

        fig_percent = go.Figure()
        fig_percent.add_trace(
            go.Scatter(x=trend_data['Month_Year'], y=trend_data['Efficiency_Percentage'], mode='lines+markers',
                       name="Efficiency %", line=dict(color='#9467bd')))
        if primary_level == 'district' and set(months or []) == set(df_mt['Month_Year'].unique()):
            trend_data['Out_of_Time_%'] = trend_data.apply(
                lambda row: (row['application_Disposed_Out_of_time'] / row['application_Disposed'] * 100) if row[
                                                                                                                 'application_Disposed'] > 0 else 0,
                axis=1)
            fig_percent.add_trace(
                go.Scatter(x=trend_data['Month_Year'], y=trend_data['Out_of_Time_%'], mode='lines+markers',
                           name='Out-of-Time %', line=dict(color='#8c564b', dash='dash')))
        fig_percent.update_layout(title=f"{entity_label} - Efficiency %", template="plotly_white", xaxis_title="Month",
                                  yaxis_title="Percentage", yaxis=dict(range=[0, 100]))

        return html.Div([dcc.Graph(figure=fig_scaler, style={'marginBottom': '24px'}), dcc.Graph(figure=fig_percent),
                         html.Div(f"Out-of-Time (Total): {int(dfX['application_Disposed_Out_of_time'].sum()):,}",
                                  className="text-muted mt-2", style={'fontWeight': '600'})])

    if mode == 'single': return dbc.Row([dbc.Col(build_trend_figs(df1, entity1), md=12)])
    return dbc.Row([dbc.Col(build_trend_figs(df1, entity1), md=6), dbc.Col(build_trend_figs(df2, entity2), md=6)])
