import streamlit as st
import pandas as pd
import plotly.express as px
from dash import html, dcc, Input, Output, State
import dash_bootstrap_components as dbc
from app import app
from data import df_adv

# ─────────────────────────────────────────────────────────────────────────────
# Helper: compute OOT rate safely
# ─────────────────────────────────────────────────────────────────────────────
def _oot_rate(oot, total):
    return round((oot / total * 100), 2) if total > 0 else 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Prepare df_adv — rename raw CSV columns to internal names
# Raw columns: Year, Month, Service_name, District_name, Office_name,
#              application_Received, application_Disposed,
#              application_Disposed_Out_of_time, application_Disposed_with_in_time,
#              Pending, Total   ← NOTE: 'Total' already exists in CSV
# ─────────────────────────────────────────────────────────────────────────────
def _prepare_df(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=['District', 'Office', 'Service', 'OOT', 'Total', 'month_dt'])

    df = df.copy()
    df.columns = df.columns.str.strip()

    # Drop the raw 'Total' column if present — we will use application_Disposed as 'Total'
    if 'Total' in df.columns and 'application_Disposed' in df.columns:
        df = df.drop(columns=['Total'])

    rename_map = {
        'District_name':                    'District',
        'Office_name':                      'Office',
        'Service_name':                     'Service',
        'application_Disposed_Out_of_time': 'OOT',
        'application_Disposed':             'Total',
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # Build month_dt
    if 'month_dt' not in df.columns:
        if 'Year' in df.columns and 'Month' in df.columns:
            df['month_dt'] = pd.to_datetime(
                df['Year'].astype(str) + '-' + df['Month'].astype(str).str.zfill(2) + '-01',
                format='%Y-%m-%d'
            )
        else:
            df['month_dt'] = pd.NaT

    for col in ['District', 'Office', 'Service']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    for col in ['OOT', 'Total']:
        if col in df.columns:
            # Guard: ensure it is a Series not a DataFrame (duplicate columns cause this)
            if isinstance(df[col], pd.DataFrame):
                print(f"[advanced_analytics] WARNING: '{col}' is duplicated — keeping first occurrence.")
                df = df.loc[:, ~df.columns.duplicated()]
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        else:
            print(f"[advanced_analytics] WARNING: '{col}' not found. Available: {df.columns.tolist()}")
            df[col] = 0

    return df


_df = _prepare_df(df_adv)

MONTH_NAMES = {
    1: "January", 2: "February", 3: "March",     4: "April",
    5: "May",     6: "June",     7: "July",       8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}
MONTH_NAME_TO_NUM = {v: k for k, v in MONTH_NAMES.items()}


# ─────────────────────────────────────────────────────────────────────────────
# Scoring engine  (uses internal column names: OOT, Total)
# ─────────────────────────────────────────────────────────────────────────────
def _score_offices(df, selected_district, selected_year, selected_month, min_count):

    # Step 1: Snapshot for selected district / year / month
    snap = df[
        (df['District']           == selected_district) &
        (df['month_dt'].dt.year   == selected_year)     &
        (df['month_dt'].dt.month  == selected_month)
    ]
    if snap.empty:
        return pd.DataFrame(), 0.0

    # Step 2: Aggregate by office for the snapshot
    office_snap = (
        snap.groupby('Office')
        .agg(Total=('Total', 'sum'), OOT=('OOT', 'sum'))
        .reset_index()
    )
    office_snap['OOT_Rate'] = office_snap.apply(
        lambda r: _oot_rate(r['OOT'], r['Total']), axis=1
    )

    # Step 3: Compute district average BEFORE filtering (all offices included)
    district_avg_total = round(office_snap['Total'].mean(), 2)

    # Step 4: Filter out offices below min_count threshold
    if min_count > 0:
        office_snap = office_snap[office_snap['Total'] >= min_count].reset_index(drop=True)

    if office_snap.empty:
        return pd.DataFrame(), district_avg_total

    # Step 5: District & state averages (OOT rate) from filtered snapshot
    district_avg_oot = _oot_rate(office_snap['OOT'].sum(), office_snap['Total'].sum())

    state_snap = df[
        (df['month_dt'].dt.year  == selected_year) &
        (df['month_dt'].dt.month == selected_month)
    ]
    state_avg_oot = _oot_rate(state_snap['OOT'].sum(), state_snap['Total'].sum())

    # Step 6: Historical consistency streak (only for offices that passed threshold)
    valid_offices = office_snap['Office'].tolist()
    dist_history = (
        df[
            (df['District'] == selected_district) &
            (df['Office'].isin(valid_offices))
        ]
        .groupby(['Office', 'month_dt'])
        .agg(Total=('Total', 'sum'), OOT=('OOT', 'sum'))
        .reset_index()
    )
    dist_history['OOT_Rate'] = dist_history.apply(
        lambda r: _oot_rate(r['OOT'], r['Total']), axis=1
    )
    dist_monthly_avg = (
        dist_history.groupby('month_dt')
        .apply(lambda g: _oot_rate(g['OOT'].sum(), g['Total'].sum()))
        .reset_index(name='Dist_Avg')
    )
    dist_history = dist_history.merge(dist_monthly_avg, on='month_dt')
    dist_history['is_bad'] = dist_history['OOT_Rate'] > dist_history['Dist_Avg']

    cutoff = pd.Timestamp(year=selected_year, month=selected_month, day=1)
    consistency_scores = {}
    for office, grp in dist_history[dist_history['month_dt'] <= cutoff].groupby('Office'):
        bad_flags = grp.sort_values('month_dt')['is_bad'].tolist()
        streak = 0
        for flag in reversed(bad_flags):
            if flag:
                streak += 1
            else:
                break
        consistency_scores[office] = streak

    office_snap['Streak'] = office_snap['Office'].map(consistency_scores).fillna(0).astype(int)

    # Step 7: Composite score
    max_oot = office_snap['OOT_Rate'].max() or 1
    office_snap['F1_District']    = ((office_snap['OOT_Rate'] - district_avg_oot) / max_oot * 100).clip(0)
    office_snap['F2_State']       = ((office_snap['OOT_Rate'] - state_avg_oot)    / max_oot * 100).clip(0)
    office_snap['F3_Consistency'] = office_snap['Streak'].apply(
        lambda s: 100 if s >= 9 else 66 if s >= 6 else 33 if s >= 3 else 0
    )
    office_snap['Composite_Score']  = (
        0.35 * office_snap['F1_District'] +
        0.35 * office_snap['F2_State']    +
        0.30 * office_snap['F3_Consistency']
    ).round(2)
    office_snap['District_Avg_OOT'] = district_avg_oot
    office_snap['State_Avg_OOT']    = state_avg_oot

    return office_snap.sort_values('Composite_Score', ascending=False).reset_index(drop=True), district_avg_total


# ─────────────────────────────────────────────────────────────────────────────
# Reason builder
# ─────────────────────────────────────────────────────────────────────────────
def _build_reason_items(row):
    items  = []
    delta_dist  = row['OOT_Rate'] - row['District_Avg_OOT']
    delta_state = row['OOT_Rate'] - row['State_Avg_OOT']
    streak      = int(row['Streak'])

    if delta_dist > 0:
        items.append(f"📍 District: OOT rate is {delta_dist:.1f}% higher than district average ({row['District_Avg_OOT']:.1f}%).")
    else:
        items.append(f"📍 District: OOT rate is {abs(delta_dist):.1f}% below district average ({row['District_Avg_OOT']:.1f}%).")

    if delta_state > 0:
        items.append(f"🌐 State: OOT rate exceeds state average ({row['State_Avg_OOT']:.1f}%) by {delta_state:.1f}%.")
    else:
        items.append(f"🌐 State: OOT rate is {abs(delta_state):.1f}% below state average ({row['State_Avg_OOT']:.1f}%).")

    if streak >= 9:
        items.append(f"🚨 Consistency: Above district avg OOT for {streak} consecutive months — serious persistent concern.")
    elif streak >= 6:
        items.append(f"⚠️ Consistency: Flagged for {streak} consecutive months — sustained performance issue.")
    elif streak >= 3:
        items.append(f"🔔 Consistency: Above-average OOT for {streak} consecutive months — requires monitoring.")
    else:
        items.append(f"✅ Consistency: No prolonged OOT streak detected ({streak} month(s)).")

    return items


# ─────────────────────────────────────────────────────────────────────────────
# KPI card helper
# ─────────────────────────────────────────────────────────────────────────────
def _kpi_card(label, value, color="#1a3c5e", bg="#f0f6ff", border="#2d6a9f"):
    return html.Div([
        html.Div(value, style={'fontSize': '1.5rem', 'fontWeight': '700', 'color': color}),
        html.Div(label, style={'fontSize': '0.78rem', 'color': '#555', 'fontWeight': '600',
                               'textTransform': 'uppercase', 'letterSpacing': '0.5px'}),
    ], style={
        'background': bg, 'borderLeft': f'5px solid {border}',
        'padding': '14px 12px', 'borderRadius': '8px', 'textAlign': 'center'
    })


# ─────────────────────────────────────────────────────────────────────────────
# Layout
# ─────────────────────────────────────────────────────────────────────────────
districts = sorted(_df['District'].dropna().unique()) if not _df.empty else []
years     = sorted(_df['month_dt'].dt.year.unique())  if not _df.empty else []

layout = html.Div([
    html.Div([
        html.H2("🔍 Advanced Analytics & Performance Report",
                style={'color': 'white', 'margin': '0', 'fontSize': '1.6rem', 'letterSpacing': '1px'}),
        html.P("Identify worst-performing offices using district benchmarking, state averages, and consistency analysis.",
               style={'color': '#c8dff0', 'margin': '4px 0 0 0', 'fontSize': '0.95rem'}),
    ], style={
        'background': 'linear-gradient(90deg, #1a3c5e 0%, #2d6a9f 100%)',
        'padding': '18px 28px', 'borderRadius': '10px', 'marginBottom': '20px'
    }),

    dbc.Row([
        dbc.Col([
            html.Label("🏛️ Select District"),
            dcc.Dropdown(
                id='aa-district',
                options=[{'label': d, 'value': d} for d in districts],
                value=districts[0] if districts else None,
                clearable=False
            )
        ], md=3),
        dbc.Col([
            html.Label("📅 Select Year"),
            dcc.Dropdown(
                id='aa-year',
                options=[{'label': y, 'value': y} for y in years],
                value=years[-1] if years else None,
                clearable=False
            )
        ], md=2),
        dbc.Col([
            html.Label("🗓️ Select Month"),
            dcc.Dropdown(id='aa-month', clearable=False)
        ], md=3),
        dbc.Col([
            html.Label("⚙️ Min. Service Count Threshold"),
            dcc.Input(id='aa-min-count', type='number', value=100, min=0,
                      style={'width': '100%', 'padding': '6px', 'borderRadius': '4px',
                             'border': '1px solid #ccc'})
        ], md=2),
        dbc.Col([
            html.Br(),
            dbc.Button("🔍 Analyse", id='aa-run-btn', color='primary',
                       className='mt-1', style={'width': '100%'})
        ], md=2),
    ], className='mb-4'),

    html.Div(id='aa-output')

], style={'padding': '20px'})


# ─────────────────────────────────────────────────────────────────────────────
# Callbacks
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output('aa-month', 'options'),
    Output('aa-month', 'value'),
    Input('aa-district', 'value'),
    Input('aa-year',     'value'),
)
def update_months(district, year):
    if not district or not year or _df.empty:
        return [], None
    filtered = _df[(_df['District'] == district) & (_df['month_dt'].dt.year == year)]
    months   = sorted(filtered['month_dt'].dt.month.unique())
    opts     = [{'label': MONTH_NAMES[m], 'value': m} for m in months]
    return opts, (months[-1] if months else None)


@app.callback(
    Output('aa-output', 'children'),
    Input('aa-run-btn', 'n_clicks'),
    State('aa-district',  'value'),
    State('aa-year',      'value'),
    State('aa-month',     'value'),
    State('aa-min-count', 'value'),
    prevent_initial_call=True,
)
def run_analysis(n_clicks, district, year, month, min_count):
    if not all([district, year, month]):
        return dbc.Alert("Please select all filters.", color="warning")

    min_count  = int(min_count or 0)
    scored, district_avg_total = _score_offices(_df, district, int(year), int(month), min_count)

    if scored.empty:
        return dbc.Alert(
            f"No offices with Total ≥ {min_count} found for the selected filters.",
            color="warning"
        )

    district_avg_oot = scored['District_Avg_OOT'].iloc[0]
    state_avg_oot    = scored['State_Avg_OOT'].iloc[0]
    month_name       = MONTH_NAMES.get(int(month), str(month))

    # ── KPI bar ──────────────────────────────────────────────────────────────
    kpi_row = dbc.Row([
        dbc.Col(_kpi_card("Offices Analysed",
                          len(scored)), md=2),
        dbc.Col(_kpi_card("Avg Total Applications",
                          f"{district_avg_total:,.1f}",
                          color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=2),
        dbc.Col(_kpi_card("Min Count Threshold",
                          f"≥ {min_count}",
                          color="#555", bg="#f8f9fa", border="#aaa"), md=2),
        dbc.Col(_kpi_card("District OOT Avg",
                          f"{district_avg_oot:.1f}%",
                          color="#c0392b", bg="#fff5f5", border="#e74c3c"), md=2),
        dbc.Col(_kpi_card("State OOT Avg",
                          f"{state_avg_oot:.1f}%",
                          color="#7d3c98", bg="#fdf5ff", border="#9b59b6"), md=2),
        dbc.Col(_kpi_card("Flagged (≥3 months)",
                          int((scored['Streak'] >= 3).sum()),
                          color="#e67e22", bg="#fff8f0", border="#e67e22"), md=2),
    ], className='mb-4')

    # ── Threshold info banner ─────────────────────────────────────────────────
    threshold_banner = dbc.Alert(
        [
            html.Strong(f"📊 District: {district}  |  {month_name} {year}"),
            html.Br(),
            f"Average total applications across all offices (before filtering): ",
            html.Strong(f"{district_avg_total:,.1f}"),
            f"  —  Offices with Total < {min_count} have been excluded from analysis."
            if min_count > 0 else
            "  —  No minimum threshold applied (all offices included)."
        ],
        color="info", className="mb-4"
    )

    # ── Table ────────────────────────────────────────────────────────────────
    display_cols = ['Office', 'Total', 'OOT', 'OOT_Rate',
                    'District_Avg_OOT', 'State_Avg_OOT', 'Streak', 'Composite_Score']
    table_df = scored[display_cols].rename(columns={
        'OOT':              'Out of Time',
        'OOT_Rate':         'OOT Rate (%)',
        'District_Avg_OOT': 'District Avg OOT (%)',
        'State_Avg_OOT':    'State Avg OOT (%)',
        'Streak':           'Bad-Month Streak',
        'Composite_Score':  'Composite Score',
    })
    table_section = html.Div([
        html.H4(f"📊 Office-wise Performance — {month_name} {year}  |  {district}"),
        dbc.Table.from_dataframe(table_df.round(2), striped=True,
                                 bordered=True, hover=True, responsive=True, size='sm')
    ], className='mb-4')

    # ── Bar chart ────────────────────────────────────────────────────────────
    fig = px.bar(
        scored.sort_values('Composite_Score', ascending=False),
        x='Office', y='OOT_Rate', color='Composite_Score',
        color_continuous_scale='Reds',
        title=f"Out-of-Time Rate by Office — {month_name} {year}  |  {district}",
        labels={'OOT_Rate': '% Out of Time', 'Composite_Score': 'Composite Score'},
        text='OOT_Rate',
    )
    fig.add_hline(y=district_avg_oot, line_dash='dash', line_color='#2980b9',
                  annotation_text=f"District Avg OOT: {district_avg_oot:.1f}%",
                  annotation_position="top left")
    fig.add_hline(y=state_avg_oot, line_dash='dot', line_color='#8e44ad',
                  annotation_text=f"State Avg OOT: {state_avg_oot:.1f}%",
                  annotation_position="bottom right")
    fig.update_layout(xaxis_tickangle=-45, height=480)
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    chart_section = dcc.Graph(figure=fig, className='mb-4')

    # ── Worst 3 ──────────────────────────────────────────────────────────────
    worst_labels = ["🥇 Rank 1 — Worst", "🥈 Rank 2", "🥉 Rank 3"]
    worst_items  = []
    for rank, (_, row) in enumerate(scored.head(min(3, len(scored))).iterrows(), 1):
        reasons = _build_reason_items(row)
        worst_items.append(dbc.Card([
            dbc.CardHeader(html.Strong(
                f"{worst_labels[rank-1]}  |  {row['Office']}  |  "
                f"OOT: {row['OOT_Rate']:.1f}%  |  Score: {row['Composite_Score']:.1f}"
            ), style={'background': '#fff0f0', 'color': '#c0392b'}),
            dbc.CardBody([
                dbc.Row([
                    dbc.Col(_kpi_card("Total Applications", f"{int(row['Total']):,}",
                                      color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=3),
                    dbc.Col(_kpi_card("OOT Rate",     f"{row['OOT_Rate']:.1f}%",
                                      color="#c0392b", bg="#fff5f5", border="#e74c3c"), md=3),
                    dbc.Col(_kpi_card("District Avg", f"{row['District_Avg_OOT']:.1f}%",
                                      color="#2980b9", bg="#eaf4ff", border="#3498db"), md=3),
                    dbc.Col(_kpi_card("Bad Months",   f"{int(row['Streak'])}",
                                      color="#e67e22", bg="#fff8f0", border="#e67e22"), md=3),
                ], className='mb-3'),
                html.Strong("📝 Reason for Flagging:"),
                html.Ul([html.Li(r) for r in reasons])
            ])
        ], className='mb-3'))

    worst_section = html.Div([
        html.Div([
            html.H3("⚠️ Worst Performing Offices", style={'margin': '0', 'color': '#c0392b'}),
            html.P("Ranked by composite score: district benchmark (35%) + state average (35%) + consistency (30%)",
                   style={'margin': '4px 0 0 0', 'color': '#7f8c8d', 'fontSize': '0.9rem'}),
        ], style={'background': '#fff0f0', 'border': '1.5px solid #e74c3c',
                  'borderRadius': '10px', 'padding': '14px 22px', 'marginBottom': '12px'}),
        *worst_items
    ], className='mb-4')

    # ── Best 3 ───────────────────────────────────────────────────────────────
    best_labels = ["🥇 Best Office", "🥈 2nd Best", "🥉 3rd Best"]
    best_items  = []
    best_rows   = scored.tail(min(3, len(scored))).iloc[::-1].reset_index(drop=True)
    for rank, (_, row) in enumerate(best_rows.iterrows(), 1):
        reasons = _build_reason_items(row)
        best_items.append(dbc.Card([
            dbc.CardHeader(html.Strong(
                f"{best_labels[rank-1]}  |  {row['Office']}  |  "
                f"OOT: {row['OOT_Rate']:.1f}%  |  Score: {row['Composite_Score']:.1f}"
            ), style={'background': '#f0fff4', 'color': '#1e8449'}),
            dbc.CardBody([
                dbc.Row([
                    dbc.Col(_kpi_card("Total Applications", f"{int(row['Total']):,}",
                                      color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=3),
                    dbc.Col(_kpi_card("OOT Rate",     f"{row['OOT_Rate']:.1f}%",
                                      color="#1e8449", bg="#f0fff4", border="#27ae60"), md=3),
                    dbc.Col(_kpi_card("District Avg", f"{row['District_Avg_OOT']:.1f}%",
                                      color="#2980b9", bg="#eaf4ff", border="#3498db"), md=3),
                    dbc.Col(_kpi_card("Bad Months",   f"{int(row['Streak'])}",
                                      color="#27ae60", bg="#f0fff4", border="#27ae60"), md=3),
                ], className='mb-3'),
                html.Strong("📝 Performance Summary:"),
                html.Ul([html.Li(r) for r in reasons])
            ])
        ], className='mb-3'))

    best_section = html.Div([
        html.Div([
            html.H3("🏆 Best Performing Offices", style={'margin': '0', 'color': '#1e8449'}),
            html.P("Offices with lowest composite score (lowest OOT burden relative to benchmarks)",
                   style={'margin': '4px 0 0 0', 'color': '#7f8c8d', 'fontSize': '0.9rem'}),
        ], style={'background': '#f0fff4', 'border': '1.5px solid #27ae60',
                  'borderRadius': '10px', 'padding': '14px 22px', 'marginBottom': '12px'}),
        *best_items
    ], className='mb-4')

    # ── Disclaimer ────────────────────────────────────────────────────────────
    disclaimer = html.Div(
        f"Methodology: Offices with Total < {min_count} are excluded. "
        f"Average total applications (all offices before filter): {district_avg_total:,.1f}. "
        f"Composite score = District Benchmark 35% · State Average 35% · Consistency 30%. "
        f"A 'bad month' = office OOT rate exceeds district average that month. "
        f"Streaks ≥3, ≥6, ≥9 months flagged with 🔔, ⚠️, 🚨.",
        style={'background': '#f8f9fa', 'border': '1px solid #d0d7de', 'borderRadius': '8px',
               'padding': '12px 18px', 'color': '#555', 'fontSize': '0.82rem'}
    )

    return html.Div([kpi_row, threshold_banner, table_section, chart_section,
                     worst_section, best_section, disclaimer])