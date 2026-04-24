from dash import html, dcc, Input, Output, State, ALL, callback_context
import dash_bootstrap_components as dbc
import pandas as pd
import numpy as np
import plotly.express as px  # 🆕 Added for generating the popup graph
from app import app
from data import df_adv

# ═════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ═════════════════════════════════════════════════════════════════════════════
MONTH_NAMES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}


# ═════════════════════════════════════════════════════════════════════════════
# DATA PREPARATION
# ═════════════════════════════════════════════════════════════════════════════
def _prepare_df(raw):
    if raw is None or raw.empty:
        return pd.DataFrame(columns=['District', 'Office', 'Service', 'OOT', 'Total', 'month_dt'])

    df = raw.copy()
    df.columns = df.columns.str.strip()

    if 'Total' in df.columns and 'application_Disposed' in df.columns:
        df.drop(columns=['Total'], inplace=True)

    # Handle both _name (from _load_mt) and _Eng (raw CSV) column naming conventions
    _map = {
        'District_name': 'District', 'District_Eng': 'District',
        'Office_name': 'Office',     'Office_Eng': 'Office',
        'Service_name': 'Service',   'Service_Eng': 'Service',
        'application_Disposed_Out_of_time': 'OOT',
        'application_Disposed': 'Total',
    }
    df.rename(columns={k: v for k, v in _map.items() if k in df.columns}, inplace=True)

    if 'month_dt' not in df.columns:
        # Normalise year/month column names (raw CSV uses Yr/Mn, _load_mt renames to Year/Month)
        if 'Yr' in df.columns and 'Year' not in df.columns:
            df.rename(columns={'Yr': 'Year'}, inplace=True)
        if 'Mn' in df.columns and 'Month' not in df.columns:
            df.rename(columns={'Mn': 'Month'}, inplace=True)

        if 'Year' in df.columns and 'Month' in df.columns:
            df['Year'] = pd.to_numeric(df['Year'].astype(str).str.replace('\ufeff', '', regex=False).str.strip(), errors='coerce')
            df['Month'] = pd.to_numeric(df['Month'], errors='coerce')
            df.dropna(subset=['Year', 'Month'], inplace=True)
            df['month_dt'] = pd.to_datetime(
                df['Year'].astype(int).astype(str) + '-' + df['Month'].astype(int).astype(str).str.zfill(2) + '-01')
        else:
            df['month_dt'] = pd.NaT

    for c in ('District', 'Office', 'Service'):
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    for c in ('OOT', 'Total'):
        df = df.loc[:, ~df.columns.duplicated()]
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
        else:
            df[c] = 0
    return df


_df = _prepare_df(df_adv)


# ═════════════════════════════════════════════════════════════════════════════
# PRE-COMPUTED CACHES
# ═════════════════════════════════════════════════════════════════════════════
def _oot_pct(s):
    return np.where(s['Total'] > 0, (s['OOT'] / s['Total'] * 100).round(2), 0.0)


def _build_caches(df):
    if df.empty:
        return {}, {}, {}, {}, {}

    om = df.groupby(['District', 'Office', 'month_dt'], as_index=False).agg(
        Total=('Total', 'sum'), OOT=('OOT', 'sum'))
    om['OOT_Rate'] = _oot_pct(om)

    dm = df.groupby(['District', 'month_dt'], as_index=False).agg(
        Total=('Total', 'sum'), OOT=('OOT', 'sum'))
    dm['OOT_Rate'] = _oot_pct(dm)

    sm = df.groupby('month_dt', as_index=False).agg(
        Total=('Total', 'sum'), OOT=('OOT', 'sum'))
    sm['OOT_Rate'] = _oot_pct(sm)

    ssm = df.groupby(['Service', 'month_dt'], as_index=False).agg(
        Total=('Total', 'sum'), OOT=('OOT', 'sum'))
    ssm['OOT_Rate'] = _oot_pct(ssm)

    osm = df.groupby(['Office', 'Service', 'month_dt'], as_index=False).agg(
        Total=('Total', 'sum'), OOT=('OOT', 'sum'))
    osm['OOT_Rate'] = _oot_pct(osm)

    for frame in (om, dm, sm, ssm, osm):
        frame['_y'] = frame['month_dt'].dt.year
        frame['_m'] = frame['month_dt'].dt.month

    return om, dm, sm, ssm, osm


_OM, _DM, _SM, _SSM, _OSM = _build_caches(_df)


# ═════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════════
def _dist_oot(district, y, m):
    r = _DM[(_DM['District'] == district) & (_DM['_y'] == y) & (_DM['_m'] == m)]
    return float(r['OOT_Rate'].iloc[0]) if len(r) else 0.0


def _state_oot(y, m):
    r = _SM[(_SM['_y'] == y) & (_SM['_m'] == m)]
    return float(r['OOT_Rate'].iloc[0]) if len(r) else 0.0


def _compute_streaks(district, valid_offices, cutoff_ts):
    """Return {office: streak_int} for valid_offices up to cutoff_ts."""
    hist = _OM[
        (_OM['District'] == district) &
        (_OM['Office'].isin(valid_offices)) &
        (_OM['month_dt'] <= cutoff_ts)
        ].copy()
    if hist.empty:
        return {o: 0 for o in valid_offices}

    davg = _DM[(_DM['District'] == district) & (_DM['month_dt'] <= cutoff_ts)]
    davg_map = dict(zip(davg['month_dt'], davg['OOT_Rate']))
    hist['d_avg'] = hist['month_dt'].map(davg_map).fillna(0)
    hist['bad'] = hist['OOT_Rate'] > hist['d_avg']
    hist.sort_values(['Office', 'month_dt'], inplace=True)

    streaks = {}
    for off, grp in hist.groupby('Office'):
        arr = grp['bad'].values
        cnt = 0
        for val in reversed(arr):
            if val:
                cnt += 1
            else:
                break
        streaks[off] = cnt
    for o in valid_offices:
        streaks.setdefault(o, 0)
    return streaks


# ═════════════════════════════════════════════════════════════════════════════
# SCORING ENGINE (from advanced_analytics.py)
# ═════════════════════════════════════════════════════════════════════════════
def _score_offices(district, y, m, min_c=0, max_c=None):
    snap = _OM[(_OM['District'] == district) & (_OM['_y'] == y) & (_OM['_m'] == m)].copy()
    if snap.empty:
        return pd.DataFrame()

    # NEW: Calculate district average for applications
    district_avg_total = snap['Total'].mean()

    # Filter for above-average application volume
    snap = snap[snap['Total'] >= district_avg_total]

    if min_c > 0:
        snap = snap[snap['Total'] >= min_c]
    if max_c and max_c > 0:
        snap = snap[snap['Total'] <= max_c]
    snap = snap.reset_index(drop=True)
    if snap.empty:
        return pd.DataFrame()

    offices = snap['Office'].tolist()
    d_oot = _dist_oot(district, y, m)
    s_oot = _state_oot(y, m)
    cutoff_ts = pd.Timestamp(year=y, month=m, day=1)

    # F1: District Deviation
    d_min = snap['OOT_Rate'].min()
    d_max = snap['OOT_Rate'].max()

    if d_max > d_min:
        snap['F1'] = ((d_max - snap['OOT_Rate']) / (d_max - d_min) * 25).round(2)
    else:
        snap['F1'] = 25.0

    # F2: State Deviation
    if s_oot > 0:
        snap['F2'] = 12.5 + ((s_oot - snap['OOT_Rate']) / s_oot * 12.5)
        snap['F2'] = snap['F2'].clip(0, 25).round(2)
    else:
        snap['F2'] = 25.0

    # F3: Streak Score
    streaks = _compute_streaks(district, offices, cutoff_ts)
    snap['Streak'] = snap['Office'].map(streaks).fillna(0).astype(int)
    snap['F3'] = snap['Streak'].apply(
        lambda s: 0.0 if s >= 9 else 8.33 if s >= 6 else 16.67 if s >= 3 else 25.0)

    # F4: Service Concentration Score
    svc_agg = _df[
        (_df['District'] == district) & (_df['Office'].isin(offices)) &
        (_df['month_dt'].dt.year == y) & (_df['month_dt'].dt.month == m)
        ].groupby(['Office', 'Service'], as_index=False).agg(OOT=('OOT', 'sum'))

    top_svc_oot = svc_agg.groupby('Office')['OOT'].max().to_dict()
    snap['Top_Svc_OOT'] = snap['Office'].map(top_svc_oot).fillna(0)
    snap['Top_Svc_Share'] = np.where(snap['OOT'] > 0, snap['Top_Svc_OOT'] / snap['OOT'], 0)
    snap['F4'] = np.where(snap['Top_Svc_Share'] >= 0.80, 25.0, 0.0)

    # Final Composite Score
    snap['Composite_Score'] = (snap['F1'] + snap['F2'] + snap['F3'] + snap['F4']).round(2)
    snap['District_Avg_OOT'] = d_oot
    snap['State_Avg_OOT'] = s_oot

    return snap


def _calculate_district_scores(y, m, min_c=0, max_c=None):
    """Calculate composite scores for all districts."""
    districts = sorted(_df['District'].dropna().unique())
    results = []

    for district in districts:
        office_scores = _score_offices(district, y, m, min_c, max_c)
        if office_scores.empty:
            continue

        # Calculate district average composite score
        avg_composite = office_scores['Composite_Score'].mean()

        # Get offices above district average
        above_avg = office_scores[office_scores['Composite_Score'] > avg_composite]

        results.append({
            'District': district,
            'Composite_Score': round(avg_composite, 2),
            'Total_Offices': len(office_scores),
            'Above_Avg_Offices': len(above_avg),
            'office_data': office_scores.sort_values('Composite_Score').to_dict('records')
        })

    return pd.DataFrame(results)


def _calculate_service_scores(y, m, min_c=0):
    """
    Calculate scores for services that are:
    1. Above average in total applications
    2. High OOT rate
    3. Distributed across multiple offices/districts (not concentrated)
    """
    # Get state-level service data for the month
    service_data = _df[
        (_df['month_dt'].dt.year == y) &
        (_df['month_dt'].dt.month == m)
        ].groupby('Service', as_index=False).agg(
        Total=('Total', 'sum'),
        OOT=('OOT', 'sum'),
        Num_Districts=('District', 'nunique'),
        Num_Offices=('Office', 'nunique')
    )

    if service_data.empty:
        return pd.DataFrame()

    # Calculate OOT rate
    service_data['OOT_Rate'] = np.where(
        service_data['Total'] > 0,
        (service_data['OOT'] / service_data['Total'] * 100).round(2),
        0.0
    )

    # Filter 1: Above average applications
    avg_total = service_data['Total'].mean()
    service_data = service_data[service_data['Total'] >= avg_total]

    # Filter 2: Minimum count threshold
    if min_c > 0:
        service_data = service_data[service_data['Total'] >= min_c]

    # Filter 3: Distributed (at least 3 different offices AND 2 districts)
    service_data = service_data[
        (service_data['Num_Offices'] >= 3) &
        (service_data['Num_Districts'] >= 2)
        ]

    # Filter 4: High OOT rate (above state average)
    state_avg_oot = _df[
                        (_df['month_dt'].dt.year == y) &
                        (_df['month_dt'].dt.month == m)
                        ]['OOT'].sum() / _df[
                        (_df['month_dt'].dt.year == y) &
                        (_df['month_dt'].dt.month == m)
                        ]['Total'].sum() * 100

    service_data = service_data[service_data['OOT_Rate'] > state_avg_oot]

    # Calculate concentration score (lower = more distributed, better)
    # Max office share of service's total OOT
    service_concentration = []
    for service in service_data['Service']:
        svc_offices = _df[
            (_df['Service'] == service) &
            (_df['month_dt'].dt.year == y) &
            (_df['month_dt'].dt.month == m)
            ].groupby('Office')['OOT'].sum()

        if len(svc_offices) > 0:
            max_share = svc_offices.max() / svc_offices.sum() * 100
        else:
            max_share = 0
        service_concentration.append(max_share)

    service_data['Max_Office_Share'] = service_concentration
    service_data['State_Avg_OOT'] = state_avg_oot

    # Sort by OOT_Rate descending, then by distribution (lower concentration first)
    service_data = service_data.sort_values(
        ['OOT_Rate', 'Max_Office_Share'],
        ascending=[False, True]
    )

    return service_data


# ═════════════════════════════════════════════════════════════════════════════
# LAYOUT
# ═════════════════════════════════════════════════════════════════════════════
# 🆕 Generating Descending Period Options (Combines Year & Month)
_periods = sorted(_df['month_dt'].dropna().unique(), reverse=True) if len(_df) else []
period_options = [
    {'label': f"{MONTH_NAMES[pd.Timestamp(p).month]} {pd.Timestamp(p).year}",
     'value': pd.Timestamp(p).strftime('%Y-%m-%d')}
    for p in _periods
]

layout = html.Div([
    html.Div([
        html.H2("📊 Key Findings & District Performance",
                style={'color': 'white', 'margin': '0', 'fontSize': '1.6rem'}),
        html.P("Analyze worst-performing districts based on composite score | Score: 100 = best, 0 = worst",
               style={'color': '#c8dff0', 'margin': '4px 0 0 0', 'fontSize': '0.9rem'}),
    ], style={'background': 'linear-gradient(90deg,#1a3c5e,#2d6a9f)',
              'padding': '18px 28px', 'borderRadius': '10px', 'marginBottom': '20px'}),

    dbc.Row([
        dbc.Col([
            html.Label("📊 Analysis Type", style={'fontWeight': 'bold'}),
            dcc.Dropdown(
                id='findings-type',
                options=[
                    {'label': '🏛️ District', 'value': 'district'},
                    {'label': '🏢 Office', 'value': 'office'},
                    {'label': '⚙️ Service', 'value': 'service'}
                ],
                value='district',
                clearable=False,
                style={'borderRadius': '4px'}
            )
        ], md=3),
        # 🆕 Replaced Year/Month with a single descending Period dropdown
        dbc.Col([
            html.Label("📅 Date Range", style={'fontWeight': 'bold'}),
            dcc.Dropdown(
                id='findings-period',
                options=period_options,
                value=period_options[0]['value'] if period_options else None,
                clearable=False,
                style={'borderRadius': '4px'}
            )
        ], md=4),
        dbc.Col([
            html.Label("⚙️ Min Total", style={'fontWeight': 'bold'}),
            dcc.Input(
                id='findings-min-count',
                type='number',
                value=100,
                min=0,
                style={'width': '100%', 'padding': '6px', 'borderRadius': '4px', 'border': '1px solid #ccc'}
            )
        ], md=2),
        dbc.Col([
            html.Br(),
            dbc.Button(
                "🚀 Go Now",
                id='findings-go-btn',
                color='primary',
                size='lg',
                style={'width': '100%', 'fontWeight': 'bold'}
            ),
            dbc.RadioItems(
                id='findings-mode',
                options=[
                    {'label': ' Report ', 'value': 'report'},
                    {'label': ' Detail ', 'value': 'detail'}
                ],
                value='report',
                inline=True,
                style={'marginTop': '10px', 'fontWeight': 'bold', 'textAlign': 'center'}
            )
        ], md=3),
    ], className='mb-4'),

    html.Div(id='findings-output', style={'marginTop': '20px'}),

    dbc.Modal([
        dbc.ModalHeader(dbc.ModalTitle(id="trend-modal-title"), close_button=True),
        dbc.ModalBody(dcc.Graph(id="trend-modal-graph")),
    ], id="trend-modal", size="lg", is_open=False),

], style={'padding': '20px'})


# ═════════════════════════════════════════════════════════════════════════════
# CALLBACKS
# ═════════════════════════════════════════════════════════════════════════════

# Main findings callback
@app.callback(
    Output('findings-output', 'children'),
    Input('findings-go-btn', 'n_clicks'),
    State('findings-type', 'value'),
    State('findings-period', 'value'),  # 🆕 Uses the new Period value
    State('findings-min-count', 'value'),
    State('findings-mode', 'value'),
    prevent_initial_call=True,
)
def generate_findings(n_clicks, analysis_type, period, min_count, mode):
    if not period:
        return html.Div("⚠️ Please select a period.",
                        style={'color': 'red', 'fontSize': '1.1rem', 'padding': '20px'})

    # Extract Year and Month from the period selection
    selected_dt = pd.to_datetime(period)
    year, month = selected_dt.year, selected_dt.month
    month_name = MONTH_NAMES.get(month, str(month))

    if analysis_type == 'district':
        # Calculate district scores
        district_df = _calculate_district_scores(year, month, min_count or 0, None)

        if district_df.empty:
            return html.Div("No data available for the selected period.",
                            style={'color': 'orange', 'fontSize': '1.1rem', 'padding': '20px'})

        if mode == 'detail':
            display_df = district_df.sort_values('Composite_Score')
            header_text = f"📋 All Districts Performance — {month_name} {year}"
        else:
            display_df = district_df.nsmallest(5, 'Composite_Score')
            header_text = f"🔴 Worst 5 Districts — {month_name} {year}"

        # Create output
        cards = []

        for i in range(len(display_df)):
            row = display_df.iloc[i]

            # Determine color based on score
            score = row['Composite_Score']
            if score < 40:
                color = '#d32f2f'  # Red
                bg_color = '#ffebee'
            elif score < 60:
                color = '#f57c00'  # Orange
                bg_color = '#fff3e0'
            else:
                color = '#fbc02d'  # Yellow
                bg_color = '#fffde7'

            card = dbc.Card([
                dbc.CardBody([
                    html.H4(row['District'],
                            style={'color': color, 'fontWeight': 'bold', 'marginBottom': '10px'}),
                    dbc.Row([
                        dbc.Col([
                            html.Div([
                                html.Span("Composite Score: ", style={'fontWeight': '600'}),
                                html.Span(f"{score:.2f}",
                                          style={'fontSize': '1.3rem', 'fontWeight': 'bold', 'color': color}),
                            ]),
                            html.Div([
                                html.Span(f"Total Offices: {row['Total_Offices']}",
                                          style={'fontSize': '0.9rem', 'color': '#666'}),
                                html.Span(f" | Above Avg: {row['Above_Avg_Offices']}",
                                          style={'fontSize': '0.9rem', 'color': '#666'}),
                            ], style={'marginTop': '5px'}),
                        ], md=8),
                        dbc.Col([
                            dbc.Button(
                                "📊 More Info",
                                id={'type': 'more-info-btn', 'index': i},
                                color='info',
                                size='sm',
                                style={'width': '100%'}
                            )
                        ], md=4),
                    ]),
                    dbc.Collapse(
                        id={'type': 'more-info-collapse', 'index': i},
                        is_open=False,
                        children=html.Div(id={'type': 'more-info-content', 'index': i})
                    ),
                ])
            ], style={'marginBottom': '15px', 'borderLeft': f'5px solid {color}', 'backgroundColor': bg_color})

            cards.append(card)

        office_data_store = html.Div(
            id='findings-office-data-store',
            children=display_df.to_json(date_format='iso', orient='split'),
            style={'display': 'none'}
        )

        return html.Div([
            html.H3(header_text,
                    style={'color': '#1a3c5e', 'marginBottom': '20px', 'borderBottom': '3px solid #2d6a9f',
                           'paddingBottom': '10px'}),
            html.Div(cards),
            office_data_store
        ])

    elif analysis_type == 'office':
        # 🆕 NEW Office-Level Analysis: Get 10 Worst Offices Statewide + Top 1 Worst Service
        snap = _OM[(_OM['_y'] == year) & (_OM['_m'] == month)].copy()

        if min_count and min_count > 0:
            snap = snap[snap['Total'] >= min_count]

        if snap.empty:
            return html.Div("No office data available for the selected criteria.",
                            style={'color': 'orange', 'fontSize': '1.1rem', 'padding': '20px'})

        # Get worst 10 by highest OOT Rate
        worst_10 = snap.sort_values(['OOT_Rate', 'Total'], ascending=[False, False]).head(10)

        cards = []
        for _, row in worst_10.iterrows():
            off = row['Office']
            dist = row['District']
            score_val = row['OOT_Rate']
            total_apps = row['Total']

            # Find Top 1 Worst Service for this specific office
            osm_snap = _OSM[(_OSM['Office'] == off) & (_OSM['_y'] == year) & (_OSM['_m'] == month) & (_OSM['OOT'] > 0)]
            if not osm_snap.empty:
                worst_svc_row = osm_snap.loc[osm_snap['OOT'].idxmax()]  # Getting the service with most OOT applications
                svc_text = f"{worst_svc_row['Service']} ({int(worst_svc_row['OOT'])} delayed | {worst_svc_row['OOT_Rate']:.1f}% OOT)"
            else:
                svc_text = "No Out of Time Services recorded"

            # Color coding
            if score_val > 70:
                color = '#d32f2f'  # Red
                bg_color = '#ffebee'
            elif score_val > 50:
                color = '#f57c00'  # Orange
                bg_color = '#fff3e0'
            else:
                color = '#fbc02d'  # Yellow
                bg_color = '#fffde7'

            card = dbc.Card([
                dbc.CardBody([
                    html.H4(f"🏢 {off}", style={'color': color, 'fontWeight': 'bold', 'marginBottom': '5px'}),
                    html.H6(f"🏛️ District: {dist}", style={'color': '#555', 'marginBottom': '10px'}),
                    dbc.Row([
                        dbc.Col([
                            html.Div([
                                html.Span("OOT Rate: ", style={'fontWeight': '600'}),
                                html.Span(f"{score_val:.2f}%",
                                          style={'fontSize': '1.3rem', 'fontWeight': 'bold', 'color': color}),
                            ]),
                            html.Div([
                                html.Span(f"Total Applications processed: {int(total_apps):,}",
                                          style={'fontSize': '0.9rem', 'color': '#666'}),
                            ], style={'marginTop': '5px'}),
                            html.Div([
                                html.Span("⚠️ Worst Service: ", style={'fontWeight': 'bold', 'color': '#c0392b'}),
                                html.Span(svc_text, style={'fontSize': '0.9rem', 'color': '#333', 'fontWeight': '500'}),
                            ], style={'marginTop': '10px', 'padding': '8px', 'backgroundColor': 'rgba(255,0,0,0.05)',
                                      'borderRadius': '4px'}),
                        ], md=12),
                    ]),
                ])
            ], style={'marginBottom': '15px', 'borderLeft': f'5px solid {color}', 'backgroundColor': bg_color})
            cards.append(card)

        return html.Div([
            html.H3(f"🔴 Worst 10 Offices Statewide — {month_name} {year}",
                    style={'color': '#1a3c5e', 'marginBottom': '20px', 'borderBottom': '3px solid #2d6a9f',
                           'paddingBottom': '10px'}),
            html.Div(cards),
        ])

    elif analysis_type == 'service':
        # Calculate service scores
        service_df = _calculate_service_scores(year, month, min_count or 0)

        if service_df.empty:
            return html.Div("No distributed high-OOT services found for the selected period.",
                            style={'color': 'orange', 'fontSize': '1.1rem', 'padding': '20px'})

        # Get top 5 services
        top_5_services = service_df.head(5)

        # Create output cards
        cards = []
        for idx, row in top_5_services.iterrows():
            score_val = row['OOT_Rate']

            # Color coding
            if score_val > 70:
                color = '#d32f2f'  # Red
                bg_color = '#ffebee'
            elif score_val > 50:
                color = '#f57c00'  # Orange
                bg_color = '#fff3e0'
            else:
                color = '#fbc02d'  # Yellow
                bg_color = '#fffde7'

            card = dbc.Card([
                dbc.CardBody([
                    html.H4(row['Service'],
                            style={'color': color, 'fontWeight': 'bold', 'marginBottom': '10px'}),
                    dbc.Row([
                        dbc.Col([
                            html.Div([
                                html.Span("OOT Rate: ", style={'fontWeight': '600'}),
                                html.Span(f"{score_val:.2f}%",
                                          style={'fontSize': '1.3rem', 'fontWeight': 'bold', 'color': color}),
                            ]),
                            html.Div([
                                html.Span(f"Total Apps: {int(row['Total']):,} | ",
                                          style={'fontSize': '0.9rem', 'color': '#666'}),
                                html.Span(f"Districts: {int(row['Num_Districts'])} | ",
                                          style={'fontSize': '0.9rem', 'color': '#666'}),
                                html.Span(f"Offices: {int(row['Num_Offices'])}",
                                          style={'fontSize': '0.9rem', 'color': '#666'}),
                            ], style={'marginTop': '5px'}),
                            html.Div([
                                html.Span(f"Max Office Concentration: {row['Max_Office_Share']:.1f}%",
                                          style={'fontSize': '0.85rem', 'color': '#888', 'fontStyle': 'italic'}),
                            ], style={'marginTop': '3px'}),
                        ], md=12),
                    ]),
                ])
            ], style={'marginBottom': '15px', 'borderLeft': f'5px solid {color}', 'backgroundColor': bg_color})

            cards.append(card)

        return html.Div([
            html.H3(f"🔴 Top 5 Distributed High-OOT Services — {month_name} {year}",
                    style={'color': '#1a3c5e', 'marginBottom': '20px', 'borderBottom': '3px solid #2d6a9f',
                           'paddingBottom': '10px'}),
            html.P([
                "Services with: ",
                html.Strong("Above-average applications"), ", ",
                html.Strong(f"OOT rate > {service_df['State_Avg_OOT'].iloc[0]:.2f}% (state avg)"), ", ",
                html.Strong("Distributed across ≥3 offices & ≥2 districts")
            ], style={'color': '#555', 'fontSize': '0.95rem', 'marginBottom': '15px'}),
            html.Div(cards),
        ])


# More Info button callback
@app.callback(
    Output({'type': 'more-info-collapse', 'index': ALL}, 'is_open'),
    Output({'type': 'more-info-content', 'index': ALL}, 'children'),
    Input({'type': 'more-info-btn', 'index': ALL}, 'n_clicks'),
    State({'type': 'more-info-collapse', 'index': ALL}, 'is_open'),
    State('findings-office-data-store', 'children'),
    State('findings-mode', 'value'),
    prevent_initial_call=True,
)
def toggle_more_info(n_clicks, is_open, office_data_json, mode):
    import json
    from dash import no_update

    ctx = callback_context

    if not ctx.triggered or ctx.triggered[0]['value'] is None:
        return no_update, no_update

    if not office_data_json:
        return no_update, no_update

    district_df = pd.read_json(office_data_json, orient='split')
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    button_idx = json.loads(button_id)['index']

    new_is_open = [not is_open[i] if i == button_idx else is_open[i] for i in range(len(is_open))]
    contents = []

    for idx in range(len(is_open)):
        if new_is_open[idx]:
            row = district_df.iloc[idx]
            office_data = row['office_data']
            office_df = pd.DataFrame(office_data)

            if mode == 'detail':
                display_offices = office_df.sort_values('Composite_Score')
                table_title = f"🏢 All Offices in {row['District']}"
                title_color = '#2d6a9f'
            else:
                display_offices = office_df.nsmallest(3, 'Composite_Score')
                table_title = f"🔻 Worst 3 Offices in {row['District']}"
                title_color = '#c0392b'

            table_header = [
                html.Thead(html.Tr([
                    html.Th("Office", style={'backgroundColor': '#1a3c5e', 'color': 'white'}),
                    html.Th("Composite Score", style={'backgroundColor': '#1a3c5e', 'color': 'white'}),
                    html.Th("OOT Rate (%)", style={'backgroundColor': '#1a3c5e', 'color': 'white'}),
                    html.Th("Total Applications", style={'backgroundColor': '#1a3c5e', 'color': 'white'}),
                    html.Th("Streak", style={'backgroundColor': '#1a3c5e', 'color': 'white'}),
                    html.Th("Trend", style={'backgroundColor': '#1a3c5e', 'color': 'white', 'textAlign': 'center'}),
                ]))
            ]

            table_rows = []
            for _, office in display_offices.iterrows():
                score = office['Composite_Score']
                if score < 40:
                    row_color = '#ffebee'
                elif score < 60:
                    row_color = '#fff3e0'
                else:
                    row_color = '#fffde7'

                table_rows.append(html.Tr([
                    html.Td(office['Office'], style={'fontWeight': '600'}),
                    html.Td(f"{score:.2f}",
                            style={'fontWeight': 'bold', 'color': '#d32f2f' if score < 40 else '#f57c00'}),
                    html.Td(f"{office['OOT_Rate']:.2f}%"),
                    html.Td(f"{int(office['Total']):,}"),
                    html.Td(f"{int(office['Streak'])}"),
                    html.Td(
                        dbc.Button(
                            "➕",
                            id={'type': 'office-trend-btn', 'district': row['District'], 'office': office['Office']},
                            color="secondary",
                            outline=True,
                            size="sm",
                            style={'border': 'none', 'fontSize': '1.2rem', 'padding': '0'}
                        ),
                        style={'textAlign': 'center'}
                    ),
                ], style={'backgroundColor': row_color}))

            table_body = [html.Tbody(table_rows)]

            content = html.Div([
                html.Hr(),
                html.H5(table_title,
                        style={'color': title_color, 'marginTop': '15px', 'marginBottom': '10px'}),
                dbc.Table(table_header + table_body, bordered=True, hover=True,
                          responsive=True, striped=True, size='sm'),
            ], style={'padding': '10px', 'backgroundColor': '#fafafa', 'borderRadius': '5px'})

            contents.append(content)
        else:
            contents.append(html.Div())

    return new_is_open, contents


@app.callback(
    Output('trend-modal', 'is_open'),
    Output('trend-modal-title', 'children'),
    Output('trend-modal-graph', 'figure'),
    Input({'type': 'office-trend-btn', 'district': ALL, 'office': ALL}, 'n_clicks'),
    State('findings-period', 'value'),  # 🆕 Uses new Period state
    prevent_initial_call=True
)
def open_trend_modal(n_clicks, period):
    import json
    from dash import no_update

    ctx = callback_context
    if not ctx.triggered or not any(n_clicks):
        return no_update, no_update, no_update

    prop_id = ctx.triggered[0]['prop_id']
    button_id = json.loads(prop_id.rsplit('.', 1)[0])
    district = button_id['district']
    office = button_id['office']

    # 🆕 Extract Year and Month directly from the period variable
    selected_dt = pd.to_datetime(period)
    year, month = selected_dt.year, selected_dt.month

    end_date = pd.Timestamp(year=year, month=month, day=1)
    start_date = end_date - pd.DateOffset(months=5)

    hist = _OM[
        (_OM['District'] == district) &
        (_OM['Office'] == office) &
        (_OM['month_dt'] >= start_date) &
        (_OM['month_dt'] <= end_date)
        ].copy()

    dist_avg = _DM[
        (_DM['District'] == district) &
        (_DM['month_dt'] >= start_date) &
        (_DM['month_dt'] <= end_date)
        ].copy()

    if hist.empty:
        fig = px.line(title="No data available")
    else:
        hist = hist[['month_dt', 'OOT_Rate']].rename(columns={'OOT_Rate': 'Office_OOT_Rate'})
        dist_avg = dist_avg[['month_dt', 'OOT_Rate']].rename(columns={'OOT_Rate': 'District_OOT_Rate'})

        plot_df = hist.merge(dist_avg, on='month_dt', how='left')
        plot_df = plot_df.sort_values('month_dt')
        plot_df['Month_Str'] = plot_df['month_dt'].dt.strftime('%b %Y')

        fig = px.line(
            plot_df,
            x='Month_Str',
            y=['Office_OOT_Rate', 'District_OOT_Rate'],
            markers=True,
            labels={
                'Month_Str': 'Month',
                'value': 'OOT Rate (%)',
                'variable': 'Metric'
            },
            title=f"OOT % Trend - Last 6 Months"
        )

        new_names = {
            'Office_OOT_Rate': f'{office} OOT %',
            'District_OOT_Rate': f'{district} District OOT %'
        }
        fig.for_each_trace(lambda t: t.update(name=new_names.get(t.name, t.name)))

        fig.update_layout(
            xaxis_title="",
            yaxis_title="OOT Rate (%)",
            yaxis_ticksuffix="%",
            template="plotly_white",
            margin=dict(l=20, r=20, t=40, b=20),
            legend_title_text=""
        )
        fig.update_traces(hovertemplate='%{y:.2f}%')

    title = f"Trend Analysis: {office} ({district})"

    return True, title, fig