import streamlit as st
import io
import pandas as pd
import plotly.express as px
from dash import html, dcc, Input, Output, State, ALL, callback_context
import dash_bootstrap_components as dbc
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.units import cm
import base64
from app import app
from data import df_adv

# ─────────────────────────────────────────────────────────────────────────────
# Helper: compute OOT rate safely
# ─────────────────────────────────────────────────────────────────────────────
def _oot_rate(oot, total):
    return round((oot / total * 100), 2) if total > 0 else 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Prepare df_adv
# ─────────────────────────────────────────────────────────────────────────────
def _prepare_df(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=['District', 'Office', 'Service', 'OOT', 'Total', 'month_dt'])

    df = df.copy()
    df.columns = df.columns.str.strip()

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
            if isinstance(df[col], pd.DataFrame):
                df = df.loc[:, ~df.columns.duplicated()]
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        else:
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
# State-level service averages
# ─────────────────────────────────────────────────────────────────────────────
def _state_service_avg(df, year, month):
    snap = df[
        (df['month_dt'].dt.year  == year) &
        (df['month_dt'].dt.month == month)
    ]
    svc = snap.groupby('Service').agg(OOT=('OOT','sum'), Total=('Total','sum')).reset_index()
    svc['State_Svc_OOT_Rate'] = svc.apply(lambda r: _oot_rate(r['OOT'], r['Total']), axis=1)
    return svc.set_index('Service')['State_Svc_OOT_Rate'].to_dict()


# ─────────────────────────────────────────────────────────────────────────────
# Service-level consistency for an office
# ─────────────────────────────────────────────────────────────────────────────
def _service_consistency(df, office, year, month):
    """For each service in an office, compute streak of months above state avg OOT."""
    cutoff      = pd.Timestamp(year=year, month=month, day=1)
    office_hist = df[df['Office'] == office].copy()

    results = []
    for svc, grp in office_hist.groupby('Service'):
        grp = grp.sort_values('month_dt')
        for _, row in grp[grp['month_dt'] <= cutoff].iterrows():
            m, y2 = row['month_dt'].month, row['month_dt'].year
            state_snap = df[
                (df['month_dt'].dt.year  == y2) &
                (df['month_dt'].dt.month == m)  &
                (df['Service'] == svc)
            ]
            state_avg = _oot_rate(state_snap['OOT'].sum(), state_snap['Total'].sum())
            results.append({
                'Service':    svc,
                'month_dt':   row['month_dt'],
                'OOT_Rate':   _oot_rate(row['OOT'], row['Total']),
                'State_Avg':  state_avg,
                'is_bad':     _oot_rate(row['OOT'], row['Total']) > state_avg
            })

    if not results:
        return {}

    hist_df = pd.DataFrame(results)
    streaks = {}
    for svc, grp in hist_df.groupby('Service'):
        flags = grp.sort_values('month_dt')['is_bad'].tolist()
        streak = 0
        for f in reversed(flags):
            if f: streak += 1
            else: break
        streaks[svc] = streak
    return streaks


# ─────────────────────────────────────────────────────────────────────────────
# Scoring engine
# ─────────────────────────────────────────────────────────────────────────────
def _score_offices(df, selected_district, selected_year, selected_month, min_count):
    snap = df[
        (df['District']           == selected_district) &
        (df['month_dt'].dt.year   == selected_year)     &
        (df['month_dt'].dt.month  == selected_month)
    ]
    if snap.empty:
        return pd.DataFrame(), 0.0

    office_snap = (
        snap.groupby('Office')
        .agg(Total=('Total', 'sum'), OOT=('OOT', 'sum'))
        .reset_index()
    )
    office_snap['OOT_Rate'] = office_snap.apply(
        lambda r: _oot_rate(r['OOT'], r['Total']), axis=1
    )

    district_avg_total = round(office_snap['Total'].mean(), 2)

    if min_count > 0:
        office_snap = office_snap[office_snap['Total'] >= min_count].reset_index(drop=True)

    if office_snap.empty:
        return pd.DataFrame(), district_avg_total

    district_avg_oot = _oot_rate(office_snap['OOT'].sum(), office_snap['Total'].sum())

    state_snap    = df[
        (df['month_dt'].dt.year  == selected_year) &
        (df['month_dt'].dt.month == selected_month)
    ]
    state_avg_oot = _oot_rate(state_snap['OOT'].sum(), state_snap['Total'].sum())

    valid_offices = office_snap['Office'].tolist()
    dist_history  = (
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
            if flag: streak += 1
            else:    break
        consistency_scores[office] = streak

    office_snap['Streak'] = office_snap['Office'].map(consistency_scores).fillna(0).astype(int)

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
# Streak label helper
# ─────────────────────────────────────────────────────────────────────────────
def _streak_label(s):
    if s >= 9:  return f"🚨 {s} months (Critical)"
    if s >= 6:  return f"⚠️ {s} months (Sustained)"
    if s >= 3:  return f"🔔 {s} months (Watch)"
    if s >= 1:  return f"📌 {s} month(s) (Minor)"
    return "✅ No streak"

def _streak_magnitude(s):
    if s >= 9:  return "Critical — persistent systemic failure"
    if s >= 6:  return "Sustained — serious performance issue"
    if s >= 3:  return "Moderate — requires monitoring"
    if s >= 1:  return "Minor — early warning"
    return "None"


# ─────────────────────────────────────────────────────────────────────────────
# Reason builder
# ─────────────────────────────────────────────────────────────────────────────
def _build_reason_items(row):
    items       = []
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

    items.append(f"📅 Consistency: {_streak_label(streak)} — {_streak_magnitude(streak)}.")
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
# PDF Generator
# ─────────────────────────────────────────────────────────────────────────────
def _generate_pdf(year, month, min_count):
    month_name = MONTH_NAMES.get(month, str(month))
    buffer     = io.BytesIO()
    doc        = SimpleDocTemplate(buffer, pagesize=A4,
                                   leftMargin=1.5*cm, rightMargin=1.5*cm,
                                   topMargin=1.5*cm,  bottomMargin=1.5*cm)
    styles     = getSampleStyleSheet()
    story      = []

    title_style = ParagraphStyle('title', parent=styles['Title'],
                                 fontSize=16, textColor=colors.HexColor('#1a3c5e'),
                                 spaceAfter=6)
    h2_style    = ParagraphStyle('h2', parent=styles['Heading2'],
                                 fontSize=13, textColor=colors.HexColor('#2d6a9f'),
                                 spaceBefore=14, spaceAfter=4)
    h3_style    = ParagraphStyle('h3', parent=styles['Heading3'],
                                 fontSize=11, textColor=colors.HexColor('#c0392b'),
                                 spaceBefore=10, spaceAfter=3)
    body_style  = ParagraphStyle('body', parent=styles['Normal'],
                                 fontSize=9, spaceAfter=3)
    label_style = ParagraphStyle('label', parent=styles['Normal'],
                                 fontSize=8, textColor=colors.HexColor('#555555'))

    def tbl(data, col_widths=None):
        t = Table(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND',  (0,0), (-1,0), colors.HexColor('#1a3c5e')),
            ('TEXTCOLOR',   (0,0), (-1,0), colors.white),
            ('FONTSIZE',    (0,0), (-1,0), 8),
            ('FONTSIZE',    (0,1), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1),
             [colors.HexColor('#f0f6ff'), colors.white]),
            ('GRID',        (0,0), (-1,-1), 0.4, colors.HexColor('#d0d7de')),
            ('ALIGN',       (1,0), (-1,-1), 'CENTER'),
            ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING',  (0,0), (-1,-1), 3),
            ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ]))
        return t

    # ── Cover ────────────────────────────────────────────────────────────────
    story.append(Paragraph(
        f"Monthly Performance Report — {month_name} {year}", title_style))
    story.append(Paragraph(
        f"Generated for all districts  |  Min. Office Threshold: {min_count}  |  "
        f"Metric: Out-of-Time (OOT) Disposal Rate", body_style))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor('#2d6a9f'), spaceAfter=10))

    # State-level service averages
    state_svc_avg = _state_service_avg(_df, year, month)

    districts = sorted(_df['District'].dropna().unique())

    for district in districts:
        scored, dist_avg_total = _score_offices(_df, district, year, month, min_count)
        if scored.empty:
            continue

        story.append(Paragraph(f"District: {district}", h2_style))
        story.append(HRFlowable(width="100%", thickness=0.5,
                                 color=colors.HexColor('#aaaaaa'), spaceAfter=4))

        dist_avg_oot   = scored['District_Avg_OOT'].iloc[0]
        state_avg_oot  = scored['State_Avg_OOT'].iloc[0]

        story.append(Paragraph(
            f"District OOT Avg: <b>{dist_avg_oot:.1f}%</b>  |  "
            f"State OOT Avg: <b>{state_avg_oot:.1f}%</b>  |  "
            f"Avg Total Applications: <b>{dist_avg_total:,.0f}</b>  |  "
            f"Offices analysed: <b>{len(scored)}</b>", body_style))
        story.append(Spacer(1, 6))

        worst3 = scored.head(3)
        for rank, (_, row) in enumerate(worst3.iterrows(), 1):
            streak    = int(row['Streak'])
            delta_d   = row['OOT_Rate'] - row['District_Avg_OOT']
            delta_s   = row['OOT_Rate'] - row['State_Avg_OOT']

            story.append(Paragraph(
                f"Rank {rank} — {row['Office']}", h3_style))

            # Office summary table
            off_data = [
                ['Metric', 'Value'],
                ['Total Applications',       f"{int(row['Total']):,}"],
                ['Out-of-Time (OOT)',         f"{int(row['OOT']):,}"],
                ['Office OOT Rate',           f"{row['OOT_Rate']:.1f}%"],
                ['District Avg OOT',          f"{row['District_Avg_OOT']:.1f}%"],
                ['State Avg OOT',             f"{row['State_Avg_OOT']:.1f}%"],
                ['vs District Avg',           f"{'▲' if delta_d>0 else '▼'} {abs(delta_d):.1f}%"],
                ['vs State Avg',              f"{'▲' if delta_s>0 else '▼'} {abs(delta_s):.1f}%"],
                ['Consecutive Bad Months',    f"{streak}"],
                ['Consistency Magnitude',     _streak_magnitude(streak)],
            ]
            story.append(tbl(off_data, col_widths=[7*cm, 7*cm]))
            story.append(Spacer(1, 6))

            # Service-level breakdown for this office
            snap = _df[
                (_df['District']          == district) &
                (_df['Office']            == row['Office']) &
                (_df['month_dt'].dt.year  == year) &
                (_df['month_dt'].dt.month == month)
            ]
            svc_snap = (
                snap.groupby('Service')
                .agg(Total=('Total','sum'), OOT=('OOT','sum'))
                .reset_index()
            )
            svc_snap['OOT_Rate']      = svc_snap.apply(
                lambda r: _oot_rate(r['OOT'], r['Total']), axis=1)
            svc_snap['State_Svc_Avg'] = svc_snap['Service'].map(
                lambda s: state_svc_avg.get(s, 0.0))
            svc_snap['vs_State']      = svc_snap['OOT_Rate'] - svc_snap['State_Svc_Avg']

            # Consistency streaks per service
            svc_streaks = _service_consistency(_df, row['Office'], year, month)

            svc_rows = [['Service', 'Total', 'OOT', 'OOT%',
                          'State Svc Avg%', 'vs State', 'Streak', 'Magnitude']]
            for _, sr in svc_snap.sort_values('OOT_Rate', ascending=False).iterrows():
                st = svc_streaks.get(sr['Service'], 0)
                svc_rows.append([
                    sr['Service'],
                    f"{int(sr['Total']):,}",
                    f"{int(sr['OOT']):,}",
                    f"{sr['OOT_Rate']:.1f}%",
                    f"{sr['State_Svc_Avg']:.1f}%",
                    f"{'▲' if sr['vs_State']>0 else '▼'} {abs(sr['vs_State']):.1f}%",
                    _streak_label(st),
                    _streak_magnitude(st),
                ])
            story.append(Paragraph("Service-level Breakdown:", label_style))
            story.append(tbl(svc_rows,
                              col_widths=[4.5*cm,1.5*cm,1.2*cm,1.3*cm,
                                          2*cm,1.5*cm,3*cm,3.5*cm]))
            story.append(Spacer(1, 10))

        story.append(Spacer(1, 14))

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


# ─────────────────────────────────────────────────────────────────────────────
# Layout
# ─────────────────────────────────────────────────────────────────────────────
districts = sorted(_df['District'].dropna().unique()) if not _df.empty else []
years     = sorted(_df['month_dt'].dt.year.unique())  if not _df.empty else []

layout = html.Div([
    html.Div([
        html.H2("🔍 Advanced Analytics & Performance Report",
                style={'color': 'white', 'margin': '0',
                       'fontSize': '1.6rem', 'letterSpacing': '1px'}),
        html.P("Identify worst-performing offices using district benchmarking, "
               "state averages, and consistency analysis.",
               style={'color': '#c8dff0', 'margin': '4px 0 0 0', 'fontSize': '0.95rem'}),
    ], style={
        'background': 'linear-gradient(90deg, #1a3c5e 0%, #2d6a9f 100%)',
        'padding': '18px 28px', 'borderRadius': '10px', 'marginBottom': '20px'
    }),

    dbc.Row([
        dbc.Col([
            html.Label("🏛️ Select District"),
            dcc.Dropdown(id='aa-district',
                         options=[{'label': d, 'value': d} for d in districts],
                         value=districts[0] if districts else None,
                         clearable=False)
        ], md=3),
        dbc.Col([
            html.Label("📅 Select Year"),
            dcc.Dropdown(id='aa-year',
                         options=[{'label': y, 'value': y} for y in years],
                         value=years[-1] if years else None,
                         clearable=False)
        ], md=2),
        dbc.Col([
            html.Label("🗓️ Select Month"),
            dcc.Dropdown(id='aa-month', clearable=False)
        ], md=3),
        dbc.Col([
            html.Label("⚙️ Min. Office Total Threshold"),
            dcc.Input(id='aa-min-count', type='number', value=100, min=0,
                      style={'width': '100%', 'padding': '6px',
                             'borderRadius': '4px', 'border': '1px solid #ccc'})
        ], md=2),
        dbc.Col([
            html.Br(),
            dbc.Row([
                dbc.Col(
                    dbc.Button("🔍 Analyse", id='aa-run-btn',
                               color='primary', style={'width': '100%'}), width=6),
                dbc.Col(
                    dbc.Button("📄 Monthly General Prediction",
                               id='aa-pdf-btn', color='warning',
                               style={'width': '100%', 'fontSize': '0.78rem'}), width=6),
            ])
        ], md=2),
    ], className='mb-4'),

    # PDF download link (hidden until generated)
    html.Div(id='aa-pdf-download'),
    dcc.Download(id='aa-pdf-file'),

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
    Output('aa-pdf-file',     'data'),
    Output('aa-pdf-download', 'children'),
    Input('aa-pdf-btn',  'n_clicks'),
    State('aa-year',     'value'),
    State('aa-month',    'value'),
    State('aa-min-count','value'),
    prevent_initial_call=True,
)
def generate_pdf(n_clicks, year, month, min_count):
    if not year or not month:
        return None, dbc.Alert("Please select Year and Month first.", color="warning")
    try:
        pdf_bytes  = _generate_pdf(int(year), int(month), int(min_count or 0))
        b64        = base64.b64encode(pdf_bytes).decode()
        month_name = MONTH_NAMES.get(int(month), str(month))
        filename   = f"Monthly_Report_{month_name}_{year}.pdf"
        return (
            dict(content=b64, filename=filename, base64=True, type='application/pdf'),
            dbc.Alert(f"✅ PDF ready: {filename}", color="success", duration=4000)
        )
    except Exception as e:
        return None, dbc.Alert(f"❌ PDF generation failed: {e}", color="danger")


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

    min_count = int(min_count or 0)
    scored, district_avg_total = _score_offices(
        _df, district, int(year), int(month), min_count)

    if scored.empty:
        return dbc.Alert(
            f"No offices with Total ≥ {min_count} found for the selected filters.",
            color="warning")

    district_avg_oot = scored['District_Avg_OOT'].iloc[0]
    state_avg_oot    = scored['State_Avg_OOT'].iloc[0]
    month_name       = MONTH_NAMES.get(int(month), str(month))
    state_svc_avg    = _state_service_avg(_df, int(year), int(month))

    # ── KPI bar ──────────────────────────────────────────────────────────────
    kpi_row = dbc.Row([
        dbc.Col(_kpi_card("Offices Analysed",        len(scored)), md=2),
        dbc.Col(_kpi_card("Avg Total Applications",  f"{district_avg_total:,.1f}",
                          color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=2),
        dbc.Col(_kpi_card("Min Count Threshold",     f"≥ {min_count}",
                          color="#555",    bg="#f8f9fa", border="#aaa"), md=2),
        dbc.Col(_kpi_card("District OOT Avg",        f"{district_avg_oot:.1f}%",
                          color="#c0392b", bg="#fff5f5", border="#e74c3c"), md=2),
        dbc.Col(_kpi_card("State OOT Avg",           f"{state_avg_oot:.1f}%",
                          color="#7d3c98", bg="#fdf5ff", border="#9b59b6"), md=2),
        dbc.Col(_kpi_card("Flagged (≥3 months)",     int((scored['Streak'] >= 3).sum()),
                          color="#e67e22", bg="#fff8f0", border="#e67e22"), md=2),
    ], className='mb-4')

    threshold_banner = dbc.Alert([
        html.Strong(f"📊 District: {district}  |  {month_name} {year}"),
        html.Br(),
        f"Average total applications (all offices before filtering): ",
        html.Strong(f"{district_avg_total:,.1f}"),
        (f"  —  Offices with Total < {min_count} excluded."
         if min_count > 0 else "  —  No threshold applied.")
    ], color="info", className="mb-4")

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
                                 bordered=True, hover=True,
                                 responsive=True, size='sm')
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
                  annotation_text=f"District Avg: {district_avg_oot:.1f}%",
                  annotation_position="top left")
    fig.add_hline(y=state_avg_oot, line_dash='dot', line_color='#8e44ad',
                  annotation_text=f"State Avg: {state_avg_oot:.1f}%",
                  annotation_position="bottom right")
    fig.update_layout(xaxis_tickangle=-45, height=480)
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    chart_section = dcc.Graph(figure=fig, className='mb-4')

    # ── Worst offices with More Detail button ────────────────────────────────
    worst_labels = ["🥇 Rank 1 — Worst", "🥈 Rank 2", "🥉 Rank 3"]
    worst_items  = []
    for rank, (_, row) in enumerate(scored.head(min(3, len(scored))).iterrows(), 1):
        reasons = _build_reason_items(row)
        office  = row['Office']
        streak  = int(row['Streak'])

        worst_items.append(dbc.Card([
            dbc.CardHeader(html.Strong(
                f"{worst_labels[rank-1]}  |  {office}  |  "
                f"OOT: {row['OOT_Rate']:.1f}%  |  Score: {row['Composite_Score']:.1f}"
            ), style={'background': '#fff0f0', 'color': '#c0392b'}),
            dbc.CardBody([
                dbc.Row([
                    dbc.Col(_kpi_card("Total Applications", f"{int(row['Total']):,}",
                                      color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=3),
                    dbc.Col(_kpi_card("OOT Rate",           f"{row['OOT_Rate']:.1f}%",
                                      color="#c0392b", bg="#fff5f5", border="#e74c3c"), md=3),
                    dbc.Col(_kpi_card("District Avg OOT",   f"{row['District_Avg_OOT']:.1f}%",
                                      color="#2980b9", bg="#eaf4ff", border="#3498db"), md=3),
                    dbc.Col(_kpi_card("State Avg OOT",      f"{row['State_Avg_OOT']:.1f}%",
                                      color="#7d3c98", bg="#fdf5ff", border="#9b59b6"), md=3),
                ], className='mb-2'),
                dbc.Row([
                    dbc.Col(_kpi_card("Consecutive Bad Months", _streak_label(streak),
                                      color="#e67e22", bg="#fff8f0", border="#e67e22"), md=6),
                    dbc.Col(_kpi_card("Consistency Magnitude",  _streak_magnitude(streak),
                                      color="#7f8c8d", bg="#f8f9fa", border="#aaa"), md=6),
                ], className='mb-3'),
                html.Strong("📝 Reason for Flagging:"),
                html.Ul([html.Li(r) for r in reasons]),

                # More Detail button
                dbc.Button(
                    "🔎 More Detail (Service Breakdown)",
                    id={'type': 'aa-detail-btn', 'index': office},
                    color='outline-danger', size='sm', className='mt-2'
                ),
                # Collapsible service detail panel
                dbc.Collapse(
                    html.Div(id={'type': 'aa-detail-panel', 'index': office}),
                    id={'type': 'aa-detail-collapse', 'index': office},
                    is_open=False
                )
            ])
        ], className='mb-3'))

    worst_section = html.Div([
        html.Div([
            html.H3("⚠️ Worst Performing Offices",
                    style={'margin': '0', 'color': '#c0392b'}),
            html.P("Ranked by composite score: district benchmark (35%) + "
                   "state average (35%) + consistency (30%)",
                   style={'margin': '4px 0 0 0',
                          'color': '#7f8c8d', 'fontSize': '0.9rem'}),
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
        streak  = int(row['Streak'])
        best_items.append(dbc.Card([
            dbc.CardHeader(html.Strong(
                f"{best_labels[rank-1]}  |  {row['Office']}  |  "
                f"OOT: {row['OOT_Rate']:.1f}%  |  Score: {row['Composite_Score']:.1f}"
            ), style={'background': '#f0fff4', 'color': '#1e8449'}),
            dbc.CardBody([
                dbc.Row([
                    dbc.Col(_kpi_card("Total Applications", f"{int(row['Total']):,}",
                                      color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=3),
                    dbc.Col(_kpi_card("OOT Rate",           f"{row['OOT_Rate']:.1f}%",
                                      color="#1e8449", bg="#f0fff4", border="#27ae60"), md=3),
                    dbc.Col(_kpi_card("District Avg OOT",   f"{row['District_Avg_OOT']:.1f}%",
                                      color="#2980b9", bg="#eaf4ff", border="#3498db"), md=3),
                    dbc.Col(_kpi_card("State Avg OOT",      f"{row['State_Avg_OOT']:.1f}%",
                                      color="#7d3c98", bg="#fdf5ff", border="#9b59b6"), md=3),
                ], className='mb-2'),
                dbc.Row([
                    dbc.Col(_kpi_card("Consecutive Bad Months", _streak_label(streak),
                                      color="#27ae60", bg="#f0fff4", border="#27ae60"), md=6),
                    dbc.Col(_kpi_card("Consistency Magnitude",  _streak_magnitude(streak),
                                      color="#7f8c8d", bg="#f8f9fa", border="#aaa"), md=6),
                ], className='mb-3'),
                html.Strong("📝 Performance Summary:"),
                html.Ul([html.Li(r) for r in reasons])
            ])
        ], className='mb-3'))

    best_section = html.Div([
        html.Div([
            html.H3("🏆 Best Performing Offices",
                    style={'margin': '0', 'color': '#1e8449'}),
            html.P("Offices with lowest composite score (lowest OOT burden relative to benchmarks)",
                   style={'margin': '4px 0 0 0',
                          'color': '#7f8c8d', 'fontSize': '0.9rem'}),
        ], style={'background': '#f0fff4', 'border': '1.5px solid #27ae60',
                  'borderRadius': '10px', 'padding': '14px 22px', 'marginBottom': '12px'}),
        *best_items
    ], className='mb-4')

    disclaimer = html.Div(
        f"Methodology: Offices with Total < {min_count} excluded. "
        f"Avg total (all offices before filter): {district_avg_total:,.1f}. "
        f"Composite = District 35% · State 35% · Consistency 30%. "
        f"Bad month = office OOT > district avg. Streaks ≥3/6/9 → 🔔/⚠️/🚨.",
        style={'background': '#f8f9fa', 'border': '1px solid #d0d7de',
               'borderRadius': '8px', 'padding': '12px 18px',
               'color': '#555', 'fontSize': '0.82rem'}
    )

    return html.Div([
        kpi_row, threshold_banner, table_section, chart_section,
        worst_section, best_section, disclaimer
    ])


# ─────────────────────────────────────────────────────────────────────────────
# Callback: toggle + populate service detail panel
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output({'type': 'aa-detail-collapse', 'index': ALL}, 'is_open'),
    Output({'type': 'aa-detail-panel',    'index': ALL}, 'children'),
    Input({'type':  'aa-detail-btn',      'index': ALL}, 'n_clicks'),
    State({'type':  'aa-detail-collapse', 'index': ALL}, 'is_open'),
    State('aa-district',  'value'),
    State('aa-year',      'value'),
    State('aa-month',     'value'),
    prevent_initial_call=True,
)
def toggle_detail(n_clicks_list, is_open_list, district, year, month):
    if not callback_context.triggered_id:
        return is_open_list, [dash.no_update] * len(is_open_list)

    office = callback_context.triggered_id['index']
    year   = int(year)
    month  = int(month)

    # Find the index of the triggered office in the ALL list
    all_ids    = [t['id']['index'] for t in callback_context.inputs_list[0]]
    target_idx = all_ids.index(office)

    # Toggle only the triggered collapse; leave others unchanged
    new_open = list(is_open_list)
    new_open[target_idx] = not is_open_list[target_idx]

    # If closing, just return without rebuilding content
    panels = [dash.no_update] * len(is_open_list)
    if not new_open[target_idx]:
        return new_open, panels

    # Build service detail content
    state_svc_avg = _state_service_avg(_df, year, month)

    snap = _df[
        (_df['District']          == district) &
        (_df['Office']            == office)   &
        (_df['month_dt'].dt.year  == year)     &
        (_df['month_dt'].dt.month == month)
    ]
    svc_snap = (
        snap.groupby('Service')
        .agg(Total=('Total', 'sum'), OOT=('OOT', 'sum'))
        .reset_index()
    )
    svc_snap['OOT_Rate']      = svc_snap.apply(
        lambda r: _oot_rate(r['OOT'], r['Total']), axis=1)
    svc_snap['State_Svc_Avg'] = svc_snap['Service'].map(
        lambda s: state_svc_avg.get(s, 0.0))
    svc_snap['vs_State']      = svc_snap['OOT_Rate'] - svc_snap['State_Svc_Avg']
    svc_snap                  = svc_snap.sort_values('OOT_Rate', ascending=False)

    svc_streaks = _service_consistency(_df, office, year, month)

    rows = []
    for _, sr in svc_snap.iterrows():
        st        = svc_streaks.get(sr['Service'], 0)
        is_bad    = sr['vs_State'] > 0
        row_color = '#fff5f5' if is_bad else '#f0fff4'
        rows.append(
            html.Tr([
                html.Td(sr['Service'],                     style={'fontWeight': '600'}),
                html.Td(f"{int(sr['Total']):,}",           style={'textAlign': 'center'}),
                html.Td(f"{int(sr['OOT']):,}",             style={'textAlign': 'center'}),
                html.Td(f"{sr['OOT_Rate']:.1f}%",         style={'textAlign': 'center',
                                                                   'color': '#c0392b' if is_bad else '#1e8449',
                                                                   'fontWeight': '700'}),
                html.Td(f"{sr['State_Svc_Avg']:.1f}%",    style={'textAlign': 'center'}),
                html.Td(
                    f"{'▲' if is_bad else '▼'} {abs(sr['vs_State']):.1f}%",
                    style={'textAlign': 'center',
                           'color': '#c0392b' if is_bad else '#1e8449',
                           'fontWeight': '600'}
                ),
                html.Td(_streak_label(st),                 style={'textAlign': 'center'}),
                html.Td(_streak_magnitude(st),             style={'textAlign': 'center',
                                                                   'fontSize': '0.8rem', 'color': '#555'}),
            ], style={'background': row_color})
        )

    detail_table = html.Div([
        html.H5(f"📋 Service Breakdown — {office}",
                style={'color': '#1a3c5e', 'marginTop': '16px', 'marginBottom': '8px'}),
        html.P(
            "Each service shows its OOT rate vs state-wide average for that service. "
            "Streak = consecutive months this office's OOT for this service exceeded "
            "the state service average.",
            style={'fontSize': '0.85rem', 'color': '#666', 'marginBottom': '8px'}
        ),
        dbc.Table([
            html.Thead(html.Tr([
                html.Th("Service"),
                html.Th("Total",          style={'textAlign': 'center'}),
                html.Th("OOT",            style={'textAlign': 'center'}),
                html.Th("OOT Rate",       style={'textAlign': 'center'}),
                html.Th("State Svc Avg",  style={'textAlign': 'center'}),
                html.Th("vs State Avg",   style={'textAlign': 'center'}),
                html.Th("Streak",         style={'textAlign': 'center'}),
                html.Th("Magnitude",      style={'textAlign': 'center'}),
            ]), style={'background': '#1a3c5e', 'color': 'white'}),
            html.Tbody(rows)
        ], bordered=True, hover=True, responsive=True, size='sm')
    ], style={'padding': '10px', 'background': '#fafafa',
              'border': '1px solid #ddd', 'borderRadius': '8px'})

    panels[target_idx] = detail_table
    return new_open, panels