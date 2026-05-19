import io
import dash
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import html, dcc, Input, Output, State
import dash_bootstrap_components as dbc

from app import app
from data import FY_DATA

# ─────────────────────────────────────────────────────────────────────────────
# STYLE CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
_CARD = {
    'background': 'white', 'border': '1px solid #d0d7de',
    'borderRadius': '10px', 'padding': '22px 26px',
    'marginBottom': '24px', 'boxShadow': '0 2px 8px rgba(0,0,0,0.06)',
}
_H4 = {
    'color': '#1a3c5e', 'borderBottom': '2px solid #2d6a9f',
    'paddingBottom': '8px', 'marginBottom': '16px',
}

# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
layout = html.Div([

    # ── Header banner ────────────────────────────────────────────────────────
    html.Div([
        html.H2("🎯 Insights & Action Centre",
                style={'color': 'white', 'margin': '0', 'fontSize': '1.6rem'}),
        html.P("Executive Scorecard  ·  Quadrant Analysis  ·  Chronic Alerts  ·  Service Health Matrix  ·  State Export",
               style={'color': '#c8dff0', 'margin': '4px 0 0 0', 'fontSize': '0.9rem'}),
    ], style={'background': 'linear-gradient(90deg,#1a3c5e,#2d6a9f)',
              'padding': '18px 28px', 'borderRadius': '10px', 'marginBottom': '24px'}),

    # ── 1. Executive Scorecard ───────────────────────────────────────────────
    html.Div([
        html.H4("📊 Executive Scorecard", style=_H4),
        html.Div(id='ins-scorecard'),
        dcc.Download(id='ins-summary-dl'),
        html.Div(
            dbc.Button("⬇️ Download State Summary (Excel)", id='ins-export-btn',
                       color='primary', outline=True, className='mt-3'),
            style={'textAlign': 'right'},
        ),
    ], style=_CARD),

    # ── 2. Volume vs Efficiency Quadrant ─────────────────────────────────────
    html.Div([
        html.H4("🔵 Volume vs Efficiency Quadrant", style=_H4),
        html.P(
            "Each dot = one office.  X-axis = Applications Received  |  Y-axis = OOT%  |  Dashed lines = medians",
            style={'color': '#666', 'fontSize': '0.88rem', 'marginBottom': '10px'},
        ),
        dcc.Graph(id='ins-quadrant', style={'height': '560px'}),
    ], style=_CARD),

    # ── 3. Chronically Red Offices ───────────────────────────────────────────
    html.Div([
        html.H4("🚨 Chronically Underperforming Offices  (≥ 6 Months Above State Avg)", style=_H4),
        html.P(
            "Offices whose Out-of-Time % has continuously exceeded the state monthly average "
            "for 6 or more consecutive months ending at the latest data month.",
            style={'color': '#666', 'fontSize': '0.88rem', 'marginBottom': '12px'},
        ),
        html.Div(id='ins-chronic'),
    ], style=_CARD),

    # ── 4. Service Health Matrix ─────────────────────────────────────────────
    html.Div([
        html.H4("🗺️ Service Health Matrix  (District × Service — OOT%)", style=_H4),
        html.P(
            "Top 20 services by OOT count shown.  Darker red = higher OOT%.  "
            "Use this to decide: system-level fix vs local/district fix.",
            style={'color': '#666', 'fontSize': '0.88rem', 'marginBottom': '10px'},
        ),
        dcc.Graph(id='ins-heatmap', style={'height': '650px'}),
    ], style=_CARD),

], style={'padding': '20px'})


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK 1 — Executive Scorecard
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(Output('ins-scorecard', 'children'), Input('fy-store', 'data'))
def cb_scorecard(fy):
    df    = FY_DATA[fy]['df'].copy()
    df_tt = FY_DATA[fy]['df_tt'].copy()
    label = FY_DATA[fy]['label']

    total_recv = int(df['Received'].sum())
    total_disp = int(df['Disposed'].sum())
    total_oot  = int(df['Disposed_Out'].sum())
    total_pend = int(df['Pending'].sum())
    oot_pct    = round(total_oot / total_disp * 100, 1) if total_disp > 0 else 0.0

    dist_sum = df_tt.groupby('District_Eng').agg(
        OOT=('Disposed_Out', 'sum'), Disp=('Disposed', 'sum')
    ).reset_index()
    dist_sum['OOT_P'] = np.where(dist_sum['Disp'] > 0,
                                  dist_sum['OOT'] / dist_sum['Disp'] * 100, 0)
    red   = int((dist_sum['OOT_P'] >= 70).sum())
    amber = int(((dist_sum['OOT_P'] >= 40) & (dist_sum['OOT_P'] < 70)).sum())
    green = int((dist_sum['OOT_P'] < 40).sum())

    def _kpi(title, value, sub='', color='#1a3c5e', bg='#eef4fb'):
        return html.Div([
            html.Div(title, style={
                'fontSize': '0.75rem', 'color': '#555',
                'marginBottom': '4px', 'fontWeight': '700',
                'textTransform': 'uppercase', 'letterSpacing': '0.5px',
            }),
            html.Div(str(value), style={
                'fontSize': '1.9rem', 'fontWeight': '800', 'color': color,
            }),
            html.Div(sub, style={'fontSize': '0.73rem', 'color': '#888', 'marginTop': '2px'}),
        ], style={
            'background': bg, 'borderRadius': '10px', 'padding': '18px 16px',
            'textAlign': 'center', 'flex': '1', 'minWidth': '130px',
        })

    return html.Div([
        _kpi("Total Received",     f"{total_recv:,}",  label),
        _kpi("Total Disposed",     f"{total_disp:,}",  label),
        _kpi("Overall OOT %",      f"{oot_pct}%",      f"{total_oot:,} cases",
             color='#c0392b', bg='#fdecea'),
        _kpi("Total Pending",      f"{total_pend:,}",  label,
             color='#e67e22', bg='#fef5e7'),
        _kpi("🔴 Red Districts",   str(red),           "OOT ≥ 70%",
             color='#c0392b', bg='#fdecea'),
        _kpi("🟠 Amber Districts", str(amber),         "40% ≤ OOT < 70%",
             color='#d35400', bg='#fef5e7'),
        _kpi("🟢 Green Districts", str(green),         "OOT < 40%",
             color='#27ae60', bg='#eafaf1'),
    ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '14px'})


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK 2 — Quadrant Chart
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(Output('ins-quadrant', 'figure'), Input('fy-store', 'data'))
def cb_quadrant(fy):
    df = FY_DATA[fy]['df'].copy()

    agg = df.groupby(['Office_Eng', 'District_Eng']).agg(
        Volume=('Received', 'sum'),
        OOT=('Disposed_Out', 'sum'),
        Disposed=('Disposed', 'sum'),
    ).reset_index()
    agg['OOT_Pct'] = np.where(
        agg['Disposed'] > 0, agg['OOT'] / agg['Disposed'] * 100, 0
    ).round(1)

    med_vol = float(agg['Volume'].median())
    med_oot = float(agg['OOT_Pct'].median())

    def _q(row):
        hv = row['Volume']  >= med_vol
        ho = row['OOT_Pct'] >= med_oot
        if  hv and  ho:  return '🔴 High Vol + High OOT (Immediate Intervention)'
        if  hv and not ho: return '🟢 High Vol + Low OOT (Star Performer)'
        if not hv and  ho: return '🟠 Low Vol + High OOT (Structural Issue)'
        return '🔵 Low Vol + Low OOT (Healthy)'

    agg['Quadrant'] = agg.apply(_q, axis=1)

    fig = px.scatter(
        agg,
        x='Volume', y='OOT_Pct',
        color='Quadrant',
        hover_name='Office_Eng',
        hover_data={
            'District_Eng': True, 'Volume': True,
            'OOT_Pct': True, 'Quadrant': False,
        },
        color_discrete_map={
            '🔴 High Vol + High OOT (Immediate Intervention)': '#e74c3c',
            '🟢 High Vol + Low OOT (Star Performer)':          '#2ecc71',
            '🟠 Low Vol + High OOT (Structural Issue)':        '#e67e22',
            '🔵 Low Vol + Low OOT (Healthy)':                  '#3498db',
        },
        labels={'Volume': 'Applications Received', 'OOT_Pct': 'OOT %'},
        title=f"Office Quadrant Analysis — {FY_DATA[fy]['label']}",
    )
    fig.add_vline(x=med_vol, line_dash='dash', line_color='grey', opacity=0.5,
                  annotation_text='Median Volume', annotation_position='top right')
    fig.add_hline(y=med_oot, line_dash='dash', line_color='grey', opacity=0.5,
                  annotation_text='Median OOT%', annotation_position='bottom right')
    fig.update_layout(
        template='plotly_white',
        legend=dict(
            orientation='h', yanchor='bottom', y=-0.32,
            xanchor='center', x=0.5, font=dict(size=11),
        ),
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK 3 — Chronically Red Offices
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(Output('ins-chronic', 'children'), Input('fy-store', 'data'))
def cb_chronic(fy):
    df = FY_DATA[fy]['df'].copy()
    df['Yr'] = df['Yr'].astype(str).str.replace('\ufeff', '', regex=False).str.strip()
    df['Mn'] = pd.to_numeric(df['Mn'], errors='coerce').fillna(0).astype(int)

    monthly = df.groupby(['Yr', 'Mn', 'Office_Eng', 'District_Eng']).agg(
        OOT=('Disposed_Out', 'sum'), Disposed=('Disposed', 'sum')
    ).reset_index()

    state_m = df.groupby(['Yr', 'Mn']).agg(
        st_OOT=('Disposed_Out', 'sum'), st_Disp=('Disposed', 'sum')
    ).reset_index()
    state_m['State_OOT'] = np.where(
        state_m['st_Disp'] > 0, state_m['st_OOT'] / state_m['st_Disp'] * 100, 0
    )

    monthly = monthly.merge(state_m[['Yr', 'Mn', 'State_OOT']], on=['Yr', 'Mn'], how='left')
    monthly['OOT_Pct'] = np.where(
        monthly['Disposed'] > 0, monthly['OOT'] / monthly['Disposed'] * 100, 0
    )
    monthly['above_avg'] = monthly['OOT_Pct'] > monthly['State_OOT']
    monthly['month_dt'] = pd.to_datetime(
        monthly['Yr'] + '-' + monthly['Mn'].astype(str).str.zfill(2) + '-01',
        errors='coerce',
    )
    monthly.sort_values(['Office_Eng', 'month_dt'], inplace=True)

    records = []
    for office, grp in monthly.groupby('Office_Eng'):
        district = grp['District_Eng'].iloc[-1]
        flags    = grp['above_avg'].tolist()
        streak   = 0
        for f in reversed(flags):
            if f:  streak += 1
            else:  break
        if streak >= 6:
            records.append({
                'Office':      office,
                'District':    district,
                'Streak':      streak,
                'Latest OOT%': round(float(grp['OOT_Pct'].iloc[-1]), 1),
                'State Avg%':  round(float(grp['State_OOT'].iloc[-1]), 1),
            })

    if not records:
        return dbc.Alert(
            "✅ No offices found with 6+ consecutive months above the state OOT average. "
            "This is a good sign!",
            color='success',
        )

    df_c = pd.DataFrame(records).sort_values('Streak', ascending=False)

    rows = []
    for _, r in df_c.iterrows():
        badge_bg = '#c0392b' if r['Streak'] >= 9 else '#e67e22'
        rows.append(html.Tr([
            html.Td(dbc.Badge(
                f"⚠️ {r['Streak']} months",
                style={'background': badge_bg, 'fontSize': '0.85rem', 'padding': '5px 10px'},
            )),
            html.Td(html.Strong(r['Office'])),
            html.Td(r['District']),
            html.Td(f"{r['Latest OOT%']}%", style={'color': '#c0392b', 'fontWeight': '700'}),
            html.Td(f"{r['State Avg%']}%", style={'color': '#555'}),
        ]))

    return dbc.Table(
        [
            html.Thead(html.Tr([
                html.Th("Streak"), html.Th("Office"), html.Th("District"),
                html.Th("Latest OOT%"), html.Th("State Avg%"),
            ]), style={'background': '#fdecea'}),
            html.Tbody(rows),
        ],
        bordered=True, hover=True, responsive=True, size='sm',
        style={'fontSize': '0.9rem'},
    )


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK 4 — Service Health Matrix
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(Output('ins-heatmap', 'figure'), Input('fy-store', 'data'))
def cb_heatmap(fy):
    df = FY_DATA[fy]['df'].copy()

    agg = df.groupby(['District_Eng', 'Service_Eng']).agg(
        OOT=('Disposed_Out', 'sum'), Disposed=('Disposed', 'sum')
    ).reset_index()
    agg['OOT_Pct'] = np.where(
        agg['Disposed'] > 0, agg['OOT'] / agg['Disposed'] * 100, 0
    ).round(1)

    top_svc = agg.groupby('Service_Eng')['OOT'].sum().nlargest(20).index.tolist()
    agg = agg[agg['Service_Eng'].isin(top_svc)]

    pivot = agg.pivot_table(
        index='District_Eng', columns='Service_Eng', values='OOT_Pct', aggfunc='mean'
    ).fillna(0)

    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale='Reds',
        hoverongaps=False,
        hovertemplate='District: %{y}<br>Service: %{x}<br>OOT%%: %{z:.1f}%%<extra></extra>',
        colorbar=dict(title='OOT %'),
    ))
    fig.update_layout(
        title=f"Service Health Matrix — {FY_DATA[fy]['label']}  (Top 20 Services by OOT Count)",
        template='plotly_white',
        xaxis=dict(tickangle=-45, tickfont=dict(size=9)),
        yaxis=dict(tickfont=dict(size=10)),
        margin=dict(l=220, b=220, t=60),
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK 5 — One-Click State Summary Export (Excel)
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output('ins-summary-dl', 'data'),
    Input('ins-export-btn', 'n_clicks'),
    State('fy-store', 'data'),
    prevent_initial_call=True,
)
def cb_export(n_clicks, fy):
    if not n_clicks:
        return dash.no_update

    df    = FY_DATA[fy]['df'].copy()
    df_tt = FY_DATA[fy]['df_tt'].copy()
    label = FY_DATA[fy]['label']

    total_recv = int(df['Received'].sum())
    total_disp = int(df['Disposed'].sum())
    total_oot  = int(df['Disposed_Out'].sum())
    total_pend = int(df['Pending'].sum())
    oot_pct    = round(total_oot / total_disp * 100, 2) if total_disp > 0 else 0.0

    # ── Sheet 1 : KPI Summary ────────────────────────────────────────────────
    kpi_df = pd.DataFrame([
        {'KPI': 'Financial Year',   'Value': label},
        {'KPI': 'Total Received',   'Value': total_recv},
        {'KPI': 'Total Disposed',   'Value': total_disp},
        {'KPI': 'Total OOT Cases',  'Value': total_oot},
        {'KPI': 'Overall OOT %',    'Value': f"{oot_pct}%"},
        {'KPI': 'Total Pending',    'Value': total_pend},
    ])

    # ── Sheet 2 : District Summary ────────────────────────────────────────────
    dist_df = df_tt.groupby('District_Eng').agg(
        Received=('Received', 'sum'),
        Disposed=('Disposed', 'sum'),
        OOT_Cases=('Disposed_Out', 'sum'),
        Pending=('Pending', 'sum'),
    ).reset_index().rename(columns={'District_Eng': 'District'})
    dist_df['OOT_%'] = np.where(
        dist_df['Disposed'] > 0,
        (dist_df['OOT_Cases'] / dist_df['Disposed'] * 100).round(2), 0,
    )
    dist_df = dist_df.sort_values('OOT_%', ascending=False)

    # ── Sheet 3 : Service Summary ─────────────────────────────────────────────
    svc_df = df.groupby('Service_Eng').agg(
        Received=('Received', 'sum'),
        Disposed=('Disposed', 'sum'),
        OOT_Cases=('Disposed_Out', 'sum'),
    ).reset_index().rename(columns={'Service_Eng': 'Service'})
    svc_df['OOT_%'] = np.where(
        svc_df['Disposed'] > 0,
        (svc_df['OOT_Cases'] / svc_df['Disposed'] * 100).round(2), 0,
    )
    svc_df = svc_df.sort_values('OOT_%', ascending=False)

    # ── Sheet 4 : Top 20 Worst Offices ────────────────────────────────────────
    off_df = df.groupby(['Office_Eng', 'District_Eng']).agg(
        Received=('Received', 'sum'),
        Disposed=('Disposed', 'sum'),
        OOT_Cases=('Disposed_Out', 'sum'),
        Pending=('Pending', 'sum'),
    ).reset_index().rename(columns={'Office_Eng': 'Office', 'District_Eng': 'District'})
    off_df['OOT_%'] = np.where(
        off_df['Disposed'] > 0,
        (off_df['OOT_Cases'] / off_df['Disposed'] * 100).round(2), 0,
    )
    off_df = off_df[off_df['Received'] > 0].nlargest(20, 'OOT_%')

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        kpi_df.to_excel(writer,  sheet_name='KPI Summary',          index=False)
        dist_df.to_excel(writer, sheet_name='District Summary',      index=False)
        svc_df.to_excel(writer,  sheet_name='Service Summary',       index=False)
        off_df.to_excel(writer,  sheet_name='Top 20 Worst Offices',  index=False)
    buf.seek(0)

    filename = f"State_Summary_{label.replace(' ', '_').replace('-', '_')}.xlsx"
    return dcc.send_bytes(buf.read(), filename=filename)
