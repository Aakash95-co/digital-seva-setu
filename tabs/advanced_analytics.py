import dash
import io
import json
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import html, dcc, Input, Output, State, ALL, callback_context
import dash_bootstrap_components as dbc
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors as rl_colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table as RLTable,
    TableStyle, HRFlowable, KeepTogether
)
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import base64
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

SUNBURST_COLORS = ['#FFAE57', '#FF7853', '#EA5151', '#CC3F57', '#9A2555',
                   '#2d6a9f', '#1abc9c', '#e67e22', '#3498db', '#8e44ad']
SUNBURST_BG = '#2E2733'

STREAK_TOOLTIP = (
    "Bad-Month Streak: Consecutive months (ending at selected month) where "
    "this office OOT% exceeded the district OOT% for that month.\n"
    "📌 1-2 = Minor  |  🔔 3-5 = Watch  |  ⚠️ 6-8 = Sustained  |  🚨 ≥9 = Critical"
)

COMPOSITE_TOOLTIP = (
    "Score (100=best, 0=worst)\n"
    "Composite = F1 + F2 + F3 + F4 (Max 25 pts each)\n"
    "F1 (District): [(Max_Dist_OOT - Office_OOT) / (Max_Dist_OOT - Min_Dist_OOT)] × 25\n"
    "F2 (State): 12.5 + ((State_Avg - Office_OOT) / State_Avg) × 12.5. (Match State=12.5, 0% OOT=25, ≥2x State=0)\n"
    "F3 (Streak): No streak=25, ≥3=16.6, ≥6=8.3, ≥9=0\n"
    "F4 (Concentration): Top service OOT <80% of Office Total = 25, else 0"
)


# ═════════════════════════════════════════════════════════════════════════════
# DATA PREPARATION  (runs once at import)
# ═════════════════════════════════════════════════════════════════════════════
def _prepare_df(raw):
    if raw is None or raw.empty:
        return pd.DataFrame(columns=['District', 'Office', 'Service', 'OOT', 'Total', 'month_dt'])

    df = raw.copy()
    df.columns = df.columns.str.strip()

    if 'Total' in df.columns and 'application_Disposed' in df.columns:
        df.drop(columns=['Total'], inplace=True)

    _map = {
        'District_name': 'District', 'Office_name': 'Office',
        'Service_name': 'Service',
        'application_Disposed_Out_of_time': 'OOT',
        'application_Disposed': 'Total',
    }
    df.rename(columns={k: v for k, v in _map.items() if k in df.columns}, inplace=True)

    if 'month_dt' not in df.columns:
        if 'Year' in df.columns and 'Month' in df.columns:
            df['month_dt'] = pd.to_datetime(
                df['Year'].astype(str) + '-' + df['Month'].astype(str).str.zfill(2) + '-01')
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
# PRE-COMPUTED CACHES  (all groupby done once)
# ═════════════════════════════════════════════════════════════════════════════
def _oot_pct(s):
    return np.where(s['Total'] > 0, (s['OOT'] / s['Total'] * 100).round(2), 0.0)


def _build_caches(df):
    if df.empty:
        empty = pd.DataFrame(columns=['OOT', 'Total', 'OOT_Rate', 'month_dt'])
        return empty.copy(), empty.copy(), empty.copy(), empty.copy(), empty.copy()

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

    # add pre-computed year/month int columns for fast filtering
    for frame in (om, dm, sm, ssm, osm):
        frame['_y'] = frame['month_dt'].dt.year
        frame['_m'] = frame['month_dt'].dt.month

    return om, dm, sm, ssm, osm


_OM, _DM, _SM, _SSM, _OSM = _build_caches(_df)


# ═════════════════════════════════════════════════════════════════════════════
# FAST LOOKUP HELPERS
# ═════════════════════════════════════════════════════════════════════════════
def _oot_rate(oot, total):
    return round(oot / total * 100, 2) if total > 0 else 0.0


def _dist_oot(district, y, m):
    r = _DM[(_DM['District'] == district) & (_DM['_y'] == y) & (_DM['_m'] == m)]
    return float(r['OOT_Rate'].iloc[0]) if len(r) else 0.0


def _state_oot(y, m):
    r = _SM[(_SM['_y'] == y) & (_SM['_m'] == m)]
    return float(r['OOT_Rate'].iloc[0]) if len(r) else 0.0


def _state_svc_avg(y, m):
    s = _SSM[(_SSM['_y'] == y) & (_SSM['_m'] == m)]
    return dict(zip(s['Service'], s['OOT_Rate']))


# ═════════════════════════════════════════════════════════════════════════════
# STREAK COMPUTATION  (vectorised per district)
# ═════════════════════════════════════════════════════════════════════════════
def _compute_streaks(district, valid_offices, cutoff_ts):
    """Return {office: streak_int} for valid_offices up to cutoff_ts."""
    hist = _OM[
        (_OM['District'] == district) &
        (_OM['Office'].isin(valid_offices)) &
        (_OM['month_dt'] <= cutoff_ts)
        ].copy()
    if hist.empty:
        return {o: 0 for o in valid_offices}

    # district avg per month
    davg = _DM[(_DM['District'] == district) & (_DM['month_dt'] <= cutoff_ts)]
    davg_map = dict(zip(davg['month_dt'], davg['OOT_Rate']))
    hist['d_avg'] = hist['month_dt'].map(davg_map).fillna(0)
    hist['bad'] = hist['OOT_Rate'] > hist['d_avg']
    hist.sort_values(['Office', 'month_dt'], inplace=True)

    streaks = {}
    for off, grp in hist.groupby('Office'):
        flags = grp['bad'].values
        s = 0
        for f in reversed(flags):
            if f:
                s += 1
            else:
                break
        streaks[off] = s
    for o in valid_offices:
        streaks.setdefault(o, 0)
    return streaks


def _service_consistency(office, y, m):
    cutoff = pd.Timestamp(year=y, month=m, day=1)
    oh = _OSM[(_OSM['Office'] == office) & (_OSM['month_dt'] <= cutoff)].copy()
    if oh.empty:
        return {}
    mg = oh.merge(
        _SSM[['Service', 'month_dt', 'OOT_Rate']].rename(columns={'OOT_Rate': 'sa'}),
        on=['Service', 'month_dt'], how='left')
    mg['sa'] = mg['sa'].fillna(0)
    mg['bad'] = mg['OOT_Rate'] > mg['sa']
    mg.sort_values(['Service', 'month_dt'], inplace=True)
    out = {}
    for svc, grp in mg.groupby('Service'):
        flags = grp['bad'].values
        s = 0
        for f in reversed(flags):
            if f:
                s += 1
            else:
                break
        out[svc] = s
    return out


# ═════════════════════════════════════════════════════════════════════════════
# STREAK DISPLAY HELPERS
# ═════════════════════════════════════════════════════════════════════════════
def _streak_label(s):
    if s >= 9: return f"🚨 {s} months (Critical)"
    if s >= 6: return f"⚠️ {s} months (Sustained)"
    if s >= 3: return f"🔔 {s} months (Watch)"
    if s >= 1: return f"📌 {s} month(s) (Minor)"
    return "✅ No streak"


def _streak_mag(s):
    if s >= 9: return "Critical — persistent systemic failure"
    if s >= 6: return "Sustained — serious performance issue"
    if s >= 3: return "Moderate — requires monitoring"
    if s >= 1: return "Minor — early warning"
    return "None"


# ═════════════════════════════════════════════════════════════════════════════
# SCORING ENGINE
# ═════════════════════════════════════════════════════════════════════════════
def _score_offices(district, y, m, min_c, max_c):
    snap = _OM[(_OM['District'] == district) & (_OM['_y'] == y) & (_OM['_m'] == m)].copy()
    if snap.empty:
        return pd.DataFrame(), 0.0

    avg_total = round(snap['Total'].mean(), 2)

    if min_c > 0:
        snap = snap[snap['Total'] >= min_c]
    if max_c and max_c > 0:
        snap = snap[snap['Total'] <= max_c]
    snap = snap.reset_index(drop=True)
    if snap.empty:
        return pd.DataFrame(), avg_total

    offices = snap['Office'].tolist()
    d_oot = _dist_oot(district, y, m)
    s_oot = _state_oot(y, m)
    cutoff_ts = pd.Timestamp(year=y, month=m, day=1)

    # ---------------------------------------------------------
    # F1: District Deviation (Min-Max, Negative Indicator)
    # Worth up to 25 points
    # ---------------------------------------------------------
    d_min = snap['OOT_Rate'].min()
    d_max = snap['OOT_Rate'].max()

    if d_max > d_min:
        snap['F1'] = ((d_max - snap['OOT_Rate']) / (d_max - d_min)) * 25.0
    else:
        # If all offices tie, they all get full points
        snap['F1'] = 25.0

    # ---------------------------------------------------------
    # F2: State Deviation (Min-Max, Negative Indicator)
    # Worth up to 25 points
    # ---------------------------------------------------------
    if s_oot > 0:
        # Difference is positive if the office is BETTER (lower OOT) than state
        state_diff = s_oot - snap['OOT_Rate']

        # Scale: 12.5 points for matching state, +/- based on the difference
        snap['F2'] = 12.5 + ((state_diff / s_oot) * 12.5)

        # Clip the scores so they never go below 0 or above 25
        snap['F2'] = snap['F2'].clip(lower=0.0, upper=25.0)
    else:
        # If the state average is somehow perfectly 0%, any OOT is bad.
        snap['F2'] = np.where(snap['OOT_Rate'] <= 0, 25.0, 0.0)

    # ---------------------------------------------------------
    # F3: Streak Score
    # Worth up to 25 points
    # ---------------------------------------------------------
    streaks = _compute_streaks(district, offices, cutoff_ts)
    snap['Streak'] = snap['Office'].map(streaks).fillna(0).astype(int)

    # Map streak severity to positive points
    # 0 points for critical (>=9), partial points for mid, full 25 points for no streak (<3)
    snap['F3'] = snap['Streak'].apply(
        lambda s: 0.0 if s >= 9 else 8.33 if s >= 6 else 16.67 if s >= 3 else 25.0)

    # ---------------------------------------------------------
    # F4: Service Concentration Score
    # Worth up to 25 points (Binary: 25 or 0)
    # ---------------------------------------------------------
    svc_agg = _df[
        (_df['District'] == district) & (_df['Office'].isin(offices)) &
        (_df['month_dt'].dt.year == y) & (_df['month_dt'].dt.month == m)
        ].groupby(['Office', 'Service'], as_index=False).agg(OOT=('OOT', 'sum'))

    # Evaluate Top Service OOT against the Office's TOTAL OOT
    top_svc_oot = svc_agg.groupby('Office')['OOT'].max().to_dict()

    snap['Top_Svc_OOT'] = snap['Office'].map(top_svc_oot).fillna(0)
    # Share of top service's OOT from total OOT
    snap['Top_Svc_Share'] = np.where(snap['OOT'] > 0, snap['Top_Svc_OOT'] / snap['OOT'], 0)

    # 25 points if worst service accounts for 80%+ of TOTAL OOT (concentrated = good) or if OOT is 0, else 0 points
    snap['F4'] = np.where((snap['Top_Svc_Share'] >= 0.80) | (snap['OOT'] == 0), 25.0, 0.0)

    # ---------------------------------------------------------
    # Final Composite Score (Sum of F1 + F2 + F3 + F4)
    # Maximum possible is 100, minimum is 0
    # ---------------------------------------------------------
    snap['Composite_Score'] = (snap['F1'] + snap['F2'] + snap['F3'] + snap['F4']).round(2)

    # ---------------------------------------------------------
    # Scaled Score (Relative 0-100 normalization based on min/max)
    # ---------------------------------------------------------
    c_min = snap['Composite_Score'].min()
    c_max = snap['Composite_Score'].max()
    if c_max > c_min:
        snap['Scaled_Score'] = ((snap['Composite_Score'] - c_min) / (c_max - c_min) * 100.0).round(2)
    else:
        snap['Scaled_Score'] = 100.0

    snap['District_Avg_OOT'] = d_oot
    snap['State_Avg_OOT'] = s_oot

    # Sort so the worst offices (lowest scores) are at the top
    return snap.sort_values('Composite_Score').reset_index(drop=True), avg_total


# ═════════════════════════════════════════════════════════════════════════════
# REASON BUILDER
# ═════════════════════════════════════════════════════════════════════════════
def _reasons(row):
    dd = row['OOT_Rate'] - row['District_Avg_OOT']
    ds = row['OOT_Rate'] - row['State_Avg_OOT']
    st = int(row['Streak'])
    ts = row.get('Top_Svc_Share', 0)
    arrow = lambda v: f"▲ {abs(v):.1f}% above" if v > 0 else f"▼ {abs(v):.1f}% below"
    items = [
        f"📍 District OOT {row['District_Avg_OOT']:.1f}%: {arrow(dd)} district avg.",
        f"🌐 State OOT {row['State_Avg_OOT']:.1f}%: {arrow(ds)} state avg.",
        f"📅 Consistency: {_streak_label(st)} — {_streak_mag(st)}.",
    ]
    if ts >= 0.80:
        items.append(f"⚠️ Concentrated: top service OOT = {ts * 100:.1f}% of Office Total → F4 penalty applied.")
    else:
        items.append(f"✅ Distributed: top service OOT = {ts * 100:.1f}% of Office Total (no concentration penalty).")
    return items


# ═════════════════════════════════════════════════════════════════════════════
# KPI CARD
# ═════════════════════════════════════════════════════════════════════════════
def _kpi(label, value, color="#1a3c5e", bg="#f0f6ff", border="#2d6a9f", tip=None):
    return html.Div([
        html.Div(str(value), style={'fontSize': '1.4rem', 'fontWeight': '700', 'color': color}),
        html.Div(label, style={
            'fontSize': '0.75rem', 'color': '#555', 'fontWeight': '600',
            'textTransform': 'uppercase', 'letterSpacing': '0.5px'}),
    ], style={
        'background': bg, 'borderLeft': f'5px solid {border}',
        'padding': '12px 10px', 'borderRadius': '8px', 'textAlign': 'center',
        'cursor': 'help' if tip else 'default', 'title': tip or ''})


# ═════════════════════════════════════════════════════════════════════════════
# SUNBURST (Plotly native — no external JS)
# ═════════════════════════════════════════════════════════════════════════════
def _sunburst_figure(district, y, m):
    raw = _df[
        (_df['District'] == district) &
        (_df['month_dt'].dt.year == y) &
        (_df['month_dt'].dt.month == m)
        ].groupby(['Office', 'Service'], as_index=False).agg(OOT=('OOT', 'sum'))
    raw = raw[raw['OOT'] > 0]
    if raw.empty:
        return None

    ids, labels, parents, values, clrs = [], [], [], [], []

    # centre node
    dist_oot = int(raw['OOT'].sum())
    ids.append(district)
    labels.append(district)
    parents.append('')
    values.append(dist_oot)
    clrs.append('#2d6a9f')

    offices = sorted(raw['Office'].unique())
    for idx, office in enumerate(offices):
        oid = f"{district}/{office}"
        off_oot = int(raw.loc[raw['Office'] == office, 'OOT'].sum())
        ids.append(oid)
        labels.append(office)
        parents.append(district)
        values.append(off_oot)
        clrs.append(SUNBURST_COLORS[idx % len(SUNBURST_COLORS)])

        svcs = raw[raw['Office'] == office].sort_values('OOT', ascending=False)
        for _, sr in svcs.iterrows():
            sid = f"{oid}/{sr['Service']}"
            ids.append(sid)
            labels.append(sr['Service'])
            parents.append(oid)
            values.append(int(sr['OOT']))
            clrs.append(SUNBURST_COLORS[idx % len(SUNBURST_COLORS)])

    fig = go.Figure(go.Sunburst(
        ids=ids, labels=labels, parents=parents, values=values,
        branchvalues='total',
        marker=dict(colors=clrs, line=dict(color=SUNBURST_BG, width=2)),
        hovertemplate='<b>%{label}</b><br>OOT: %{value:,}<extra></extra>',
        textinfo='label+value',
        insidetextorientation='radial',
        maxdepth=3,
    ))
    month_name = MONTH_NAMES.get(m, str(m))
    fig.update_layout(
        title=dict(
            text=f"🌐 OOT Sunburst — {district} | {month_name} {y}",
            font=dict(color='#ddd', size=16)),
        paper_bgcolor=SUNBURST_BG,
        plot_bgcolor=SUNBURST_BG,
        margin=dict(t=50, l=10, r=10, b=10),
        height=560,
        font=dict(color='#eee'),
    )
    return fig


# ═════════════════════════════════════════════════════════════════════════════
# PDF GENERATOR (clean, wrapped text, no overflows)
# ═════════════════════════════════════════════════════════════════════════════
def _generate_pdf(year, month, min_c, max_c):
    month_name = MONTH_NAMES.get(month, str(month))
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm)
    sty = getSampleStyleSheet()

    # -- custom styles --
    sTitle = ParagraphStyle('sTitle', parent=sty['Title'], fontSize=14,
                            textColor=rl_colors.HexColor('#1a3c5e'), spaceAfter=4)
    sH2 = ParagraphStyle('sH2', parent=sty['Heading2'], fontSize=11,
                         textColor=rl_colors.HexColor('#2d6a9f'), spaceBefore=10, spaceAfter=2)
    sH3 = ParagraphStyle('sH3', parent=sty['Heading3'], fontSize=9,
                         textColor=rl_colors.HexColor('#c0392b'), spaceBefore=6, spaceAfter=2)
    sBody = ParagraphStyle('sBody', parent=sty['Normal'], fontSize=7.5, spaceAfter=2)
    sSmall = ParagraphStyle('sSmall', parent=sty['Normal'], fontSize=6.5,
                            textColor=rl_colors.HexColor('#555'))
    # cell style: for wrapping inside table cells
    sCell = ParagraphStyle('sCell', parent=sty['Normal'], fontSize=6.5,
                           leading=8, alignment=TA_LEFT)
    sCellC = ParagraphStyle('sCellC', parent=sty['Normal'], fontSize=6.5,
                            leading=8, alignment=TA_CENTER)

    def _p(txt, style=sCell):
        return Paragraph(str(txt), style)

    hdr_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#1a3c5e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.white),
        ('FONTSIZE', (0, 0), (-1, -1), 6.5),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
         [rl_colors.HexColor('#f0f6ff'), rl_colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.4, rl_colors.HexColor('#d0d7de')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ])

    page_w = A4[0] - 2.4 * cm  # usable width

    def _make_tbl(data, col_widths):
        t = RLTable(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(hdr_style)
        return t

    story = []
    story.append(Paragraph(f"Monthly Performance Report — {month_name} {year}", sTitle))
    story.append(Paragraph(
        f"All districts | Min: {min_c} | Max: {max_c or 'None'} | "
        f"Score: 100=best, 0=worst", sBody))
    story.append(HRFlowable(width="100%", thickness=1,
                            color=rl_colors.HexColor('#2d6a9f'), spaceAfter=6))
    story.append(Paragraph(
        "<b>Formula:</b> Score = F1 + F2 + F3 + F4 (Max 100) | "
        "F1=District OOT diff | F2=State OOT diff | F3=Streak penalty | "
        "F4=25 if top service OOT <80% of Office Total else 0", sSmall))
    story.append(Spacer(1, 6))

    svc_avg = _state_svc_avg(year, month)
    districts = sorted(_df['District'].dropna().unique())

    # column widths for office summary table (adjusted to fit 12 columns)
    off_cw = [3.4 * cm, 1.1 * cm, 1 * cm, 1.1 * cm, 1.1 * cm, 1.1 * cm,
              1.1 * cm, 1.1 * cm, 1.2 * cm, 1.2 * cm, 1.1 * cm, 1.1 * cm]
    # column widths for service table (service name gets more space)
    svc_cw = [4 * cm, 1.1 * cm, 1 * cm, 1 * cm, 1.2 * cm, 1.2 * cm, 1.1 * cm, 2.5 * cm, 2.8 * cm]

    for district in districts:
        scored, avg_t = _score_offices(district, year, month, min_c, max_c)
        if scored.empty:
            continue

        d_oot = scored['District_Avg_OOT'].iloc[0]
        s_oot = scored['State_Avg_OOT'].iloc[0]

        story.append(Paragraph(f"District: {district}", sH2))
        story.append(HRFlowable(width="100%", thickness=0.4,
                                color=rl_colors.HexColor('#aaa'), spaceAfter=2))
        story.append(Paragraph(
            f"District OOT: <b>{d_oot:.1f}%</b> | State OOT: <b>{s_oot:.1f}%</b> | "
            f"Avg Total: <b>{avg_t:,.0f}</b> | Offices: <b>{len(scored)}</b>", sBody))
        story.append(Spacer(1, 3))

        # ---------- top-level district table ----------
        off_data = [[
            _p('<b>Office</b>'), _p('<b>Total</b>', sCellC), _p('<b>OOT</b>', sCellC),
            _p('<b>OOT%</b>', sCellC), _p('<b>Dist%</b>', sCellC),
            _p('<b>State%</b>', sCellC), _p('<b>vs Dist</b>', sCellC),
            _p('<b>vs State</b>', sCellC), _p('<b>Streak</b>', sCellC),
            _p('<b>TopSvc%</b>', sCellC), _p('<b>Score</b>', sCellC),
            _p('<b>Scaled</b>', sCellC),
        ]]
        for _, r in scored.iterrows():
            dd = r['OOT_Rate'] - r['District_Avg_OOT']
            ds = r['OOT_Rate'] - r['State_Avg_OOT']
            off_data.append([
                _p(r['Office']),
                _p(f"{int(r['Total']):,}", sCellC),
                _p(f"{int(r['OOT']):,}", sCellC),
                _p(f"{r['OOT_Rate']:.1f}%", sCellC),
                _p(f"{r['District_Avg_OOT']:.1f}%", sCellC),
                _p(f"{r['State_Avg_OOT']:.1f}%", sCellC),
                _p(f"{'▲' if dd > 0 else '▼'}{abs(dd):.1f}%", sCellC),
                _p(f"{'▲' if ds > 0 else '▼'}{abs(ds):.1f}%", sCellC),
                _p(str(int(r['Streak'])), sCellC),
                _p(f"{r['Top_Svc_Share'] * 100:.1f}%", sCellC),
                _p(f"{r['Composite_Score']:.1f}", sCellC),
                _p(f"{r['Scaled_Score']:.1f}", sCellC),
            ])
        story.append(_make_tbl(off_data, off_cw))
        story.append(Spacer(1, 4))

        # ---------- worst 3 detail ----------
        worst3 = scored.head(min(3, len(scored)))
        for rank, (_, row) in enumerate(worst3.iterrows(), 1):
            streak = int(row['Streak'])
            ts = row.get('Top_Svc_Share', 0)
            dd = row['OOT_Rate'] - row['District_Avg_OOT']
            ds = row['OOT_Rate'] - row['State_Avg_OOT']

            story.append(Paragraph(f"Rank {rank} Worst — {row['Office']}", sH3))

            # Office metrics
            met = [
                ['Metric', 'Value'],
                ['Total', f"{int(row['Total']):,}"],
                ['OOT', f"{int(row['OOT']):,}"],
                ['OOT Rate', f"{row['OOT_Rate']:.1f}%"],
                ['District OOT', f"{row['District_Avg_OOT']:.1f}%"],
                ['State OOT', f"{row['State_Avg_OOT']:.1f}%"],
                ['vs District', f"{'▲' if dd > 0 else '▼'} {abs(dd):.1f}%"],
                ['vs State', f"{'▲' if ds > 0 else '▼'} {abs(ds):.1f}%"],
                ['Bad-Month Streak', str(streak)],
                ['Streak Level', _streak_mag(streak)],
                ['Top Svc / Office Total', f"{ts * 100:.1f}%"],
                ['Concentration', 'Penalised' if ts >= 0.80 else 'OK (distributed)'],
                ['Composite Score', f"{row['Composite_Score']:.1f}"],
                ['Scaled Score', f"{row['Scaled_Score']:.1f}"],
            ]
            met_para = [[_p(c[0]), _p(c[1], sCellC)] for c in met]
            story.append(_make_tbl(met_para, [6 * cm, 5 * cm]))
            story.append(Spacer(1, 3))

            # Service breakdown
            svc_raw = _df[
                (_df['District'] == district) & (_df['Office'] == row['Office']) &
                (_df['month_dt'].dt.year == year) & (_df['month_dt'].dt.month == month)
                ].groupby('Service', as_index=False).agg(Total=('Total', 'sum'), OOT=('OOT', 'sum'))
            svc_raw['OOT_Rate'] = svc_raw.apply(lambda r: _oot_rate(r['OOT'], r['Total']), axis=1)
            svc_raw['StAvg'] = svc_raw['Service'].map(lambda s: svc_avg.get(s, 0.0))
            svc_raw['vs'] = svc_raw['OOT_Rate'] - svc_raw['StAvg']
            t_total = svc_raw['Total'].sum()
            svc_raw['Share'] = svc_raw['OOT'].apply(
                lambda o: round(o / t_total * 100, 1) if t_total > 0 else 0)
            svc_streaks = _service_consistency(row['Office'], year, month)
            svc_raw = svc_raw.sort_values('OOT_Rate', ascending=False)

            svc_data = [[
                _p('<b>Service</b>'), _p('<b>Total</b>', sCellC), _p('<b>OOT</b>', sCellC),
                _p('<b>OOT%</b>', sCellC), _p('<b>St.Avg%</b>', sCellC),
                _p('<b>vs St</b>', sCellC), _p('<b>OOT/Tot%</b>', sCellC),
                _p('<b>Streak</b>', sCellC), _p('<b>Level</b>', sCellC),
            ]]
            for _, sr in svc_raw.iterrows():
                st = svc_streaks.get(sr['Service'], 0)
                svc_data.append([
                    _p(sr['Service']),  # wraps automatically
                    _p(f"{int(sr['Total']):,}", sCellC),
                    _p(f"{int(sr['OOT']):,}", sCellC),
                    _p(f"{sr['OOT_Rate']:.1f}%", sCellC),
                    _p(f"{sr['StAvg']:.1f}%", sCellC),
                    _p(f"{'▲' if sr['vs'] > 0 else '▼'}{abs(sr['vs']):.1f}%", sCellC),
                    _p(f"{sr['Share']:.1f}%", sCellC),
                    _p(_streak_label(st), sCellC),
                    _p(_streak_mag(st), sCellC),
                ])
            story.append(Paragraph("Service Breakdown:", sSmall))
            story.append(KeepTogether([_make_tbl(svc_data, svc_cw)]))
            story.append(Spacer(1, 6))

        story.append(Spacer(1, 8))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═════════════════════════════════════════════════════════════════════════════
# LAYOUT
# ═════════════════════════════════════════════════════════════════════════════
_districts = sorted(_df['District'].dropna().unique()) if len(_df) else []

layout = html.Div([
    html.Div([
        html.H2("🔍 Advanced Analytics & Performance Report",
                style={'color': 'white', 'margin': '0', 'fontSize': '1.6rem'}),
        html.P("Score: 100 = best, 0 = worst  |  Click District OOT KPI for sunburst",
               style={'color': '#c8dff0', 'margin': '4px 0 0 0', 'fontSize': '0.9rem'}),
    ], style={'background': 'linear-gradient(90deg,#1a3c5e,#2d6a9f)',
              'padding': '18px 28px', 'borderRadius': '10px', 'marginBottom': '20px'}),

    dbc.Row([
        dbc.Col([html.Label("🏛️ District"),
                 dcc.Dropdown(id='aa-district',
                              options=[{'label': d, 'value': d} for d in _districts],
                              value=_districts[0] if _districts else None,
                              clearable=False)], md=3),
        dbc.Col([html.Label("📅 Reporting Period"),
                 dcc.Dropdown(id='aa-period', clearable=False)], md=3),
        dbc.Col([html.Label("⚙️ Min Total"),
                 dcc.Input(id='aa-min-count', type='number', value=100, min=0,
                           style={'width': '100%', 'padding': '6px',
                                  'borderRadius': '4px', 'border': '1px solid #ccc'})], md=1),
        dbc.Col([html.Label("⚙️ Max Total"),
                 dcc.Input(id='aa-max-count', type='number', value=None, min=0,
                           placeholder='No limit',
                           style={'width': '100%', 'padding': '6px',
                                  'borderRadius': '4px', 'border': '1px solid #ccc'})], md=1),
        dbc.Col([
            html.Br(),
            dbc.Row([
                dbc.Col(dbc.Button("🔍 Analyse", id='aa-run-btn',
                                   color='primary', style={'width': '100%'}), width=6),
                dbc.Col(dbc.Button("📄 PDF Report", id='aa-pdf-btn', color='warning',
                                   style={'width': '100%', 'fontSize': '0.75rem'}), width=6),
            ])
        ], md=4),
    ], className='mb-4'),

    html.Div(id='aa-pdf-download'),
    dcc.Download(id='aa-pdf-file'),
    html.Div(id='aa-output'),

    # Hidden sunburst modal
    dbc.Modal([
        dbc.ModalHeader(dbc.ModalTitle("🌐 OOT Sunburst — District → Office → Service"),
                        close_button=True),
        dbc.ModalBody(id='aa-sunburst-body'),
    ], id='aa-sunburst-modal', size='xl', is_open=False),

], style={'padding': '20px'})


# ═════════════════════════════════════════════════════════════════════════════
# CALLBACKS
# ═════════════════════════════════════════════════════════════════════════════

# -- Reporting Period dropdown --
@app.callback(
    Output('aa-period', 'options'),
    Output('aa-period', 'value'),
    Input('aa-district', 'value'),
)
def update_periods(district):
    if not district or _df.empty:
        return [], None

    # Get unique datetimes for this district and sort descending
    dates = _OM[_OM['District'] == district]['month_dt'].drop_duplicates().sort_values(ascending=False)

    opts = []
    for d in dates:
        # Build label e.g., "March 2025"
        lbl = f"{MONTH_NAMES.get(d.month, str(d.month))} {d.year}"
        # Value string formatting for parsing later
        val = f"{d.year}-{d.month}"
        opts.append({'label': lbl, 'value': val})

    return opts, opts[0]['value'] if opts else None


# -- PDF --
@app.callback(
    Output('aa-pdf-file', 'data'),
    Output('aa-pdf-download', 'children'),
    Input('aa-pdf-btn', 'n_clicks'),
    State('aa-period', 'value'),
    State('aa-min-count', 'value'), State('aa-max-count', 'value'),
    prevent_initial_call=True,
)
def gen_pdf(_, period, min_c, max_c):
    if not period:
        return None, dbc.Alert("Select Reporting Period.", color="warning")
    try:
        year, month = map(int, period.split('-'))
        pdf = _generate_pdf(year, month,
                            int(min_c or 0), int(max_c) if max_c else None)
        mn = MONTH_NAMES.get(month, str(month))
        fn = f"Report_{mn}_{year}.pdf"
        return dict(content=base64.b64encode(pdf).decode(),
                    filename=fn, base64=True, type='application/pdf'), \
            dbc.Alert(f"✅ {fn}", color="success", duration=4000)
    except Exception as e:
        return None, dbc.Alert(f"❌ {e}", color="danger")


# -- Sunburst modal (opens on District OOT KPI click) --
@app.callback(
    Output('aa-sunburst-modal', 'is_open'),
    Output('aa-sunburst-body', 'children'),
    Input('aa-dist-oot-kpi', 'n_clicks'),
    State('aa-district', 'value'),
    State('aa-period', 'value'),
    prevent_initial_call=True,
)
def open_sunburst(n, district, period):
    if not all([district, period]):
        return False, ""
    year, month = map(int, period.split('-'))
    fig = _sunburst_figure(district, year, month)
    if fig is None:
        return True, dbc.Alert("No OOT data for this selection.", color="warning")
    return True, dcc.Graph(figure=fig, style={'height': '560px'})


# -- Main analysis --
@app.callback(
    Output('aa-output', 'children'),
    Input('aa-run-btn', 'n_clicks'),
    State('aa-district', 'value'), State('aa-period', 'value'),
    State('aa-min-count', 'value'), State('aa-max-count', 'value'),
    prevent_initial_call=True,
)
def run_analysis(_, district, period, min_c, max_c):
    if not all([district, period]):
        return dbc.Alert("Select all filters.", color="warning")

    year, month = map(int, period.split('-'))
    min_c = int(min_c or 0)
    max_c = int(max_c) if max_c else None

    scored, avg_t = _score_offices(district, year, month, min_c, max_c)
    if scored.empty:
        return dbc.Alert(f"No offices in Total [{min_c}–{max_c or '∞'}].", color="warning")

    d_oot = scored['District_Avg_OOT'].iloc[0]
    s_oot = scored['State_Avg_OOT'].iloc[0]
    month_name = MONTH_NAMES.get(month, str(month))

    # ── KPI row (District OOT is clickable → opens sunburst) ──────────────
    kpi_row = dbc.Row([
        dbc.Col(_kpi("Offices", str(len(scored))), md=2),
        dbc.Col(_kpi("Avg Total (pre-filter)", f"{avg_t:,.0f}",
                     color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=2),
        dbc.Col(
            html.Div(
                _kpi("District OOT  🔍click",
                     f"{d_oot:.1f}%",
                     color="#c0392b", bg="#fff5f5", border="#e74c3c",
                     tip="Click to open sunburst: District → Office → Service by OOT"),
                id='aa-dist-oot-kpi',
                style={'cursor': 'pointer'},
                n_clicks=0,
            ), md=2),
        dbc.Col(_kpi("State OOT", f"{s_oot:.1f}%",
                     color="#7d3c98", bg="#fdf5ff", border="#9b59b6"), md=2),
        dbc.Col(_kpi("≥3 Bad Months", str(int((scored['Streak'] >= 3).sum())),
                     color="#e67e22", bg="#fff8f0", border="#e67e22"), md=2),
        dbc.Col(_kpi("Concentrated", str(int((scored['Top_Svc_Share'] >= 0.80).sum())),
                     color="#c0392b", bg="#fff5f5", border="#e74c3c",
                     tip="Offices where one service's OOT is ≥80% of Office Total"), md=2),
    ], className='mb-3')

    info = dbc.Alert([
        html.Strong(f"📊 {district} | {month_name} {year}"), html.Br(),
        f"Pre-filter avg total: ", html.Strong(f"{avg_t:,.1f}"),
        f" | Filter: [{min_c}, {max_c or '∞'}]",
    ], color="info", className="mb-3")

    formula = dbc.Alert([
        html.Strong("📐 Formula  "),
        html.Code("Composite Score = F1 + F2 + F3 + F4  (Max 100 points)"), html.Br(),
        html.Small([
            html.B("F1 (25 pts)"), " Dist Min-Max: [(Max_Dist − Office) / (Max_Dist − Min_Dist)] × 25  |  ",
            html.B("F2 (25 pts)"), " State Baseline: 12.5 + ((State_OOT − Office_OOT) / State_OOT) × 12.5  |  ",
            html.B("F3 (25 pts)"), " Streak Score: <3=25, ≥3=16.6, ≥6=8.3, ≥9=0  |  ",
            html.B("F4 (25 pts)"), " Service Score: Top Svc OOT <80% of Total = 25, else 0  |  ",
            html.B("100=best"),
        ], style={'color': '#555', 'fontSize': '0.82rem'})
    ], color="light", className="mb-3", style={'border': '1px solid #d0d7de'})

    # ── Table ─────────────────────────────────────────────────────────────
    col_names = ['Office', 'Total', 'Out of Time', 'OOT Rate(%)', 'Bad-Month Streak ⓘ', 'Score(100=best) ⓘ',
                 'Scaled Score']
    thead = html.Thead(html.Tr([html.Th(col) for col in col_names]))

    tbody_rows = []
    for _, row in scored.iterrows():
        f1, f2, f3, f4 = row['F1'], row['F2'], row['F3'], row['F4']
        comp = row['Composite_Score']
        tooltip_text = f"F1: {f1:.2f}\nF2: {f2:.2f}\nF3: {f3:.2f}\nF4: {f4:.2f}"

        tbody_rows.append(html.Tr([
            html.Td(row['Office']),
            html.Td(int(row['Total'])),
            html.Td(int(row['OOT'])),
            html.Td(f"{row['OOT_Rate']:.2f}"),
            html.Td(int(row['Streak'])),
            html.Td(html.Span(f"{comp:.2f}", title=tooltip_text,
                              style={'cursor': 'help', 'borderBottom': '1px dotted #2d6a9f'})),
            html.Td(f"{row['Scaled_Score']:.2f}")
        ]))
    tbody = html.Tbody(tbody_rows)

    table_sec = html.Div([
        html.H4(f"📊 Office Performance — {month_name} {year} | {district}", className='mb-1'),
        html.Small([
            "Sorted worst→best.  ",
            html.Span("ⓘ Bad-Month Streak", style={'cursor': 'help', 'textDecoration': 'underline dotted',
                                                   'color': '#e67e22'},
                      title=STREAK_TOOLTIP),
            "  |  ",
            html.Span("ⓘ Composite Score", style={'cursor': 'help', 'textDecoration': 'underline dotted',
                                                  'color': '#2d6a9f'},
                      title=COMPOSITE_TOOLTIP),
        ], style={'color': '#888'}),
        dbc.Table([thead, tbody], striped=True, bordered=True, hover=True, responsive=True, size='sm')
    ], className='mb-4')

    # ── Bar chart ─────────────────────────────────────────────────────────
    fig = px.bar(
        scored.sort_values('Composite_Score'),
        x='Office', y='OOT_Rate', color='Composite_Score',
        color_continuous_scale='RdYlGn', range_color=[0, 100],
        title=f"OOT Rate by Office — {month_name} {year} | {district}",
        labels={'OOT_Rate': '% Out of Time', 'Composite_Score': 'Score'},
        text='OOT_Rate')
    fig.add_hline(y=d_oot, line_dash='dash', line_color='#2980b9',
                  annotation_text=f"District: {d_oot:.1f}%", annotation_position="top left")
    fig.add_hline(y=s_oot, line_dash='dot', line_color='#8e44ad',
                  annotation_text=f"State: {s_oot:.1f}%", annotation_position="bottom right")
    fig.update_layout(xaxis_tickangle=-45, height=460)
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    chart_sec = dcc.Graph(figure=fig, className='mb-4')

    # ── Worst 3 ───────────────────────────────────────────────────────────
    def _rank_card(row, rank, wlabel, header_bg, header_clr, streak_bg, streak_border):
        reasons = _reasons(row)
        office = row['Office']
        streak = int(row['Streak'])
        ts = row.get('Top_Svc_Share', 0)
        return dbc.Card([
            dbc.CardHeader(html.Strong(
                f"{wlabel} | {office} | OOT: {row['OOT_Rate']:.1f}% | "
                f"Score: {row['Composite_Score']:.1f}/100"),
                style={'background': header_bg, 'color': header_clr}),
            dbc.CardBody([
                dbc.Row([
                    dbc.Col(_kpi("Total", f"{int(row['Total']):,}",
                                 color="#1a3c5e", bg="#eaf4ff", border="#2d6a9f"), md=2),
                    dbc.Col(_kpi("OOT Rate", f"{row['OOT_Rate']:.1f}%",
                                 color=header_clr, bg=header_bg,
                                 border=header_clr.replace('#fff', '#e74')), md=2),
                    dbc.Col(_kpi("Dist OOT", f"{row['District_Avg_OOT']:.1f}%",
                                 color="#2471a3", bg="#eaf4ff", border="#3498db"), md=2),
                    dbc.Col(_kpi("State OOT", f"{row['State_Avg_OOT']:.1f}%",
                                 color="#7d3c98", bg="#fdf5ff", border="#9b59b6"), md=2),
                    dbc.Col(html.Div([
                        html.Div(_streak_label(streak),
                                 style={'fontSize': '1rem', 'fontWeight': '700', 'color': streak_border}),
                        html.Div("Bad-Month Streak",
                                 style={'fontSize': '0.72rem', 'color': '#555',
                                        'fontWeight': '600', 'textTransform': 'uppercase'}),
                        html.Div(_streak_mag(streak),
                                 style={'fontSize': '0.72rem', 'color': '#888'}),
                    ], style={'background': streak_bg, 'borderLeft': f'5px solid {streak_border}',
                              'padding': '10px', 'borderRadius': '8px', 'textAlign': 'center',
                              'title': STREAK_TOOLTIP, 'cursor': 'help'}), md=2),
                    dbc.Col(_kpi("Top Svc Share", f"{ts * 100:.1f}%",
                                 color="#e74c3c" if ts >= 0.80 else "#27ae60",
                                 bg="#fff5f5" if ts >= 0.80 else "#f0fff4",
                                 border="#e74c3c" if ts >= 0.80 else "#27ae60",
                                 tip="≥80% of Total Received → F4 concentration penalty"), md=2),
                ], className='mb-2'),
                html.Strong("📝 Reason:"),
                html.Ul([html.Li(r) for r in reasons]),
                dbc.Button("🔎 Service Breakdown",
                           id={'type': 'aa-detail-btn', 'index': office},
                           color='outline-danger', size='sm', className='mt-2'),
                dbc.Collapse(
                    html.Div(id={'type': 'aa-detail-panel', 'index': office}),
                    id={'type': 'aa-detail-collapse', 'index': office}, is_open=False),
            ])
        ], className='mb-3')

    labels_w = ["🥇 Rank 1 — Worst", "🥈 Rank 2", "🥉 Rank 3"]
    worst_cards = [
        _rank_card(row, i, labels_w[i], '#fff0f0', '#c0392b', '#fff8f0', '#e67e22')
        for i, (_, row) in enumerate(scored.head(min(3, len(scored))).iterrows())
    ]
    worst_sec = html.Div([
        html.Div([
            html.H3("⚠️ Worst Performing Offices", style={'margin': '0', 'color': '#c0392b'}),
            html.P("Lowest composite score first.",
                   style={'margin': '4px 0', 'color': '#7f8c8d', 'fontSize': '0.88rem'}),
        ], style={'background': '#fff0f0', 'border': '1.5px solid #e74c3c',
                  'borderRadius': '10px', 'padding': '14px 22px', 'marginBottom': '12px'}),
        *worst_cards
    ], className='mb-4')

    # ── Best 3 ────────────────────────────────────────────────────────────
    labels_b = ["🥇 Best", "🥈 2nd Best", "🥉 3rd Best"]
    best_rows = scored.tail(min(3, len(scored))).iloc[::-1].reset_index(drop=True)
    best_cards = [
        _rank_card(row, i, labels_b[i], '#f0fff4', '#1e8449', '#f0fff4', '#27ae60')
        for i, (_, row) in enumerate(best_rows.iterrows())
    ]
    best_sec = html.Div([
        html.Div([
            html.H3("🏆 Best Performing Offices", style={'margin': '0', 'color': '#1e8449'}),
            html.P("Highest composite score.", style={'margin': '4px 0', 'color': '#7f8c8d', 'fontSize': '0.88rem'}),
        ], style={'background': '#f0fff4', 'border': '1.5px solid #27ae60',
                  'borderRadius': '10px', 'padding': '14px 22px', 'marginBottom': '12px'}),
        *best_cards
    ], className='mb-4')

    note = html.Div([
        html.Strong("ℹ️ Methodology: "),
        f"Filter Total∈[{min_c},{max_c or '∞'}]. "
        f"F1/F2 = magnitude of office OOT vs district/state snapshot. "
        f"F4 = 25 if top service OOT <80% of Office Total else 0. "
        f"Normalised 0–100 then inverted. 100=best."
    ], style={'background': '#f8f9fa', 'border': '1px solid #d0d7de',
              'borderRadius': '8px', 'padding': '12px 18px', 'color': '#555', 'fontSize': '0.82rem'})

    return html.Div([kpi_row, info, formula, table_sec, chart_sec,
                     worst_sec, best_sec, note])


# ═════════════════════════════════════════════════════════════════════════════
# SERVICE DETAIL (expand/collapse)
# ═════════════════════════════════════════════════════════════════════════════
@app.callback(
    Output({'type': 'aa-detail-collapse', 'index': ALL}, 'is_open'),
    Output({'type': 'aa-detail-panel', 'index': ALL}, 'children'),
    Input({'type': 'aa-detail-btn', 'index': ALL}, 'n_clicks'),
    State({'type': 'aa-detail-collapse', 'index': ALL}, 'is_open'),
    State('aa-district', 'value'),
    State('aa-period', 'value'),
    prevent_initial_call=True,
)
def toggle_detail(n_list, is_open_list, district, period):
    if not callback_context.triggered_id:
        return is_open_list, [dash.no_update] * len(is_open_list)

    office = callback_context.triggered_id['index']
    y, m = map(int, period.split('-'))
    all_ids = [t['id']['index'] for t in callback_context.inputs_list[0]]
    idx = all_ids.index(office)

    new_open = list(is_open_list)
    new_open[idx] = not is_open_list[idx]
    panels = [dash.no_update] * len(is_open_list)

    if not new_open[idx]:
        return new_open, panels

    svc_avg = _state_svc_avg(y, m)
    snap = _df[
        (_df['District'] == district) & (_df['Office'] == office) &
        (_df['month_dt'].dt.year == y) & (_df['month_dt'].dt.month == m)
        ].groupby('Service', as_index=False).agg(Total=('Total', 'sum'), OOT=('OOT', 'sum'))

    snap['OOT_Rate'] = snap.apply(lambda r: _oot_rate(r['OOT'], r['Total']), axis=1)
    snap['StAvg'] = snap['Service'].map(lambda s: svc_avg.get(s, 0.0))
    snap['vs'] = snap['OOT_Rate'] - snap['StAvg']
    t_total = snap['Total'].sum()
    snap['Share'] = snap['OOT'].apply(
        lambda o: round(o / t_total * 100, 1) if t_total > 0 else 0)
    snap.sort_values('OOT_Rate', ascending=False, inplace=True)
    svc_str = _service_consistency(office, y, m)

    rows = []
    for _, sr in snap.iterrows():
        st = svc_str.get(sr['Service'], 0)
        bad = sr['vs'] > 0
        conc = sr['Share'] >= 80
        bg = '#fff5f5' if (bad or conc) else '#f0fff4'
        rows.append(html.Tr([
            html.Td(sr['Service'], style={'fontWeight': '600'}),
            html.Td(f"{int(sr['Total']):,}", style={'textAlign': 'center'}),
            html.Td(f"{int(sr['OOT']):,}", style={'textAlign': 'center'}),
            html.Td(f"{sr['OOT_Rate']:.1f}%", style={'textAlign': 'center',
                                                     'color': '#c0392b' if bad else '#1e8449', 'fontWeight': '700'}),
            html.Td(f"{sr['StAvg']:.1f}%", style={'textAlign': 'center'}),
            html.Td(f"{'▲' if bad else '▼'} {abs(sr['vs']):.1f}%",
                    style={'textAlign': 'center', 'color': '#c0392b' if bad else '#1e8449',
                           'fontWeight': '600'}),
            html.Td(f"{sr['Share']:.1f}%", style={'textAlign': 'center',
                                                  'color': '#e74c3c' if conc else '#555',
                                                  'fontWeight': '700' if conc else '400'}),
            html.Td(_streak_label(st), style={'textAlign': 'center'}),
            html.Td(_streak_mag(st), style={'textAlign': 'center',
                                            'fontSize': '0.78rem', 'color': '#555'}),
        ], style={'background': bg}))

    panels[idx] = html.Div([
        html.H5(f"📋 Service Breakdown — {office}",
                style={'color': '#1a3c5e', 'marginTop': '14px', 'marginBottom': '6px'}),
        html.P("OOT% vs state service avg. OOT/Tot % = Service OOT / Office Total. "
               "Red if ≥80% (F4 penalty). Streak = months office svc > state svc avg.",
               style={'fontSize': '0.82rem', 'color': '#666', 'marginBottom': '6px'}),
        dbc.Table([
            html.Thead(html.Tr([
                html.Th("Service"), html.Th("Total", style={'textAlign': 'center'}),
                html.Th("OOT", style={'textAlign': 'center'}),
                html.Th("OOT Rate", style={'textAlign': 'center'}),
                html.Th("State Avg", style={'textAlign': 'center'}),
                html.Th("vs State", style={'textAlign': 'center'}),
                html.Th("OOT/Tot %", style={'textAlign': 'center'}),
                html.Th("Streak", style={'textAlign': 'center'}),
                html.Th("Level", style={'textAlign': 'center'}),
            ]), style={'background': '#1a3c5e', 'color': 'white'}),
            html.Tbody(rows)
        ], bordered=True, hover=True, responsive=True, size='sm')
    ], style={'padding': '10px', 'background': '#fafafa',
              'border': '1px solid #ddd', 'borderRadius': '8px'})

    return new_open, panels