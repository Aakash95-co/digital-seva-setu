import dash
import io
import numpy as np
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
import json
from app import app
from data import df_adv

# ─────────────────────────────────────────────────────────────────────────────
# Helper
# ─────────────────────────────────────────────────────────────────────────────
def _oot_rate(oot, total):
    return round((oot / total * 100), 2) if total > 0 else 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Prepare df
# ─────────────────────────────────────────────────────────────────────────────
def _prepare_df(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=['District','Office','Service','OOT','Total','month_dt'])
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
                df['Year'].astype(str) + '-' +
                df['Month'].astype(str).str.zfill(2) + '-01', format='%Y-%m-%d')
        else:
            df['month_dt'] = pd.NaT
    for col in ['District','Office','Service']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    for col in ['OOT','Total']:
        if col in df.columns:
            if isinstance(df[col], pd.DataFrame):
                df = df.loc[:, ~df.columns.duplicated()]
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        else:
            df[col] = 0
    return df


_df = _prepare_df(df_adv)

MONTH_NAMES = {
    1:"January",2:"February",3:"March",4:"April",
    5:"May",6:"June",7:"July",8:"August",
    9:"September",10:"October",11:"November",12:"December"
}

# ─────────────────────────────────────────────────────────────────────────────
# Pre-compute monthly caches ONCE at startup
# ─────────────────────────────────────────────────────────────────────────────
def _build_caches(df):
    off_monthly = (
        df.groupby(['District','Office','month_dt'])
        .agg(Total=('Total','sum'), OOT=('OOT','sum')).reset_index()
    )
    off_monthly['OOT_Rate'] = np.where(
        off_monthly['Total'] > 0,
        (off_monthly['OOT'] / off_monthly['Total'] * 100).round(2), 0.0)

    dist_monthly = (
        df.groupby(['District','month_dt'])
        .agg(Total=('Total','sum'), OOT=('OOT','sum')).reset_index()
    )
    dist_monthly['OOT_Rate'] = np.where(
        dist_monthly['Total'] > 0,
        (dist_monthly['OOT'] / dist_monthly['Total'] * 100).round(2), 0.0)

    state_monthly = (
        df.groupby('month_dt')
        .agg(Total=('Total','sum'), OOT=('OOT','sum')).reset_index()
    )
    state_monthly['OOT_Rate'] = np.where(
        state_monthly['Total'] > 0,
        (state_monthly['OOT'] / state_monthly['Total'] * 100).round(2), 0.0)

    state_svc_monthly = (
        df.groupby(['Service','month_dt'])
        .agg(Total=('Total','sum'), OOT=('OOT','sum')).reset_index()
    )
    state_svc_monthly['OOT_Rate'] = np.where(
        state_svc_monthly['Total'] > 0,
        (state_svc_monthly['OOT'] / state_svc_monthly['Total'] * 100).round(2), 0.0)

    off_svc_monthly = (
        df.groupby(['Office','Service','month_dt'])
        .agg(Total=('Total','sum'), OOT=('OOT','sum')).reset_index()
    )
    off_svc_monthly['OOT_Rate'] = np.where(
        off_svc_monthly['Total'] > 0,
        (off_svc_monthly['OOT'] / off_svc_monthly['Total'] * 100).round(2), 0.0)

    return off_monthly, dist_monthly, state_monthly, state_svc_monthly, off_svc_monthly


(_OFF_MONTHLY, _DIST_MONTHLY, _STATE_MONTHLY,
 _STATE_SVC_MONTHLY, _OFF_SVC_MONTHLY) = _build_caches(_df)


# ─────────────────────────────────────────────────────────────────────────────
# Snapshot averages (selected month only)
# ─────────────────────────────────────────────────────────────────────────────
def _dist_snap_oot(district, year, month):
    row = _DIST_MONTHLY[
        (_DIST_MONTHLY['District']           == district) &
        (_DIST_MONTHLY['month_dt'].dt.year   == year)     &
        (_DIST_MONTHLY['month_dt'].dt.month  == month)
    ]
    return row['OOT_Rate'].iloc[0] if not row.empty else 0.0


def _state_snap_oot(year, month):
    row = _STATE_MONTHLY[
        (_STATE_MONTHLY['month_dt'].dt.year  == year) &
        (_STATE_MONTHLY['month_dt'].dt.month == month)
    ]
    return row['OOT_Rate'].iloc[0] if not row.empty else 0.0


def _state_service_avg(year, month):
    snap = _STATE_SVC_MONTHLY[
        (_STATE_SVC_MONTHLY['month_dt'].dt.year  == year) &
        (_STATE_SVC_MONTHLY['month_dt'].dt.month == month)
    ]
    return snap.set_index('Service')['OOT_Rate'].to_dict()


# ─────────────────────────────────────────────────────────────────────────────
# Service consistency (vs state svc avg streak)
# ─────────────────────────────────────────────────────────────────────────────
def _service_consistency(office, year, month):
    cutoff   = pd.Timestamp(year=year, month=month, day=1)
    off_hist = _OFF_SVC_MONTHLY[
        (_OFF_SVC_MONTHLY['Office']    == office) &
        (_OFF_SVC_MONTHLY['month_dt'] <= cutoff)
    ].copy()
    if off_hist.empty:
        return {}
    merged = off_hist.merge(
        _STATE_SVC_MONTHLY[['Service','month_dt','OOT_Rate']]
        .rename(columns={'OOT_Rate':'State_Avg'}),
        on=['Service','month_dt'], how='left'
    )
    merged['State_Avg'] = merged['State_Avg'].fillna(0)
    merged['is_bad']    = merged['OOT_Rate'] > merged['State_Avg']
    streaks = {}
    for svc, grp in merged.groupby('Service'):
        flags  = grp.sort_values('month_dt')['is_bad'].tolist()
        streak = 0
        for f in reversed(flags):
            if f: streak += 1
            else: break
        streaks[svc] = streak
    return streaks


# ─────────────────────────────────────────────────────────────────────────────
# Streak helpers
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

STREAK_TOOLTIP = (
    "Bad-Month Streak: Consecutive months ending at selected month where this office OOT% "
    "exceeded the district average OOT% for that month.\n"
    "🔔 ≥3 = Watch  |  ⚠️ ≥6 = Sustained  |  🚨 ≥9 = Critical"
)

COMPOSITE_TOOLTIP = (
    "Composite Score (100 = best, 0 = worst)\n"
    "Formula: Score = 100 − norm(0.25·F1 + 0.25·F2 + 0.25·F3 + 0.25·F4)\n"
    "F1: Office OOT% vs District OOT% (current month) — magnitude above/below avg\n"
    "F2: Office OOT% vs State OOT% (current month) — magnitude above/below avg\n"
    "F3: Streak penalty → ≥9=100, ≥6=66, ≥3=33, else 0\n"
    "F4: Service concentration → 25 if no service ≥80% of OOT, 0 if one service ≥80%\n"
    "norm = rescale raw to 0–100 across offices. Higher score = better office."
)


# ─────────────────────────────────────────────────────────────────────────────
# SCORING ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def _score_offices(selected_district, selected_year, selected_month, min_count, max_count):
    cutoff = pd.Timestamp(year=selected_year, month=selected_month, day=1)

    # Snapshot for selected district / month
    snap = _OFF_MONTHLY[
        (_OFF_MONTHLY['District']           == selected_district) &
        (_OFF_MONTHLY['month_dt'].dt.year   == selected_year)     &
        (_OFF_MONTHLY['month_dt'].dt.month  == selected_month)
    ].copy()
    if snap.empty:
        return pd.DataFrame(), 0.0

    district_avg_total = round(snap['Total'].mean(), 2)

    # Apply thresholds
    if min_count > 0:
        snap = snap[snap['Total'] >= min_count]
    if max_count and max_count > 0:
        snap = snap[snap['Total'] <= max_count]
    snap = snap.reset_index(drop=True)
    if snap.empty:
        return pd.DataFrame(), district_avg_total

    valid_offices = snap['Office'].tolist()

    # Current month district & state OOT
    dist_oot  = _dist_snap_oot(selected_district, selected_year, selected_month)
    state_oot = _state_snap_oot(selected_year, selected_month)

    # ── F1: District OOT magnitude ──────────────────────────────────────────
    snap['F1_Raw'] = snap['OOT_Rate'] - dist_oot
    f1_max = snap['F1_Raw'].abs().max() or 1
    snap['F1']     = ((snap['F1_Raw'] / f1_max) * 100).clip(-100, 100)

    # ── F2: State OOT magnitude ──────────────────────────────────────────────
    snap['F2_Raw'] = snap['OOT_Rate'] - state_oot
    f2_max = snap['F2_Raw'].abs().max() or 1
    snap['F2']     = ((snap['F2_Raw'] / f2_max) * 100).clip(-100, 100)

    # ── F3: Consistency streak ───────────────────────────────────────────────
    # district monthly avg per month (for bad-month determination)
    dist_hist_map = _DIST_MONTHLY[
        (_DIST_MONTHLY['District']  == selected_district) &
        (_DIST_MONTHLY['month_dt'] <= cutoff)
    ].set_index('month_dt')['OOT_Rate'].to_dict()

    off_hist = _OFF_MONTHLY[
        (_OFF_MONTHLY['District']   == selected_district) &
        (_OFF_MONTHLY['Office'].isin(valid_offices))      &
        (_OFF_MONTHLY['month_dt']  <= cutoff)
    ].copy()
    off_hist['Dist_Avg_Month'] = off_hist['month_dt'].map(dist_hist_map).fillna(0)
    off_hist['is_bad']         = off_hist['OOT_Rate'] > off_hist['Dist_Avg_Month']

    streaks = {}
    for office, grp in off_hist.groupby('Office'):
        flags  = grp.sort_values('month_dt')['is_bad'].tolist()
        streak = 0
        for f in reversed(flags):
            if f: streak += 1
            else: break
        streaks[office] = streak

    snap['Streak'] = snap['Office'].map(streaks).fillna(0).astype(int)
    snap['F3']     = snap['Streak'].apply(
        lambda s: 100.0 if s >= 9 else 66.0 if s >= 6 else 33.0 if s >= 3 else 0.0)

    # ── F4: Service concentration (25 = good/distributed, 0 = concentrated) ─
    svc_snap = _df[
        (_df['District']          == selected_district) &
        (_df['Office'].isin(valid_offices))             &
        (_df['month_dt'].dt.year  == selected_year)     &
        (_df['month_dt'].dt.month == selected_month)
    ].groupby(['Office','Service']).agg(OOT=('OOT','sum')).reset_index()

    def _top_share(office):
        sub = svc_snap[svc_snap['Office'] == office]
        if sub.empty or sub['OOT'].sum() == 0:
            return 0.0
        return sub['OOT'].max() / sub['OOT'].sum()

    snap['Top_Svc_Share'] = snap['Office'].apply(_top_share)
    # 25 = distributed (no service ≥80%), 0 = concentrated (one service ≥80%)
    # This means F4=25 is GOOD — so for the penalty score we INVERT:
    # penalty = 25 - F4_raw  →  concentrated gets higher penalty
    snap['F4_raw'] = snap['Top_Svc_Share'].apply(
        lambda sh: 0.0 if sh >= 0.80 else 25.0)
    # Convert to 0–100 penalty: concentrated = 100, distributed = 0
    snap['F4'] = snap['Top_Svc_Share'].apply(
        lambda sh: 100.0 if sh >= 0.80 else 0.0)

    # ── Raw composite (higher = worse) ──────────────────────────────────────
    raw = (0.25 * snap['F1'] +
           0.25 * snap['F2'] +
           0.25 * snap['F3'] +
           0.25 * snap['F4'])

    raw_min, raw_max = raw.min(), raw.max()
    if raw_max > raw_min:
        raw_norm = (raw - raw_min) / (raw_max - raw_min) * 100
    else:
        raw_norm = pd.Series([50.0] * len(raw), index=raw.index)

    snap['Composite_Score']  = (100 - raw_norm).round(2)
    snap['District_Avg_OOT'] = dist_oot
    snap['State_Avg_OOT']    = state_oot

    return (
        snap.sort_values('Composite_Score', ascending=True).reset_index(drop=True),
        district_avg_total
    )


# ─────────────────────────────────────────────────────────────────────────────
# Reason builder
# ─────────────────────────────────────────────────────────────────────────────
def _build_reason_items(row):
    items      = []
    delta_dist = row['OOT_Rate'] - row['District_Avg_OOT']
    delta_st   = row['OOT_Rate'] - row['State_Avg_OOT']
    streak     = int(row['Streak'])
    top_sh     = row.get('Top_Svc_Share', 0)

    items.append(
        f"📍 District avg OOT {row['District_Avg_OOT']:.1f}%: "
        f"{'▲ ' + str(abs(round(delta_dist,1))) + '% above' if delta_dist > 0 else '▼ ' + str(abs(round(delta_dist,1))) + '% below'} district avg."
    )
    items.append(
        f"🌐 State avg OOT {row['State_Avg_OOT']:.1f}%: "
        f"{'▲ ' + str(abs(round(delta_st,1))) + '% above' if delta_st > 0 else '▼ ' + str(abs(round(delta_st,1))) + '% below'} state avg."
    )
    items.append(f"📅 Consistency: {_streak_label(streak)} — {_streak_magnitude(streak)}.")
    if top_sh >= 0.80:
        items.append(
            f"⚠️ Service Concentration: Top service drives {top_sh*100:.1f}% of OOT — "
            f"F4 penalty applied (office flagged as single-service failure)."
        )
    else:
        items.append(
            f"✅ Service Distribution OK: Top service = {top_sh*100:.1f}% of OOT "
            f"(below 80% threshold — no concentration penalty)."
        )
    return items


# ─────────────────────────────────────────────────────────────────────────────
# KPI card
# ─────────────────────────────────────────────────────────────────────────────
def _kpi_card(label, value, color="#1a3c5e", bg="#f0f6ff", border="#2d6a9f", tooltip=None):
    return html.Div([
        html.Div(str(value), style={'fontSize':'1.4rem','fontWeight':'700','color':color}),
        html.Div(label, style={'fontSize':'0.75rem','color':'#555','fontWeight':'600',
                               'textTransform':'uppercase','letterSpacing':'0.5px'}),
    ], style={
        'background':bg, 'borderLeft':f'5px solid {border}',
        'padding':'12px 10px','borderRadius':'8px','textAlign':'center',
        'cursor':'help' if tooltip else 'default',
        'title': tooltip or ''
    })


# ─────────────────────────────────────────────────────────────────────────────
# Sunburst data builder  (District → Office → Service  sized by OOT)
# ─────────────────────────────────────────────────────────────────────────────
def _build_sunburst_data(district, year, month):
    COLORS = ['#FFAE57','#FF7853','#EA5151','#CC3F57','#9A2555']

    snap = _df[
        (_df['District']          == district) &
        (_df['month_dt'].dt.year  == year)     &
        (_df['month_dt'].dt.month == month)
    ]
    if snap.empty:
        return None

    # Office → Service aggregation
    agg = (snap.groupby(['Office','Service'])
               .agg(OOT=('OOT','sum'), Total=('Total','sum'))
               .reset_index())
    agg = agg[agg['OOT'] > 0]
    if agg.empty:
        return None

    offices      = agg['Office'].unique()
    office_color = {o: COLORS[i % len(COLORS)] for i, o in enumerate(sorted(offices))}

    children = []
    for office in sorted(offices):
        sub     = agg[agg['Office'] == office].sort_values('OOT', ascending=False)
        off_oot = int(sub['OOT'].sum())
        svc_nodes = []
        for _, row in sub.iterrows():
            svc_nodes.append({
                'name':      row['Service'],
                'value':     int(row['OOT']),
                'itemStyle': {'color': office_color[office], 'opacity': 0.75},
                'label':     {'color': '#fff'},
            })
        children.append({
            'name':      office,
            'value':     off_oot,
            'itemStyle': {'color': office_color[office]},
            'label':     {'color': '#fff'},
            'children':  svc_nodes,
        })

    dist_oot = int(agg['OOT'].sum())
    data = [{
        'name':      district,
        'value':     dist_oot,
        'itemStyle': {'color': '#2d6a9f'},
        'label':     {'color': '#fff', 'rotate': 0},
        'children':  children,
    }]
    return data


# ─────────────────────────────────────────────────────────────────────────────
# Sunburst chart component
# ─────────────────────────────────────────────────────────────────────────────
def _sunburst_component(district, year, month):
    data = _build_sunburst_data(district, year, month)
    if data is None:
        return dbc.Alert("No OOT data available for sunburst chart.", color="warning")

    data_json = json.dumps(data)
    month_name = MONTH_NAMES.get(month, str(month))

    # ECharts sunburst via clientside JS injected in a div
    chart_id = f"sunburst-{district.replace(' ','_')}-{year}-{month}"
    return html.Div([
        html.H5(f"🌐 OOT Sunburst — {district}  |  {month_name} {year}",
                style={'color':'#1a3c5e','marginBottom':'8px'}),
        html.P("Centre = District → Ring 2 = Offices → Outer Ring = Services  |  "
               "Size = Out-of-Time count",
               style={'fontSize':'0.82rem','color':'#888','marginBottom':'6px'}),
        html.Div(
            id=chart_id,
            **{'data-sunburst': data_json},
            style={'width':'100%','height':'520px','background':'#2E2733',
                   'borderRadius':'12px'}
        ),
        # Inline script to render ECharts
        html.Script(f"""
            (function() {{
                function tryRender() {{
                    var el = document.getElementById('{chart_id}');
                    if (!el) {{ setTimeout(tryRender, 200); return; }}
                    if (typeof echarts === 'undefined') {{ setTimeout(tryRender, 300); return; }}
                    var chart = echarts.init(el);
                    var rawData = {data_json};
                    var colors  = ['#FFAE57','#FF7853','#EA5151','#CC3F57','#9A2555'];
                    chart.setOption({{
                        backgroundColor: '#2E2733',
                        tooltip: {{
                            trigger: 'item',
                            formatter: function(p) {{
                                return '<b>' + p.name + '</b><br/>OOT: ' + p.value;
                            }}
                        }},
                        series: [{{
                            type: 'sunburst',
                            center: ['50%','50%'],
                            data: rawData,
                            sort: function(a, b) {{
                                return b.getValue() - a.getValue();
                            }},
                            label: {{
                                rotate: 'radial',
                                color: '#fff',
                                fontSize: 11
                            }},
                            itemStyle: {{
                                borderColor: '#2E2733',
                                borderWidth: 2
                            }},
                            emphasis: {{
                                focus: 'ancestor'
                            }},
                            levels: [
                                {{}},
                                {{
                                    r0: 0, r: 60,
                                    label: {{ rotate: 0, fontSize: 13, fontWeight: 'bold' }}
                                }},
                                {{
                                    r0: 65, r: 180,
                                    label: {{ fontSize: 11 }}
                                }},
                                {{
                                    r0: 185, r: 260,
                                    itemStyle: {{ shadowBlur: 4, shadowColor: '#FFAE57' }},
                                    label: {{
                                        rotate: 'tangential',
                                        fontSize: 9,
                                        color: '#FFAE57'
                                    }}
                                }}
                            ]
                        }}]
                    }});
                    window.addEventListener('resize', function() {{ chart.resize(); }});
                }}
                tryRender();
            }})();
        """),
    ], style={'marginBottom':'20px'})


# ─────────────────────────────────────────────────────────────────────────────
# PDF Generator
# ─────────────────────────────────────────────────────────────────────────────
def _generate_pdf(year, month, min_count, max_count):
    month_name = MONTH_NAMES.get(month, str(month))
    buffer     = io.BytesIO()
    doc        = SimpleDocTemplate(buffer, pagesize=A4,
                                   leftMargin=1.5*cm, rightMargin=1.5*cm,
                                   topMargin=1.5*cm,  bottomMargin=1.5*cm)
    styles     = getSampleStyleSheet()
    story      = []

    title_style = ParagraphStyle('title', parent=styles['Title'],
                                 fontSize=15, textColor=colors.HexColor('#1a3c5e'), spaceAfter=4)
    h2_style    = ParagraphStyle('h2', parent=styles['Heading2'],
                                 fontSize=12, textColor=colors.HexColor('#2d6a9f'),
                                 spaceBefore=12, spaceAfter=3)
    h3_style    = ParagraphStyle('h3', parent=styles['Heading3'],
                                 fontSize=10, textColor=colors.HexColor('#c0392b'),
                                 spaceBefore=8, spaceAfter=2)
    body_style  = ParagraphStyle('body', parent=styles['Normal'], fontSize=8, spaceAfter=2)
    label_style = ParagraphStyle('label', parent=styles['Normal'],
                                 fontSize=7.5, textColor=colors.HexColor('#555555'))

    def _tbl(data, col_widths=None):
        t = Table(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND',     (0,0),(-1,0), colors.HexColor('#1a3c5e')),
            ('TEXTCOLOR',      (0,0),(-1,0), colors.white),
            ('FONTSIZE',       (0,0),(-1,-1), 7.5),
            ('ROWBACKGROUNDS', (0,1),(-1,-1),
             [colors.HexColor('#f0f6ff'), colors.white]),
            ('GRID',           (0,0),(-1,-1), 0.4, colors.HexColor('#d0d7de')),
            ('ALIGN',          (1,0),(-1,-1), 'CENTER'),
            ('VALIGN',         (0,0),(-1,-1), 'MIDDLE'),
            ('TOPPADDING',     (0,0),(-1,-1), 2),
            ('BOTTOMPADDING',  (0,0),(-1,-1), 2),
        ]))
        return t

    story.append(Paragraph(f"Monthly Performance Report — {month_name} {year}", title_style))
    story.append(Paragraph(
        f"All districts  |  Min: {min_count}  |  Max: {max_count or 'None'}  |  "
        f"Score: 100=best, 0=worst  |  District & State OOT = current month snapshot",
        body_style))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor('#2d6a9f'), spaceAfter=8))
    story.append(Paragraph(
        "Formula: Score = 100 − norm(0.25·F1 + 0.25·F2 + 0.25·F3 + 0.25·F4)  |  "
        "F1=District OOT magnitude  |  F2=State OOT magnitude  |  "
        "F3=Streak penalty  |  F4=100 if top service ≥80% OOT else 0",
        label_style))
    story.append(Spacer(1, 8))

    state_svc_avg = _state_service_avg(year, month)
    districts     = sorted(_df['District'].dropna().unique())

    for district in districts:
        scored, dist_avg_total = _score_offices(district, year, month, min_count, max_count)
        if scored.empty:
            continue

        dist_oot  = scored['District_Avg_OOT'].iloc[0]
        state_oot = scored['State_Avg_OOT'].iloc[0]

        story.append(Paragraph(f"District: {district}", h2_style))
        story.append(HRFlowable(width="100%", thickness=0.5,
                                 color=colors.HexColor('#aaaaaa'), spaceAfter=3))
        story.append(Paragraph(
            f"District OOT: <b>{dist_oot:.1f}%</b>  |  "
            f"State OOT: <b>{state_oot:.1f}%</b>  |  "
            f"Avg Total: <b>{dist_avg_total:,.0f}</b>  |  "
            f"Offices: <b>{len(scored)}</b>", body_style))
        story.append(Spacer(1, 4))

        worst3 = scored.head(3)
        for rank, (_, row) in enumerate(worst3.iterrows(), 1):
            streak = int(row['Streak'])
            delta_d = row['OOT_Rate'] - row['District_Avg_OOT']
            delta_s = row['OOT_Rate'] - row['State_Avg_OOT']
            top_sh  = row.get('Top_Svc_Share', 0)

            story.append(Paragraph(f"Rank {rank} Worst — {row['Office']}", h3_style))
            off_data = [
                ['Metric',                  'Value'],
                ['Total Applications',       f"{int(row['Total']):,}"],
                ['Out-of-Time (OOT)',        f"{int(row['OOT']):,}"],
                ['Office OOT Rate',          f"{row['OOT_Rate']:.1f}%"],
                ['District OOT (month)',      f"{row['District_Avg_OOT']:.1f}%"],
                ['State OOT (month)',         f"{row['State_Avg_OOT']:.1f}%"],
                ['vs District',              f"{'▲' if delta_d>0 else '▼'} {abs(delta_d):.1f}%"],
                ['vs State',                 f"{'▲' if delta_s>0 else '▼'} {abs(delta_s):.1f}%"],
                ['Consecutive Bad Months',   f"{streak}"],
                ['Streak Magnitude',         _streak_magnitude(streak)],
                ['Top Service OOT Share',    f"{top_sh*100:.1f}%"],
                ['F4 Concentration',         'Penalised' if top_sh >= 0.80 else 'OK (distributed)'],
                ['Composite Score (100=best)', f"{row['Composite_Score']:.1f}"],
            ]
            story.append(_tbl(off_data, col_widths=[8*cm, 7*cm]))
            story.append(Spacer(1, 4))

            snap_svc = _df[
                (_df['District']          == district) &
                (_df['Office']            == row['Office']) &
                (_df['month_dt'].dt.year  == year) &
                (_df['month_dt'].dt.month == month)
            ].groupby('Service').agg(Total=('Total','sum'), OOT=('OOT','sum')).reset_index()
            snap_svc['OOT_Rate']      = snap_svc.apply(lambda r: _oot_rate(r['OOT'], r['Total']), axis=1)
            snap_svc['State_Svc_Avg'] = snap_svc['Service'].map(lambda s: state_svc_avg.get(s, 0.0))
            snap_svc['vs_State']      = snap_svc['OOT_Rate'] - snap_svc['State_Svc_Avg']
            total_oot_off             = snap_svc['OOT'].sum()
            snap_svc['OOT_Share']     = snap_svc['OOT'].apply(
                lambda o: round(o/total_oot_off*100,1) if total_oot_off > 0 else 0)
            svc_streaks = _service_consistency(row['Office'], year, month)

            svc_rows = [['Service','Total','OOT','OOT%','State Avg%','vs State','OOT Share','Streak','Magnitude']]
            for _, sr in snap_svc.sort_values('OOT_Rate', ascending=False).iterrows():
                st = svc_streaks.get(sr['Service'], 0)
                svc_rows.append([
                    sr['Service'],
                    f"{int(sr['Total']):,}",
                    f"{int(sr['OOT']):,}",
                    f"{sr['OOT_Rate']:.1f}%",
                    f"{sr['State_Svc_Avg']:.1f}%",
                    f"{'▲' if sr['vs_State']>0 else '▼'} {abs(sr['vs_State']):.1f}%",
                    f"{sr['OOT_Share']:.1f}%",
                    _streak_label(st),
                    _streak_magnitude(st),
                ])
            story.append(Paragraph("Service Breakdown:", label_style))
            story.append(_tbl(svc_rows,
                               col_widths=[3.5*cm,1.3*cm,1.1*cm,1.2*cm,
                                           1.7*cm,1.4*cm,1.4*cm,2.8*cm,3*cm]))
            story.append(Spacer(1, 8))
        story.append(Spacer(1, 10))

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


# ─────────────────────────────────────────────────────────────────────────────
# Layout
# ─────────────────────────────────────────────────────────────────────────────
districts = sorted(_df['District'].dropna().unique()) if not _df.empty else []
years     = sorted(_df['month_dt'].dt.year.unique())  if not _df.empty else []

layout = html.Div([
    # ECharts CDN
    html.Script(src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"),

    html.Div([
        html.H2("🔍 Advanced Analytics & Performance Report",
                style={'color':'white','margin':'0','fontSize':'1.6rem','letterSpacing':'1px'}),
        html.P("Composite score: 100 = best, 0 = worst  |  "
               "District & State OOT = selected month snapshot",
               style={'color':'#c8dff0','margin':'4px 0 0 0','fontSize':'0.9rem'}),
    ], style={
        'background':'linear-gradient(90deg,#1a3c5e 0%,#2d6a9f 100%)',
        'padding':'18px 28px','borderRadius':'10px','marginBottom':'20px'
    }),

    dbc.Row([
        dbc.Col([html.Label("🏛️ District"),
                 dcc.Dropdown(id='aa-district',
                              options=[{'label':d,'value':d} for d in districts],
                              value=districts[0] if districts else None,
                              clearable=False)], md=3),
        dbc.Col([html.Label("📅 Year"),
                 dcc.Dropdown(id='aa-year',
                              options=[{'label':y,'value':y} for y in years],
                              value=years[-1] if years else None,
                              clearable=False)], md=1),
        dbc.Col([html.Label("🗓️ Month"),
                 dcc.Dropdown(id='aa-month', clearable=False)], md=2),
        dbc.Col([html.Label("⚙️ Min Total"),
                 dcc.Input(id='aa-min-count', type='number', value=100, min=0,
                           style={'width':'100%','padding':'6px',
                                  'borderRadius':'4px','border':'1px solid #ccc'})], md=1),
        dbc.Col([html.Label("⚙️ Max Total"),
                 dcc.Input(id='aa-max-count', type='number', value=None, min=0,
                           placeholder='No limit',
                           style={'width':'100%','padding':'6px',
                                  'borderRadius':'4px','border':'1px solid #ccc'})], md=1),
        dbc.Col([
            html.Br(),
            dbc.Row([
                dbc.Col(dbc.Button("🔍 Analyse", id='aa-run-btn',
                                   color='primary', style={'width':'100%'}), width=6),
                dbc.Col(dbc.Button("📄 Monthly Report PDF", id='aa-pdf-btn',
                                   color='warning',
                                   style={'width':'100%','fontSize':'0.75rem'}), width=6),
            ])
        ], md=4),
    ], className='mb-4'),

    html.Div(id='aa-pdf-download'),
    dcc.Download(id='aa-pdf-file'),
    html.Div(id='aa-output'),

], style={'padding':'20px'})


# ─────────────────────────────────────────────────────────────────────────────
# Callbacks
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output('aa-month','options'),
    Output('aa-month','value'),
    Input('aa-district','value'),
    Input('aa-year',    'value'),
)
def update_months(district, year):
    if not district or not year or _df.empty:
        return [], None
    filtered = _df[(_df['District']==district) & (_df['month_dt'].dt.year==year)]
    months   = sorted(filtered['month_dt'].dt.month.unique())
    opts     = [{'label':MONTH_NAMES[m],'value':m} for m in months]
    return opts, (months[-1] if months else None)


@app.callback(
    Output('aa-pdf-file',     'data'),
    Output('aa-pdf-download', 'children'),
    Input('aa-pdf-btn',   'n_clicks'),
    State('aa-year',      'value'),
    State('aa-month',     'value'),
    State('aa-min-count', 'value'),
    State('aa-max-count', 'value'),
    prevent_initial_call=True,
)
def generate_pdf(n_clicks, year, month, min_count, max_count):
    if not year or not month:
        return None, dbc.Alert("Please select Year and Month first.", color="warning")
    try:
        pdf_bytes  = _generate_pdf(int(year), int(month),
                                   int(min_count or 0),
                                   int(max_count) if max_count else None)
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
    Output('aa-output','children'),
    Input('aa-run-btn',   'n_clicks'),
    State('aa-district',  'value'),
    State('aa-year',      'value'),
    State('aa-month',     'value'),
    State('aa-min-count', 'value'),
    State('aa-max-count', 'value'),
    prevent_initial_call=True,
)
def run_analysis(n_clicks, district, year, month, min_count, max_count):
    if not all([district, year, month]):
        return dbc.Alert("Please select all filters.", color="warning")

    min_count = int(min_count or 0)
    max_count = int(max_count) if max_count else None
    year, month = int(year), int(month)

    scored, district_avg_total = _score_offices(
        district, year, month, min_count, max_count)

    if scored.empty:
        return dbc.Alert(
            f"No offices found within Total [{min_count} – {max_count or '∞'}].",
            color="warning")

    dist_oot   = scored['District_Avg_OOT'].iloc[0]
    state_oot  = scored['State_Avg_OOT'].iloc[0]
    month_name = MONTH_NAMES.get(month, str(month))

    # ── KPI row ───────────────────────────────────────────────────────────────
    kpi_row = dbc.Row([
        dbc.Col(_kpi_card("Offices Analysed",   str(len(scored))), md=2),
        dbc.Col(_kpi_card("Avg Total (pre-filter)",
                          f"{district_avg_total:,.0f}",
                          color="#1a3c5e",bg="#eaf4ff",border="#2d6a9f"), md=2),
        dbc.Col(_kpi_card("District OOT",
                          f"{dist_oot:.1f}%",
                          color="#c0392b",bg="#fff5f5",border="#e74c3c"), md=2),
        dbc.Col(_kpi_card("State OOT",
                          f"{state_oot:.1f}%",
                          color="#7d3c98",bg="#fdf5ff",border="#9b59b6"), md=2),
        dbc.Col(_kpi_card("Flagged ≥3 Bad Months",
                          str(int((scored['Streak']>=3).sum())),
                          color="#e67e22",bg="#fff8f0",border="#e67e22"), md=2),
        dbc.Col(_kpi_card("Single-Svc Concentrated",
                          str(int((scored['Top_Svc_Share']>=0.80).sum())),
                          color="#c0392b",bg="#fff5f5",border="#e74c3c"), md=2),
    ], className='mb-3')

    threshold_banner = dbc.Alert([
        html.Strong(f"📊 {district}  |  {month_name} {year}"),
        html.Br(),
        f"Pre-filter avg total: ", html.Strong(f"{district_avg_total:,.1f}"),
        f"  |  Filter: Total ∈ [{min_count}, {max_count or '∞'}]",
    ], color="info", className="mb-3")

    # Formula banner
    formula_banner = dbc.Alert([
        html.Strong("📐 Composite Score Formula  "),
        html.Code("Score = 100 − norm( 0.25·F1 + 0.25·F2 + 0.25·F3 + 0.25·F4 )",
                  style={'fontSize':'0.85rem'}),
        html.Br(),
        html.Small([
            html.B("F1"), " Office OOT% vs District OOT% (current month magnitude)  |  ",
            html.B("F2"), " Office OOT% vs State OOT% (current month magnitude)  |  ",
            html.B("F3"), " Streak penalty (≥9→100, ≥6→66, ≥3→33, else 0)  |  ",
            html.B("F4"), " Concentration: 100 if one service ≥80% of OOT, else 0  |  ",
            html.B("norm"), " = rescale to 0–100 across offices  |  ",
            html.B("100 = best, 0 = worst"),
        ], style={'color':'#555','fontSize':'0.82rem'})
    ], color="light", className="mb-3", style={'border':'1px solid #d0d7de'})

    # ── Office table ──────────────────────────────────────────────────────────
    disp = scored[['Office','Total','OOT','OOT_Rate',
                   'District_Avg_OOT','State_Avg_OOT',
                   'Streak','Top_Svc_Share','Composite_Score']].copy()
    disp['Top_Svc_Share'] = (disp['Top_Svc_Share']*100).round(1)
    disp = disp.rename(columns={
        'OOT':              'Out of Time',
        'OOT_Rate':         'OOT Rate (%)',
        'District_Avg_OOT': 'District OOT (%)',
        'State_Avg_OOT':    'State OOT (%)',
        'Streak':           'Bad-Month Streak ⓘ',
        'Top_Svc_Share':    'Top Svc OOT Share (%)',
        'Composite_Score':  'Score(100=best) ⓘ',
    })
    table_section = html.Div([
        html.H4(f"📊 Office Performance — {month_name} {year}  |  {district}",
                className='mb-1'),
        html.Small("Sorted worst→best. Hover ⓘ headers for explanation.",
                   style={'color':'#888'}),
        dbc.Table.from_dataframe(disp.round(2), striped=True, bordered=True,
                                 hover=True, responsive=True, size='sm')
    ], className='mb-4')

    # ── Bar chart ─────────────────────────────────────────────────────────────
    fig = px.bar(
        scored.sort_values('Composite_Score'),
        x='Office', y='OOT_Rate', color='Composite_Score',
        color_continuous_scale='RdYlGn', range_color=[0,100],
        title=f"OOT Rate by Office — {month_name} {year}  |  {district}",
        labels={'OOT_Rate':'% Out of Time','Composite_Score':'Score (100=best)'},
        text='OOT_Rate',
    )
    fig.add_hline(y=dist_oot, line_dash='dash', line_color='#2980b9',
                  annotation_text=f"District OOT: {dist_oot:.1f}%",
                  annotation_position="top left")
    fig.add_hline(y=state_oot, line_dash='dot', line_color='#8e44ad',
                  annotation_text=f"State OOT: {state_oot:.1f}%",
                  annotation_position="bottom right")
    fig.update_layout(xaxis_tickangle=-45, height=460)
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    chart_section = dcc.Graph(figure=fig, className='mb-4')

    # ── Sunburst ──────────────────────────────────────────────────────────────
    sunburst_section = html.Div([
        html.Hr(),
        _sunburst_component(district, year, month),
        html.Hr(),
    ], className='mb-4')

    # ── Worst 3 ───────────────────────────────────────────────────────────────
    worst_labels = ["🥇 Rank 1 — Worst","🥈 Rank 2","🥉 Rank 3"]
    worst_items  = []
    for rank, (_, row) in enumerate(scored.head(min(3,len(scored))).iterrows(), 1):
        reasons = _build_reason_items(row)
        office  = row['Office']
        streak  = int(row['Streak'])
        top_sh  = row.get('Top_Svc_Share', 0)

        worst_items.append(dbc.Card([
            dbc.CardHeader(html.Strong(
                f"{worst_labels[rank-1]}  |  {office}  |  "
                f"OOT: {row['OOT_Rate']:.1f}%  |  Score: {row['Composite_Score']:.1f}/100"
            ), style={'background':'#fff0f0','color':'#c0392b'}),
            dbc.CardBody([
                dbc.Row([
                    dbc.Col(_kpi_card("Total",         f"{int(row['Total']):,}",
                                      color="#1a3c5e",bg="#eaf4ff",border="#2d6a9f"), md=2),
                    dbc.Col(_kpi_card("OOT Rate",      f"{row['OOT_Rate']:.1f}%",
                                      color="#c0392b",bg="#fff5f5",border="#e74c3c"), md=2),
                    dbc.Col(_kpi_card("District OOT",  f"{row['District_Avg_OOT']:.1f}%",
                                      color="#2471a3",bg="#eaf4ff",border="#3498db"), md=2),
                    dbc.Col(_kpi_card("State OOT",     f"{row['State_Avg_OOT']:.1f}%",
                                      color="#7d3c98",bg="#fdf5ff",border="#9b59b6"), md=2),
                    dbc.Col(
                        html.Div([
                            html.Div(_streak_label(streak),
                                     style={'fontSize':'1rem','fontWeight':'700','color':'#e67e22'}),
                            html.Div("Bad-Month Streak",
                                     style={'fontSize':'0.72rem','color':'#555',
                                            'fontWeight':'600','textTransform':'uppercase'}),
                            html.Div(_streak_magnitude(streak),
                                     style={'fontSize':'0.72rem','color':'#888'}),
                        ], style={'background':'#fff8f0','borderLeft':'5px solid #e67e22',
                                  'padding':'10px','borderRadius':'8px','textAlign':'center',
                                  'title':STREAK_TOOLTIP,'cursor':'help'}),
                    md=2),
                    dbc.Col(_kpi_card(
                        "Top Svc OOT Share",
                        f"{top_sh*100:.1f}%",
                        color="#e74c3c" if top_sh>=0.80 else "#27ae60",
                        bg="#fff5f5"    if top_sh>=0.80 else "#f0fff4",
                        border="#e74c3c" if top_sh>=0.80 else "#27ae60",
                        tooltip="% of this office's total OOT coming from top service. "
                                "≥80% triggers F4 concentration penalty."
                    ), md=2),
                ], className='mb-2'),
                html.Strong("📝 Reason for Flagging:"),
                html.Ul([html.Li(r) for r in reasons]),
                dbc.Button("🔎 Service Breakdown",
                           id={'type':'aa-detail-btn','index':office},
                           color='outline-danger', size='sm', className='mt-2'),
                dbc.Collapse(
                    html.Div(id={'type':'aa-detail-panel','index':office}),
                    id={'type':'aa-detail-collapse','index':office},
                    is_open=False
                )
            ])
        ], className='mb-3'))

    worst_section = html.Div([
        html.Div([
            html.H3("⚠️ Worst Performing Offices",
                    style={'margin':'0','color':'#c0392b'}),
            html.P("Sorted lowest composite score first. "
                   "Score = 100−norm(0.25·F1+0.25·F2+0.25·F3+0.25·F4)",
                   style={'margin':'4px 0 0 0','color':'#7f8c8d','fontSize':'0.88rem'}),
        ], style={'background':'#fff0f0','border':'1.5px solid #e74c3c',
                  'borderRadius':'10px','padding':'14px 22px','marginBottom':'12px'}),
        *worst_items
    ], className='mb-4')

    # ── Best 3 ────────────────────────────────────────────────────────────────
    best_labels = ["🥇 Best","🥈 2nd Best","🥉 3rd Best"]
    best_items  = []
    best_rows   = scored.tail(min(3,len(scored))).iloc[::-1].reset_index(drop=True)
    for rank,(_, row) in enumerate(best_rows.iterrows(), 1):
        reasons = _build_reason_items(row)
        streak  = int(row['Streak'])
        top_sh  = row.get('Top_Svc_Share',0)
        best_items.append(dbc.Card([
            dbc.CardHeader(html.Strong(
                f"{best_labels[rank-1]}  |  {row['Office']}  |  "
                f"OOT: {row['OOT_Rate']:.1f}%  |  Score: {row['Composite_Score']:.1f}/100"
            ), style={'background':'#f0fff4','color':'#1e8449'}),
            dbc.CardBody([
                dbc.Row([
                    dbc.Col(_kpi_card("Total",        f"{int(row['Total']):,}",
                                      color="#1a3c5e",bg="#eaf4ff",border="#2d6a9f"), md=2),
                    dbc.Col(_kpi_card("OOT Rate",     f"{row['OOT_Rate']:.1f}%",
                                      color="#1e8449",bg="#f0fff4",border="#27ae60"), md=2),
                    dbc.Col(_kpi_card("District OOT", f"{row['District_Avg_OOT']:.1f}%",
                                      color="#2471a3",bg="#eaf4ff",border="#3498db"), md=2),
                    dbc.Col(_kpi_card("State OOT",    f"{row['State_Avg_OOT']:.1f}%",
                                      color="#7d3c98",bg="#fdf5ff",border="#9b59b6"), md=2),
                    dbc.Col(
                        html.Div([
                            html.Div(_streak_label(streak),
                                     style={'fontSize':'1rem','fontWeight':'700','color':'#27ae60'}),
                            html.Div("Bad-Month Streak",
                                     style={'fontSize':'0.72rem','color':'#555',
                                            'fontWeight':'600','textTransform':'uppercase'}),
                            html.Div(_streak_magnitude(streak),
                                     style={'fontSize':'0.72rem','color':'#888'}),
                        ], style={'background':'#f0fff4','borderLeft':'5px solid #27ae60',
                                  'padding':'10px','borderRadius':'8px','textAlign':'center',
                                  'title':STREAK_TOOLTIP,'cursor':'help'}),
                    md=2),
                    dbc.Col(_kpi_card(
                        "Top Svc OOT Share", f"{top_sh*100:.1f}%",
                        color="#e74c3c" if top_sh>=0.80 else "#27ae60",
                        bg="#fff5f5"    if top_sh>=0.80 else "#f0fff4",
                        border="#e74c3c" if top_sh>=0.80 else "#27ae60"), md=2),
                ], className='mb-2'),
                html.Strong("📝 Performance Summary:"),
                html.Ul([html.Li(r) for r in reasons])
            ])
        ], className='mb-3'))

    best_section = html.Div([
        html.Div([
            html.H3("🏆 Best Performing Offices",
                    style={'margin':'0','color':'#1e8449'}),
            html.P("Highest composite score = lowest OOT burden relative to benchmarks",
                   style={'margin':'4px 0 0 0','color':'#7f8c8d','fontSize':'0.88rem'}),
        ], style={'background':'#f0fff4','border':'1.5px solid #27ae60',
                  'borderRadius':'10px','padding':'14px 22px','marginBottom':'12px'}),
        *best_items
    ], className='mb-4')

    disclaimer = html.Div([
        html.Strong("ℹ️ Methodology: "),
        f"Filter Total ∈ [{min_count},{max_count or '∞'}]. "
        f"F1/F2 = magnitude of office OOT vs district/state snapshot (current month). "
        f"F4 = 100 penalty if top service ≥80% of office OOT, else 0 (distributed = no penalty). "
        f"Raw normalised 0–100 across offices then inverted. 100=best."
    ], style={'background':'#f8f9fa','border':'1px solid #d0d7de',
              'borderRadius':'8px','padding':'12px 18px','color':'#555','fontSize':'0.82rem'})

    return html.Div([
        kpi_row, threshold_banner, formula_banner,
        table_section, chart_section, sunburst_section,
        worst_section, best_section, disclaimer
    ])


# ─────────────────────────────────────────────────────────────────────────────
# Service detail panel callback
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output({'type':'aa-detail-collapse','index':ALL},'is_open'),
    Output({'type':'aa-detail-panel',   'index':ALL},'children'),
    Input({'type':'aa-detail-btn',      'index':ALL},'n_clicks'),
    State({'type':'aa-detail-collapse', 'index':ALL},'is_open'),
    State('aa-district','value'),
    State('aa-year',    'value'),
    State('aa-month',   'value'),
    prevent_initial_call=True,
)
def toggle_detail(n_clicks_list, is_open_list, district, year, month):
    if not callback_context.triggered_id:
        return is_open_list, [dash.no_update]*len(is_open_list)

    office     = callback_context.triggered_id['index']
    year, month = int(year), int(month)
    all_ids    = [t['id']['index'] for t in callback_context.inputs_list[0]]
    target_idx = all_ids.index(office)

    new_open             = list(is_open_list)
    new_open[target_idx] = not is_open_list[target_idx]
    panels               = [dash.no_update]*len(is_open_list)

    if not new_open[target_idx]:
        return new_open, panels

    state_svc_avg = _state_service_avg(year, month)

    snap = _df[
        (_df['District']          == district) &
        (_df['Office']            == office)   &
        (_df['month_dt'].dt.year  == year)     &
        (_df['month_dt'].dt.month == month)
    ]
    svc_snap = (
        snap.groupby('Service')
        .agg(Total=('Total','sum'), OOT=('OOT','sum'))
        .reset_index()
    )
    svc_snap['OOT_Rate']      = svc_snap.apply(lambda r: _oot_rate(r['OOT'],r['Total']), axis=1)
    svc_snap['State_Svc_Avg'] = svc_snap['Service'].map(lambda s: state_svc_avg.get(s,0.0))
    svc_snap['vs_State']      = svc_snap['OOT_Rate'] - svc_snap['State_Svc_Avg']
    total_oot                 = svc_snap['OOT'].sum()
    svc_snap['OOT_Share']     = svc_snap['OOT'].apply(
        lambda o: round(o/total_oot*100,1) if total_oot>0 else 0.0)
    svc_snap                  = svc_snap.sort_values('OOT_Rate', ascending=False)
    svc_streaks               = _service_consistency(office, year, month)

    rows = []
    for _, sr in svc_snap.iterrows():
        st      = svc_streaks.get(sr['Service'], 0)
        is_bad  = sr['vs_State'] > 0
        domina  = sr['OOT_Share'] >= 80
        bg      = '#fff5f5' if (is_bad or domina) else '#f0fff4'
        rows.append(html.Tr([
            html.Td(sr['Service'],              style={'fontWeight':'600'}),
            html.Td(f"{int(sr['Total']):,}",    style={'textAlign':'center'}),
            html.Td(f"{int(sr['OOT']):,}",      style={'textAlign':'center'}),
            html.Td(f"{sr['OOT_Rate']:.1f}%",   style={'textAlign':'center',
                                                         'color':'#c0392b' if is_bad else '#1e8449',
                                                         'fontWeight':'700'}),
            html.Td(f"{sr['State_Svc_Avg']:.1f}%", style={'textAlign':'center'}),
            html.Td(f"{'▲' if is_bad else '▼'} {abs(sr['vs_State']):.1f}%",
                    style={'textAlign':'center',
                           'color':'#c0392b' if is_bad else '#1e8449','fontWeight':'600'}),
            html.Td(f"{sr['OOT_Share']:.1f}%",
                    style={'textAlign':'center',
                           'color':'#e74c3c' if domina else '#555',
                           'fontWeight':'700' if domina else '400'}),
            html.Td(_streak_label(st),          style={'textAlign':'center'}),
            html.Td(_streak_magnitude(st),      style={'textAlign':'center',
                                                         'fontSize':'0.78rem','color':'#555'}),
        ], style={'background':bg}))

    detail_table = html.Div([
        html.H5(f"📋 Service Breakdown — {office}",
                style={'color':'#1a3c5e','marginTop':'14px','marginBottom':'6px'}),
        html.P(
            "OOT% vs state-wide average for each service.  "
            "OOT Share = % of office total OOT from that service "
            "(red if ≥80% — triggers F4 concentration penalty).  "
            "Streak = consecutive months office OOT for this service > state service avg.",
            style={'fontSize':'0.82rem','color':'#666','marginBottom':'6px'}
        ),
        dbc.Table([
            html.Thead(html.Tr([
                html.Th("Service"),
                html.Th("Total",         style={'textAlign':'center'}),
                html.Th("OOT",           style={'textAlign':'center'}),
                html.Th("OOT Rate",      style={'textAlign':'center'}),
                html.Th("State Svc Avg", style={'textAlign':'center'}),
                html.Th("vs State",      style={'textAlign':'center'}),
                html.Th("OOT Share",     style={'textAlign':'center'}),
                html.Th("Streak",        style={'textAlign':'center'}),
                html.Th("Magnitude",     style={'textAlign':'center'}),
            ]), style={'background':'#1a3c5e','color':'white'}),
            html.Tbody(rows)
        ], bordered=True, hover=True, responsive=True, size='sm')
    ], style={'padding':'10px','background':'#fafafa',
              'border':'1px solid #ddd','borderRadius':'8px'})

    panels[target_idx] = detail_table
    return new_open, panels