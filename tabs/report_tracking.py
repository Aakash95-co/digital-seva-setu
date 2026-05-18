import base64
import dash
import pandas as pd
from dash import html, dcc, Input, Output, State, ALL, callback_context
import dash_bootstrap_components as dbc
from app import app
from data import df_adv

# ═════════════════════════════════════════════════════════════════════════════
# DATA PREP  (derive dropdown options from the loaded dataset)
# ═════════════════════════════════════════════════════════════════════════════
def _prep_options(raw):
    if raw is None or raw.empty:
        return [], [], []
    df = raw.copy()
    df.columns = df.columns.str.strip()
    _map = {
        'District_name': 'District', 'District_Eng': 'District',
        'Office_name': 'Office',     'Office_Eng': 'Office',
        'Service_name': 'Service',   'Service_Eng': 'Service',
    }
    df.rename(columns={k: v for k, v in _map.items() if k in df.columns}, inplace=True)
    for c in ('District', 'Office', 'Service'):
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    districts = sorted(df['District'].dropna().unique()) if 'District' in df.columns else []
    offices   = sorted(df['Office'].dropna().unique())   if 'Office'   in df.columns else []
    services  = sorted(df['Service'].dropna().unique())  if 'Service'  in df.columns else []
    return districts, offices, services


_districts, _offices, _services = _prep_options(df_adv)

_district_opts = [{'label': d, 'value': d} for d in _districts]
_office_opts   = [{'label': o, 'value': o} for o in _offices]
_service_opts  = [{'label': s, 'value': s} for s in _services]


# ═════════════════════════════════════════════════════════════════════════════
# LAYOUT
# ═════════════════════════════════════════════════════════════════════════════
layout = html.Div([
    # ── Header ──────────────────────────────────────────────────────────────
    html.Div([
        html.H2("📁 Report & Tracking",
                style={'color': 'white', 'margin': '0', 'fontSize': '1.6rem'}),
        html.P("Upload reports for record-keeping  |  Track districts, offices or services over time",
               style={'color': '#c8dff0', 'margin': '4px 0 0 0', 'fontSize': '0.9rem'}),
    ], style={'background': 'linear-gradient(90deg,#1a3c5e,#2d6a9f)',
              'padding': '18px 28px', 'borderRadius': '10px', 'marginBottom': '24px'}),

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 1 — REPORTING (file upload)
    # ══════════════════════════════════════════════════════════════════════
    html.Div([
        html.H4("📤 Reporting — Upload Files",
                style={'color': '#1a3c5e', 'borderBottom': '2px solid #2d6a9f',
                       'paddingBottom': '8px', 'marginBottom': '16px'}),
        html.P("Upload any report document (PDF, Excel, CSV, Word, images, etc.) "
               "for record-keeping and future reference.",
               style={'color': '#666', 'fontSize': '0.92rem', 'marginBottom': '14px'}),

        dcc.Upload(
            id='rt-upload',
            children=html.Div([
                html.Div("📂", style={'fontSize': '2.5rem', 'marginBottom': '8px'}),
                html.Strong("Drag & Drop files here"),
                html.Span(" or ", style={'color': '#888'}),
                html.A("Browse", style={'color': '#2d6a9f', 'textDecoration': 'underline',
                                        'cursor': 'pointer'}),
                html.Br(),
                html.Small("Supported: PDF, XLSX, XLS, CSV, DOCX, JPG, PNG, …",
                           style={'color': '#888'}),
            ], style={'textAlign': 'center', 'padding': '10px'}),
            style={
                'width': '100%', 'minHeight': '130px',
                'lineHeight': '1.6', 'borderWidth': '2px',
                'borderStyle': 'dashed', 'borderColor': '#2d6a9f',
                'borderRadius': '10px', 'backgroundColor': '#f8fbff',
                'display': 'flex', 'alignItems': 'center',
                'justifyContent': 'center', 'cursor': 'pointer',
            },
            multiple=True,
        ),
        html.Div(id='rt-upload-output', style={'marginTop': '14px'}),
    ], style={
        'background': 'white', 'border': '1px solid #d0d7de',
        'borderRadius': '10px', 'padding': '22px 26px', 'marginBottom': '28px',
        'boxShadow': '0 2px 8px rgba(0,0,0,0.06)',
    }),

    # ══════════════════════════════════════════════════════════════════════
    # SECTION 2 — TRACKING
    # ══════════════════════════════════════════════════════════════════════
    html.Div([
        html.H4("📌 Tracking — Monitor Performance",
                style={'color': '#1a3c5e', 'borderBottom': '2px solid #2d6a9f',
                       'paddingBottom': '8px', 'marginBottom': '16px'}),
        html.P("Select a category and item to track monthly. "
               "Use Track / Untrack to manage your watchlist.",
               style={'color': '#666', 'fontSize': '0.92rem', 'marginBottom': '18px'}),

        dbc.Row([
            # Category selector
            dbc.Col([
                html.Label("📊 Category", style={'fontWeight': 'bold', 'marginBottom': '4px'}),
                dcc.Dropdown(
                    id='rt-track-category',
                    options=[
                        {'label': '🏛️ District', 'value': 'district'},
                        {'label': '🏢 Office',   'value': 'office'},
                        {'label': '⚙️ Service',  'value': 'service'},
                    ],
                    value='district',
                    clearable=False,
                ),
            ], md=3),

            # Dynamic value dropdown (populated by callback)
            dbc.Col([
                html.Label("🔎 Select Item", style={'fontWeight': 'bold', 'marginBottom': '4px'}),
                dcc.Dropdown(
                    id='rt-track-value',
                    options=_district_opts,
                    value=None,
                    placeholder='Choose…',
                    clearable=True,
                ),
            ], md=5),

            # Month dropdown
            dbc.Col([
                html.Label("📅 Month (optional)", style={'fontWeight': 'bold', 'marginBottom': '4px'}),
                dcc.Dropdown(
                    id='rt-track-month',
                    options=[
                        {'label': m, 'value': m}
                        for m in ['January', 'February', 'March', 'April',
                                  'May', 'June', 'July', 'August',
                                  'September', 'October', 'November', 'December']
                    ],
                    value=None,
                    placeholder='All months',
                    clearable=True,
                ),
            ], md=2),

            # Buttons
            dbc.Col([
                html.Label("\u00a0", style={'display': 'block', 'marginBottom': '4px'}),
                dbc.Row([
                    dbc.Col(
                        dbc.Button("➕ Track", id='rt-track-btn', color='success',
                                   style={'width': '100%', 'fontWeight': 'bold'}),
                        width=6),
                    dbc.Col(
                        dbc.Button("➖ Untrack", id='rt-untrack-btn', color='danger',
                                   outline=True, style={'width': '100%', 'fontWeight': 'bold'}),
                        width=6),
                ], className='g-1'),
            ], md=2),
        ], className='mb-3 align-items-end'),

        # Feedback / status message
        html.Div(id='rt-track-status', style={'marginBottom': '12px'}),

        # Tracked items list
        html.Div([
            html.H6("📋 Tracked Items",
                    style={'color': '#1a3c5e', 'fontWeight': 'bold', 'marginBottom': '10px'}),
            html.Div(id='rt-tracked-items-list',
                     children=html.P("No items tracked yet.",
                                     style={'color': '#888', 'fontStyle': 'italic'})),
        ], style={
            'background': '#f8fbff', 'border': '1px solid #d0e4f4',
            'borderRadius': '8px', 'padding': '14px 18px',
        }),
    ], style={
        'background': 'white', 'border': '1px solid #d0d7de',
        'borderRadius': '10px', 'padding': '22px 26px',
        'boxShadow': '0 2px 8px rgba(0,0,0,0.06)',
    }),

    # Hidden store for tracked items (client-side state)
    dcc.Store(id='rt-tracked-store', data=[]),

], style={'padding': '20px'})


# ═════════════════════════════════════════════════════════════════════════════
# CALLBACKS
# ═════════════════════════════════════════════════════════════════════════════

# -- Populate item dropdown based on category --------------------------------
@app.callback(
    Output('rt-track-value', 'options'),
    Output('rt-track-value', 'value'),
    Input('rt-track-category', 'value'),
)
def update_track_items(category):
    if category == 'district':
        return _district_opts, None
    elif category == 'office':
        return _office_opts, None
    elif category == 'service':
        return _service_opts, None
    return [], None


# -- Handle file upload -------------------------------------------------------
@app.callback(
    Output('rt-upload-output', 'children'),
    Input('rt-upload', 'contents'),
    State('rt-upload', 'filename'),
    prevent_initial_call=True,
)
def handle_upload(contents_list, filenames):
    if not contents_list:
        return dash.no_update

    items = []
    for content, name in zip(contents_list, filenames):
        ext = name.rsplit('.', 1)[-1].upper() if '.' in name else '?'
        icon_map = {
            'PDF': '📄', 'XLSX': '📊', 'XLS': '📊', 'CSV': '📋',
            'DOCX': '📝', 'DOC': '📝', 'JPG': '🖼️', 'JPEG': '🖼️',
            'PNG': '🖼️', 'ZIP': '🗜️',
        }
        icon = icon_map.get(ext, '📎')
        # Estimate file size from base64 payload
        try:
            b64 = content.split(',', 1)[1]
            size_kb = round(len(base64.b64decode(b64)) / 1024, 1)
            size_str = f"{size_kb} KB" if size_kb < 1024 else f"{size_kb / 1024:.1f} MB"
        except Exception:
            size_str = "—"

        items.append(
            dbc.Alert([
                html.Span(f"{icon} ", style={'fontSize': '1.2rem'}),
                html.Strong(name),
                html.Span(f"  ({ext}, {size_str})",
                          style={'fontSize': '0.85rem', 'color': '#555', 'marginLeft': '6px'}),
                html.Span(" ✅ Uploaded", style={'color': '#27ae60', 'marginLeft': '10px',
                                                 'fontWeight': '600', 'fontSize': '0.9rem'}),
            ], color='success', className='py-2 px-3 mb-1',
               style={'borderRadius': '6px', 'fontSize': '0.95rem'})
        )

    return html.Div(items)


# -- Track / Untrack buttons -------------------------------------------------
@app.callback(
    Output('rt-tracked-store', 'data'),
    Output('rt-track-status', 'children'),
    Input('rt-track-btn', 'n_clicks'),
    Input('rt-untrack-btn', 'n_clicks'),
    State('rt-track-category', 'value'),
    State('rt-track-value', 'value'),
    State('rt-track-month', 'value'),
    State('rt-tracked-store', 'data'),
    prevent_initial_call=True,
)
def handle_track(track_clicks, untrack_clicks, category, item_value, month, tracked):
    ctx = callback_context
    if not ctx.triggered or not item_value:
        return tracked, dbc.Alert("⚠️ Please select an item first.", color='warning',
                                  duration=3000, className='py-2')

    trigger = ctx.triggered[0]['prop_id'].split('.')[0]
    entry = {'category': category, 'value': item_value, 'month': month or 'All'}
    entry_key = f"{category}::{item_value}::{month or 'All'}"

    existing_keys = [f"{e['category']}::{e['value']}::{e['month']}" for e in tracked]

    if trigger == 'rt-track-btn':
        if entry_key in existing_keys:
            return tracked, dbc.Alert(f"ℹ️ '{item_value}' is already tracked.", color='info',
                                      duration=3000, className='py-2')
        updated = tracked + [entry]
        return updated, dbc.Alert(f"✅ '{item_value}' added to tracking.", color='success',
                                  duration=3000, className='py-2')

    elif trigger == 'rt-untrack-btn':
        updated = [e for e in tracked if f"{e['category']}::{e['value']}::{e['month']}" != entry_key]
        if len(updated) == len(tracked):
            return tracked, dbc.Alert(f"⚠️ '{item_value}' was not in the tracked list.", color='warning',
                                      duration=3000, className='py-2')
        return updated, dbc.Alert(f"🗑️ '{item_value}' removed from tracking.", color='danger',
                                  duration=3000, className='py-2')

    return tracked, dash.no_update


# -- Render tracked items list -----------------------------------------------
@app.callback(
    Output('rt-tracked-items-list', 'children'),
    Input('rt-tracked-store', 'data'),
)
def render_tracked(tracked):
    if not tracked:
        return html.P("No items tracked yet.", style={'color': '#888', 'fontStyle': 'italic'})

    cat_icons = {'district': '🏛️', 'office': '🏢', 'service': '⚙️'}
    cat_colors = {'district': '#1a3c5e', 'office': '#2d6a9f', 'service': '#1abc9c'}

    rows = []
    for e in tracked:
        cat = e.get('category', '?')
        icon = cat_icons.get(cat, '📌')
        color = cat_colors.get(cat, '#555')
        rows.append(
            dbc.Badge(
                [html.Span(icon + ' '), html.Strong(e['value']),
                 html.Span(f"  ({cat.title()})", style={'fontWeight': '400', 'opacity': '0.85'}),
                 html.Span(f"  | {e['month']}", style={'fontSize': '0.82rem', 'opacity': '0.75'})],
                color='light',
                style={'border': f'1.5px solid {color}', 'color': color,
                       'padding': '6px 10px', 'borderRadius': '6px',
                       'marginRight': '8px', 'marginBottom': '8px',
                       'fontSize': '0.9rem', 'display': 'inline-block'},
            )
        )
    return html.Div(rows)
