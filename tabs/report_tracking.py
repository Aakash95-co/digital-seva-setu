import os
import base64
import datetime
import json

import dash
from dash import html, dcc, Input, Output, State, ALL, callback_context
import dash_bootstrap_components as dbc

from app import app
from data import df_adv

# ═════════════════════════════════════════════════════════════════════════════
# UPLOAD FOLDER  (saved one level above this file → <project>/uploads/)
# ═════════════════════════════════════════════════════════════════════════════
_UPLOAD_DIR = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', 'uploads')
)
os.makedirs(_UPLOAD_DIR, exist_ok=True)

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

_ICON_MAP = {
    'PDF': '📄', 'XLSX': '📊', 'XLS': '📊', 'CSV': '📋',
    'DOCX': '📝', 'DOC': '📝', 'JPG': '🖼️', 'JPEG': '🖼️',
    'PNG': '🖼️', 'ZIP': '🗜️',
}

# Unique key for a tracked entry
def _entry_key(e):
    return f"D:{e.get('district','')};O:{e.get('office','')};S:{e.get('service','')}"


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
    # SECTION 1 — REPORTING (file upload + saved-files list)
    # ══════════════════════════════════════════════════════════════════════
    html.Div([
        html.H4("📤 Reporting — Upload Files",
                style={'color': '#1a3c5e', 'borderBottom': '2px solid #2d6a9f',
                       'paddingBottom': '8px', 'marginBottom': '16px'}),
        html.P("Upload reports (PDF, Excel, CSV, Word, images …). "
               "Files are saved on the server and listed below for future reference.",
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
                'width': '100%', 'minHeight': '130px', 'lineHeight': '1.6',
                'borderWidth': '2px', 'borderStyle': 'dashed', 'borderColor': '#2d6a9f',
                'borderRadius': '10px', 'backgroundColor': '#f8fbff',
                'display': 'flex', 'alignItems': 'center',
                'justifyContent': 'center', 'cursor': 'pointer',
            },
            multiple=True,
        ),
        html.Div(id='rt-upload-status', style={'marginTop': '10px'}),

        # ── Saved files table ────────────────────────────────────────────
        html.Div([
            html.H6("📋 Uploaded Files",
                    style={'color': '#1a3c5e', 'fontWeight': 'bold', 'marginBottom': '10px'}),
            html.Div(id='rt-uploads-list'),
        ], style={'background': '#f8fbff', 'border': '1px solid #d0e4f4',
                  'borderRadius': '8px', 'padding': '14px 18px', 'marginTop': '16px'}),

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
        html.P("Select any combination of District, Office and Service, "
               "then click Track to add a row to the watchlist.",
               style={'color': '#666', 'fontSize': '0.92rem', 'marginBottom': '18px'}),

        dbc.Row([
            dbc.Col([
                html.Label("🏛️ District", style={'fontWeight': 'bold', 'marginBottom': '4px'}),
                dcc.Dropdown(
                    id='rt-track-district',
                    options=_district_opts,
                    value=None, placeholder='Select district…', clearable=True,
                ),
            ], md=4),
            dbc.Col([
                html.Label("🏢 Office", style={'fontWeight': 'bold', 'marginBottom': '4px'}),
                dcc.Dropdown(
                    id='rt-track-office',
                    options=_office_opts,
                    value=None, placeholder='Select office…', clearable=True,
                ),
            ], md=4),
            dbc.Col([
                html.Label("⚙️ Service", style={'fontWeight': 'bold', 'marginBottom': '4px'}),
                dcc.Dropdown(
                    id='rt-track-service',
                    options=_service_opts,
                    value=None, placeholder='Select service…', clearable=True,
                ),
            ], md=3),
            dbc.Col([
                html.Label("\u00a0", style={'display': 'block', 'marginBottom': '4px'}),
                dbc.Button("➕ Track", id='rt-track-btn', color='success',
                           style={'width': '100%', 'fontWeight': 'bold'}),
            ], md=1),
        ], className='mb-3 align-items-end'),

        html.Div(id='rt-track-status', style={'marginBottom': '12px'}),

        # ── Watchlist table ──────────────────────────────────────────────
        html.Div([
            html.H6("📋 Tracking Watchlist",
                    style={'color': '#1a3c5e', 'fontWeight': 'bold', 'marginBottom': '10px'}),
            html.Div(id='rt-tracked-table'),
        ], style={'background': '#f8fbff', 'border': '1px solid #d0e4f4',
                  'borderRadius': '8px', 'padding': '14px 18px'}),

    ], style={
        'background': 'white', 'border': '1px solid #d0d7de',
        'borderRadius': '10px', 'padding': '22px 26px',
        'boxShadow': '0 2px 8px rgba(0,0,0,0.06)',
    }),

    # ── Persistent stores (browser localStorage → survives page refresh) ─
    dcc.Store(id='rt-tracked-store', storage_type='local', data=[]),
    dcc.Store(id='rt-uploads-store', storage_type='local', data=[]),

], style={'padding': '20px'})


# ═════════════════════════════════════════════════════════════════════════════
# CALLBACKS
# ═════════════════════════════════════════════════════════════════════════════

# ── 1. Save uploaded files to disk + update metadata store ──────────────────
@app.callback(
    Output('rt-uploads-store', 'data'),
    Output('rt-upload-status', 'children'),
    Input('rt-upload', 'contents'),
    State('rt-upload', 'filename'),
    State('rt-uploads-store', 'data'),
    prevent_initial_call=True,
)
def handle_upload(contents_list, filenames, current_uploads):
    if not contents_list:
        return current_uploads, dash.no_update

    existing_names = {e['filename'] for e in (current_uploads or [])}
    new_entries    = list(current_uploads or [])
    alerts         = []

    for content, name in zip(contents_list, filenames):
        ext  = name.rsplit('.', 1)[-1].upper() if '.' in name else '?'
        icon = _ICON_MAP.get(ext, '📎')
        try:
            _header, b64 = content.split(',', 1)
            file_bytes   = base64.b64decode(b64)
            size_kb      = round(len(file_bytes) / 1024, 1)
            size_str     = f"{size_kb} KB" if size_kb < 1024 else f"{size_kb / 1024:.1f} MB"

            ts        = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_name = "".join(c if (c.isalnum() or c in '._- ()') else '_' for c in name)
            save_name = f"{ts}_{safe_name}"
            with open(os.path.join(_UPLOAD_DIR, save_name), 'wb') as fh:
                fh.write(file_bytes)

            upload_dt = datetime.datetime.now().strftime('%d-%b-%Y %H:%M')

            if name not in existing_names:
                new_entries.append({
                    'filename': name,
                    'saved_as': save_name,
                    'ext':      ext,
                    'size':     size_str,
                    'uploaded': upload_dt,
                })
                existing_names.add(name)

            alerts.append(
                dbc.Alert([
                    html.Span(f"{icon} ", style={'fontSize': '1.2rem'}),
                    html.Strong(name),
                    html.Span(f"  ({ext}, {size_str})",
                              style={'fontSize': '0.85rem', 'color': '#555', 'marginLeft': '6px'}),
                    html.Span(" ✅ Saved", style={'color': '#27ae60', 'marginLeft': '10px',
                                                   'fontWeight': '600', 'fontSize': '0.9rem'}),
                ], color='success', className='py-2 px-3 mb-1',
                   style={'borderRadius': '6px', 'fontSize': '0.95rem'}),
            )
        except Exception as ex:
            alerts.append(
                dbc.Alert(f"❌ Error saving {name}: {ex}", color='danger', className='py-2')
            )

    return new_entries, html.Div(alerts)


# ── 2. Render saved-files table ──────────────────────────────────────────────
@app.callback(
    Output('rt-uploads-list', 'children'),
    Input('rt-uploads-store', 'data'),
)
def render_uploads(uploads):
    if not uploads:
        return html.P("No files uploaded yet.",
                      style={'color': '#888', 'fontStyle': 'italic'})

    rows = []
    for i, e in enumerate(uploads, 1):
        icon = _ICON_MAP.get(e.get('ext', ''), '📎')
        rows.append(html.Tr([
            html.Td(str(i), style={'color': '#888', 'width': '36px', 'textAlign': 'center'}),
            html.Td([html.Span(icon + " "), html.Strong(e.get('filename', '—'))]),
            html.Td(e.get('ext', '—'),      style={'color': '#2d6a9f', 'textAlign': 'center'}),
            html.Td(e.get('size', '—'),     style={'color': '#555', 'textAlign': 'right'}),
            html.Td(e.get('uploaded', '—'), style={'color': '#555', 'fontSize': '0.85rem'}),
        ]))

    return dbc.Table(
        [
            html.Thead(html.Tr([
                html.Th("#"),
                html.Th("File Name"),
                html.Th("Type",        style={'textAlign': 'center'}),
                html.Th("Size",        style={'textAlign': 'right'}),
                html.Th("Uploaded On"),
            ]), style={'background': '#eef4fb'}),
            html.Tbody(rows),
        ],
        bordered=True, hover=True, responsive=True, size='sm',
        style={'marginBottom': '0', 'fontSize': '0.9rem'},
    )


# ── 3. Add item to tracking store ────────────────────────────────────────────
@app.callback(
    Output('rt-tracked-store', 'data'),
    Output('rt-track-status', 'children'),
    Input('rt-track-btn', 'n_clicks'),
    State('rt-track-district', 'value'),
    State('rt-track-office',   'value'),
    State('rt-track-service',  'value'),
    State('rt-tracked-store',  'data'),
    prevent_initial_call=True,
)
def handle_track(n_clicks, district, office, service, tracked):
    if not any([district, office, service]):
        return tracked, dbc.Alert(
            "⚠️ Please select at least one of District, Office or Service.",
            color='warning', duration=3000, className='py-2',
        )

    entry = {
        'district': district or '',
        'office':   office   or '',
        'service':  service  or '',
    }
    key = _entry_key(entry)

    if key in [_entry_key(e) for e in tracked]:
        return tracked, dbc.Alert(
            "ℹ️ This combination is already in the watchlist.",
            color='info', duration=3000, className='py-2',
        )

    return tracked + [entry], dbc.Alert(
        "✅ Added to tracking watchlist.",
        color='success', duration=2000, className='py-2',
    )


# ── 4. Remove row from tracking store (per-row Untrack button) ───────────────
@app.callback(
    Output('rt-tracked-store', 'data', allow_duplicate=True),
    Input({'type': 'rt-untrack-row', 'index': ALL}, 'n_clicks'),
    State('rt-tracked-store', 'data'),
    prevent_initial_call=True,
)
def handle_untrack_row(n_clicks_list, tracked):
    ctx = callback_context
    if not ctx.triggered or not any(n for n in (n_clicks_list or []) if n):
        return tracked

    triggered_prop = ctx.triggered[0]['prop_id']
    try:
        id_dict = json.loads(triggered_prop.rsplit('.', 1)[0])
        key     = id_dict['index']
    except Exception:
        return tracked

    return [e for e in tracked if _entry_key(e) != key]


# ── 5. Render tracking watchlist table ───────────────────────────────────────
@app.callback(
    Output('rt-tracked-table', 'children'),
    Input('rt-tracked-store', 'data'),
)
def render_tracked_table(tracked):
    if not tracked:
        return html.P(
            "No items tracked yet. Use the dropdowns above to add a combination.",
            style={'color': '#888', 'fontStyle': 'italic'},
        )

    rows = []
    for i, e in enumerate(tracked, 1):
        district = e.get('district') or '—'
        office   = e.get('office')   or '—'
        service  = e.get('service')  or '—'
        key      = _entry_key(e)
        rows.append(html.Tr([
            html.Td(str(i), style={'color': '#888', 'width': '36px', 'textAlign': 'center'}),
            html.Td(district),
            html.Td(office),
            html.Td(service),
            html.Td(
                dbc.Button(
                    "Untrack", size='sm', color='danger', outline=True,
                    id={'type': 'rt-untrack-row', 'index': key},
                    style={'padding': '2px 12px', 'fontSize': '0.82rem'},
                )
            ),
        ]))

    return dbc.Table(
        [
            html.Thead(html.Tr([
                html.Th("#"),
                html.Th("District"),
                html.Th("Office"),
                html.Th("Service"),
                html.Th("Action"),
            ]), style={'background': '#eef4fb'}),
            html.Tbody(rows),
        ],
        bordered=True, hover=True, responsive=True, size='sm',
        style={'marginBottom': '0', 'fontSize': '0.9rem'},
    )
