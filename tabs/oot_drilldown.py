import numpy as np
import pandas as pd
import plotly.graph_objects as go
from dash import html, dcc, Input, Output
import dash_bootstrap_components as dbc
from app import app
from data import df as df_raw

# ─────────────────────────────────────────────────────────────────────────────
# ONE-TIME DATA PREP
# ─────────────────────────────────────────────────────────────────────────────
_df = df_raw.copy()
_df["Yr"] = _df["Yr"].astype(str).str.strip().str.replace("\ufeff", "", regex=False)
_df["Mn"] = pd.to_numeric(_df["Mn"], errors="coerce").fillna(0).astype(int)
_df["month_dt"] = pd.to_datetime(
    _df["Yr"] + "-" + _df["Mn"].astype(str).str.zfill(2) + "-01",
    format="%Y-%m-%d", errors="coerce"
)
_df.dropna(subset=["month_dt"], inplace=True)
_df["Month_Year"] = _df["month_dt"].dt.strftime("%b-%Y")
_df["Disposed_Out"] = pd.to_numeric(_df["Disposed_Out"], errors="coerce").fillna(0).astype(int)
_df["Received"] = pd.to_numeric(_df["Received"], errors="coerce").fillna(0).astype(int)

_all_months = sorted(_df["Month_Year"].unique(), key=lambda x: pd.to_datetime(x, format="%b-%Y"))
_month_opts = [{"label": m, "value": m} for m in _all_months]
_default_months = _all_months[-3:] if len(_all_months) >= 3 else _all_months

# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
layout = dbc.Container([

    # Header
    html.Div([
        html.H2("🔥 Out-of-Time Drilldown", style={"color": "white", "margin": 0, "fontWeight": "600"}),
        html.P(
            "Per-Item Discovery: Find worst items, then drill down to their specific 3 worst offenders.",
            style={"color": "#ffcccc", "margin": "6px 0 0 0", "fontSize": "0.95rem"}
        ),
    ], style={
        "background": "linear-gradient(135deg, #8b0000 0%, #e63946 100%)",
        "padding": "22px 30px", "borderRadius": "12px", "marginBottom": "24px",
        "boxShadow": "0 4px 12px rgba(0,0,0,0.15)"
    }),

    # ── Filter Bar ──────────────────────────────────────────────────────────
    dbc.Card(dbc.CardBody([
        dbc.Row([
            # 1. Date Filter
            dbc.Col([
                html.Label("📅 Date Range", className="fw-bold mb-1 text-dark"),
                dcc.Dropdown(
                    id="oot-month-filter",
                    options=_month_opts,
                    value=_default_months,
                    multi=True,
                    placeholder="Select months…",
                    style={"borderRadius": "6px"}
                ),
            ], md=4),

            # 2. Toggle: Primary Focus
            dbc.Col([
                html.Label("🎯 What to rank first?", className="fw-bold mb-1 text-dark"),
                html.Div([
                    dbc.RadioItems(
                        id="oot-primary-dim",
                        options=[
                            {"label": " Worst Services", "value": "Service"},
                            {"label": " Worst Offices", "value": "Office"},
                        ],
                        value="Service",
                        inline=True,
                        className="fw-semibold"
                    )
                ], className="pt-2"),
            ], md=4),

            # 3. Discovery Matrix Sizing
            dbc.Col([
                html.Label("🏆 Matrix Size (N)", className="fw-bold mb-1 text-dark"),
                dbc.Input(id="oot-n-primary", type="number", min=1, max=50, value=6),
                html.Small("Find top N, then show 3 worst for each", className="text-muted"),
            ], md=4),

        ], className="g-3"),
    ]), className="mb-4 border-0 shadow-sm", style={"borderRadius": "12px"}),

    # ── Status bar ──────────────────────────────────────────────────────────
    html.Div(id="oot-status-bar", className="mb-3"),

    # ── Heatmap ─────────────────────────────────────────────────────────────
    dbc.Card([
        dbc.CardHeader(
            html.H5("🔥 Discovery Heatmap", className="mb-0 fw-bold text-dark"),
            style={"backgroundColor": "#ffffff", "borderBottom": "2px solid #f1f3f5", "padding": "16px 20px"}
        ),
        dbc.CardBody(
            dcc.Loading(
                color="#e63946",
                children=dcc.Graph(
                    id="oot-heatmap",
                    config={"displayModeBar": False, "scrollZoom": False},
                    style={"height": "750px"}
                )
            )
        ),
    ], className="border-0 shadow-sm mb-4", style={"borderRadius": "12px", "overflow": "hidden"}),

], fluid=True, style={"padding": "24px", "backgroundColor": "#f8f9fa"})


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _filtered_df(months):
    if not months:
        return _df
    return _df[_df["Month_Year"].isin(months)]


def _build_outputs(months, primary_dim, primary_n):
    primary_n = max(1, int(primary_n or 6))

    # HARDCODED: Always find exactly 3 worst offenders per item
    SECONDARY_M = 3

    base = _filtered_df(months)
    if base.empty:
        return None, None, 0, primary_n

    state_total_oot = max(int(base["Disposed_Out"].sum()), 1)

    # ─────────────────────────────────────────────────────────────────────
    # SCENARIO A: Rank Services First
    # ─────────────────────────────────────────────────────────────────────
    if primary_dim == "Service":
        # 1. Global Services above average volume
        srv_global = base.groupby("Service_Eng").agg(OOT=("Disposed_Out", "sum"),
                                                     Received=("Received", "sum")).reset_index()
        avg_rec = srv_global["Received"].mean() if not srv_global.empty else 0
        srv_global = srv_global[srv_global["Received"] > avg_rec]

        if srv_global.empty:
            return None, None, state_total_oot, primary_n

        # Rank by OOT %
        srv_global["OOT_Rate"] = (srv_global["OOT"] / srv_global["Received"].replace(0, np.nan) * 100).fillna(0).round(
            2)
        valid_services = srv_global.sort_values("OOT_Rate", ascending=False).head(primary_n)["Service_Eng"].tolist()

        # 2. For EACH of those services, find its specific Top 3 Offices
        valid_offices = []
        for srv in valid_services:
            df_s = base[base["Service_Eng"] == srv]
            off_s = df_s.groupby("Office_Eng").agg(OOT=("Disposed_Out", "sum"),
                                                   Received=("Received", "sum")).reset_index()
            off_s = off_s[off_s["Received"] > 0]
            off_s["OOT_Rate"] = (off_s["OOT"] / off_s["Received"] * 100).fillna(0).round(2)

            top_o = off_s.sort_values(["OOT_Rate", "OOT"], ascending=[False, False]).head(SECONDARY_M)
            valid_offices.extend(top_o["Office_Eng"].tolist())

        valid_offices = list(dict.fromkeys(valid_offices))  # Remove duplicates

    # ─────────────────────────────────────────────────────────────────────
    # SCENARIO B: Rank Offices First
    # ─────────────────────────────────────────────────────────────────────
    else:
        # 1. Global Offices above average volume
        off_global = base.groupby("Office_Eng").agg(OOT=("Disposed_Out", "sum"),
                                                    Received=("Received", "sum")).reset_index()
        avg_rec = off_global["Received"].mean() if not off_global.empty else 0
        off_global = off_global[off_global["Received"] > avg_rec]

        if off_global.empty:
            return None, None, state_total_oot, primary_n

        # Rank by OOT %
        off_global["OOT_Rate"] = (off_global["OOT"] / off_global["Received"].replace(0, np.nan) * 100).fillna(0).round(
            2)
        valid_offices = off_global.sort_values("OOT_Rate", ascending=False).head(primary_n)["Office_Eng"].tolist()

        # 2. For EACH of those offices, find its specific Top 3 Services
        valid_services = []
        for off in valid_offices:
            df_o = base[base["Office_Eng"] == off]
            srv_o = df_o.groupby("Service_Eng").agg(OOT=("Disposed_Out", "sum"),
                                                    Received=("Received", "sum")).reset_index()
            srv_o = srv_o[srv_o["Received"] > 0]
            srv_o["OOT_Rate"] = (srv_o["OOT"] / srv_o["Received"] * 100).fillna(0).round(2)

            top_s = srv_o.sort_values(["OOT_Rate", "OOT"], ascending=[False, False]).head(SECONDARY_M)
            valid_services.extend(top_s["Service_Eng"].tolist())

        valid_services = list(dict.fromkeys(valid_services))  # Remove duplicates

    # ─────────────────────────────────────────────────────────────────────
    # 3. BUILD THE GUARANTEED COVERAGE MATRICES (OOT & RECEIVED)
    # ─────────────────────────────────────────────────────────────────────
    if not valid_services or not valid_offices:
        return None, None, state_total_oot, primary_n

    base_f = base[base["Service_Eng"].isin(valid_services) & base["Office_Eng"].isin(valid_offices)]

    # We aggregate both OOT and Received at the cell level
    cell = base_f.groupby(["Office_Eng", "Service_Eng"]).agg(OOT=("Disposed_Out", "sum"),
                                                             Received=("Received", "sum")).reset_index()

    pivot_oot = cell.pivot(index="Office_Eng", columns="Service_Eng", values="OOT").reindex(index=valid_offices,
                                                                                            columns=valid_services).fillna(
        0)
    pivot_rec = cell.pivot(index="Office_Eng", columns="Service_Eng", values="Received").reindex(index=valid_offices,
                                                                                                 columns=valid_services).fillna(
        0)

    return pivot_oot, pivot_rec, state_total_oot, primary_n


def _shorten(label, n=40):
    return label if len(label) <= n else label[: n - 1] + "…"


def _build_heatmap_fig(pivot_oot, pivot_rec, months):
    y_labels = pivot_oot.index.tolist()
    x_labels = pivot_oot.columns.tolist()
    z_vals = pivot_oot.values.tolist()

    # Generate Date string for hover
    month_str = ", ".join(months) if months else "All Available Data"

    # Clean, basic HTML for Plotly hover text
    hover_text = []
    for off in y_labels:
        row_hover = []
        for srv in x_labels:
            cell_oot = int(pivot_oot.loc[off, srv])
            cell_rec = int(pivot_rec.loc[off, srv])

            row_hover.append(
                f"<b>Service:</b> {srv}<br>"
                f"<b>Office:</b> {off}<br>"
                f"<b>Date Range:</b> {month_str}<br>"
                f"────────────────<br>"
                f"<b>Total Received:</b> {cell_rec:,}<br>"
                f"<b>Total OOT:</b> {cell_oot:,}"
            )
        hover_text.append(row_hover)

    fig = go.Figure(go.Heatmap(
        z=z_vals,
        x=[_shorten(s, 35) for s in x_labels],
        y=[_shorten(o, 40) for o in y_labels],
        text=hover_text,
        hoverinfo="text",
        colorscale=[
            [0.00, "#f8f9fa"],  # 0 values = soft off-white background
            [0.01, "#ffe3e3"],  # very light red
            [0.20, "#ff9999"],
            [0.50, "#e63946"],  # vibrant red
            [0.80, "#bd0026"],  # dark red
            [1.00, "#6b0000"],  # deep maroon
        ],
        colorbar=dict(
            title=dict(text="OOT Volume", side="right", font=dict(size=12, color="#495057")),
            tickfont=dict(size=11, color="#495057"),
            thickness=14,
            outlinewidth=0
        ),
        xgap=3, ygap=3,  # Thicker gaps for cleaner separation
    ))

    fig.update_layout(
        xaxis=dict(
            tickfont=dict(size=10, color="#495057"),
            title=dict(text="Service", font=dict(color="#6c757d")),
            tickangle=-40,
            side="bottom",
        ),
        yaxis=dict(
            tickfont=dict(size=10, color="#495057"),
            title=dict(text="Office", font=dict(color="#6c757d")),
            autorange="reversed",
        ),
        margin=dict(l=300, r=40, t=20, b=240),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(family="Segoe UI, Arial, sans-serif"),
        hoverlabel=dict(
            bgcolor="white",
            font_size=13,
            bordercolor="#e63946",
            font_family="Segoe UI, Arial, sans-serif"
        ),
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output("oot-heatmap", "figure"),
    Output("oot-status-bar", "children"),
    Input("oot-month-filter", "value"),
    Input("oot-primary-dim", "value"),
    Input("oot-n-primary", "value"),
)
def update_oot(months, primary_dim, primary_n):
    pivot_oot, pivot_rec, state_total_oot, p_n = _build_outputs(
        months, primary_dim, primary_n
    )

    def _empty_fig(msg):
        f = go.Figure()
        f.update_layout(
            annotations=[dict(text=msg, showarrow=False, xref="paper", yref="paper", x=0.5, y=0.5,
                              font=dict(size=16, color="#adb5bd"))],
            paper_bgcolor="#ffffff", plot_bgcolor="#ffffff"
        )
        return f

    if pivot_oot is None or (hasattr(pivot_oot, "empty") and pivot_oot.empty):
        warn = dbc.Alert("⚠️ No data matches the current filters.", color="warning",
                         className="rounded-3 shadow-sm border-0")
        return _empty_fig("No data available for these parameters"), warn

    fig = _build_heatmap_fig(pivot_oot, pivot_rec, months)

    month_str = ", ".join(months) if months else "All Available Months"
    primary_txt = "Services" if primary_dim == "Service" else "Offices"

    # Calculate exact rows x cols plotted
    n_rows, n_cols = pivot_oot.shape

    status = dbc.Alert([
        html.B("📅 "), f"{month_str}  |  ",
        html.B("State OOT: "), f"{state_total_oot:,}  |  ",
        html.B("Ranked First: "), f"{primary_txt}  ",
        html.Span("(Filtered to > Average Volume)", style={"fontSize": "0.85rem", "color": "#8b0000"}),
        "  |  ",
        html.B("Matrix Plotted: "), f"{n_rows} Rows × {n_cols} Columns",
    ], color="danger", className="py-2 px-3 border-0 shadow-sm rounded-3 mb-2", style={"backgroundColor": "#f8d7da"})

    return fig, status