import calendar
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from dash import html, dcc, Input, Output
import dash_bootstrap_components as dbc
from app import app
from data import FY_DATA

_MONTH_ABR = {v: k for k, v in enumerate(calendar.month_abbr) if v}


# ─────────────────────────────────────────────────────────────────────────────
# DATA PREP
# ─────────────────────────────────────────────────────────────────────────────
def _prep_df(raw_df):
    df = raw_df.copy()
    df["Yr"] = df["Yr"].astype(str).str.strip().str.replace("\ufeff", "", regex=False)
    df["Mn"] = pd.to_numeric(df["Mn"], errors="coerce").fillna(0).astype(int)
    df["Disposed_Out"] = pd.to_numeric(df["Disposed_Out"], errors="coerce").fillna(0).astype(int)
    df["Received"] = pd.to_numeric(df["Received"], errors="coerce").fillna(0).astype(int)
    return df


def _shorten(label, n=35):
    return label if len(label) <= n else label[: n - 1] + "…"


def _build_data(fy, month="ALL", top_n=5):
    top_n = int(top_n or 5)
    df = _prep_df(FY_DATA[fy]["df"])
    fy_label = FY_DATA[fy]["label"]

    # ── Month filter ─────────────────────────────────────────────────────────
    month_label = "All Months"
    if month and month != "ALL":
        mn = _MONTH_ABR.get(month.split("-")[0], 0)
        if mn:
            df = df[df["Mn"] == mn]
            month_label = month

    state_total_oot = max(int(df["Disposed_Out"].sum()), 1)

    # ── Top N Offices by raw OOT count ──────────────────────────────────────
    off_agg = (
        df.groupby("Office_Eng")
        .agg(OOT=("Disposed_Out", "sum"), Received=("Received", "sum"))
        .reset_index()
    )
    off_agg["OOT_Rate"] = (
        off_agg["OOT"] / off_agg["Received"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    off_agg["Share"] = (off_agg["OOT"] / state_total_oot * 100).round(1)
    top_offices = off_agg.nlargest(top_n, "OOT")["Office_Eng"].tolist()

    # ── Top N Services by raw OOT count ─────────────────────────────────────
    srv_agg = (
        df.groupby("Service_Eng")
        .agg(OOT=("Disposed_Out", "sum"), Received=("Received", "sum"))
        .reset_index()
    )
    srv_agg["OOT_Rate"] = (
        srv_agg["OOT"] / srv_agg["Received"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    srv_agg["Share"] = (srv_agg["OOT"] / state_total_oot * 100).round(1)
    top_services = srv_agg.nlargest(top_n, "OOT")["Service_Eng"].tolist()

    # ── NxN Intersection ────────────────────────────────────────────────────
    base_f = df[
        df["Office_Eng"].isin(top_offices) & df["Service_Eng"].isin(top_services)
    ]
    cell = (
        base_f.groupby(["Office_Eng", "Service_Eng"])
        .agg(OOT=("Disposed_Out", "sum"), Received=("Received", "sum"))
        .reset_index()
    )
    pivot_oot = (
        cell.pivot(index="Office_Eng", columns="Service_Eng", values="OOT")
        .reindex(index=top_offices, columns=top_services)
        .fillna(0)
    )
    pivot_rec = (
        cell.pivot(index="Office_Eng", columns="Service_Eng", values="Received")
        .reindex(index=top_offices, columns=top_services)
        .fillna(0)
    )

    off_top = (
        off_agg[off_agg["Office_Eng"].isin(top_offices)]
        .set_index("Office_Eng")
        .reindex(top_offices)
    )
    srv_top = (
        srv_agg[srv_agg["Service_Eng"].isin(top_services)]
        .set_index("Service_Eng")
        .reindex(top_services)
    )

    matrix_oot = int(pivot_oot.values.sum())
    matrix_pct = round(matrix_oot / state_total_oot * 100, 1)

    return dict(
        pivot_oot=pivot_oot, pivot_rec=pivot_rec,
        off_top=off_top, srv_top=srv_top,
        state_total_oot=state_total_oot,
        matrix_oot=matrix_oot, matrix_pct=matrix_pct,
        fy_label=fy_label,
        top5_offices=top_offices, top5_services=top_services,
        top_n=top_n,
        month_label=month_label,
    )


# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
layout = dbc.Container([

    # Header
    html.Div([
        html.H2("🔥 Out-of-Time Drilldown", style={"color": "white", "margin": 0, "fontWeight": "600"}),
        html.P(
            "Top N Offices & Top N Services driving Out-of-Time cases — and their intersection.",
            style={"color": "#ffcccc", "margin": "6px 0 0 0", "fontSize": "0.95rem"}
        ),
    ], style={
        "background": "linear-gradient(135deg, #8b0000 0%, #e63946 100%)",
        "padding": "22px 30px", "borderRadius": "12px", "marginBottom": "24px",
        "boxShadow": "0 4px 12px rgba(0,0,0,0.15)"
    }),

    # ── Filters ─────────────────────────────────────────────────────────────
    dbc.Card([
        dbc.CardBody([
            dbc.Row([
                dbc.Col([
                    html.Label("📅 Month", className="fw-semibold text-secondary mb-1",
                               style={"fontSize": "0.82rem"}),
                    dcc.Dropdown(
                        id="oot-month",
                        value="ALL",
                        clearable=False,
                        style={"fontSize": "0.85rem"},
                    ),
                ], md=5),
                dbc.Col([
                    html.Label("🔢 Matrix Size", className="fw-semibold text-secondary mb-1",
                               style={"fontSize": "0.82rem"}),
                    dcc.Input(
                        id="oot-top-n",
                        type="number",
                        value=5,
                        min=2,
                        max=20,
                        step=1,
                        debounce=True,
                        style={"width": "80px", "padding": "6px", "borderRadius": "4px",
                               "border": "1px solid #ccc", "fontSize": "0.9rem"},
                    ),
                ], md=5),
            ], align="center"),
        ], style={"padding": "12px 20px"}),
    ], className="border-0 shadow-sm mb-4", style={"borderRadius": "10px", "backgroundColor": "#fff"}),

    # Status bar
    html.Div(id="oot-status-bar", className="mb-4"),

    # ── Bar Charts Row ───────────────────────────────────────────────────────
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardHeader(
                    html.H6(id="oot-office-bar-title",
                            children="🏢 Top Offices by OOT Volume",
                            className="mb-0 fw-bold text-dark"),
                    style={"backgroundColor": "#fff5f5", "borderBottom": "2px solid #f8d7da"}
                ),
                dbc.CardBody(
                    dcc.Loading(
                        color="#e63946",
                        children=dcc.Graph(
                            id="oot-office-bar",
                            config={"displayModeBar": False},
                            style={"height": "280px"}
                        )
                    )
                ),
            ], className="border-0 shadow-sm h-100", style={"borderRadius": "12px"}),
        ], md=6),

        dbc.Col([
            dbc.Card([
                dbc.CardHeader(
                    html.H6(id="oot-service-bar-title",
                            children="⚙️ Top Services by OOT Volume",
                            className="mb-0 fw-bold text-dark"),
                    style={"backgroundColor": "#fff5f5", "borderBottom": "2px solid #f8d7da"}
                ),
                dbc.CardBody(
                    dcc.Loading(
                        color="#e63946",
                        children=dcc.Graph(
                            id="oot-service-bar",
                            config={"displayModeBar": False},
                            style={"height": "280px"}
                        )
                    )
                ),
            ], className="border-0 shadow-sm h-100", style={"borderRadius": "12px"}),
        ], md=6),
    ], className="mb-4"),

    # ── NxN Annotated Heatmap ────────────────────────────────────────────────
    dbc.Card([
        dbc.CardHeader(
            html.Div([
                html.H5(id="oot-heatmap-title",
                        children="🔥 5 × 5 Intersection Heatmap",
                        className="mb-0 fw-bold text-dark d-inline"),
                html.Span(
                    " — Each cell shows OOT count (bold) + % of State OOT",
                    style={"fontSize": "0.85rem", "color": "#6c757d", "marginLeft": "10px"}
                ),
            ]),
            style={"backgroundColor": "#ffffff", "borderBottom": "2px solid #f1f3f5", "padding": "16px 20px"}
        ),
        dbc.CardBody(
            dcc.Loading(
                color="#e63946",
                children=dcc.Graph(
                    id="oot-heatmap",
                    config={"displayModeBar": False},
                    style={"height": "520px"}
                )
            )
        ),
    ], className="border-0 shadow-sm", style={"borderRadius": "12px", "overflow": "hidden"}),

    # ── Matrix Summary ───────────────────────────────────────────────────────
    html.Div(id="oot-matrix-summary", className="mt-3"),

], fluid=True, style={"padding": "24px", "backgroundColor": "#f8f9fa"})


# ─────────────────────────────────────────────────────────────────────────────
# FIGURE BUILDERS
# ─────────────────────────────────────────────────────────────────────────────
def _bar_fig(names, oot_vals, rates, color):
    labels = [_shorten(n, 45) for n in names]
    fig = go.Figure(go.Bar(
        x=oot_vals,
        y=labels,
        orientation="h",
        marker_color=color,
        text=[f"{v:,}  ({r:.1f}% of State OOT)" for v, r in zip(oot_vals, rates)],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>OOT: %{x:,}<extra></extra>",
    ))
    fig.update_layout(
        xaxis=dict(title="OOT Cases", tickfont=dict(size=10)),
        yaxis=dict(autorange="reversed", tickfont=dict(size=10)),
        margin=dict(l=10, r=130, t=10, b=30),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#f8f9fa",
        height=280,
        font=dict(family="Segoe UI, Arial, sans-serif"),
    )
    return fig


def _heatmap_fig(pivot_oot, pivot_rec, state_total_oot):
    y_labels_full = pivot_oot.index.tolist()
    x_labels_full = pivot_oot.columns.tolist()
    y_labels = [_shorten(o, 40) for o in y_labels_full]
    x_labels = [_shorten(s, 35) for s in x_labels_full]

    z = pivot_oot.values
    max_val = z.max() if z.max() > 0 else 1

    annotations = []
    hover_text = []
    for i, off in enumerate(y_labels_full):
        row_hover = []
        for j, srv in enumerate(x_labels_full):
            cell_oot = int(pivot_oot.iloc[i, j])
            cell_rec = int(pivot_rec.iloc[i, j])
            share = round(cell_oot / state_total_oot * 100, 1)
            oot_rate = round(cell_oot / cell_rec * 100, 1) if cell_rec > 0 else 0.0
            font_color = "white" if (cell_oot / max_val) > 0.45 else "#222222"
            annotations.append(dict(
                x=x_labels[j],
                y=y_labels[i],
                text=f"<b>{cell_oot:,}</b><br>{share:.1f}%",
                showarrow=False,
                font=dict(size=10, color=font_color),
                align="center",
                xref="x", yref="y",
            ))
            row_hover.append(
                f"<b>Office:</b> {off}<br>"
                f"<b>Service:</b> {srv}<br>"
                f"────────────────<br>"
                f"<b>Received:</b> {cell_rec:,}<br>"
                f"<b>OOT:</b> {cell_oot:,}<br>"
                f"<b>% of State OOT:</b> {share:.1f}%<br>"
                f"<b>OOT Rate:</b> {oot_rate:.1f}%"
            )
        hover_text.append(row_hover)

    fig = go.Figure(go.Heatmap(
        z=z.tolist(),
        x=x_labels,
        y=y_labels,
        text=hover_text,
        hoverinfo="text",
        colorscale=[
            [0.00, "#f8f9fa"],
            [0.01, "#ffe3e3"],
            [0.20, "#ff9999"],
            [0.50, "#e63946"],
            [0.80, "#bd0026"],
            [1.00, "#6b0000"],
        ],
        colorbar=dict(
            title=dict(text="OOT Volume", side="right", font=dict(size=11, color="#495057")),
            tickfont=dict(size=10, color="#495057"),
            thickness=12,
            outlinewidth=0,
        ),
        xgap=4, ygap=4,
    ))
    fig.update_layout(
        annotations=annotations,
        xaxis=dict(
            tickfont=dict(size=10, color="#495057"),
            title=dict(text="Service", font=dict(color="#6c757d")),
            tickangle=-35,
            side="bottom",
        ),
        yaxis=dict(
            tickfont=dict(size=10, color="#495057"),
            title=dict(text="Office", font=dict(color="#6c757d")),
            autorange="reversed",
        ),
        margin=dict(l=280, r=60, t=20, b=200),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(family="Segoe UI, Arial, sans-serif"),
        hoverlabel=dict(
            bgcolor="white", font_size=12,
            bordercolor="#e63946",
            font_family="Segoe UI, Arial, sans-serif",
        ),
        height=520,
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACKS
# ─────────────────────────────────────────────────────────────────────────────

# -- Populate month dropdown when FY changes --
@app.callback(
    Output("oot-month", "options"),
    Output("oot-month", "value"),
    Input("fy-store", "data"),
)
def update_month_options(fy):
    opts = [{"label": "All Months", "value": "ALL"}]
    for o in FY_DATA[fy]["month_options"]:
        if o["value"] != "ALL_MONTHS":
            opts.append({"label": o["label"], "value": o["value"]})
    return opts, "ALL"


# -- Main analysis --
@app.callback(
    Output("oot-office-bar", "figure"),
    Output("oot-service-bar", "figure"),
    Output("oot-heatmap", "figure"),
    Output("oot-status-bar", "children"),
    Output("oot-heatmap-title", "children"),
    Output("oot-office-bar-title", "children"),
    Output("oot-service-bar-title", "children"),
    Output("oot-matrix-summary", "children"),   # ← add this line
    Input("fy-store", "data"),
    Input("oot-month", "value"),
    Input("oot-top-n", "value"),
)
def update_oot(fy, month, top_n):
    top_n = int(top_n or 5)
    d = _build_data(fy, month=month or "ALL", top_n=top_n)

    # Sort ascending so worst appears at top of horizontal bar chart
    off_sorted = d["off_top"].sort_values("OOT", ascending=True)
    fig_off = _bar_fig(
        off_sorted.index.tolist(),
        off_sorted["OOT"].tolist(),
        off_sorted["Share"].tolist(),
        "#e63946",
    )

    srv_sorted = d["srv_top"].sort_values("OOT", ascending=True)
    fig_srv = _bar_fig(
        srv_sorted.index.tolist(),
        srv_sorted["OOT"].tolist(),
        srv_sorted["Share"].tolist(),
        "#bd0026",
    )

    fig_heat = _heatmap_fig(d["pivot_oot"], d["pivot_rec"], d["state_total_oot"])

    # Find the single worst cell
    flat = [
        (off, srv, int(d["pivot_oot"].loc[off, srv]))
        for off in d["top5_offices"]
        for srv in d["top5_services"]
    ]
    flat.sort(key=lambda x: x[2], reverse=True)
    top1_off, top1_srv, top1_oot = flat[0]
    top1_pct = round(top1_oot / d["state_total_oot"] * 100, 1)

    n = top_n
    status = dbc.Alert([
        html.B(f"📅 {d['fy_label']}"),
        "  |  ",
        html.B("Period: "), d["month_label"],
        "  |  ",
        html.B("State Total OOT: "), f"{d['state_total_oot']:,}",
        "  |  ",
        html.B(f"Top {n}×{n} Matrix: "),
        html.Span(
            f"{d['matrix_oot']:,} cases ({d['matrix_pct']}% of state OOT)",
            style={"color": "#8b0000", "fontWeight": "600"}
        ),
        "  |  ",
        html.B("Worst Cell: "),
        html.Span(
            f"{_shorten(top1_off, 30)} × {_shorten(top1_srv, 30)} → {top1_oot:,} ({top1_pct}% of state OOT)",
            style={"color": "#8b0000", "fontWeight": "600"}
        ),
    ], color="danger", className="py-2 px-3 border-0 shadow-sm rounded-3",
       style={"backgroundColor": "#f8d7da", "fontSize": "0.9rem"})

    heatmap_title = f"🔥 {n} × {n} Intersection Heatmap"
    office_bar_title = f"🏢 Top {n} Offices by OOT Volume"
    service_bar_title = f"⚙️ Top {n} Services by OOT Volume"

    n_cells = top_n * top_n
    summary = dbc.Alert([
        html.Span("📊 "),
        html.B(f"{n}×{n} = {n_cells} entities"),
        f" (Top {n} Offices × Top {n} Services) account for ",
        html.Span(
            f"{d['matrix_pct']}% of total State OOT",
            style={"color": "#8b0000", "fontWeight": "700", "fontSize": "1.05rem"}
        ),
        f"  —  {d['matrix_oot']:,} out of {d['state_total_oot']:,} cases",
    ], color="warning", className="py-2 px-4 shadow-sm rounded-3 text-center",
   style={"backgroundColor": "#fff3cd", "border": "1px solid #ffc107", "fontSize": "0.95rem"})

    return fig_off, fig_srv, fig_heat, status, heatmap_title, office_bar_title, service_bar_title, summary