import numpy as np
import pandas as pd
import plotly.graph_objects as go
from dash import html, dcc, Input, Output
import dash_bootstrap_components as dbc
from app import app
from data import FY_DATA

TOP_N = 5  # Always 5 offices × 5 services


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


def _build_data(fy):
    df = _prep_df(FY_DATA[fy]["df"])
    fy_label = FY_DATA[fy]["label"]
    state_total_oot = max(int(df["Disposed_Out"].sum()), 1)

    # ── Top 5 Offices by raw OOT count ──────────────────────────────────────
    off_agg = (
        df.groupby("Office_Eng")
        .agg(OOT=("Disposed_Out", "sum"), Received=("Received", "sum"))
        .reset_index()
    )
    off_agg["OOT_Rate"] = (
        off_agg["OOT"] / off_agg["Received"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    top5_offices = off_agg.nlargest(TOP_N, "OOT")["Office_Eng"].tolist()

    # ── Top 5 Services by raw OOT count ─────────────────────────────────────
    srv_agg = (
        df.groupby("Service_Eng")
        .agg(OOT=("Disposed_Out", "sum"), Received=("Received", "sum"))
        .reset_index()
    )
    srv_agg["OOT_Rate"] = (
        srv_agg["OOT"] / srv_agg["Received"].replace(0, np.nan) * 100
    ).fillna(0).round(1)
    top5_services = srv_agg.nlargest(TOP_N, "OOT")["Service_Eng"].tolist()

    # ── 5×5 Intersection ────────────────────────────────────────────────────
    base_f = df[
        df["Office_Eng"].isin(top5_offices) & df["Service_Eng"].isin(top5_services)
    ]
    cell = (
        base_f.groupby(["Office_Eng", "Service_Eng"])
        .agg(OOT=("Disposed_Out", "sum"), Received=("Received", "sum"))
        .reset_index()
    )
    pivot_oot = (
        cell.pivot(index="Office_Eng", columns="Service_Eng", values="OOT")
        .reindex(index=top5_offices, columns=top5_services)
        .fillna(0)
    )
    pivot_rec = (
        cell.pivot(index="Office_Eng", columns="Service_Eng", values="Received")
        .reindex(index=top5_offices, columns=top5_services)
        .fillna(0)
    )

    off_top = (
        off_agg[off_agg["Office_Eng"].isin(top5_offices)]
        .set_index("Office_Eng")
        .reindex(top5_offices)
    )
    srv_top = (
        srv_agg[srv_agg["Service_Eng"].isin(top5_services)]
        .set_index("Service_Eng")
        .reindex(top5_services)
    )

    matrix_oot = int(pivot_oot.values.sum())
    matrix_pct = round(matrix_oot / state_total_oot * 100, 1)

    return dict(
        pivot_oot=pivot_oot, pivot_rec=pivot_rec,
        off_top=off_top, srv_top=srv_top,
        state_total_oot=state_total_oot,
        matrix_oot=matrix_oot, matrix_pct=matrix_pct,
        fy_label=fy_label,
        top5_offices=top5_offices, top5_services=top5_services,
    )


# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
layout = dbc.Container([

    # Header
    html.Div([
        html.H2("🔥 Out-of-Time Drilldown", style={"color": "white", "margin": 0, "fontWeight": "600"}),
        html.P(
            "Top 5 Offices & Top 5 Services driving Out-of-Time cases — and their intersection.",
            style={"color": "#ffcccc", "margin": "6px 0 0 0", "fontSize": "0.95rem"}
        ),
    ], style={
        "background": "linear-gradient(135deg, #8b0000 0%, #e63946 100%)",
        "padding": "22px 30px", "borderRadius": "12px", "marginBottom": "24px",
        "boxShadow": "0 4px 12px rgba(0,0,0,0.15)"
    }),

    # Status bar
    html.Div(id="oot-status-bar", className="mb-4"),

    # ── Bar Charts Row ───────────────────────────────────────────────────────
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardHeader(
                    html.H6("🏢 Top 5 Offices by OOT Volume", className="mb-0 fw-bold text-dark"),
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
                    html.H6("⚙️ Top 5 Services by OOT Volume", className="mb-0 fw-bold text-dark"),
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

    # ── 5×5 Annotated Heatmap ────────────────────────────────────────────────
    dbc.Card([
        dbc.CardHeader(
            html.Div([
                html.H5("🔥 5 × 5 Intersection Heatmap", className="mb-0 fw-bold text-dark d-inline"),
                html.Span(
                    " — Each cell shows OOT count (bold) + OOT rate %",
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
        text=[f"{v:,}  ({r:.1f}%)" for v, r in zip(oot_vals, rates)],
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


def _heatmap_fig(pivot_oot, pivot_rec):
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
            rate = round(cell_oot / cell_rec * 100, 1) if cell_rec > 0 else 0.0
            font_color = "white" if (cell_oot / max_val) > 0.45 else "#222222"
            annotations.append(dict(
                x=x_labels[j],
                y=y_labels[i],
                text=f"<b>{cell_oot:,}</b><br>{rate:.1f}%",
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
                f"<b>OOT Rate:</b> {rate:.1f}%"
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
# CALLBACK
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output("oot-office-bar", "figure"),
    Output("oot-service-bar", "figure"),
    Output("oot-heatmap", "figure"),
    Output("oot-status-bar", "children"),
    Input("fy-store", "data"),
)
def update_oot(fy):
    d = _build_data(fy)

    # Sort ascending so worst appears at top of horizontal bar chart
    off_sorted = d["off_top"].sort_values("OOT", ascending=True)
    fig_off = _bar_fig(
        off_sorted.index.tolist(),
        off_sorted["OOT"].tolist(),
        off_sorted["OOT_Rate"].tolist(),
        "#e63946",
    )

    srv_sorted = d["srv_top"].sort_values("OOT", ascending=True)
    fig_srv = _bar_fig(
        srv_sorted.index.tolist(),
        srv_sorted["OOT"].tolist(),
        srv_sorted["OOT_Rate"].tolist(),
        "#bd0026",
    )

    fig_heat = _heatmap_fig(d["pivot_oot"], d["pivot_rec"])

    # Find the single worst cell for the claim
    flat = [
        (off, srv, int(d["pivot_oot"].loc[off, srv]))
        for off in d["top5_offices"]
        for srv in d["top5_services"]
    ]
    flat.sort(key=lambda x: x[2], reverse=True)
    top1_off, top1_srv, top1_oot = flat[0]
    top1_pct = round(top1_oot / d["state_total_oot"] * 100, 1)

    status = dbc.Alert([
        html.B(f"📅 {d['fy_label']}"),
        "  |  ",
        html.B("State Total OOT: "), f"{d['state_total_oot']:,}",
        "  |  ",
        html.B("Top 5×5 Matrix: "),
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

    return fig_off, fig_srv, fig_heat, status