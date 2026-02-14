import math
import os
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, Input, Output, dash_table, dcc, html

from Anode_Trending import (
    compute_retrofit_requirements,
    derive_projected_year,
    load_remaining_life,
    parse_inputs_sheet,
)

ANODE_PATH = os.getenv(
    "ANODE_XLSX_PATH",
    "/Users/syahnaz.omar/Library/CloudStorage/OneDrive-Personal/Desktop/sesco/Working File/Anode Trending.xlsx",
)
SKG_PATH = os.getenv(
    "SKG_XLSX_PATH",
    "/Users/syahnaz.omar/Library/CloudStorage/OneDrive-Personal/Desktop/sesco/Working File/SKG Database/SKG Asset Dimension 2025.xlsx",
)
TARGET_YEAR = 2050
ELEVATION_THRESHOLD = -30.0

ACCENT = "#00A19A"
FADED = "#D6DCE5"
BG = "#F5F7FA"
CARD_BG = "#FFFFFF"
TEXT = "#1F2933"


def metric_card(title, value_id):
    return html.Div(
        style={
            "backgroundColor": CARD_BG,
            "borderRadius": "10px",
            "boxShadow": "0 2px 8px rgba(0,0,0,0.06)",
            "padding": "12px 14px",
            "marginBottom": "10px",
        },
        children=[
            html.Div(title, style={"fontSize": "12px", "color": "#6B7280", "marginBottom": "6px"}),
            html.Div(id=value_id, style={"fontSize": "22px", "fontWeight": "600", "color": TEXT}),
        ],
    )

DEPLETION_DATE_COLS = [
    "Depletion \nMay 2001",
    "Depletion \nFeb 2014",
    "Depletion \nApril 2016",
    "Depletion \nJuly 2024",
]


def clean_depletion_value(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.lower() in {"nan", "n/a", "na", "none", "-", "null"}:
        return None
    return text


def latest_depletion_stage(row):
    # latest date is the right-most depletion column
    for col in reversed(DEPLETION_DATE_COLS):
        value = clean_depletion_value(row.get(col))
        if value is not None:
            return value
    return None


def anode_sort_key(anode_no: str):
    text = str(anode_no)
    m = re.search(r"BAN(\d+)", text.upper())
    if m:
        return int(m.group(1))
    return 10**9


def rectangle_perimeter_points(n: int, width: float = 120.0, height: float = 80.0):
    if n <= 0:
        return []
    if n == 1:
        return [(0.0, height / 2.0)]

    perimeter = 2.0 * (width + height)
    points = []
    step = perimeter / n
    for i in range(n):
        d = i * step
        if d <= width:
            x = -width / 2.0 + d
            y = height / 2.0
        elif d <= width + height:
            x = width / 2.0
            y = height / 2.0 - (d - width)
        elif d <= 2 * width + height:
            x = width / 2.0 - (d - width - height)
            y = -height / 2.0
        else:
            x = -width / 2.0
            y = -height / 2.0 + (d - 2 * width - height)
        points.append((x, y))
    return points


def build_rectangular_anode_layout(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.assign(X_Rect=pd.Series(dtype=float), Y_Rect=pd.Series(dtype=float), Z_Elevation=pd.Series(dtype=float))
    out = []
    for elev in sorted(df["Elevation (m)"].dropna().unique().tolist(), reverse=True):
        layer = df[df["Elevation (m)"] == elev].copy()
        layer["sort_key"] = layer["Anode No"].apply(anode_sort_key)
        layer = layer.sort_values(["sort_key", "Anode No"])
        points = rectangle_perimeter_points(len(layer))
        for (_, row), (x, y) in zip(layer.iterrows(), points):
            rec = row.to_dict()
            rec["X_Rect"] = round(float(x), 2)
            rec["Y_Rect"] = round(float(y), 2)
            rec["Z_Elevation"] = round(float(elev), 2)
            out.append(rec)
    return pd.DataFrame(out)


def build_performance_curve(anodes: pd.DataFrame, design_year: float, projected_year: float) -> pd.DataFrame:
    start = int(math.floor(design_year))
    end = int(math.ceil(projected_year))
    years = list(range(start, end + 1))
    total = len(anodes)
    recs = []
    for y in years:
        active = int((anodes["Anode Life"] >= y).sum())
        recs.append({"Year": y, "Active_Anodes": active, "Depleted_Anodes": total - active})
    return pd.DataFrame(recs)


def add_theme(fig, height=320, xangle=0):
    fig.update_layout(
        margin=dict(l=40, r=20, t=60, b=40),
        paper_bgcolor=CARD_BG,
        plot_bgcolor=CARD_BG,
        template="plotly_white",
        font=dict(family="Helvetica, Arial, sans-serif", color=TEXT),
        height=height,
    )
    fig.update_xaxes(tickangle=xangle)
    return fig


def load_platform_meta() -> dict:
    try:
        psc = pd.read_excel(SKG_PATH, sheet_name="PSC", header=0)
        row = psc[psc["PLATFORM"].astype(str).str.upper() == "M1PQ-A"]
        if row.empty:
            return {}
        one = row.iloc[0]
        return {
            "cluster": str(one.get("CLUSTER", "-")).strip(),
            "type": str(one.get("TYPE", "-")).strip(),
            "status": str(one.get("OPERATING STATUS", "-")).strip(),
        }
    except Exception:
        return {}


def load_data():
    info = parse_inputs_sheet(ANODE_PATH)
    anodes = load_remaining_life(ANODE_PATH)
    projected_year = derive_projected_year(ANODE_PATH, float(anodes["Anode Life"].max()))

    m1 = pd.read_excel(ANODE_PATH, sheet_name="M1PQ-A")
    m1 = m1[[
        "Anode No",
        "Anode Category",
        "EL. Zone (m)",
        "Depletion Category",
        "Depletion \nMay 2001",
        "Depletion \nFeb 2014",
        "Depletion \nApril 2016",
        "Depletion \nJuly 2024",
    ]].copy()
    m1["Anode No"] = m1["Anode No"].astype(str).str.strip()
    m1 = m1[(m1["Anode No"].str.lower() != "nan") & (m1["Anode No"] != "")]
    m1["Anode Category"] = m1["Anode Category"].astype(str).str.strip()
    m1["Latest Depletion"] = m1.apply(latest_depletion_stage, axis=1)
    m1["Depletion Category"] = pd.to_numeric(m1["Depletion Category"], errors="coerce")
    m1["EL. Zone (m)"] = pd.to_numeric(m1["EL. Zone (m)"], errors="coerce")

    anodes = anodes.merge(
        m1[["Anode No", "Latest Depletion", "Depletion Category"]],
        on="Anode No",
        how="left",
    )

    m1_model = m1[["Anode No", "Anode Category", "EL. Zone (m)", "Latest Depletion"]].copy()
    m1_model = m1_model.rename(columns={"EL. Zone (m)": "Elevation (m)"})
    m1_model = m1_model[m1_model["Anode No"].str.upper().str.contains(r"BAN\d+", regex=True, na=False)]
    m1_model["Elevation (m)"] = pd.to_numeric(m1_model["Elevation (m)"], errors="coerce")
    m1_model = m1_model.dropna(subset=["Elevation (m)"])
    m1_model["Latest Depletion"] = m1_model["Latest Depletion"].apply(clean_depletion_value)
    m1_model = m1_model.dropna(subset=["Latest Depletion"])
    m1_model["Elevation Band"] = pd.cut(
        m1_model["Elevation (m)"],
        bins=[-1000, ELEVATION_THRESHOLD, 1000],
        labels=[f"<= {ELEVATION_THRESHOLD:.0f} m", f"> {ELEVATION_THRESHOLD:.0f} m"],
    ).astype(str)
    m1_model = build_rectangular_anode_layout(m1_model)

    anodes["Elevation Band"] = pd.cut(
        anodes["Elevation (m)"],
        bins=[-1000, ELEVATION_THRESHOLD, 1000],
        labels=[f"<= {ELEVATION_THRESHOLD:.0f} m", f"> {ELEVATION_THRESHOLD:.0f} m"],
    ).astype(str)

    perf_all = build_performance_curve(anodes, info["design_year"], projected_year)
    depleted_all = anodes[anodes["Anode Life"] <= projected_year].copy()

    allocation, summary = compute_retrofit_requirements(
        anodes=anodes,
        info=info,
        projected_year=projected_year,
        target_year=TARGET_YEAR,
        elevation_threshold=ELEVATION_THRESHOLD,
    )

    rl = pd.read_excel(ANODE_PATH, sheet_name="Remaining Life")
    total_original = int(m1[m1["Anode Category"].str.lower() == "original"]["Anode No"].nunique())
    total_current = int(m1["Anode No"].nunique())
    depleted_val = pd.to_numeric(rl.get("Depleted"), errors="coerce")
    replaced_val = pd.to_numeric(rl.get("Replaced"), errors="coerce")
    depleted = int(depleted_val.max()) if depleted_val is not None and depleted_val.notna().any() else 0
    replaced = int(replaced_val.max()) if replaced_val is not None and replaced_val.notna().any() else 0

    return {
        "info": info,
        "anodes": anodes,
        "perf_all": perf_all,
        "depleted_all": depleted_all,
        "projected_year": projected_year,
        "allocation": allocation,
        "summary": summary,
        "meta": load_platform_meta(),
        "model_df": m1_model,
        "anode_detail": {
            "total_original": total_original,
            "total_current": total_current,
            "depleted": depleted,
            "replaced": replaced,
        },
    }


def apply_filters(df: pd.DataFrame, category: str, elevation_band: str) -> pd.DataFrame:
    out = df.copy()
    if category != "All":
        out = out[out["Anode Category"] == category]
    if elevation_band != "All":
        out = out[out["Elevation Band"] == elevation_band]
    return out


def apply_stage_filter(df: pd.DataFrame, selected_stages, stage_col: str = "Latest Depletion") -> pd.DataFrame:
    if df.empty:
        return df
    selected = selected_stages or ["__ALL_STAGES__"]
    if "__ALL_STAGES__" in selected:
        return df
    clean = (
        df[stage_col]
        .astype(str)
        .str.strip()
        .replace({"nan": None, "N/A": None, "n/a": None, "NA": None, "na": None, "-": None, "": None})
    )
    return df[clean.isin(set(selected))]


data = load_data()
anodes_df = data["anodes"]
projected_year = data["projected_year"]
info = data["info"]
model_df = data["model_df"].copy()
stage_values = sorted(
    set(
        pd.concat(
            [
                anodes_df["Latest Depletion"],
                data["model_df"]["Latest Depletion"],
            ],
            ignore_index=True,
        )
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: ~s.str.lower().isin({"n/a", "na", "nan", "none", "-", "null", ""})]
        .tolist()
    )
)
stage_options = [{"label": "All stages", "value": "__ALL_STAGES__"}] + [{"label": s, "value": s} for s in stage_values]

meta_txt = ""
if data["meta"]:
    meta_txt = (
        f"Cluster: {data['meta'].get('cluster', '-')} | "
        f"Type: {data['meta'].get('type', '-')} | "
        f"Status: {data['meta'].get('status', '-')}"
    )
anode_detail = data["anode_detail"]
anode_detail_txt = (
    f"Total Original Anodes: {anode_detail['total_original']:,} | "
    f"Total Current Anodes: {anode_detail['total_current']:,} | "
    f"Depleted: {anode_detail['depleted']:,} | "
    f"Replaced: {anode_detail['replaced']:,}"
)

app = Dash(__name__)

app.layout = html.Div(
    style={
        "fontFamily": "Helvetica, Arial, sans-serif",
        "backgroundColor": BG,
        "padding": "16px",
        "maxWidth": "1280px",
        "margin": "0 auto",
    },
    children=[
        html.Div(
            style={
                "backgroundColor": CARD_BG,
                "padding": "16px 20px",
                "borderRadius": "10px",
                "boxShadow": "0 2px 8px rgba(0,0,0,0.06)",
                "marginBottom": "16px",
            },
            children=[
                html.Div(
                    style={"display": "flex", "justifyContent": "space-between", "gap": "16px", "alignItems": "stretch"},
                    children=[
                        html.Div(
                            style={"flex": 1},
                            children=[
                                html.H2("Anode Trending Dashboard", style={"margin": 0, "color": TEXT}),
                                html.Div(
                                    f"Design Year: {round(info['design_year']):.0f} | Jul-24 Reference: {round(info['jul24_year']):.0f} | "
                                    f"Projected Life: {round(projected_year):.0f} | Target: {TARGET_YEAR}",
                                    style={"marginTop": "8px", "fontSize": "13px", "color": "#4B5563"},
                                ),
                                html.Div(meta_txt, style={"marginTop": "4px", "fontSize": "13px", "color": "#4B5563"}),
                                html.Div(anode_detail_txt, style={"marginTop": "4px", "fontSize": "13px", "color": "#4B5563"}),
                            ],
                        ),
                        html.Div(
                            style={
                                "width": "220px",
                                "backgroundColor": "#F8FAFC",
                                "border": "1px solid #E2E8F0",
                                "borderRadius": "10px",
                                "padding": "8px",
                                "textAlign": "center",
                            },
                            children=[
                                html.Img(
                                    src=app.get_asset_url("M1PQ-A.png"),
                                    style={"width": "100%", "height": "120px", "objectFit": "cover", "borderRadius": "8px"},
                                ),
                                html.Div("M1PQ-A", style={"marginTop": "6px", "fontWeight": "600", "color": TEXT}),
                            ],
                        ),
                    ],
                ),
                html.Div(
                    style={"display": "flex", "gap": "16px", "marginTop": "12px"},
                    children=[
                        html.Div(
                            style={"flex": 1},
                            children=[
                                html.Label("Anode Category"),
                                dcc.Dropdown(
                                    id="anode-category",
                                    options=[
                                        {"label": "All", "value": "All"},
                                        {"label": "Original", "value": "Original"},
                                        {"label": "Retrofit", "value": "Retrofit"},
                                    ],
                                    value="All",
                                    clearable=False,
                                ),
                            ],
                        ),
                        html.Div(
                            style={"flex": 1},
                            children=[
                                html.Label("Elevation Filter"),
                                dcc.RadioItems(
                                    id="elevation-band",
                                    options=[
                                        {"label": "All", "value": "All"},
                                        {"label": f"> {ELEVATION_THRESHOLD:.0f} m", "value": f"> {ELEVATION_THRESHOLD:.0f} m"},
                                        {"label": f"<= {ELEVATION_THRESHOLD:.0f} m", "value": f"<= {ELEVATION_THRESHOLD:.0f} m"},
                                    ],
                                    value="All",
                                    labelStyle={
                                        "display": "inline-block",
                                        "marginRight": "8px",
                                        "padding": "6px 10px",
                                        "border": "1px solid #E2E8F0",
                                        "borderRadius": "999px",
                                        "backgroundColor": "#FFFFFF",
                                        "cursor": "pointer",
                                        "fontSize": "13px",
                                    },
                                    inputStyle={"marginRight": "6px"},
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),
        html.Div(
            style={"display": "flex", "gap": "12px"},
            children=[
                html.Div(
                    style={"width": "260px"},
                    children=[
                        metric_card("Total Anodes (Filtered)", "metric-total"),
                        metric_card("Depleted by Projected Year", "metric-depleted"),
                        metric_card("Projected Life Year", "metric-projected"),
                        metric_card("Min Retrofit Needed to 2050", "metric-retrofit"),
                    ],
                ),
                html.Div(
                    style={"flex": 1},
                    children=[
                        html.Div(
                            style={"display": "grid", "gridTemplateColumns": "1fr 1fr", "gap": "12px"},
                            children=[
                                dcc.Graph(id="fig-performance", style={"backgroundColor": CARD_BG, "borderRadius": "10px", "height": "320px"}),
                                dcc.Graph(id="fig-depletion-year", style={"backgroundColor": CARD_BG, "borderRadius": "10px", "height": "320px"}),
                                dcc.Graph(id="fig-elevation", style={"backgroundColor": CARD_BG, "borderRadius": "10px", "height": "320px"}),
                                dcc.Graph(id="fig-stage", style={"backgroundColor": CARD_BG, "borderRadius": "10px", "height": "320px"}),
                            ],
                        ),
                        html.Div(
                            style={
                                "marginTop": "16px",
                                "backgroundColor": CARD_BG,
                                "borderRadius": "10px",
                                "boxShadow": "0 2px 8px rgba(0,0,0,0.06)",
                                "padding": "8px",
                            },
                            children=[
                                html.Div(
                                    style={"padding": "4px 8px 8px 8px"},
                                    children=[
                                        html.Label("Depletion Stage"),
                                        dcc.Dropdown(
                                            id="stage-filter",
                                            options=stage_options,
                                            value=["__ALL_STAGES__"],
                                            multi=True,
                                            placeholder="Select depletion stage",
                                        ),
                                    ],
                                ),
                                dcc.Graph(id="fig-anode-3d", style={"height": "520px"}),
                            ],
                        ),
                        html.Div(
                            style={
                                "marginTop": "16px",
                                "backgroundColor": CARD_BG,
                                "borderRadius": "10px",
                                "boxShadow": "0 2px 8px rgba(0,0,0,0.06)",
                                "padding": "8px",
                            },
                            children=[dcc.Graph(id="fig-retrofit", style={"height": "340px"})],
                        ),
                        html.Div(
                            style={
                                "marginTop": "16px",
                                "backgroundColor": CARD_BG,
                                "borderRadius": "10px",
                                "boxShadow": "0 2px 8px rgba(0,0,0,0.06)",
                                "padding": "10px",
                            },
                            children=[
                                html.H4("Anodes Depleting Within Projection", style={"margin": "0 0 8px 0", "color": TEXT}),
                                dash_table.DataTable(
                                    id="depleted-table",
                                    columns=[
                                        {"name": "Anode No", "id": "Anode No"},
                                        {"name": "Category", "id": "Anode Category"},
                                        {"name": "Elevation (m)", "id": "Elevation (m)"},
                                        {"name": "Anode Life", "id": "Anode Life"},
                                        {"name": "Latest Depletion", "id": "Latest Depletion"},
                                    ],
                                    page_size=10,
                                    style_table={"overflowX": "auto"},
                                    style_header={"backgroundColor": "#EEF2F7", "fontWeight": "bold", "color": TEXT},
                                    style_cell={"padding": "8px", "fontSize": "12px", "color": TEXT},
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),
    ],
)


@app.callback(
    Output("fig-performance", "figure"),
    Output("fig-depletion-year", "figure"),
    Output("fig-elevation", "figure"),
    Output("fig-stage", "figure"),
    Output("fig-anode-3d", "figure"),
    Output("fig-retrofit", "figure"),
    Output("depleted-table", "data"),
    Output("metric-total", "children"),
    Output("metric-depleted", "children"),
    Output("metric-projected", "children"),
    Output("metric-retrofit", "children"),
    Input("anode-category", "value"),
    Input("elevation-band", "value"),
    Input("stage-filter", "value"),
)
def update_dashboard(category, elevation_band, selected_stages):
    anodes_base = apply_filters(anodes_df, category, elevation_band)
    model_base = apply_filters(data["model_df"], category, elevation_band)
    filtered = apply_stage_filter(anodes_base, selected_stages)
    model_filtered = apply_stage_filter(model_base, selected_stages)

    # Retrofit allocation is stage-scoped via elevations present in filtered anodes.
    if filtered.empty:
        alloc_filtered = data["allocation"].iloc[0:0].copy()
    else:
        selected_elev = set(pd.to_numeric(filtered["Elevation (m)"], errors="coerce").round(1).dropna().tolist())
        alloc_filtered = data["allocation"][
            pd.to_numeric(data["allocation"]["Elevation (m)"], errors="coerce").round(1).isin(selected_elev)
        ].copy()

    perf_all = build_performance_curve(anodes_base, info["design_year"], projected_year)
    perf_filtered = build_performance_curve(filtered, info["design_year"], projected_year)
    fig_perf = go.Figure()
    fig_perf.add_trace(
        go.Scatter(
            x=perf_all["Year"],
            y=perf_all["Active_Anodes"],
            mode="lines",
            line=dict(color=FADED, width=2),
            name="Active (All)",
        )
    )
    fig_perf.add_trace(
        go.Scatter(
            x=perf_filtered["Year"],
            y=perf_filtered["Active_Anodes"],
            mode="lines+markers",
            line=dict(color=ACCENT, width=3),
            name="Active (Filtered)",
        )
    )
    fig_perf.add_vline(x=projected_year, line_width=1.5, line_dash="dash", line_color="#EF4444")
    fig_perf.update_layout(title="Performance Trend: Design to Projected Life", xaxis_title="Year", yaxis_title="Number of Anodes")
    fig_perf = add_theme(fig_perf)

    depleted = filtered[filtered["Anode Life"] <= projected_year].copy()
    dep_counts = (
        depleted.assign(Depletion_Year=depleted["Anode Life"].apply(math.floor).astype(int))
        .groupby("Depletion_Year", as_index=False)
        .size()
        .rename(columns={"size": "Count"})
    )
    fig_dep = go.Figure()
    if dep_counts.empty:
        fig_dep.add_annotation(text="No anodes in filter", showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper")
    else:
        fig_dep.add_trace(go.Bar(x=dep_counts["Depletion_Year"], y=dep_counts["Count"], marker_color="#205295"))
    fig_dep.add_vline(x=projected_year, line_width=1.5, line_dash="dash", line_color="#EF4444")
    fig_dep.update_layout(title="Depletion Count by Year (Filtered)", xaxis_title="Depletion Year", yaxis_title="Number of Anodes")
    fig_dep = add_theme(fig_dep)

    elev_counts = (
        filtered.assign(ElevationRounded=filtered["Elevation (m)"].round(1))
        .groupby("ElevationRounded", as_index=False)
        .size()
        .rename(columns={"size": "Count"})
    )
    elev_counts = elev_counts.sort_values("ElevationRounded", ascending=False)
    fig_elev = go.Figure()
    if elev_counts.empty:
        fig_elev.add_annotation(text="No data in filter", showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper")
    else:
        fig_elev.add_trace(go.Bar(x=elev_counts["ElevationRounded"].astype(str), y=elev_counts["Count"], marker_color="#3FA796"))
    fig_elev.update_layout(title="Anode Distribution by Elevation", xaxis_title="Elevation (m)", yaxis_title="Number of Anodes")
    fig_elev = add_theme(fig_elev, xangle=45)

    stage_counts = (
        filtered["Latest Depletion"]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: ~s.str.lower().isin({"n/a", "na", "nan", "none", "-", "null", ""})]
        .value_counts()
        .reset_index()
    )
    stage_counts.columns = ["Stage", "Count"]
    fig_stage = go.Figure()
    if stage_counts.empty:
        fig_stage.add_annotation(text="No stage data", showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper")
    else:
        fig_stage.add_trace(
            go.Pie(
                labels=stage_counts["Stage"],
                values=stage_counts["Count"],
                hole=0.4,
                marker=dict(colors=px.colors.qualitative.Set2),
                textinfo="percent+label",
            )
        )
    fig_stage.update_layout(title="Anode Stage Mix (Latest Depletion Date)", legend_title_text="Depletion Stage")
    fig_stage = add_theme(fig_stage)

    stage_order = ["1-25%", "26-50%", "51-75%", "76-100%", "Depleted 2014"]
    stage_color_map = {
        "1-25%": "#0EA5E9",
        "26-50%": "#10B981",
        "51-75%": "#F59E0B",
        "76-100%": "#EF4444",
        "Depleted 2014": "#7C3AED",
    }
    if model_filtered.empty:
        fig_3d = go.Figure()
        fig_3d.add_annotation(text="No anode points in current filter", showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper")
    else:
        fig_3d = px.scatter_3d(
            model_filtered,
            x="X_Rect",
            y="Y_Rect",
            z="Z_Elevation",
            color="Latest Depletion",
            symbol="Anode Category",
            hover_name="Anode No",
            hover_data={"X_Rect": False, "Y_Rect": False, "Z_Elevation": True, "Latest Depletion": True, "Anode Category": True},
            category_orders={"Latest Depletion": stage_order},
            color_discrete_map=stage_color_map,
        )
    fig_3d.update_layout(
        title="Anode Layout by Elevation and Depletion Stage",
        legend_title_text="Stage / Category",
        scene=dict(
            xaxis_title="X Axis",
            yaxis_title="Y Axis",
            zaxis_title="Elevation (m)",
        ),
        margin=dict(l=10, r=10, t=60, b=10),
        paper_bgcolor=CARD_BG,
        plot_bgcolor=CARD_BG,
        template="plotly_white",
        font=dict(family="Helvetica, Arial, sans-serif", color=TEXT),
    )

    alloc = alloc_filtered.sort_values("Elevation (m)", ascending=False)
    fig_retro = go.Figure()
    if alloc.empty:
        fig_retro.add_annotation(text="No retrofit points in current stage filter", showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper")
    else:
        fig_retro.add_trace(
            go.Bar(
                x=alloc["Elevation (m)"].round(1).astype(str),
                y=alloc["Retrofit_Qty_Required"],
                marker_color="#0EA5E9",
                text=alloc["Retrofit_Qty_Required"],
                textposition="outside",
            )
        )
    fig_retro.update_layout(
        title=f"Minimum Retrofit Allocation Above {ELEVATION_THRESHOLD:.0f} m (Target {TARGET_YEAR})",
        xaxis_title="Elevation (m)",
        yaxis_title="Required Retrofit Quantity",
    )
    fig_retro = add_theme(fig_retro, height=340, xangle=45)

    table = depleted[["Anode No", "Anode Category", "Elevation (m)", "Anode Life", "Latest Depletion"]].copy()
    table = table.sort_values(["Anode Life", "Elevation (m)", "Anode No"]).head(120)
    table["Elevation (m)"] = table["Elevation (m)"].round(1)
    table["Anode Life"] = table["Anode Life"].round(0).astype(int)

    metric_total = f"{len(filtered):,}"
    metric_depleted = f"{len(depleted):,}"
    metric_projected = f"{round(projected_year):.0f}"
    metric_retrofit = f"{int(alloc['Retrofit_Qty_Required'].sum()):,}"

    return (
        fig_perf,
        fig_dep,
        fig_elev,
        fig_stage,
        fig_3d,
        fig_retro,
        table.to_dict("records"),
        metric_total,
        metric_depleted,
        metric_projected,
        metric_retrofit,
    )


if __name__ == "__main__":
    app.run(debug=False)
