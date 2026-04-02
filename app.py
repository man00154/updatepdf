import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import openpyxl

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sify DC – Capacity Intelligence",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# THEME CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
NAVY  = "#0a0e1a"
DARK1 = "#0f1628"
DARK2 = "#141c35"
CARD  = "#1a2340"
BORD  = "#2a3a6a"
BLUE  = "#2a5298"
LBLUE = "#3d72d9"
CYAN  = "#00d4ff"
TEXT  = "#c9d8f0"
MUTED = "#7a92c0"
WHITE = "#ffffff"
GREEN = "#00c986"
AMBER = "#f5a623"
RED   = "#ff4d6d"

st.markdown(f"""
<style>
html,body,[class*="css"]{{background:{NAVY};color:{TEXT};font-family:'Segoe UI',sans-serif}}
.stApp{{background:{NAVY}}}
section[data-testid="stSidebar"]{{background:{DARK1};border-right:1px solid {BORD}}}
section[data-testid="stSidebar"] *{{color:{TEXT}!important}}
.stTabs [data-baseweb="tab-list"]{{background:{DARK1};border-radius:12px;padding:6px;gap:4px;border:1px solid {BORD}}}
.stTabs [data-baseweb="tab"]{{background:transparent;color:{MUTED};border-radius:8px;padding:8px 18px;font-weight:600;font-size:.88rem;border:none}}
.stTabs [aria-selected="true"]{{background:{BLUE};color:{WHITE}!important;box-shadow:0 2px 12px rgba(42,82,152,.5)}}
.stTabs [data-baseweb="tab-panel"]{{background:transparent;padding-top:1.2rem}}
.stSelectbox>div>div,.stMultiSelect>div>div{{background:{DARK2};border:1px solid {BORD};border-radius:8px;color:{TEXT}}}
.stTextInput>div>div{{background:{DARK2};border:1px solid {BORD};border-radius:8px}}
.stTextInput input{{color:{TEXT}}}
.stButton>button{{background:linear-gradient(135deg,{BLUE},{LBLUE});color:{WHITE};border:none;border-radius:8px;padding:8px 22px;font-weight:700;font-size:.9rem;box-shadow:0 4px 14px rgba(42,82,152,.4);transition:all .2s}}
.stButton>button:hover{{transform:translateY(-2px);box-shadow:0 6px 20px rgba(42,82,152,.6)}}
.stDataFrame{{border:1px solid {BORD};border-radius:10px;overflow:hidden}}
.stDataFrame table{{background:{DARK2};color:{TEXT}}}
.stDataFrame th{{background:{BLUE};color:{WHITE};padding:10px}}
.stDataFrame td{{border-bottom:1px solid {BORD};padding:8px 12px}}
[data-testid="metric-container"]{{background:{CARD};border:1px solid {BORD};border-radius:12px;padding:16px 20px;box-shadow:0 4px 16px rgba(0,0,0,.3)}}
[data-testid="metric-container"] label{{color:{MUTED};font-size:.82rem;font-weight:600;letter-spacing:.04em;text-transform:uppercase}}
[data-testid="metric-container"] [data-testid="stMetricValue"]{{color:{CYAN};font-size:1.7rem;font-weight:800}}
.kpi-card{{background:{CARD};border:1px solid {BORD};border-radius:14px;padding:22px 24px;box-shadow:0 4px 20px rgba(0,0,0,.35);text-align:center;height:100%}}
.kpi-val{{font-size:2rem;font-weight:900;line-height:1.1}}
.kpi-label{{font-size:.78rem;color:{MUTED};font-weight:700;text-transform:uppercase;letter-spacing:.06em;margin-top:6px}}
.kpi-sub{{font-size:.76rem;color:{GREEN};margin-top:4px}}
.section-title{{font-size:1.2rem;font-weight:800;color:{WHITE};border-left:4px solid {CYAN};padding-left:14px;margin:24px 0 16px;letter-spacing:.02em}}
.badge{{display:inline-block;background:{BLUE};color:{WHITE};border-radius:20px;padding:2px 12px;font-size:.76rem;font-weight:700;margin:2px}}
.result-box{{background:{DARK2};border:1px solid {BORD};border-radius:10px;padding:18px 22px;margin:10px 0;font-size:1rem;color:{TEXT}}}
.result-big{{font-size:2.6rem;font-weight:900;background:linear-gradient(90deg,{CYAN},{LBLUE});-webkit-background-clip:text;-webkit-text-fill-color:transparent}}
.hero{{background:linear-gradient(135deg,{DARK2} 0%,{DARK1} 50%,{NAVY} 100%);border:1px solid {BORD};border-radius:16px;padding:28px 36px;margin-bottom:24px;box-shadow:0 8px 32px rgba(0,0,0,.4)}}
.hero h1{{font-size:2rem;font-weight:900;color:{WHITE};margin:0;letter-spacing:.02em}}
.hero p{{color:{MUTED};margin:6px 0 0;font-size:.95rem}}
::-webkit-scrollbar{{width:6px;height:6px}}
::-webkit-scrollbar-track{{background:{DARK1}}}
::-webkit-scrollbar-thumb{{background:{BORD};border-radius:3px}}
::-webkit-scrollbar-thumb:hover{{background:{BLUE}}}
header[data-testid="stHeader"]{{display:none!important;visibility:hidden!important;height:0!important}}
#MainMenu{{display:none!important}}
footer{{display:none!important}}
.stDeployButton{{display:none!important}}
[data-testid="stToolbar"]{{display:none!important}}
.viewerBadge_container__1QSob{{display:none!important}}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL LOADING  –  supports both .xlsx (openpyxl) and .xls (xlrd)
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_DIR = Path(__file__).parent / "excel_files"

SECTION_MARKERS = {
    "Billing Model", "Space", "Power Capacity", "Power Usage",
    "Seating Space", "Revenue", "DEMARC", "RHS", "SHS",
    "ONSITE TAPE ROTATION", "OFFSITE TAPE ROTATION",
    "SAFE VAULT", "STORE SPACE",
}


def _is_section_row(vals) -> bool:
    non = [v for v in vals if v and str(v).strip() not in ("", "None")]
    if not non:
        return False
    hits = sum(1 for v in non if str(v).strip() in SECTION_MARKERS)
    return hits / len(non) > 0.25


def _build_cols(g_row, c_row):
    cur_g = ""
    cols = []
    for g, c in zip(g_row, c_row):
        if g and g not in ("None", ""):
            cur_g = g
        label = c if c and c not in ("None", "") else ""
        if label:
            cols.append(f"{cur_g} | {label}" if cur_g else label)
        else:
            cols.append(cur_g or "_col")
    return cols


def _clean_df(df: pd.DataFrame) -> "pd.DataFrame | None":
    df = df.dropna(axis=1, how="all")
    df = df[df.apply(
        lambda r: any(str(v).strip() not in ("", "None", "nan") for v in r), axis=1
    )]
    df = df[~df.apply(
        lambda r: r.astype(str).str.contains(
            r"#DIV|#REF|#N/A|#VALUE", regex=True, na=False
        ).any(), axis=1
    )]
    for col in df.columns:
        conv = pd.to_numeric(df[col], errors="coerce")
        if conv.notna().sum() / max(len(df), 1) > 0.45:
            df[col] = conv
    return df if len(df) >= 1 else None


def _detect_header(raw_rows):
    """Return (hdr_start, g_row, c_row) from first 6 rows."""
    def rs(r):
        return [str(v).strip() if v is not None else "" for v in r]

    r1 = rs(raw_rows[0])
    r2 = rs(raw_rows[1]) if len(raw_rows) > 1 else []
    r3 = rs(raw_rows[2]) if len(raw_rows) > 2 else []

    if _is_section_row(r1) and _is_section_row(r2) and r3:
        return 4, r2, r3
    elif _is_section_row(r1) and not _is_section_row(r2):
        return 3, r1, r2
    elif r1:
        return 2, [""] * len(r1), r1
    return None, None, None


def _load_ws(ws) -> "pd.DataFrame | None":
    """Load an openpyxl worksheet."""
    if ws.max_row < 2 or ws.max_column < 2:
        return None
    raw = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 6), values_only=True))
    hdr_start, g_row, c_row = _detect_header(raw)
    if hdr_start is None:
        return None
    cols = _build_cols(g_row, c_row)
    data = []
    for row in ws.iter_rows(min_row=hdr_start, values_only=True):
        vals = list(row)[: len(cols)]
        if any(v is not None and str(v).strip() not in ("", "None") for v in vals):
            data.append(vals)
    if not data:
        return None
    return _clean_df(pd.DataFrame(data, columns=cols))


def _load_xls_ws(sheet) -> "pd.DataFrame | None":
    """Load a xlrd worksheet (legacy .xls)."""
    try:
        import xlrd  # noqa: F401
        nrows, ncols = sheet.nrows, sheet.ncols
        if nrows < 2 or ncols < 2:
            return None

        def cv(r, c):
            try:
                v = sheet.cell_value(r, c)
                return str(v).strip() if v is not None and str(v).strip() else ""
            except Exception:
                return ""

        raw = [[cv(ri, ci) for ci in range(ncols)] for ri in range(min(nrows, 6))]
        hdr_start, g_row, c_row = _detect_header(raw)
        if hdr_start is None:
            return None
        cols = _build_cols(g_row, c_row)
        data = []
        for ri in range(hdr_start, nrows):
            vals = [cv(ri, ci) for ci in range(ncols)][: len(cols)]
            if any(v and v not in ("", "None") for v in vals):
                data.append(vals)
        if not data:
            return None
        return _clean_df(pd.DataFrame(data, columns=cols))
    except Exception:
        return None


def _label(stem: str) -> str:
    s = re.sub(r"Customer_and_Capacity_Tracker_", "", stem, flags=re.I)
    s = re.sub(r"_\d{10,}$", "", s)
    return s.replace("_", " ").strip()


@st.cache_data(show_spinner=False)
def load_all() -> dict:
    """Return  {location_label: {sheet_name: DataFrame}}  for every Excel file."""
    result: dict = {}

    # ── .xlsx ──────────────────────────────────────────────────────────────
    for fpath in sorted(EXCEL_DIR.glob("*.xlsx")):
        label = _label(fpath.stem)
        try:
            wb = openpyxl.load_workbook(str(fpath), data_only=True)
        except Exception:
            continue
        sheets = {}
        for sn in wb.sheetnames:
            try:
                df = _load_ws(wb[sn])
                if df is not None and len(df) >= 2:
                    sheets[sn] = df
            except Exception:
                pass
        wb.close()
        if sheets:
            result[label] = sheets

    # ── .xls ───────────────────────────────────────────────────────────────
    for fpath in sorted(EXCEL_DIR.glob("*.xls")):
        label = _label(fpath.stem)
        try:
            import xlrd
            wb = xlrd.open_workbook(str(fpath))
            sheets = {}
            for sn in wb.sheet_names():
                try:
                    df = _load_xls_ws(wb.sheet_by_name(sn))
                    if df is not None and len(df) >= 2:
                        sheets[sn] = df
                except Exception:
                    pass
            if sheets:
                result[label] = sheets
        except Exception:
            continue

    return result


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def combined_df(data: dict) -> pd.DataFrame:
    """Concatenate ALL sheets from ALL locations into one DataFrame."""
    frames = []
    for loc, sheets in data.items():
        for sn, df in sheets.items():
            tmp = df.copy()
            tmp.insert(0, "_Sheet", sn)
            tmp.insert(0, "_Location", loc)
            frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def num_cols(df: pd.DataFrame) -> list:
    return [c for c in df.columns
            if pd.api.types.is_numeric_dtype(df[c]) and not c.startswith("_")]


def txt_cols(df: pd.DataFrame) -> list:
    return [c for c in df.columns
            if not pd.api.types.is_numeric_dtype(df[c]) and not c.startswith("_")]


def find_col(df: pd.DataFrame, *pats) -> "str | None":
    for p in pats:
        for c in df.columns:
            if re.search(p, c, re.I):
                return c
    return None


def fmt(n):
    if n is None:
        return "–"
    try:
        n = float(n)
        if abs(n) >= 1_000_000:
            return f"{n / 1_000_000:.2f} M"
        if abs(n) >= 1_000:
            return f"{n / 1_000:.1f} K"
        return f"{n:,.2f}"
    except Exception:
        return str(n)


def kpi_html(value, label, sub="", color=CYAN):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return (
        f'<div class="kpi-card">'
        f'<div class="kpi-val" style="color:{color}">{value}</div>'
        f'<div class="kpi-label">{label}</div>{sub_html}</div>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# OPERATIONS ENGINE
# ─────────────────────────────────────────────────────────────────────────────
OPERATIONS = [
    "Sum", "Mean (Avg)", "Median", "Min", "Max",
    "Count", "Std Deviation", "Variance", "Range (Max-Min)",
    "Top N Values", "Bottom N Values",
    "Cumulative Sum", "Rank (Desc)", "% of Total",
]


def run_op(df: pd.DataFrame, col: str, op: str,
           group_by: str = None, top_n: int = 10):
    series = pd.to_numeric(df[col], errors="coerce").dropna()
    if series.empty:
        return None, "No numeric data."

    if group_by and group_by in df.columns:
        g = df[[group_by, col]].copy()
        g[col] = pd.to_numeric(g[col], errors="coerce")
        g = g.dropna(subset=[col])
        fn = {
            "Sum": "sum", "Mean (Avg)": "mean", "Median": "median",
            "Min": "min", "Max": "max", "Count": "count",
            "Std Deviation": "std", "Variance": "var",
        }.get(op, "sum")
        res = g.groupby(group_by)[col].agg(fn).reset_index()
        res[col] = res[col].round(3)
        return res, f"{op} of '{col}' by '{group_by}'"

    if op == "Sum":               v = series.sum()
    elif op == "Mean (Avg)":      v = series.mean()
    elif op == "Median":          v = series.median()
    elif op == "Min":             v = series.min()
    elif op == "Max":             v = series.max()
    elif op == "Count":           v = len(series)
    elif op == "Std Deviation":   v = series.std()
    elif op == "Variance":        v = series.var()
    elif op == "Range (Max-Min)": v = series.max() - series.min()
    elif op == "% of Total":      v = 100.0
    elif op == "Top N Values":
        res = df[[col]].copy()
        res[col] = pd.to_numeric(res[col], errors="coerce")
        return res.dropna().nlargest(top_n, col).reset_index(drop=True), f"Top {top_n} of '{col}'"
    elif op == "Bottom N Values":
        res = df[[col]].copy()
        res[col] = pd.to_numeric(res[col], errors="coerce")
        return res.dropna().nsmallest(top_n, col).reset_index(drop=True), f"Bottom {top_n} of '{col}'"
    elif op == "Cumulative Sum":
        res = pd.DataFrame({"Row": range(len(series)), col: series.cumsum().values})
        return res, f"Cumulative Sum of '{col}'"
    elif op == "Rank (Desc)":
        res = df[[col]].copy()
        res[col] = pd.to_numeric(res[col], errors="coerce")
        res = res.dropna()
        res["Rank"] = res[col].rank(ascending=False).astype(int)
        return res.sort_values("Rank").reset_index(drop=True), f"Rank by '{col}'"
    else:
        v = series.sum()
    return round(float(v), 4), f"{op} of '{col}'"


# ─────────────────────────────────────────────────────────────────────────────
# CHART FACTORY  –  15 chart types
# ─────────────────────────────────────────────────────────────────────────────
CHART_TYPES = [
    "Bar Chart", "Grouped Bar", "Stacked Bar",
    "Line Chart", "Scatter Plot", "Area Chart",
    "Bubble Chart", "Heatmap (Correlation)", "Box Plot",
    "Violin Plot", "Funnel Chart", "Waterfall / Cumulative",
    "3-D Scatter", "Radar Chart", "Histogram",
]

CHART_DESC = {
    "Bar Chart":              "Compare a numeric metric across categorical groups.",
    "Grouped Bar":            "Side-by-side comparison of multiple numeric columns.",
    "Stacked Bar":            "Show composition and total simultaneously.",
    "Line Chart":             "Trend analysis across ordered rows / time series.",
    "Scatter Plot":           "Correlation between two numeric variables.",
    "Area Chart":             "Cumulative volume trends with filled area.",
    "Bubble Chart":           "Three-dimensional numeric relationships (X, Y, size).",
    "Heatmap (Correlation)":  "Spot which numeric columns are correlated.",
    "Box Plot":               "Distribution, spread, median and outliers.",
    "Violin Plot":            "Full probability distribution shape.",
    "Funnel Chart":           "Staged capacity utilisation visualisation.",
    "Waterfall / Cumulative": "Running-total analysis (e.g. cumulative power).",
    "3-D Scatter":            "Three-axis numeric exploration.",
    "Radar Chart":            "Multi-axis comparison across metrics.",
    "Histogram":              "Frequency distribution of a numeric variable.",
}

CHART_NEEDS = {
    "Bar Chart":              {"x_cat", "y_num"},
    "Grouped Bar":            {"x_cat"},
    "Stacked Bar":            {"x_cat"},
    "Line Chart":             {"y_num"},
    "Scatter Plot":           {"x_num", "y_num", "color"},
    "Area Chart":             {"y_num"},
    "Bubble Chart":           {"x_num", "y_num", "size"},
    "Heatmap (Correlation)":  set(),
    "Box Plot":               {"x_cat", "y_num"},
    "Violin Plot":            {"x_cat", "y_num"},
    "Funnel Chart":           {"x_cat", "y_num"},
    "Waterfall / Cumulative": {"y_num"},
    "3-D Scatter":            {"x_num", "y_num", "z_num"},
    "Radar Chart":            set(),
    "Histogram":              {"y_num", "color"},
}


def _base_layout() -> dict:
    return dict(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,22,40,0.85)",
        font=dict(color=TEXT, family="Segoe UI"),
        margin=dict(l=40, r=30, t=60, b=50),
    )


def make_chart(ct: str, df: pd.DataFrame,
               x=None, y=None, color=None, size=None, z=None,
               title="") -> go.Figure:
    lay = _base_layout()
    nc = num_cols(df)
    tc = txt_cols(df)
    if not x and tc:
        x = tc[0]
    if not y and nc:
        y = nc[0]

    try:
        if ct == "Bar Chart":
            if x and y:
                agg = df.groupby(x)[y].sum().reset_index().sort_values(y, ascending=False).head(30)
                fig = px.bar(agg, x=x, y=y, color=y, title=title, color_continuous_scale="Blues")
            else:
                fig = go.Figure()

        elif ct == "Grouped Bar":
            ys = [c for c in nc if c != x][:4]
            if x and ys:
                agg = df.groupby(x)[ys].sum().reset_index().head(20)
                fig = px.bar(agg, x=x, y=ys, barmode="group", title=title,
                             color_discrete_sequence=px.colors.qualitative.Bold)
            else:
                fig = go.Figure()

        elif ct == "Stacked Bar":
            ys = [c for c in nc if c != x][:4]
            if x and ys:
                agg = df.groupby(x)[ys].sum().reset_index().head(20)
                fig = px.bar(agg, x=x, y=ys, barmode="stack", title=title,
                             color_discrete_sequence=px.colors.qualitative.Pastel)
            else:
                fig = go.Figure()

        elif ct == "Line Chart":
            if y:
                sub = df[[c for c in [x, y] if c]].dropna().reset_index(drop=True)
                kw = dict(y=y, title=title, markers=True, color_discrete_sequence=[CYAN])
                if x:
                    kw["x"] = x
                fig = px.line(sub, **kw)
            else:
                fig = go.Figure()

        elif ct == "Scatter Plot":
            if x and y:
                sub = df.dropna(subset=[c for c in [x, y] if c])
                kw = dict(x=x, y=y, title=title, opacity=0.7, color_discrete_sequence=[CYAN])
                if color and color in df.columns:
                    kw["color"] = color
                    kw.pop("color_discrete_sequence")
                fig = px.scatter(sub, **kw)
            else:
                fig = go.Figure()

        elif ct == "Area Chart":
            if y:
                sub = df[[c for c in [x, y] if c]].dropna().reset_index(drop=True)
                kw = dict(y=y, title=title, color_discrete_sequence=[LBLUE])
                if x:
                    kw["x"] = x
                fig = px.area(sub, **kw)
            else:
                fig = go.Figure()

        elif ct == "Bubble Chart":
            if len(nc) >= 3:
                xc = x if x in nc else nc[0]
                yc = y if y in nc else nc[1]
                sc = size if size in nc else nc[2]
                fig = px.scatter(df.dropna(subset=[xc, yc, sc]),
                                 x=xc, y=yc, size=sc, color=sc,
                                 color_continuous_scale="Blues", opacity=0.75, title=title)
            else:
                fig = go.Figure()

        elif ct == "Heatmap (Correlation)":
            cols = nc[:14]
            if len(cols) >= 2:
                corr = df[cols].corr().round(2)
                fig = go.Figure(go.Heatmap(
                    z=corr.values, x=corr.columns, y=corr.index,
                    colorscale="RdBu", zmid=0,
                    text=corr.values.round(2), texttemplate="%{text}",
                ))
            else:
                fig = go.Figure()

        elif ct == "Box Plot":
            if y:
                kw = dict(y=y, title=title,
                          color_discrete_sequence=px.colors.qualitative.Bold)
                if x:
                    kw["x"] = x
                fig = px.box(df.dropna(subset=[y]), **kw)
            else:
                fig = go.Figure()

        elif ct == "Violin Plot":
            if y:
                kw = dict(y=y, title=title, box=True, points="outliers",
                          color_discrete_sequence=[LBLUE])
                if x:
                    kw["x"] = x
                fig = px.violin(df.dropna(subset=[y]), **kw)
            else:
                fig = go.Figure()

        elif ct == "Funnel Chart":
            if x and y:
                agg = df.groupby(x)[y].sum().reset_index().sort_values(y, ascending=False).head(20)
                fig = go.Figure(go.Funnel(
                    y=agg[x].astype(str), x=agg[y],
                    textinfo="value+percent total",
                ))
            else:
                fig = go.Figure()

        elif ct == "Waterfall / Cumulative":
            if y:
                s = df[y].dropna().head(30)
                fig = go.Figure(go.Waterfall(
                    x=list(range(len(s))), y=s.tolist(),
                    measure=["relative"] * len(s),
                    text=[f"{v:.1f}" for v in s],
                    connector=dict(line=dict(color=BORD)),
                    increasing=dict(marker_color=GREEN),
                    decreasing=dict(marker_color=RED),
                ))
            else:
                fig = go.Figure()

        elif ct == "3-D Scatter":
            if len(nc) >= 3:
                xc = x if x in nc else nc[0]
                yc = y if y in nc else nc[1]
                zc = z if z in nc else nc[2]
                sub = df[[xc, yc, zc]].dropna().head(500)
                fig = go.Figure(go.Scatter3d(
                    x=sub[xc], y=sub[yc], z=sub[zc], mode="markers",
                    marker=dict(size=5, color=sub[zc], colorscale="Blues",
                                opacity=0.8, showscale=True),
                ))
                fig.update_layout(scene=dict(
                    xaxis_title=xc, yaxis_title=yc, zaxis_title=zc, bgcolor=DARK2,
                ))
            else:
                fig = go.Figure()

        elif ct == "Radar Chart":
            cols = nc[:8]
            if len(cols) >= 3:
                vals = df[cols].mean().tolist()
                vals += [vals[0]]
                fig = go.Figure(go.Scatterpolar(
                    r=vals, theta=cols + [cols[0]],
                    fill="toself", line_color=CYAN,
                    fillcolor="rgba(0,212,255,0.15)",
                ))
                fig.update_layout(polar=dict(
                    radialaxis=dict(visible=True, gridcolor=BORD),
                    angularaxis=dict(gridcolor=BORD),
                    bgcolor=DARK2,
                ))
            else:
                fig = go.Figure()

        elif ct == "Histogram":
            if y:
                kw = dict(x=y, nbins=30, title=title, opacity=0.85,
                          color_discrete_sequence=[LBLUE])
                if color and color in df.columns:
                    kw["color"] = color
                    kw.pop("color_discrete_sequence")
                fig = px.histogram(df.dropna(subset=[y]), **kw)
            else:
                fig = go.Figure()

        else:
            fig = go.Figure()

        fig.update_layout(title=title, **lay)
        return fig

    except Exception as exc:
        fig = go.Figure()
        fig.add_annotation(text=f"Chart error: {exc}", x=0.5, y=0.5,
                           showarrow=False, font=dict(color=RED, size=14))
        fig.update_layout(**lay)
        return fig


# ─────────────────────────────────────────────────────────────────────────────
# SMART SEARCH ENGINE  –  Compound multi-intent query parser
# ─────────────────────────────────────────────────────────────────────────────

# Known DC location name fragments
_LOCATIONS = [
    "airoli", "bangalore", "bengaluru", "chennai", "kolkata", "calcutta",
    "noida", "rabale", "vashi",
]

# Aggregation intent keywords → operation label
_OP_KW: dict[str, str] = {
    "sum": "Sum", "total": "Sum", "add": "Sum", "aggregate": "Sum",
    "average": "Mean (Avg)", "avg": "Mean (Avg)", "mean": "Mean (Avg)",
    "median": "Median",
    "minimum": "Min", "min": "Min", "lowest": "Min", "smallest": "Min", "least": "Min",
    "maximum": "Max", "max": "Max", "highest": "Max", "largest": "Max", "biggest": "Max",
    "count": "Count", "how many": "Count", "number of": "Count",
    "std": "Std Deviation", "deviation": "Std Deviation",
    "variance": "Variance",
    "top": "Top N Values", "best": "Top N Values",
    "bottom": "Bottom N Values", "worst": "Bottom N Values",
    "cumulative": "Cumulative Sum", "running": "Cumulative Sum",
    "rank": "Rank (Desc)", "ranking": "Rank (Desc)",
}

# Column concept keywords → regex patterns to search in column names
_COL_CONCEPTS: list[tuple[list[str], str]] = [
    (["power", "kw", "kilowatt"],           r"power|kw|kilowatt"),
    (["capacity", "purchased", "subscribed"],r"capacity|purchased|subscribed"),
    (["usage", "use", "consumption", "used"],r"usage|use|in use|consumption"),
    (["revenue", "mrc", "billing", "charge"],r"revenue|mrc|billing|charge"),
    (["rack", "racks"],                      r"rack"),
    (["customer", "client", "name"],         r"customer|client|name"),
    (["caged"],                              r"caged"),
    (["uncaged"],                            r"caged"),
    (["ownership", "owned"],                 r"ownership"),
    (["space", "area"],                      r"space|area"),
]

# List / filter intent keywords (these do NOT trigger aggregation)
_LIST_KW = {
    "list", "show", "display", "get", "fetch", "give",
    "find", "what", "which", "who", "where", "detail", "details",
}

# Stop-words to drop before keyword matching
_STOP = {
    "a", "an", "the", "of", "in", "for", "to", "is", "are", "was", "were",
    "from", "all", "me", "per", "with", "across", "by", "that", "this",
    "their", "its", "at", "on", "be", "as", "at", "has", "have",
}


def _detect_location_filter(clause: str, df: pd.DataFrame) -> pd.DataFrame:
    """Return rows matching any location name mentioned in clause."""
    if "_Location" not in df.columns:
        return df
    matched_locs = []
    q = clause.lower()
    for loc_kw in _LOCATIONS:
        if loc_kw in q:
            for actual_loc in df["_Location"].unique():
                if loc_kw in actual_loc.lower():
                    matched_locs.append(actual_loc)
    if matched_locs:
        return df[df["_Location"].isin(set(matched_locs))]
    return df


def _best_col_for_concept(concept_words: list, df: pd.DataFrame) -> "str | None":
    """Return the best numeric column matching concept words."""
    nc = num_cols(df)
    for word in concept_words:
        for c in nc:
            if word in c.lower():
                return c
    return None


def _detect_num_col(clause: str, df: pd.DataFrame) -> "str | None":
    """Find the most relevant numeric column for the clause."""
    q = clause.lower()
    for keywords, regex in _COL_CONCEPTS:
        if any(kw in q for kw in keywords):
            nc = num_cols(df)
            for c in nc:
                if re.search(regex, c, re.I):
                    return c
    nc = num_cols(df)
    for priority in (r"power|kw", r"capacity", r"usage|use", r"revenue", r"rack"):
        for c in nc:
            if re.search(priority, c, re.I):
                return c
    return nc[0] if nc else None


def _detect_text_filter(clause: str, df: pd.DataFrame) -> "tuple[pd.DataFrame, list]":
    """Apply text/category filters from the clause, return (filtered_df, matched_keywords)."""
    q = clause.lower()
    words = re.sub(r"[^\w\s]", " ", q).split()
    keywords = [w for w in words
                if w not in _STOP and w not in _OP_KW and w not in _LIST_KW and len(w) > 2]

    tc = [c for c in df.columns if not c.startswith("_")]
    mask = pd.Series([True] * len(df), index=df.index)
    matched = []

    for kw in keywords:
        kw_mask = pd.Series([False] * len(df), index=df.index)
        for col in tc:
            try:
                kw_mask |= df[col].astype(str).str.lower().str.contains(
                    re.escape(kw), na=False)
            except Exception:
                pass
        if kw_mask.any():
            mask &= kw_mask
            matched.append(kw)

    return df[mask].copy(), matched


def _detect_top_n(clause: str) -> int:
    """Extract top-N from clause, e.g. 'top 10 customers' → 10."""
    m = re.search(r"\b(?:top|bottom|best|worst)\s+(\d+)\b", clause, re.I)
    return int(m.group(1)) if m else 10


def _detect_groupby(clause: str, df: pd.DataFrame) -> "str | None":
    """Detect group-by column from 'by X' or 'per X' patterns."""
    m = re.search(r"\b(?:by|per|group by|grouped by)\s+([\w\s]+?)(?:\s+and|\s+in|\s+for|$)",
                  clause, re.I)
    if not m:
        return None
    target = m.group(1).strip().lower()
    if "location" in target or "city" in target or "site" in target:
        return "_Location" if "_Location" in df.columns else None
    for c in df.columns:
        if target in c.lower() and not c.startswith("_"):
            return c
    return None


def _parse_clause_intent(clause: str) -> "str | None":
    """Return 'list', 'aggregate', or 'compare' for a clause."""
    q = clause.lower()
    has_agg = any(kw in q for kw in _OP_KW)
    has_list = any(kw in q for kw in _LIST_KW)
    if has_agg:
        return "aggregate"
    if has_list:
        return "list"
    return "list"  # default


def parse_and_execute(query: str, df: pd.DataFrame) -> list:
    """
    Parse a complex, compound natural-language query and return a list of result dicts.
    Each result has: {title, type: 'table'|'scalar'|'grouped', data, description}
    """
    if df.empty:
        return [{"title": "No data", "type": "error", "description": "DataFrame is empty."}]

    results = []

    # ── Split on ' and ' respecting phrases ───────────────────────────────────
    clauses = re.split(r"\s+and\s+", query.strip(), flags=re.I)
    if not clauses:
        clauses = [query]

    for clause in clauses:
        clause = clause.strip()
        if not clause:
            continue

        # 1. Detect operation
        q = clause.lower()
        detected_op = None
        for kw, op in _OP_KW.items():
            if kw in q:
                detected_op = op
                break

        # 2. Filter by location
        work = _detect_location_filter(clause, df)

        # 3. Detect group-by
        grp = _detect_groupby(clause, work)

        # 4. Detect top-N
        top_n = _detect_top_n(clause)

        # 5. Detect numeric column
        num_col = _detect_num_col(clause, work)

        # 6. Apply text filters (entity / category filters)
        filtered, matched_kws = _detect_text_filter(clause, work)

        # 7. Decide intent
        intent = _parse_clause_intent(clause)

        # ── AGGREGATE INTENT ──────────────────────────────────────────────
        if detected_op and num_col and num_col in filtered.columns:
            if detected_op in ("Top N Values", "Bottom N Values") and grp is None:
                res, desc = run_op(filtered, num_col, detected_op, None, top_n)
                results.append({
                    "title": desc,
                    "type": "table",
                    "data": res if isinstance(res, pd.DataFrame) else pd.DataFrame(),
                    "description": f"Clause: *{clause}*",
                })
            elif grp:
                res, desc = run_op(filtered, num_col, detected_op, grp, top_n)
                results.append({
                    "title": desc,
                    "type": "grouped",
                    "data": res if isinstance(res, pd.DataFrame) else pd.DataFrame(),
                    "description": f"Clause: *{clause}*",
                    "x_col": grp,
                    "y_col": num_col,
                })
            else:
                res, desc = run_op(filtered, num_col, detected_op, None, top_n)
                loc_note = ""
                if "_Location" in filtered.columns:
                    locs = filtered["_Location"].unique().tolist()
                    loc_note = f" | Locations: {', '.join(locs)}" if locs else ""
                results.append({
                    "title": desc,
                    "type": "scalar",
                    "data": res,
                    "description": f"Clause: *{clause}*{loc_note}",
                    "rows_used": len(filtered),
                })

        # ── LIST INTENT ───────────────────────────────────────────────────
        else:
            loc_note = ""
            if "_Location" in filtered.columns:
                locs = filtered["_Location"].unique().tolist()
                loc_note = f" ({', '.join(locs)})" if locs else ""
            kw_note = f"Filtered by: {', '.join(matched_kws)}" if matched_kws else "All records"
            results.append({
                "title": f"Records{loc_note} — {kw_note}",
                "type": "table",
                "data": filtered,
                "description": f"Clause: *{clause}*",
            })

    return results if results else [
        {"title": "No results", "type": "error",
         "description": "Could not interpret the query."}
    ]


# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────
with st.spinner("Loading Excel files…"):
    ALL = load_all()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:16px 0 20px">
      <div style="font-size:2.2rem">🏢</div>
      <div style="font-size:.95rem;font-weight:900;color:{WHITE};letter-spacing:.04em">
        SIFY DATA CENTRE</div>
      <div style="font-size:.7rem;color:{MUTED};margin-top:2px">
        Capacity Intelligence Platform</div>
    </div>""", unsafe_allow_html=True)

    all_locs = sorted(ALL.keys())
    if not all_locs:
        st.error("No Excel files found in the  excel_files/  folder.")
        st.stop()

    sel_locs = st.multiselect("📍 Locations", all_locs, default=all_locs)
    all_sheet_opts = sorted({sn for loc in sel_locs for sn in ALL.get(loc, {})})
    sel_sheets = st.multiselect("📋 Sheets", all_sheet_opts, default=all_sheet_opts)

    st.markdown("---")
    n_loc = len(sel_locs)
    n_sh  = sum(len(ALL.get(l, {})) for l in sel_locs)
    st.markdown(
        f'<div style="font-size:.78rem;color:{MUTED}">Loaded '
        f'<b style="color:{CYAN}">{n_loc}</b> locations | '
        f'<b style="color:{CYAN}">{n_sh}</b> sheets</div>',
        unsafe_allow_html=True,
    )

# ─── apply filters ────────────────────────────────────────────────────────────
fdata = {
    loc: {sn: df for sn, df in ALL[loc].items() if sn in sel_sheets}
    for loc in sel_locs if loc in ALL
}
fdata = {k: v for k, v in fdata.items() if v}

COMB = combined_df(fdata)
CUST = COMB.copy()

# ─────────────────────────────────────────────────────────────────────────────
# HEADER HERO BANNER
# ─────────────────────────────────────────────────────────────────────────────
n_sheets_loaded = sum(len(v) for v in fdata.values())
st.markdown(f"""
<div class="hero">
  <h1>🏢 Sify Data Centre – Capacity Intelligence</h1>
  <p>{n_loc} locations · {n_sheets_loaded} sheets · Real-time analytics &amp; smart query</p>
</div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
T = st.tabs([
    "📊 Dashboard",
    "🔍 Data Explorer",
    "⚙️ Operations",
    "📈 Charts",
    "🧠 Smart Search",
    "🌐 Cross-Location",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 – DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with T[0]:
    cust_c = find_col(CUST, r"customer.?name")
    caged_c = find_col(CUST, r"caged")
    cap_c  = find_col(CUST, r"total capacity purchased|total capacity")
    use_c  = find_col(CUST, r"capacity in use")
    kw_c   = find_col(CUST, r"usage in kw")
    rack_s = find_col(CUST, r"space.*subscription|^subscription$")
    rack_u = find_col(CUST, r"space.*in use|^in use$")
    own_c  = find_col(CUST, r"ownership")
    rev_c  = find_col(CUST, r"total revenue|revenue")

    def _sum(col):
        if col and col in CUST.columns:
            return pd.to_numeric(CUST[col], errors="coerce").sum()
        return 0

    total_cust = CUST[cust_c].dropna().nunique() if cust_c else len(CUST)
    total_cap  = _sum(cap_c)
    total_use  = _sum(use_c)
    total_kw   = _sum(kw_c)
    racks_s    = _sum(rack_s)
    racks_u    = _sum(rack_u)
    total_rev  = _sum(rev_c)

    caged_n = uncaged_n = 0
    if caged_c:
        cv = CUST[caged_c].astype(str).str.upper()
        caged_n   = cv.str.contains("CAGED",   na=False).sum()
        uncaged_n = cv.str.contains("UNCAGED", na=False).sum()

    util_pct = (total_use / total_cap * 100) if total_cap > 0 else 0
    rack_pct = (racks_u  / racks_s   * 100) if racks_s   > 0 else 0

    # ── KPI Row 1 ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Key Performance Indicators</div>',
                unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(kpi_html(fmt(total_cust), "Total Customers",
                         f"Across {n_loc} locations", CYAN), unsafe_allow_html=True)
    c2.markdown(kpi_html(fmt(total_cap),  "Total Capacity (KW)",
                         "Power purchased", LBLUE), unsafe_allow_html=True)
    c3.markdown(kpi_html(fmt(total_use),  "Capacity in Use (KW)",
                         f"{util_pct:.1f}% utilisation", GREEN), unsafe_allow_html=True)
    c4.markdown(kpi_html(fmt(total_kw),   "Actual Usage (KW)",
                         "Metered consumption", AMBER), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── KPI Row 2 ─────────────────────────────────────────────────────────
    c5, c6, c7, c8 = st.columns(4)
    c5.markdown(kpi_html(fmt(caged_n),   "Caged Customers",   "Dedicated caged space",  CYAN),  unsafe_allow_html=True)
    c6.markdown(kpi_html(fmt(uncaged_n), "Uncaged Customers", "Shared / open hall",     LBLUE), unsafe_allow_html=True)
    c7.markdown(kpi_html(fmt(racks_s),   "Racks Subscribed",  "Total racks contracted", GREEN), unsafe_allow_html=True)
    c8.markdown(kpi_html(fmt(total_rev), "Total Revenue",     "Monthly MRC",            RED),   unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Gauges ────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Utilisation Gauges</div>',
                unsafe_allow_html=True)
    g1, g2 = st.columns(2)

    def _gauge(val, label, bar_color):
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=min(val, 100),
            title={"text": label, "font": {"color": TEXT, "size": 14}},
            gauge={
                "axis": {"range": [0, 100], "tickcolor": TEXT},
                "bar":  {"color": bar_color},
                "bgcolor": DARK2,
                "steps": [
                    {"range": [0,  50], "color": "#1a2a1a"},
                    {"range": [50, 80], "color": "#2a2a1a"},
                    {"range": [80,100], "color": "#2a1a1a"},
                ],
                "threshold": {"line": {"color": RED, "width": 3}, "value": 80},
            },
            number={"suffix": "%", "font": {"color": bar_color}},
        ))
        fig.update_layout(**_base_layout(), height=270)
        return fig

    g1.plotly_chart(_gauge(util_pct, "Power Utilisation (%)", LBLUE), use_container_width=True)
    g2.plotly_chart(_gauge(rack_pct, "Rack Occupancy (%)",    GREEN), use_container_width=True)

    # ── Capacity vs Usage by Location ─────────────────────────────────────
    st.markdown('<div class="section-title">Capacity vs Usage by Location</div>',
                unsafe_allow_html=True)
    if cap_c and use_c and "_Location" in CUST.columns:
        la = CUST.groupby("_Location").agg(
            Capacity_Purchased=(cap_c, lambda x: pd.to_numeric(x, errors="coerce").sum()),
            Capacity_in_Use   =(use_c, lambda x: pd.to_numeric(x, errors="coerce").sum()),
        ).reset_index()
        fig_la = px.bar(la, x="_Location",
                        y=["Capacity_Purchased", "Capacity_in_Use"],
                        barmode="group",
                        labels={"_Location": "Location", "value": "KW"},
                        color_discrete_map={"Capacity_Purchased": LBLUE,
                                            "Capacity_in_Use": GREEN})
        fig_la.update_layout(**_base_layout(), height=360)
        st.plotly_chart(fig_la, use_container_width=True)
    elif "_Location" in CUST.columns:
        lc = CUST["_Location"].value_counts().reset_index()
        lc.columns = ["Location", "Records"]
        fig_la = px.bar(lc, x="Location", y="Records",
                        color="Records", color_continuous_scale="Blues")
        fig_la.update_layout(**_base_layout(), height=320)
        st.plotly_chart(fig_la, use_container_width=True)

    # ── Pie Charts ────────────────────────────────────────────────────────
    if caged_c or own_c:
        st.markdown('<div class="section-title">Space &amp; Ownership Split</div>',
                    unsafe_allow_html=True)
        p1, p2 = st.columns(2)
        if caged_c:
            cv = CUST[caged_c].astype(str).str.upper().str.strip()
            cv = cv[cv.isin(["CAGED", "UNCAGED"])]
            if not cv.empty:
                pie = cv.value_counts().reset_index()
                pie.columns = ["Type", "Count"]
                fp = px.pie(pie, names="Type", values="Count",
                            title="Caged vs Uncaged",
                            color_discrete_sequence=[CYAN, LBLUE], hole=0.45)
                fp.update_layout(**_base_layout(), height=300)
                p1.plotly_chart(fp, use_container_width=True)
        if own_c:
            ov = CUST[own_c].astype(str).str.upper().str.strip()
            ov = ov[ov.str.len().between(1, 20)]
            if not ov.empty:
                own = ov.value_counts().reset_index()
                own.columns = ["Type", "Count"]
                fo = px.pie(own, names="Type", values="Count",
                            title="Ownership Split",
                            color_discrete_sequence=[GREEN, AMBER], hole=0.45)
                fo.update_layout(**_base_layout(), height=300)
                p2.plotly_chart(fo, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – DATA EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
with T[1]:
    st.markdown('<div class="section-title">Data Explorer – Column-Level Search</div>',
                unsafe_allow_html=True)

    de_loc = st.selectbox("Location", ["All"] + sorted(fdata.keys()), key="de_loc")
    if de_loc == "All":
        view = COMB.copy()
    else:
        sh_map = fdata.get(de_loc, {})
        de_sh  = st.selectbox("Sheet", list(sh_map.keys()), key="de_sh")
        view   = sh_map.get(de_sh, pd.DataFrame()).copy()
        view.insert(0, "_Location", de_loc)

    if not view.empty:
        disp_cols = [c for c in view.columns if not c.startswith("_")]
        sc1, sc2 = st.columns([1, 2])
        with sc1:
            search_col = st.selectbox("Search in column",
                                      ["Any column"] + disp_cols, key="de_sc")
        with sc2:
            search_val = st.text_input("Search value",
                                       placeholder="Type to filter…", key="de_sv")

        vw = view.copy()
        if search_val.strip():
            sv = search_val.strip().lower()
            if search_col == "Any column":
                mask = vw.apply(
                    lambda r: r.astype(str).str.lower().str.contains(sv, na=False).any(),
                    axis=1)
            else:
                mask = vw[search_col].astype(str).str.lower().str.contains(sv, na=False)
            vw = vw[mask]

        show_c = st.multiselect("Columns to display",
                                disp_cols,
                                default=disp_cols[:min(14, len(disp_cols))],
                                key="de_cols")
        out = vw[[c for c in show_c if c in vw.columns]] if show_c else vw[disp_cols]

        st.markdown(f'<span class="badge">{len(out):,} rows</span>',
                    unsafe_allow_html=True)
        st.dataframe(out.head(500), use_container_width=True, height=420)
        st.download_button("⬇ Download CSV",
                           out.to_csv(index=False).encode(), "sify_data.csv",
                           "text/csv")

        with st.expander("📊 Column Statistics"):
            nc = num_cols(vw)
            if nc:
                st.dataframe(vw[nc].describe().round(3).T, use_container_width=True)
    else:
        st.info("No data for selected filters.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – OPERATIONS
# ══════════════════════════════════════════════════════════════════════════════
with T[2]:
    st.markdown('<div class="section-title">Aggregate Operations Engine</div>',
                unsafe_allow_html=True)

    op_src = st.selectbox("Data source",
                          ["All Locations"] + sorted(fdata.keys()), key="op_src")
    if op_src == "All Locations":
        op_df = CUST.copy()
    else:
        op_sh_map = fdata.get(op_src, {})
        op_sh = st.selectbox("Sheet", list(op_sh_map.keys()), key="op_sh")
        op_df = op_sh_map.get(op_sh, pd.DataFrame()).copy()

    if not op_df.empty:
        nc_op = num_cols(op_df)
        tc_op = txt_cols(op_df)

        oc1, oc2, oc3 = st.columns(3)
        with oc1:
            op_col = st.selectbox("📐 Numeric Column",
                                  nc_op if nc_op else ["(none)"], key="op_col")
        with oc2:
            op_op = st.selectbox("🔧 Operation", OPERATIONS, key="op_op")
        with oc3:
            op_grp = st.selectbox("📦 Group By (optional)",
                                  ["(none)"] + tc_op, key="op_grp")

        op_n = 10
        if op_op in ("Top N Values", "Bottom N Values"):
            op_n = st.slider("N", 5, 50, 10, key="op_n")

        fc1, fc2 = st.columns(2)
        with fc1:
            op_fc = st.selectbox("🔎 Pre-filter Column (optional)",
                                 ["(none)"] + tc_op, key="op_fc")
        with fc2:
            op_fv = ""
            if op_fc != "(none)":
                op_fv = st.text_input("Filter value", key="op_fv")

        if st.button("▶ Run Operation", key="op_run") and nc_op:
            wdf = op_df.copy()
            if op_fc != "(none)" and op_fv.strip():
                wdf = wdf[wdf[op_fc].astype(str).str.lower()
                          .str.contains(op_fv.strip().lower(), na=False)]

            grp = None if op_grp == "(none)" else op_grp
            res, desc = run_op(wdf, op_col, op_op, grp, op_n)

            st.markdown(
                f'<div class="result-box"><b style="color:{MUTED}">{desc}</b>',
                unsafe_allow_html=True)
            if isinstance(res, pd.DataFrame):
                st.dataframe(res.head(100), use_container_width=True)
                st.download_button("⬇ Download", res.to_csv(index=False).encode(),
                                   "op_result.csv", "text/csv")
            elif res is not None:
                st.markdown(f'<div class="result-big">{fmt(res)}</div>',
                            unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

            if isinstance(res, pd.DataFrame) and grp and len(res) > 1:
                fig_op = px.bar(res.head(30), x=grp, y=op_col,
                                color=op_col, color_continuous_scale="Blues", title=desc)
                fig_op.update_layout(**_base_layout(), height=380)
                st.plotly_chart(fig_op, use_container_width=True)
    else:
        st.info("No data for selected source.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – CHARTS
# ══════════════════════════════════════════════════════════════════════════════
with T[3]:
    st.markdown('<div class="section-title">Chart Studio – All 15 Chart Types</div>',
                unsafe_allow_html=True)

    ch_src = st.selectbox("Data source",
                          ["All Locations"] + sorted(fdata.keys()), key="ch_src")
    if ch_src == "All Locations":
        ch_df = CUST.copy()
    else:
        ch_sh_map = fdata.get(ch_src, {})
        ch_sh = st.selectbox("Sheet", list(ch_sh_map.keys()), key="ch_sh")
        ch_df = ch_sh_map.get(ch_sh, pd.DataFrame()).copy()
        if "_Location" not in ch_df.columns:
            ch_df.insert(0, "_Location", ch_src)

    if not ch_df.empty:
        nc_ch = num_cols(ch_df)
        tc_ch = txt_cols(ch_df)

        sel_ct = st.selectbox("📊 Chart Type", CHART_TYPES, key="ch_ct")
        st.markdown(
            f'<div style="background:{DARK2};border:1px solid {BORD};border-radius:8px;'
            f'padding:10px 14px;font-size:.82rem;color:{MUTED};margin-bottom:10px">'
            f'ℹ {CHART_DESC.get(sel_ct, "")}</div>',
            unsafe_allow_html=True)

        needs = CHART_NEEDS.get(sel_ct, set())
        ax1, ax2 = st.columns(2)
        ax3, ax4 = st.columns(2)
        x_v = y_v = color_v = size_v = z_v = None

        if "x_cat" in needs and tc_ch:
            with ax1:
                xs = st.selectbox("X Axis (category)", ["(auto)"] + tc_ch, key="ch_x")
                x_v = None if xs == "(auto)" else xs
        elif "x_num" in needs and nc_ch:
            with ax1:
                xs = st.selectbox("X Axis (numeric)", ["(auto)"] + nc_ch, key="ch_xn")
                x_v = None if xs == "(auto)" else xs

        if "y_num" in needs and nc_ch:
            with ax2:
                ys = st.selectbox("Y Axis / Value", ["(auto)"] + nc_ch, key="ch_y")
                y_v = None if ys == "(auto)" else ys

        if "color" in needs:
            with ax3:
                cs = st.selectbox("Color by", ["(none)"] + tc_ch + nc_ch, key="ch_col")
                color_v = None if cs == "(none)" else cs

        if "size" in needs and nc_ch:
            with ax4:
                ss = st.selectbox("Bubble Size", ["(auto)"] + nc_ch, key="ch_sz")
                size_v = None if ss == "(auto)" else ss

        if "z_num" in needs and nc_ch:
            with ax4:
                zs = st.selectbox("Z Axis", ["(auto)"] + nc_ch, key="ch_z")
                z_v = None if zs == "(auto)" else zs

        if st.button("🎨 Generate Chart", key="ch_gen"):
            fig = make_chart(sel_ct, ch_df, x_v, y_v, color_v, size_v, z_v,
                             title=f"{sel_ct} – {ch_src}")
        else:
            fig = make_chart(sel_ct, ch_df.head(200), title=f"{sel_ct} – Preview")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("📚 Quick Gallery – All 15 Chart Types"):
            g_cols = st.columns(3)
            for i, ct in enumerate(CHART_TYPES):
                with g_cols[i % 3]:
                    st.markdown(f"**{ct}**")
                    fig_sm = make_chart(ct, ch_df.head(80), title="")
                    fig_sm.update_layout(height=220, margin=dict(l=8, r=8, t=10, b=8))
                    st.plotly_chart(fig_sm, use_container_width=True)
    else:
        st.info("No data for selected source.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – SMART SEARCH
# ══════════════════════════════════════════════════════════════════════════════
with T[4]:
    st.markdown('<div class="section-title">🧠 Smart Search &amp; Query Engine</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    <div style="background:{DARK2};border:1px solid {BORD};border-radius:10px;
         padding:16px 20px;margin-bottom:16px;font-size:.86rem;color:{MUTED}">
    <b style="color:{TEXT}">Compound queries supported — use <code>and</code> to chain multiple intents:</b><br><br>
    &nbsp;• <code>list all caged customers and total sum of power in kw</code><br>
    &nbsp;• <code>show uncaged customers in noida and average capacity purchased</code><br>
    &nbsp;• <code>count caged customers by location and total revenue</code><br>
    &nbsp;• <code>top 10 customers by revenue in bangalore and average usage in kw</code><br>
    &nbsp;• <code>list customers in rabale and sum power kw and count uncaged</code><br>
    &nbsp;• <code>maximum capacity airoli and minimum capacity noida</code><br>
    &nbsp;• <code>show all customers and total sum revenue and average power kw</code>
    </div>""", unsafe_allow_html=True)

    query   = st.text_input(
        "🔍 Enter your query",
        placeholder="e.g. list all caged customers and total sum of power in kw",
        key="sq_q")
    sq_locs = st.multiselect("Limit to locations (optional)",
                             ["All"] + sorted(fdata.keys()),
                             default=["All"], key="sq_locs")

    if st.button("🚀 Run Query", key="sq_run") and query.strip():
        pool = CUST.copy()
        if "All" not in sq_locs and sq_locs and "_Location" in pool.columns:
            pool = pool[pool["_Location"].isin(sq_locs)]

        results = parse_and_execute(query, pool)

        st.markdown(
            f'<div style="background:{DARK2};border:1px solid {BORD};border-radius:8px;'
            f'padding:10px 16px;margin-bottom:14px;font-size:.85rem;color:{MUTED}">'
            f'Query parsed into <b style="color:{CYAN}">{len(results)}</b> result block(s)</div>',
            unsafe_allow_html=True)

        all_tables = []

        for i, res in enumerate(results):
            st.markdown(
                f'<div class="section-title">{i+1}. {res["title"]}</div>',
                unsafe_allow_html=True)
            st.caption(res.get("description", ""))

            rtype = res.get("type", "error")

            if rtype == "scalar":
                val = res["data"]
                rows_used = res.get("rows_used", "?")
                if isinstance(val, (int, float)):
                    st.markdown(f"""
                    <div style="background:{DARK2};border:1px solid {BORD};
                         border-radius:14px;padding:24px 32px;text-align:center;margin:8px 0 18px">
                      <div style="font-size:.82rem;color:{MUTED};text-transform:uppercase;
                           font-weight:700;letter-spacing:.06em">{res['title']}</div>
                      <div class="result-big">{fmt(val)}</div>
                      <div style="font-size:.76rem;color:{MUTED};margin-top:8px">
                        Computed from {rows_used:,} rows</div>
                    </div>""", unsafe_allow_html=True)
                else:
                    st.info("Could not compute value.")

            elif rtype in ("table", "grouped"):
                data = res.get("data", pd.DataFrame())
                if isinstance(data, pd.DataFrame) and not data.empty:
                    meta  = [c for c in ["_Location", "_Sheet"] if c in data.columns]
                    show_c = [c for c in data.columns if not c.startswith("_")][:22]
                    display = data[meta + show_c]
                    st.markdown(
                        f'<span class="badge">{len(display):,} rows</span>',
                        unsafe_allow_html=True)
                    st.dataframe(display.head(400), use_container_width=True, height=340)
                    all_tables.append((res["title"], display))

                    # Auto-chart for grouped results
                    if rtype == "grouped":
                        xc = res.get("x_col", "")
                        yc = res.get("y_col", "")
                        if xc and yc and xc in data.columns and yc in data.columns:
                            fig_g = px.bar(data.head(30), x=xc, y=yc,
                                           color=yc, color_continuous_scale="Blues",
                                           title=res["title"])
                            fig_g.update_layout(**_base_layout(), height=350)
                            st.plotly_chart(fig_g, use_container_width=True)

                    # Location breakdown
                    if "_Location" in data.columns and data["_Location"].nunique() > 1:
                        lb = data["_Location"].value_counts().reset_index()
                        lb.columns = ["Location", "Rows"]
                        fig_lb = px.bar(lb, x="Location", y="Rows",
                                        color="Rows", color_continuous_scale="Blues",
                                        title="Row count by Location")
                        fig_lb.update_layout(**_base_layout(), height=300)
                        st.plotly_chart(fig_lb, use_container_width=True)

                    # Numeric distribution
                    nc_res = num_cols(data)
                    if nc_res:
                        fig_h = px.histogram(data.dropna(subset=[nc_res[0]]),
                                             x=nc_res[0], nbins=25,
                                             title=f"Distribution – {nc_res[0]}",
                                             color_discrete_sequence=[LBLUE])
                        fig_h.update_layout(**_base_layout(), height=280)
                        st.plotly_chart(fig_h, use_container_width=True)
                else:
                    st.warning("No matching rows for this clause.")

            else:
                st.error(res.get("description", "Query error."))

            st.markdown("<hr style='border-color:{};margin:10px 0 20px'>".format(BORD),
                        unsafe_allow_html=True)

        # Combined download of all table results
        if all_tables:
            combined_out = pd.concat(
                [df.assign(_ResultBlock=title) for title, df in all_tables],
                ignore_index=True, sort=False
            )
            st.download_button(
                "⬇ Download All Results (CSV)",
                combined_out.to_csv(index=False).encode(),
                "smart_query_results.csv", "text/csv"
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – CROSS-LOCATION
# ══════════════════════════════════════════════════════════════════════════════
with T[5]:
    st.markdown('<div class="section-title">Cross-Location Comparison</div>',
                unsafe_allow_html=True)

    nc_all = num_cols(CUST)
    if not nc_all:
        st.info("No numeric columns found for cross-location comparison.")
    else:
        xl1, xl2, xl3 = st.columns(3)
        with xl1: xl_col = st.selectbox("📐 Metric", nc_all, key="xl_col")
        with xl2: xl_op  = st.selectbox("🔧 Aggregation",
                                        ["Sum", "Mean (Avg)", "Max", "Min", "Count"],
                                        key="xl_op")
        with xl3: xl_ct  = st.selectbox("📊 Chart style",
                                        ["Bar Chart", "Line Chart", "Box Plot", "Radar Chart"],
                                        key="xl_ct")

        rows = []
        for loc in sel_locs:
            loc_df = (CUST[CUST["_Location"] == loc]
                      if "_Location" in CUST.columns else CUST)
            if not loc_df.empty and xl_col in loc_df.columns:
                val, _ = run_op(loc_df, xl_col, xl_op)
                if isinstance(val, (int, float)):
                    rows.append({"Location": loc, xl_col: val})
        xl_agg = (pd.DataFrame(rows).sort_values(xl_col, ascending=False)
                  if rows else pd.DataFrame())

        if not xl_agg.empty:
            k1, k2, k3 = st.columns(3)
            k1.metric("🏆 Highest", xl_agg.iloc[0]["Location"],
                      fmt(xl_agg.iloc[0][xl_col]))
            k2.metric("📉 Lowest",  xl_agg.iloc[-1]["Location"],
                      fmt(xl_agg.iloc[-1][xl_col]))
            k3.metric("Σ Network", "", fmt(xl_agg[xl_col].sum()))

            if xl_ct == "Radar Chart" and len(xl_agg) >= 3:
                vals = xl_agg[xl_col].tolist()
                locs = xl_agg["Location"].tolist()
                fig_xl = go.Figure(go.Scatterpolar(
                    r=vals + [vals[0]], theta=locs + [locs[0]],
                    fill="toself", line_color=CYAN,
                    fillcolor="rgba(0,212,255,0.15)",
                ))
                fig_xl.update_layout(polar=dict(
                    radialaxis=dict(visible=True, gridcolor=BORD),
                    angularaxis=dict(gridcolor=BORD),
                    bgcolor=DARK2,
                ), **_base_layout(), height=440)
            else:
                fig_xl = make_chart(xl_ct, xl_agg, "Location", xl_col,
                                    title=f"{xl_op} of {xl_col} – All Locations")
                fig_xl.update_layout(height=420)
            st.plotly_chart(fig_xl, use_container_width=True)

            st.markdown('<div class="section-title">Summary Table</div>',
                        unsafe_allow_html=True)
            st.dataframe(xl_agg, use_container_width=True)

        if len(nc_all) >= 2 and len(sel_locs) >= 2 and "_Location" in CUST.columns:
            with st.expander("🔥 Full Metrics Heatmap Across Locations"):
                hcols = nc_all[:12]
                hrows = []
                for loc in sel_locs:
                    ld  = CUST[CUST["_Location"] == loc]
                    row = {"Location": loc}
                    for c in hcols:
                        row[c] = pd.to_numeric(ld[c], errors="coerce").sum()
                    hrows.append(row)
                hdf = pd.DataFrame(hrows).set_index("Location")
                hn  = hdf.div(hdf.max()).fillna(0)
                fig_hm = go.Figure(go.Heatmap(
                    z=hn.values,
                    x=[c[:18] for c in hn.columns],
                    y=hn.index,
                    colorscale="Blues",
                    text=hdf.values.round(1),
                    texttemplate="%{text}",
                ))
                fig_hm.update_layout(
                    title="Normalised Metrics Heatmap",
                    **_base_layout(),
                    height=max(300, len(sel_locs) * 60),
                )
                st.plotly_chart(fig_hm, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<hr style="border-color:{BORD};margin-top:2rem">
<div style="text-align:center;color:{MUTED};font-size:.75rem;padding:10px 0 20px">
  Sify Data Centre – Capacity Intelligence Platform &nbsp;|&nbsp;
  {n_loc} locations · {n_sheets_loaded} sheets · 15 chart types · Smart query engine
</div>""", unsafe_allow_html=True)
