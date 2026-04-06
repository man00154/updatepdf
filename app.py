import os
import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import openpyxl
from openai import OpenAI as _OpenAI

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sify DC – Capacity Excel Query",
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
.hero p{{color:{WHITE};margin:6px 0 0;font-size:.95rem}}
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
# EXCEL LOADING ENGINE
# ─────────────────────────────────────────────────────────────────────────────
_BASE = Path(__file__).parent


def _excel_dirs():
    candidates = [
        _BASE / "excel_files",
        _BASE / "attached_assets",
        _BASE,
    ]
    return [d for d in candidates if d.is_dir()]


SECTION_MARKERS = {
    "Billing Model", "Space", "Power Capacity", "Power Usage",
    "Seating Space", "Revenue", "DEMARC", "RHS", "SHS",
    "ONSITE TAPE ROTATION", "OFFSITE TAPE ROTATION",
    "SAFE VAULT", "STORE SPACE", "Contract Information",
    "Floor / Module", "Customer Name",
}

HEADER_INDICATORS = {
    "customer name", "floor", "sr. no", "sno", "s.no",
    "customer", "subscription", "caged", "uncaged", "uom", "in use",
    "power subscription", "billing model", "subscription mode",
    "ownership", "per unit rate", "mrc", "description",
}


def _is_section_row(vals) -> bool:
    non = [str(v).strip() for v in vals if v and str(v).strip() not in ("", "None")]
    if not non:
        return False
    hits = sum(1 for v in non if v in SECTION_MARKERS)
    return hits / len(non) > 0.30


def _is_header_row(vals) -> bool:
    non = [str(v).strip().lower() for v in vals if v and str(v).strip() not in ("", "None")]
    if not non:
        return False
    hits = sum(1 for v in non if any(ind in v for ind in HEADER_INDICATORS))
    return hits / len(non) >= 0.20


def _actual_col_count(rows: list) -> int:
    last = 0
    for row in rows:
        for i in range(len(row) - 1, -1, -1):
            if row[i] is not None and str(row[i]).strip() not in ("", "None"):
                last = max(last, i + 1)
                break
    return last


def _detect_header(raw_rows):
    def rs(r):
        return [str(v).strip() if v is not None else "" for v in r]

    rows = [rs(r) for r in raw_rows[:10]]

    for i in range(min(5, len(rows) - 1)):
        r1 = rows[i]
        r2 = rows[i + 1] if i + 1 < len(rows) else []
        if r2 and (_is_section_row(r1) or _is_header_row(r1)) and _is_header_row(r2):
            data_start = i + 3
            return data_start, r1, r2

    for i in range(min(5, len(rows))):
        r = rows[i]
        if _is_header_row(r):
            data_start = i + 2
            return data_start, [""] * len(r), r

    if rows:
        return 2, [""] * len(rows[0]), rows[0]
    return None, None, None


def _build_cols(g_row: list, c_row: list) -> list:
    cur_g = ""
    cols = []
    seen: dict = {}
    for g, c in zip(g_row, c_row):
        g_s = str(g).strip() if g else ""
        c_s = str(c).strip() if c else ""
        if g_s and g_s not in ("None", ""):
            cur_g = g_s
        label = c_s if c_s and c_s not in ("None", "") else ""
        raw = f"{cur_g} | {label}" if (cur_g and label) else (label or cur_g or "_col")
        cnt = seen.get(raw, 0)
        seen[raw] = cnt + 1
        cols.append(raw if cnt == 0 else f"{raw}.{cnt}")
    return cols


def _clean_df(df: pd.DataFrame) -> "pd.DataFrame | None":
    df = df.dropna(axis=1, how="all")
    df = df[df.apply(
        lambda r: any(str(v).strip() not in ("", "None", "nan") for v in r), axis=1
    )]
    try:
        err_mask = df.apply(
            lambda r: r.astype(str).str.contains(
                r"#DIV|#REF|#N/A|#VALUE", regex=True, na=False
            ).all(), axis=1
        )
        df = df[~err_mask]
    except Exception:
        pass
    for col in df.columns:
        try:
            conv = pd.to_numeric(df[col], errors="coerce")
            col_lower = col.lower()
            is_name_col = any(kw in col_lower for kw in (
                "name", "customer", "floor", "module", "model", "mode",
                "ownership", "caged", "uom", "remarks", "description",
            ))
            if not is_name_col and conv.notna().sum() / max(len(df), 1) > 0.50:
                df[col] = conv
        except Exception:
            pass
    return df.reset_index(drop=True) if len(df) >= 1 else None


def _load_ws(ws) -> "pd.DataFrame | None":
    if ws.max_row < 2:
        return None
    sample = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10), values_only=True))
    actual_ncols = _actual_col_count(sample)
    if actual_ncols < 2:
        return None
    sample = [row[:actual_ncols] for row in sample]
    hdr_start, g_row, c_row = _detect_header(sample)
    if hdr_start is None:
        return None
    g_row = list(g_row)[:actual_ncols]
    c_row = list(c_row)[:actual_ncols]
    while len(g_row) < actual_ncols:
        g_row.append("")
    while len(c_row) < actual_ncols:
        c_row.append("")
    cols = _build_cols(g_row, c_row)
    data = []
    for row in ws.iter_rows(min_row=hdr_start, max_col=actual_ncols, values_only=True):
        vals = list(row)[:len(cols)]
        while len(vals) < len(cols):
            vals.append(None)
        if any(v is not None and str(v).strip() not in ("", "None") for v in vals):
            data.append(vals)
    if not data:
        return None
    return _clean_df(pd.DataFrame(data, columns=cols))


def _load_xls_ws(sheet) -> "pd.DataFrame | None":
    try:
        import xlrd
        nrows, ncols = sheet.nrows, sheet.ncols
        if nrows < 2 or ncols < 2:
            return None

        def cv(r, c):
            try:
                cell = sheet.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    return None
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_NUMBER:
                    iv = int(v)
                    return iv if iv == v else v
                return str(v).strip() if v is not None else None
            except Exception:
                return None

        sample = [[cv(ri, ci) for ci in range(ncols)] for ri in range(min(nrows, 10))]
        actual_ncols = _actual_col_count(sample)
        if actual_ncols < 2:
            return None
        sample = [row[:actual_ncols] for row in sample]
        hdr_start, g_row, c_row = _detect_header(sample)
        if hdr_start is None:
            return None
        g_row = list(g_row)[:actual_ncols]
        c_row = list(c_row)[:actual_ncols]
        while len(g_row) < actual_ncols:
            g_row.append("")
        while len(c_row) < actual_ncols:
            c_row.append("")
        cols = _build_cols(g_row, c_row)
        data = []
        for ri in range(hdr_start - 1, nrows):
            vals = [cv(ri, ci) for ci in range(actual_ncols)][:len(cols)]
            while len(vals) < len(cols):
                vals.append(None)
            if any(v is not None and str(v).strip() not in ("", "None") for v in vals):
                data.append(vals)
        if not data:
            return None
        return _clean_df(pd.DataFrame(data, columns=cols))
    except Exception:
        return None


def _label(fpath: Path) -> str:
    s = re.sub(r"Customer_and_Capacity_Tracker_", "", fpath.stem, flags=re.I)
    s = re.sub(r"_\d{10,}$", "", s)
    s = re.sub(r"_\d{2}[A-Za-z]{3}\d{2,4}", "", s)
    s = re.sub(r"[_]+", " ", s).strip()
    s = re.sub(r"\(\s*(\d+)\s*\)", r"(\1)", s).strip()
    return s if s else fpath.stem


def _file_timestamp(fpath: Path) -> int:
    m = re.search(r"_(\d{10,})(?:\.[^.]+)?$", fpath.stem)
    return int(m.group(1)) if m else 0


@st.cache_data(show_spinner=False)
def load_all() -> dict:
    """Load all 10 Excel files with all their sheets into a nested dict."""
    dirs = _excel_dirs()

    label_map: dict = {}

    def _register(fpath: Path, kind: str) -> None:
        base = _label(fpath)
        ts = _file_timestamp(fpath)
        existing = label_map.get(base)
        if existing is None or ts > existing[0]:
            label_map[base] = (ts, fpath, kind)

    seen_names: set = set()
    for d in dirs:
        for fpath in sorted(d.glob("*.xlsx")):
            if fpath.name not in seen_names:
                seen_names.add(fpath.name)
                _register(fpath, "xlsx")
        for fpath in sorted(d.glob("*.xls")):
            if fpath.suffix.lower() == ".xls" and fpath.name not in seen_names:
                seen_names.add(fpath.name)
                _register(fpath, "xls")

    result: dict = {}
    for label, (_, fpath, kind) in sorted(label_map.items()):
        if kind == "xlsx":
            try:
                wb = openpyxl.load_workbook(str(fpath), data_only=True, read_only=False)
            except Exception:
                try:
                    wb = openpyxl.load_workbook(str(fpath), data_only=False)
                except Exception:
                    continue
            sheets: dict = {}
            for sn in wb.sheetnames:
                try:
                    df = _load_ws(wb[sn])
                    if df is not None and len(df) >= 1:
                        sheets[sn] = df
                except Exception:
                    pass
            wb.close()
            if sheets:
                result[label] = sheets
        else:
            try:
                import xlrd
                wb = xlrd.open_workbook(str(fpath))
                sheets: dict = {}
                for sn in wb.sheet_names():
                    try:
                        df = _load_xls_ws(wb.sheet_by_name(sn))
                        if df is not None and len(df) >= 1:
                            sheets[sn] = df
                    except Exception:
                        pass
                if sheets:
                    result[label] = sheets
            except Exception:
                continue

    return result


# ─────────────────────────────────────────────────────────────────────────────
# DATA HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def combined_df(data: dict, loc_filter=None, sheet_filter=None) -> pd.DataFrame:
    frames = []
    for loc, sheets in data.items():
        if loc_filter and loc not in loc_filter:
            continue
        for sn, df in sheets.items():
            if sheet_filter and sn not in sheet_filter:
                continue
            tmp = df.copy()
            tmp.insert(0, "_Sheet", sn)
            tmp.insert(0, "_Location", loc)
            frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True, sort=False)
    return combined.reset_index(drop=True)


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
    if col not in df.columns:
        return None, f"Column '{col}' not found."
    series = pd.to_numeric(df[col], errors="coerce").dropna()
    if series.empty:
        return None, "No numeric data in column."

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
# CHART FACTORY
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
    "Grouped Bar":            "Side-by-side comparison of multiple numeric columns across groups.",
    "Stacked Bar":            "Show composition and total simultaneously across groups.",
    "Line Chart":             "Trend analysis across ordered rows or time-series data.",
    "Scatter Plot":           "Correlation between two numeric variables.",
    "Area Chart":             "Cumulative volume trends with filled area.",
    "Bubble Chart":           "Three-dimensional numeric relationships (X, Y, size).",
    "Heatmap (Correlation)":  "Spot which numeric columns are correlated.",
    "Box Plot":               "Distribution, spread, median, and outliers.",
    "Violin Plot":            "Full probability distribution shape.",
    "Funnel Chart":           "Staged capacity utilisation visualisation.",
    "Waterfall / Cumulative": "Running total analysis.",
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
                s = pd.to_numeric(df[y], errors="coerce").dropna().head(30)
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
# SMART QUERY ENGINE  – fully corrected AND/OR compound query parser
# ─────────────────────────────────────────────────────────────────────────────

_LOCATIONS_KW = [
    "airoli", "bangalore", "bengaluru", "chennai", "kolkata", "calcutta",
    "noida", "rabale", "vashi",
]

_OP_KW: dict = {
    "sum": "Sum", "total": "Sum", "add": "Sum", "aggregate": "Sum",
    "average": "Mean (Avg)", "avg": "Mean (Avg)", "mean": "Mean (Avg)",
    "median": "Median",
    "minimum": "Min", "min": "Min", "lowest": "Min", "smallest": "Min", "least": "Min",
    "maximum": "Max", "max": "Max", "highest": "Max", "largest": "Max", "biggest": "Max",
    "count": "Count", "number of": "Count", "how many": "Count",
    "std": "Std Deviation", "deviation": "Std Deviation", "standard deviation": "Std Deviation",
    "variance": "Variance",
    "top": "Top N Values", "best": "Top N Values",
    "bottom": "Bottom N Values", "worst": "Bottom N Values",
    "cumulative": "Cumulative Sum", "running": "Cumulative Sum",
    "rank": "Rank (Desc)", "ranking": "Rank (Desc)",
}

_COL_CONCEPT_WORDS: set = {
    "power", "kw", "kilowatt", "kva", "capacity", "purchased", "subscribed",
    "usage", "use", "consumption", "revenue", "mrc", "billing", "charge",
    "rack", "racks", "space", "area", "floor", "sitting", "seat",
    "rate", "unit", "contract", "term", "expiry", "frequency",
}

_COL_CONCEPTS: list = [
    (["allocated capacity", "allocated kw"],              r"allocated.*kw|kw.*allocated"),
    (["subscribed capacity kw", "capacity to be given in kw"], r"subscribed.*capacity.*kw|capacity.*to.*be.*given.*kw"),
    (["capacity to be given", "capacity remaining"],      r"capacity.*to.*be.*given|remaining.*capacity"),
    (["reserved capacity"],                               r"reserved.*capacity"),
    (["capacity purchased", "total capacity"],            r"total.*capacity|capacity.*purchased"),
    (["capacity in use", "used capacity"],                r"capacity.*in.*use"),
    (["in use"],                                          r"in.*use"),
    (["yet to be given", "yet to be billed"],             r"yet.*to.*be|yet.*billed"),
    (["power", "kw", "kilowatt", "kva"],                  r"total.*capacity.*kw|capacity.*kw|power|kilowatt|kva|\bkw\b"),
    (["capacity", "purchased", "subscribed"],             r"capacity|purchased|subscribed"),
    (["usage", "consumption", "kw hr", "kwhr"],           r"kw.*hr|kwhr|usage|consumption"),
    (["space revenue", "space including capacity"],       r"space.*revenue|space.*including.*capacity"),
    (["additional capacity revenue", "additional capacity charges"], r"additional.*capacity.*revenue|additional.*capacity.*charge"),
    (["power usage revenue", "power revenue"],            r"power.*usage.*revenue|power.*revenue"),
    (["seating space revenue", "seating revenue"],        r"seating.*space.*revenue|seating.*revenue"),
    (["other items", "any other"],                        r"any.*other|other.*items"),
    (["total revenue", "total mrc"],                      r"total.*revenue|total.*mrc"),
    (["revenue", "mrc", "billing", "charge"],             r"revenue|mrc"),
    (["per unit rate", "per unit", "mrc rate"],           r"per.*unit.*rate|per.*unit"),
    (["subscription", "sitting space subscription"],      r"subscription"),
    (["sitting space", "seat", "seating"],                r"sit|seat"),
    (["rack", "racks"],                                   r"\brack\b"),
    (["space", "area"],                                   r"space|area"),
    (["contract start", "start date"],                    r"contract.*start|start.*date"),
    (["expiry", "expiry date", "current expiry"],         r"expir"),
    (["term", "term of contract", "years"],               r"term.*contract|term.*year"),
    (["billing frequency", "frequency"],                  r"billing.*freq|frequency"),
    (["sales order", "so ref"],                           r"sales.*order|so.*ref"),
    (["contract", "contract info"],                       r"contract"),
    (["customer", "client", "name"],                      r"customer.*name|client.*name"),
    (["rate", "per unit"],                                r"per.*unit|rate"),
    (["remarks", "notes"],                                r"remarks|notes"),
]

_LIST_KW = {
    "list", "show", "display", "get", "fetch", "give",
    "find", "what", "which", "who", "where", "detail", "details",
    "all", "every",
}

_STOP = {
    "a", "an", "the", "of", "in", "for", "to", "is", "are", "was", "were",
    "from", "me", "per", "with", "across", "by", "that", "this",
    "their", "its", "at", "on", "be", "as", "has", "have", "do",
}

_NON_FILTER_WORDS: set = (
    set(_OP_KW.keys()) |
    _COL_CONCEPT_WORDS |
    _LIST_KW |
    _STOP |
    set(_LOCATIONS_KW) |
    {"and", "or", "not", "caged", "uncaged", "by", "location", "customer",
     "customers", "site", "dc", "data", "centre", "center"}
)


def _loc_filter(clause: str, df: pd.DataFrame) -> pd.DataFrame:
    """Filter rows by location keyword found in clause."""
    if "_Location" not in df.columns:
        return df
    q = clause.lower()
    matched_locs = []
    for loc in df["_Location"].unique():
        loc_lower = loc.lower()
        for kw in _LOCATIONS_KW:
            if kw in q and kw in loc_lower:
                matched_locs.append(loc)
                break
    return df[df["_Location"].isin(matched_locs)] if matched_locs else df


def _detect_num_col(clause: str, df: pd.DataFrame) -> "str | None":
    """Identify which numeric column the query is referring to."""
    q = clause.lower()
    nc = num_cols(df)
    if not nc:
        return None

    for keywords, regex in _COL_CONCEPTS:
        if any(kw in q for kw in keywords):
            for c in nc:
                if re.search(regex, c, re.I):
                    return c

    for priority in (r"total.*capacity|capacity.*kw", r"capacity.*in.*use|in.*use",
                     r"power|kw", r"capacity", r"revenue", r"rack"):
        for c in nc:
            if re.search(priority, c, re.I):
                return c
    return nc[0] if nc else None


def _detect_groupby(clause: str, df: pd.DataFrame) -> "str | None":
    m = re.search(r"\b(?:by|per|grouped?\s*by)\s+([\w\s]+?)(?:\s+(?:and|in|for|or)|$)",
                  clause, re.I)
    if not m:
        return None
    target = m.group(1).strip().lower()
    if any(kw in target for kw in ("location", "city", "site", "dc")):
        return "_Location" if "_Location" in df.columns else None
    for c in df.columns:
        if not c.startswith("_") and target.replace(" ", "") in c.lower().replace(" ", ""):
            return c
    return None


def _detect_caged_filter(clause: str, df: pd.DataFrame) -> "tuple[pd.DataFrame, list]":
    q = clause.lower()
    applied = []

    caged_col = None
    for c in df.columns:
        if re.search(r"\bcaged\b", c, re.I):
            caged_col = c
            break

    if caged_col is None:
        return df, applied

    col_vals = df[caged_col].astype(str).str.strip().str.upper()

    want_uncaged = bool(re.search(r"\buncaged\b", q))
    want_caged = bool(re.search(r"\bcaged\b", q)) and not want_uncaged

    if want_uncaged:
        mask = col_vals.isin(["UNCAGED", "UN-CAGED", "UN CAGED"])
        if mask.any():
            applied.append("uncaged")
            return df[mask].copy(), applied
    elif want_caged:
        mask = col_vals.isin(["CAGED"])
        if mask.any():
            applied.append("caged")
            return df[mask].copy(), applied

    return df, applied


def _detect_customer_filter(clause: str, df: pd.DataFrame) -> "tuple[pd.DataFrame, list]":
    q = clause.lower()

    quoted = re.findall(r'"([^"]+)"|\'([^\']+)\'', q)
    candidates = [a or b for a, b in quoted]

    cust_col = find_col(df, r"customer.*name|client.*name")
    if not cust_col:
        return df, []

    if candidates:
        for cand in candidates:
            mask = df[cust_col].astype(str).str.lower().str.contains(
                re.escape(cand.lower()), na=False)
            if mask.any():
                return df[mask].copy(), [cand]

    return df, []


def _detect_billing_model_filter(clause: str, df: pd.DataFrame) -> "tuple[pd.DataFrame, list]":
    q = clause.lower()
    applied = []
    work = df

    pw_sub_col = find_col(work, r"power.*subscription.*model|billing.*model.*power.*subscription")
    if pw_sub_col:
        vals = work[pw_sub_col].astype(str).str.strip().str.upper()
        if re.search(r"\brated\b", q):
            mask = vals.str.contains("RATED", na=False)
            if mask.any():
                work = work[mask].copy()
                applied.append("Power Sub Model = Rated")
        elif re.search(r"\bsubscribed\b", q):
            mask = vals.str.contains("SUBSCRIBED|SUBSCRIB", na=False)
            if mask.any():
                work = work[mask].copy()
                applied.append("Power Sub Model = Subscribed")

    pw_use_col = find_col(work, r"power.*usage.*model|billing.*model.*power.*usage")
    if pw_use_col:
        vals = work[pw_use_col].astype(str).str.strip().str.upper()
        if re.search(r"\bbundled\b", q) and not applied:
            mask = vals.str.contains("BUNDLED", na=False)
            if mask.any():
                work = work[mask].copy()
                applied.append("Power Usage = Bundled")
        elif re.search(r"\bmetered\b", q) and not applied:
            mask = vals.str.contains("METERED", na=False)
            if mask.any():
                work = work[mask].copy()
                applied.append("Power Usage = Metered")

    own_col = find_col(work, r"\brhs\b|\bshs\b|ownership")
    if own_col and not applied:
        vals = work[own_col].astype(str).str.strip().str.upper()
        if re.search(r"\brhs\b", q):
            mask = vals.str.contains("RHS", na=False)
            if mask.any():
                work = work[mask].copy()
                applied.append("Ownership = RHS")
        elif re.search(r"\bshs\b", q):
            mask = vals.str.contains("SHS", na=False)
            if mask.any():
                work = work[mask].copy()
                applied.append("Ownership = SHS")

    return work, applied


def _clause_has_own_filter(clause: str) -> bool:
    q = clause.lower()
    return bool(
        re.search(r"\b(caged|uncaged|rated|subscribed|bundled|metered|rhs|shs)\b", q)
        or re.findall(r'"([^"]+)"|\'([^\']+)\'', q)
    )


def _detect_text_filter(clause: str, df: pd.DataFrame) -> "tuple[pd.DataFrame, list]":
    work, applied = _detect_caged_filter(clause, df)

    if not applied:
        work, a2 = _detect_billing_model_filter(clause, work)
        applied += a2

    q = clause.lower()
    quoted = re.findall(r'"([^"]+)"|\'([^\']+)\'', q)
    if quoted:
        work2, a3 = _detect_customer_filter(clause, work)
        if a3:
            work = work2
            applied += a3

    return work, applied


def _detect_top_n(clause: str) -> int:
    m = re.search(r"\b(?:top|bottom|best|worst)\s+(\d+)\b", clause, re.I)
    return int(m.group(1)) if m else 10


def _parse_or_clauses(query: str) -> list:
    return re.split(r"\s+or\s+", query, flags=re.I)


# ─────────────────────────────────────────────────────────────────────────────
# SMART QUERY: Full-text search fallback helper
# ─────────────────────────────────────────────────────────────────────────────

def _full_text_search(query: str, df: pd.DataFrame) -> "pd.DataFrame":
    """
    Search all text/string columns for words from the query that are NOT
    in _NON_FILTER_WORDS. Returns matching rows (may be empty).
    Only triggered when no structured filter matched.
    """
    tokens = re.findall(r'\b\w[\w\s]*?\b', query.lower())
    search_terms = [
        t.strip() for t in tokens
        if t.strip() and t.strip() not in _NON_FILTER_WORDS and len(t.strip()) > 2
    ]
    if not search_terms:
        return pd.DataFrame()

    str_cols = [c for c in df.columns if not c.startswith("_") and
                df[c].dtype == object]
    if not str_cols:
        return pd.DataFrame()

    combined_mask = pd.Series([False] * len(df), index=df.index)
    for term in search_terms:
        term_mask = df[str_cols].apply(
            lambda col: col.astype(str).str.lower().str.contains(
                re.escape(term), na=False)
        ).any(axis=1)
        combined_mask = combined_mask | term_mask

    return df[combined_mask].copy()


def execute_clause(clause: str, df: pd.DataFrame,
                   context_df: "pd.DataFrame | None" = None) -> dict:
    """
    Execute a single query clause and return a result dict.
    """
    if df.empty:
        return {"title": "No data", "type": "error", "description": "DataFrame is empty."}

    q = clause.lower()

    detected_op = None
    for kw, op in sorted(_OP_KW.items(), key=lambda x: -len(x[0])):
        if re.search(r"\b" + re.escape(kw) + r"\b", q):
            detected_op = op
            break

    is_listing = any(kw in q for kw in _LIST_KW)

    inherit_context = (
        context_df is not None
        and not context_df.empty
        and detected_op is not None
        and not _clause_has_own_filter(clause)
    )
    base = context_df if inherit_context else df
    ctx_label = " [on filtered rows]" if inherit_context else ""

    work = _loc_filter(clause, base)

    filtered, matched_kws = _detect_text_filter(clause, work)

    grp = _detect_groupby(clause, filtered)
    top_n = _detect_top_n(clause)

    num_col = _detect_num_col(clause, filtered)

    loc_note = ""
    if "_Location" in filtered.columns:
        locs = filtered["_Location"].unique().tolist()
        loc_note = f" ({', '.join(locs)})" if locs and len(locs) < 5 else ""

    rows_used = len(filtered)
    filter_label = (
        ("Filtered by: " + ", ".join(matched_kws)) if matched_kws
        else ("Inherited context" if inherit_context else "All records")
    )

    if detected_op and num_col and num_col in filtered.columns:
        if detected_op in ("Top N Values", "Bottom N Values") and grp is None:
            cust_col = find_col(filtered, r"customer.*name|client.*name")
            extra_cols = []
            if cust_col and cust_col != num_col:
                extra_cols.append(cust_col)
            if "_Location" in filtered.columns:
                extra_cols.append("_Location")
            res_cols = extra_cols + [num_col]
            res_cols = [c for c in res_cols if c in filtered.columns]
            res_df = filtered[res_cols].copy()
            res_df[num_col] = pd.to_numeric(res_df[num_col], errors="coerce")
            res_df = res_df.dropna(subset=[num_col])
            if detected_op == "Top N Values":
                res_df = res_df.nlargest(top_n, num_col).reset_index(drop=True)
                desc = f"Top {top_n} of '{num_col}'"
            else:
                res_df = res_df.nsmallest(top_n, num_col).reset_index(drop=True)
                desc = f"Bottom {top_n} of '{num_col}'"
            return {"title": desc + ctx_label, "type": "table",
                    "data": res_df,
                    "description": f"*{clause}*",
                    "filter_label": filter_label, "rows_used": rows_used}

        elif grp:
            res, desc = run_op(filtered, num_col, detected_op, grp, top_n)
            return {"title": desc + ctx_label, "type": "grouped",
                    "data": res if isinstance(res, pd.DataFrame) else pd.DataFrame(),
                    "description": f"*{clause}*{loc_note}",
                    "x_col": grp, "y_col": num_col,
                    "filter_label": filter_label, "rows_used": rows_used}
        else:
            res, desc = run_op(filtered, num_col, detected_op, None, top_n)
            return {"title": desc + ctx_label, "type": "scalar", "data": res,
                    "description": f"*{clause}*{loc_note}",
                    "rows_used": rows_used, "filter_label": filter_label,
                    "num_col": num_col}

    else:
        if matched_kws or is_listing:
            return {"title": f"Records{loc_note} — {filter_label}",
                    "type": "table", "data": filtered,
                    "description": f"*{clause}*",
                    "filter_label": filter_label, "rows_used": rows_used}

        fts = _full_text_search(clause, filtered)
        if not fts.empty:
            return {
                "title": f"Text Search Results{loc_note}",
                "type": "table", "data": fts,
                "description": f"*{clause}*",
                "filter_label": f"Text match: {clause[:40]}",
                "rows_used": len(fts)
            }

        return {"title": f"Records{loc_note}", "type": "table",
                "data": filtered, "description": f"*{clause}*",
                "filter_label": filter_label, "rows_used": rows_used}


def _extract_filter_context(result: dict) -> "pd.DataFrame | None":
    """Return the filtered DataFrame from a result if it is a listing/filter result."""
    if result.get("type") == "table":
        d = result.get("data")
        if isinstance(d, pd.DataFrame) and not d.empty:
            return d
    return None


def parse_and_execute(query: str, df: pd.DataFrame) -> list:
    """
    Parse compound query with AND/OR logic and execute against real data.
    """
    if df.empty:
        return [{"title": "No data", "type": "error", "description": "DataFrame is empty."}]

    results = []
    and_clauses = re.split(r"\s+and\s+", query.strip(), flags=re.I)
    running_context: "pd.DataFrame | None" = None

    for and_clause in and_clauses:
        and_clause = and_clause.strip()
        if not and_clause:
            continue

        or_clauses = _parse_or_clauses(and_clause)

        if len(or_clauses) > 1:
            union_frames = []
            scalar_results = []
            for or_c in or_clauses:
                res = execute_clause(or_c.strip(), df)
                if res["type"] in ("table", "grouped") and isinstance(res.get("data"), pd.DataFrame):
                    union_frames.append(res["data"])
                elif res["type"] == "scalar":
                    scalar_results.append(res)

            if union_frames:
                merged = pd.concat(union_frames, ignore_index=True, sort=False).drop_duplicates()
                results.append({
                    "title": f"OR Combined — {and_clause[:60]}",
                    "type": "table", "data": merged.reset_index(drop=True),
                    "description": f"*Union of: {' | '.join(or_clauses)}*",
                    "filter_label": "OR union", "rows_used": len(merged),
                })
            elif scalar_results:
                for sr in scalar_results:
                    results.append(sr)
            else:
                for or_c in or_clauses:
                    results.append(execute_clause(or_c.strip(), df))

            running_context = None

        else:
            res = execute_clause(and_clause, df, context_df=running_context)
            results.append(res)

            if res.get("type") == "table":
                new_ctx = _extract_filter_context(res)
                if new_ctx is not None:
                    running_context = new_ctx
            elif _clause_has_own_filter(and_clause):
                filter_only = execute_clause(and_clause, df, context_df=None)
                if filter_only.get("type") == "table":
                    new_ctx = _extract_filter_context(filter_only)
                    if new_ctx is not None:
                        running_context = new_ctx

    return results or [{"title": "No results", "type": "error",
                        "description": "Could not interpret the query."}]


# ─────────────────────────────────────────────────────────────────────────────
# SMART QUERY AI ENGINE  — structured AI parse → real data execution → table/metric display
# ─────────────────────────────────────────────────────────────────────────────

# ── Exact semantic column registry (derived from actual Sify DC Excel files) ─
# Maps: semantic_key → (regex_pattern_against_column_name, priority)
_SEMANTIC_COLS: dict = {
    "total_power":              (r"Power Capacity.*Total Capacity Purchased",          1),
    "power_in_use":             (r"Power Capacity.*Capacity in Use|Power Capacity.*Usage in KW", 1),
    "power_allocated":          (r"Power Capacity.*Allocated.*Capacity",               1),
    "power_reserved":           (r"Power Capacity.*Reserved Capacity",                 1),
    "power_additional_mrc":     (r"Power Capacity.*Additional Capacity Charges",       1),
    "power_subscribed_given":   (r"Power Capacity.*Subscribed Capacity to be given",   1),
    "space_subscription":       (r"^Space \| Subscription$",                           1),
    "space_in_use":             (r"^Space \| In Use$",                                 1),
    "space_billed":             (r"^Space \| Billed$",                                 1),
    "space_reserved":           (r"Space.*Reserved Capacity",                          1),
    "total_revenue":            (r"Revenue.*Total Revenue",                            1),
    "revenue_space":            (r"Revenue.*Space revenue",                            1),
    "revenue_power":            (r"Revenue.*Power Usage revenue",                      1),
    "revenue_additional":       (r"Revenue.*Additional Capacity Revenue",              1),
    "revenue_seating":          (r"Revenue.*Seating Space",                            1),
    "revenue_other":            (r"Revenue.*Any Other Items",                          1),
    "net_rev_total":            (r"Contract Information.*Net Rev Total",               2),
    "rev_cap_power":            (r"Contract Information.*Total Rev.*Cap.*Power",       2),
    "seating_subscription":     (r"^Seating Space \| Subscription$",                  1),
    "seating_in_use":           (r"^Seating Space \| In Use$",                        1),
    "per_unit_rate":            (r"Per Unit rate|per.*unit.*rate",                     1),
}

# Maps field_hint keywords → semantic_key (in priority order)
_HINT_SEMANTIC: list = [
    # Power / Capacity
    # ── Power / capacity ─────────────────────────────────────────────────────
    ("total capacity purchased",    "total_power"),
    ("total power",                 "total_power"),
    ("sum of power",                "total_power"),
    ("sum power",                   "total_power"),
    ("power sum",                   "total_power"),
    ("total kw purchased",          "total_power"),
    ("power capacity",              "total_power"),
    ("capacity purchased",          "total_power"),
    ("sum of capacity",             "total_power"),
    ("sum capacity",                "total_power"),
    ("total capacity purchased",    "total_power"),
    ("total capacity",              "total_power"),
    ("power total",                 "total_power"),
    ("power kw",                    "total_power"),
    ("total kw",                    "total_power"),
    ("subscribed kw",               "total_power"),
    ("allocated kw",                "power_allocated"),
    # ── Power in use ──────────────────────────────────────────────────────────
    ("power in use",                "power_in_use"),
    ("capacity in use",             "power_in_use"),
    ("sum of power in use",         "power_in_use"),
    ("power used",                  "power_in_use"),
    ("power usage",                 "power_in_use"),
    ("usage in kw",                 "power_in_use"),
    ("power usage kw",              "power_in_use"),
    ("kw in use",                   "power_in_use"),
    ("power reserved",              "power_reserved"),
    # Space
    ("total space used",            "space_in_use"),
    ("space used",                  "space_in_use"),
    ("space in use",                "space_in_use"),
    ("total space",                 "space_subscription"),
    ("space subscription",          "space_subscription"),
    ("space subscribed",            "space_subscription"),
    ("space purchased",             "space_subscription"),
    ("space billed",                "space_billed"),
    # ── Revenue ──────────────────────────────────────────────────────────────
    ("total revenue",               "total_revenue"),
    ("revenue total",               "total_revenue"),
    ("total mrc",                   "total_revenue"),
    ("total monthly revenue",       "total_revenue"),
    ("sum of revenue",              "total_revenue"),
    ("mrc",                         "total_revenue"),
    ("revenue including capacity",  "revenue_space"),
    ("space revenue",               "revenue_space"),
    ("revenue from space",          "revenue_space"),
    ("power usage revenue",         "revenue_power"),
    ("power revenue",               "revenue_power"),
    ("revenue from power",          "revenue_power"),
    ("seating revenue",             "revenue_seating"),
    ("additional capacity revenue", "revenue_additional"),
    ("other revenue",               "revenue_other"),
    ("net revenue",                 "net_rev_total"),
    # Seating
    ("seating subscription",        "seating_subscription"),
    ("seating in use",              "seating_in_use"),
    ("seating space",               "seating_subscription"),
    # Rate
    ("per unit rate",               "per_unit_rate"),
    ("unit rate",                   "per_unit_rate"),
]


def _detect_unit(col_name: str) -> str:
    if not col_name:
        return ""
    c = col_name.lower()
    if "kva"  in c:                                           return "KVA"
    if any(k in c for k in ("kwhr", "kw hr", "kw-hr", "unit rate")):
        return "₹/kWh"
    if any(k in c for k in ("revenue", "mrc", "per unit rate", "charges", "billing")):
        return "₹"
    if any(k in c for k in ("kw", "kilowatt")):              return "kW"
    if any(k in c for k in ("sqft", "sq ft", "sq.ft", "subscription", "space", "area")):
        return "sq.ft"
    if "rack" in c:                                           return "racks"
    if any(k in c for k in ("seat", "sitting")):              return "seats"
    return ""


_AI_PARSER_PROMPT = """# SYSTEM PROMPT: Sify Data Centre Excel Query Engine

## IDENTITY & MISSION
You are an ultra-precise data retrieval engine for Sify Technologies Ltd. Data Centre Customer & Capacity Tracker Excel files (10 files covering all India locations, all sheets). Your job is to parse the user's natural language query into a structured JSON array that a Python executor will run against the actual DataFrames. You NEVER guess, assume, or hallucinate values. Every field_hint you choose must directly map to a real column in the data.

## CRITICAL RULES
1. ZERO TOLERANCE FOR FABRICATION — every value must be traceable to actual data.
2. DECIMAL PRECISION IS SACRED — NEVER round, truncate, or approximate. If a cell has 530.0311160714285, preserve all decimal places.
3. ALL FILES, ALL SHEETS — data spans 10 Excel files across all India DC locations (Airoli, Rabale T1/T2, Rabale Tower 4, Rabale Tower 5, Bangalore 01, Noida 01, Noida 02, Chennai, Kolkata, Vashi). The Python executor queries all of them. Do NOT restrict the location unless the user explicitly names one.
4. CASE-INSENSITIVE MATCHING — caged/CAGED/Caged all mean the same. Apply same logic for rated/subscribed/bundled/metered.
5. RETURN ONLY raw JSON array — no markdown, no prose, no code fences. Output must be parseable by json.loads().

## JSON OUTPUT FORMAT
Return a JSON array. Each element is one operation:
{
  "id": "op1",
  "type": "list" | "aggregate" | "count",
  "label": "short human-readable label for this result card",
  "filter": {
    "caged":      true | false | null,
    "uncaged":    true | false | null,
    "rated":      true | false | null,
    "subscribed": true | false | null,
    "bundled":    true | false | null,
    "metered":    true | false | null,
    "rhs":        true | false | null,
    "shs":        true | false | null
  },
  "location": ["airoli","noida"] | null,
  "operation": "sum"|"avg"|"mean"|"min"|"max"|"count"|"std"|"median"|"variance"|"range"|"count_nonzero"|"top"|"bottom" | null,
  "field_hint": "<exact phrase from the list below>" | null,
  "top_n": integer | null,
  "group_by_location": true | false
}

## EXACT field_hint PHRASES (use verbatim — Python maps these to real columns):
Power / Capacity:
- "total capacity purchased"    — Total KW/KVA purchased (subscribed capacity)
- "power in use"                — Capacity in Use (KW/KVA currently consumed)
- "power allocated"             — Allocated / subscribed capacity in KW to be given
- "power usage kw"              — Actual Usage in KW (metered consumption)

Space / Racks:
- "total space"                 — Space Subscription (sqft)
- "space in use"                — Space In Use (sqft)
- "space billed"                — Space Billed (sqft)
- "seating subscription"        — Seating Space Subscription (seats)
- "seating in use"              — Seating Space In Use (seats)

Revenue (all are ₹/month MRC):
- "total revenue"               — Total Monthly Revenue (MRC)
- "space revenue"               — Revenue from Space including capacity
- "power revenue"               — Revenue from Power Usage
- "seating revenue"             — Revenue from Seating Space
- "additional capacity revenue" — Additional Capacity Revenue
- "net revenue"                 — Net Revenue Total (Contract Information)

Rate:
- "per unit rate"               — Per Unit Rate / tariff (₹/KW-HR)

## OPERATION → FUNCTION MAPPING
| User says                                | operation value |
|------------------------------------------|-----------------|
| sum / total / aggregate / add up         | "sum"           |
| average / avg / mean                     | "avg"           |
| minimum / min / lowest / smallest        | "min"           |
| maximum / max / highest / largest        | "max"           |
| count / how many / number of             | "count"         |
| list / show / display / get              → type="list", operation=null |
| standard deviation / std / deviation     | "std"           |
| median / middle value                    | "median"        |
| variance                                 | "variance"      |
| range / spread                           | "range"         |
| top N / largest N                        | "top" + top_n   |
| bottom N / smallest N                    | "bottom" + top_n|

## LOCATION ALIASES (resolve BROADLY — always include all sub-locations)
- "airoli"                          → ["airoli"]
- "vashi"                           → ["vashi"]
- "rabale" / "rabale tower"         → ["rabale"]   (matches T1, T2, Tower4, Tower5)
- "rabale t1" / "rabale tower 1"    → ["rabale t1"]
- "rabale t2" / "rabale tower 2"    → ["rabale t2"]
- "rabale tower 4"                  → ["rabale tower 4"]
- "rabale tower 5" / "rabale t5"    → ["rabale tower 5"]
- "noida"                           → ["noida"]    (matches BOTH Noida 01 AND Noida 02)
- "noida 01" / "noida-01"           → ["noida 01"]
- "noida 02" / "noida-02"           → ["noida 02"]
- "bangalore" / "bengaluru"         → ["bangalore"]
- "chennai"                         → ["chennai"]
- "kolkata" / "calcutta"            → ["kolkata"]
- "mumbai" / "mumbai region"        → ["airoli","rabale","vashi"]
- no location mentioned             → null (query ALL locations)

## FILTER SEMANTICS
- "caged customers"      → filter.caged = true
- "uncaged customers"    → filter.uncaged = true
- "rated customers"      → filter.rated = true (Power Subscription Model = Rated)
- "subscribed customers" → filter.subscribed = true (Power Subscription Model = Subscribed)
- "metered customers"    → filter.metered = true (Power Usage Model = Metered)
- "bundled customers"    → filter.bundled = true (Power Usage Model = Bundled)
- Combine with AND: caged + rated → filter.caged=true AND filter.rated=true
- "all customers" (no qualifier) → all filters null

## FILTER PROPAGATION RULE
If sub-query N is a "list" with a filter, and sub-query N+1 is an "aggregate" on the SAME subject with no explicit filter, inherit the filter from sub-query N.
Example: "list caged customers and sum their capacity in use" →
  op1: type=list, filter.caged=true
  op2: type=aggregate, filter.caged=true (inherited), field_hint="power in use", operation="sum"

## COMPLEX QUERY DECOMPOSITION
Queries joined by "and", "or", commas, semicolons, or "also" = multiple operations in the array.
Execute each independently on the filtered dataset.
- "and" between filters = intersection (BOTH must be true)
- "or" between locations = union (include rows from EITHER location)
- "and" between actions on the same filter = same filter, multiple operation objects

## UNIT AWARENESS (inform the label — do NOT change field_hint)
- Power/capacity columns → KW or KVA (varies per row's UoM column)
- Space columns → Sq Ft
- Rack/seating columns → Racks or Seats
- Revenue columns → ₹/month (MRC)
- Rate columns → ₹/KW-HR
Always include the unit in the label string (e.g., "Total Power Purchased (KW)").

## EXAMPLES

Query: "sum of power purchased"
→ [{"id":"op1","type":"aggregate","label":"Total Power Purchased (KW/KVA)","filter":null,"location":null,"operation":"sum","field_hint":"total capacity purchased","top_n":null,"group_by_location":false}]

Query: "total revenue by location"
→ [{"id":"op1","type":"aggregate","label":"Total Revenue (₹/month) by Location","filter":null,"location":null,"operation":"sum","field_hint":"total revenue","top_n":null,"group_by_location":true}]

Query: "std deviation of per unit rate"
→ [{"id":"op1","type":"aggregate","label":"Std Dev of Per Unit Rate","filter":null,"location":null,"operation":"std","field_hint":"per unit rate","top_n":null,"group_by_location":false}]

Query: "list caged customers in noida"
→ [{"id":"op1","type":"list","label":"Caged Customers — Noida","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":["noida"],"operation":null,"field_hint":null,"top_n":null,"group_by_location":false}]

Query: "list all caged customers AND sum capacity in use AND total power used AND list rated customers AND show customers in airoli or noida"
→ [
  {"id":"op1","type":"list","label":"Caged Customers (All Locations)","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false},
  {"id":"op2","type":"aggregate","label":"Capacity in Use — Caged (KW/KVA)","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"operation":"sum","field_hint":"power in use","top_n":null,"group_by_location":false},
  {"id":"op3","type":"aggregate","label":"Total Power Purchased — Caged (KW/KVA)","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"operation":"sum","field_hint":"total capacity purchased","top_n":null,"group_by_location":false},
  {"id":"op4","type":"list","label":"Rated Customers (All Locations)","filter":{"caged":null,"uncaged":null,"rated":true,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false},
  {"id":"op5","type":"aggregate","label":"Total Power Used — Rated (KW)","filter":{"caged":null,"uncaged":null,"rated":true,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"operation":"sum","field_hint":"power in use","top_n":null,"group_by_location":false},
  {"id":"op6","type":"list","label":"Customers in Airoli or Noida","filter":{"caged":null,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":["airoli","noida"],"operation":null,"field_hint":null,"top_n":null,"group_by_location":false}
]

Query: "how many customers per location"
→ [{"id":"op1","type":"count","label":"Customer Count by Location","filter":null,"location":null,"operation":"count","field_hint":null,"top_n":null,"group_by_location":true}]

Query: "top 5 customers by revenue"
→ [{"id":"op1","type":"aggregate","label":"Top 5 Customers by Revenue","filter":null,"location":null,"operation":"top","field_hint":"total revenue","top_n":5,"group_by_location":false}]

Query: "minimum per unit rate in bangalore"
→ [{"id":"op1","type":"aggregate","label":"Min Per Unit Rate — Bangalore","filter":null,"location":["bangalore"],"operation":"min","field_hint":"per unit rate","top_n":null,"group_by_location":false}]

Query: "median power in use across all locations"
→ [{"id":"op1","type":"aggregate","label":"Median Power In Use (KW/KVA)","filter":null,"location":null,"operation":"median","field_hint":"power in use","top_n":null,"group_by_location":true}]
"""


def _robust_to_numeric(series: pd.Series) -> pd.Series:
    """
    Precisely convert a Series to float64, handling:
      • Comma-formatted numbers  : "1,23,456.78"  → 123456.78
      • Currency prefixes        : "₹ 1,234.56"   → 1234.56
      • Whitespace / % suffix    : "  12.5 %  "   → 12.5
      • Already float/int values : preserved as-is (no repr drift)
      • #DIV, #REF errors        : → NaN
      • Dash / blank / None      : → NaN
    Preserves full float64 precision (uses Decimal internally for string parsing).
    """
    from decimal import Decimal, InvalidOperation

    def _parse(v):
        if v is None:
            return np.nan
        if isinstance(v, bool):
            return np.nan
        if isinstance(v, (int, float)):
            if pd.isna(v):
                return np.nan
            return float(v)           # already numeric — no repr drift
        s = str(v).strip()
        if not s or s in ("-", "–", "—", "nan", "NaN", "None", "N/A",
                          "#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?"):
            return np.nan
        # Strip currency / whitespace
        s = re.sub(r"[₹$£€\s]", "", s)
        s = s.rstrip("%").strip()
        if not s:
            return np.nan
        # Remove commas (handles both Indian 12,34,567 and Western 1,234,567)
        s = s.replace(",", "")
        try:
            return float(Decimal(s))   # Decimal avoids float-parse drift
        except (InvalidOperation, ValueError):
            return np.nan

    return series.apply(_parse)


def _fmt_decimal(val: float, unit: str = "") -> str:
    """
    Format a float for display with appropriate decimal precision:
      • Near-zero difference from int  : show as integer
      • Large revenue values           : 2 dp with comma separators
      • Small rates / ratios           : up to 6 significant digits
    Avoids showing floating-point drift like 881.451999999997.
    """
    if pd.isna(val):
        return "N/A"

    # Round to 10 sig figs to eliminate float drift (e.g. 881.451999999997 → 881.452)
    import math
    if val == 0:
        rounded = 0.0
    else:
        mag = math.floor(math.log10(abs(val)))
        rounded = round(val, max(0, 9 - mag))

    # Decide decimal places
    if rounded == int(rounded) and abs(rounded) < 1e12:
        disp_val = f"{int(rounded):,}"
    elif abs(rounded) >= 10_000:
        disp_val = f"{rounded:,.2f}"
    elif abs(rounded) >= 1:
        disp_val = f"{rounded:,.4f}".rstrip("0").rstrip(".")
    else:
        disp_val = f"{rounded:.6g}"

    return disp_val


def _resolve_col_by_semantic(df: pd.DataFrame, field_hint: str) -> "tuple[str|None, str]":
    """Return (column_name, reason_string) for field_hint — semantic first, fuzzy fallback."""
    hint_lower = (field_hint or "").lower().strip()
    nc = num_cols(df)

    for kw, sem_key in _HINT_SEMANTIC:
        if kw in hint_lower:
            pattern, _ = _SEMANTIC_COLS[sem_key]
            for c in df.columns:
                if re.search(pattern, c, re.I):
                    return c, f"matched '{kw}' → '{sem_key}'"

    hint_words = [w for w in re.split(r"\W+", hint_lower) if len(w) > 2]
    for c in nc:
        if any(w in c.lower() for w in hint_words):
            return c, f"fuzzy word match ({hint_words})"

    if nc:
        return nc[0], "fallback: first numeric column"
    return None, "no numeric column found"


# ── OpenAI client ─────────────────────────────────────────────────────────────

def _get_openai_client():
    replit_base = os.environ.get("AI_INTEGRATIONS_OPENAI_BASE_URL", "")
    replit_key  = os.environ.get("AI_INTEGRATIONS_OPENAI_API_KEY", "")
    std_key     = os.environ.get("OPENAI_API_KEY", "")
    if not std_key:
        try:    std_key = st.secrets.get("OPENAI_API_KEY", "")
        except Exception: pass
    if replit_base and replit_key:
        return _OpenAI(base_url=replit_base, api_key=replit_key)
    if std_key:
        return _OpenAI(api_key=std_key)
    return None


# ── AI Query Parser ───────────────────────────────────────────────────────────

def parse_query_with_ai(query: str) -> "list | tuple":
    import json
    client = _get_openai_client()
    if client is None:
        return ("config_error",
                "No OpenAI API key found. Add OPENAI_API_KEY in Streamlit secrets or environment.")
    try:
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": _AI_PARSER_PROMPT},
                {"role": "user",   "content": query},
            ],
            max_tokens=2048,
            temperature=0,
        )
        raw = resp.choices[0].message.content or "[]"
        raw = re.sub(r"^```(?:json)?\s*", "", raw.strip(), flags=re.I)
        raw = re.sub(r"\s*```$", "", raw.strip())
        ops = json.loads(raw)
        return ops if isinstance(ops, list) else []
    except Exception as e:
        return ("parse_error", str(e))


# ── Filter application ────────────────────────────────────────────────────────

def _apply_op_filters(df: pd.DataFrame,
                      filter_dict: dict,
                      locations: "list | None") -> pd.DataFrame:
    result = df.copy()

    if locations and "_Location" in result.columns:
        loc_mask = pd.Series(False, index=result.index)
        for loc_kw in locations:
            loc_mask |= result["_Location"].str.lower().str.contains(
                re.escape(loc_kw.lower()), na=False)
        result = result[loc_mask]

    if not filter_dict or result.empty:
        return result

    caged_col  = find_col(result, r"\bcaged\b|caged.*uncaged|space.*caged")
    pw_sub_col = find_col(result, r"Power Subscription Model|power.*subscription.*model")
    pw_use_col = find_col(result, r"Power Usage Model|power.*usage.*model")
    own_col    = find_col(result, r"\brhs\b|\bshs\b|ownership")

    if filter_dict.get("caged") and caged_col:
        result = result[result[caged_col].astype(str).str.upper().str.strip() == "CAGED"]
    elif filter_dict.get("uncaged") and caged_col:
        result = result[result[caged_col].astype(str).str.upper().str.strip().isin(
            ["UNCAGED", "UN-CAGED", "UN CAGED"])]

    if filter_dict.get("rated") and pw_sub_col:
        result = result[result[pw_sub_col].astype(str).str.upper().str.contains("RATED", na=False)]
    elif filter_dict.get("subscribed") and pw_sub_col:
        result = result[result[pw_sub_col].astype(str).str.upper().str.contains("SUBSCRIB", na=False)]

    if filter_dict.get("bundled") and pw_use_col:
        result = result[result[pw_use_col].astype(str).str.upper().str.contains("BUNDLED", na=False)]
    elif filter_dict.get("metered") and pw_use_col:
        result = result[result[pw_use_col].astype(str).str.upper().str.contains("METERED", na=False)]

    if filter_dict.get("rhs") and own_col:
        result = result[result[own_col].astype(str).str.upper().str.contains("RHS", na=False)]
    elif filter_dict.get("shs") and own_col:
        result = result[result[own_col].astype(str).str.upper().str.contains("SHS", na=False)]

    return result


# ── Execute AI Operations ─────────────────────────────────────────────────────

_EXTENDED_OPS = {
    "sum":         lambda s: s.sum(),
    "avg":         lambda s: s.mean(),
    "mean":        lambda s: s.mean(),
    "min":         lambda s: s.min(),
    "max":         lambda s: s.max(),
    "count":       lambda s: float(len(s)),
    "std":         lambda s: s.std(ddof=1),
    "median":      lambda s: s.median(),
    "variance":    lambda s: s.var(ddof=1),
    "var":         lambda s: s.var(ddof=1),
    "range":       lambda s: s.max() - s.min(),
    "sum_abs":     lambda s: s.abs().sum(),
    "count_nonzero": lambda s: float((s != 0).sum()),
    "pct_nonzero": lambda s: float((s != 0).sum()) / max(len(s), 1) * 100,
}


def execute_ai_operations(operations: list, df: pd.DataFrame) -> list:
    results = []

    for op in operations:
        op_type     = op.get("type", "list")
        label       = op.get("label", "Result")
        filter_dict = op.get("filter") or {}
        locations   = op.get("location")
        operation   = (op.get("operation") or "sum").lower().strip()
        field_hint  = op.get("field_hint") or ""
        top_n       = int(op.get("top_n") or 10)
        grp_by_loc  = bool(op.get("group_by_location"))

        filtered = _apply_op_filters(df, filter_dict, locations)

        if filtered.empty:
            results.append({"type": "empty", "label": label,
                            "message": "No records match this filter."})
            continue

        # ── LIST ─────────────────────────────────────────────────────────────
        if op_type == "list":
            meta  = [c for c in ["_Location", "_Sheet"] if c in filtered.columns]
            data  = [c for c in filtered.columns if not c.startswith("_")][:30]
            disp  = filtered[meta + data].reset_index(drop=True)
            disp.index += 1
            results.append({"type": "table", "label": label,
                            "data": disp, "row_count": len(disp)})

        # ── AGGREGATE / TOP / BOTTOM ─────────────────────────────────────────
        elif op_type in ("aggregate", "top", "bottom"):
            col, reason = _resolve_col_by_semantic(filtered, field_hint)

            if not col or col not in filtered.columns:
                results.append({"type": "error", "label": label,
                                "message": f"No column matched '{field_hint}'."})
                continue

            unit   = _detect_unit(col)

            # Use _robust_to_numeric to handle all decimal/string/currency formats
            series = _robust_to_numeric(filtered[col])
            valid  = series.dropna()
            total  = len(series)

            # TOP table
            if operation in ("top", "largest"):
                cname = find_col(filtered, r"customer.*name|client.*name")
                extra = [c for c in ["_Location", cname] if c and c in filtered.columns]
                sub   = filtered[extra + [col]].copy()
                sub[col] = _robust_to_numeric(sub[col])
                sub = sub.dropna(subset=[col]).nlargest(top_n, col).reset_index(drop=True)
                sub.index += 1
                results.append({"type": "table", "label": label, "data": sub,
                                "row_count": len(sub), "unit": unit, "column": col,
                                "col_reason": reason})
                continue

            # BOTTOM table
            if operation in ("bottom", "smallest"):
                cname = find_col(filtered, r"customer.*name|client.*name")
                extra = [c for c in ["_Location", cname] if c and c in filtered.columns]
                sub   = filtered[extra + [col]].copy()
                sub[col] = _robust_to_numeric(sub[col])
                sub = sub.dropna(subset=[col]).nsmallest(top_n, col).reset_index(drop=True)
                sub.index += 1
                results.append({"type": "table", "label": label, "data": sub,
                                "row_count": len(sub), "unit": unit, "column": col,
                                "col_reason": reason})
                continue

            if valid.empty:
                results.append({"type": "error", "label": label,
                                "message": f"Column '{col}' has no numeric values."})
                continue

            # Compute with full precision
            op_fn = _EXTENDED_OPS.get(operation)
            if op_fn is None:
                # Try known aliases
                for alias_key in ("sum",):
                    if alias_key in operation:
                        op_fn = _EXTENDED_OPS[alias_key]
                        break
                if op_fn is None:
                    op_fn = _EXTENDED_OPS["sum"]

            val = op_fn(valid)

            # Per-location breakdown using _robust_to_numeric
            loc_breakdown = None
            if grp_by_loc and "_Location" in filtered.columns:
                grp = (
                    filtered.groupby("_Location")[col]
                    .apply(lambda x: _robust_to_numeric(x).sum())
                    .reset_index()
                )
                col_label = f"{col} ({unit})" if unit else col
                grp.columns = ["Location", col_label]
                grp = grp.sort_values(col_label, ascending=False).reset_index(drop=True)
                grp.index += 1
                loc_breakdown = grp

            # Auto per-location breakdown for sum/avg/std/median
            auto_loc = None
            if operation in ("sum", "avg", "mean", "median", "std") \
                    and "_Location" in filtered.columns:
                grp2 = (
                    filtered.groupby("_Location")[col]
                    .apply(lambda x: _robust_to_numeric(x).sum()
                           if operation == "sum" else
                           _robust_to_numeric(x).mean())
                    .reset_index()
                )
                col_lbl2 = f"{col} ({unit})" if unit else col
                grp2.columns = ["Location", col_lbl2]
                grp2 = grp2.sort_values(col_lbl2, ascending=False).reset_index(drop=True)
                grp2.index += 1
                auto_loc = grp2

            results.append({
                "type": "scalar", "label": label,
                "value": val, "unit": unit, "column": col,
                "col_reason": reason,
                "row_count": total, "valid_count": len(valid),
                "operation": operation, "loc_breakdown": loc_breakdown,
                "auto_loc": auto_loc,
            })

        # ── COUNT ─────────────────────────────────────────────────────────────
        elif op_type == "count":
            if grp_by_loc and "_Location" in filtered.columns:
                grp = (filtered.groupby("_Location").size()
                       .reset_index(name="Count")
                       .sort_values("Count", ascending=False)
                       .reset_index(drop=True))
                grp.index += 1
                results.append({"type": "table", "label": label,
                                "data": grp, "row_count": grp["Count"].sum()})
            else:
                results.append({
                    "type": "scalar", "label": label,
                    "value": float(len(filtered)), "unit": "customers",
                    "column": "", "col_reason": "count of filtered rows",
                    "row_count": len(filtered), "valid_count": len(filtered),
                    "operation": "count", "loc_breakdown": None, "auto_loc": None,
                })

    return results


# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────
with st.spinner("Loading all Excel files…"):
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
        st.error("No Excel files found. Place your Excel files in the 'excel_files/' folder.")
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
    st.markdown("<br>", unsafe_allow_html=True)
    for loc in all_locs:
        n = len(ALL.get(loc, {}))
        active = "✓" if loc in sel_locs else "○"
        color = GREEN if loc in sel_locs else MUTED
        st.markdown(
            f'<div style="font-size:.76rem;color:{MUTED};padding:2px 0">'
            f'<span style="color:{color}">{active}</span> {loc} '
            f'<span style="color:{GREEN};font-weight:700">({n} sheets)</span></div>',
            unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# BUILD FILTERED DATA
# ─────────────────────────────────────────────────────────────────────────────
fdata = {loc: {sn: df for sn, df in sheets.items()
               if not sel_sheets or sn in sel_sheets}
         for loc, sheets in ALL.items() if loc in sel_locs}
fdata = {loc: sheets for loc, sheets in fdata.items() if sheets}

COMB = combined_df(fdata)

CUST_frames = []
for loc, sheets in fdata.items():
    for sn, df in sheets.items():
        tmp = df.copy()
        tmp.insert(0, "_Sheet", sn)
        tmp.insert(0, "_Location", loc)
        CUST_frames.append(tmp)

if CUST_frames:
    CUST = pd.concat(CUST_frames, ignore_index=True, sort=False).reset_index(drop=True)
else:
    CUST = pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# HERO HEADER
# ─────────────────────────────────────────────────────────────────────────────
total_rec = len(CUST)
st.markdown(f"""
<div class="hero">
  <h1>🏢 Sify DC — Capacity Intelligence Platform</h1>
  <p>All {len(all_locs)} locations · {sum(len(s) for s in ALL.values())} sheets ·
     {total_rec:,} records loaded</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
T = st.tabs(["📊 KPI Overview", "🗂 Data Explorer", "⚙️ Operations",
             "📈 Charts", "🧠 Smart Query", "🌐 Cross-Location"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 – KPI OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with T[0]:
    st.markdown('<div class="section-title">Key Performance Indicators — All Locations</div>',
                unsafe_allow_html=True)

    if CUST.empty:
        st.warning("No data loaded. Please check your Excel files.")
    else:
        nc = num_cols(CUST)
        tc = txt_cols(CUST)

        cust_c   = find_col(CUST, r"customer.*name|client.*name")
        caged_c  = find_col(CUST, r"\bcaged\b")
        own_c    = find_col(CUST, r"\brhs\b|\bshs\b|ownership")
        sub_mode_c = find_col(CUST, r"subscription.*mode\s*\(rack|space.*subscription.*mode")
        pw_sub_c = find_col(CUST, r"power.*subscription.*model|billing.*model.*power.*subscription")
        pw_use_m_c = find_col(CUST, r"power.*usage.*model|billing.*model.*power.*usage")

        space_sub_c  = find_col(CUST, r"space\s*\|\s*subscription$|^space.*subscription$")
        space_inuse_c = find_col(CUST, r"space.*in.*use")
        space_ytbg_c = find_col(CUST, r"yet.*to.*be.*given|yet.*billed")
        space_res_c  = find_col(CUST, r"reserved.*capacity")
        space_rate_c = find_col(CUST, r"per.*unit.*rate|per.*unit.*mrc")
        rack_c       = find_col(CUST, r"\brack\b")

        cap_c        = find_col(CUST, r"total.*capacity.*purchased|total.*capacity|capacity.*purchased")
        use_c        = find_col(CUST, r"capacity.*in.*use")
        cap_ytbg_c   = find_col(CUST, r"capacity.*to.*be.*given")
        cap_res_c    = find_col(CUST, r"reserved.*capacity")
        sub_kw_c     = find_col(CUST, r"subscribed.*capacity.*kw|capacity.*to.*be.*given.*kw")
        alloc_kw_c   = find_col(CUST, r"allocated.*capacity.*kw|\"allocated\".*kw|allocated.*kw")

        pu_sub_c     = find_col(CUST, r"power.*usage.*subscription|kw.*hr.*subscription")
        pu_inuse_c   = find_col(CUST, r"power.*usage.*in.*use")
        pu_ytbg_c    = find_col(CUST, r"power.*usage.*yet|yet.*to.*be.*given")

        seat_sub_c   = find_col(CUST, r"seating.*space.*subscription|sitting.*space.*subscription|sitting.*space")
        seat_inuse_c = find_col(CUST, r"seating.*space.*in.*use|sitting.*space.*in.*use")

        rev_space_c  = find_col(CUST, r"space.*revenue.*including|space.*revenue")
        rev_addcap_c = find_col(CUST, r"additional.*capacity.*revenue|additional.*capacity.*charge")
        rev_pwuse_c  = find_col(CUST, r"power.*usage.*revenue")
        rev_seat_c   = find_col(CUST, r"seating.*space.*revenue|seating.*revenue")
        rev_other_c  = find_col(CUST, r"any.*other.*items|other.*items")
        rev_total_c  = find_col(CUST, r"total.*revenue")
        rev_freq_c   = find_col(CUST, r"billing.*frequency|frequency")
        rev_so_c     = find_col(CUST, r"sales.*order|so.*ref")
        rev_mrc_c    = find_col(CUST, r"total.*mrc|mrc")

        con_start_c  = find_col(CUST, r"contract.*start|start.*date")
        con_term_c   = find_col(CUST, r"term.*contract|term.*year")
        con_expiry_c = find_col(CUST, r"current.*expiry|expiry.*date|expir")
        con_remarks_c = find_col(CUST, r"remarks")

        def _n(col):
            if col and col in CUST.columns:
                return pd.to_numeric(CUST[col], errors="coerce").sum()
            return None

        def _avg(col):
            if col and col in CUST.columns:
                return pd.to_numeric(CUST[col], errors="coerce").mean()
            return None

        def _cnt_val(col, val):
            if col and col in CUST.columns:
                return int((CUST[col].astype(str).str.upper().str.strip() == val.upper()).sum())
            return None

        k = st.columns(5)
        total_customers = CUST[cust_c].dropna().nunique() if cust_c else len(CUST)
        k[0].markdown(kpi_html(f"{total_customers:,}", "Unique Customers",
                               "across all locations", CYAN), unsafe_allow_html=True)

        if "_Location" in CUST.columns:
            k[1].markdown(kpi_html(f"{CUST['_Location'].nunique()}", "Active Locations",
                                   f"{sum(len(s) for s in fdata.values())} sheets", LBLUE),
                          unsafe_allow_html=True)

        k[2].markdown(kpi_html(f"{len(CUST):,}", "Total Records",
                               "All sheets combined", MUTED), unsafe_allow_html=True)

        if cap_c:
            tot_cap = _n(cap_c)
            k[3].markdown(kpi_html(fmt(tot_cap), "Total Capacity Purchased",
                                   "Power Capacity section", GREEN), unsafe_allow_html=True)

        if use_c:
            tot_use = _n(use_c)
            k[4].markdown(kpi_html(fmt(tot_use), "Capacity In Use",
                                   "Power Capacity section", AMBER), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Billing Model</div>', unsafe_allow_html=True)
        bm_cols = st.columns(4)

        if caged_c:
            cage_vals = CUST[caged_c].astype(str).str.upper().str.strip()
            n_caged   = (cage_vals == "CAGED").sum()
            n_uncaged = (cage_vals == "UNCAGED").sum()
            bm_cols[0].markdown(kpi_html(f"{n_caged}", "Caged",
                                         f"Uncaged: {n_uncaged}", CYAN), unsafe_allow_html=True)

        if own_c:
            rhs_c = _cnt_val(own_c, "RHS")
            shs_c = _cnt_val(own_c, "SHS")
            if rhs_c is not None or shs_c is not None:
                bm_cols[1].markdown(
                    kpi_html(f"{rhs_c or 0}", "RHS",
                             f"SHS: {shs_c or 0}", LBLUE), unsafe_allow_html=True)

        if pw_sub_c:
            rated = _cnt_val(pw_sub_c, "RATED")
            subsc = _cnt_val(pw_sub_c, "SUBSCRIBED")
            bm_cols[2].markdown(
                kpi_html(f"{rated or 0}", "Power Sub: Rated",
                         f"Subscribed: {subsc or 0}", AMBER), unsafe_allow_html=True)

        if pw_use_m_c:
            bundled = _cnt_val(pw_use_m_c, "BUNDLED")
            metered = _cnt_val(pw_use_m_c, "METERED")
            bm_cols[3].markdown(
                kpi_html(f"{bundled or 0}", "Power Usage: Bundled",
                         f"Metered: {metered or 0}", GREEN), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Space</div>', unsafe_allow_html=True)
        sp_cols = st.columns(5)

        if sub_mode_c:
            rack_m = _cnt_val(sub_mode_c, "RACK")
            u_m    = _cnt_val(sub_mode_c, "U SPACE")
            sqft_m = _cnt_val(sub_mode_c, "SQFT SPACE")
            sp_cols[0].markdown(
                kpi_html(f"{rack_m or 0}", "Subscription Mode: Rack",
                         f"U Space: {u_m or 0} | SqFt: {sqft_m or 0}", CYAN),
                unsafe_allow_html=True)

        if space_sub_c:
            v = _n(space_sub_c)
            if v is not None:
                sp_cols[1].markdown(kpi_html(fmt(v), "Space Subscription",
                                             space_sub_c[:25], GREEN), unsafe_allow_html=True)

        if space_inuse_c:
            v = _n(space_inuse_c)
            if v is not None:
                sp_cols[2].markdown(kpi_html(fmt(v), "Space In Use",
                                             space_inuse_c[:25], AMBER), unsafe_allow_html=True)

        if space_ytbg_c:
            v = _n(space_ytbg_c)
            if v is not None:
                sp_cols[3].markdown(kpi_html(fmt(v), "Yet To Be Given / Billed",
                                             space_ytbg_c[:25], RED), unsafe_allow_html=True)

        if space_rate_c:
            v = _avg(space_rate_c)
            if v is not None:
                sp_cols[4].markdown(kpi_html(fmt(v), "Avg Per Unit Rate (MRC)",
                                             space_rate_c[:25], LBLUE), unsafe_allow_html=True)
        elif rack_c:
            v = _n(rack_c)
            if v is not None:
                sp_cols[4].markdown(kpi_html(fmt(v), "Total Racks / Space",
                                             rack_c[:25], LBLUE), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Power Capacity</div>', unsafe_allow_html=True)
        pc_cols = st.columns(5)

        if cap_c:
            pc_cols[0].markdown(kpi_html(fmt(_n(cap_c)), "Total Capacity Purchased",
                                         cap_c[:25], GREEN), unsafe_allow_html=True)
        if use_c:
            pc_cols[1].markdown(kpi_html(fmt(_n(use_c)), "Capacity In Use",
                                         use_c[:25], AMBER), unsafe_allow_html=True)
        if cap_ytbg_c:
            v = _n(cap_ytbg_c)
            if v is not None:
                pc_cols[2].markdown(kpi_html(fmt(v), "Capacity To Be Given",
                                             cap_ytbg_c[:25], RED), unsafe_allow_html=True)
        if sub_kw_c:
            v = _n(sub_kw_c)
            if v is not None:
                pc_cols[3].markdown(kpi_html(fmt(v), "Subscribed Cap. To Give (KW)",
                                             sub_kw_c[:25], LBLUE), unsafe_allow_html=True)
        if alloc_kw_c:
            v = _n(alloc_kw_c)
            if v is not None:
                pc_cols[4].markdown(kpi_html(fmt(v), "Allocated Capacity KW",
                                             alloc_kw_c[:25], CYAN), unsafe_allow_html=True)
        elif cap_c and use_c:
            t_cap = _n(cap_c) or 0
            t_use = _n(use_c) or 0
            util  = (t_use / t_cap * 100) if t_cap > 0 else 0
            pc_cols[4].markdown(kpi_html(f"{util:.1f}%", "Utilisation Rate",
                                         "Capacity In Use / Purchased", AMBER),
                                unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        show_pu = pu_sub_c or pu_inuse_c or pu_ytbg_c
        if show_pu:
            st.markdown('<div class="section-title">Power Usage</div>', unsafe_allow_html=True)
            pu_cols = st.columns(4)
            i = 0
            for col, label, color in [
                (pu_sub_c,   "KW-HR/Month Subscription", GREEN),
                (pu_inuse_c, "Power Usage In Use",        AMBER),
                (pu_ytbg_c,  "Power Usage Yet To Give",   RED),
            ]:
                if col:
                    v = _n(col)
                    if v is not None:
                        pu_cols[i].markdown(kpi_html(fmt(v), label, col[:25], color),
                                            unsafe_allow_html=True)
                        i += 1
            st.markdown("<br>", unsafe_allow_html=True)

        show_seat = seat_sub_c or seat_inuse_c
        if show_seat:
            st.markdown('<div class="section-title">Seating Space</div>', unsafe_allow_html=True)
            ss_cols = st.columns(3)
            if seat_sub_c:
                v = _n(seat_sub_c)
                if v is not None:
                    ss_cols[0].markdown(kpi_html(fmt(v), "Sitting Space Subscription",
                                                 seat_sub_c[:25], CYAN), unsafe_allow_html=True)
            if seat_inuse_c:
                v = _n(seat_inuse_c)
                if v is not None:
                    ss_cols[1].markdown(kpi_html(fmt(v), "Sitting Space In Use",
                                                 seat_inuse_c[:25], AMBER), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        rev_c = rev_total_c or rev_mrc_c
        st.markdown('<div class="section-title">Revenue (Monthly)</div>', unsafe_allow_html=True)
        rv_cols = st.columns(5)
        rv_items = [
            (rev_space_c,  "Space Revenue (incl. Capacity)", CYAN),
            (rev_addcap_c, "Additional Capacity Revenue",    LBLUE),
            (rev_pwuse_c,  "Power Usage Revenue",            GREEN),
            (rev_seat_c,   "Seating Space Revenue",          AMBER),
            (rev_other_c,  "Any Other Items",                MUTED),
        ]
        filled = 0
        for col, label, color in rv_items:
            if col and filled < 5:
                v = _n(col)
                if v is not None:
                    rv_cols[filled].markdown(kpi_html(fmt(v), label, col[:25], color),
                                             unsafe_allow_html=True)
                    filled += 1

        st.markdown("<br>", unsafe_allow_html=True)
        rv2_cols = st.columns(4)
        rv2_items = [
            (rev_total_c,  "Total Revenue",     GREEN),
            (rev_mrc_c,    "Total MRC",         LBLUE),
        ]
        filled2 = 0
        for col, label, color in rv2_items:
            if col and filled2 < 4:
                v = _n(col)
                if v is not None:
                    rv2_cols[filled2].markdown(kpi_html(fmt(v), label, col[:25], color),
                                               unsafe_allow_html=True)
                    filled2 += 1

        if rev_freq_c:
            freq_counts = CUST[rev_freq_c].dropna().value_counts()
            top_freq = freq_counts.index[0] if not freq_counts.empty else "—"
            rv2_cols[min(filled2, 3)].markdown(
                kpi_html(str(top_freq), "Top Billing Frequency",
                         f"{len(freq_counts)} types", AMBER), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Contract Information</div>', unsafe_allow_html=True)
        ci_cols = st.columns(4)
        ci_i = 0

        if con_start_c:
            non_null = CUST[con_start_c].dropna()
            ci_cols[ci_i].markdown(
                kpi_html(f"{len(non_null):,}", "Contracts With Start Date",
                         con_start_c[:25], CYAN), unsafe_allow_html=True)
            ci_i += 1

        if con_term_c:
            v = _avg(con_term_c)
            if v is not None:
                ci_cols[ci_i].markdown(kpi_html(f"{v:.1f} yr", "Avg Contract Term",
                                                con_term_c[:25], GREEN), unsafe_allow_html=True)
                ci_i += 1

        if con_expiry_c:
            non_null = CUST[con_expiry_c].dropna()
            ci_cols[ci_i].markdown(
                kpi_html(f"{len(non_null):,}", "Contracts With Expiry Date",
                         con_expiry_c[:25], AMBER), unsafe_allow_html=True)
            ci_i += 1

        if rev_so_c:
            so_count = CUST[rev_so_c].dropna().nunique()
            ci_cols[min(ci_i, 3)].markdown(
                kpi_html(f"{so_count:,}", "Unique Sales Orders",
                         rev_so_c[:25], LBLUE), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Per-Location Summary</div>', unsafe_allow_html=True)
        if "_Location" in CUST.columns:
            agg_cols = [c for c in [cap_c, use_c, rev_total_c or rev_mrc_c,
                                    rev_space_c, rev_pwuse_c] if c]
            if agg_cols:
                loc_agg = CUST.groupby("_Location")[agg_cols].apply(
                    lambda g: g.apply(pd.to_numeric, errors="coerce").sum()
                ).reset_index()
                loc_agg.columns = ["Location"] + agg_cols
                if cust_c:
                    loc_agg["Customer Count"] = (
                        CUST.groupby("_Location")[cust_c]
                        .apply(lambda g: g.dropna().nunique()).values)
                else:
                    loc_agg["Customer Count"] = (
                        CUST.groupby("_Location").size().values)
                st.dataframe(loc_agg.round(2), use_container_width=True)
            else:
                lc = CUST["_Location"].value_counts().reset_index()
                lc.columns = ["Location", "Records"]
                st.dataframe(lc, use_container_width=True)

        if cap_c and use_c:
            st.markdown('<div class="section-title">Utilisation Gauges</div>',
                        unsafe_allow_html=True)
            g1, g2 = st.columns(2)
            t_cap = _n(cap_c) or 0
            t_use = _n(use_c) or 0
            util_pct = min((t_use / t_cap * 100) if t_cap > 0 else 0, 100)
            rack_pct = util_pct
            if rack_c:
                t_rack = _n(rack_c) or 0
                rack_pct = min((t_use / t_rack * 100) if t_rack > 0 else util_pct, 100)

            def _gauge(val, label, bar_color):
                fig = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=min(float(val), 100),
                    title={"text": label, "font": {"color": TEXT, "size": 14}},
                    gauge={
                        "axis": {"range": [0, 100], "tickcolor": TEXT},
                        "bar":  {"color": bar_color},
                        "bgcolor": DARK2,
                        "steps": [
                            {"range": [0,  50], "color": "#1a2a1a"},
                            {"range": [50, 80], "color": "#2a2a1a"},
                            {"range": [80, 100], "color": "#2a1a1a"},
                        ],
                        "threshold": {"line": {"color": RED, "width": 3}, "value": 80},
                    },
                    number={"suffix": "%", "font": {"color": bar_color}},
                ))
                fig.update_layout(**_base_layout(), height=270)
                return fig

            g1.plotly_chart(_gauge(util_pct, "Capacity Utilisation (%)", LBLUE),
                            use_container_width=True)
            g2.plotly_chart(_gauge(rack_pct, "Space/Rack Utilisation (%)", GREEN),
                            use_container_width=True)

        if cap_c and use_c and "_Location" in CUST.columns:
            st.markdown('<div class="section-title">Capacity vs Usage by Location</div>',
                        unsafe_allow_html=True)
            la = CUST.groupby("_Location").agg(
                Capacity_Purchased=(cap_c, lambda x: pd.to_numeric(x, errors="coerce").sum()),
                Capacity_in_Use   =(use_c, lambda x: pd.to_numeric(x, errors="coerce").sum()),
            ).reset_index()
            fig_la = px.bar(la, x="_Location",
                            y=["Capacity_Purchased", "Capacity_in_Use"],
                            barmode="group",
                            labels={"_Location": "Location", "value": "Units"},
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

        st.markdown('<div class="section-title">Space &amp; Revenue Split</div>',
                    unsafe_allow_html=True)
        pie_cols = st.columns(3)

        if caged_c:
            cv = CUST[caged_c].astype(str).str.upper().str.strip()
            pie_d = cv.value_counts().reset_index()
            pie_d.columns = ["Status", "Count"]
            if not pie_d.empty:
                fig_p1 = px.pie(pie_d, names="Status", values="Count",
                                title="Caged vs Uncaged",
                                color_discrete_sequence=[CYAN, LBLUE, GREEN, AMBER])
                fig_p1.update_layout(**_base_layout(), height=300)
                pie_cols[0].plotly_chart(fig_p1, use_container_width=True)

        if pw_sub_c:
            pie_d2 = CUST[pw_sub_c].dropna().value_counts().reset_index()
            pie_d2.columns = ["Model", "Count"]
            if not pie_d2.empty:
                fig_p2 = px.pie(pie_d2, names="Model", values="Count",
                                title="Power Subscription Model",
                                color_discrete_sequence=[LBLUE, GREEN, AMBER, RED])
                fig_p2.update_layout(**_base_layout(), height=300)
                pie_cols[1].plotly_chart(fig_p2, use_container_width=True)

        if pw_use_m_c:
            pie_d3 = CUST[pw_use_m_c].dropna().value_counts().reset_index()
            pie_d3.columns = ["Model", "Count"]
            if not pie_d3.empty:
                fig_p3 = px.pie(pie_d3, names="Model", values="Count",
                                title="Power Usage Model",
                                color_discrete_sequence=[GREEN, AMBER, CYAN, RED])
                fig_p3.update_layout(**_base_layout(), height=300)
                pie_cols[2].plotly_chart(fig_p3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – DATA EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
with T[1]:
    st.markdown('<div class="section-title">Data Explorer</div>', unsafe_allow_html=True)

    if CUST.empty:
        st.warning("No data loaded.")
    else:
        de1, de2, de3 = st.columns(3)
        with de1:
            de_loc = st.selectbox("📍 Location", ["All"] + sorted(fdata.keys()), key="de_loc")
        with de2:
            if de_loc != "All":
                sheet_opts = sorted(fdata.get(de_loc, {}).keys())
            else:
                sheet_opts = sorted({sn for sheets in fdata.values() for sn in sheets})
            de_sh = st.selectbox("📋 Sheet", ["All"] + sheet_opts, key="de_sh")
        with de3:
            de_search = st.text_input("🔍 Search (any column)", key="de_search", placeholder="type to filter…")

        view_df = CUST.copy()
        if de_loc != "All" and "_Location" in view_df.columns:
            view_df = view_df[view_df["_Location"] == de_loc]
        if de_sh != "All" and "_Sheet" in view_df.columns:
            view_df = view_df[view_df["_Sheet"] == de_sh]
        if de_search.strip():
            mask = view_df.apply(
                lambda r: r.astype(str).str.lower().str.contains(
                    de_search.lower(), na=False).any(), axis=1)
            view_df = view_df[mask]

        st.markdown(
            f'<span class="badge">{len(view_df):,} rows</span> '
            f'<span class="badge" style="background:{DARK2}">{len(view_df.columns)} cols</span>',
            unsafe_allow_html=True)
        st.dataframe(view_df.head(1000), use_container_width=True, height=480)

        dl_cols = st.columns(3)
        dl_cols[0].download_button(
            "⬇ Download CSV", view_df.to_csv(index=False).encode(),
            "sify_data.csv", "text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – OPERATIONS
# ══════════════════════════════════════════════════════════════════════════════
with T[2]:
    st.markdown('<div class="section-title">Operations Engine</div>', unsafe_allow_html=True)

    if CUST.empty:
        st.warning("No data loaded.")
    else:
        op1, op2, op3, op4 = st.columns(4)
        with op1:
            op_loc = st.selectbox("📍 Location", ["All"] + sorted(fdata.keys()), key="op_loc")
        with op2:
            op_sh_opts = (sorted(fdata.get(op_loc, {}).keys())
                          if op_loc != "All" else sorted({sn for s in fdata.values() for sn in s}))
            op_sh = st.selectbox("📋 Sheet", ["All"] + op_sh_opts, key="op_sh")

        op_df = CUST.copy()
        if op_loc != "All" and "_Location" in op_df.columns:
            op_df = op_df[op_df["_Location"] == op_loc]
        if op_sh != "All" and "_Sheet" in op_df.columns:
            op_df = op_df[op_df["_Sheet"] == op_sh]

        nc_op = num_cols(op_df)
        tc_op = txt_cols(op_df)

        with op3:
            op_col = st.selectbox("📐 Column", nc_op if nc_op else ["—"], key="op_col")
        with op4:
            op_op = st.selectbox("🔧 Operation", OPERATIONS, key="op_op")

        op5, op6 = st.columns(2)
        with op5:
            op_grp = st.selectbox("🗂 Group By (optional)",
                                  ["None"] + [c for c in tc_op if not c.startswith("_")],
                                  key="op_grp")
        with op6:
            op_n = st.number_input("N (Top/Bottom N)", min_value=1, max_value=100,
                                   value=10, step=1, key="op_n")

        if st.button("▶ Run Operation", key="op_run") and op_col and op_col != "—":
            grp = op_grp if op_grp != "None" else None
            res, desc = run_op(op_df, op_col, op_op, grp, int(op_n))
            st.markdown(f'<div class="section-title">{desc}</div>', unsafe_allow_html=True)
            if isinstance(res, pd.DataFrame):
                st.dataframe(res, use_container_width=True)
                if grp:
                    fig_op = px.bar(res.head(30), x=grp, y=op_col, color=op_col,
                                    color_continuous_scale="Blues", title=desc)
                    fig_op.update_layout(**_base_layout(), height=350)
                    st.plotly_chart(fig_op, use_container_width=True)
            elif isinstance(res, (int, float)):
                st.markdown(f"""
                <div style="background:{DARK2};border:1px solid {BORD};border-radius:14px;
                     padding:28px;text-align:center;margin:12px 0">
                  <div style="font-size:.8rem;color:{MUTED};text-transform:uppercase;
                       font-weight:700;letter-spacing:.08em">{desc}</div>
                  <div class="result-big">{fmt(res)}</div>
                </div>""", unsafe_allow_html=True)
            else:
                st.info(str(res))


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – CHARTS
# ══════════════════════════════════════════════════════════════════════════════
with T[3]:
    st.markdown('<div class="section-title">Chart Studio</div>', unsafe_allow_html=True)

    if CUST.empty:
        st.warning("No data loaded.")
    else:
        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            ch_loc = st.selectbox("📍 Location", ["All"] + sorted(fdata.keys()), key="ch_loc")
        with ch2:
            ch_sh_opts = (sorted(fdata.get(ch_loc, {}).keys())
                          if ch_loc != "All" else sorted({sn for s in fdata.values() for sn in s}))
            ch_sh = st.selectbox("📋 Sheet", ["All"] + ch_sh_opts, key="ch_sh")
        with ch3:
            ch_type = st.selectbox("📊 Chart Type", CHART_TYPES, key="ch_type")

        st.caption(CHART_DESC.get(ch_type, ""))

        ch_df = CUST.copy()
        if ch_loc != "All" and "_Location" in ch_df.columns:
            ch_df = ch_df[ch_df["_Location"] == ch_loc]
        if ch_sh != "All" and "_Sheet" in ch_df.columns:
            ch_df = ch_df[ch_df["_Sheet"] == ch_sh]

        nc_ch = num_cols(ch_df)
        tc_ch = txt_cols(ch_df)

        needs = CHART_NEEDS.get(ch_type, set())
        ca, cb, cc, cd = st.columns(4)
        x_val = ca.selectbox("X-axis / Category",
                              ["—"] + tc_ch + nc_ch, key="ch_x") if "x_cat" in needs or "x_num" in needs else None
        y_val = cb.selectbox("Y-axis / Value",
                              ["—"] + nc_ch, key="ch_y") if "y_num" in needs else None
        col_val = cc.selectbox("Color",
                               ["—"] + tc_ch + nc_ch, key="ch_col") if "color" in needs else None
        sz_val = cd.selectbox("Size",
                              ["—"] + nc_ch, key="ch_sz") if "size" in needs else None
        z_val = ca.selectbox("Z-axis",
                             ["—"] + nc_ch, key="ch_z") if "z_num" in needs else None

        if st.button("🎨 Generate Chart", key="ch_run"):
            kw = dict(
                x=x_val if x_val and x_val != "—" else None,
                y=y_val if y_val and y_val != "—" else None,
                color=col_val if col_val and col_val != "—" else None,
                size=sz_val if sz_val and sz_val != "—" else None,
                z=z_val if z_val and z_val != "—" else None,
                title=f"{ch_type} — {ch_loc} / {ch_sh}",
            )
            fig = make_chart(ch_type, ch_df, **kw)
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – SMART QUERY  (Structured AI parse → real data execution)
# ══════════════════════════════════════════════════════════════════════════════
with T[4]:
    st.markdown('<div class="section-title">🧠 Smart Query — AI-Powered Structured Query Engine</div>',
                unsafe_allow_html=True)

    # ── Data source ───────────────────────────────────────────────────────────
    sq_src_opts = ["All Locations & All Sheets"] + sorted(fdata.keys())
    sq_src = st.selectbox("📂 Query data source", sq_src_opts, key="sq_src")

    if sq_src == "All Locations & All Sheets":
        pool_base = CUST.copy()
    else:
        loc_frames = []
        for sn, df_loc in fdata.get(sq_src, {}).items():
            tmp = df_loc.copy()
            tmp.insert(0, "_Sheet", sn)
            tmp.insert(0, "_Location", sq_src)
            loc_frames.append(tmp)
        pool_base = pd.concat(loc_frames, ignore_index=True, sort=False) if loc_frames else pd.DataFrame()

    if not pool_base.empty:
        n_locs   = pool_base["_Location"].nunique() if "_Location" in pool_base.columns else 1
        n_sheets = pool_base["_Sheet"].nunique()    if "_Sheet"    in pool_base.columns else 1
        st.markdown(
            f'<div style="font-size:.78rem;color:{MUTED};margin-bottom:10px">'
            f'Query pool: <b style="color:{CYAN}">{len(pool_base):,}</b> records · '
            f'<b style="color:{CYAN}">{n_locs}</b> location(s) · '
            f'<b style="color:{CYAN}">{n_sheets}</b> sheet(s)</div>',
            unsafe_allow_html=True)

    # ── Query input ───────────────────────────────────────────────────────────
    query = st.text_area(
        "🔍 Enter your query",
        placeholder=(
            "e.g. List all caged customers AND sum capacity in use AND total power used "
            "AND list rated customers AND show customers in airoli or noida"
        ),
        key="sq_q",
        height=90,
    )

    sq_locs = st.multiselect(
        "📍 Restrict to locations (optional)",
        options=sorted(fdata.keys()),
        default=[],
        key="sq_locs"
    )

    c_run, _ = st.columns([1, 6])
    with c_run:
        run_clicked = st.button("🚀 Run Query", key="sq_run")

    if "sq_results_history" not in st.session_state:
        st.session_state["sq_results_history"] = []

    # ── Run ───────────────────────────────────────────────────────────────────
    if run_clicked and query.strip():
        if pool_base.empty:
            st.error("No data loaded. Please check your Excel files.")
        else:
            pool = pool_base.copy()
            if sq_locs and "_Location" in pool.columns:
                pool = pool[pool["_Location"].isin(sq_locs)]
            if pool.empty:
                st.warning("No records for selected locations.")
            else:
                with st.spinner("🤖 Parsing and executing query…"):
                    ops_raw = parse_query_with_ai(query.strip())

                if isinstance(ops_raw, tuple):
                    err_type, err_msg = ops_raw
                    st.error(f"**{'Config' if err_type == 'config_error' else 'Parse'} Error**: {err_msg}")
                elif not ops_raw:
                    st.warning("Could not parse query. Please try rephrasing.")
                else:
                    results = execute_ai_operations(ops_raw, pool)
                    st.session_state["sq_results_history"].append({
                        "query":   query.strip(),
                        "source":  sq_src,
                        "records": len(pool),
                        "results": results,
                    })

    # ── Display ───────────────────────────────────────────────────────────────
    if st.session_state.get("sq_results_history"):
        for hist in reversed(st.session_state["sq_results_history"]):
            # Query bubble
            st.markdown(f"""
            <div style="background:{DARK2};border:1px solid {BORD};border-radius:12px;
                 padding:14px 18px;margin:14px 0 8px;display:flex;align-items:flex-start;gap:12px">
              <div style="font-size:1.1rem;min-width:28px">❓</div>
              <div>
                <div style="font-size:.72rem;color:{MUTED};font-weight:700;
                     text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px">
                  {hist['source']} · {hist['records']:,} records</div>
                <div style="font-size:.95rem;color:{WHITE};font-weight:600">{hist['query']}</div>
              </div>
            </div>""", unsafe_allow_html=True)

            # Results
            for res in hist["results"]:
                rtype = res["type"]
                label = res["label"]

                # ── Table result ──────────────────────────────────────────────
                if rtype == "table":
                    st.markdown(
                        f'<div style="font-size:.8rem;color:{CYAN};font-weight:700;'
                        f'text-transform:uppercase;letter-spacing:.06em;margin:12px 0 4px">'
                        f'📋 {label} &nbsp;<span style="color:{MUTED};font-weight:400;'
                        f'font-size:.75rem">— {res["row_count"]:,} row(s)</span></div>',
                        unsafe_allow_html=True)
                    st.dataframe(res["data"], use_container_width=True)
                    csv_bytes = res["data"].to_csv(index=True).encode("utf-8")
                    st.download_button(
                        f"⬇️ Download {label} (CSV)",
                        data=csv_bytes,
                        file_name=f"{label.replace(' ','_')}.csv",
                        mime="text/csv",
                        key=f"dl_{label}_{id(res)}"
                    )

                # ── Scalar / metric result ────────────────────────────────────
                elif rtype == "scalar":
                    unit        = res.get("unit", "")
                    val         = res["value"]
                    col_n       = res.get("column", "")
                    col_reason  = res.get("col_reason", "")
                    valid_count = res.get("valid_count", res["row_count"])
                    total_count = res["row_count"]
                    operation   = (res.get("operation") or "sum").upper()
                    pct_valid   = (valid_count / total_count * 100) if total_count else 0

                    # Format value — use _fmt_decimal to avoid float drift
                    if unit == "₹":
                        import math as _math
                        _rval = round(val, max(0, 9 - (int(_math.floor(_math.log10(abs(val)))) if val else 0)))
                        if _rval >= 1_00_00_000:
                            disp = f"₹ {_rval/1_00_00_000:,.2f} Cr"
                        elif _rval >= 1_00_000:
                            disp = f"₹ {_rval/1_00_000:,.2f} L"
                        else:
                            disp = f"₹ {_rval:,.2f}"
                    elif unit == "customers":
                        disp = f"{int(val):,}"
                        unit = "customers"
                    else:
                        _raw = _fmt_decimal(val, unit)
                        disp = f"{_raw} {unit}".strip() if unit else _raw

                    warn_color = AMBER if pct_valid < 50 else GREEN

                    st.markdown(f"""
                    <div style="background:{DARK2};border:1px solid {BORD};border-radius:14px;
                         padding:20px 28px;margin:10px 0">
                      <div style="font-size:.72rem;color:{MUTED};font-weight:700;
                           text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px">
                        {operation} · {label}</div>
                      <div style="font-size:2.4rem;font-weight:900;color:{CYAN};
                           letter-spacing:-.01em;line-height:1.1">{disp}
                        <span style="font-size:1rem;color:{MUTED};font-weight:500"> {unit}</span>
                      </div>
                      <div style="font-size:.74rem;color:{MUTED};margin-top:8px;
                           border-top:1px solid {BORD};padding-top:8px">
                        📊 Column used: <b style="color:{TEXT}">{col_n}</b><br>
                        ✅ <b style="color:{warn_color}">{valid_count:,}</b> of
                        <b style="color:{TEXT}">{total_count:,}</b> rows had numeric values
                        ({pct_valid:.0f}%)
                        {"⚠️ Many blank/text values — sum may be partial" if pct_valid < 50 else ""}
                      </div>
                    </div>""", unsafe_allow_html=True)

                    # Auto per-location table
                    auto_loc = res.get("auto_loc")
                    if auto_loc is not None and not auto_loc.empty:
                        with st.expander("📍 Per-location breakdown", expanded=True):
                            st.dataframe(auto_loc, use_container_width=True)

                elif rtype == "empty":
                    st.info(f"**{label}**: {res['message']}")
                elif rtype == "error":
                    st.warning(f"**{label}**: {res['message']}")

            st.markdown(f"<hr style='border-color:{BORD};margin:10px 0 16px'>",
                        unsafe_allow_html=True)

        if st.button("🗑 Clear Results", key="sq_clear"):
            st.session_state["sq_results_history"] = []
            st.rerun()
    else:
        st.markdown(f"""
        <div style="background:{DARK2};border:1px dashed {BORD};border-radius:12px;
             padding:32px;text-align:center;color:{MUTED};margin-top:16px">
          <div style="font-size:2rem;margin-bottom:10px">🤖</div>
          <div style="font-size:.95rem;color:{TEXT}">
            Enter a query above and click <b style="color:{CYAN}">Run Query</b>
          </div>
          <div style="font-size:.82rem;margin-top:10px">
            Examples:<br>
            <i>sum of power</i> &nbsp;·&nbsp;
            <i>total space used</i> &nbsp;·&nbsp;
            <i>total revenue</i> &nbsp;·&nbsp;
            <i>list all caged customers AND sum capacity in use AND total revenue by location</i>
          </div>
        </div>""", unsafe_allow_html=True)




# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – CROSS-LOCATION
# ══════════════════════════════════════════════════════════════════════════════
with T[5]:
    st.markdown('<div class="section-title">Cross-Location Comparison</div>',
                unsafe_allow_html=True)

    nc_all = num_cols(CUST)
    if not nc_all:
        st.info("No numeric columns found.")
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
            loc_df = CUST[CUST["_Location"] == loc] if "_Location" in CUST.columns else CUST
            if not loc_df.empty and xl_col in loc_df.columns:
                val, _ = run_op(loc_df, xl_col, xl_op)
                if isinstance(val, (int, float)):
                    rows.append({"Location": loc, xl_col: val})
        xl_agg = (pd.DataFrame(rows).sort_values(xl_col, ascending=False)
                  if rows else pd.DataFrame())

        if not xl_agg.empty:
            k1, k2, k3 = st.columns(3)
            k1.metric("🏆 Highest", xl_agg.iloc[0]["Location"],  fmt(xl_agg.iloc[0][xl_col]))
            k2.metric("📉 Lowest",  xl_agg.iloc[-1]["Location"], fmt(xl_agg.iloc[-1][xl_col]))
            k3.metric("Σ Network",  "", fmt(xl_agg[xl_col].sum()))

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
                                    title=f"{xl_op} of {xl_col} — All Locations")
                fig_xl.update_layout(height=420)
            st.plotly_chart(fig_xl, use_container_width=True)

            st.markdown('<div class="section-title">Summary Table</div>', unsafe_allow_html=True)
            st.dataframe(xl_agg.round(3), use_container_width=True)
            st.download_button("⬇ Download CSV",
                               xl_agg.to_csv(index=False).encode(),
                               "cross_location.csv", "text/csv")
        else:
            st.info("No numeric data available for selected metric and locations.")
