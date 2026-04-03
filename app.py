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
    """Return list of directories to search for Excel files, in priority order."""
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

# Words that strongly indicate a row IS a column-header row
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
    """Detect if a row looks like a column-header row."""
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
    """
    Return (data_start_row_1based, group_header_row, col_header_row).

    Scans the first 10 rows looking for:
      - A group-header row (contains SECTION_MARKERS)
      - Immediately followed by a column-header row (contains HEADER_INDICATORS)
    Data starts on the row after the column-header row.

    Falls back to single-row header detection, then row-1 as header.
    """
    def rs(r):
        return [str(v).strip() if v is not None else "" for v in r]

    rows = [rs(r) for r in raw_rows[:10]]

    # Try two-row header: section row then col-header row
    for i in range(min(5, len(rows) - 1)):
        r1 = rows[i]
        r2 = rows[i + 1] if i + 1 < len(rows) else []
        if r2 and (_is_section_row(r1) or _is_header_row(r1)) and _is_header_row(r2):
            data_start = i + 3  # 1-based: rows i+1 and i+2 are headers, data at i+3
            return data_start, r1, r2

    # Try single-row header (col-header only, no preceding section row)
    for i in range(min(5, len(rows))):
        r = rows[i]
        if _is_header_row(r):
            data_start = i + 2  # 1-based: row i+1 is header, data at i+2
            return data_start, [""] * len(r), r

    # Last resort: treat row 0 as header
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
            # Only convert if >50% of values are numeric AND the column doesn't look like
            # a text/name column based on its name
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


@st.cache_data(show_spinner=False)
def load_all() -> dict:
    """Return {location_label: {sheet_name: DataFrame}} for every Excel file found."""
    result: dict = {}
    used_labels: set = set()

    def _safe_label(fpath: Path) -> str:
        base = _label(fpath)
        if base not in used_labels:
            return base
        full = re.sub(r"Customer_and_Capacity_Tracker_", "", fpath.stem, flags=re.I)
        full = re.sub(r"_\d{10,}$", "", full).replace("_", " ").strip()
        return full

    dirs = _excel_dirs()
    seen_files: set = set()

    for d in dirs:
        for fpath in sorted(d.glob("*.xlsx")):
            if fpath.name in seen_files:
                continue
            seen_files.add(fpath.name)
            label = _safe_label(fpath)
            used_labels.add(label)
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

        for fpath in sorted(d.glob("*.xls")):
            if fpath.suffix.lower() != ".xls":
                continue
            if fpath.name in seen_files:
                continue
            seen_files.add(fpath.name)
            label = _safe_label(fpath)
            used_labels.add(label)
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
# CHART FACTORY – 15 chart types
# ─────────────────────────────────────────────────────────────────────────────
CHART_TYPES = [
    "Bar Chart", "Grouped Bar", "Stacked Bar",
    "Line Chart", "Scatter Plot", "Area Chart",
    "Bubble Chart", "Heatmap (Correlation)", "Box Plot",
    "Violin Plot", "Funnel Chart", "Waterfall / Cumulative",
    "3-D Scatter", "Radar Chart", "Histogram",
]

CHART_DESC = {
    "Bar Chart":              "Compare a numeric metric across categorical groups (e.g., power per customer).",
    "Grouped Bar":            "Side-by-side comparison of multiple numeric columns across groups.",
    "Stacked Bar":            "Show composition and total simultaneously across groups.",
    "Line Chart":             "Trend analysis across ordered rows or time-series data.",
    "Scatter Plot":           "Correlation between two numeric variables; colour-coded by a third.",
    "Area Chart":             "Cumulative volume trends with filled area for visual emphasis.",
    "Bubble Chart":           "Three-dimensional numeric relationships (X, Y, size).",
    "Heatmap (Correlation)":  "Instantly spot which numeric columns are correlated.",
    "Box Plot":               "Distribution, spread, median, and outliers for numeric columns.",
    "Violin Plot":            "Full probability distribution shape for numeric columns.",
    "Funnel Chart":           "Staged capacity utilisation or sales-pipeline visualisation.",
    "Waterfall / Cumulative": "Running total analysis, e.g., cumulative power consumed.",
    "3-D Scatter":            "Three-axis numeric exploration for high-dimensional data.",
    "Radar Chart":            "Multi-axis comparison of a single entity across metrics.",
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
# SMART QUERY ENGINE – Full AND/OR compound query parser
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

# Column concept keywords that identify WHAT numeric column to operate on
# These must NOT be used for row-level text filtering
_COL_CONCEPT_WORDS: set = {
    "power", "kw", "kilowatt", "kva", "capacity", "purchased", "subscribed",
    "usage", "use", "consumption", "revenue", "mrc", "billing", "charge",
    "rack", "racks", "space", "area", "floor", "sitting", "seat",
    "rate", "unit", "contract", "term", "expiry", "frequency",
}

_COL_CONCEPTS: list = [
    (["power", "kw", "kilowatt", "kva"],                 r"total.*capacity.*kw|capacity.*kw|power|kw|kilowatt|kva"),
    (["capacity purchased", "total capacity"],            r"total.*capacity|capacity.*purchased"),
    (["capacity in use", "used capacity", "in use"],      r"capacity.*in.*use|in use"),
    (["capacity", "purchased", "subscribed"],             r"capacity|purchased|subscribed"),
    (["usage", "use", "consumption"],                     r"usage|in use|consumption"),
    (["revenue", "mrc", "billing", "charge"],             r"revenue|mrc|billing|charge"),
    (["rack", "racks"],                                   r"rack"),
    (["customer", "client", "name"],                      r"customer.*name|client.*name"),
    (["space", "area"],                                   r"space|area"),
    (["sitting", "seat"],                                 r"sit|seat"),
    (["rate", "per unit"],                                r"per.*unit|rate"),
    (["contract", "term", "expiry"],                      r"contract|term|expir"),
    (["frequency", "billing frequency"],                  r"billing.*freq|frequency"),
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

# All words that describe operations or columns — these should NOT be used to
# filter rows in the data (would wrongly exclude all rows that don't contain
# these words in text columns)
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
    """Filter rows by location if location keyword found in clause."""
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

    # Fallback: pick first numeric column that matches common priority patterns
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
    """
    Apply caged/uncaged filter.
    Returns (filtered_df, applied_filter_labels).
    """
    q = clause.lower()
    applied = []

    # Find the caged/uncaged column — look for "caged" but NOT "uncaged" in col name
    caged_col = None
    for c in df.columns:
        if re.search(r"\bcaged\b", c, re.I):
            caged_col = c
            break

    if caged_col is None:
        return df, applied

    col_vals = df[caged_col].astype(str).str.strip().str.upper()

    # Determine what the user is asking for
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
    """
    Look for specific customer names mentioned in the query.
    Only applies the filter if a name is found in the data.
    """
    q = clause.lower()

    # Extract candidate name-like tokens (capitalized words not in stop lists)
    # Look for quoted strings first
    quoted = re.findall(r'"([^"]+)"|\'([^\']+)\'', q)
    candidates = [a or b for a, b in quoted]

    # If no quoted names, try to find multi-word proper nouns from the query
    # by matching against actual customer names in the data
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


def _detect_text_filter(clause: str, df: pd.DataFrame) -> "tuple[pd.DataFrame, list]":
    """
    Apply targeted text filtering for caged/uncaged and specific customer names.
    Does NOT filter on generic column-concept words (power, kw, capacity, etc.)
    to avoid incorrectly removing all rows.
    """
    # First apply caged/uncaged filter (most specific)
    work, applied = _detect_caged_filter(clause, df)

    # Then try customer name filter only if explicitly quoted
    q = clause.lower()
    quoted = re.findall(r'"([^"]+)"|\'([^\']+)\'', q)
    if quoted:
        work2, a2 = _detect_customer_filter(clause, work)
        if a2:
            work = work2
            applied += a2

    return work, applied


def _detect_top_n(clause: str) -> int:
    m = re.search(r"\b(?:top|bottom|best|worst)\s+(\d+)\b", clause, re.I)
    return int(m.group(1)) if m else 10


def _parse_or_clauses(query: str) -> list:
    """Split on ' or ' at top level, returning list of OR clauses."""
    return re.split(r"\s+or\s+", query, flags=re.I)


def execute_clause(clause: str, df: pd.DataFrame) -> dict:
    """Execute a single clause and return a result dict."""
    if df.empty:
        return {"title": "No data", "type": "error", "description": "DataFrame is empty."}

    q = clause.lower()

    # Detect operation keyword
    detected_op = None
    for kw, op in sorted(_OP_KW.items(), key=lambda x: -len(x[0])):
        if re.search(r"\b" + re.escape(kw) + r"\b", q):
            detected_op = op
            break

    # Is this a listing query?
    is_listing = any(kw in q for kw in _LIST_KW)

    # Apply location filter
    work = _loc_filter(clause, df)

    # Apply caged/uncaged + specific name filter
    filtered, matched_kws = _detect_text_filter(clause, work)

    grp = _detect_groupby(clause, filtered)
    top_n = _detect_top_n(clause)
    num_col = _detect_num_col(clause, filtered)

    loc_note = ""
    if "_Location" in filtered.columns:
        locs = filtered["_Location"].unique().tolist()
        loc_note = f" ({', '.join(locs)})" if locs and len(locs) < 5 else ""

    rows_used = len(filtered)

    # If an operation is detected, run it
    if detected_op and num_col and num_col in filtered.columns:
        if detected_op in ("Top N Values", "Bottom N Values") and grp is None:
            res, desc = run_op(filtered, num_col, detected_op, None, top_n)
            return {"title": desc, "type": "table",
                    "data": res if isinstance(res, pd.DataFrame) else pd.DataFrame(),
                    "description": f"*{clause}*"}
        elif grp:
            res, desc = run_op(filtered, num_col, detected_op, grp, top_n)
            return {"title": desc, "type": "grouped",
                    "data": res if isinstance(res, pd.DataFrame) else pd.DataFrame(),
                    "description": f"*{clause}*{loc_note}",
                    "x_col": grp, "y_col": num_col}
        else:
            res, desc = run_op(filtered, num_col, detected_op, None, top_n)
            return {"title": desc, "type": "scalar", "data": res,
                    "description": f"*{clause}*{loc_note}",
                    "rows_used": rows_used}
    else:
        # Listing or filtering result
        kw_note = f"Filtered by: {', '.join(matched_kws)}" if matched_kws else "All records"
        if not is_listing and not matched_kws:
            # No listing keyword, no filter, no operation → return summary count
            return {"title": f"Records{loc_note}", "type": "table",
                    "data": filtered,
                    "description": f"*{clause}*"}
        return {"title": f"Records{loc_note} — {kw_note}",
                "type": "table", "data": filtered,
                "description": f"*{clause}*"}


def parse_and_execute(query: str, df: pd.DataFrame) -> list:
    """
    Parse compound query with AND/OR logic.
    AND → independent result blocks.
    OR  → union of filtered rows.
    """
    if df.empty:
        return [{"title": "No data", "type": "error", "description": "DataFrame is empty."}]

    results = []
    and_clauses = re.split(r"\s+and\s+", query.strip(), flags=re.I)

    for and_clause in and_clauses:
        and_clause = and_clause.strip()
        if not and_clause:
            continue

        or_clauses = _parse_or_clauses(and_clause)

        if len(or_clauses) > 1:
            union_frames = []
            for or_c in or_clauses:
                res = execute_clause(or_c.strip(), df)
                if res["type"] in ("table", "grouped") and isinstance(res.get("data"), pd.DataFrame):
                    union_frames.append(res["data"])

            if union_frames:
                merged = pd.concat(union_frames, ignore_index=True, sort=False).drop_duplicates()
                merged = merged.reset_index(drop=True)
                results.append({
                    "title": f"OR Combined — {and_clause[:60]}",
                    "type": "table",
                    "data": merged,
                    "description": f"*Union of: {' | '.join(or_clauses)}*"
                })
            else:
                for or_c in or_clauses:
                    results.append(execute_clause(or_c.strip(), df))
        else:
            results.append(execute_clause(and_clause, df))

    return results or [{"title": "No results", "type": "error",
                        "description": "Could not interpret the query."}]


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
        st.error("No Excel files found. Place files in the 'excel_files/' or 'attached_assets/' folder.")
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

# Customer-facing combined DataFrame (all sheets from selected locations)
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

        cap_c   = find_col(CUST, r"total.*capacity.*kw|total.*capacity.*kva|total.*capacity|capacity.*purchased")
        use_c   = find_col(CUST, r"capacity.*in.*use|in.*use")
        rev_c   = find_col(CUST, r"total.*revenue|total.*mrc|revenue|mrc")
        power_c = find_col(CUST, r"power.*kw|kw.*power|allocated.*kw")
        rack_c  = find_col(CUST, r"\brack\b")
        cust_c  = find_col(CUST, r"customer.*name|client.*name")
        caged_c = find_col(CUST, r"\bcaged\b")
        sub_c   = find_col(CUST, r"subscription.*in.*use|in.*use")
        own_c   = find_col(CUST, r"\brhs\b|\bshs\b|ownership")

        # Row 1 – primary KPIs
        k = st.columns(5)
        total_customers = CUST[cust_c].dropna().nunique() if cust_c else len(CUST)
        k[0].markdown(kpi_html(f"{total_customers:,}", "Unique Customers", "across all locations", CYAN),
                      unsafe_allow_html=True)

        if "_Location" in CUST.columns:
            k[1].markdown(kpi_html(f"{CUST['_Location'].nunique()}", "Active Locations",
                                   f"{sum(len(s) for s in fdata.values())} sheets", LBLUE),
                          unsafe_allow_html=True)

        if cap_c:
            tot_cap = pd.to_numeric(CUST[cap_c], errors="coerce").sum()
            k[2].markdown(kpi_html(fmt(tot_cap), "Total Capacity Purchased", cap_c[:30], GREEN),
                          unsafe_allow_html=True)

        if use_c:
            tot_use = pd.to_numeric(CUST[use_c], errors="coerce").sum()
            k[3].markdown(kpi_html(fmt(tot_use), "Total Capacity In Use", use_c[:30], AMBER),
                          unsafe_allow_html=True)

        if rev_c:
            tot_rev = pd.to_numeric(CUST[rev_c], errors="coerce").sum()
            k[4].markdown(kpi_html(fmt(tot_rev), "Total Revenue (MRC)", "Monthly Recurring", GREEN),
                          unsafe_allow_html=True)
        elif nc:
            v = pd.to_numeric(CUST[nc[-1]], errors="coerce").sum()
            k[4].markdown(kpi_html(fmt(v), nc[-1][:30], "", MUTED), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Row 2 – power & space KPIs
        k2 = st.columns(5)
        if power_c:
            tot_pw = pd.to_numeric(CUST[power_c], errors="coerce").sum()
            k2[0].markdown(kpi_html(fmt(tot_pw), "Total Allocated Power (KW)", power_c[:30], RED),
                           unsafe_allow_html=True)

        if caged_c:
            cage_vals = CUST[caged_c].astype(str).str.upper().str.strip()
            n_caged   = (cage_vals == "CAGED").sum()
            n_uncaged = (cage_vals == "UNCAGED").sum()
            k2[1].markdown(kpi_html(f"{n_caged}", "Caged", f"Uncaged: {n_uncaged}", CYAN),
                           unsafe_allow_html=True)

        if cap_c and use_c:
            t_cap = pd.to_numeric(CUST[cap_c], errors="coerce").sum()
            t_use = pd.to_numeric(CUST[use_c], errors="coerce").sum()
            util  = (t_use / t_cap * 100) if t_cap > 0 else 0
            k2[2].markdown(kpi_html(f"{util:.1f}%", "Utilisation Rate",
                                    "Capacity Used / Purchased", AMBER), unsafe_allow_html=True)

        if sub_c and sub_c != use_c:
            tot_sub = pd.to_numeric(CUST[sub_c], errors="coerce").sum()
            k2[3].markdown(kpi_html(fmt(tot_sub), "Subscription In Use", sub_c[:30], LBLUE),
                           unsafe_allow_html=True)

        if rack_c:
            tot_rack = pd.to_numeric(CUST[rack_c], errors="coerce").sum()
            k2[4].markdown(kpi_html(fmt(tot_rack), "Total Racks / Space", rack_c[:30], GREEN),
                           unsafe_allow_html=True)
        elif len(nc) >= 5:
            v = pd.to_numeric(CUST[nc[4]], errors="coerce").sum()
            k2[4].markdown(kpi_html(fmt(v), nc[4][:30], "", MUTED), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Per-location summary table
        st.markdown('<div class="section-title">Per-Location Summary</div>', unsafe_allow_html=True)
        if "_Location" in CUST.columns:
            agg_cols = [c for c in [cap_c, use_c, rev_c, power_c] if c]
            if agg_cols:
                loc_agg = CUST.groupby("_Location")[agg_cols].apply(
                    lambda g: g.apply(pd.to_numeric, errors="coerce").sum()
                ).reset_index()
                loc_agg.columns = ["Location"] + agg_cols
                if cust_c:
                    loc_agg["Customer Count"] = (
                        CUST.groupby("_Location")[cust_c]
                        .apply(lambda g: g.dropna().nunique())
                        .values
                    )
                else:
                    loc_agg["Customer Count"] = (
                        CUST.groupby("_Location").size().values
                    )
                st.dataframe(loc_agg.round(2), use_container_width=True)
            else:
                lc = CUST["_Location"].value_counts().reset_index()
                lc.columns = ["Location", "Records"]
                st.dataframe(lc, use_container_width=True)

        # Gauge charts
        if cap_c and use_c:
            st.markdown('<div class="section-title">Utilisation Gauges</div>', unsafe_allow_html=True)
            g1, g2 = st.columns(2)
            t_cap = pd.to_numeric(CUST[cap_c], errors="coerce").sum()
            t_use = pd.to_numeric(CUST[use_c], errors="coerce").sum()
            util_pct = min((t_use / t_cap * 100) if t_cap > 0 else 0, 100)

            if rack_c:
                t_rack = pd.to_numeric(CUST[rack_c], errors="coerce").sum()
                rack_pct = min((t_use / t_rack * 100) if t_rack > 0 else util_pct, 100)
            else:
                rack_pct = util_pct

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

        # Capacity vs Usage bar chart by location
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

        # Caged / Uncaged pie
        if caged_c:
            st.markdown('<div class="section-title">Space & Ownership Split</div>',
                        unsafe_allow_html=True)
            p1, p2 = st.columns(2)
            cv = CUST[caged_c].astype(str).str.upper().str.strip()
            cv_valid = cv[cv.isin(["CAGED", "UNCAGED"])]
            if not cv_valid.empty:
                pie = cv_valid.value_counts().reset_index()
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
                    own_d = ov.value_counts().reset_index()
                    own_d.columns = ["Type", "Count"]
                    fo = px.pie(own_d, names="Type", values="Count",
                                title="RHS/SHS Split",
                                color_discrete_sequence=[GREEN, AMBER], hole=0.45)
                    fo.update_layout(**_base_layout(), height=300)
                    p2.plotly_chart(fo, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – DATA EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
with T[1]:
    st.markdown('<div class="section-title">Data Explorer — All Locations &amp; Sheets</div>',
                unsafe_allow_html=True)

    de_loc = st.selectbox("Location", ["All"] + sorted(fdata.keys()), key="de_loc")
    if de_loc == "All":
        view = COMB.copy()
    else:
        sh_map = fdata.get(de_loc, {})
        de_sh  = st.selectbox("Sheet", list(sh_map.keys()), key="de_sh")
        view   = sh_map.get(de_sh, pd.DataFrame()).copy()
        if "_Location" not in view.columns:
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

        show_c = st.multiselect("Columns to display", disp_cols,
                                default=disp_cols[:min(14, len(disp_cols))],
                                key="de_cols")
        out = vw[[c for c in show_c if c in vw.columns]] if show_c else vw[disp_cols]

        st.markdown(f'<span class="badge">{len(out):,} rows</span>', unsafe_allow_html=True)
        st.dataframe(out.head(500), use_container_width=True, height=420)
        st.download_button("⬇ Download CSV",
                           out.to_csv(index=False).encode(), "sify_data.csv", "text/csv")

        with st.expander("📊 Column Statistics"):
            nc2 = num_cols(vw)
            if nc2:
                st.dataframe(vw[nc2].describe().round(3).T, use_container_width=True)
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
                                  ["(none)"] + tc_op + ["_Location", "_Sheet"], key="op_grp")

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
                op_fv = st.text_input("Filter value (exact match or partial)", key="op_fv")

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
                st.markdown(f'<div style="color:{MUTED};font-size:.8rem">Computed from '
                            f'{len(wdf):,} rows</div>', unsafe_allow_html=True)
            else:
                st.warning("No result. Check column and filter.")
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
    st.markdown('<div class="section-title">Chart Studio — All 15 Chart Types</div>',
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

        with st.expander("📚 Quick Gallery — All 15 Chart Types"):
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
# TAB 4 – SMART QUERY
# ══════════════════════════════════════════════════════════════════════════════
with T[4]:
    st.markdown('<div class="section-title">🧠 Smart Query — Compound AND/OR Query Engine</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    <div style="background:{DARK2};border:1px solid {BORD};border-radius:10px;
         padding:16px 20px;margin-bottom:16px;font-size:.86rem;color:{MUTED}">
    <b style="color:{TEXT}">Use <code>and</code> to chain results, <code>or</code> for union of filters:</b><br><br>
    &nbsp;• <code>list all caged customers</code><br>
    &nbsp;• <code>show uncaged customers in noida</code><br>
    &nbsp;• <code>total sum of capacity purchased</code><br>
    &nbsp;• <code>count caged customers by location</code><br>
    &nbsp;• <code>total revenue bangalore and average capacity in use noida</code><br>
    &nbsp;• <code>maximum capacity airoli and minimum capacity noida</code><br>
    &nbsp;• <code>show customers in airoli or noida</code><br>
    &nbsp;• <code>list caged customers and total sum capacity purchased</code><br>
    &nbsp;• <code>sum capacity in use by location</code>
    </div>""", unsafe_allow_html=True)

    query   = st.text_input(
        "🔍 Enter your query",
        placeholder="e.g. list all caged customers and total sum of capacity purchased",
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
            st.markdown(f'<div class="section-title">{i+1}. {res["title"]}</div>',
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
                    st.info(f"Result: {val}")

            elif rtype in ("table", "grouped"):
                data = res.get("data", pd.DataFrame())
                if isinstance(data, pd.DataFrame) and not data.empty:
                    meta   = [c for c in ["_Location", "_Sheet"] if c in data.columns]
                    show_c = [c for c in data.columns if not c.startswith("_")][:25]
                    display = data[meta + show_c] if meta else data[show_c]
                    st.markdown(f'<span class="badge">{len(display):,} rows</span>',
                                unsafe_allow_html=True)
                    st.dataframe(display.head(500), use_container_width=True, height=340)
                    all_tables.append((res["title"], display))

                    if rtype == "grouped":
                        xc = res.get("x_col", "")
                        yc = res.get("y_col", "")
                        if xc and yc and xc in data.columns and yc in data.columns:
                            fig_g = px.bar(data.head(30), x=xc, y=yc,
                                           color=yc, color_continuous_scale="Blues",
                                           title=res["title"])
                            fig_g.update_layout(**_base_layout(), height=350)
                            st.plotly_chart(fig_g, use_container_width=True)

                    if "_Location" in data.columns and data["_Location"].nunique() > 1:
                        lb = data["_Location"].value_counts().reset_index()
                        lb.columns = ["Location", "Rows"]
                        fig_lb = px.bar(lb, x="Location", y="Rows",
                                        color="Rows", color_continuous_scale="Blues",
                                        title="Row count by Location")
                        fig_lb.update_layout(**_base_layout(), height=300)
                        st.plotly_chart(fig_lb, use_container_width=True)

                    nc_res = num_cols(data)
                    if nc_res:
                        fig_h = px.histogram(data.dropna(subset=[nc_res[0]]),
                                             x=nc_res[0], nbins=25,
                                             title=f"Distribution — {nc_res[0]}",
                                             color_discrete_sequence=[LBLUE])
                        fig_h.update_layout(**_base_layout(), height=280)
                        st.plotly_chart(fig_h, use_container_width=True)
                else:
                    st.warning("No matching rows for this clause.")

            else:
                st.error(res.get("description", "Query error."))

            st.markdown(f"<hr style='border-color:{BORD};margin:10px 0 20px'>",
                        unsafe_allow_html=True)

        if all_tables:
            combined_out = pd.concat(
                [df.assign(_ResultBlock=title) for title, df in all_tables],
                ignore_index=True, sort=False)
            st.download_button("⬇ Download All Results (CSV)",
                               combined_out.to_csv(index=False).encode(),
                               "smart_query_results.csv", "text/csv")


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
            st.info("No data available for this metric across selected locations.")
