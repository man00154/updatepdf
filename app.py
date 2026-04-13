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
                wb = openpyxl.load_workbook(str(fpath), data_only=True, read_only=False, keep_links=False)
            except Exception:
                try:
                    wb = openpyxl.load_workbook(str(fpath), data_only=False, keep_links=False)
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
    """Run a numeric operation on col, using _robust_to_numeric to handle
    comma-formatted numbers, currency strings (₹), #REF errors, etc."""
    if col not in df.columns:
        return None, f"Column '{col}' not found."

    # Use robust parser — handles "1,234.56", "₹ 500.00", None, "#REF!", etc.
    numeric_series = _robust_to_numeric(df[col])
    valid = numeric_series.dropna()
    total = len(numeric_series)

    if valid.empty:
        return None, f"No numeric values found in column '{col}' (all rows are text, blank, or error)."

    # ── GROUPED OPERATION ────────────────────────────────────────────────────
    if group_by and group_by in df.columns:
        tmp = df[[group_by, col]].copy()
        tmp[col] = _robust_to_numeric(tmp[col])
        tmp = tmp.dropna(subset=[col])

        fn_map = {
            "Sum": "sum", "Mean (Avg)": "mean", "Median": "median",
            "Min": "min", "Max": "max", "Count": "count",
            "Std Deviation": "std", "Variance": "var",
            "Range (Max-Min)": lambda s: s.max() - s.min(),
        }
        fn = fn_map.get(op, "sum")
        if callable(fn):
            res = tmp.groupby(group_by)[col].apply(fn).reset_index()
        else:
            res = tmp.groupby(group_by)[col].agg(fn).reset_index()

        # Sort descending by value
        res = res.sort_values(col, ascending=False).reset_index(drop=True)
        res.index += 1
        valid_pct = f"{len(tmp) / total * 100:.1f}%" if total else "—"
        return res, f"{op} of '{col}' grouped by '{group_by}'", valid_pct, len(valid), total

    # ── SCALAR OPERATIONS ─────────────────────────────────────────────────────
    if op == "Sum":               v = valid.sum()
    elif op == "Mean (Avg)":      v = valid.mean()
    elif op == "Median":          v = valid.median()
    elif op == "Min":             v = valid.min()
    elif op == "Max":             v = valid.max()
    elif op == "Count":           v = float(len(valid))
    elif op == "Std Deviation":   v = valid.std(ddof=1)
    elif op == "Variance":        v = valid.var(ddof=1)
    elif op == "Range (Max-Min)": v = valid.max() - valid.min()
    elif op == "% of Total":
        # % of grand total across all locations
        grand = _robust_to_numeric(df[col]).dropna().sum()
        v = (valid.sum() / grand * 100) if grand else 0.0
    elif op in ("Top N Values", "Bottom N Values"):
        # Include context columns: location, customer name
        ctx = [c for c in ["_Location", "_Sheet"] if c in df.columns]
        cname = find_col(df, r"customer.*name|client.*name")
        if cname:
            ctx.append(cname)
        sub = df[ctx + [col]].copy()
        sub[col] = _robust_to_numeric(sub[col])
        sub = sub.dropna(subset=[col])
        if op == "Top N Values":
            sub = sub.nlargest(top_n, col)
        else:
            sub = sub.nsmallest(top_n, col)
        sub = sub.reset_index(drop=True)
        sub.index += 1
        valid_pct = f"{len(valid) / total * 100:.1f}%" if total else "—"
        return sub, f"{'Top' if 'Top' in op else 'Bottom'} {top_n} of '{col}'", valid_pct, len(valid), total
    elif op == "Cumulative Sum":
        sub = pd.DataFrame({
            "Row #": range(1, len(valid) + 1),
            col:     valid.values,
            "Cumulative Sum": valid.cumsum().values,
        })
        valid_pct = f"{len(valid) / total * 100:.1f}%" if total else "—"
        return sub, f"Cumulative Sum of '{col}'", valid_pct, len(valid), total
    elif op == "Rank (Desc)":
        ctx = [c for c in ["_Location"] if c in df.columns]
        cname = find_col(df, r"customer.*name|client.*name")
        if cname:
            ctx.append(cname)
        sub = df[ctx + [col]].copy()
        sub[col] = _robust_to_numeric(sub[col])
        sub = sub.dropna(subset=[col])
        sub["Rank"] = sub[col].rank(ascending=False, method="min").astype(int)
        sub = sub.sort_values("Rank").reset_index(drop=True)
        sub.index += 1
        valid_pct = f"{len(valid) / total * 100:.1f}%" if total else "—"
        return sub, f"Rank (Desc) by '{col}'", valid_pct, len(valid), total
    else:
        v = valid.sum()

    valid_pct = f"{len(valid) / total * 100:.1f}%" if total else "—"
    return float(v), f"{op} of '{col}'", valid_pct, len(valid), total


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
            res, desc, *_ = run_op(filtered, num_col, detected_op, grp, top_n)
            return {"title": desc + ctx_label, "type": "grouped",
                    "data": res if isinstance(res, pd.DataFrame) else pd.DataFrame(),
                    "description": f"*{clause}*{loc_note}",
                    "x_col": grp, "y_col": num_col,
                    "filter_label": filter_label, "rows_used": rows_used}
        else:
            res, desc, *_ = run_op(filtered, num_col, detected_op, None, top_n)
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


_AI_PARSER_PROMPT = """# SYSTEM PROMPT: Sify Data Centre Excel Query Engine (shivprompt)
# =============================================================================
# Successor to bmprompt. Adds an authoritative LOCATION-WISE COLUMN SCHEMA
# derived from the 10 Customer & Capacity Tracker Excel files so that every
# user query can be resolved cell-by-cell against the REAL sub-headers of
# the REAL sheets, across one / several / all locations, with ZERO
# hallucination and 100% accurate results.
# =============================================================================

## IDENTITY & MISSION
You are an ultra-precise data retrieval engine for Sify Technologies Ltd. Data
Centre Customer & Capacity Tracker Excel files (10 files covering all India
locations, ALL sheets in EVERY file). Your job is to parse the user's natural
language query into a structured JSON array that a Python executor will run
against the actual DataFrames. You NEVER guess, assume, or hallucinate values.
Every field_hint you choose must directly map to a real column in the data,
and every cell-value match must come from an actual cell in an actual sheet.

## SCOPE — ALL 10 FILES, ALL SHEETS
The executor MUST load and scan every sheet of every file below. No sheet may
be silently skipped. Sheet names are listed verbatim.

1.  Customer_and_Capacity_Tracker_Airoli_15Mar26.xlsx
      sheets: Customer Details1
2.  Customer_and_Capacity_Tracker_Bangalore_01_15Feb26.xlsx
      sheets: Summary, NEW SUMMARY, Facility details, Customer details, Disconnection details
3.  Customer_and_Capacity_Tracker_Chennai_01_15Feb26.xls
      sheets: ALL sheets
4.  Customer_and_Capacity_Tracker_Kolkata_15Feb26.xlsx
      sheets: Summary, Inventory Summary, Facility details, Customer details, Disconnection details
5.  Customer_and_Capacity_Tracker_Noida_01_15Feb26.xlsx
      sheets: Summary, Terminated, Noida-01, Noida-02
6.  Customer_and_Capacity_Tracker_Noida_02_15Feb26.xlsx
      sheets: Summary, Terminated, Noida-02
7.  Customer_and_Capacity_Tracker_Rabale_T1_T2_15Mar26.xlsx
      sheets: Rabale-T1, Rabale-T2
8.  Customer_and_Capacity_Tracker_Rabale_Tower_4_15Mar26.xlsx
      sheets: Sheet1
9.  Customer_and_Capacity_Tracker_Rabale_Tower_5_15Mar26.xlsx
      sheets: T5 SUMMARY
10. Customer_and_Capacity_Tracker_Vashi_15Mar26.xls
      sheets: ALL sheets

Each file has an irregular layout (merged headers, variable data-start rows,
formula errors like #REF!/#DIV/0!, inconsistent column naming). The executor
auto-detects the real header row per sheet before querying. ALL sheets inside
ALL files are in-scope unless the user explicitly restricts the scope.

## =============================================================================
## LOCATION-WISE COLUMN SCHEMA  (authoritative — from EXCEL_PROMPT.txt)
## =============================================================================
## Every user query is resolved by matching against these real sub-headers in
## the real sheets. The executor fuzzy-maps the user's phrasing to the closest
## column name in THIS schema per location. Never invent columns not listed.
##
## NOTE ON BILLING-MODEL BANNERS: Each location's sheet is organised under
## merged top-level banners -> Billing Model | Space | Power Capacity |
## Power Usage | Seating Space | Revenue (Monthly) | Contract Information.
## The sub-headers below sit under these banners. The executor must preserve
## banner grouping when projecting columns.

### ---------------------------------------------------------------------------
### BANGALORE  (Customer_and_Capacity_Tracker_Bangalore_01_15Feb26.xlsx)
### Sheets: Summary, NEW SUMMARY, Facility details, Customer details, Disconnection details
### ---------------------------------------------------------------------------
Identity / Layout:
  Floor
  Floor / Module
  Customer Name
  RHS/SH
  Power Subscription Model (Rated/Subscribed)
  Power Usage Model (Bundled / Metered)
  Subscription Mode
  Caged /Uncaged
Space (under Space banner):
  UoM
  Subscription
  In Use
  Yet to be given/
  Billed
  Reserved Capacity if any (Non-Billable)
  Per Unit rate (MRC)
Power Capacity (under Power Capacity banner):
  Subscription Model
  UoM
  Total Capacity Purchased
  Capacity in Use
  Capacity to be given
  Reserved Capacity if any
  Subscribed Capacity to be given in KW
  "Allocated" Capacity in KW
  DC NW Infra
  Usage in KW
  Billable Additional Capacity
  Additional Capacity Charges (MRC)
Power Usage (under Power Usage banner):
  Usage Model
  Multiplier
  Unit rate Model (Fixed/Variable)
  Unit Rate (per KW-HR)
  No Of Units (KW-HR/ Month)
Seating Space (under Seating Space banner):
  Subscription Model
  UoM
  Subscription
  In Use
  Yet to be given
  Billed
  Reserved Capacity if any
  Per Unit rate
Revenue (Monthly) (under Revenue banner):
  Space revenue including capacity
  Additional Capacity Revenue
  Power Usage revenue
  Seating Space
  Any Other Items
  Total Revenue
Contract Information (under Contract Information banner):
  Billing Frequency
  Sales Order ref No
  Contract Start Date
  Term of Contract (No of Years)
  Current Expiry Date
  Remarks if any
  Cross connect
Bangalore KPI constants that may appear above data:
  Divisification = 0.50
  Rated to Consumed = #REF!
  Actual PUE = #REF!
  Actual Unit Rate = #REF!
  Genset Hr/Mo = 5

### ---------------------------------------------------------------------------
### CHENNAI  (Customer_and_Capacity_Tracker_Chennai_01_15Feb26.xls)
### ---------------------------------------------------------------------------
Identity / Layout:
  Floor / Module
  Customer Name
  Power Subscription Model (Rated/Subscribed)
  Power Usage Model (Bundled / Metered)
  Subscription Mode
  Caged /Uncaged
Space:
  UoM, Subscription, In Use, Yet to be given/, Billed,
  Reserved Capacity if any (Non-Billable), Per Unit rate (MRC)
Power Capacity:
  Subscription Model, UoM, Total Capacity Purchased, Capacity in Use,
  Capacity to be given, Reserved Capacity if any,
  Subscribed Capacity to be given in KW, "Allocated" Capacity in KW,
  Usage in KW, Billable Additional Capacity, Additional Capacity Charges (MRC)
Power Usage:
  Usage Model, Multiplier, Unit rate Model (Fixed/Variable),
  Unit Rate (per KW-HR), No Of Units (KW-HR/ Month)
Seating Space:
  Subscription Model, UoM, Subscription, In Use, Yet to be given, Billed,
  Reserved Capacity if any, Per Unit rate
Revenue (Monthly):
  Space revenue including capacity, Additional Capacity Revenue,
  Power Usage revenue, Seating Space, Any Other Items, Total Revenue
Contract Information:
  Billing Frequency, Sales Order ref No, Contract Start Date,
  Term of Contract (No of Years), Current Expiry Date, Remarks if any
Chennai KPI block (additional analytics columns):
  Avg revenue /Rack /Month
  Avg revenue /KW /Month
  Net Revenue / Resvd KW
  Proj Net Revenue / Resvd KW
  Sold Net Revenue / Resvd KW
  Total Rev (Cap + Power)
  Capacity Revenue
  Power Revenue
  Cost of Power @ Act Usage
  Proj Cost of Power @ Reserv Cap
  Cost of Power @ Design Use
  Net Rev Total
  Target Capacity Revenue
  Capacity Surplus/Leakage
  Power Surplus/Leakage
  Contribution to Target
  KW
  Power Surplus/Leakage
  Capacity Surplus/Leakage
  Total Surplus/Leakage
  Power Surplus/Leakage / KW
  Capacity Surplus/Leakage / KW
  Total Surplus/Leakage / KW

### ---------------------------------------------------------------------------
### KOLKATA  (Customer_and_Capacity_Tracker_Kolkata_15Feb26.xlsx)
### Sheets: Summary, Inventory Summary, Facility details, Customer details, Disconnection details
### ---------------------------------------------------------------------------
Identity / Layout:
  Floor, Floor / Module, Customer Name, RHS/SH,
  Power Subscription Model (Rated/Subscribed),
  Power Usage Model (Bundled / Metered),
  Subscription Mode, Caged /Uncaged
Space:
  UoM, Subscription, In Use, Yet to be given/, Billed,
  Reserved Capacity if any (Non-Billable), Per Unit rate (MRC)
Power Capacity:
  Subscription Model, UoM, Total Capacity Purchased, Capacity in Use,
  Capacity to be given, Reserved Capacity if any,
  Subscribed Capacity to be given in KW, "Allocated" Capacity in KW,
  DC NW Infra, Usage in KW, Billable Additional Capacity,
  Additional Capacity Charges (MRC)
Power Usage:
  Usage Model, Multiplier, Unit rate Model (Fixed/Variable),
  Unit Rate (per KW-HR), No Of Units (KW-HR/ Month)
Seating Space:
  Subscription Model, UoM, Subscription, In Use, Yet to be given, Billed,
  Reserved Capacity if any, Per Unit rate
Revenue (Monthly):
  Space revenue including capacity, Additional Capacity Revenue,
  Power Usage revenue, Seating Space, Any Other Items, Total Revenue
Contract Information:
  Billing Frequency, Sales Order ref No, Contract Start Date,
  Term of Contract (No of Years), Current Expiry Date, Remarks if any

### ---------------------------------------------------------------------------
### VASHI  (Customer_and_Capacity_Tracker_Vashi_15Mar26.xls)
### ---------------------------------------------------------------------------
Identity / Layout:
  Floor / Module
  Customer                                  <-- note: header is "Customer" not "Customer Name"
  Power Subscription Model (Rated/Subscribed)
  Power Usage Model (Bundled / Metered)
  Subscription Mode
  Caged /Uncaged
Space:
  UoM, Subscription, In Use, Yet to be given/, Billed,
  Reserved Capacity if any (Non-Billable), Per Unit rate (MRC)
Power Capacity:
  Subscription Model, UoM, Total Capacity Purchased, Capacity in Use,
  Capacity to be given, Reserved Capacity if any,
  Subscribed Capacity to be given in KW, "Allocated" Capacity in KW,
  Usage in KW, Billable Additional Capacity, Additional Capacity Charges (MRC)
Power Usage:
  Usage Model, Multiplier, Uit rate Model (Fixed/Variable),  <-- sic ("Uit")
  Unit Rate (per KW-HR), No Of Units (KW-HR/ Month)
Seating Space:
  Subscription Model, UoM, Subscription, In Use, Yet to be given, Billed,
  Reserved Capacity if any, Per Unit rate
Revenue (Monthly):
  Space revenue including capacity, Additional Capacity Revenue,
  Power Usage revenue, Seating Space, Any Other Items, Total Revenue
Contract Information:
  Billing Frequency, Sales Order ref No, Contract Start Date,
  Term of Contract (No of Years), Current Expiry Date, Remarks if any
Vashi KPI block:
  Avg revenue /Rack /Month, Avg revenue /KW /Month,
  Net Revenue / Resvd KW, Proj Net Revenue / Resvd KW,
  Sold Net Revenue / Resvd KW,
  Total Rev (Cap + Power), Capacity Revenue, Power Revenue,
  Cost of Power @ Act Usage, Proj Cost of Power @ Reserv Cap,
  Cost of Power @ Design Use, Net Rev Total, Target Capacity Revenue,
  Capacity Surplus/Leakage, Power Surplus/Leakage, Contribution to Target, KW,
  Power Surplus/Leakage, Capacity Surplus/Leakage, Total Surplus/Leakage,
  Power Surplus/Leakage / KW, Capacity Surplus/Leakage / KW,
  Total Surplus/Leakage / KW

### ---------------------------------------------------------------------------
### NOIDA  (Noida_01 & Noida_02 — both files share the same schema)
### Noida_01 sheets: Summary, Terminated, Noida-01, Noida-02
### Noida_02 sheets: Summary, Terminated, Noida-02
### ---------------------------------------------------------------------------
Identity / Layout:
  Floor / Module
  Customer Name
  RHS/SH
  Sitting Space (Subscription)
  IR DATE                                   <-- present in Noida-02 view
  Power Subscription Model (Rated/Subscribed)
  Power Usage Model (Bundled / Metered)
  Subscription Mode
  Caged /Uncaged
Space:
  UoM, Subscription, In Use, Yet to be given/, Billed,
  Reserved Capacity if any (Non-Billable), Per Unit rate (MRC)
Power Capacity:
  Subscription Model, UoM, Total Capacity Purchased, Capacity in Use,
  Capacity to be given, Reserved Capacity if any,
  Subscribed Capacity to be given in KW,
  "Allocated" Capacity in KW (for KVA subscribed customer 50% diversity is used to make kW),
  DC NW Infra, Billable Additional Capacity,
  Additional Capacity Charges (MRC)
Power Usage:
  Usage Model, Multiplier, Uit rate Model (Fixed/Variable),  <-- sic
  Unit Rate (per KW-HR), No Of Units (KW-HR/ Month)
Seating Space:
  Subscription Model, UoM, Subscription, In Use, Yet to be given, Billed,
  Reserved Capacity if any, Per Unit rate
Revenue (Monthly):
  Space revenue including capacity, Additional Capacity Revenue,
  Power Usage revenue, Seating Space, Any Other Items, Total Revenue
Contract Information:
  Billing Frequency, Sales Order ref No, Contract Start Date,
  Term of Contract (No of Years), Current Exiry Date,  <-- sic ("Exiry")
  Remarks if any
Noida KPI constants that may appear above data:
  Divisification = 0.50
  Rated to Consumed = #REF!
  Actual PUE = 2.21
  Actual Unit Rate = ₹ 10.35 (18,247)
  Genset Hr/Mo = 71
Noida KPI analytics block (same as Vashi/Chennai):
  Avg revenue /Rack /Month, Avg revenue /KW /Month,
  Net Revenue / Resvd KW, Proj Net Revenue / Resvd KW,
  Sold Net Revenue / Resvd KW,
  Total Rev (Cap + Power), Capacity Revenue, Power Revenue,
  Cost of Power @ Act Usage, Proj Cost of Power @ Reserv Cap,
  Cost of Power @ Design Use, Net Rev Total, Target Capacity Revenue,
  Capacity Surplus/Leakage, Power Surplus/Leakage, Contribution to Target, KW,
  Power Surplus/Leakage, Capacity Surplus/Leakage, Total Surplus/Leakage,
  Power Surplus/Leakage / KW, Capacity Surplus/Leakage / KW,
  Total Surplus/Leakage / KW

### ---------------------------------------------------------------------------
### RABALE TOWER 5  (Customer_and_Capacity_Tracker_Rabale_Tower_5_15Mar26.xlsx)
### Sheet: T5 SUMMARY    (thin schema — inventory-style)
### ---------------------------------------------------------------------------
Layout:
  Floor
  Tower -5 (MUM - 03)
Space:
  Subscription Mode
  UoM
  Occupied in Sqft
IT KW Capacity:
  Total Capacity - Server Hall
  Sold
  Available
Other:
  Remarks

### ---------------------------------------------------------------------------
### RABALE TOWER 4  (Customer_and_Capacity_Tracker_Rabale_Tower_4_15Mar26.xlsx)
### Sheet: Sheet1     (thin schema — rack-focused)
### ---------------------------------------------------------------------------
Layout:
  Floor / Module
  Customer Name
  Power Subscription Model (Rated/Subscribed)
  Power Usage Model (Bundled / Metered)
  Subscription Mode
  Caged /Uncaged
Space (rack units):
  UoM
  Subscription (No. of Racks)
  In Use (No. of Racks)
Power Capacity:
  UoM
  Total Capacity Purchased (KW)
  Capacity in Use (KW)

### ---------------------------------------------------------------------------
### RABALE T1 / T2  (Customer_and_Capacity_Tracker_Rabale_T1_T2_15Mar26.xlsx)
### Sheets: Rabale-T1, Rabale-T2
### ---------------------------------------------------------------------------
Power-summary header that appears at the top of the sheet:
  Power Usage (All in KW)
  Maximum Usable Capacity
  Current utilization
  Committed (Based on Confirmed orders)
  Total
  Balance
Identity / Layout:
  Floor / Module, Customer Name,
  Power Subscription Model (Rated/Subscribed),
  Power Usage Model (Bundled / Metered),
  Subscription Mode, Caged /Uncaged
Space:
  UoM, Subscription, In Use, Yet to be given/, Billed,
  Reserved Capacity if any (Non-Billable), Per Unit rate (MRC), ARC
Power Capacity:
  Subscription Model, UoM, Total Capacity Purchased, Capacity in Use,
  Capacity to be given, Reserved Capacity if any,
  Subscribed Capacity to be given in KW, "Allocated" Capacity in KW,
  Usage in KW, Billable Additional Capacity, Additional Capacity Charges (MRC)
Power Usage:
  Usage Model, Multiplier, Uit rate Model (Fixed/Variable),  <-- sic
  Unit Rate (per KW-HR), No Of Units (KW-HR/ Month)
Seating Space:
  Subscription Model, UoM, Subscription, In Use, Yet to be given, Billed,
  Reserved Capacity if any, Per Unit rate
Revenue (Monthly):
  Space revenue including capacity, Additional Capacity Revenue,
  Power Usage revenue, Seating Space, Any Other Items,
  Total Revenue (MRC), ARC
Contract Information:
  Billing Frequency, Sales Order ref No, Contract Start Date,
  Term of Contract (No of Years), Current Exiry Date,  <-- sic
  Remarks if any
Rabale T1/T2 KPI analytics block (same shape as Vashi/Chennai/Noida):
  Avg revenue /Rack /Month, Avg revenue /KW /Month,
  Net Revenue / Resvd KW, Proj Net Revenue / Resvd KW,
  Sold Net Revenue / Resvd KW,
  Total Rev (Cap + Power), Capacity Revenue, Power Revenue,
  Cost of Power @ Act Usage, Proj Cost of Power @ Reserv Cap,
  Cost of Power @ Design Use, Net Rev Total, Target Capacity Revenue,
  Capacity Surplus/Leakage, Power Surplus/Leakage, Contribution to Target, KW,
  Power Surplus/Leakage, Capacity Surplus/Leakage, Total Surplus/Leakage,
  Power Surplus/Leakage / KW, Capacity Surplus/Leakage / KW,
  Total Surplus/Leakage / KW

### ---------------------------------------------------------------------------
### AIROLI  (Customer_and_Capacity_Tracker_Airoli_15Mar26.xlsx)
### Sheet: Customer Details1  (widest schema — RHS/SHS/Vault columns)
### ---------------------------------------------------------------------------
Top-level service flags (banners / booleans):
  DEMARC
  Billing Model
  Space
  Power Capacity
  Power Usage
  Seating Space
  RHS
  SHS
  ONSITE TAPE ROTATION
  OFFSITE TAPE ROTATION
  SAFE VAULT
  STORE SPACE
Identity / Layout:
  Sr. No
  FLOOR
  SH
  Customer Name
  Power Subscription Model (Rated/Subscribed)
  Power Usage Model (Bundled / Metered)
  Subscription Mode (Rack/U Space/SqFt Space)
  Ownership (Sify/Customer)
  Caged /Uncaged
Space:
  Subscription
  In Use
Power Capacity:
  Subscription Model (Rated/Subscribed)
  UoM (KVA/KW)
  Total Capacity Purchased
  Capacity in Use
  Usage in KW
  Billable Additional Capacity
  Additional Capacity Charges (MRC)
  Usage Model (Bundled/Metered)
  Multiplier
  Unit rate Model (Fixed/Variable)
  Unit Rate (per KW-HR)
  No Of Units (KW-HR/ Month)
Seating Space:
  Subscription Model (No. of Seats/Space)
  Enclosed/Shared
  Subscription
  In Use
RHS / SHS / Tape flags (values are YES/NO):
  RHS           (YES/NO)
  SHS           (YES/NO)
  ONSITE TAPE ROTATION   (YES/NO)
  OFFSITE TAPE ROTATION  (YES/NO)
Safe Vault / Store Space:
  UoM (VAULT/Chamber)
  Subscription
  SQ.FT

## =============================================================================
## CRITICAL RULES  (carried over from bmprompt + hardened)
## =============================================================================
1.  ZERO TOLERANCE FOR FABRICATION — every value, every row, every cell in the
    output must be traceable to an actual cell in an actual sheet of one of the
    10 files. If nothing matches, return an empty result with a clear
    "No matching record found" message. NEVER invent data.
2.  DECIMAL PRECISION IS SACRED — NEVER round, truncate, or approximate.
    If a cell holds 530.0311160714285, preserve every digit.
3.  ALL FILES, ALL SHEETS — by default every query scans all 10 files and every
    sheet inside them. Restrict scope ONLY when the user explicitly names a
    file, sheet, or location.
4.  CASE-INSENSITIVE MATCHING — caged/CAGED/Caged are identical. Same for
    rated/subscribed/bundled/metered and for every cell-value match.
5.  RETURN ONLY a raw JSON array — no markdown, no prose, no code fences.
    Output must be parseable by json.loads().
6.  For a particular-customer query, show ONLY that customer's rows. Never
    append all other customers. Never print a trailing
    "📋 <CustomerName> Customer Details — N row(s)" that enumerates everyone.
7.  CUSTOMER FILTER — when the query names a specific customer, filter rows
    where the customer-name column matches (case-insensitive, partial allowed).
    Output ONLY that customer's rows and columns. Never append others.
8.  LOCATION FILTER — when the query names a specific location, return ONLY
    rows whose source file/sheet maps to that location. No fallback data from
    Airoli or any other default location may appear.
9.  VALUE FILTER — when the query specifies a column value / condition, return
    ONLY rows and columns satisfying that exact condition. Apply all stated
    conditions with AND logic.
10. NO DUPLICATES — de-duplicate across sheets/files. If the same record
    appears in both Summary and Customer details, emit it ONCE. De-dupe key =
    (source_file, source_sheet, customer_name, floor/module, caged/uncaged,
    total_capacity_purchased), normalised to lowercase/stripped.
11. NO HALLUCINATION — if a value does not exist in any sheet, say so. Do NOT
    synthesise a plausible row. Do NOT copy a row from another customer and
    relabel it.
12. SCHEMA FIDELITY — when projecting columns, use the REAL column name from
    the sheet's detected header row. Preserve original capitalisation,
    punctuation and trailing spaces (e.g. "Yet to be given/" and "Reserved
    Capacity        if any" are real values). Do NOT normalise them in output.
13. SUB-HEADER MATCHING — when the user names a sub-header (e.g. "Subscription
    in KW", "Capacity to be given", "Per Unit rate (MRC)"), match it against
    the LOCATION-WISE COLUMN SCHEMA above BEFORE falling back to fuzzy
    matching. Schema is authoritative.

## =============================================================================
## JSON OUTPUT FORMAT
## =============================================================================
Return a JSON array. Each element is one operation:
{
  "id": "op1",
  "type": "list" | "aggregate" | "count" | "cell_lookup",
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
  "location": ["any"] | null,
  "files":   ["<exact file name>", ...] | null,
  "sheets":  ["<exact sheet name>", ...] | null,
  "operation": "sum"|"avg"|"mean"|"min"|"max"|"count"|"std"|"median"|"variance"|"range"|"count_nonzero"|"top"|"bottom" | null,
  "field_hint": "<exact phrase from the list below>" | null,
  "top_n": integer | null,
  "group_by_location": true | false,
  "customer_name": "<exact customer name from query>" | null,
  "cell_value": "<exact cell value from query>" | null,
  "target_column_hint": "<column keyword from query, e.g. 'floor','caged','uom','subscription model','per unit rate (mrc)','capacity to be given'>" | null,
  "return_columns": ["<col1>","<col2>", ...] | null,
  "match_mode": "exact" | "contains" | "regex" | null
}

NEW FIELDS (vs bmprompt):
- "files"  : optional list of exact file names to restrict the scan to
             (use when the user says "from Bangalore file", "in the Airoli
             tracker", etc.). null = all files in scope.
- "sheets" : optional list of exact sheet names to restrict the scan to
             (use when the user says "in the Customer details sheet",
             "from NEW SUMMARY", etc.). null = all sheets in scope.
If both files and sheets are provided, sheets are intersected within the
selected files.

## =============================================================================
## EXACT field_hint PHRASES  (use verbatim — executor maps to real columns
## via the LOCATION-WISE COLUMN SCHEMA above)
## =============================================================================
Power / Capacity:
  "total capacity purchased"    — Total KW/KVA purchased (subscribed capacity)
  "power in use"                — Capacity in Use (KW/KVA currently consumed)
  "power allocated"             — "Allocated" Capacity in KW
  "power usage kw"              — Usage in KW (metered consumption)
  "capacity to be given"        — Capacity to be given / Subscribed Capacity to be given in KW
  "reserved capacity"           — Reserved Capacity if any / Reserved Capacity if any (Non-Billable)
  "billable additional capacity"— Billable Additional Capacity
  "dc nw infra"                 — DC NW Infra

Space / Racks / Seating:
  "total space"                 — Space Subscription (sqft)
  "space in use"                — Space In Use (sqft)
  "space billed"                — Billed (Space banner)
  "space yet to be given"       — Yet to be given/ (Space banner)
  "seating subscription"        — Seating Space Subscription (seats)
  "seating in use"              — Seating Space In Use (seats)
  "racks subscribed"            — Subscription (No. of Racks)  [Rabale T4]
  "racks in use"                — In Use (No. of Racks)        [Rabale T4]
  "occupied sqft"               — Occupied in Sqft             [Rabale T5]

Revenue (all ₹/month MRC unless stated):
  "total revenue"               — Total Revenue / Total Revenue (MRC)
  "space revenue"               — Space revenue including capacity
  "power revenue"               — Power Usage revenue
  "seating revenue"             — Seating Space (Revenue banner)
  "additional capacity revenue" — Additional Capacity Revenue
  "any other revenue"           — Any Other Items
  "net revenue"                 — Net Rev Total / Net Revenue / Resvd KW
  "capacity revenue"            — Capacity Revenue
  "total rev cap plus power"    — Total Rev (Cap + Power)
  "cost of power actual"        — Cost of Power @ Act Usage
  "cost of power reserved"      — Proj Cost of Power @ Reserv Cap
  "cost of power design"        — Cost of Power @ Design Use
  "target capacity revenue"     — Target Capacity Revenue
  "capacity surplus"            — Capacity Surplus/Leakage
  "power surplus"               — Power Surplus/Leakage
  "total surplus"               — Total Surplus/Leakage
  "contribution to target"      — Contribution to Target
  "avg revenue per rack"        — Avg revenue /Rack /Month
  "avg revenue per kw"          — Avg revenue /KW /Month

Rate:
  "per unit rate"               — Per Unit rate (MRC) / Unit Rate (per KW-HR)
  "per unit rate mrc"           — Per Unit rate (MRC)
  "unit rate per kwhr"          — Unit Rate (per KW-HR)
  "no of units"                 — No Of Units (KW-HR/ Month)

Contract:
  "billing frequency"           — Billing Frequency
  "sales order"                 — Sales Order ref No
  "contract start"              — Contract Start Date
  "contract term"               — Term of Contract (No of Years)
  "contract expiry"             — Current Expiry Date / Current Exiry Date
  "remarks"                     — Remarks if any
  "cross connect"               — Cross connect (Bangalore only)
  "arc"                         — ARC (Rabale T1/T2 only)

Identity / Classification:
  "floor"                       — Floor / FLOOR / Floor / Module
  "floor module"                — Floor / Module
  "customer name"               — Customer Name / Customer
  "rhs sh"                      — RHS/SH / SH / RHS / SHS
  "power subscription model"    — Power Subscription Model (Rated/Subscribed)
  "power usage model"           — Power Usage Model (Bundled / Metered)
  "subscription mode"           — Subscription Mode / Subscription Mode (Rack/U Space/SqFt Space)
  "caged uncaged"               — Caged /Uncaged
  "uom"                         — UoM / UoM (KVA/KW) / UoM (VAULT/Chamber)
  "ownership"                   — Ownership (Sify/Customer)        [Airoli]
  "ir date"                     — IR DATE                          [Noida-02]
  "sitting space subscription"  — Sitting Space (Subscription)     [Noida]
  "enclosed shared"             — Enclosed/Shared                  [Airoli]
  "usage model"                 — Usage Model / Usage Model (Bundled/Metered)
  "unit rate model"             — Unit rate Model (Fixed/Variable) / Uit rate Model (Fixed/Variable)
  "multiplier"                  — Multiplier
  "onsite tape"                 — ONSITE TAPE ROTATION             [Airoli]
  "offsite tape"                — OFFSITE TAPE ROTATION            [Airoli]
  "safe vault"                  — SAFE VAULT                       [Airoli]
  "store space"                 — STORE SPACE                      [Airoli]

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
| which rows have <col> = <val> / find     → type="cell_lookup", operation=null |

## LOCATION ALIASES  (resolve BROADLY — always include all sub-locations)
- "airoli"                          → ["airoli"]
- "vashi"                           → ["vashi"]
- "rabale" / "rabale tower"         → ["rabale"]   (T1 + T2 + Tower4 + Tower5)
- "rabale t1" / "rabale tower 1"    → ["rabale t1"]
- "rabale t2" / "rabale tower 2"    → ["rabale t2"]
- "rabale tower 4"                  → ["rabale tower 4"]
- "rabale tower 5" / "rabale t5"    → ["rabale tower 5"]
- "noida"                           → ["noida"]    (Noida_01 + Noida_02)
- "noida 01" / "noida-01"           → ["noida 01"]
- "noida 02" / "noida-02"           → ["noida 02"]
- "bangalore" / "bengaluru"         → ["bangalore"]
- "chennai"                         → ["chennai"]
- "kolkata" / "calcutta"            → ["kolkata"]
- "mumbai" / "mumbai region"        → ["airoli","rabale","vashi"]
- no location mentioned             → null  (query ALL 10 files)

## FILE / SHEET ALIASES  (for the new "files" and "sheets" JSON fields)
Resolve a user phrase to the exact file name from the SCOPE list above:
- "airoli file" / "airoli tracker"          → Customer_and_Capacity_Tracker_Airoli_15Mar26.xlsx
- "bangalore file"                          → Customer_and_Capacity_Tracker_Bangalore_01_15Feb26.xlsx
- "chennai file"                            → Customer_and_Capacity_Tracker_Chennai_01_15Feb26.xls
- "kolkata file"                            → Customer_and_Capacity_Tracker_Kolkata_15Feb26.xlsx
- "noida 01 file"                           → Customer_and_Capacity_Tracker_Noida_01_15Feb26.xlsx
- "noida 02 file"                           → Customer_and_Capacity_Tracker_Noida_02_15Feb26.xlsx
- "rabale t1 t2 file"                       → Customer_and_Capacity_Tracker_Rabale_T1_T2_15Mar26.xlsx
- "rabale tower 4 file"                     → Customer_and_Capacity_Tracker_Rabale_Tower_4_15Mar26.xlsx
- "rabale tower 5 file"                     → Customer_and_Capacity_Tracker_Rabale_Tower_5_15Mar26.xlsx
- "vashi file"                              → Customer_and_Capacity_Tracker_Vashi_15Mar26.xls

Sheet names must be resolved against the exact list in SCOPE — examples:
- "customer details sheet in bangalore"     → files=["..._Bangalore_01_..."], sheets=["Customer details"]
- "new summary in bangalore"                → sheets=["NEW SUMMARY"]
- "t5 summary"                              → sheets=["T5 SUMMARY"]
- "rabale t1 sheet"                         → sheets=["Rabale-T1"]
- "terminated sheet in noida"               → sheets=["Terminated"]
- "disconnection details"                   → sheets=["Disconnection details"]

## FILTER SEMANTICS
- "caged customers"      → filter.caged = true
- "uncaged customers"    → filter.uncaged = true
- "rated customers"      → filter.rated = true (Power Subscription Model = Rated)
- "subscribed customers" → filter.subscribed = true (Power Subscription Model = Subscribed)
- "metered customers"    → filter.metered = true (Power Usage Model = Metered)
- "bundled customers"    → filter.bundled = true (Power Usage Model = Bundled)
- "rhs customers"        → filter.rhs = true
- "shs customers"        → filter.shs = true
- Combine with AND: caged + rated → filter.caged=true AND filter.rated=true
- "all customers" (no qualifier) → all filters null

## FILTER PROPAGATION RULE
If sub-query N is a "list" with a filter, and sub-query N+1 is an "aggregate"
on the SAME subject with no explicit filter, inherit the filter from N.

## COMPLEX QUERY DECOMPOSITION
Queries joined by "and", "or", commas, semicolons, or "also" = multiple
operations in the array. Execute each independently on the filtered dataset.
- "and" between filters   = intersection (BOTH must be true)
- "or" between locations  = union (include rows from EITHER location)
- "and" between actions on the same filter = same filter, multiple op objects

## CUSTOMER / LOCATION / VALUE STRICT RULES
(identical to bmprompt — filter to the matched set ONLY, never append
fallback or default-location rows, never print a trailing "— N row(s)"
footer that lists unrelated customers)

## =============================================================================
## CELL-VALUE QUERY RULES
## =============================================================================
Use type = "cell_lookup" whenever the user names a specific CELL VALUE in a
specific sub-header and wants the rows that match it. Examples:
- "show all rows where Floor = 3rd Floor"
- "list customers whose UoM is KVA"
- "find all entries where Caged/Uncaged = Caged in Bangalore"
- "which customers are on RHS in Airoli"
- "show rows where Power Subscription Model is Rated"
- "give me everything where Ownership = Customer"   (Airoli)
- "customers whose Subscription Mode = Rack"
- "rows where Unit rate Model = Fixed"
- "rows where Billing Frequency = Monthly in Kolkata"
- "show entries where Capacity to be given > 0 in Noida"

### Required JSON fields for cell_lookup
- "type"               : "cell_lookup"
- "target_column_hint" : the sub-header in the user's own words (e.g.
                         "floor", "uom", "caged/uncaged", "ownership",
                         "subscription mode", "power subscription model",
                         "unit rate model", "rhs/sh", "enclosed/shared",
                         "billing frequency", "per unit rate (mrc)",
                         "capacity to be given"). Executor fuzzy-maps this
                         via the LOCATION-WISE COLUMN SCHEMA.
- "cell_value"         : the exact value the user typed (e.g. "3rd Floor",
                         "KVA", "Caged", "Rated", "Rack", "Customer",
                         "Fixed", "RHS", "Monthly").
- "match_mode"         : "exact" (default) | "contains" | "regex"
- "return_columns"     : OPTIONAL list of columns to return. If null,
                         the executor returns a safe default projection:
                         Source_File, Source_Sheet, Customer Name,
                         Floor/Module, Caged/Uncaged, UoM,
                         Total Capacity Purchased, Capacity in Use,
                         plus the target column itself.
- "location"           : honours LOCATION ALIASES — null = scan all 10 files.
- "files" / "sheets"   : optional tight scope — see FILE / SHEET ALIASES.
- "filter"             : may be combined with cell_lookup.
- "customer_name"      : may be combined with cell_lookup.

### Executor semantics for cell_lookup
1. Load every sheet of every file in scope (respecting location / files /
   sheets filters if any).
2. For each sheet, fuzzy-map target_column_hint → the real column name in
   that sheet's detected header row, using the LOCATION-WISE COLUMN SCHEMA
   as the primary map. If the column does not exist in that sheet, SKIP
   the sheet silently (do not fabricate a column).
3. Normalise both the cell and the query value: strip whitespace, collapse
   internal whitespace, casefold. Numeric-looking values are compared as
   float after _robust_to_numeric.
4. Apply match_mode:
   - "exact"    : normalised cell == normalised query value
   - "contains" : normalised query is a substring of normalised cell
   - "regex"    : re.search(query, cell, re.I)
5. Keep ONLY rows where the match succeeds. Never include non-matching rows.
6. De-duplicate across sheets/files using the key in CRITICAL RULE 10.
7. Project to return_columns (or the safe default). Always include
   Source_File and Source_Sheet so the user can trace each row.
8. If ZERO rows match across every sheet of every file in scope, return a
   single empty result object with label
   "No matching record found for <target_column_hint> = <cell_value>".
   Do NOT fall back to any other dataset. Do NOT show unrelated rows.
9. NEVER invent a column value that was not in the source cell. NEVER fuse
   cells from different rows.

### Forbidden behaviours
- Returning rows whose target column value does NOT match cell_value.
- Returning the same row twice because it appears in two sheets.
- Returning rows from a location / file / sheet the user did not ask for.
- Returning a "closest guess" row when no exact match exists.
- Adding a trailing "— N row(s)" footer that lists extraneous customers.
- Silently dropping a sheet that DOES contain the target column because
  it is in an odd file.

## UNIT AWARENESS  (inform the label — do NOT change field_hint)
- Power/capacity → KW or KVA (varies per row's UoM column)
- Space          → Sq Ft
- Racks/seating  → Racks or Seats
- Revenue        → ₹/month (MRC)
- Rate           → ₹/KW-HR
Always include the unit in the label string (e.g. "Total Power Purchased (KW)").

## =============================================================================
## EXAMPLES
## =============================================================================

Query: "show customer name"
→ show only the particular customer's details. No all-customers list.
  No trailing "📋 <Customer> — N row(s)".

Query: "sum of power purchased"
→ [{"id":"op1","type":"aggregate","label":"Total Power Purchased (KW/KVA)","filter":null,"location":null,"files":null,"sheets":null,"operation":"sum","field_hint":"total capacity purchased","top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "total revenue by location"
→ [{"id":"op1","type":"aggregate","label":"Total Revenue (₹/month) by Location","filter":null,"location":null,"files":null,"sheets":null,"operation":"sum","field_hint":"total revenue","top_n":null,"group_by_location":true,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "list caged customers in noida"
→ [{"id":"op1","type":"list","label":"Caged Customers — Noida","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":["noida"],"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "show Wipro customer details"
→ [{"id":"op1","type":"list","label":"Wipro Customer Details","filter":{"caged":null,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":"Wipro","cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

### Cell-value examples — matched against LOCATION-WISE COLUMN SCHEMA

Query: "show all rows where Floor = 3rd Floor"
→ [{"id":"op1","type":"cell_lookup","label":"Rows where Floor = 3rd Floor","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"3rd Floor","target_column_hint":"floor","return_columns":null,"match_mode":"exact"}]

Query: "list all customers whose UoM is KVA"
→ [{"id":"op1","type":"cell_lookup","label":"Customers with UoM = KVA","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"KVA","target_column_hint":"uom","return_columns":["Source_File","Source_Sheet","Customer Name","UoM","Total Capacity Purchased","Capacity in Use"],"match_mode":"exact"}]

Query: "find rows where Caged/Uncaged = Caged in Bangalore Customer details sheet"
→ [{"id":"op1","type":"cell_lookup","label":"Caged rows — Bangalore (Customer details)","filter":null,"location":["bangalore"],"files":["Customer_and_Capacity_Tracker_Bangalore_01_15Feb26.xlsx"],"sheets":["Customer details"],"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Caged","target_column_hint":"caged/uncaged","return_columns":null,"match_mode":"exact"}]

Query: "which customers have Power Subscription Model = Rated in Airoli"
→ [{"id":"op1","type":"cell_lookup","label":"Rated rows — Airoli","filter":null,"location":["airoli"],"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Rated","target_column_hint":"power subscription model","return_columns":["Source_File","Source_Sheet","Customer Name","Power Subscription Model (Rated/Subscribed)","UoM (KVA/KW)","Total Capacity Purchased"],"match_mode":"exact"}]

Query: "rows where Subscription Mode contains Rack"
→ [{"id":"op1","type":"cell_lookup","label":"Rows where Subscription Mode contains 'Rack'","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Rack","target_column_hint":"subscription mode","return_columns":null,"match_mode":"contains"}]

Query: "show Wipro rows where UoM = KVA"
→ [{"id":"op1","type":"cell_lookup","label":"Wipro rows where UoM = KVA","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":"Wipro","cell_value":"KVA","target_column_hint":"uom","return_columns":null,"match_mode":"exact"}]

Query: "rows where Ownership = Customer in Airoli and sum their capacity in use"
→ [
  {"id":"op1","type":"cell_lookup","label":"Ownership = Customer — Airoli","filter":null,"location":["airoli"],"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Customer","target_column_hint":"ownership","return_columns":null,"match_mode":"exact"},
  {"id":"op2","type":"aggregate","label":"Capacity in Use — Ownership=Customer, Airoli","filter":null,"location":["airoli"],"files":null,"sheets":null,"operation":"sum","field_hint":"power in use","top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Customer","target_column_hint":"ownership","return_columns":null,"match_mode":"exact"}
]

Query: "rows where Billing Frequency = Monthly in Kolkata"
→ [{"id":"op1","type":"cell_lookup","label":"Billing Frequency = Monthly — Kolkata","filter":null,"location":["kolkata"],"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Monthly","target_column_hint":"billing frequency","return_columns":null,"match_mode":"exact"}]

Query: "rows where Unit rate Model = Fixed across all locations"
→ [{"id":"op1","type":"cell_lookup","label":"Unit rate Model = Fixed (All Locations)","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Fixed","target_column_hint":"unit rate model","return_columns":null,"match_mode":"exact"}]

Query: "show RHS = YES rows in Airoli"
→ [{"id":"op1","type":"cell_lookup","label":"RHS = YES — Airoli","filter":null,"location":["airoli"],"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"YES","target_column_hint":"rhs sh","return_columns":null,"match_mode":"exact"}]

Query: "list all caged customers AND sum capacity in use AND total power used AND list rated customers AND show customers in airoli or noida"
→ [
  {"id":"op1","type":"list","label":"Caged Customers (All Locations)","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null},
  {"id":"op2","type":"aggregate","label":"Capacity in Use — Caged (KW/KVA)","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"files":null,"sheets":null,"operation":"sum","field_hint":"power in use","top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null},
  {"id":"op3","type":"aggregate","label":"Total Power Purchased — Caged (KW/KVA)","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"files":null,"sheets":null,"operation":"sum","field_hint":"total capacity purchased","top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null},
  {"id":"op4","type":"list","label":"Rated Customers (All Locations)","filter":{"caged":null,"uncaged":null,"rated":true,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null},
  {"id":"op5","type":"aggregate","label":"Total Power Used — Rated (KW)","filter":{"caged":null,"uncaged":null,"rated":true,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"files":null,"sheets":null,"operation":"sum","field_hint":"power in use","top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null},
  {"id":"op6","type":"list","label":"Customers in Airoli or Noida","filter":{"caged":null,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":["airoli","noida"],"files":null,"sheets":null,"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}
]

Query: "how many customers per location"
→ [{"id":"op1","type":"count","label":"Customer Count by Location","filter":null,"location":null,"files":null,"sheets":null,"operation":"count","field_hint":null,"top_n":null,"group_by_location":true,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "top 5 customers by revenue"
→ [{"id":"op1","type":"aggregate","label":"Top 5 Customers by Revenue","filter":null,"location":null,"files":null,"sheets":null,"operation":"top","field_hint":"total revenue","top_n":5,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "minimum per unit rate in bangalore"
→ [{"id":"op1","type":"aggregate","label":"Min Per Unit Rate — Bangalore","filter":null,"location":["bangalore"],"files":null,"sheets":null,"operation":"min","field_hint":"per unit rate","top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "median power in use across all locations"
→ [{"id":"op1","type":"aggregate","label":"Median Power In Use (KW/KVA)","filter":null,"location":null,"files":null,"sheets":null,"operation":"median","field_hint":"power in use","top_n":null,"group_by_location":true,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "show sold vs available KW in Rabale Tower 5"
→ [{"id":"op1","type":"list","label":"Rabale Tower 5 — Sold vs Available (KW)","filter":null,"location":["rabale tower 5"],"files":["Customer_and_Capacity_Tracker_Rabale_Tower_5_15Mar26.xlsx"],"sheets":["T5 SUMMARY"],"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":["Source_File","Source_Sheet","Floor","Tower -5 (MUM - 03)","Total Capacity - Server Hall","Sold","Available","Remarks"],"match_mode":null}]

Query: "list racks subscribed vs in use in Rabale Tower 4"
→ [{"id":"op1","type":"list","label":"Rabale Tower 4 — Racks Subscribed vs In Use","filter":null,"location":["rabale tower 4"],"files":["Customer_and_Capacity_Tracker_Rabale_Tower_4_15Mar26.xlsx"],"sheets":["Sheet1"],"operation":null,"field_hint":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":["Source_File","Source_Sheet","Floor / Module","Customer Name","Subscription (No. of Racks)","In Use (No. of Racks)","Total Capacity Purchased (KW)","Capacity in Use (KW)"],"match_mode":null}]
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
    "sum":           lambda s: s.sum(),
    "total":         lambda s: s.sum(),
    "avg":           lambda s: s.mean(),
    "mean":          lambda s: s.mean(),
    "average":       lambda s: s.mean(),
    "min":           lambda s: s.min(),
    "minimum":       lambda s: s.min(),
    "max":           lambda s: s.max(),
    "maximum":       lambda s: s.max(),
    "count":         lambda s: float(len(s)),
    "std":           lambda s: s.std(ddof=1),
    "stddev":        lambda s: s.std(ddof=1),
    "median":        lambda s: s.median(),
    "variance":      lambda s: s.var(ddof=1),
    "var":           lambda s: s.var(ddof=1),
    "range":         lambda s: s.max() - s.min(),
    "spread":        lambda s: s.max() - s.min(),
    "sum_abs":       lambda s: s.abs().sum(),
    "count_nonzero": lambda s: float((s != 0).sum()),
    "pct_nonzero":   lambda s: float((s != 0).sum()) / max(len(s), 1) * 100,
    "multiply":      lambda s: s.prod(),
    "product":       lambda s: s.prod(),
    "percentage":    lambda s: s.sum(),
    "cumulative":    lambda s: s.sum(),
}




# ─────────────────────────────────────────────────────────────────────────────
# CANONICAL COLUMN REGISTRY  (v3 — maps all 38 canonical two-level names to
# real sub-header regex patterns as produced by _build_cols() from all 10
# Sify DC Excel files across all archetypes: CUSTOMER_DETAILS, AIROLI_BANNER,
# CAPACITY_GRID, FACILITY_GRID, SINGLE_LEVEL, CUSTOM_SUMMARY)
# ─────────────────────────────────────────────────────────────────────────────
_CANONICAL_REGISTRY: "dict[str, list[str]]" = {
    # ── BILLING MODEL band ──────────────────────────────────────────────────
    "Billing Model | Unnamed: 1_level_1": [
        r"billing model.*power subscription",
        r"power subscription model",
        r"billing model.*subscription",
        r"billing model.*model",
        r"billing model",
    ],
    "Billing Model | Rated": [
        r"billing model.*power subscription",
        r"power subscription model",
        r"billing model.*subscription",
        r"billing model.*model",
        r"billing model",
    ],
    "Billing Model | Subscribed": [
        r"billing model.*power subscription",
        r"power subscription model",
        r"billing model.*subscription",
        r"billing model.*model",
        r"billing model",
    ],
    "Billing Model | Metered": [
        r"power usage model",
        r"billing model.*usage",
        r"metered.*model",
        r"usage model",
    ],
    # ── SPACE band ──────────────────────────────────────────────────────────
    "Space | Seating Space": [
        r"seating space.*subscription",
        r"space.*seating.*subscription",
        r"seating space",
        r"seating.*subscription",
    ],
    "Space | Unnamed: 5_level_1": [
        r"subscription mode",
        r"space.*subscription.*mode",
        r"space.*mode",
    ],
    "Space | Floor": [
        r"space.*floor.*module",
        r"space.*floor",
        r"floor.*module",
        r"floor",
    ],
    "Space | Caged/Uncaged": [
        r"caged\s*/\s*uncaged",
        r"caged.*uncaged",
        r"space.*caged",
        r"caged",
    ],
    # ── POWER CAPACITY band ─────────────────────────────────────────────────
    "Power Capacity | Contracted": [
        r"power capacity.*total capacity purchased",
        r"total capacity purchased",
        r"power capacity.*contracted",
        r"contracted.*capacity",
        r"total.*purchased",
    ],
    "Power Capacity | Consumed": [
        r"power capacity.*capacity in use",
        r"capacity in use",
        r"power capacity.*consumed",
        r"power.*consumed",
        r"in use.*kw",
    ],
    "Power Capacity | Available": [
        r"power capacity.*capacity to be given",
        r"capacity to be given",
        r"power capacity.*available",
        r"power.*available",
        r"to be given",
    ],
    "Power Capacity | Unnamed: 3_level_1": [
        r"power capacity.*subscription.*kw",
        r"power capacity.*kw.*kva",
        r"power capacity.*subscription",
        r"subscription.*kw",
    ],
    "Power Capacity | Rated Load": [
        r"power capacity.*rated.*load",
        r"power capacity.*subscription.*model.*value",
        r"rated.*load",
        r"subscription.*model.*value",
        r"rated.*capacity",
    ],
    "Power Capacity | Actual Load": [
        r"power capacity.*actual.*load",
        r"actual.*load.*kw",
        r"power capacity.*usage.*kw",
        r"actual.*usage",
        r"actual load",
    ],
    "Power Capacity | KW-HR/Month": [
        r"power capacity.*no.*of.*units.*kw.*hr",
        r"no.*of.*units.*kw.*hr",
        r"kw.*hr.*month",
        r"units.*kw.*hr",
        r"kw-hr.*month",
    ],
    "Power Capacity | Unit Rate": [
        r"power capacity.*unit rate.*per.*kw",
        r"unit rate.*per.*kw",
        r"per unit rate.*mrc",
        r"per.*kw.*hr",
        r"unit rate",
    ],
    "Power Capacity | No. of Units": [
        r"power capacity.*no\.?\s*of\s*units",
        r"no\.?\s*of\s*units.*kw",
        r"number.*of.*units",
        r"no.*units",
    ],
    "Power Capacity | Raw Power (Genset)": [
        r"power capacity.*raw power.*genset",
        r"raw power.*genset.*transformer",
        r"raw power.*genset",
        r"genset.*kw",
        r"generator.*kw",
    ],
    "Power Capacity | Raw Power (Transformer)": [
        r"power capacity.*raw power.*transformer",
        r"raw power.*transformer",
        r"transformer.*kw",
        r"utility.*sanction.*load.*kva",
        r"utility.*sanction",
    ],
    "Power Capacity | Raw Power (Demand)": [
        r"power capacity.*raw power.*demand",
        r"raw power.*demand",
        r"contract.*demand",
        r"utility.*demand",
        r"sanction.*demand",
    ],
    # ── ACTUAL PUE ──────────────────────────────────────────────────────────
    "Actual PUE | Power Usage": [
        r"actual pue.*power usage",
        r"actual pue",
        r"pue.*power",
        r"\bpue\b",
    ],
    # ── RATED TO CONSUMED ───────────────────────────────────────────────────
    "Rated to Consumed | Ratio": [
        r"rated to consumed.*ratio",
        r"rated.*consumed.*ratio",
        r"rated to consumed",
        r"rated.*consumed",
    ],
    "Rated to Consumed | Unnamed: X_level_1": [
        r"rated to consumed",
        r"rated.*consumed",
    ],
    # ── GENSET HR/MO ────────────────────────────────────────────────────────
    "Genset Hr/Mo | Seating Space": [
        r"genset hr.*mo.*seating",
        r"genset hr.*seating",
        r"genset hr.*mo",
        r"genset hr",
        r"genset.*hour.*month",
    ],
    "Genset Hr/Mo | Unnamed: X_level_1": [
        r"genset hr.*mo",
        r"genset hr",
        r"genset.*hour.*month",
    ],
    # ── REVENUE band ────────────────────────────────────────────────────────
    "Revenue | Monthly": [
        r"revenue.*total revenue",
        r"total revenue",
        r"revenue.*monthly",
        r"monthly.*revenue",
        r"revenue.*mrc",
    ],
    "Additional Charges | MRC": [
        r"additional.*charges.*mrc",
        r"additional capacity charges",
        r"additional.*capacity.*mrc",
        r"additional charges",
        r"additional.*mrc",
    ],
    "Multiplier | Unnamed: X_level_1": [
        r"multiplier",
        r"revenue.*multiplier",
    ],
    # ── CAPACITY band (Rabale T1/T2 CAPACITY_GRID) ──────────────────────────
    "Capacity | Total Purchased": [
        r"capacity.*maximum usable",
        r"maximum usable capacity",
        r"capacity.*total.*purchased",
        r"total.*purchased.*capacity",
        r"max.*usable.*capacity",
    ],
    "Capacity | In Use": [
        r"capacity.*current utilization",
        r"current utilization",
        r"capacity.*in use",
        r"utilization.*kw",
    ],
    "Capacity | Reserved": [
        r"capacity.*committed",
        r"committed.*confirmed",
        r"capacity.*reserved",
        r"reserved.*capacity",
        r"committed",
    ],
    "Capacity | Surplus": [
        r"capacity.*surplus",
        r"surplus.*balance",
        r"surplus",
        r"capacity.*balance.*positive",
    ],
    "Capacity | Leakage": [
        r"capacity.*leakage",
        r"leakage.*balance",
        r"leakage",
        r"capacity.*balance.*negative",
    ],
    # ── IDENTITY / LAYOUT columns (all archetypes) ──────────────────────────
    "Customer Name | Unnamed: X_level_1": [
        r"customer.*name.*customer",
        r"customer name",
        r"client name",
        r"customer.*name",
    ],
    "Floor | Unnamed: X_level_1": [
        r"floor.*module",
        r"\bfloor\b",
    ],
    "Module | Unnamed: X_level_1": [
        r"floor.*module",
        r"\bmodule\b",
        r"\bfloor\b",
    ],
    "Description | Unnamed: X_level_1": [
        r"\bdescription\b",
        r"facility.*description",
    ],
    "Remarks | Unnamed: X_level_1": [
        r"remarks.*if.*any",
        r"\bremarks\b",
        r"remark",
    ],
}


def execute_ai_operations(operations: list, df: pd.DataFrame) -> list:
    results = []

    for op in operations:
        op_type           = op.get("type", "list")
        label             = op.get("label", "Result")
        filter_dict       = op.get("filter") or {}
        locations         = op.get("location")
        operation         = (op.get("operation") or "sum").lower().strip()
        field_hint        = op.get("field_hint") or ""
        top_n             = int(op.get("top_n") or 10)
        grp_by_loc        = bool(op.get("group_by_location"))
        canonical_columns = op.get("canonical_columns") or []
        files_scope       = op.get("files") or []
        sheets_scope      = op.get("sheets") or []
        customer_name     = op.get("customer_name") or ""
        cell_value        = op.get("cell_value") or ""
        target_col_hint   = op.get("target_column_hint") or ""
        match_mode        = (op.get("match_mode") or "contains").lower().strip()
        return_columns    = op.get("return_columns") or []

        filtered = _apply_op_filters(df, filter_dict, locations)

        # Apply file/sheet scope restrictors
        if files_scope and "_Location" in filtered.columns:
            file_mask = pd.Series(False, index=filtered.index)
            for fkw in files_scope:
                file_mask |= filtered["_Location"].str.lower().str.contains(
                    re.escape(fkw.lower()), na=False)
            filtered = filtered[file_mask]

        if sheets_scope and "_Sheet" in filtered.columns:
            sheet_mask = pd.Series(False, index=filtered.index)
            for skw in sheets_scope:
                sheet_mask |= filtered["_Sheet"].str.lower().str.contains(
                    re.escape(skw.lower()), na=False)
            filtered = filtered[sheet_mask]

        # Apply customer name filter
        if customer_name:
            cname_col = find_col(filtered, r"customer.*name|client.*name")
            if cname_col and cname_col in filtered.columns:
                filtered = filtered[
                    filtered[cname_col].astype(str).str.lower().str.contains(
                        re.escape(customer_name.lower()), na=False)]

        if filtered.empty:
            results.append({"type": "empty", "label": label,
                            "message": "No records match this filter."})
            continue

        # ── CANONICAL COLUMN RESOLUTION ────────────────────────────────────────
        def _resolve_canonical_column(df_: pd.DataFrame, canonical: str) -> "str | None":
            """
            Resolve a canonical two-level name 'Parent | Sub' to an actual DataFrame
            column. Resolution order:
              1. Exact match (column already named canonically)
              2. _CANONICAL_REGISTRY regex patterns (precise per-archetype mapping)
              3. Fuzzy parent+sub token matching (robust fallback)
            Never fabricates — returns None if nothing matches.
            """
            # 1. Exact column name match
            if canonical in df_.columns:
                return canonical

            # 2. Registry-driven regex matching (ordered by specificity)
            patterns = _CANONICAL_REGISTRY.get(canonical, [])
            for pat in patterns:
                for c in df_.columns:
                    try:
                        if re.search(pat, c, re.I):
                            return c
                    except re.error:
                        pass

            # 3. Fuzzy fallback using parent / sub token overlap
            if " | " in canonical:
                parent_raw, sub_raw = canonical.split(" | ", 1)
                parent = parent_raw.strip().lower()
                sub    = sub_raw.strip().lower()
                is_unnamed_sub = (sub.startswith("unnamed") or
                                  sub.startswith("x_level") or
                                  "level_1" in sub)

                # 3a. Parent contains AND (sub contains OR sub is unnamed artifact)
                for c in df_.columns:
                    cl = c.lower()
                    if parent in cl:
                        if is_unnamed_sub or sub in cl:
                            return c

                # 3b. Parent-only when sub is an Unnamed artifact
                if is_unnamed_sub:
                    for c in df_.columns:
                        if parent in c.lower():
                            return c

                # 3c. Token overlap — any parent word AND any meaningful sub word
                parent_words = [w for w in re.split(r"\W+", parent) if len(w) > 2]
                sub_words    = [w for w in re.split(r"\W+", sub) if len(w) > 2
                                and w not in ("unnamed", "level", "level1")]
                for c in df_.columns:
                    cl = c.lower()
                    p_match = any(w in cl for w in parent_words)
                    s_match = any(w in cl for w in sub_words) if sub_words else True
                    if p_match and s_match:
                        return c

                # 3d. Sub-word-only last resort for strong multi-word sub-headers
                if sub_words and len(sub_words) >= 2:
                    for c in df_.columns:
                        if sum(1 for w in sub_words if w in c.lower()) >= 2:
                            return c
            else:
                # Single-level canonical — keyword token match
                kw    = canonical.strip().lower()
                words = [w for w in re.split(r"\W+", kw) if len(w) > 2]
                for c in df_.columns:
                    if any(w in c.lower() for w in words):
                        return c
            return None

        # ── COLUMN_FETCH ──────────────────────────────────────────────────────
        if op_type == "column_fetch":
            if not canonical_columns:
                results.append({"type": "error", "label": label,
                                "message": "column_fetch requires canonical_columns."})
                continue

            # Build per-sheet output
            sheet_frames = []
            for loc in filtered["_Location"].unique() if "_Location" in filtered.columns else [""]:
                loc_df = filtered[filtered["_Location"] == loc] if "_Location" in filtered.columns else filtered
                for sn in loc_df["_Sheet"].unique() if "_Sheet" in loc_df.columns else [""]:
                    sn_df = loc_df[loc_df["_Sheet"] == sn] if "_Sheet" in loc_df.columns else loc_df
                    if sn_df.empty:
                        continue
                    row_data = {}
                    any_resolved = False
                    for canonical in canonical_columns:
                        real_col = _resolve_canonical_column(sn_df, canonical)
                        if real_col:
                            row_data[canonical] = sn_df[real_col].values
                            any_resolved = True
                        else:
                            row_data[canonical] = [""] * len(sn_df)
                    if not any_resolved:
                        continue
                    chunk = pd.DataFrame(row_data)
                    if "_Location" in sn_df.columns:
                        chunk.insert(0, "Source_File", sn_df["_Location"].values)
                    if "_Sheet" in sn_df.columns:
                        chunk.insert(1, "Source_Sheet", sn_df["_Sheet"].values)
                    sheet_frames.append(chunk)

            if not sheet_frames:
                results.append({"type": "empty", "label": label,
                                "message": "No matching columns found in any sheet."})
                continue

            combined_chunk = pd.concat(sheet_frames, ignore_index=True)
            combined_chunk = combined_chunk.reset_index(drop=True)
            combined_chunk.index += 1
            results.append({"type": "table", "label": label,
                            "data": combined_chunk, "row_count": len(combined_chunk)})
            continue

        # ── CELL_LOOKUP ────────────────────────────────────────────────────────
        elif op_type == "cell_lookup":
            if not target_col_hint and not canonical_columns:
                results.append({"type": "error", "label": label,
                                "message": "cell_lookup requires target_column_hint."})
                continue

            # Resolve target column
            target_col = None
            if target_col_hint and target_col_hint.lower() != "any":
                # Try canonical resolution first
                target_col = _resolve_canonical_column(filtered, target_col_hint)
                if not target_col:
                    # Fuzzy fallback
                    target_col, _ = _resolve_col_by_semantic(filtered, target_col_hint)

            any_col_mode = (not target_col_hint or target_col_hint.lower() == "any")

            if not any_col_mode and (not target_col or target_col not in filtered.columns):
                results.append({"type": "empty", "label": label,
                                "message": f"Column '{target_col_hint}' not found in data."})
                continue

            # Match rows
            if any_col_mode:
                # Scan every column
                mask = pd.Series(False, index=filtered.index)
                matched_col_info = []
                for c in filtered.columns:
                    if c.startswith("_"):
                        continue
                    col_str = filtered[c].astype(str)
                    if match_mode == "exact":
                        col_mask = col_str.str.lower() == str(cell_value).lower()
                    elif match_mode == "regex":
                        try:
                            col_mask = col_str.str.contains(str(cell_value), flags=re.I, na=False)
                        except re.error:
                            col_mask = col_str.str.lower().str.contains(
                                re.escape(str(cell_value).lower()), na=False)
                    else:
                        col_mask = col_str.str.lower().str.contains(
                            re.escape(str(cell_value).lower()), na=False)
                    mask = mask | col_mask
                matched_df = filtered[mask].copy()
            else:
                col_str = filtered[target_col].astype(str)
                # Handle numeric/comparison operators in cell_value
                cv = str(cell_value).strip()
                if cv.startswith(">") or cv.startswith("<") or cv.startswith("="):
                    try:
                        num_series = _robust_to_numeric(filtered[target_col])
                        import operator as op_mod
                        ops_map = {
                            ">=": op_mod.ge, "<=": op_mod.le,
                            ">": op_mod.gt,  "<": op_mod.lt,  "=": op_mod.eq,
                        }
                        for sym, fn in sorted(ops_map.items(), key=lambda x: -len(x[0])):
                            if cv.startswith(sym):
                                num_val = float(cv[len(sym):].strip())
                                mask = fn(num_series, num_val)
                                matched_df = filtered[mask.fillna(False)].copy()
                                break
                        else:
                            matched_df = filtered[pd.Series(False, index=filtered.index)].copy()
                    except Exception:
                        matched_df = filtered[pd.Series(False, index=filtered.index)].copy()
                elif match_mode == "exact":
                    mask = col_str.str.lower() == cv.lower()
                    matched_df = filtered[mask].copy()
                elif match_mode == "regex":
                    try:
                        mask = col_str.str.contains(cv, flags=re.I, na=False)
                    except re.error:
                        mask = col_str.str.lower().str.contains(
                            re.escape(cv.lower()), na=False)
                    matched_df = filtered[mask].copy()
                else:
                    mask = col_str.str.lower().str.contains(
                        re.escape(cv.lower()), na=False)
                    matched_df = filtered[mask].copy()

            if matched_df.empty:
                results.append({
                    "type": "empty", "label": label,
                    "message": f"No rows found where '{target_col_hint}' = '{cell_value}'."})
                continue

            # Build projection
            if return_columns:
                proj_cols = []
                for rc in return_columns:
                    real = _resolve_canonical_column(matched_df, rc)
                    if real and real in matched_df.columns:
                        proj_cols.append(real)
                display_df = matched_df[
                    [c for c in ["_Location", "_Sheet"] if c in matched_df.columns] + proj_cols
                ].copy()
            else:
                meta = [c for c in ["_Location", "_Sheet"] if c in matched_df.columns]
                data = [c for c in matched_df.columns if not c.startswith("_")][:30]
                # Also add any resolved canonical columns
                for canonical in canonical_columns:
                    real = _resolve_canonical_column(matched_df, canonical)
                    if real and real not in data and real in matched_df.columns:
                        data.append(real)
                display_df = matched_df[meta + data].copy()

            display_df = display_df.rename(columns={"_Location": "Source_File", "_Sheet": "Source_Sheet"})
            display_df = display_df.reset_index(drop=True)
            display_df.index += 1
            results.append({"type": "table", "label": label,
                            "data": display_df, "row_count": len(display_df)})
            continue

        # ── LIST ─────────────────────────────────────────────────────────────
        if op_type == "list":
            meta  = [c for c in ["_Location", "_Sheet"] if c in filtered.columns]
            data  = [c for c in filtered.columns if not c.startswith("_")][:30]
            # Include requested canonical columns
            for canonical in canonical_columns:
                real = _resolve_canonical_column(filtered, canonical)
                if real and real not in data and real in filtered.columns:
                    data.append(real)
            disp  = filtered[meta + data].reset_index(drop=True)
            disp.index += 1
            results.append({"type": "table", "label": label,
                            "data": disp, "row_count": len(disp)})

        # ── AGGREGATE / TOP / BOTTOM ─────────────────────────────────────────
        elif op_type in ("aggregate", "top", "bottom"):
            # First try canonical columns for field resolution
            col = None
            reason = ""
            if canonical_columns:
                for canonical in canonical_columns:
                    real = _resolve_canonical_column(filtered, canonical)
                    if real and real in filtered.columns:
                        col = real
                        reason = f"canonical '{canonical}' → '{real}'"
                        break

            if not col:
                col, reason = _resolve_col_by_semantic(filtered, field_hint)

            if not col or col not in filtered.columns:
                results.append({"type": "error", "label": label,
                                "message": f"No column matched '{field_hint}'."})
                continue

            unit   = _detect_unit(col)
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

            # Percentage: compute as % of total sum
            if operation in ("percentage", "percent", "pct"):
                total_sum = _robust_to_numeric(df[col]).sum() if col in df.columns else valid.sum()
                val = (valid.sum() / total_sum * 100) if total_sum else 0.0
                unit = "%"
            else:
                op_fn = _EXTENDED_OPS.get(operation)
                if op_fn is None:
                    for alias_key in ("sum",):
                        if alias_key in operation:
                            op_fn = _EXTENDED_OPS[alias_key]
                            break
                    if op_fn is None:
                        op_fn = _EXTENDED_OPS["sum"]
                val = op_fn(valid)

            # Per-location breakdown
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
            if operation in ("sum", "total", "avg", "average", "mean", "median", "std") \
                    and "_Location" in filtered.columns:
                grp2 = (
                    filtered.groupby("_Location")[col]
                    .apply(lambda x: _robust_to_numeric(x).sum()
                           if operation in ("sum", "total") else
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
    st.markdown('<div class="section-title">Operation Summary Dashboard</div>',
                unsafe_allow_html=True)

    if CUST.empty:
        st.warning("No data loaded. Please check your Excel files.")
    else:
        # ──────────────────────────────────────────────────────────────────────
        # Resolve all working columns from the combined customer dataframe
        # ──────────────────────────────────────────────────────────────────────
        cust_c     = find_col(CUST, r"customer.*name|client.*name|^customer$")
        caged_c    = find_col(CUST, r"\bcaged\b")
        own_c      = find_col(CUST, r"\brhs\b|\bshs\b|ownership")
        sub_mode_c = find_col(CUST, r"subscription.*mode\s*\(rack|space.*subscription.*mode|^subscription.*mode$")

        space_sub_c   = find_col(CUST, r"space\s*\|\s*subscription$|^space.*subscription$|^subscription$")
        space_inuse_c = find_col(CUST, r"space.*in.*use|^in.*use$")

        cap_c      = find_col(CUST, r"total.*capacity.*purchased|total.*capacity|capacity.*purchased")
        use_c      = find_col(CUST, r"capacity.*in.*use")
        rack_c     = find_col(CUST, r"^rack$|^racks$|\brack\b(?!.*space)")

        # ──────────────────────────────────────────────────────────────────────
        # Safe numeric helpers (column may be None or absent)
        # ──────────────────────────────────────────────────────────────────────
        def _nsum(df, col):
            if col and col in df.columns:
                v = _robust_to_numeric(df[col]).sum()
                try:
                    return float(v) if pd.notna(v) else 0.0
                except Exception:
                    return 0.0
            return 0.0

        # ──────────────────────────────────────────────────────────────────────
        # Per-location authoritative metrics pulled from the FACILITY / SUMMARY
        # sheets of each uploaded Excel file.  These are the source of truth
        # for Built / Installed / Operational / Available KW and Racks and for
        # White Space (sqft).  All values are read with openpyxl / pandas from
        # the actual files on disk – nothing is hard-coded.
        # ──────────────────────────────────────────────────────────────────────
        @st.cache_data(show_spinner=False)
        def _build_location_facility_metrics() -> pd.DataFrame:
            """
            Walk every configured excel dir, inspect the facility / summary
            sheet of every workbook and extract per-location KPIs.
            Returns one row per location with columns:
              Location, Built_kW, Installed_kW, Operational_kW, Available_kW,
              Built_Racks, Installed_Racks, Operational_Racks, Available_Racks,
              Built_Sqft, Operational_Sqft, Available_Sqft,
              Avg_PUE, Min_PUE, Max_PUE
            Missing values default to 0.
            """
            import openpyxl as _opx

            rows: "list[dict]" = []

            def _flt(v):
                """Coerce a cell to float; return 0.0 if not parseable."""
                if v is None:
                    return 0.0
                if isinstance(v, (int, float)):
                    try:
                        f = float(v)
                        return f if np.isfinite(f) else 0.0
                    except Exception:
                        return 0.0
                s = str(v).strip()
                if not s:
                    return 0.0
                # strip units / commas
                s2 = re.sub(r"[^\d\.\-]", "", s)
                if not s2 or s2 in (".", "-", "-."):
                    return 0.0
                try:
                    return float(s2)
                except Exception:
                    return 0.0

            def _label_from_fname(fname: str) -> str:
                stem = Path(fname).stem
                s = stem.replace("Customer_and_Capacity_Tracker_", "")
                s = re.sub(r"_\d{1,2}[A-Za-z]{3}\d{2,4}.*$", "", s)
                s = s.strip("_ ").replace("_", " ")
                return s or stem

            def _iter_source_files():
                seen = set()
                try:
                    dirs = _excel_dirs()
                except Exception:
                    dirs = [Path(".")]
                for d in dirs:
                    try:
                        if not d.exists():
                            continue
                        for p in sorted(d.iterdir()):
                            if p.suffix.lower() in (".xlsx", ".xls") and p.name not in seen:
                                seen.add(p.name)
                                yield p
                    except Exception:
                        continue

            for fpath in _iter_source_files():
                loc_label = _label_from_fname(fpath.name)
                metrics = dict(
                    Location=loc_label,
                    Built_kW=0.0, Installed_kW=0.0,
                    Operational_kW=0.0, Available_kW=0.0,
                    Built_Racks=0.0, Installed_Racks=0.0,
                    Operational_Racks=0.0, Available_Racks=0.0,
                    Built_Sqft=0.0, Operational_Sqft=0.0, Available_Sqft=0.0,
                    Avg_PUE=0.0, Min_PUE=0.0, Max_PUE=0.0,
                )

                try:
                    if fpath.suffix.lower() == ".xlsx":
                        wb = _opx.load_workbook(fpath, data_only=True)
                        sheets = {sn: wb[sn] for sn in wb.sheetnames}

                        # Collect all rows from all sheets as a flat list of
                        # tuples for text-based discovery of totals.
                        flat = []  # (sheet, row_index, values_tuple, row_label_lc)
                        for sn, ws in sheets.items():
                            for ri, row in enumerate(ws.iter_rows(values_only=True)):
                                vals = list(row)
                                first_text = ""
                                for v in vals:
                                    if v is not None and str(v).strip():
                                        first_text = str(v).strip().lower()
                                        break
                                flat.append((sn, ri, vals, first_text))

                        # ── PUE ─────────────────────────────────────────────
                        pue_vals = []
                        for sn, ri, vals, lbl in flat:
                            if "pue" in lbl and "rack" not in lbl:
                                for v in vals:
                                    f = _flt(v)
                                    if 1.0 < f < 5.0:
                                        pue_vals.append(f)
                                        break
                            # "TOTAL <loc>" row in NEW SUMMARY has PUE in last numeric col
                            if lbl.startswith("total ") and any(
                                (isinstance(c, str) and "pue" in str(c).lower())
                                for s2, r2, vs2, _ in flat if s2 == sn and r2 < ri
                                for c in vs2
                            ):
                                nums = [_flt(v) for v in vals if _flt(v) and 1.0 < _flt(v) < 5.0]
                                pue_vals.extend(nums)
                        pue_vals = [p for p in pue_vals if 1.0 < p < 5.0]
                        if pue_vals:
                            metrics["Avg_PUE"] = float(np.mean(pue_vals))
                            metrics["Min_PUE"] = float(np.min(pue_vals))
                            metrics["Max_PUE"] = float(np.max(pue_vals))

                        # ── Facility / Summary row search ──────────────────
                        # Look for the "Total" row in a Facility details sheet
                        # and extract: IT capacity Installed kW, IT Power Sold,
                        # Allocated IT Power, IT kW Usage, MAX Rack Capacity,
                        # Rack Space sold, Available space, Designed White
                        # Space, Used White Space, Available White Space.
                        header_row = None
                        header_idx = {}
                        fac_sheet = None
                        all_data_rows = []  # per-floor rows collected after header
                        total_row = None
                        for sn, ws in sheets.items():
                            snl = sn.lower()
                            if "facility" in snl or "summary" in snl or "inventory" in snl:
                                full_rows = list(ws.iter_rows(values_only=True))
                                hdr_ri = -1
                                for ri, row in enumerate(full_rows):
                                    joined = " ".join(
                                        str(c).lower() for c in row if c is not None
                                    )
                                    if ("it capacity installed" in joined
                                            or "rack capacity" in joined
                                            or "white space" in joined):
                                        header_row = [
                                            (str(c).strip().lower() if c is not None else "")
                                            for c in row
                                        ]
                                        for ci, h in enumerate(header_row):
                                            if h:
                                                header_idx[h] = ci
                                        fac_sheet = sn
                                        hdr_ri = ri
                                        break
                                if hdr_ri < 0:
                                    continue
                                # Gather candidate data / total rows after header
                                for ri2, row2 in enumerate(full_rows):
                                    if ri2 <= hdr_ri + 1:   # skip the subheader row
                                        continue
                                    first = ""
                                    for c in row2:
                                        if c is not None and str(c).strip():
                                            first = str(c).strip().lower()
                                            break
                                    if not first:
                                        continue
                                    if first.startswith("total"):
                                        total_row = row2
                                    else:
                                        # Count numeric cells — data rows have several
                                        num_ct = sum(1 for c in row2 if _flt(c) != 0)
                                        if num_ct >= 3:
                                            all_data_rows.append(row2)
                                break

                        if header_idx:
                            def _pick_from(row_, patterns):
                                for pat in patterns:
                                    for h, ci in header_idx.items():
                                        if re.search(pat, h):
                                            if ci < len(row_):
                                                f = _flt(row_[ci])
                                                if f:
                                                    return f
                                return 0.0

                            def _sum_from(patterns):
                                total = 0.0
                                for pat in patterns:
                                    col_idx = None
                                    for h, ci in header_idx.items():
                                        if re.search(pat, h):
                                            col_idx = ci
                                            break
                                    if col_idx is None:
                                        continue
                                    for r_ in all_data_rows:
                                        if col_idx < len(r_):
                                            total += _flt(r_[col_idx])
                                    if total:
                                        return total
                                return 0.0

                            def _resolve(patterns):
                                # Prefer Total row if it has a non-zero value; else sum data rows
                                if total_row is not None:
                                    v = _pick_from(total_row, patterns)
                                    if v:
                                        return v
                                return _sum_from(patterns)

                            metrics["Built_kW"]       = _resolve([r"it capacity installed", r"installed kw"])
                            metrics["Installed_kW"]   = metrics["Built_kW"]
                            metrics["Operational_kW"] = _resolve([r"it kw usage", r"kw usage", r"actual load kw"])
                            sold_kw = _resolve([r"it power sold", r"allocated it power", r"power sold"])
                            if sold_kw:
                                metrics["Available_kW"] = max(metrics["Built_kW"] - sold_kw, 0.0)
                            else:
                                metrics["Available_kW"] = max(metrics["Built_kW"] - metrics["Operational_kW"], 0.0)

                            metrics["Built_Racks"]       = _resolve([r"max rack capacity", r"rack capacity.*design", r"^design$"])
                            metrics["Installed_Racks"]   = metrics["Built_Racks"]
                            metrics["Operational_Racks"] = _resolve([r"rack space sold", r"racks which can be placed", r"^sold$"])
                            metrics["Available_Racks"]   = _resolve([r"^available space$", r"^available$"])

                            metrics["Built_Sqft"]       = _resolve([r"designed white space", r"design.*white.*space"])
                            metrics["Operational_Sqft"] = _resolve([r"used white space"])
                            metrics["Available_Sqft"]   = _resolve([r"avaialble white space", r"available white space"])

                        # ── Rabale T1/T2 archetype (Power Usage block) ─────
                        if metrics["Built_kW"] == 0:
                            for sn, ws in sheets.items():
                                snl = sn.lower()
                                if "rabale" in snl or "t1" in snl or "t2" in snl:
                                    # Row starting with "UPS Capacity" — sum max & current
                                    ups_max = 0.0
                                    ups_cur = 0.0
                                    space_max = 0.0
                                    space_cur = 0.0
                                    for row in ws.iter_rows(values_only=True):
                                        if not row or row[0] is None:
                                            continue
                                        lbl = str(row[0]).lower()
                                        if "ups capacity" in lbl:
                                            ups_max += _flt(row[1] if len(row) > 1 else 0)
                                            ups_cur += _flt(row[2] if len(row) > 2 else 0)
                                        elif lbl.startswith("space at") and "sq" in lbl:
                                            space_max += _flt(row[1] if len(row) > 1 else 0)
                                            space_cur += _flt(row[2] if len(row) > 2 else 0)
                                    if ups_max > 0:
                                        metrics["Built_kW"]        = ups_max
                                        metrics["Installed_kW"]    = ups_max
                                        metrics["Operational_kW"]  = ups_cur
                                        metrics["Available_kW"]    = max(ups_max - ups_cur, 0.0)
                                    if space_max > 0:
                                        metrics["Built_Sqft"]       = space_max
                                        metrics["Operational_Sqft"] = space_cur
                                        metrics["Available_Sqft"]   = max(space_max - space_cur, 0.0)

                        # ── T5 SUMMARY archetype: columns are Total Capacity, Sold, Available ──
                        if metrics["Built_kW"] == 0 and "T5 SUMMARY" in sheets:
                            ws = sheets["T5 SUMMARY"]
                            tot_cap = tot_sold = tot_avail = 0.0
                            for row in ws.iter_rows(values_only=True):
                                # columns (0-based): 7=Total Capacity, 8=Sold, 10=Available (per inspection)
                                if len(row) > 10:
                                    tot_cap   += _flt(row[7])
                                    tot_sold  += _flt(row[8])
                                    tot_avail += _flt(row[10])
                            if tot_cap > 0 or tot_sold > 0:
                                metrics["Built_kW"]       = tot_cap
                                metrics["Installed_kW"]   = tot_cap
                                metrics["Operational_kW"] = tot_sold
                                metrics["Available_kW"]   = tot_avail if tot_avail else max(tot_cap - tot_sold, 0.0)

                    else:
                        # .xls legacy path – fall back to pandas
                        try:
                            xls = pd.ExcelFile(fpath)
                            for sn in xls.sheet_names:
                                try:
                                    df_raw = pd.read_excel(xls, sheet_name=sn, header=None)
                                except Exception:
                                    continue
                                # Very light extraction – look for a row that
                                # contains "Total" as first non-null cell
                                for _, row in df_raw.iterrows():
                                    first = None
                                    for v in row:
                                        if pd.notna(v) and str(v).strip():
                                            first = str(v).strip().lower()
                                            break
                                    if first and first.startswith("total"):
                                        nums = [_flt(v) for v in row]
                                        nums = [n for n in nums if n > 0]
                                        if nums and metrics["Built_kW"] == 0:
                                            # pick the largest plausible kW-like value
                                            big = [n for n in nums if n > 50]
                                            if big:
                                                metrics["Built_kW"]       = max(big)
                                                metrics["Installed_kW"]   = max(big)
                                                metrics["Operational_kW"] = sorted(big)[len(big) // 2] if len(big) > 1 else max(big) * 0.8
                                                metrics["Available_kW"]   = max(metrics["Built_kW"] - metrics["Operational_kW"], 0.0)
                                        break
                        except Exception:
                            pass

                except Exception:
                    pass

                # ──────────────────────────────────────────────────────────
                # Fallback / supplement from CUSTOMER-level data in CUST:
                # For this specific file's location rows, aggregate
                # Total Capacity Purchased, Capacity in Use, and Subscription
                # (rack count).  Only used if facility-level not available.
                # ──────────────────────────────────────────────────────────
                try:
                    if "_Location" in CUST.columns:
                        loc_norm = loc_label.lower().split()[0]  # "rabale", "noida", etc.
                        mask = CUST["_Location"].astype(str).str.lower().str.contains(
                            re.escape(loc_norm), na=False
                        )
                        # If more specific (e.g. "noida 01" vs "noida 02"),
                        # refine by the second token if present
                        tokens = loc_label.lower().split()
                        if len(tokens) >= 2:
                            mask2 = CUST["_Location"].astype(str).str.lower().str.contains(
                                re.escape(" ".join(tokens[:2])), na=False
                            )
                            if mask2.any():
                                mask = mask2
                        sub = CUST[mask]
                        if not sub.empty:
                            if metrics["Built_kW"] == 0 and cap_c and cap_c in sub.columns:
                                metrics["Built_kW"]     = _nsum(sub, cap_c)
                                metrics["Installed_kW"] = metrics["Built_kW"]
                            if metrics["Operational_kW"] == 0 and use_c and use_c in sub.columns:
                                metrics["Operational_kW"] = _nsum(sub, use_c)
                            if metrics["Available_kW"] == 0:
                                metrics["Available_kW"] = max(
                                    metrics["Built_kW"] - metrics["Operational_kW"], 0.0
                                )
                            if metrics["Operational_Racks"] == 0 and space_sub_c and space_sub_c in sub.columns:
                                metrics["Operational_Racks"] = _nsum(sub, space_sub_c)
                            if metrics["Built_Racks"] == 0 and metrics["Operational_Racks"]:
                                metrics["Built_Racks"]     = metrics["Operational_Racks"]
                                metrics["Installed_Racks"] = metrics["Operational_Racks"]
                except Exception:
                    pass

                rows.append(metrics)

            if not rows:
                return pd.DataFrame(columns=[
                    "Location", "Built_kW", "Installed_kW", "Operational_kW",
                    "Available_kW", "Built_Racks", "Installed_Racks",
                    "Operational_Racks", "Available_Racks", "Built_Sqft",
                    "Operational_Sqft", "Available_Sqft",
                    "Avg_PUE", "Min_PUE", "Max_PUE",
                ])

            df = pd.DataFrame(rows)

            # ── Sanity clamps: protect against source-data header/col drift
            # (e.g. a "9000" value landing under an "Available space" header
            # when the real meaning was white-space sq-ft).  A location can
            # never have more Operational or Available racks than Built.
            def _clamp_racks(row):
                b = row["Built_Racks"]
                if b > 0:
                    if row["Operational_Racks"] > b:
                        row["Operational_Racks"] = b
                    if row["Available_Racks"] > b:
                        row["Available_Racks"] = max(b - row["Operational_Racks"], 0.0)
                return row
            df = df.apply(_clamp_racks, axis=1)

            def _clamp_kw(row):
                b = row["Built_kW"]
                if b > 0:
                    if row["Operational_kW"] > b * 1.5:
                        row["Operational_kW"] = b
                    if row["Available_kW"] > b:
                        row["Available_kW"] = max(b - row["Operational_kW"], 0.0)
                return row
            df = df.apply(_clamp_kw, axis=1)

            # Derived "Unconsumed" (allocated but not yet used) and "Actual"
            df["Unconsumed_kW"]    = (df["Built_kW"] - df["Operational_kW"]).clip(lower=0)
            df["Actual_kW"]        = df["Operational_kW"]
            df["Unconsumed_Racks"] = (df["Built_Racks"] - df["Operational_Racks"]).clip(lower=0)
            df["Actual_Racks"]     = df["Operational_Racks"]
            df["Wasted_Sqft"]      = (
                df["Built_Sqft"] - df["Operational_Sqft"] - df["Available_Sqft"]
            ).clip(lower=0)
            return df

        LOC_DF = _build_location_facility_metrics()

        # ──────────────────────────────────────────────────────────────────────
        # TOP ROW – headline KPI tiles (PUE / Power Usage / IT Load / Racks / SqFt)
        # ──────────────────────────────────────────────────────────────────────
        if LOC_DF.empty:
            st.info("No facility-level KPIs could be parsed from the uploaded files.")
        else:
            tot_built_kw  = float(LOC_DF["Built_kW"].sum())
            tot_inst_kw   = float(LOC_DF["Installed_kW"].sum())
            tot_oper_kw   = float(LOC_DF["Operational_kW"].sum())
            tot_avail_kw  = float(LOC_DF["Available_kW"].sum())

            tot_built_rk  = float(LOC_DF["Built_Racks"].sum())
            tot_inst_rk   = float(LOC_DF["Installed_Racks"].sum())
            tot_oper_rk   = float(LOC_DF["Operational_Racks"].sum())
            tot_avail_rk  = float(LOC_DF["Available_Racks"].sum())

            tot_built_sf  = float(LOC_DF["Built_Sqft"].sum())
            tot_oper_sf   = float(LOC_DF["Operational_Sqft"].sum())
            tot_avail_sf  = float(LOC_DF["Available_Sqft"].sum())

            pue_non_zero  = LOC_DF[LOC_DF["Avg_PUE"] > 0]
            if not pue_non_zero.empty:
                avg_pue = float(pue_non_zero["Avg_PUE"].mean())
                min_pue = float(pue_non_zero["Min_PUE"].replace(0, np.nan).min())
                max_pue = float(pue_non_zero["Max_PUE"].max())
                if not np.isfinite(min_pue):
                    min_pue = avg_pue
            else:
                avg_pue = min_pue = max_pue = 0.0

            # Compose 5 headline tiles identical in spirit to the reference dashboard
            pct_power = (tot_oper_kw / tot_built_kw * 100) if tot_built_kw > 0 else 0.0
            # "IT Load" ≈ sold/installed IT kW consumption (Actual) vs Installed
            pct_it    = (tot_oper_kw / tot_inst_kw * 100) if tot_inst_kw > 0 else 0.0

            hc = st.columns(5)
            hc[0].markdown(kpi_html(
                f"{avg_pue:.2f}" if avg_pue else "–",
                "Avg PUE",
                f"Min: {min_pue:.2f} &nbsp;•&nbsp; Max: {max_pue:.2f}" if avg_pue else "not in source files",
                GREEN if avg_pue and avg_pue < 2 else AMBER,
            ), unsafe_allow_html=True)

            hc[1].markdown(kpi_html(
                f"{tot_oper_kw:,.2f} kW",
                "Power Usage",
                f"{pct_power:.2f}% of {tot_built_kw:,.0f} kW",
                CYAN,
            ), unsafe_allow_html=True)

            hc[2].markdown(kpi_html(
                f"{tot_oper_kw:,.2f} kW",
                "IT Load",
                f"{pct_it:.2f}% of {tot_inst_kw:,.0f} kW",
                LBLUE,
            ), unsafe_allow_html=True)

            hc[3].markdown(kpi_html(
                f"{int(tot_avail_rk):,}",
                "Rack Space",
                "Racks Available",
                GREEN,
            ), unsafe_allow_html=True)

            hc[4].markdown(kpi_html(
                f"{tot_avail_sf:,.2f}",
                "White Space",
                "Sqft Available",
                AMBER,
            ), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ──────────────────────────────────────────────────────────────────
            # RACK POWER USAGE section
            # ──────────────────────────────────────────────────────────────────
            st.markdown('<div class="section-title">Rack Power Usage</div>',
                        unsafe_allow_html=True)
            pwr = st.columns(4)
            used_pct_pwr = (tot_oper_kw / tot_built_kw * 100) if tot_built_kw > 0 else 0.0
            free_pct_pwr = 100 - used_pct_pwr if tot_built_kw > 0 else 0.0
            pwr[0].markdown(kpi_html(
                f"{tot_built_kw:,.0f} kW", "Total Built", "All Floors", CYAN
            ), unsafe_allow_html=True)
            pwr[1].markdown(kpi_html(
                f"{tot_inst_kw:,.0f} kW", "Total Installed", "All Floors", LBLUE
            ), unsafe_allow_html=True)
            pwr[2].markdown(kpi_html(
                f"{tot_oper_kw:,.2f} kW", "Total Operational",
                f"{used_pct_pwr:.1f}% Used", AMBER
            ), unsafe_allow_html=True)
            pwr[3].markdown(kpi_html(
                f"{tot_avail_kw:,.2f} kW", "Total Available",
                f"{free_pct_pwr:.1f}% Free", GREEN
            ), unsafe_allow_html=True)

            # Bar chart per location – Built / Operational / Available / Unconsumed / Actual
            rpu_df = LOC_DF[LOC_DF["Built_kW"] > 0].copy().sort_values("Built_kW", ascending=True)
            if not rpu_df.empty:
                fig_rpu = go.Figure()
                fig_rpu.add_trace(go.Bar(name="Built",       x=rpu_df["Location"], y=rpu_df["Built_kW"],       marker_color=LBLUE))
                fig_rpu.add_trace(go.Bar(name="Operational", x=rpu_df["Location"], y=rpu_df["Operational_kW"], marker_color=RED))
                fig_rpu.add_trace(go.Bar(name="Available",   x=rpu_df["Location"], y=rpu_df["Available_kW"],   marker_color=GREEN))
                fig_rpu.add_trace(go.Bar(name="Unconsumed",  x=rpu_df["Location"], y=rpu_df["Unconsumed_kW"],  marker_color=AMBER))
                fig_rpu.add_trace(go.Bar(name="Actual",      x=rpu_df["Location"], y=rpu_df["Actual_kW"],      marker_color=CYAN))
                fig_rpu.update_layout(
                    **_base_layout(),
                    barmode="group",
                    height=380,
                    yaxis_title="kW",
                    xaxis_title="Location",
                    legend=dict(orientation="h", y=-0.25),
                )
                st.plotly_chart(fig_rpu, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ──────────────────────────────────────────────────────────────────
            # RACK SPACE section
            # ──────────────────────────────────────────────────────────────────
            st.markdown('<div class="section-title">Rack Space</div>',
                        unsafe_allow_html=True)
            rsp = st.columns(4)
            used_pct_rk = (tot_oper_rk / tot_built_rk * 100) if tot_built_rk > 0 else 0.0
            free_pct_rk = 100 - used_pct_rk if tot_built_rk > 0 else 0.0
            rsp[0].markdown(kpi_html(
                f"{int(tot_built_rk):,}", "Total Built", "All Floors", CYAN
            ), unsafe_allow_html=True)
            rsp[1].markdown(kpi_html(
                f"{int(tot_inst_rk):,}", "Total Installed", "All Floors", LBLUE
            ), unsafe_allow_html=True)
            rsp[2].markdown(kpi_html(
                f"{int(tot_oper_rk):,}", "Total Operational",
                f"{used_pct_rk:.1f}% Used", AMBER
            ), unsafe_allow_html=True)
            rsp[3].markdown(kpi_html(
                f"{int(tot_avail_rk):,}", "Total Available",
                f"{free_pct_rk:.1f}% Free", GREEN
            ), unsafe_allow_html=True)

            rsk_df = LOC_DF[LOC_DF["Built_Racks"] > 0].copy().sort_values("Built_Racks", ascending=True)
            if not rsk_df.empty:
                fig_rsk = go.Figure()
                fig_rsk.add_trace(go.Bar(name="Built",       x=rsk_df["Location"], y=rsk_df["Built_Racks"],       marker_color=LBLUE))
                fig_rsk.add_trace(go.Bar(name="Operational", x=rsk_df["Location"], y=rsk_df["Operational_Racks"], marker_color=RED))
                fig_rsk.add_trace(go.Bar(name="Available",   x=rsk_df["Location"], y=rsk_df["Available_Racks"],   marker_color=GREEN))
                fig_rsk.add_trace(go.Bar(name="Unconsumed",  x=rsk_df["Location"], y=rsk_df["Unconsumed_Racks"],  marker_color=AMBER))
                fig_rsk.add_trace(go.Bar(name="Actual",      x=rsk_df["Location"], y=rsk_df["Actual_Racks"],      marker_color=CYAN))
                fig_rsk.update_layout(
                    **_base_layout(),
                    barmode="group",
                    height=380,
                    yaxis_title="Racks",
                    xaxis_title="Location",
                    legend=dict(orientation="h", y=-0.25),
                )
                st.plotly_chart(fig_rsk, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ──────────────────────────────────────────────────────────────────
            # WHITE SPACE section
            # ──────────────────────────────────────────────────────────────────
            st.markdown('<div class="section-title">White Space</div>',
                        unsafe_allow_html=True)
            ws_df = LOC_DF[LOC_DF["Built_Sqft"] > 0].copy().sort_values("Built_Sqft", ascending=True)
            if not ws_df.empty:
                fig_ws = go.Figure()
                fig_ws.add_trace(go.Bar(name="Built Space",       x=ws_df["Location"], y=ws_df["Built_Sqft"],       marker_color=LBLUE))
                fig_ws.add_trace(go.Bar(name="Operational Space", x=ws_df["Location"], y=ws_df["Operational_Sqft"], marker_color=RED))
                fig_ws.add_trace(go.Bar(name="Available Space",   x=ws_df["Location"], y=ws_df["Available_Sqft"],   marker_color=GREEN))
                fig_ws.add_trace(go.Bar(name="Wasted Space",      x=ws_df["Location"], y=ws_df["Wasted_Sqft"],      marker_color=AMBER))
                fig_ws.update_layout(
                    **_base_layout(),
                    barmode="group",
                    height=380,
                    yaxis_title="Sq.ft",
                    xaxis_title="Location",
                    legend=dict(orientation="h", y=-0.25),
                )
                st.plotly_chart(fig_ws, use_container_width=True)
            else:
                st.caption("White Space (sqft) figures are not present in the uploaded facility sheets.")

            st.markdown("<br>", unsafe_allow_html=True)

            # ──────────────────────────────────────────────────────────────────
            # PER-LOCATION OPERATIONAL SUMMARY TABLE (like the cards in ref image)
            # ──────────────────────────────────────────────────────────────────
            st.markdown('<div class="section-title">Per-Location Summary</div>',
                        unsafe_allow_html=True)
            summary_df = LOC_DF[[
                "Location",
                "Avg_PUE", "Built_kW", "Operational_kW", "Available_kW",
                "Built_Racks", "Operational_Racks", "Available_Racks",
                "Built_Sqft", "Available_Sqft",
            ]].copy()
            summary_df.columns = [
                "Location", "Avg PUE",
                "Built kW", "Operational kW", "Available kW",
                "Built Racks", "Operational Racks", "Available Racks",
                "Built SqFt", "Available SqFt",
            ]
            summary_df = summary_df.round(2)
            st.dataframe(summary_df, use_container_width=True, height=380)

            st.markdown("<br>", unsafe_allow_html=True)

            # ──────────────────────────────────────────────────────────────────
            # BILLING MODEL BREAKDOWN (from customer-level CUST dataframe)
            # ──────────────────────────────────────────────────────────────────
            st.markdown('<div class="section-title">Billing Model</div>',
                        unsafe_allow_html=True)
            bm_cols = st.columns(4)

            def _cnt_contains(col, pat):
                if col and col in CUST.columns:
                    return int(
                        CUST[col].astype(str).str.upper().str.strip()
                        .str.contains(pat, na=False, regex=True).sum()
                    )
                return 0

            if caged_c:
                n_caged   = _cnt_contains(caged_c, r"^CAGED$|^CAGE$")
                n_uncaged = _cnt_contains(caged_c, r"UNCAGED")
                bm_cols[0].markdown(kpi_html(
                    f"{n_caged}", "Caged", f"Uncaged: {n_uncaged}", CYAN
                ), unsafe_allow_html=True)

            pw_sub_c   = find_col(CUST, r"power.*subscription.*model|billing.*model.*power.*subscription")
            pw_use_m_c = find_col(CUST, r"power.*usage.*model|billing.*model.*power.*usage")

            if pw_sub_c:
                n_rated = _cnt_contains(pw_sub_c, r"RATED")
                n_sub   = _cnt_contains(pw_sub_c, r"SUBSCRIBED")
                bm_cols[1].markdown(kpi_html(
                    f"{n_rated}", "Rated", f"Subscribed: {n_sub}", LBLUE
                ), unsafe_allow_html=True)

            if pw_use_m_c:
                n_bund = _cnt_contains(pw_use_m_c, r"BUNDLED")
                n_met  = _cnt_contains(pw_use_m_c, r"METERED")
                bm_cols[2].markdown(kpi_html(
                    f"{n_bund}", "Bundled", f"Metered: {n_met}", GREEN
                ), unsafe_allow_html=True)

            if own_c:
                n_sify = _cnt_contains(own_c, r"SIFY")
                n_cust = _cnt_contains(own_c, r"CUSTOMER")
                if n_sify or n_cust:
                    bm_cols[3].markdown(kpi_html(
                        f"Sify: {n_sify}", "Ownership",
                        f"Customer: {n_cust}", AMBER
                    ), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ──────────────────────────────────────────────────────────────────
            # UTILISATION GAUGES
            # ──────────────────────────────────────────────────────────────────
            if tot_built_kw > 0 or tot_built_rk > 0:
                st.markdown('<div class="section-title">Utilisation Gauges</div>',
                            unsafe_allow_html=True)
                g1, g2 = st.columns(2)

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

                pwr_util = (tot_oper_kw / tot_built_kw * 100) if tot_built_kw > 0 else 0
                rk_util  = (tot_oper_rk / tot_built_rk * 100) if tot_built_rk > 0 else 0
                g1.plotly_chart(_gauge(pwr_util, "Power Capacity Utilisation (%)", LBLUE),
                                use_container_width=True)
                g2.plotly_chart(_gauge(rk_util, "Rack Space Utilisation (%)", GREEN),
                                use_container_width=True)


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
        st.warning("No data loaded. Please check that Excel files are present.")
    else:
        # ── Row 1: Location + Sheet filters ──────────────────────────────────
        op1, op2 = st.columns(2)
        with op1:
            op_loc = st.selectbox("📍 Filter by Location",
                                  ["All"] + sorted(fdata.keys()), key="op_loc")
        with op2:
            op_sh_opts = (sorted(fdata.get(op_loc, {}).keys())
                          if op_loc != "All"
                          else sorted({sn for s in fdata.values() for sn in s}))
            op_sh = st.selectbox("📋 Filter by Sheet",
                                 ["All"] + op_sh_opts, key="op_sh")

        # Apply location / sheet filter
        op_df = CUST.copy()
        if op_loc != "All" and "_Location" in op_df.columns:
            op_df = op_df[op_df["_Location"].str.contains(op_loc, case=False, na=False)]
        if op_sh != "All" and "_Sheet" in op_df.columns:
            op_df = op_df[op_df["_Sheet"] == op_sh]

        st.caption(f"🔢 **{len(op_df):,}** rows available "
                   f"({'all locations' if op_loc == 'All' else op_loc}"
                   f"{', ' + op_sh if op_sh != 'All' else ''})")

        # ── ALL selectable columns (not just dtype-numeric) ──────────────────
        # _robust_to_numeric will parse them correctly at run time
        all_cols = [c for c in op_df.columns if not c.startswith("_")]
        nc_op    = num_cols(op_df)   # already-numeric (dtype) — highlighted
        tc_op    = txt_cols(op_df)   # text / categoricals

        # Group-by candidates: categorical + Location/Sheet meta
        grp_candidates = (
            [c for c in ["_Location", "_Sheet"] if c in op_df.columns]
            + [c for c in tc_op if c not in ("_Location", "_Sheet")]
        )

        # ── Row 2: Column + Operation + Group-by ─────────────────────────────
        op3, op4, op5 = st.columns([2, 2, 2])
        with op3:
            # Show ALL columns; mark already-numeric with ★
            col_display = {c: (f"★ {c}" if c in nc_op else c) for c in all_cols}
            op_col_disp = st.selectbox(
                "📐 Column  (★ = already numeric)",
                options=list(col_display.values()),
                key="op_col_disp",
            )
            # Map back to real column name
            op_col = next((k for k, v in col_display.items() if v == op_col_disp), op_col_disp)

        with op4:
            op_op = st.selectbox("🔧 Operation", OPERATIONS, key="op_op")

        with op5:
            op_grp = st.selectbox(
                "🗂 Group By  (optional)",
                ["None"] + [c for c in grp_candidates if c != op_col],
                key="op_grp",
            )

        op6, op7 = st.columns([1, 3])
        with op6:
            op_n = st.number_input("N  (for Top / Bottom N)",
                                   min_value=1, max_value=500, value=10, step=1, key="op_n")

        # ── Quick column preview (before running) ─────────────────────────────
        if op_col and op_col in op_df.columns:
            preview_series = _robust_to_numeric(op_df[op_col])
            valid_preview  = preview_series.dropna()
            total_preview  = len(preview_series)
            pct_v          = valid_preview.shape[0] / max(total_preview, 1) * 100
            col_p1, col_p2, col_p3, col_p4 = st.columns(4)
            col_p1.metric("Total Rows",   f"{total_preview:,}")
            col_p2.metric("Numeric Rows", f"{len(valid_preview):,}")
            col_p3.metric("Valid %",      f"{pct_v:.1f}%")
            if not valid_preview.empty:
                col_p4.metric("Quick Sum", _fmt_decimal(valid_preview.sum()))

        # ── Run button ────────────────────────────────────────────────────────
        if st.button("▶  Run Operation", key="op_run", use_container_width=True):
            if not op_col or op_col not in op_df.columns:
                st.error("Please select a valid column.")
            else:
                grp = op_grp if op_grp != "None" else None
                out = run_op(op_df, op_col, op_op, grp, int(op_n))

                # run_op always returns 5-tuple; first element may be None
                if out[0] is None:
                    # Error string in out[1]
                    st.error(out[1])
                else:
                    result, desc, valid_pct, valid_count, total_count = out

                    st.markdown(
                        f'<div class="section-title">📊 {desc}</div>',
                        unsafe_allow_html=True,
                    )

                    # ── Scalar result ─────────────────────────────────────────
                    if isinstance(result, (int, float)):
                        unit  = _detect_unit(op_col)
                        raw   = _fmt_decimal(result, unit)
                        disp  = f"₹ {_fmt_decimal(result)} {unit}".strip() if unit == "₹" else raw

                        r1, r2, r3 = st.columns(3)
                        with r1:
                            st.markdown(f"""
                            <div style="background:{DARK2};border:1px solid {BORD};
                                 border-radius:14px;padding:24px 28px;text-align:center">
                              <div style="font-size:.75rem;color:{MUTED};font-weight:700;
                                   text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">
                                {op_op}</div>
                              <div style="font-size:2.2rem;font-weight:900;color:{CYAN};
                                   line-height:1.1">{disp}</div>
                              <div style="font-size:.8rem;color:{MUTED};margin-top:8px">
                                {op_col}</div>
                            </div>""", unsafe_allow_html=True)
                        with r2:
                            st.metric("Rows used", f"{valid_count:,} / {total_count:,}")
                            st.metric("Valid %", valid_pct)
                        with r3:
                            if unit:
                                st.info(f"Unit detected: **{unit}**")
                            st.caption(f"Source: {op_loc} › {op_sh}")

                    # ── Table result (grouped / Top-N / etc.) ─────────────────
                    elif isinstance(result, pd.DataFrame):
                        st.dataframe(result, use_container_width=True)

                        # Stat pills
                        sp1, sp2, sp3 = st.columns(3)
                        sp1.metric("Rows returned", len(result))
                        sp2.metric("Rows used (valid)", f"{valid_count:,} / {total_count:,}")
                        sp3.metric("Valid %", valid_pct)

                        # Bar chart when grouped and column exists in result
                        if grp and op_col in result.columns and grp in result.columns:
                            plot_df = result.head(30).copy()
                            fig_op = px.bar(
                                plot_df,
                                x=grp, y=op_col,
                                color=op_col,
                                color_continuous_scale="Blues",
                                title=desc,
                                labels={op_col: op_op, grp: grp},
                                text=op_col,
                            )
                            fig_op.update_traces(texttemplate="%{text:,.2f}", textposition="outside")
                            fig_op.update_layout(**_base_layout(), height=420,
                                                 xaxis_tickangle=-35)
                            st.plotly_chart(fig_op, use_container_width=True)

                        # Download button
                        csv_bytes = result.to_csv(index=True).encode()
                        st.download_button(
                            "⬇️  Download result as CSV",
                            data=csv_bytes,
                            file_name=f"ops_result_{op_col[:20]}_{op_op}.csv",
                            mime="text/csv",
                            key="op_download",
                        )
                    else:
                        st.info(str(result))


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


# ═══════════════════════════════════════════════════════════════════════════════
# SMART QUERY — ENHANCED DATA INGESTION + ACCURATE SCHEMA MAPPING  (additive)
# These functions are used ONLY by the Smart Query tab (T[4]).
# All other tabs and all existing functions/globals remain completely unchanged.
# ═══════════════════════════════════════════════════════════════════════════════

# ── Comprehensive field-hint → column pattern registry ─────────────────────────
# Covers every real column name variant observed across all 10 DC Excel files.
_SQ_FIELD_PATTERNS: dict = {
    # Power / Capacity ─────────────────────────────────────────────────────────
    "total capacity purchased": [
        r"Power Capacity\s*\|\s*Total Capacity Purchased",
        r"^Total Capacity Purchased",
        r"Total Capacity Purchased \(KW\)",
        r"capacity.*purchased",
        r"subscribed.*kw|kw.*subscribed",
    ],
    "power in use": [
        r"Power Capacity\s*\|\s*Capacity in Use",
        r"Power Capacity\s*\|\s*Usage in KW",
        r"Capacity in Use \(KW\)",
        r"^Capacity in Use$",
        r"capacity\s+in\s+use",
        r"usage\s+in\s+kw",
    ],
    "power allocated": [
        r'Power Capacity\s*\|\s*"?Allocated"?\s*Capacity',
        r"allocated.*capacity.*kw",
        r"allocated.*kw",
    ],
    "power usage kw": [
        r"Power Capacity\s*\|\s*Usage in KW",
        r"Power Usage.*Raw Power",
        r"Power Usage \(All in KW\)",
        r"usage\s+in\s+kw",
        r"actual.*usage.*kw",
        r"Actual Load KVA",
    ],
    "subscribed capacity kw": [
        r"Power Capacity\s*\|\s*Subscribed Capacity to be given in KW",
        r"subscribed.*capacity.*to.*be.*given",
        r"capacity.*to.*be.*given.*kw",
    ],
    "capacity to be given": [
        r"Power Capacity\s*\|\s*Capacity to be given",
        r"capacity.*to.*be.*given",
    ],
    "reserved capacity": [
        r"Power Capacity\s*\|\s*Reserved Capacity",
        r"reserved.*capacity",
        r"Seating Space\s*\|\s*Reserved Capacity",
    ],
    "additional capacity charges": [
        r"Power Capacity\s*\|\s*Additional Capacity Charges",
        r"Power Capacity\s*\|\s*Billable Additional Capacity",
        r"additional.*capacity.*charges",
        r"billable.*additional.*capacity",
    ],
    # Space / Racks ─────────────────────────────────────────────────────────────
    "total space": [
        r"^Space \| Subscription$",
        r"Space\s*\|\s*Subscription",
        r"Sitting Space \(Subscription\)",
        r"space.*subscription",
        r"Subscription\(No\. of Racks\)",
    ],
    "space in use": [
        r"^Space \| In Use$",
        r"Space\s*\|\s*In Use",
        r"In Use\(No\. of Racks\)",
        r"space.*in.*use",
    ],
    "space billed": [
        r"^Space \| Billed$",
        r"Space\s*\|\s*Billed",
        r"space.*billed",
        r"Seating Space\s*\|\s*Billed",
    ],
    "space yet to be given": [
        r"Space\s*\|\s*Yet to be given",
        r"yet.*to.*be.*given",
        r"Seating Space\s*\|\s*Yet to be given",
    ],
    "seating subscription": [
        r"^Seating Space \| Subscription$",
        r"Seating Space\s*\|\s*Subscription",
        r"seating.*space.*subscription",
        r"sitting.*space.*subscription",
    ],
    "seating in use": [
        r"^Seating Space \| In Use$",
        r"Seating Space\s*\|\s*In Use",
        r"seating.*space.*in.*use",
    ],
    # Revenue ───────────────────────────────────────────────────────────────────
    "total revenue": [
        r"Revenue \(Monthly\s*\)\s*\|\s*Total Revenue",
        r"Revenue.*\|\s*Total Revenue",
        r"^Total Revenue$",
        r"total.*revenue",
        r"total.*mrc",
        r"Total Rev \(Cap \+ Power\)",
        r"Total\s*\|\s*\d",   # summary total rows
    ],
    "space revenue": [
        r"Revenue \(Monthly\s*\)\s*\|\s*Space revenue",
        r"Revenue.*\|\s*Space revenue",
        r"space.*revenue.*including",
        r"space.*revenue",
    ],
    "power revenue": [
        r"Revenue \(Monthly\s*\)\s*\|\s*Power Usage revenue",
        r"Revenue.*\|\s*Power Usage revenue",
        r"Contract Information\s*\|\s*Power Revenue",
        r"power.*usage.*revenue",
        r"power.*revenue",
    ],
    "additional capacity revenue": [
        r"Revenue \(Monthly\s*\)\s*\|\s*Additional Capacity Revenue",
        r"Revenue.*\|\s*Additional Capacity Revenue",
        r"additional.*capacity.*revenue",
        r"Contract Information\s*\|\s*Capacity Revenue",
    ],
    "seating revenue": [
        r"Revenue \(Monthly\s*\)\s*\|\s*Seating Space",
        r"Revenue.*\|\s*Seating Space",
        r"seating.*space.*revenue",
        r"seating.*revenue",
    ],
    "net revenue": [
        r"Contract Information\s*\|\s*Net Rev Total",
        r"Contract Information\s*\|\s*Total Rev \(Cap \+ Power\)",
        r"net.*rev.*total",
        r"net.*revenue",
        r"Total Rev \(Cap \+ Power\)",
    ],
    # Rate ──────────────────────────────────────────────────────────────────────
    "per unit rate": [
        r"Power Usage\s*\|\s*Unit Rate \(per KW-HR\)",
        r"Power Usage\s*\|\s*Unit rate.*KW",
        r"Space\s*\|\s*Per Unit rate",
        r"Seating Space\s*\|\s*Per Unit rate",
        r"per.*unit.*rate",
        r"unit.*rate.*kw",
    ],
    # Contract ──────────────────────────────────────────────────────────────────
    "contract start": [
        r"Contract Information\s*\|\s*Contract Start",
        r"contract.*start",
        r"start.*date",
    ],
    "contract expiry": [
        r"Contract Information\s*\|\s*Current Ex[ip]iry Date",
        r"expiry.*date",
        r"current.*expiry",
    ],
    "contract term": [
        r"Contract Information\s*\|\s*Term of Contract",
        r"term.*of.*contract",
        r"contract.*term",
        r"no.*of.*years",
    ],
}

# Extended phrase → field key aliases (checked before _HINT_SEMANTIC)
_SQ_HINT_ALIASES: list = [
    # Power
    ("total capacity purchased",     "total capacity purchased"),
    ("total power purchased",        "total capacity purchased"),
    ("total kw purchased",           "total capacity purchased"),
    ("total kva purchased",          "total capacity purchased"),
    ("power purchased",              "total capacity purchased"),
    ("capacity purchased",           "total capacity purchased"),
    ("subscribed capacity",          "total capacity purchased"),
    ("total capacity",               "total capacity purchased"),
    ("power capacity",               "total capacity purchased"),
    ("sum of power",                 "total capacity purchased"),
    ("total power",                  "total capacity purchased"),
    ("power kw",                     "total capacity purchased"),
    ("total kw",                     "total capacity purchased"),
    ("power in use",                 "power in use"),
    ("capacity in use",              "power in use"),
    ("power used",                   "power in use"),
    ("power usage",                  "power in use"),
    ("usage in kw",                  "power in use"),
    ("kw in use",                    "power in use"),
    ("power usage kw",               "power usage kw"),
    ("actual usage",                 "power usage kw"),
    ("raw power",                    "power usage kw"),
    ("power allocated",              "power allocated"),
    ("allocated capacity",           "power allocated"),
    ("allocated kw",                 "power allocated"),
    ("capacity to be given",         "capacity to be given"),
    ("subscribed capacity kw",       "subscribed capacity kw"),
    ("reserved capacity",            "reserved capacity"),
    ("additional capacity charges",  "additional capacity charges"),
    # Space
    ("total space",                  "total space"),
    ("space subscription",           "total space"),
    ("space subscribed",             "total space"),
    ("space purchased",              "total space"),
    ("sitting space subscription",   "total space"),
    ("total racks",                  "total space"),
    ("number of racks",              "total space"),
    ("space in use",                 "space in use"),
    ("space used",                   "space in use"),
    ("racks in use",                 "space in use"),
    ("space billed",                 "space billed"),
    ("space yet to be given",        "space yet to be given"),
    ("seating subscription",         "seating subscription"),
    ("seating space subscription",   "seating subscription"),
    ("seating in use",               "seating in use"),
    ("seating space in use",         "seating in use"),
    # Revenue
    ("total revenue",                "total revenue"),
    ("total mrc",                    "total revenue"),
    ("total monthly revenue",        "total revenue"),
    ("sum of revenue",               "total revenue"),
    ("sum revenue",                  "total revenue"),
    ("revenue total",                "total revenue"),
    ("mrc",                          "total revenue"),
    ("space revenue",                "space revenue"),
    ("revenue from space",           "space revenue"),
    ("space including capacity",     "space revenue"),
    ("power revenue",                "power revenue"),
    ("power usage revenue",          "power revenue"),
    ("revenue from power",           "power revenue"),
    ("additional capacity revenue",  "additional capacity revenue"),
    ("seating revenue",              "seating revenue"),
    ("net revenue",                  "net revenue"),
    ("net rev",                      "net revenue"),
    ("total rev",                    "net revenue"),
    # Rate
    ("per unit rate",                "per unit rate"),
    ("unit rate",                    "per unit rate"),
    ("tariff",                       "per unit rate"),
    ("rate per kw",                  "per unit rate"),
]


def _sq_resolve_field(df: "pd.DataFrame", field_hint: str) -> "tuple[str|None, str]":
    """
    Enhanced column resolver for Smart Query tab only.
    Priority order:
      1. Extended _SQ_FIELD_PATTERNS via _SQ_HINT_ALIASES (covers all DC Excel column variants)
      2. Original _HINT_SEMANTIC / _SEMANTIC_COLS (legacy fallback)
      3. Fuzzy multi-word match across ALL columns (not just dtype-numeric)
      4. First numeric column
    Returns (column_name, reason_string).
    """
    hint_lower = (field_hint or "").lower().strip()
    nc = num_cols(df)

    # ── 1. Extended pattern registry ──────────────────────────────────────────
    resolved_key = None
    for alias, key in _SQ_HINT_ALIASES:
        if alias in hint_lower:
            resolved_key = key
            break
    if resolved_key is None:
        for key in _SQ_FIELD_PATTERNS:
            if key in hint_lower or hint_lower in key:
                resolved_key = key
                break

    if resolved_key and resolved_key in _SQ_FIELD_PATTERNS:
        for pat in _SQ_FIELD_PATTERNS[resolved_key]:
            for c in df.columns:
                if re.search(pat, c, re.I):
                    return c, f"extended schema: '{resolved_key}' via «{pat}»"

    # ── 2. Original _HINT_SEMANTIC / _SEMANTIC_COLS ───────────────────────────
    for kw, sem_key in _HINT_SEMANTIC:
        if kw in hint_lower:
            pattern, _ = _SEMANTIC_COLS[sem_key]
            for c in df.columns:
                if re.search(pattern, c, re.I):
                    return c, f"original semantic: '{kw}' → '{sem_key}'"

    # ── 3. Fuzzy word match across all columns ────────────────────────────────
    hint_words = [w for w in re.split(r"\W+", hint_lower) if len(w) > 2]
    best_col, best_score = None, 0
    for c in df.columns:
        c_lower = c.lower()
        score = sum(1 for w in hint_words if w in c_lower)
        if score > best_score:
            best_score, best_col = score, c
    if best_col and best_score > 0:
        return best_col, f"fuzzy match ({hint_words}, score={best_score})"

    # ── 4. Fallback ───────────────────────────────────────────────────────────
    if nc:
        return nc[0], "fallback: first numeric column"
    return None, "no column found"


# Known non-customer row patterns for ingestion filter
_SQ_JUNK_NAMES: frozenset = frozenset({
    "customer name", "sr. no", "sno", "s.no", "no.", "sl. no",
    "total", "sub total", "subtotal", "grand total", "summary",
    "description", "floor", "module", "floor / module",
    "power summary", "nan", "none", "", "remark", "remarks",
    "total bangalore", "total kolkata", "total noida", "total chennai",
    "total vashi", "total airoli", "total rabale", "total mumbai",
    "uom", "uom (kva/kw)", "value",
})


def _sq_preprocess_pool(df: "pd.DataFrame") -> "pd.DataFrame":
    """
    Data ingestion enrichment for Smart Query pool (additive — does NOT
    alter any global state or other tabs).

    Step 1 — Remove non-customer rows:
        Rows whose customer-name cell is blank, a section header, a serial
        number, or a known aggregate label are dropped.

    Step 2 — Promote object columns to float64:
        Columns whose names suggest numeric content (power, revenue, capacity,
        space, rack, rate …) are run through _robust_to_numeric().  A column
        is only replaced if ≥ 20 % of its non-null values parse successfully —
        this keeps genuinely categorical columns untouched while converting
        columns that happen to contain "₹ 1,234.56" or "1,23,456.78" strings.
    """
    if df.empty:
        return df

    result = df.copy()

    # ── Step 1: Remove non-customer rows ─────────────────────────────────────
    cust_col = find_col(result, r"customer.*name|client.*name|DEMARC.*Customer Name")
    if cust_col:
        def _is_real_customer(v):
            s = str(v).strip()
            if not s or s.lower() in _SQ_JUNK_NAMES:
                return False
            if re.fullmatch(r"\d+", s):      # pure serial number
                return False
            if re.fullmatch(r"[-–—]+", s):   # dash placeholder
                return False
            if len(s) < 2:
                return False
            return True

        valid_mask = result[cust_col].apply(_is_real_customer)
        if valid_mask.sum() > 0:
            result = result[valid_mask].reset_index(drop=True)

    # ── Step 2: Promote suspected numeric columns to float64 ─────────────────
    _NUMERIC_KW = {
        "subscription", "in use", "billed", "reserved", "capacity",
        "purchased", "allocated", "revenue", "mrc", "rate", "charge",
        "kw", "kva", "rack", "space", "seat", "sitting", "kwhr",
        "quantity", "qty", "amount", "total", "yet to be given",
        "subscribed", "usage", "consumption", "additional",
    }
    metadata_cols = {c for c in result.columns if c.startswith("_")}

    for col in result.columns:
        if col in metadata_cols:
            continue
        if pd.api.types.is_numeric_dtype(result[col]):
            continue                           # already float/int — skip
        col_lower = col.lower()
        if not any(kw in col_lower for kw in _NUMERIC_KW):
            continue                           # not a candidate — skip

        converted = _robust_to_numeric(result[col])
        non_null  = result[col].notna().sum()
        parsed_ok = converted.notna().sum()
        if non_null > 0 and parsed_ok / max(non_null, 1) >= 0.20:
            result[col] = converted            # promote to numeric

    return result


def _sq_execute_with_schema(ops_raw: list, pool: "pd.DataFrame") -> list:
    """
    Wrapper around execute_ai_operations that temporarily substitutes the
    enhanced column resolver _sq_resolve_field for _resolve_col_by_semantic
    so that all 10 DC Excel column name variants are correctly matched.
    Restores the original resolver after execution.
    """
    import sys as _sys
    _mod = _sys.modules[__name__]
    _orig = getattr(_mod, "_resolve_col_by_semantic", None)
    try:
        # Patch global resolver with the enhanced version
        _mod._resolve_col_by_semantic = _sq_resolve_field
        results = execute_ai_operations(ops_raw, pool)
    finally:
        # Always restore original, even on exception
        if _orig is not None:
            _mod._resolve_col_by_semantic = _orig
    return results


# ═══════════════════════════════════════════════════════════════════════════════
# SMART QUERY: CUSTOMER-WISE LOOKUP HELPERS  (additive — T[4] only)
# No existing functions, globals, or other tabs are changed.
# ═══════════════════════════════════════════════════════════════════════════════

# Profile metadata columns shown in the customer card
_SQ_CUST_PROFILE_PATS: list = [
    (r"\bfloor\b|\bmodule\b",                                         "Floor / Module"),
    (r"\bSH\b|sub.*hall",                                             "Sub-Hall"),
    (r"caged.*uncaged|caged",                                         "Caged / Uncaged"),
    (r"ownership.*sify.*cust|ownership",                              "Ownership"),
    (r"subscription.*mode|space.*subscription.*mode",                 "Subscription Mode"),
    (r"power.*subscription.*model|billing.*model.*power.*subscr",     "Power Subscription Model"),
    (r"power.*usage.*model|billing.*model.*power.*usage",             "Power Usage Model"),
    (r"uom.*kva|uom",                                                 "UoM (KW/KVA)"),
    (r"billing.*frequency|frequency",                                 "Billing Frequency"),
]

# Data column patterns to include in the customer metrics table (in display order)
_SQ_CUST_DATA_PATS: list = [
    (r"Power Capacity.*Total Capacity Purchased|Total Capacity Purchased", "Power Capacity Purchased"),
    (r"Power Capacity.*Capacity in Use|^Capacity in Use$",                 "Power Capacity in Use"),
    (r"Power Capacity.*Usage in KW|Usage in KW",                           "Power Usage (KW)"),
    (r'Power Capacity.*"?Allocated"?.*Capacity',                           "Allocated Capacity (KW)"),
    (r"Power Capacity.*Subscribed Capacity.*KW",                           "Subscribed Capacity to be Given (KW)"),
    (r"Power Capacity.*Capacity to be given",                              "Capacity to be Given"),
    (r"Power Capacity.*Reserved Capacity",                                 "Reserved Capacity"),
    (r"Power Capacity.*Additional Capacity Charges",                       "Additional Capacity Charges (₹)"),
    (r"^Space \| Subscription$|Space.*Subscription",                       "Space Subscription"),
    (r"^Space \| In Use$|Space.*In Use",                                   "Space In Use"),
    (r"^Space \| Billed$|Space.*Billed",                                   "Space Billed"),
    (r"^Space \| Yet to be given",                                         "Space Yet to be Given"),
    (r"Seating Space.*Subscription",                                        "Seating Subscription"),
    (r"Seating Space.*In Use",                                              "Seating In Use"),
    (r"Revenue.*Space revenue|Space revenue",                               "Revenue — Space (₹)"),
    (r"Revenue.*Power Usage revenue",                                       "Revenue — Power (₹)"),
    (r"Revenue.*Additional Capacity Revenue",                               "Revenue — Add. Capacity (₹)"),
    (r"Revenue.*Seating Space",                                             "Revenue — Seating (₹)"),
    (r"Revenue.*Total Revenue|Total Revenue",                               "Total Revenue (₹)"),
    (r"Contract Information.*Net Rev Total",                                "Net Revenue Total (₹)"),
    (r"Contract Information.*Total Rev.*Cap.*Power",                        "Total Rev (Cap + Power) (₹)"),
    (r"Power Usage.*Unit Rate|Per Unit rate",                               "Per Unit Rate (₹/kWh)"),
    (r"Contract Information.*Term of Contract",                             "Contract Term (Years)"),
    (r"Contract Information.*Contract Start",                               "Contract Start Date"),
    (r"Contract Information.*Current Ex",                                   "Contract Expiry Date"),
    (r"Contract Information.*Sales Order",                                  "Sales Order Ref"),
]


def _sq_detect_customer_query(query: str) -> "tuple[str, str]":
    """
    Extract customer name from a natural-language query.
    Returns (customer_name, remaining_field_hint).
    Returns ("", query) when no customer name pattern is detected.

    Patterns recognised:
      "power capacity purchased for Oracle"         → ("Oracle", "power capacity purchased")
      "Oracle's power capacity"                     → ("Oracle", "power capacity")
      "show data for CISCO SYSTEMS"                 → ("CISCO SYSTEMS", "show data")
      "total revenue of Wipro"                      → ("Wipro", "total revenue")
      "customer Mahindra capacity in use"           → ("Mahindra", "capacity in use")
    """
    q = query.strip()

    # Pattern 1: "... for <CustomerName>" — most common natural phrasing
    m = re.search(
        r"\bfor\s+([A-Za-z][\w\s&().,'\"/-]{2,60}?)(?:\s+(?:at|in|across|from|by|across)\b|\s*$)",
        q, re.I
    )
    if m:
        cust = m.group(1).strip().rstrip(",.?")
        field = re.sub(r"\bfor\s+" + re.escape(cust), "", q, flags=re.I).strip()
        return cust, field

    # Pattern 2: "<CustomerName>'s <field>"
    m = re.search(r"^([A-Za-z][\w\s&().,'/-]{1,60}?)\s*[''`]s\s+(.+)$", q, re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()

    # Pattern 3: "customer <CustomerName>"
    m = re.search(r"\bcustomer\s+([A-Z][\w\s&().,'/-]{1,60}?)(?:\s+and\b|\s*$)", q, re.I)
    if m:
        cust = m.group(1).strip().rstrip(",.?")
        field = re.sub(r"\bcustomer\s+" + re.escape(cust), "", q, flags=re.I).strip()
        return cust, field

    # Pattern 4: "... of <CapitalisedCustomer>" (only if capitalised, avoids "sum of power")
    m = re.search(
        r"\bof\s+([A-Z][A-Za-z][\w\s&().,'/-]{1,60}?)(?:\s+(?:at|in|across|from|and|for)\b|\s*$)",
        q
    )
    if m:
        cust = m.group(1).strip().rstrip(",.?")
        if len(cust) > 3:
            field = re.sub(r"\bof\s+" + re.escape(cust), "", q, flags=re.I).strip()
            return cust, field

    return "", q


def _sq_find_customers(search: str, df: "pd.DataFrame") -> "pd.DataFrame":
    """
    Return all rows in df whose customer-name column contains `search`
    (case-insensitive substring).  Tries DEMARC | Customer Name first,
    then any column matching customer.*name.
    """
    cust_col = find_col(df,
        r"DEMARC.*Customer Name",
        r"customer.*name",
        r"client.*name",
    )
    if not cust_col or not search.strip():
        return pd.DataFrame()
    mask = df[cust_col].astype(str).str.lower().str.contains(
        re.escape(search.strip().lower()), na=False
    )
    return df[mask].copy().reset_index(drop=True)


def _sq_build_customer_profile(rows: "pd.DataFrame",
                                customer_search: str,
                                field_hint: str) -> dict:
    """
    Build a richly structured result dict for a customer-wise lookup.
    Keys:
      type, label, customer, row_count,
      profile_df   — metadata profile table
      focus_col    — specifically requested column (if field_hint given)
      focus_val    — sum of that column for this customer
      focus_unit   — unit string
      metrics_df   — all numeric data columns with values
      raw_df       — full detail table for download
    """
    cust_col = find_col(rows,
        r"DEMARC.*Customer Name", r"customer.*name", r"client.*name")

    # ── Canonical customer display name ──────────────────────────────────────
    if cust_col:
        names = rows[cust_col].dropna().astype(str).str.strip().unique()
        names = [n for n in names if n and n.lower() not in ("none","nan","")]
        display_name = names[0] if names else customer_search
    else:
        display_name = customer_search

    # ── Profile card ─────────────────────────────────────────────────────────
    profile_rows = []
    if "_Location" in rows.columns:
        locs = rows["_Location"].unique().tolist()
        profile_rows.append({"Field": "Location(s)", "Value": ", ".join(locs)})
    if "_Sheet" in rows.columns:
        sheets = rows["_Sheet"].unique().tolist()
        profile_rows.append({"Field": "Sheet(s)", "Value": ", ".join(sheets)})
    profile_rows.append({"Field": "Matched Rows", "Value": str(len(rows))})

    for pat, label in _SQ_CUST_PROFILE_PATS:
        c = find_col(rows, pat)
        if c and c in rows.columns:
            vals = rows[c].dropna().astype(str).str.strip().unique()
            vals = [v for v in vals if v and v.lower() not in ("none","nan","")]
            if vals:
                profile_rows.append({"Field": label, "Value": ", ".join(vals[:5])})

    profile_df = pd.DataFrame(profile_rows) if profile_rows else pd.DataFrame()

    # ── Metrics table (all numeric data columns) ──────────────────────────────
    metric_rows = []
    seen_cols: set = set()
    for pat, metric_label in _SQ_CUST_DATA_PATS:
        c = find_col(rows, pat)
        if not c or c in seen_cols or c.startswith("_"):
            continue
        seen_cols.add(c)
        series = _robust_to_numeric(rows[c])
        valid = series.dropna()
        if valid.empty:
            continue
        total_val = valid.sum()
        unit = _detect_unit(c)
        display_val = (
            f"₹ {total_val:,.2f}" if unit == "₹"
            else f"{_fmt_decimal(total_val)} {unit}".strip()
        )
        metric_rows.append({
            "Metric":  metric_label,
            "Value":   display_val,
            "_raw_val": total_val,
            "_col":    c,
            "_unit":   unit,
            "_n_rows": f"{len(valid)}/{len(rows)}",
        })

    metrics_df = (
        pd.DataFrame(metric_rows).drop(columns=["_raw_val","_col","_unit","_n_rows"])
        if metric_rows else pd.DataFrame()
    )
    # Keep the raw detail separately for the focus value lookup
    _metric_detail = metric_rows

    # ── Focus column (specifically requested metric) ──────────────────────────
    focus_col, focus_reason, focus_val, focus_unit = None, "", None, ""
    if field_hint.strip():
        focus_col, focus_reason = _sq_resolve_field(rows, field_hint)
        if focus_col and focus_col in rows.columns:
            _s = _robust_to_numeric(rows[focus_col]).dropna()
            focus_val  = float(_s.sum()) if not _s.empty else None
            focus_unit = _detect_unit(focus_col)

    # ── Full raw detail table ─────────────────────────────────────────────────
    meta_c = [c for c in ["_Location", "_Sheet"] if c in rows.columns]
    data_c = [c for c in rows.columns if not c.startswith("_")]
    raw_df = rows[meta_c + data_c].reset_index(drop=True)
    raw_df.index += 1

    return {
        "type":         "customer_lookup",
        "label":        f"Customer: {display_name}",
        "customer":     display_name,
        "row_count":    len(rows),
        "profile_df":   profile_df,
        "metrics_df":   metrics_df,
        "focus_col":    focus_col,
        "focus_val":    focus_val,
        "focus_unit":   focus_unit,
        "focus_reason": focus_reason,
        "raw_df":       raw_df,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# SHAKTISHIV — ENHANCED CUSTOMER LOOKUP HELPERS (additive only, T[4])
# Fixes multi-column customer name detection (DEMARC | Customer Name vs
# Customer Name) so Airoli, Bangalore, Kolkata, Noida etc. are all found.
# ═══════════════════════════════════════════════════════════════════════════════

# All column name patterns that can hold a customer/client name across all DCs
_SK_CUST_COL_PATS = [
    r"DEMARC.*Customer Name",
    r"^Customer Name$",
    r"^Customer Name\b",
    r"customer.*name",
    r"client.*name",
]


def _sk_all_customer_cols(df: "pd.DataFrame") -> list:
    """Return every column in df whose name matches any customer-name pattern."""
    hits = []
    for pat in _SK_CUST_COL_PATS:
        for col in df.columns:
            if re.search(pat, col, re.I) and col not in hits:
                hits.append(col)
    return hits


def _sk_find_customers_all(search: str, df: "pd.DataFrame") -> "pd.DataFrame":
    """
    Case-insensitive substring search across ALL customer-name columns in df.
    Combines results so rows from Airoli (DEMARC | Customer Name), Bangalore,
    Kolkata, Noida (Customer Name) are all returned together.
    Deduplicates by original index.
    """
    if not search.strip() or df.empty:
        return pd.DataFrame()
    term  = search.strip().lower()
    cols  = _sk_all_customer_cols(df)
    if not cols:
        return pd.DataFrame()
    combined_idx: set = set()
    for col in cols:
        mask = df[col].astype(str).str.lower().str.contains(
            re.escape(term), na=False
        )
        combined_idx.update(df.index[mask].tolist())
    if not combined_idx:
        return pd.DataFrame()
    return df.loc[sorted(combined_idx)].copy().reset_index(drop=True)


def _sk_canonical_name(rows: "pd.DataFrame", fallback: str) -> str:
    """Return the most common canonical customer name from the matched rows."""
    cols = _sk_all_customer_cols(rows)
    for col in cols:
        if col in rows.columns:
            vals = rows[col].dropna().astype(str).str.strip().unique()
            vals = [v for v in vals if v and v.lower() not in ("none","nan","")]
            if vals:
                return vals[0]
    return fallback


def _sk_build_per_loc_metrics(gdf: "pd.DataFrame") -> "pd.DataFrame":
    """Build a metrics table for a single location/sheet group of rows."""
    rows_ = []
    seen_ : set = set()
    for pat, label in _SQ_CUST_DATA_PATS:
        col = find_col(gdf, pat)
        if not col or col in seen_ or col.startswith("_"):
            continue
        seen_.add(col)
        s = _robust_to_numeric(gdf[col]).dropna()
        if s.empty:
            continue
        v    = float(s.sum())
        unit = _detect_unit(col)
        disp = (f"₹ {v:,.2f}" if unit == "₹"
                else f"{_fmt_decimal(v)} {unit}".strip())
        rows_.append({"Metric": label, "Value": disp})
    return pd.DataFrame(rows_) if rows_ else pd.DataFrame()


def _sk_build_per_loc_profile(gdf: "pd.DataFrame") -> "pd.DataFrame":
    """Build a profile card for a single location/sheet group of rows."""
    rows_ = []
    for pat, label in _SQ_CUST_PROFILE_PATS:
        col = find_col(gdf, pat)
        if not col or col.startswith("_"):
            continue
        vals = gdf[col].dropna().astype(str).str.strip().unique().tolist()
        vals = [v for v in vals if v and v.lower() not in ("none","nan","")]
        if vals:
            rows_.append({"Field": label, "Value": ", ".join(vals[:4])})
    return pd.DataFrame(rows_) if rows_ else pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 – SMART QUERY  (merged: AI query + customer name-wise full row search)
# ─────────────────────────────────────────────────────────────────────────────
with T[4]:
    st.markdown('<div class="section-title">🧠 Smart Query — AI-Powered Structured Query Engine</div>',
                unsafe_allow_html=True)

    # ── Data source ───────────────────────────────────────────────────────────
    sq_src_opts = ["All Locations & All Sheets"] + sorted(fdata.keys())
    sq_src = st.selectbox("📂 Query data source", sq_src_opts, key="sq_src")

    if sq_src == "All Locations & All Sheets":
        pool_base = _sq_preprocess_pool(CUST.copy())
    else:
        loc_frames = []
        for sn, df_loc in fdata.get(sq_src, {}).items():
            tmp = df_loc.copy()
            tmp.insert(0, "_Sheet", sn)
            tmp.insert(0, "_Location", sq_src)
            loc_frames.append(tmp)
        _raw_pool = pd.concat(loc_frames, ignore_index=True, sort=False) if loc_frames else pd.DataFrame()
        pool_base = _sq_preprocess_pool(_raw_pool)

    if not pool_base.empty:
        n_locs   = pool_base["_Location"].nunique() if "_Location" in pool_base.columns else 1
        n_sheets = pool_base["_Sheet"].nunique()    if "_Sheet"    in pool_base.columns else 1
        # Count how many columns were promoted to numeric by pre-processing
        _nc_after = num_cols(pool_base)
        st.markdown(
            f'<div style="font-size:.78rem;color:{MUTED};margin-bottom:6px">'
            f'Query pool: <b style="color:{CYAN}">{len(pool_base):,}</b> records · '
            f'<b style="color:{CYAN}">{n_locs}</b> location(s) · '
            f'<b style="color:{CYAN}">{n_sheets}</b> sheet(s) · '
            f'<b style="color:{CYAN}">{len(_nc_after)}</b> numeric columns available</div>',
            unsafe_allow_html=True)

        # ── Schema awareness panel ────────────────────────────────────────────
        with st.expander("📐 Data Schema — available columns & field hints", expanded=False):
            _schema_cols = [c for c in pool_base.columns if not c.startswith("_")]
            _num_set = set(_nc_after)
            _schema_rows = []
            for c in _schema_cols:
                dtype  = "numeric" if c in _num_set else "text"
                sample = pool_base[c].dropna()
                s_val  = str(sample.iloc[0])[:40] if not sample.empty else "—"
                _schema_rows.append({"Column": c, "Type": dtype, "Sample value": s_val})
            if _schema_rows:
                st.dataframe(pd.DataFrame(_schema_rows), use_container_width=True)
            st.markdown(
                f'<div style="font-size:.76rem;color:{MUTED};margin-top:8px">'
                f'<b>Key field hints for Smart Query:</b><br>'
                f'<code>total capacity purchased</code> · '
                f'<code>power in use</code> · '
                f'<code>total revenue</code> · '
                f'<code>space in use</code> · '
                f'<code>per unit rate</code> · '
                f'<code>seating subscription</code> · '
                f'<code>net revenue</code>'
                f'</div>',
                unsafe_allow_html=True)

    # ── NEW STRUCTURED QUERY ENGINE (no AI — 100 % accurate, direct DataFrame ops) ──
    # Two query modes:
    #   Mode A: Customer Name Search  — find ALL rows matching a customer across every file & sheet
    #   Mode B: Column & Operations  — pick a real column, apply operations, optional filters
    # Location-wise column headers derived from locationheaders.txt (actual Excel sub-headers)
    # ─────────────────────────────────────────────────────────────────────────────────

    # ── Embedded location-wise canonical column headers (from locationheaders.txt) ──
    _LOC_CANON_HEADERS = {
        "Airoli": [
            "Sr. No", "FLOOR", "SH", "Customer Name",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode (Rack/U Space/SqFt Space)",
            "Ownership(Sify/Customer)", "Caged /Uncaged",
            "Space | Subscription", "Space | In Use",
            "Power Capacity | Subscription Model (Rated/Subscribed)",
            "Power Capacity | UoM (KVA/KW)",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Usage in KW",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Capacity | Usage Model (Bundled/Metered)",
            "Power Capacity | Multiplier",
            "Power Capacity | Unit rate Model (Fixed/Variable)",
            "Power Capacity | Unit Rate (per KW-HR)",
            "Power Capacity | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model (No. of Seats/Space)",
            "Seating Space | Enclosed/Shared",
            "Seating Space | Subscription", "Seating Space | In Use",
        ],
        "Bangalore 01": [
            "Floor", "Floor / Module", "Customer Name", "RHS/SH",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM", "Space | Subscription", "Space | In Use",
            "Space | Yet to be given/", "Space | Billed",
            "Space | Reserved Capacity if any (Non-Billable)",
            "Space | Per Unit rate (MRC)",
            "Power Capacity | Subscription Model",
            "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Capacity to be given",
            "Power Capacity | Reserved Capacity if any",
            "Power Capacity | Subscribed Capacity to be given in KW",
            "Power Capacity | Allocated Capacity in KW",
            "Power Capacity | DC NW Infra",
            "Power Capacity | Usage in KW",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Usage | Usage Model", "Power Usage | Multiplier",
            "Power Usage | Unit rate Model (Fixed/Variable)",
            "Power Usage | Unit Rate (per KW-HR)",
            "Power Usage | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model", "Seating Space | UoM",
            "Seating Space | Subscription", "Seating Space | In Use",
            "Seating Space | Yet to be given", "Seating Space | Billed",
            "Seating Space | Reserved Capacity if any",
            "Seating Space | Per Unit rate",
            "Revenue | Space revenue including capacity",
            "Revenue | Additional Capacity Revenue",
            "Revenue | Power Usage revenue",
            "Revenue | Seating Space", "Revenue | Any Other Items",
            "Revenue | Total Revenue", "Revenue | Billing Frequency",
            "Contract | Sales Order ref No",
            "Contract | Contract Start Date",
            "Contract | Term of Contract (No of Years)",
            "Contract | Current Expiry Date",
            "Contract | Remarks if any", "Contract | Cross connect",
        ],
        "Chennai 01": [
            "Floor / Module", "Customer Name",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM", "Space | Subscription", "Space | In Use",
            "Space | Yet to be given/", "Space | Billed",
            "Space | Reserved Capacity if any (Non-Billable)",
            "Space | Per Unit rate (MRC)",
            "Power Capacity | Subscription Model",
            "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Capacity to be given",
            "Power Capacity | Reserved Capacity if any",
            "Power Capacity | Subscribed Capacity to be given in KW",
            "Power Capacity | Allocated Capacity in KW",
            "Power Capacity | Usage in KW",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Usage | Usage Model", "Power Usage | Multiplier",
            "Power Usage | Unit rate Model (Fixed/Variable)",
            "Power Usage | Unit Rate (per KW-HR)",
            "Power Usage | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model", "Seating Space | UoM",
            "Seating Space | Subscription", "Seating Space | In Use",
            "Seating Space | Yet to be given", "Seating Space | Billed",
            "Seating Space | Reserved Capacity if any",
            "Seating Space | Per Unit rate",
            "Revenue | Space revenue including capacity",
            "Revenue | Additional Capacity Revenue",
            "Revenue | Power Usage revenue",
            "Revenue | Seating Space", "Revenue | Any Other Items",
            "Revenue | Total Revenue", "Revenue | Billing Frequency",
            "Contract | Sales Order ref No",
            "Contract | Contract Start Date",
            "Contract | Term of Contract (No of Years)",
            "Contract | Current Expiry Date",
            "Contract | Remarks if any",
            "Analytics | Avg revenue /Rack /Month",
            "Analytics | Avg revenue /KW /Month",
            "Analytics | Net Revenue / Resvd KW",
            "Analytics | Total Rev (Cap + Power)",
            "Analytics | Capacity Revenue", "Analytics | Power Revenue",
            "Analytics | Net Rev Total",
            "Analytics | Power Surplus/Leakage",
            "Analytics | Capacity Surplus/Leakage",
            "Analytics | Total Surplus/Leakage",
        ],
        "Kolkata": [
            "Floor", "Floor / Module", "Customer Name", "RHS/SH",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM", "Space | Subscription", "Space | In Use",
            "Space | Yet to be given/", "Space | Billed",
            "Space | Reserved Capacity if any (Non-Billable)",
            "Space | Per Unit rate (MRC)",
            "Power Capacity | Subscription Model",
            "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Capacity to be given",
            "Power Capacity | Reserved Capacity if any",
            "Power Capacity | Subscribed Capacity to be given in KW",
            "Power Capacity | Allocated Capacity in KW",
            "Power Capacity | DC NW Infra",
            "Power Capacity | Usage in KW",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Usage | Usage Model", "Power Usage | Multiplier",
            "Power Usage | Unit rate Model (Fixed/Variable)",
            "Power Usage | Unit Rate (per KW-HR)",
            "Power Usage | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model", "Seating Space | UoM",
            "Seating Space | Subscription", "Seating Space | In Use",
            "Seating Space | Yet to be given", "Seating Space | Billed",
            "Seating Space | Reserved Capacity if any",
            "Seating Space | Per Unit rate",
            "Revenue | Space revenue including capacity",
            "Revenue | Additional Capacity Revenue",
            "Revenue | Power Usage revenue",
            "Revenue | Seating Space", "Revenue | Any Other Items",
            "Revenue | Total Revenue", "Revenue | Billing Frequency",
            "Contract | Sales Order ref No",
            "Contract | Contract Start Date",
            "Contract | Term of Contract (No of Years)",
            "Contract | Current Expiry Date",
            "Contract | Remarks if any",
        ],
        "Vashi": [
            "Floor / Module", "Customer",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM", "Space | Subscription", "Space | In Use",
            "Space | Yet to be given/", "Space | Billed",
            "Space | Reserved Capacity if any (Non-Billable)",
            "Space | Per Unit rate (MRC)",
            "Power Capacity | Subscription Model",
            "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Capacity to be given",
            "Power Capacity | Reserved Capacity if any",
            "Power Capacity | Subscribed Capacity to be given in KW",
            "Power Capacity | Allocated Capacity in KW",
            "Power Capacity | Usage in KW",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Usage | Usage Model", "Power Usage | Multiplier",
            "Power Usage | Uit rate Model (Fixed/Variable)",
            "Power Usage | Unit Rate (per KW-HR)",
            "Power Usage | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model", "Seating Space | UoM",
            "Seating Space | Subscription", "Seating Space | In Use",
            "Seating Space | Yet to be given", "Seating Space | Billed",
            "Seating Space | Reserved Capacity if any",
            "Seating Space | Per Unit rate",
            "Revenue | Space revenue including capacity",
            "Revenue | Additional Capacity Revenue",
            "Revenue | Power Usage revenue",
            "Revenue | Seating Space", "Revenue | Any Other Items",
            "Revenue | Total Revenue", "Revenue | Billing Frequency",
            "Contract | Sales Order ref No",
            "Contract | Contract Start Date",
            "Contract | Term of Contract (No of Years)",
            "Contract | Current Expiry Date",
            "Contract | Remarks if any",
            "Analytics | Avg revenue /Rack /Month",
            "Analytics | Avg revenue /KW /Month",
            "Analytics | Net Revenue / Resvd KW",
            "Analytics | Total Rev (Cap + Power)",
            "Analytics | Capacity Revenue", "Analytics | Power Revenue",
            "Analytics | Net Rev Total",
            "Analytics | Power Surplus/Leakage",
            "Analytics | Capacity Surplus/Leakage",
            "Analytics | Total Surplus/Leakage",
        ],
        "Noida 01": [
            "Floor / Module", "Customer Name", "RHS/SH",
            "Sitting Space (Subscription)", "IR DATE",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM", "Space | Subscription", "Space | In Use",
            "Space | Yet to be given/", "Space | Billed",
            "Space | Reserved Capacity if any (Non-Billable)",
            "Space | Per Unit rate (MRC)",
            "Power Capacity | Subscription Model", "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Capacity to be given",
            "Power Capacity | Reserved Capacity if any",
            "Power Capacity | Subscribed Capacity to be given in KW",
            "Power Capacity | Allocated Capacity in KW (for KVA subscribed customer 50% diversity)",
            "Power Capacity | DC NW Infra",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Usage | Usage Model", "Power Usage | Multiplier",
            "Power Usage | Uit rate Model (Fixed/Variable)",
            "Power Usage | Unit Rate (per KW-HR)",
            "Power Usage | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model", "Seating Space | UoM",
            "Seating Space | Subscription", "Seating Space | In Use",
            "Seating Space | Yet to be given", "Seating Space | Billed",
            "Seating Space | Reserved Capacity if any",
            "Seating Space | Per Unit rate",
            "Revenue | Space revenue including capacity",
            "Revenue | Additional Capacity Revenue",
            "Revenue | Power Usage revenue",
            "Revenue | Seating Space", "Revenue | Any Other Items",
            "Revenue | Total Revenue", "Revenue | Billing Frequency",
            "Contract | Sales Order ref No",
            "Contract | Contract Start Date",
            "Contract | Term of Contract (No of Years)",
            "Contract | Current Expiry Date",
            "Contract | Remarks if any",
            "Analytics | Avg revenue /Rack /Month",
            "Analytics | Avg revenue /KW /Month",
            "Analytics | Net Revenue / Resvd KW",
            "Analytics | Total Rev (Cap + Power)",
            "Analytics | Capacity Revenue", "Analytics | Power Revenue",
            "Analytics | Net Rev Total",
            "Analytics | Power Surplus/Leakage",
            "Analytics | Capacity Surplus/Leakage",
            "Analytics | Total Surplus/Leakage",
        ],
        "Noida 02": [
            "Floor / Module", "Customer Name", "RHS/SH",
            "Sitting Space (Subscription)", "IR DATE",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM", "Space | Subscription", "Space | In Use",
            "Space | Yet to be given/", "Space | Billed",
            "Space | Reserved Capacity if any (Non-Billable)",
            "Space | Per Unit rate (MRC)",
            "Power Capacity | Subscription Model", "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Capacity to be given",
            "Power Capacity | Reserved Capacity if any",
            "Power Capacity | Subscribed Capacity to be given in KW",
            "Power Capacity | Allocated Capacity in KW (for KVA subscribed customer 50% diversity)",
            "Power Capacity | DC NW Infra",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Usage | Usage Model", "Power Usage | Multiplier",
            "Power Usage | Uit rate Model (Fixed/Variable)",
            "Power Usage | Unit Rate (per KW-HR)",
            "Power Usage | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model", "Seating Space | UoM",
            "Seating Space | Subscription", "Seating Space | In Use",
            "Seating Space | Yet to be given", "Seating Space | Billed",
            "Seating Space | Reserved Capacity if any",
            "Seating Space | Per Unit rate",
            "Revenue | Space revenue including capacity",
            "Revenue | Additional Capacity Revenue",
            "Revenue | Power Usage revenue",
            "Revenue | Seating Space", "Revenue | Any Other Items",
            "Revenue | Total Revenue", "Revenue | Billing Frequency",
            "Contract | Sales Order ref No",
            "Contract | Contract Start Date",
            "Contract | Term of Contract (No of Years)",
            "Contract | Current Expiry Date",
            "Contract | Remarks if any",
        ],
        "Rabale T1 T2": [
            "Floor / Module", "Customer Name",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM", "Space | Subscription", "Space | In Use",
            "Space | Yet to be given/", "Space | Billed",
            "Space | Reserved Capacity if any (Non-Billable)",
            "Space | Per Unit rate (MRC)", "Space | ARC",
            "Power Capacity | Subscription Model", "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased",
            "Power Capacity | Capacity in Use",
            "Power Capacity | Capacity to be given",
            "Power Capacity | Reserved Capacity if any",
            "Power Capacity | Subscribed Capacity to be given in KW",
            "Power Capacity | Allocated Capacity in KW",
            "Power Capacity | Usage in KW",
            "Power Capacity | Billable Additional Capacity",
            "Power Capacity | Additional Capacity Charges (MRC)",
            "Power Usage | Usage Model", "Power Usage | Multiplier",
            "Power Usage | Uit rate Model (Fixed/Variable)",
            "Power Usage | Unit Rate (per KW-HR)",
            "Power Usage | No Of Units (KW-HR/ Month)",
            "Seating Space | Subscription Model", "Seating Space | UoM",
            "Seating Space | Subscription", "Seating Space | In Use",
            "Seating Space | Yet to be given", "Seating Space | Billed",
            "Seating Space | Reserved Capacity if any",
            "Seating Space | Per Unit rate",
            "Revenue | Space revenue including capacity",
            "Revenue | Additional Capacity Revenue",
            "Revenue | Power Usage revenue",
            "Revenue | Seating Space", "Revenue | Any Other Items",
            "Revenue | Total Revenue (MRC)", "Revenue | ARC",
            "Revenue | Billing Frequency",
            "Contract | Sales Order ref No",
            "Contract | Contract Start Date",
            "Contract | Term of Contract (No of Years)",
            "Contract | Current Expiry Date",
            "Contract | Remarks if any",
            "Analytics | Avg revenue /Rack /Month",
            "Analytics | Avg revenue /KW /Month",
            "Analytics | Net Revenue / Resvd KW",
            "Analytics | Total Rev (Cap + Power)",
            "Analytics | Capacity Revenue", "Analytics | Power Revenue",
            "Analytics | Net Rev Total",
            "Analytics | Power Surplus/Leakage",
            "Analytics | Capacity Surplus/Leakage",
            "Analytics | Total Surplus/Leakage",
            "Power Usage (KW) | Maximum Usable Capacity",
            "Power Usage (KW) | Current utilization",
            "Power Usage (KW) | Committed (Based on Confirmed orders)",
            "Power Usage (KW) | Total", "Power Usage (KW) | Balance",
        ],
        "Rabale Tower 4": [
            "Floor / Module", "Customer Name",
            "Power Subscription Model (Rated/Subscribed)",
            "Power Usage Model (Bundled / Metered)",
            "Subscription Mode", "Caged /Uncaged",
            "Space | UoM",
            "Space | Subscription(No. of Racks)", "Space | In Use(No. of Racks)",
            "Power Capacity | UoM",
            "Power Capacity | Total Capacity Purchased (KW)",
            "Power Capacity | Capacity in Use (KW)",
        ],
        "Rabale Tower 5": [
            "Floor", "Tower -5 (MUM - 03)",
            "Space | Subscription Mode", "Space | UoM",
            "Space | Occupied in Sqft",
            "IT KW CAPACITY | Total Capacity - Server Hall",
            "IT KW CAPACITY | Sold", "IT KW CAPACITY | Available",
            "Remarks",
        ],
    }

    # ── Operations map (all supported numeric operations) ─────────────────────
    _SQ_ALL_OPS = {
        "Sum":                   ("sum",      lambda s: s.sum()),
        "Average (Mean)":        ("avg",      lambda s: s.mean()),
        "Median":                ("median",   lambda s: s.median()),
        "Min":                   ("min",      lambda s: s.min()),
        "Max":                   ("max",      lambda s: s.max()),
        "Count":                 ("count",    lambda s: float(len(s))),
        "Count Non-Zero":        ("count_nz", lambda s: float((s != 0).sum())),
        "Std Deviation":         ("std",      lambda s: s.std(ddof=1)),
        "Variance":              ("var",      lambda s: s.var(ddof=1)),
        "Range (Max − Min)":     ("range",    lambda s: s.max() - s.min()),
        "Product (Multiply)":    ("prod",     lambda s: s.prod()),
        "% of Grand Total":      ("pct",      lambda s: (s.sum() / s.sum() * 100) if s.sum() != 0 else 0.0),
        "Cumulative Sum":        ("cumsum",   lambda s: s.cumsum().iloc[-1] if not s.empty else 0.0),
        "Top 10 Values":         ("top10",    None),
        "Bottom 10 Values":      ("bot10",    None),
        "Show All Matching Rows": ("rows",   None),
    }

    def _sq_fuzzy_col(df_: pd.DataFrame, canonical: str) -> "str | None":
        """Fuzzy-match a canonical column hint to a real column in df_."""
        if canonical in df_.columns:
            return canonical
        c_lo = canonical.lower()
        # Try progressively looser matches
        # 1. Contains the full canonical name (case-insensitive)
        for c in df_.columns:
            if c_lo in c.lower() or c.lower() in c_lo:
                return c
        # 2. Strip pipe grouping — match sub-header part only
        sub = c_lo.split("|")[-1].strip() if "|" in c_lo else c_lo
        if len(sub) > 3:
            for c in df_.columns:
                if sub in c.lower():
                    return c
        # 3. All words match
        words = [w for w in re.split(r"\W+", c_lo) if len(w) > 3]
        for c in df_.columns:
            if words and all(w in c.lower() for w in words):
                return c
        return None

    # ── Mode selector ─────────────────────────────────────────────────────────
    _sq_mode_opts = [
        "🔍 Customer Name Search",
        "📊 Column & Operations Query",
    ]
    _sq_mode = st.radio(
        "Select Query Mode",
        _sq_mode_opts,
        horizontal=True,
        key="sq_mode_select",
    )

    sq_locs = st.multiselect(
        "📍 Restrict to locations (optional — leave blank for all)",
        options=sorted(fdata.keys()),
        default=[],
        key="sq_locs",
    )

    if "sq_results_history" not in st.session_state:
        st.session_state["sq_results_history"] = []

    # ══════════════════════════════════════════════════════════════════════════
    # MODE A — CUSTOMER NAME SEARCH
    # Searches every file, every sheet for rows where Customer Name column
    # contains the entered text (case-insensitive, partial match).
    # Returns ALL columns for every matching row — zero hallucination.
    # ══════════════════════════════════════════════════════════════════════════
    if _sq_mode == "🔍 Customer Name Search":
        st.markdown(
            f'<div style="font-size:.82rem;color:{MUTED};margin-bottom:8px">'
            f'Enter a customer name (or part of it) to search <b>all 10 Excel files</b> '
            f'and <b>all sheets</b>. Every matching row with all its columns is returned.'
            f'</div>',
            unsafe_allow_html=True,
        )

        _cust_col1, _cust_col2 = st.columns([3, 1])
        with _cust_col1:
            _cust_input = st.text_input(
                "👤 Customer Name",
                placeholder="e.g.  Wipro  |  Oracle  |  YES BANK  |  CISCO  |  Tata",
                key="sq_cust_input",
            )
        with _cust_col2:
            _cust_match = st.selectbox(
                "Match Mode",
                ["Contains (partial)", "Exact", "Starts with"],
                key="sq_cust_match",
            )

        _cust_sheets_opts = sorted({
            sn for loc in (sq_locs if sq_locs else fdata.keys())
            for sn in fdata.get(loc, {})
        })
        _cust_sheets_sel = st.multiselect(
            "📄 Filter by sheets (optional)",
            options=_cust_sheets_opts,
            default=[],
            key="sq_cust_sheets",
        )

        _cust_additional_col = st.multiselect(
            "📋 Show only these columns in results (leave blank = show all columns)",
            options=[c for c in pool_base.columns if not c.startswith("_")] if not pool_base.empty else [],
            default=[],
            key="sq_cust_cols",
        )

        _crun_c, _ = st.columns([1, 6])
        with _crun_c:
            _cust_run_btn = st.button("🔍 Search Customer", key="sq_cust_run")

        if _cust_run_btn and _cust_input.strip():
            _search_pool = combined_df(ALL)  # ALL files, ALL sheets
            if sq_locs and "_Location" in _search_pool.columns:
                _search_pool = _search_pool[_search_pool["_Location"].isin(sq_locs)]
            if _cust_sheets_sel and "_Sheet" in _search_pool.columns:
                _search_pool = _search_pool[_search_pool["_Sheet"].isin(_cust_sheets_sel)]

            # Find customer name column
            _cn_col = find_col(_search_pool, r"customer.*name|client.*name|^customer$")
            if _cn_col is None:
                # Also try bare "Customer" for Vashi
                _cn_col = find_col(_search_pool, r"\bcustomer\b")
            _srch = _cust_input.strip()
            _srch_lo = _srch.lower()

            if _cn_col and _cn_col in _search_pool.columns:
                _vals = _search_pool[_cn_col].astype(str)
                if _cust_match == "Exact":
                    _mask = _vals.str.lower() == _srch_lo
                elif _cust_match == "Starts with":
                    _mask = _vals.str.lower().str.startswith(_srch_lo)
                else:
                    _mask = _vals.str.lower().str.contains(re.escape(_srch_lo), na=False)
                _found = _search_pool[_mask].copy()
            else:
                # Fallback: search all text columns
                _str_cols = [c for c in _search_pool.columns
                             if not c.startswith("_") and _search_pool[c].dtype == object]
                _mask2 = pd.Series(False, index=_search_pool.index)
                for _sc in _str_cols:
                    _mask2 |= _search_pool[_sc].astype(str).str.lower().str.contains(
                        re.escape(_srch_lo), na=False)
                _found = _search_pool[_mask2].copy()

            if _found.empty:
                st.warning(f"No rows found for customer **'{_srch}'** in the selected scope.")
            else:
                # Display summary card
                _f_locs  = sorted(_found["_Location"].unique()) if "_Location" in _found.columns else []
                _f_sheets = sorted(_found["_Sheet"].unique())   if "_Sheet"    in _found.columns else []
                st.markdown(
                    f'<div style="background:{DARK2};border:2px solid {CYAN};'
                    f'border-radius:14px;padding:20px 26px;margin:14px 0">'
                    f'<div style="font-size:1.05rem;font-weight:900;color:{WHITE}">'
                    f'✅ Found: <span style="color:{CYAN}">{_srch}</span></div>'
                    f'<div style="font-size:.84rem;color:{TEXT};margin-top:6px">'
                    f'<b style="color:{CYAN}">{len(_found):,}</b> matching row(s) across '
                    f'<b style="color:{CYAN}">{len(_f_locs)}</b> location(s) · '
                    f'<b style="color:{CYAN}">{len(_f_sheets)}</b> sheet(s)</div>'
                    f'<div style="margin-top:8px;font-size:.77rem;color:{MUTED}">Locations: '
                    + (", ".join(f'<span style="color:{GREEN}">{l}</span>' for l in _f_locs) or "—")
                    + f'</div><div style="font-size:.77rem;color:{MUTED};margin-top:3px">Sheets: '
                    + (", ".join(f'<span style="color:{AMBER}">{s}</span>' for s in _f_sheets[:20]) or "—")
                    + f'</div></div>',
                    unsafe_allow_html=True,
                )

                # Row-count validation breakdown
                with st.expander("🔬 Validation — row count per file & sheet", expanded=False):
                    if "_Location" in _found.columns and "_Sheet" in _found.columns:
                        _val_df = (
                            _found.groupby(["_Location", "_Sheet"])
                            .size().reset_index(name="Row Count")
                            .sort_values(["_Location", "_Sheet"])
                        )
                        _val_df.index = range(1, len(_val_df) + 1)
                        st.dataframe(_val_df, use_container_width=True)

                # Build display dataframe
                _meta_c = [c for c in ["_Location", "_Sheet"] if c in _found.columns]
                _data_c = (
                    _cust_additional_col
                    if _cust_additional_col
                    else [c for c in _found.columns if not c.startswith("_")]
                )
                _disp = _found[_meta_c + [c for c in _data_c if c in _found.columns]].copy()
                _disp.index = range(1, len(_disp) + 1)

                st.markdown(
                    f'<div style="font-size:.8rem;color:{CYAN};font-weight:700;'
                    f'text-transform:uppercase;letter-spacing:.05em;margin:14px 0 6px">'
                    f'📋 All Data — {len(_found):,} row(s) for "{_srch}"</div>',
                    unsafe_allow_html=True,
                )
                st.dataframe(_disp, use_container_width=True)
                st.download_button(
                    f"⬇️ Download CSV — {_srch[:40]}",
                    _disp.to_csv(index=False).encode("utf-8"),
                    f"customer_{_srch.replace(' ','_')[:40]}.csv",
                    "text/csv",
                    key="sq_cust_dl",
                )

                # Save to results history (table format compatible with display block)
                _res_entry = {
                    "type": "table", "label": f"Customer Search: {_srch}",
                    "data": _disp, "row_count": len(_found),
                }
                st.session_state["sq_results_history"].append({
                    "query":   f"Customer search: {_srch}",
                    "source":  sq_src,
                    "records": len(_found),
                    "results": [_res_entry],
                })

    # ══════════════════════════════════════════════════════════════════════════
    # MODE B — COLUMN & OPERATIONS QUERY
    # User selects: location → column (from real loaded columns, guided by
    # locationheaders.txt) → operation → optional filters.
    # 100 % accurate — executed directly on the real DataFrames. No AI/LLM.
    # ══════════════════════════════════════════════════════════════════════════
    else:
        st.markdown(
            f'<div style="font-size:.82rem;color:{MUTED};margin-bottom:8px">'
            f'Select a location, column, and operation to run directly on the real Excel data. '
            f'All columns listed are real sub-headers from the actual Excel files. '
            f'No guessing — 100% accurate results.</div>',
            unsafe_allow_html=True,
        )

        _op_r1, _op_r2 = st.columns([2, 3])
        with _op_r1:
            _op_loc = st.selectbox(
                "📍 Location",
                ["All Locations"] + sorted(fdata.keys()),
                key="sq_op_loc",
            )

        # Build the working pool for this location
        if _op_loc == "All Locations":
            _op_pool = pool_base.copy()
            if sq_locs and "_Location" in _op_pool.columns:
                _op_pool = _op_pool[_op_pool["_Location"].isin(sq_locs)]
        else:
            _op_frames = []
            for _sn, _df_loc in fdata.get(_op_loc, {}).items():
                _tmp = _df_loc.copy()
                _tmp.insert(0, "_Sheet", _sn)
                _tmp.insert(0, "_Location", _op_loc)
                _op_frames.append(_tmp)
            _op_pool = pd.concat(_op_frames, ignore_index=True, sort=False) if _op_frames else pd.DataFrame()

        # Real columns from loaded data
        _op_real_cols = [c for c in _op_pool.columns if not c.startswith("_")] if not _op_pool.empty else []
        _op_num_cols  = [c for c in _op_real_cols if pd.api.types.is_numeric_dtype(_op_pool[c])] if not _op_pool.empty else []
        _op_txt_cols  = [c for c in _op_real_cols if c not in _op_num_cols]

        with _op_r2:
            _op_col_type = st.radio(
                "Column type",
                ["📊 Numeric columns", "📝 Text / Category columns"],
                horizontal=True, key="sq_op_col_type",
            )

        _col_list = _op_num_cols if "Numeric" in _op_col_type else _op_txt_cols

        _op_r3, _op_r4 = st.columns([3, 2])
        with _op_r3:
            # Search box to filter column list
            _col_search = st.text_input(
                "🔎 Search column (partial match)",
                placeholder="e.g. Revenue | Capacity | Customer | Subscription",
                key="sq_col_search",
            )
            _col_filtered = (
                [c for c in _col_list if _col_search.strip().lower() in c.lower()]
                if _col_search.strip() else _col_list
            )
            _sel_col = st.selectbox(
                f"📋 Select Column — {len(_col_filtered)} available",
                ["— pick a column —"] + _col_filtered,
                key="sq_sel_col",
            )
        with _op_r4:
            if "Numeric" in _op_col_type:
                _sel_op_name = st.selectbox(
                    "⚙️ Operation",
                    list(_SQ_ALL_OPS.keys()),
                    key="sq_sel_op",
                )
            else:
                _sel_op_name = st.selectbox(
                    "⚙️ Operation",
                    ["Count (non-null)", "Unique Count", "Value Counts", "Show All Rows"],
                    key="sq_sel_op",
                )

        # Optional filters row
        _op_f1, _op_f2, _op_f3 = st.columns([2, 2, 1])
        with _op_f1:
            _op_cust_filter = st.text_input(
                "👤 Customer Name Filter (optional)",
                placeholder="e.g.  Wipro  |  Oracle",
                key="sq_op_cust_filter",
            )
        with _op_f2:
            _op_sheet_opts = sorted({
                sn for loc in (sq_locs if sq_locs else (
                    [_op_loc] if _op_loc != "All Locations" else fdata.keys()))
                for sn in fdata.get(loc, {})
            })
            _op_sheets = st.multiselect(
                "📄 Sheets (optional)",
                options=_op_sheet_opts, default=[],
                key="sq_op_sheets",
            )
        with _op_f3:
            _op_grp_loc = st.checkbox("Group by Location", value=True, key="sq_op_grp")

        _op_top_n = 10
        if _sel_op_name in ("Top 10 Values", "Bottom 10 Values"):
            _op_top_n = st.number_input(
                "N", min_value=1, max_value=100, value=10, step=5, key="sq_op_top_n"
            )

        _run_op_c, _ = st.columns([1, 6])
        with _run_op_c:
            _op_run_btn = st.button("▶ Run Query", key="sq_op_run")

        if _op_run_btn and _sel_col != "— pick a column —":
            # Apply sheet filter
            _work = _op_pool.copy()
            if _op_sheets and "_Sheet" in _work.columns:
                _work = _work[_work["_Sheet"].isin(_op_sheets)]

            # Apply customer name filter
            if _op_cust_filter.strip():
                _cn_c2 = find_col(_work, r"customer.*name|client.*name|^customer$")
                if _cn_c2 and _cn_c2 in _work.columns:
                    _work = _work[
                        _work[_cn_c2].astype(str).str.lower().str.contains(
                            re.escape(_op_cust_filter.strip().lower()), na=False)
                    ]

            if _work.empty:
                st.warning("No records match the applied filters.")
            elif _sel_col not in _work.columns:
                st.error(f"Column '{_sel_col}' not found in selected scope.")
            else:
                _q_label = (
                    f"{_sel_op_name} of '{_sel_col}'"
                    + (f" [Customer: {_op_cust_filter}]" if _op_cust_filter.strip() else "")
                    + (f" [{_op_loc}]" if _op_loc != "All Locations" else " [All Locations]")
                )

                # ── Text column operations ─────────────────────────────────────
                if "Numeric" not in _op_col_type:
                    _tc_data = _work[_sel_col].dropna().astype(str)
                    if _sel_op_name == "Count (non-null)":
                        st.metric("Count (non-null)", f"{len(_tc_data):,}")
                    elif _sel_op_name == "Unique Count":
                        st.metric("Unique Values", f"{_tc_data.nunique():,}")
                    elif _sel_op_name == "Value Counts":
                        _vc = _tc_data.value_counts().reset_index()
                        _vc.columns = [_sel_col, "Count"]
                        _vc.index = range(1, len(_vc) + 1)
                        st.dataframe(_vc, use_container_width=True)
                    else:  # Show All Rows
                        _meta_r = [c for c in ["_Location", "_Sheet"] if c in _work.columns]
                        _cn_txt_r = find_col(_work, r"customer.*name|client.*name|^customer$")
                        _txt_extra = [_cn_txt_r] if (_cn_txt_r and _cn_txt_r not in _meta_r and _cn_txt_r in _work.columns) else []
                        _disp_r = _work[_meta_r + _txt_extra + [_sel_col]].copy()
                        _disp_r.index = range(1, len(_disp_r) + 1)
                        st.dataframe(_disp_r, use_container_width=True)
                        st.download_button(
                            "⬇️ Download CSV",
                            _disp_r.to_csv(index=False).encode(),
                            f"query_{_sel_col[:30]}.csv", "text/csv",
                            key="sq_txt_dl",
                        )
                    _res_txt = {
                        "type": "table", "label": _q_label,
                        "data": _work[[c for c in ["_Location", "_Sheet", _sel_col] if c in _work.columns]],
                        "row_count": len(_work),
                    }
                    st.session_state["sq_results_history"].append({
                        "query": _q_label, "source": sq_src,
                        "records": len(_work), "results": [_res_txt],
                    })

                # ── Numeric column operations ──────────────────────────────────
                else:
                    _num_s = _robust_to_numeric(_work[_sel_col]).dropna()
                    _total_rows = len(_work)
                    _valid_rows = len(_num_s)

                    if _num_s.empty:
                        st.warning(
                            f"Column **'{_sel_col}'** has no numeric values in the current scope."
                        )
                    else:
                        _op_key, _op_fn = _SQ_ALL_OPS[_sel_op_name][0], _SQ_ALL_OPS[_sel_op_name][1]
                        _unit = _detect_unit(_sel_col)

                        # Special handling: Top N / Bottom N / Show All Rows
                        if _sel_op_name in ("Top 10 Values", "Bottom 10 Values", "Show All Matching Rows"):
                            _meta_c2 = [c for c in ["_Location", "_Sheet"] if c in _work.columns]
                            _cn_c3   = find_col(_work, r"customer.*name|client.*name|^customer$")
                            _show_c  = _meta_c2 + ([_cn_c3] if _cn_c3 else []) + [_sel_col]
                            _show_df = _work[[c for c in _show_c if c in _work.columns]].copy()
                            _show_df[_sel_col] = _robust_to_numeric(_show_df[_sel_col])
                            _show_df = _show_df.dropna(subset=[_sel_col])
                            if _sel_op_name == "Top 10 Values":
                                _show_df = _show_df.nlargest(int(_op_top_n), _sel_col)
                            elif _sel_op_name == "Bottom 10 Values":
                                _show_df = _show_df.nsmallest(int(_op_top_n), _sel_col)
                            _show_df = _show_df.reset_index(drop=True)
                            _show_df.index = range(1, len(_show_df) + 1)
                            st.dataframe(_show_df, use_container_width=True)
                            st.download_button(
                                f"⬇️ Download {_sel_op_name} CSV",
                                _show_df.to_csv(index=False).encode(),
                                f"query_{_sel_op_name.replace(' ','_')[:20]}.csv",
                                "text/csv", key="sq_topn_dl",
                            )
                            _res_e = {
                                "type": "table", "label": _q_label,
                                "data": _show_df, "row_count": len(_show_df),
                            }
                            st.session_state["sq_results_history"].append({
                                "query": _q_label, "source": sq_src,
                                "records": len(_work), "results": [_res_e],
                            })

                        elif _sel_op_name == "% of Grand Total":
                            # Each location's share of the grand total
                            _grand = _num_s.sum()
                            if _op_grp_loc and "_Location" in _work.columns:
                                _pct_rows = []
                                for _ploc, _pgrp in _work.groupby("_Location"):
                                    _ps = _robust_to_numeric(_pgrp[_sel_col]).dropna()
                                    _pct_rows.append({
                                        "Location": _ploc,
                                        f"{_sel_col} (Sum)": round(_ps.sum(), 4),
                                        "% of Grand Total": round(_ps.sum() / _grand * 100, 2) if _grand else 0,
                                    })
                                _pct_df = pd.DataFrame(_pct_rows).sort_values(
                                    "% of Grand Total", ascending=False
                                ).reset_index(drop=True)
                                _pct_df.index = range(1, len(_pct_df) + 1)
                                st.dataframe(_pct_df, use_container_width=True)
                            else:
                                st.metric(
                                    f"Grand Total of '{_sel_col}'",
                                    f"{_grand:,.2f} {_unit}".strip(),
                                )
                            _res_pct = {
                                "type": "scalar", "label": _q_label,
                                "value": float(_grand), "unit": _unit,
                                "column": _sel_col, "col_reason": "sum for % calc",
                                "row_count": _total_rows,
                                "valid_count": _valid_rows,
                                "operation": "pct", "loc_breakdown": None, "auto_loc": None,
                            }
                            st.session_state["sq_results_history"].append({
                                "query": _q_label, "source": sq_src,
                                "records": _total_rows, "results": [_res_pct],
                            })

                        elif _sel_op_name == "Cumulative Sum":
                            _cn_cum = find_col(_work, r"customer.*name|client.*name|^customer$")
                            _cum_base = _work.copy()
                            _cum_base[_sel_col] = _robust_to_numeric(_cum_base[_sel_col])
                            _cum_base = _cum_base.dropna(subset=[_sel_col]).reset_index(drop=True)
                            _cum_dict = {"Row #": range(1, len(_cum_base) + 1)}
                            if _cn_cum and _cn_cum in _cum_base.columns:
                                _cum_dict["Customer Name"] = _cum_base[_cn_cum].astype(str).values
                            _cum_dict[_sel_col] = _cum_base[_sel_col].values
                            _cum_dict["Cumulative Sum"] = _cum_base[_sel_col].cumsum().values
                            _cum_df = pd.DataFrame(_cum_dict)
                            st.dataframe(_cum_df, use_container_width=True)
                            _res_cum = {
                                "type": "table", "label": _q_label,
                                "data": _cum_df, "row_count": len(_cum_df),
                            }
                            st.session_state["sq_results_history"].append({
                                "query": _q_label, "source": sq_src,
                                "records": _total_rows, "results": [_res_cum],
                            })

                        else:
                            # Standard scalar operation
                            _val = float(_op_fn(_num_s))
                            _pct_v = _valid_rows / _total_rows * 100 if _total_rows else 0

                            # Format display value
                            if _unit == "₹":
                                if abs(_val) >= 1_00_00_000:
                                    _val_disp = f"₹ {_val/1_00_00_000:,.2f} Cr"
                                elif abs(_val) >= 1_00_000:
                                    _val_disp = f"₹ {_val/1_00_000:,.2f} L"
                                else:
                                    _val_disp = f"₹ {_val:,.2f}"
                            else:
                                _val_disp = f"{_val:,.4f} {_unit}".strip() if abs(_val) < 1e10 else f"{_val:.4e}"

                            st.markdown(f"""
                            <div style="background:{DARK2};border:1px solid {BORD};
                                 border-radius:14px;padding:22px 30px;margin:10px 0">
                              <div style="font-size:.72rem;color:{MUTED};font-weight:700;
                                   text-transform:uppercase;letter-spacing:.07em;
                                   margin-bottom:8px">{_sel_op_name}</div>
                              <div style="font-size:2.2rem;font-weight:900;color:{CYAN};
                                   line-height:1.1">{_val_disp}</div>
                              <div style="font-size:.74rem;color:{MUTED};margin-top:10px;
                                   border-top:1px solid {BORD};padding-top:8px">
                                📊 Column: <b style="color:{TEXT}">{_sel_col}</b><br>
                                ✅ <b style="color:{GREEN}">{_valid_rows:,}</b> of
                                <b style="color:{TEXT}">{_total_rows:,}</b> rows had numeric values
                                ({_pct_v:.0f}%)
                                {"⚠️ Many blank/text values — result may be partial" if _pct_v < 50 else ""}
                              </div>
                            </div>""", unsafe_allow_html=True)

                            # Per-location breakdown
                            if _op_grp_loc and "_Location" in _work.columns:
                                _loc_rows_2 = []
                                for _ll, _lg in _work.groupby("_Location"):
                                    _ls = _robust_to_numeric(_lg[_sel_col]).dropna()
                                    if not _ls.empty:
                                        _lv = float(_op_fn(_ls))
                                        _loc_rows_2.append({
                                            "Location": _ll,
                                            f"{_sel_op_name} of {_sel_col}": round(_lv, 4),
                                            "Valid Rows": len(_ls),
                                        })
                                if _loc_rows_2:
                                    _loc_df2 = pd.DataFrame(_loc_rows_2).sort_values(
                                        f"{_sel_op_name} of {_sel_col}", ascending=False
                                    ).reset_index(drop=True)
                                    _loc_df2.index = range(1, len(_loc_df2) + 1)
                                    with st.expander("📍 Per-location breakdown", expanded=True):
                                        st.dataframe(_loc_df2, use_container_width=True)

                            # Per-customer breakdown (mandatory customer name in results)
                            _cn_sc2 = find_col(_work, r"customer.*name|client.*name|^customer$")
                            if _cn_sc2 and _cn_sc2 in _work.columns:
                                _cust_sc_rows = []
                                for _cc, _cg in _work.groupby(_cn_sc2):
                                    _cs2 = _robust_to_numeric(_cg[_sel_col]).dropna()
                                    if not _cs2.empty:
                                        _cust_sc_rows.append({
                                            "Customer Name": _cc,
                                            f"{_sel_op_name} of {_sel_col}": round(float(_op_fn(_cs2)), 4),
                                            "Valid Rows": len(_cs2),
                                        })
                                if _cust_sc_rows:
                                    _cust_sc_df = pd.DataFrame(_cust_sc_rows).sort_values(
                                        f"{_sel_op_name} of {_sel_col}", ascending=False
                                    ).reset_index(drop=True)
                                    _cust_sc_df.index = range(1, len(_cust_sc_df) + 1)
                                    with st.expander("👤 Per-customer breakdown", expanded=True):
                                        st.dataframe(_cust_sc_df, use_container_width=True)
                                        st.download_button(
                                            "⬇️ Download Customer Breakdown CSV",
                                            _cust_sc_df.to_csv(index=False).encode("utf-8"),
                                            f"customer_breakdown_{_sel_col[:30].replace(' ','_')}.csv",
                                            "text/csv",
                                            key="sq_cust_sc_dl",
                                        )

                            # Save to history
                            _auto_loc2 = None
                            if _op_grp_loc and "_Location" in _work.columns:
                                _grp2b = _work.groupby("_Location")[_sel_col].apply(
                                    lambda x: _robust_to_numeric(x).sum()
                                ).reset_index()
                                _grp2b.columns = ["Location", _sel_col]
                                _grp2b = _grp2b.sort_values(_sel_col, ascending=False).reset_index(drop=True)
                                _grp2b.index += 1
                                _auto_loc2 = _grp2b
                            _res_sc = {
                                "type": "scalar", "label": _q_label,
                                "value": _val, "unit": _unit,
                                "column": _sel_col, "col_reason": "direct selection",
                                "row_count": _total_rows,
                                "valid_count": _valid_rows,
                                "operation": _sel_op_name.lower(),
                                "loc_breakdown": None, "auto_loc": _auto_loc2,
                            }
                            st.session_state["sq_results_history"].append({
                                "query": _q_label, "source": sq_src,
                                "records": _total_rows, "results": [_res_sc],
                            })

    run_clicked = False  # Compatibility — old AI path disabled in this mode


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

                # ── Customer lookup result (additive) ─────────────────────
                elif rtype == "customer_lookup":
                    cust_disp = res["customer"]
                    n_rows    = res["row_count"]
                    st.markdown(
                        f'<div style="background:{DARK2};border:2px solid {CYAN};'
                        f'border-radius:14px;padding:20px 24px;margin:12px 0">'
                        f'<div style="font-size:1.1rem;font-weight:900;color:{WHITE};'
                        f'margin-bottom:4px">👤 {cust_disp}</div>'
                        f'<div style="font-size:.8rem;color:{MUTED}">'
                        f'{n_rows} matching row(s) across all locations</div></div>',
                        unsafe_allow_html=True
                    )

                    # Focus metric (the specifically requested value)
                    if res.get("focus_col") and res.get("focus_val") is not None:
                        _fval  = res["focus_val"]
                        _funit = res.get("focus_unit","")
                        _fcol  = res["focus_col"].split("|")[-1].strip() if "|" in res["focus_col"] else res["focus_col"]
                        _fdisplay = (
                            f"₹ {_fval:,.2f}" if _funit == "₹"
                            else f"{_fmt_decimal(_fval)} {_funit}".strip()
                        )
                        st.markdown(
                            f'<div style="background:{CARD};border:1px solid {BORD};'
                            f'border-radius:12px;padding:20px 28px;margin:8px 0;'
                            f'display:inline-block;min-width:260px">'
                            f'<div style="font-size:.75rem;color:{MUTED};font-weight:700;'
                            f'text-transform:uppercase;letter-spacing:.06em">{_fcol}</div>'
                            f'<div class="result-big">{_fdisplay}</div>'
                            f'<div style="font-size:.72rem;color:{CYAN};margin-top:6px">'
                            f'for {cust_disp}</div></div>',
                            unsafe_allow_html=True
                        )

                    # Two-column layout: profile + metrics
                    _pc1, _pc2 = st.columns([1, 2])
                    with _pc1:
                        if not res["profile_df"].empty:
                            st.markdown(
                                f'<div style="font-size:.78rem;color:{CYAN};font-weight:700;'
                                f'text-transform:uppercase;letter-spacing:.06em;'
                                f'margin:10px 0 4px">📋 Profile</div>',
                                unsafe_allow_html=True
                            )
                            _pf = res["profile_df"].copy()
                            _pf.index = [""] * len(_pf)
                            st.dataframe(_pf, use_container_width=True, hide_index=True)

                    with _pc2:
                        if not res["metrics_df"].empty:
                            st.markdown(
                                f'<div style="font-size:.78rem;color:{CYAN};font-weight:700;'
                                f'text-transform:uppercase;letter-spacing:.06em;'
                                f'margin:10px 0 4px">📊 Metrics</div>',
                                unsafe_allow_html=True
                            )
                            _mf = res["metrics_df"].copy()
                            _mf.index = [""] * len(_mf)
                            st.dataframe(_mf, use_container_width=True, hide_index=True)

                    # Full detail table in expander
                    with st.expander(f"🔎 Full detail table — {n_rows} row(s)", expanded=False):
                        st.dataframe(res["raw_df"], use_container_width=True)
                        st.download_button(
                            "⬇ Download CSV",
                            res["raw_df"].to_csv(index=False).encode(),
                            f"customer_{cust_disp.replace(' ','_')[:30]}.csv",
                            "text/csv",
                            key=f"dl_cust_{hash(cust_disp)}",
                        )
                # ── End customer lookup ────────────────────────────────────

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


    # ══════════════════════════════════════════════════════════════════════════
    # COLUMN HEADER LOOKUP — Find any column by exact/partial header name
    # across all 10 DC Excel files and all sheets, with optional aggregations
    # ══════════════════════════════════════════════════════════════════════════

    st.markdown(f"<hr style='border-color:{BORD};margin:32px 0 20px'>",
                unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">🔎 Column Header Lookup — Search Any Column Across All Files & Sheets</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div style="font-size:.85rem;color:{MUTED};margin-bottom:16px">'
        f'Type a column header name (exact or partial) to find and retrieve data from '
        f'<b>all matching columns across all 10 DC Excel files and all sheets</b>. '
        f'Numeric columns support aggregation metrics (Sum, Avg, Min, Max …); '
        f'text columns support Count, Unique Count, and Value Counts.</div>',
        unsafe_allow_html=True,
    )

    _chl_all_cols = sorted(
        [c for c in CUST.columns if not c.startswith("_")]
    ) if not CUST.empty else []

    _chl_r1a, _chl_r1b = st.columns([4, 2])
    with _chl_r1a:
        _chl_header_search = st.text_input(
            "🔍 Column Header Name (partial match, case-insensitive)",
            placeholder="e.g.  Customer Name  |  Subscription  |  Power Capacity  |  MRC  |  Revenue",
            key="chl_header_search",
        )
    with _chl_r1b:
        _chl_loc_sel = st.selectbox(
            "📍 Location Filter",
            ["All Locations"] + sorted(fdata.keys()),
            key="chl_loc_sel",
        )

    _chl_search_lower = _chl_header_search.strip().lower()
    _chl_matched = (
        [c for c in _chl_all_cols if _chl_search_lower in c.lower()]
        if _chl_search_lower else _chl_all_cols
    )

    _chl_r2a, _chl_r2b = st.columns([4, 2])
    with _chl_r2a:
        _chl_selected_col = st.selectbox(
            f"📋 Matched Columns — {len(_chl_matched)} found",
            ["— pick a column —"] + _chl_matched,
            key="chl_selected_col",
        )
    with _chl_r2b:
        _chl_sheet_opts = ["All Sheets"]
        if _chl_loc_sel != "All Locations":
            _chl_sheet_opts += sorted(fdata.get(_chl_loc_sel, {}).keys())
        _chl_sheet_sel = st.selectbox("📄 Sheet Filter", _chl_sheet_opts, key="chl_sheet_sel")

    _chl_is_numeric = False
    if _chl_selected_col != "— pick a column —" and not CUST.empty and _chl_selected_col in CUST.columns:
        _chl_probe = CUST[_chl_selected_col].dropna()
        _chl_is_numeric = (
            pd.to_numeric(_chl_probe, errors="coerce").notna().sum() /
            max(len(_chl_probe), 1) > 0.4
        )

    _chl_r3a, _chl_r3b, _chl_r3c = st.columns([3, 2, 1])
    with _chl_r3a:
        if _chl_is_numeric:
            _chl_metric = st.selectbox(
                "📐 Metric (optional — numeric column)",
                ["— none (show raw data) —", "Sum", "Average", "Min", "Max",
                 "Count", "Median", "Std Dev", "Count Non-Zero"],
                key="chl_metric",
            )
        else:
            _chl_metric = st.selectbox(
                "📐 Operation (optional — text column)",
                ["— none (show raw data) —", "Count (non-null)", "Unique Count", "Value Counts"],
                key="chl_metric",
            )
    with _chl_r3b:
        _chl_max_rows = st.number_input(
            "Max rows to show", min_value=10, max_value=500, value=50, step=10,
            key="chl_max_rows",
        )
    with _chl_r3c:
        st.markdown("<br>", unsafe_allow_html=True)
        _chl_run = st.button("🔎 Lookup", key="chl_run", use_container_width=True)

    _chl_val_filter = st.text_input(
        "🔑 Value Filter — enter any value to filter rows in the selected column (optional, partial match)",
        placeholder="e.g.  Colt  |  100  |  Caged  |  Bangalore  |  MRC",
        key="chl_val_filter",
    )

    if _chl_run:
        if _chl_selected_col == "— pick a column —":
            st.warning("Please search for and select a column header to look up.")
        elif CUST.empty:
            st.warning("No data loaded. Check that the Excel files are accessible.")
        else:
            _chl_df = CUST.copy()
            if _chl_loc_sel != "All Locations" and "_Location" in _chl_df.columns:
                _chl_df = _chl_df[_chl_df["_Location"] == _chl_loc_sel]
            if _chl_sheet_sel != "All Sheets" and "_Sheet" in _chl_df.columns:
                _chl_df = _chl_df[_chl_df["_Sheet"] == _chl_sheet_sel]
            if _chl_val_filter.strip() and _chl_selected_col in _chl_df.columns:
                _chl_mask = (
                    _chl_df[_chl_selected_col]
                    .astype(str)
                    .str.contains(re.escape(_chl_val_filter.strip()), case=False, na=False)
                )
                _chl_df = _chl_df[_chl_mask]

            if _chl_selected_col not in _chl_df.columns:
                st.error(
                    f"Column ‘{_chl_selected_col}’ not found "
                    f"in the selected scope."
                )
            else:
                _chl_col_data = _chl_df[_chl_selected_col].dropna()
                _chl_scope = (
                    _chl_loc_sel if _chl_loc_sel != "All Locations"
                    else "All Locations"
                )
                if _chl_sheet_sel != "All Sheets":
                    _chl_scope += f" / {_chl_sheet_sel}"

                _chl_no_metric = _chl_metric == "— none (show raw data) —"

                if not _chl_no_metric:
                    if _chl_is_numeric:
                        _chl_num = pd.to_numeric(_chl_col_data, errors="coerce").dropna()
                        _chl_num_ops = {
                            "Sum":            _chl_num.sum(),
                            "Average":        _chl_num.mean(),
                            "Min":            _chl_num.min(),
                            "Max":            _chl_num.max(),
                            "Count":          float(len(_chl_num)),
                            "Median":         _chl_num.median(),
                            "Std Dev":        _chl_num.std(ddof=1),
                            "Count Non-Zero": float((_chl_num != 0).sum()),
                        }
                        _chl_result = _chl_num_ops.get(_chl_metric, "N/A")
                        _chl_fmt = (
                            f"{_chl_result:,.2f}"
                            if isinstance(_chl_result, float) else str(_chl_result)
                        )
                        st.markdown(f"""
        <div class="result-box">
          <div style="font-size:.82rem;color:{MUTED};margin-bottom:6px">
            <b>{_chl_metric}</b> of
            <b style="color:{CYAN}">{_chl_selected_col}</b>
            &nbsp;·&nbsp; {_chl_scope}
          </div>
          <div class="result-big">{_chl_fmt}</div>
          <div style="font-size:.78rem;color:{MUTED};margin-top:8px">
            Based on <b style="color:{CYAN}">{len(_chl_num):,}</b> numeric records
          </div>
        </div>""", unsafe_allow_html=True)

                        if _chl_loc_sel == "All Locations" and "_Location" in CUST.columns:
                            _chl_num_op_fn = {
                                "Sum":            lambda s: s.sum(),
                                "Average":        lambda s: s.mean(),
                                "Min":            lambda s: s.min(),
                                "Max":            lambda s: s.max(),
                                "Count":          lambda s: float(len(s)),
                                "Median":         lambda s: s.median(),
                                "Std Dev":        lambda s: s.std(ddof=1),
                                "Count Non-Zero": lambda s: float((s != 0).sum()),
                            }
                            _chl_grp_rows = []
                            for _chl_gloc, _chl_gdf in CUST.groupby("_Location"):
                                _chl_gs = pd.to_numeric(
                                    _chl_gdf.get(_chl_selected_col, pd.Series()),
                                    errors="coerce",
                                ).dropna()
                                if not _chl_gs.empty:
                                    _chl_grp_rows.append({
                                        "Location":     _chl_gloc,
                                        _chl_metric:    _chl_num_op_fn[_chl_metric](_chl_gs),
                                        "Records Used": len(_chl_gs),
                                    })
                            if _chl_grp_rows:
                                _chl_grp_df = (
                                    pd.DataFrame(_chl_grp_rows)
                                    .sort_values(_chl_metric, ascending=False)
                                    .reset_index(drop=True)
                                )
                                st.markdown(
                                    f'<div style="font-size:.85rem;color:{MUTED};'
                                    f'margin:14px 0 6px"><b>Per-Location Breakdown:</b></div>',
                                    unsafe_allow_html=True,
                                )
                                st.dataframe(_chl_grp_df, use_container_width=True)
                    else:
                        if _chl_metric == "Count (non-null)":
                            st.markdown(f"""
        <div class="result-box">
          <div style="font-size:.82rem;color:{MUTED};margin-bottom:6px">
            <b>Count (non-null)</b> in
            <b style="color:{CYAN}">{_chl_selected_col}</b>
            &nbsp;·&nbsp; {_chl_scope}
          </div>
          <div class="result-big">{len(_chl_col_data):,}</div>
        </div>""", unsafe_allow_html=True)
                        elif _chl_metric == "Unique Count":
                            st.markdown(f"""
        <div class="result-box">
          <div style="font-size:.82rem;color:{MUTED};margin-bottom:6px">
            <b>Unique Count</b> in
            <b style="color:{CYAN}">{_chl_selected_col}</b>
            &nbsp;·&nbsp; {_chl_scope}
          </div>
          <div class="result-big">{_chl_col_data.nunique():,}</div>
        </div>""", unsafe_allow_html=True)
                        elif _chl_metric == "Value Counts":
                            _chl_vc = (
                                _chl_col_data.astype(str)
                                .value_counts()
                                .reset_index()
                            )
                            _chl_vc.columns = [_chl_selected_col, "Count"]
                            st.markdown(
                                f'<div style="font-size:.85rem;color:{MUTED};margin:8px 0 4px">'
                                f'<b>Value Counts — {_chl_selected_col}</b>'
                                f' · {_chl_scope}</div>',
                                unsafe_allow_html=True,
                            )
                            st.dataframe(
                                _chl_vc.head(int(_chl_max_rows)),
                                use_container_width=True,
                            )

                _chl_show_cols = [
                    c for c in ["_Location", "_Sheet", _chl_selected_col]
                    if c in _chl_df.columns
                ]
                _chl_display_df = (
                    _chl_df[_chl_show_cols]
                    .dropna(subset=[_chl_selected_col])
                    .reset_index(drop=True)
                )
                st.markdown(
                    f'<div style="font-size:.85rem;color:{MUTED};margin:12px 0 6px">'
                    f'<b>Raw Data</b> — '
                    f'<b style="color:{CYAN}">{_chl_selected_col}</b> '
                    f'({min(int(_chl_max_rows), len(_chl_display_df))} of '
                    f'{len(_chl_display_df)} rows · {_chl_scope})</div>',
                    unsafe_allow_html=True,
                )
                st.dataframe(
                    _chl_display_df.head(int(_chl_max_rows)),
                    use_container_width=True,
                )
                st.download_button(
                    "⬇ Download Column Data CSV",
                    _chl_display_df.to_csv(index=False).encode("utf-8"),
                    f"col_{'_'.join(_chl_selected_col[:25].split()).replace('|','_')}.csv",
                    "text/csv",
                    key="chl_dl",
                )




    # ══════════════════════════════════════════════════════════════════════════
    # FETCH ANY CELL VALUE — by VALUE (not by position)
    # Search for a specific value across ALL 10 DC Excel files and ALL sheets.
    # Returns every cell (file, sheet, row, column) that contains the value.
    # ══════════════════════════════════════════════════════════════════════════

    st.markdown(f"<hr style='border-color:{BORD};margin:32px 0 20px'>",
                unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">📌 Fetch Any Cell Value — By Value Across All Files & Sheets</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div style="font-size:.85rem;color:{MUTED};margin-bottom:16px">'
        f'Enter a value (text, number, date, or partial string) to find every cell that '
        f'contains it across <b>all 10 DC Excel files and all sheets</b>. '
        f'Results show the exact file, sheet, row and column for each match — '
        f'no need to know the row position in advance.</div>',
        unsafe_allow_html=True,
    )

    _fcv_vc1, _fcv_vc2 = st.columns([4, 1])
    with _fcv_vc1:
        _fcv_val_input = st.text_input(
            "🔍 Value to search (partial match, case-insensitive)",
            placeholder="e.g.  530.031   |   YES BANK   |   2024-03-15   |   CAGED   |   Rated",
            key="fcv_val_input",
        )
    with _fcv_vc2:
        st.markdown("<div style='margin-top:28px'></div>", unsafe_allow_html=True)
        _fcv_val_run = st.button("🔍 Find Value", key="fcv_val_run", use_container_width=True)

    # Optional filters: restrict to a specific file or sheet
    _fcv_opt_c1, _fcv_opt_c2 = st.columns([2, 2])
    with _fcv_opt_c1:
        _fcv_filter_locs = st.multiselect(
            "🏢 Limit to DC file(s) (leave empty = ALL)",
            options=sorted(ALL.keys()),
            default=[],
            key="fcv_filter_locs",
        )
    with _fcv_opt_c2:
        _fcv_opt_sheets_all: list = []
        _for_locs = _fcv_filter_locs if _fcv_filter_locs else sorted(ALL.keys())
        for _fl in _for_locs:
            _fcv_opt_sheets_all += sorted(ALL.get(_fl, {}).keys())
        _fcv_opt_sheets_all = sorted(set(_fcv_opt_sheets_all))
        _fcv_filter_sheets = st.multiselect(
            "📄 Limit to sheet(s) (leave empty = ALL)",
            options=_fcv_opt_sheets_all,
            default=[],
            key="fcv_filter_sheets",
        )

    if _fcv_val_run:
        if not _fcv_val_input.strip():
            st.warning("Please enter a value to search.")
        else:
            _fcv_search_term = _fcv_val_input.strip().lower()
            _fcv_hit_rows: list = []   # list of dicts: {File, Sheet, Row, Column, Value}

            # Iterate over all files and sheets
            _fcv_search_locs = _fcv_filter_locs if _fcv_filter_locs else sorted(ALL.keys())
            for _fcv_loc in _fcv_search_locs:
                _fcv_loc_sheets = ALL.get(_fcv_loc, {})
                _fcv_search_sheets = (
                    [s for s in _fcv_filter_sheets if s in _fcv_loc_sheets]
                    if _fcv_filter_sheets else sorted(_fcv_loc_sheets.keys())
                )
                for _fcv_sh in _fcv_search_sheets:
                    _fcv_df = _fcv_loc_sheets.get(_fcv_sh)
                    if _fcv_df is None or _fcv_df.empty:
                        continue
                    _fcv_data_cols = [c for c in _fcv_df.columns if not c.startswith("_")]
                    for _fcv_col in _fcv_data_cols:
                        # Vectorised substring search across the column
                        _fcv_col_str = _fcv_df[_fcv_col].astype(str).str.lower()
                        _fcv_match_mask = _fcv_col_str.str.contains(
                            re.escape(_fcv_search_term), na=False
                        )
                        for _fcv_ridx in _fcv_df.index[_fcv_match_mask]:
                            _fcv_raw_val = _fcv_df.at[_fcv_ridx, _fcv_col]
                            _fcv_hit_rows.append({
                                "File (Location)": _fcv_loc,
                                "Sheet": _fcv_sh,
                                "Row": int(_fcv_ridx) + 1,
                                "Column": _fcv_col,
                                "Cell Value": str(_fcv_raw_val),
                            })

            if not _fcv_hit_rows:
                st.warning(
                    f"**'{_fcv_val_input}'** not found in any cell across "
                    f"{'selected' if (_fcv_filter_locs or _fcv_filter_sheets) else 'all'} "
                    f"Excel files and sheets. Try a shorter or partial value."
                )
            else:
                _fcv_hits_df = pd.DataFrame(_fcv_hit_rows)
                _fcv_n_hits  = len(_fcv_hits_df)
                _fcv_n_files = _fcv_hits_df["File (Location)"].nunique()
                _fcv_n_sheets = _fcv_hits_df["Sheet"].nunique()
                _fcv_n_cols  = _fcv_hits_df["Column"].nunique()

                st.markdown(
                    f'<div style="background:{DARK2};border:2px solid {CYAN};'
                    f'border-radius:14px;padding:20px 26px;margin:14px 0">'
                    f'<div style="font-size:1.1rem;font-weight:900;color:{WHITE};margin-bottom:6px">'
                    f'📌 Found <span style="color:{CYAN}">{_fcv_n_hits}</span> cell(s) '
                    f'matching <i>"{_fcv_val_input}"</i></div>'
                    f'<div style="font-size:.82rem;color:{TEXT}">'
                    f'Across <b style="color:{CYAN}">{_fcv_n_files}</b> DC file(s) · '
                    f'<b style="color:{CYAN}">{_fcv_n_sheets}</b> sheet(s) · '
                    f'<b style="color:{CYAN}">{_fcv_n_cols}</b> unique column(s)</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

                # Summary: count per file + sheet
                with st.expander("🔬 Validation — Match count per file & sheet", expanded=False):
                    _fcv_val_sum = (
                        _fcv_hits_df.groupby(["File (Location)", "Sheet"])
                        .size().reset_index(name="Matches")
                        .sort_values(["File (Location)", "Sheet"])
                    )
                    _fcv_val_sum.index = range(1, len(_fcv_val_sum) + 1)
                    st.dataframe(_fcv_val_sum, use_container_width=True)

                # Full results table
                st.markdown(
                    f'<div style="font-size:.8rem;color:{CYAN};font-weight:700;'
                    f'text-transform:uppercase;letter-spacing:.05em;margin:16px 0 6px">'
                    f'📋 All Matching Cells ({_fcv_n_hits})</div>',
                    unsafe_allow_html=True,
                )
                _fcv_hits_df.index = range(1, len(_fcv_hits_df) + 1)
                st.dataframe(_fcv_hits_df, use_container_width=True)

                # Full row context for first N matches
                with st.expander(
                    f"🔎 Full row context for first {min(5, _fcv_n_hits)} match(es)",
                    expanded=False,
                ):
                    for _fi, _frow in enumerate(_fcv_hit_rows[:5]):
                        _fctx_loc   = _frow["File (Location)"]
                        _fctx_sh    = _frow["Sheet"]
                        _fctx_ridx  = _frow["Row"] - 1   # 0-based
                        _fctx_col   = _frow["Column"]
                        _fctx_df    = ALL.get(_fctx_loc, {}).get(_fctx_sh)
                        if _fctx_df is None or _fctx_ridx >= len(_fctx_df):
                            continue
                        _fctx_data_cols = [c for c in _fctx_df.columns if not c.startswith("_")]
                        _fctx_row_df = _fctx_df.iloc[[_fctx_ridx]][_fctx_data_cols].T.reset_index()
                        _fctx_row_df.columns = ["Column", "Value"]
                        _fctx_row_df.index   = range(1, len(_fctx_row_df) + 1)
                        st.markdown(
                            f'**Match {_fi+1}** — '
                            f'`{_fctx_loc}` › `{_fctx_sh}` › Row {_frow["Row"]} › '
                            f'`{_fctx_col}`'
                        )
                        st.dataframe(_fctx_row_df, use_container_width=True)

                st.download_button(
                    f"⬇️ Download all {_fcv_n_hits} matching cells (CSV)",
                    _fcv_hits_df.to_csv(index=False).encode("utf-8"),
                    f"cell_search_{_fcv_val_input.replace(' ','_')[:30]}.csv",
                    "text/csv",
                    key="fcv_dl_hits",
                )
    # ══════════════════════════════════════════════════════════════════════════
    # SHAMBHUSHIV — ENHANCED CUSTOMER LOOKUP
    # • Searches ALL 10 DC Excel files and ALL sheets regardless of the
    #   "Query data source" dropdown above (which only affects Smart Query).
    # • _sk_full_pool is built fresh from ALL (the global dict of all loaded
    #   Excel files) so every location/sheet is always searchable.
    # • Location dropdown shows ALL locations from all 10 Excel files.
    # • Sheet dropdown is dynamically populated from selected locations.
    # • Results are shown location-wise AND sheet-wise with per-group metrics.
    # • "Not found" summary shown for locations where customer is absent.
    # ══════════════════════════════════════════════════════════════════════════

    # ── Build the full search pool directly from ALL — all 10 Excel files,
    #    all sheets — independent of sq_src and the sidebar sel_locs/sel_sheets.
    #    We use combined_df(ALL) here (no preprocessing) so that every location
    #    and sheet from attached Excel files is always represented.
    _sk_full_pool = combined_df(ALL)   # has _Location and _Sheet columns

    st.markdown(f"<hr style='border-color:{BORD};margin:32px 0 20px'>",
                unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">👤 Customer Lookup — Name-Based Data Retrieval</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div style="font-size:.85rem;color:{MUTED};margin-bottom:16px">'
        f'Search any customer by name across <b>all 10 DC Excel files and all sheets</b>. '
        f'Works even when the same customer appears in multiple locations. '
        f'Type part of the name — partial match, case-insensitive '
        f'(e.g. <b style="color:{CYAN}">Wipro</b> finds WIPRO LIMITED, Wipro Nabard, etc.). '
        f'Use the filters below to narrow to a specific DC file or sheet.</div>',
        unsafe_allow_html=True,
    )

    # ── Location + Sheet filter options — taken directly from ALL dict
    #    so ALL 10 Excel files and ALL their sheets are always listed,
    #    regardless of whether any rows survived preprocessing.
    _sk_all_locs = sorted(ALL.keys())
    _sk_loc_sheet_map: dict = {
        _sk_l: sorted(ALL[_sk_l].keys()) for _sk_l in _sk_all_locs
    }

    _sk_fr1, _sk_fr2 = st.columns([3, 2])
    with _sk_fr1:
        _sk_sel_locs = st.multiselect(
            "🏢 Filter by DC Location (leave empty = ALL)",
            options=_sk_all_locs,
            default=[],
            key="sk_sel_locs",
        )
    with _sk_fr2:
        _sk_sheet_options = []
        if _sk_sel_locs:
            for _l in _sk_sel_locs:
                _sk_sheet_options += _sk_loc_sheet_map.get(_l, [])
            _sk_sheet_options = sorted(set(_sk_sheet_options))
        else:
            for _l in _sk_all_locs:
                _sk_sheet_options += _sk_loc_sheet_map.get(_l, [])
            _sk_sheet_options = sorted(set(_sk_sheet_options))
        _sk_sel_sheets = st.multiselect(
            "📄 Filter by Sheet (leave empty = ALL sheets)",
            options=_sk_sheet_options,
            default=[],
            key="sk_sel_sheets",
        )

    # ── Customer name + metric + button ───────────────────────────────────────
    _sk_r1, _sk_r2, _sk_r3 = st.columns([2, 2, 1])
    with _sk_r1:
        _sk_cust_input = st.text_input(
            "🔍 Customer Name (partial match)",
            placeholder="e.g. Wipro, Oracle, Cisco, TATA, YES BANK …",
            key="sk_cust_input",
        )
    with _sk_r2:
        _sk_field_input = st.text_input(
            "📐 Metric to highlight (optional)",
            placeholder="e.g. power capacity purchased, total revenue …",
            key="sk_field_input",
        )
    with _sk_r3:
        st.markdown("<div style='margin-top:28px'></div>", unsafe_allow_html=True)
        _sk_run = st.button("Find Customer", key="sk_run", use_container_width=True)

    # ── Live autocomplete: show matching names as typed ────────────────────────
    # Uses _sk_full_pool (all 10 Excel files) so autocomplete is always complete
    if _sk_cust_input.strip() and not _sk_run:
        _sk_hint_rows = _sk_find_customers_all(_sk_cust_input.strip(), _sk_full_pool)
        _sk_hint_cols = _sk_all_customer_cols(_sk_hint_rows)
        _sk_hint_names: list = []
        for _hc in _sk_hint_cols:
            if _hc in _sk_hint_rows.columns:
                vals = (_sk_hint_rows[_hc].dropna().astype(str)
                        .str.strip().unique().tolist())
                _sk_hint_names += [
                    v for v in vals
                    if v and v.lower() not in ("none","nan","")
                ]
        _sk_hint_names = sorted(set(_sk_hint_names))
        if _sk_hint_names:
            st.markdown(
                f'<div style="font-size:.77rem;color:{CYAN};font-weight:700;'
                f'margin:4px 0 2px">Matching customers ({len(_sk_hint_names)}) '
                f'across all files:</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                " &nbsp; ".join(
                    f'<span class="badge">{n[:55]}</span>'
                    for n in _sk_hint_names[:25]
                ),
                unsafe_allow_html=True,
            )

    # ── Execute lookup ─────────────────────────────────────────────────────────
    if _sk_run:
        if not _sk_cust_input.strip():
            st.warning("Please enter a customer name.")
        elif _sk_full_pool.empty:
            st.error("No data loaded. Please check your Excel files.")
        else:
            # Build the search pool: start with _sk_full_pool (all 10 Excel files)
            # so that customer lookup is ALWAYS across all locations/sheets,
            # independent of the sidebar or the sq_src dropdown above.
            _sk_pool = _sk_full_pool.copy()
            # Apply own location filter (if user selected specific locations)
            if _sk_sel_locs and "_Location" in _sk_pool.columns:
                _sk_pool = _sk_pool[_sk_pool["_Location"].isin(_sk_sel_locs)]
            # Apply own sheet filter (if user selected specific sheets)
            if _sk_sel_sheets and "_Sheet" in _sk_pool.columns:
                _sk_pool = _sk_pool[_sk_pool["_Sheet"].isin(_sk_sel_sheets)]

            # Search ALL customer-name columns across the pool
            _sk_rows = _sk_find_customers_all(_sk_cust_input.strip(), _sk_pool)

            if _sk_rows.empty:
                st.warning(
                    f"**'{_sk_cust_input}'** not found in "
                    f"{'selected locations/sheets' if (_sk_sel_locs or _sk_sel_sheets) else 'any DC file'}. "
                    f"Try a shorter or different name."
                )
            else:
                # Canonical display name from actual data
                _sk_cust_disp = _sk_canonical_name(_sk_rows, _sk_cust_input.strip())
                _sk_n_rows    = len(_sk_rows)

                # Build (location, sheet) groups
                _sk_has_loc   = "_Location" in _sk_rows.columns
                _sk_has_sheet = "_Sheet"    in _sk_rows.columns

                if _sk_has_loc and _sk_has_sheet:
                    _sk_groups = list(_sk_rows.groupby(
                        ["_Location", "_Sheet"], sort=True
                    ))
                elif _sk_has_loc:
                    _sk_groups = [
                        ((loc, ""), g)
                        for loc, g in _sk_rows.groupby("_Location", sort=True)
                    ]
                else:
                    _sk_groups = [("(All)", _sk_rows)]

                _sk_found_locs   = sorted({k[0] for k, _ in _sk_groups})
                _sk_found_sheets = sorted({k[1] for k, _ in _sk_groups if k[1]})
                _sk_n_files      = len(_sk_found_locs)
                _sk_n_sheets_    = len(_sk_groups)

                # ── Top summary card ──────────────────────────────────────
                st.markdown(
                    f'<div style="background:{DARK2};border:2px solid {CYAN};'
                    f'border-radius:14px;padding:20px 26px;margin:14px 0">'
                    f'<div style="font-size:1.1rem;font-weight:900;color:{WHITE};'
                    f'margin-bottom:6px">👤 {_sk_cust_disp}</div>'
                    f'<div style="font-size:.82rem;color:{TEXT}">'
                    f'<b style="color:{CYAN}">{_sk_n_rows}</b> matching row(s) '
                    f'found in <b style="color:{CYAN}">{_sk_n_files}</b> DC location(s) '
                    f'across <b style="color:{CYAN}">{_sk_n_sheets_}</b> sheet(s)</div>'
                    f'<div style="margin-top:10px">'
                    + " &nbsp; ".join(
                        f'<span class="badge">{k[0]}'
                        + (f' — {k[1]}' if k[1] else '')
                        + '</span>'
                        for k, _ in _sk_groups
                    )
                    + f'</div></div>',
                    unsafe_allow_html=True,
                )

                # ── Focus metric (overall, across all matched rows) ───────
                if _sk_field_input.strip():
                    _sk_focus_col, _ = _sq_resolve_field(_sk_rows, _sk_field_input.strip())
                    if _sk_focus_col and _sk_focus_col in _sk_rows.columns:
                        _sk_fs  = _robust_to_numeric(_sk_rows[_sk_focus_col]).dropna()
                        if not _sk_fs.empty:
                            _sk_fv   = float(_sk_fs.sum())
                            _sk_fu   = _detect_unit(_sk_focus_col)
                            _sk_fd   = (f"₹ {_sk_fv:,.2f}" if _sk_fu == "₹"
                                        else f"{_fmt_decimal(_sk_fv)} {_sk_fu}".strip())
                            _sk_fcol = (_sk_focus_col.split("|")[-1].strip()
                                        if "|" in _sk_focus_col else _sk_focus_col)
                            st.markdown(
                                f'<div style="background:{CARD};border:1px solid {BORD};'
                                f'border-radius:12px;padding:22px 32px;margin:10px 0;'
                                f'display:inline-block;min-width:280px">'
                                f'<div style="font-size:.75rem;color:{MUTED};font-weight:700;'
                                f'text-transform:uppercase;letter-spacing:.06em">'
                                f'{_sk_fcol} (all locations combined)</div>'
                                f'<div class="result-big">{_sk_fd}</div>'
                                f'<div style="font-size:.72rem;color:{CYAN};margin-top:6px">'
                                f'for {_sk_cust_disp}</div></div>',
                                unsafe_allow_html=True,
                            )

                # ── Overall download ──────────────────────────────────────
                _sk_dl_meta  = [c for c in ["_Location","_Sheet"] if c in _sk_rows.columns]
                _sk_dl_data  = [c for c in _sk_rows.columns if not c.startswith("_")]
                _sk_dl_df    = _sk_rows[_sk_dl_meta + _sk_dl_data].copy()
                _sk_dl_df.index = range(1, len(_sk_dl_df) + 1)
                with st.expander(
                    f"⬇ Download all {_sk_n_rows} matching rows (all locations combined)",
                    expanded=False,
                ):
                    st.dataframe(_sk_dl_df, use_container_width=True)
                    st.download_button(
                        "⬇ Download combined CSV",
                        _sk_dl_df.to_csv(index=False).encode(),
                        f"customer_{_sk_cust_disp.replace(' ','_')[:40]}_ALL.csv",
                        "text/csv",
                        key="sk_dl_all",
                    )

                # ── Per-location / per-sheet cards ────────────────────────
                st.markdown(
                    f"<hr style='border-color:{BORD};margin:24px 0 16px'>",
                    unsafe_allow_html=True,
                )
                st.markdown(
                    '<div class="section-title">'
                    '📂 Results by DC Location &amp; Sheet'
                    '</div>',
                    unsafe_allow_html=True,
                )

                for _sk_gi, (_sk_gkey, _sk_gdf) in enumerate(_sk_groups):
                    _sk_loc_n   = _sk_gkey[0] if isinstance(_sk_gkey, tuple) else str(_sk_gkey)
                    _sk_sheet_n = _sk_gkey[1] if isinstance(_sk_gkey, tuple) else ""
                    _sk_gdf     = _sk_gdf.reset_index(drop=True)
                    _sk_g_n     = len(_sk_gdf)
                    _sk_g_name  = _sk_canonical_name(_sk_gdf, _sk_cust_disp)

                    # Location card header
                    st.markdown(
                        f'<div style="background:{DARK2};border-left:4px solid {BLUE};'
                        f'border-radius:0 12px 12px 0;padding:14px 20px;margin:18px 0 8px">'
                        f'<span style="font-size:1rem;font-weight:800;color:{WHITE}">'
                        f'🏢 {_sk_loc_n}</span>'
                        + (f'&nbsp;<span style="font-size:.82rem;color:{CYAN};'
                           f'font-weight:600"> — {_sk_sheet_n}</span>'
                           if _sk_sheet_n else "")
                        + f'<br><span style="font-size:.76rem;color:{MUTED}">'
                        f'{_sk_g_n} row(s) &nbsp;·&nbsp; '
                        f'<span style="color:{GREEN}">{_sk_g_name}</span>'
                        f'</span></div>',
                        unsafe_allow_html=True,
                    )

                    # Per-group metrics & profile
                    _sk_gmetrics = _sk_build_per_loc_metrics(_sk_gdf)
                    _sk_gprofile = _sk_build_per_loc_profile(_sk_gdf)

                    _sk_ga, _sk_gb = st.columns([1, 2])
                    with _sk_ga:
                        if not _sk_gprofile.empty:
                            st.markdown(
                                f'<div style="font-size:.72rem;color:{CYAN};font-weight:700;'
                                f'text-transform:uppercase;margin-bottom:4px">'
                                f'📋 Profile</div>',
                                unsafe_allow_html=True,
                            )
                            _gpf = _sk_gprofile.copy(); _gpf.index = [""] * len(_gpf)
                            st.dataframe(_gpf, use_container_width=True, hide_index=True)

                    with _sk_gb:
                        if not _sk_gmetrics.empty:
                            st.markdown(
                                f'<div style="font-size:.72rem;color:{CYAN};font-weight:700;'
                                f'text-transform:uppercase;margin-bottom:4px">'
                                f'📊 Metrics</div>',
                                unsafe_allow_html=True,
                            )
                            _gmf = _sk_gmetrics.copy(); _gmf.index = [""] * len(_gmf)
                            st.dataframe(_gmf, use_container_width=True, hide_index=True)
                        else:
                            st.caption("No numeric metrics in this sheet for this customer.")

                    # Row detail expander for this group
                    _sk_g_meta = [c for c in ["_Location","_Sheet"] if c in _sk_gdf.columns]
                    _sk_g_data = [c for c in _sk_gdf.columns if not c.startswith("_")]
                    _sk_g_show = _sk_gdf[_sk_g_meta + _sk_g_data].copy()
                    _sk_g_show.index = range(1, len(_sk_g_show) + 1)
                    _sk_g_title = (
                        f"All columns — {_sk_loc_n}"
                        + (f" / {_sk_sheet_n}" if _sk_sheet_n else "")
                        + f"  ({_sk_g_n} row(s))"
                    )
                    with st.expander(_sk_g_title, expanded=(_sk_n_files == 1)):
                        st.dataframe(_sk_g_show, use_container_width=True)
                        _sk_g_fn = (
                            f"customer_{_sk_cust_disp.replace(' ','_')[:25]}"
                            f"_{_sk_loc_n.replace(' ','_')[:20]}.csv"
                        )
                        st.download_button(
                            "⬇ Download this sheet's CSV",
                            _sk_g_show.to_csv(index=False).encode(),
                            _sk_g_fn, "text/csv",
                            key=f"sk_dl_{_sk_gi}",
                        )

                # ── "Not found" summary for remaining locations ────────────
                _sk_absent = [
                    loc for loc in _sk_all_locs
                    if loc not in _sk_found_locs
                    and (not _sk_sel_locs or loc in _sk_sel_locs)
                ]
                if _sk_absent:
                    st.markdown(
                        f'<div style="background:{DARK2};border:1px dashed {BORD};'
                        f'border-radius:10px;padding:12px 18px;margin:20px 0;'
                        f'font-size:.8rem;color:{MUTED}">'
                        f'<b>Not found in:</b> '
                        + ", ".join(
                            f'<span style="color:{AMBER}">{l}</span>'
                            for l in _sk_absent
                        )
                        + "</div>",
                        unsafe_allow_html=True,
                    )


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
                val, *_ = run_op(loc_df, xl_col, xl_op)
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
