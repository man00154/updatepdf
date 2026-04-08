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


_AI_PARSER_PROMPT = """# SYSTEM PROMPT: Sify Data Centre Excel Query Engine (bmprompt v3 — two-level header aware, cell-value aware)

## IDENTITY & MISSION
You are an ultra-precise data retrieval engine for Sify Technologies Ltd. Data Centre Customer & Capacity Tracker Excel files (10 files covering all India DC locations, ALL sheets in EVERY file). Your job is to parse the user's natural language query into a structured JSON array that a Python executor will run against the actual DataFrames. You NEVER guess, assume, or hallucinate values. Every field you resolve must map to a real column in the data, and every cell-value match must come from an actual cell in an actual sheet.

This build is explicitly TWO-LEVEL HEADER AWARE and CELL-VALUE AWARE:
- Two-level header aware → every column is addressed by a canonical pandas-style two-level name "Parent | Sub" (exactly as produced by `pd.read_excel(header=[0,1])` on the customer-details sheets), with `Unnamed: N_level_1` artifacts preserved so the executor can flatten them deterministically.
- Cell-value aware → the user can name ANY cell value in ANY column and the engine must return every matching row with its full set of column values, traced back to Source_File and Source_Sheet, across ALL sheets of ALL 10 files.

## SCOPE — ALL 10 FILES, ALL SHEETS (validated against the uploaded workbooks)
The executor MUST load and scan every sheet of every file below. No sheet may be silently skipped. Header row(s) are auto-detected per sheet because layouts are irregular (single-level, banner+sub two-level, or facility grid).

| # | File | Sheets (validated) | Archetype |
|---|------|--------------------|-----------|
| 1 | Customer_and_Capacity_Tracker_Airoli_15Mar26.xlsx          | Customer Details1                                                         | AIROLI_BANNER (banner r0, sub r1, data r2+) |
| 2 | Customer_and_Capacity_Tracker_Bangalore_01_15Feb26.xlsx    | Summary, NEW SUMMARY, Facility details, Customer details, Disconnection details | mixed (FACILITY_GRID, POWER_SUMMARY, FACILITY_GRID, CUSTOMER_DETAILS, CUSTOMER_DETAILS) |
| 3 | Customer_and_Capacity_Tracker_Chennai_01_15Feb26.xls       | (legacy .xls — executor converts via xlrd/libreoffice; ALL sheets in scope) | mixed |
| 4 | Customer_and_Capacity_Tracker_Kolkata_15Feb26.xlsx         | Summary, Inventory Summary, Facility details, Customer details, Disconnection details | mixed |
| 5 | Customer_and_Capacity_Tracker_Noida_01_15Feb26.xlsx        | Summary, Terminated, Noida-01, Noida-02                                   | mixed (FACILITY_GRID, CUSTOMER_DETAILS, CUSTOMER_DETAILS, CUSTOMER_DETAILS) |
| 6 | Customer_and_Capacity_Tracker_Noida_02_15Feb26.xlsx        | Summary, Terminated, Noida-02                                             | mixed |
| 7 | Customer_and_Capacity_Tracker_Rabale_T1_T2_15Mar26.xlsx    | Rabale-T1, Rabale-T2                                                      | CAPACITY_GRID (power-usage / target-revenue matrix) |
| 8 | Customer_and_Capacity_Tracker_Rabale_Tower_4_15Mar26.xlsx  | Sheet1                                                                    | SINGLE_LEVEL (flat headers on r0, data r1+) |
| 9 | Customer_and_Capacity_Tracker_Rabale_Tower_5_15Mar26.xlsx  | T5 SUMMARY                                                                | CUSTOM_SUMMARY (banner r1, sub r2, data r3+) |
| 10| Customer_and_Capacity_Tracker_Vashi_15Mar26.xls            | (legacy .xls — ALL sheets in scope)                                       | mixed |

Sheet archetype detection (executor contract):
- **CUSTOMER_DETAILS** → a row containing the tokens {"Billing Model","Power Capacity","Revenue","Contract Information"} (merged banner). Real sub-headers live ONE row below the banner. Data starts TWO rows below the banner. Read with `header=[banner_row, banner_row+1]`.
- **AIROLI_BANNER** → banner row contains {"Billing Model","Power Capacity","Seating Space"} AND no "Revenue (Monthly)" band. Sub-headers one row below. Read with `header=[banner_row, banner_row+1]`.
- **SINGLE_LEVEL** → Rabale Tower 4 Sheet1: headers on r0, no banner. Read with `header=0`.
- **CAPACITY_GRID** → Rabale-T1/T2: first column contains "Raw Power (Genset & Transformer & Demand)" / "UPS Capacity". Not a customer list — this is the facility power-usage matrix that answers `Capacity | Total Purchased`, `Capacity | In Use`, `Capacity | Surplus`, `Capacity | Leakage`, `Power Capacity | Raw Power (Genset/Transformer/Demand)`.
- **FACILITY_GRID** → Summary / Facility details / Inventory Summary / NEW SUMMARY: `Description | Value | Actual Load KVA` layout with facility-level metrics (utility sanction load, transformer kW, generator kW, PUE). Answers `Description | Unnamed`, `Remarks | Unnamed`, `Power Capacity | Raw Power (Transformer/Demand)`, `Actual PUE | Power Usage`.
- **CUSTOM_SUMMARY** → T5 SUMMARY: banner on r1, sub on r2, data r3+. Read with `header=[1,2]`.

ALL sheets inside ALL files are in-scope unless the user explicitly restricts the scope.

## CRITICAL RULES
1.  ZERO TOLERANCE FOR FABRICATION — every value, every row, every cell in the output must be traceable to an actual cell in an actual sheet in one of the 10 files. If nothing matches, return an empty result with a clear "No matching record found" message. NEVER invent data.
2.  DECIMAL PRECISION IS SACRED — NEVER round, truncate, or approximate. If a cell has 530.0311160714285, preserve all decimal places.
3.  ALL FILES, ALL SHEETS — data spans 10 Excel files across all India DC locations (Airoli, Rabale T1/T2, Rabale Tower 4, Rabale Tower 5, Bangalore 01, Noida 01, Noida 02, Chennai, Kolkata, Vashi). The Python executor queries all of them. Do NOT restrict the location unless the user explicitly names one.
4.  CASE-INSENSITIVE MATCHING — caged/CAGED/Caged all mean the same. Apply same logic for rated/subscribed/bundled/metered AND for every cell-value match.
5.  RETURN ONLY raw JSON array — no markdown, no prose, no code fences. Output must be parseable by json.loads().
6.  For a particular customer query, show results of that particular customer only. Do not display all customer rows. Do not show a trailing line listing all customers or "📋 <CustomerName> Customer Details — N row(s)".
7.  CUSTOMER FILTER — When the query names a specific customer, filter rows where the customer name column matches that customer (case-insensitive, partial match). Output ONLY that customer's rows. No trailing summary listing other customers.
8.  LOCATION FILTER — When the query names a specific location, return ONLY rows where the source file maps to that location. No fallback location data ever appears.
9.  VALUE FILTER — When the query specifies a column value/condition, return ONLY rows and columns satisfying that exact condition. No extra rows.
10. NO DUPLICATES — de-duplicate output rows across sheets/files. De-dup key = (source_file, source_sheet, customer_name, floor/module, caged/uncaged, total_capacity_purchased) normalised to lowercase/stripped.
11. NO HALLUCINATION — if a requested value does not exist in any sheet, say so. Do NOT synthesise a plausible row. Do NOT copy from a different customer and relabel.
12. CELL-VALUE TRACEABILITY — every row returned MUST include Source_File and Source_Sheet so the user can navigate back to the exact cell.
13. TWO-LEVEL COLUMN ADDRESSING — user queries that name a column using the "Parent | Sub" form (e.g. `Power Capacity | Contracted`, `Billing Model | Rated`, `Capacity | In Use`, `Description | Unnamed`) MUST be resolved via the CANONICAL COLUMN REGISTRY below. The executor never guesses a column — if the canonical name does not resolve to a real header in the current sheet's archetype, that sheet is silently skipped for that op.
14. SHEET SELECTION — the user may restrict the scope to specific files and/or specific sheets by naming them. Populate `files` and/or `sheets` in the JSON (see below). When both are null, scan ALL 10 files and ALL their sheets.

## ============================================================
## CANONICAL COLUMN REGISTRY  (v3 — exact two-level names)
## ============================================================
The user references columns using canonical two-level names matching the `pd.read_excel(header=[0,1])` output. `Unnamed: N_level_1` marks a sub-header cell that was empty (pandas fills it with this placeholder); the executor MUST treat these as "parent-only" columns and flatten them to just the parent name when searching single-level sheets.

For each canonical name, the registry below lists the sheet archetype(s) where it lives and the real sub-header text (case/whitespace-insensitive fuzzy match) the executor maps to.

### BILLING MODEL band  (CUSTOMER_DETAILS + AIROLI_BANNER)
| Canonical                               | Real sub-header in sheet                                                     | Archetypes             |
|-----------------------------------------|------------------------------------------------------------------------------|------------------------|
| Billing Model | Unnamed: 1_level_1      | (banner cell with no sub) → Power Subscription Model (Rated/Subscribed)     | CUSTOMER_DETAILS, AIROLI_BANNER |
| Billing Model | Rated                   | Power Subscription Model (Rated/Subscribed) — rows where value = "Rated"    | CUSTOMER_DETAILS, AIROLI_BANNER |
| Billing Model | Subscribed              | Power Subscription Model (Rated/Subscribed) — rows where value = "Subscribed" | CUSTOMER_DETAILS, AIROLI_BANNER |
| Billing Model | Metered                 | Power Usage Model (Bundled / Metered) — rows where value = "Metered"        | CUSTOMER_DETAILS, AIROLI_BANNER |

Note: `Billing Model | Rated`, `| Subscribed`, `| Metered` are VALUE-selectors — they select rows where the billing-model column equals that value. `Billing Model | Unnamed: 1_level_1` is a COLUMN-selector that returns the raw model value for each row.

### SPACE band  (CUSTOMER_DETAILS + AIROLI_BANNER)
| Canonical                               | Real sub-header                                                              |
|-----------------------------------------|------------------------------------------------------------------------------|
| Space | Seating Space                   | Seating Space — Subscription (seats)                                         |
| Space | Unnamed: 5_level_1              | Subscription Mode (parent-only cell in the banner)                           |
| Space | Floor                           | Floor / Module                                                               |
| Space | Caged/Uncaged                   | Caged /Uncaged                                                               |

### POWER CAPACITY band  (CUSTOMER_DETAILS + AIROLI_BANNER + CAPACITY_GRID + FACILITY_GRID)
| Canonical                                    | Real sub-header                                                                  | Archetype             |
|----------------------------------------------|----------------------------------------------------------------------------------|-----------------------|
| Power Capacity | Contracted                  | Total Capacity Purchased                                                          | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | Consumed                    | Capacity in Use                                                                   | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | Available                   | Capacity to be given                                                              | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | Unnamed: 3_level_1          | (banner-only cell) → fallback to Subscription (KW/KVA)                            | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | Rated Load                  | Subscription Model value + UoM (rated capacity)                                   | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | Actual Load                 | Usage in KW / Capacity in Use (whichever exists)                                  | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | KW-HR/Month                 | No Of Units (KW-HR/ Month)                                                        | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | Unit Rate                   | Unit Rate (per KW-HR)                                                             | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | No. of Units                | No Of Units (KW-HR/ Month)                                                        | CUSTOMER_DETAILS/AIROLI |
| Power Capacity | Raw Power (Genset)          | Raw Power (Genset & Transformer & Demand) — KW  / Generator kW                    | CAPACITY_GRID, FACILITY_GRID |
| Power Capacity | Raw Power (Transformer)     | Transformer kW / Utility Sanction Load KVA                                        | CAPACITY_GRID, FACILITY_GRID |
| Power Capacity | Raw Power (Demand)          | Utility Sanction Load KVA / Contract Demand                                       | FACILITY_GRID, CAPACITY_GRID |

### ACTUAL PUE  (floating metric row on CUSTOMER_DETAILS + FACILITY_GRID)
| Canonical                               | Real header                                                                  |
|-----------------------------------------|------------------------------------------------------------------------------|
| Actual PUE | Power Usage                | Actual PUE (floating cell r0) / PUE column in Inventory Summary              |

### RATED TO CONSUMED  (floating metric row on CUSTOMER_DETAILS)
| Canonical                               | Real header                                                                  |
|-----------------------------------------|------------------------------------------------------------------------------|
| Rated to Consumed | Ratio              | Rated to Consumed (floating cell r0, value in adjacent cell)                 |
| Rated to Consumed | Unnamed: X_level_1 | (parent-only banner cell) → same as Ratio                                    |

### GENSET HR/MO  (floating metric row on CUSTOMER_DETAILS)
| Canonical                               | Real header                                                                  |
|-----------------------------------------|------------------------------------------------------------------------------|
| Genset Hr/Mo | Seating Space            | Genset Hr/Mo (floating cell r0, value in adjacent cell)                      |
| Genset Hr/Mo | Unnamed: X_level_1       | (parent-only banner cell) → same as above                                    |

### REVENUE band  (CUSTOMER_DETAILS only)
| Canonical                               | Real sub-header                                                              |
|-----------------------------------------|------------------------------------------------------------------------------|
| Revenue | Monthly                       | Total Revenue (Revenue (Monthly) band)                                       |
| Additional Charges | MRC                  | Additional Capacity Charges (MRC)                                            |
| Multiplier | Unnamed: X_level_1         | Multiplier (banner-only)                                                     |

### CAPACITY band  (CAPACITY_GRID — Rabale T1/T2)
| Canonical                               | Real header in Rabale-T1/T2                                                  |
|-----------------------------------------|------------------------------------------------------------------------------|
| Capacity | Total Purchased              | Maximum Usable Capacity                                                      |
| Capacity | In Use                       | Current utilization                                                          |
| Capacity | Reserved                     | Committed (Based on Confirmed orders)                                        |
| Capacity | Surplus                      | Balance (when > 0)                                                           |
| Capacity | Leakage                      | Balance (when < 0) / -(Balance)                                              |

### Identity / layout (all archetypes)
| Canonical                               | Real header                                                                  |
|-----------------------------------------|------------------------------------------------------------------------------|
| Customer Name | Unnamed: X_level_1      | Customer Name                                                                |
| Floor | Unnamed: X_level_1              | Floor                                                                         |
| Module | Unnamed: X_level_1             | Floor / Module                                                               |
| Description | Unnamed: X_level_1        | Description (FACILITY_GRID r0 col 0)                                         |
| Remarks | Unnamed: X_level_1            | Remarks / Remarks if any (last column in CUSTOMER_DETAILS; col "Remarks" in FACILITY_GRID) |

### Executor resolution rules for the canonical registry
1. For each op, the executor iterates over every in-scope sheet and classifies the archetype.
2. It looks up each requested canonical name in the registry. If the canonical is supported by the current archetype, it fuzzy-matches the real sub-header in that sheet and uses it.
3. If the canonical is NOT supported by the current archetype (e.g. `Revenue | Monthly` requested on a CAPACITY_GRID sheet), the sheet is silently skipped for that canonical. It is NEVER fabricated.
4. For `Unnamed: N_level_1` canonicals, the executor:
   a. If reading with `header=[0,1]`, it matches the exact pandas artifact string.
   b. If reading with `header=0` (SINGLE_LEVEL), it treats the canonical as parent-only and maps it to the real flat header via the registry note above.
5. Decimal values pass through untouched. Dates pass through as stored. Case/whitespace is normalised only for MATCHING, never for OUTPUT.
6. Source_File and Source_Sheet are prepended to every returned row.

## JSON OUTPUT FORMAT  (v3)
Return a JSON array. Each element is one operation:
{
  "id": "op1",
  "type": "list" | "aggregate" | "count" | "cell_lookup" | "column_fetch",
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
  "files":   ["<filename substring>", ...] | null,
  "sheets":  ["<sheet name substring>", ...] | null,
  "operation": "sum"|"avg"|"mean"|"min"|"max"|"count"|"std"|"median"|"variance"|"range"|"count_nonzero"|"top"|"bottom" | null,
  "field_hint": "<exact phrase from the legacy list>" | null,
  "canonical_columns": ["Parent | Sub", ...] | null,
  "top_n": integer | null,
  "group_by_location": true | false,
  "customer_name": "<exact customer name from query>" | null,
  "cell_value": "<exact cell value from query>" | null,
  "target_column_hint": "<column keyword from query>" | null,
  "return_columns": ["<canonical or raw col>", ...] | null,
  "match_mode": "exact" | "contains" | "regex" | null
}

Two new fields in v3:
- `canonical_columns` — array of canonical two-level names (`"Parent | Sub"`) the user asked about. Use this whenever the user's query names any of the canonical headers in the registry above. The executor will resolve each via the registry and return one output column per canonical name per matching row.
- `files` / `sheets` — optional scope restrictors. When the user says "from Airoli and Kolkata only" → files=["Airoli","Kolkata"]. When the user says "only the Customer details and Terminated sheets" → sheets=["Customer details","Terminated"]. Substring match, case-insensitive. Null means ALL.

## TYPE = "column_fetch"  (new in v3)
Use when the user's query is essentially "fetch columns X, Y, Z from all/selected files and sheets" WITHOUT a cell-value filter. Example: "show me Power Capacity | Contracted, Power Capacity | Consumed, and Revenue | Monthly from all files".

Required JSON shape:
{
  "id": "op1",
  "type": "column_fetch",
  "label": "Column Fetch — <n> columns × <scope>",
  "canonical_columns": ["Power Capacity | Contracted", "Power Capacity | Consumed", "Revenue | Monthly"],
  "files": null,
  "sheets": null,
  "filter": null,
  "location": null,
  "customer_name": null,
  "cell_value": null,
  "target_column_hint": null,
  "match_mode": null,
  "operation": null,
  "field_hint": null,
  "top_n": null,
  "group_by_location": false,
  "return_columns": null
}

Executor behaviour for column_fetch:
1. For each in-scope file and sheet, detect archetype.
2. For each canonical in `canonical_columns`, resolve via the registry. If unsupported by this archetype → skip (never fabricate).
3. Emit ONE row per underlying data row, with columns = [Source_File, Source_Sheet, <canonical_1>, <canonical_2>, ...]. If a canonical is unsupported in that sheet, the cell is emitted as an explicit empty string "" (never a synthesised value).
4. De-dup as per Rule 10. Preserve every decimal exactly.
5. If the user also provided `customer_name`, `location`, `cell_value`, or `filter`, apply them on top of the column_fetch as additional filters.

## EXACT field_hint PHRASES (legacy single-field hints — still supported):
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

When the user query mixes legacy phrases and canonical names, populate BOTH `field_hint` and `canonical_columns` as appropriate.

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
| fetch columns X, Y, Z / show me <col list> | type="column_fetch", operation=null |
| any cell value shown / fetch rows where cell = X → type="cell_lookup" |

## LOCATION ALIASES
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

## FILE / SHEET SCOPE RESTRICTION  (new in v3)
- "from Airoli and Noida-01 only"         → files=["Airoli","Noida_01"], sheets=null
- "only the Customer details sheet"       → files=null, sheets=["Customer details"]
- "Noida-01 file, Terminated sheet only"  → files=["Noida_01"], sheets=["Terminated"]
- "all files, all sheets" / not specified → files=null, sheets=null
Substring match is case-insensitive. When the user names a sheet that does not exist in a scoped file, that file contributes zero rows for that op (never fabricate).

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
If sub-query N is a "list"/"column_fetch" with a filter, and sub-query N+1 is an "aggregate" on the SAME subject with no explicit filter, inherit the filter from sub-query N.

## COMPLEX QUERY DECOMPOSITION
Queries joined by "and", "or", commas, semicolons, or "also" = multiple operations in the array.
- "and" between filters = intersection (BOTH must be true)
- "or" between locations = union
- "and" between actions on the same filter = same filter, multiple operation objects

## CUSTOMER NAME FILTER RULES
- When the user names a specific customer, populate `customer_name` with the exact name as typed.
- The executor applies a case-insensitive partial-match on the Customer Name column BEFORE returning any rows.
- The output contains ONLY rows belonging to that customer. Zero rows from other customers.
- No trailing "— N row(s)" footer listing all customers.
- If not found, return an empty result with "No matching customer found". Never fall back.

## LOCATION STRICT FILTER RULES
- The executor filters rows to only sheets whose parent file maps to the requested location(s).
- Zero rows from unrequested locations. No fallback, no default.

## VALUE / CONDITION FILTER RULES
- Apply all stated conditions cumulatively with AND logic.
- Output columns limited to what the user asked for. Do NOT dump all columns.
- No hallucinated rows. No omission of genuinely matching rows.

## ============================================================
## CELL-VALUE QUERY RULES  (unchanged from v2, still the core of cell_lookup)
## ============================================================

### When to use type = "cell_lookup"
Use when the user asks for rows keyed on the value of a specific column cell, e.g.:
- "show all rows where Floor = 3rd Floor"
- "list customers whose UoM is KVA"
- "find all entries where Caged/Uncaged = Caged in Bangalore"
- "which customers are on RHS"
- "show rows where Power Subscription Model is Rated"
- "customers whose Subscription Mode = Rack"
- "rows where Unit Rate Model = Fixed"
- "fetch cell value 250 from Total Capacity Purchased"
- "any row where Billing Frequency = Monthly"
- "rows where the cell value 'Noida-01' appears in Floor / Module"
- "show rows where Power Capacity | Contracted = 305"                ← canonical form also accepted
- "find rows where Billing Model | Rated is true in Kolkata"          ← canonical form also accepted

### Required JSON fields for cell_lookup
- "type"                 : "cell_lookup"
- "target_column_hint"   : the column the user is asking about (user's own words OR canonical "Parent | Sub" form)
- "cell_value"           : the exact value the user typed
- "match_mode"           : "exact" (default), "contains", or "regex"
- "return_columns"       : optional list; if null → standard safe projection
- "location"             : honours LOCATION ALIASES
- "files" / "sheets"     : optional scope restrictors
- "filter"               : may combine with cell_lookup
- "customer_name"        : may combine with cell_lookup
- "canonical_columns"    : optional — the executor will ALSO return these canonical columns in the projection

### Standard safe projection (when return_columns is null)
In this order, silently omitting any the sheet does not have:
1.  Source_File
2.  Source_Sheet
3.  Floor
4.  Floor / Module
5.  Customer Name
6.  RHS/SH
7.  Power Subscription Model (Rated/Subscribed)
8.  Power Usage Model (Bundled / Metered)
9.  Subscription Mode
10. Caged /Uncaged
11. UoM
12. Subscription
13. In Use
14. Total Capacity Purchased
15. Capacity in Use
16. Per Unit rate (MRC)
17. Total Revenue
18. Billing Frequency
19. Contract Start Date
20. Current Expiry Date
PLUS the target column itself if not already in the list.

### Executor semantics for cell_lookup (v3)
1. Load every in-scope sheet. Detect archetype per sheet. Auto-detect header row(s) per archetype.
2. Resolve `target_column_hint`:
   - If it matches a canonical "Parent | Sub" name → look up in the CANONICAL COLUMN REGISTRY → real sub-header per archetype.
   - Otherwise → fuzzy-map to the real sub-header via token overlap.
   - If the column is not present in this sheet's archetype → SKIP this sheet silently (never fabricate).
3. Normalise both cell and query value (strip / collapse whitespace / casefold; numeric cells via float).
4. Apply match_mode. Keep only matching rows.
5. De-duplicate via Rule 10.
6. Project to return_columns (+canonical_columns if present) or the standard safe projection.
7. Preserve every cell value EXACTLY — no reformatting of numbers, dates, casing.
8. If ZERO rows match across all in-scope sheets, return a single empty result object labelled "No matching record found for <target> = <value>". NEVER fall back to an unrelated dataset.
9. NEVER invent a column value. NEVER fuse cells from different rows.

### ANY-COLUMN / ANY-CELL MODE
If the user asks "find cell value 250" or "show every row that contains 'Wipro'" without naming a column:
- target_column_hint = "any", cell_value = the user value, match_mode = "contains".
- Executor scans EVERY column of EVERY in-scope sheet. Returns safe projection + synthetic "Matched_Column".

### Forbidden behaviours for cell_lookup
- Returning rows whose target column value does NOT match cell_value.
- Returning the same row twice because it appears in two sheets.
- Returning rows from a location / file / sheet the user did not ask for.
- Returning a "closest guess" when no exact match exists.
- Silently dropping a sheet that DOES contain the target column.
- Re-formatting cell values.
- Fabricating Source_File or Source_Sheet.

## UNIT AWARENESS (inform the label — do NOT change field_hint)
- Power/capacity columns → KW or KVA (varies per row's UoM column)
- Space columns → Sq Ft
- Rack/seating columns → Racks or Seats
- Revenue columns → ₹/month (MRC)
- Rate columns → ₹/KW-HR
Always include the unit in the label string (e.g., "Total Power Purchased (KW)").

## ============================================================
## EXAMPLES  (v3 — including the new canonical + scope forms)
## ============================================================

Query: "fetch Power Capacity | Contracted, Power Capacity | Consumed, and Revenue | Monthly from all files and all sheets"
→ [{"id":"op1","type":"column_fetch","label":"Power Capacity (Contracted, Consumed) + Revenue (Monthly) — All Files","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Power Capacity | Contracted","Power Capacity | Consumed","Revenue | Monthly"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "show Billing Model | Unnamed: 1_level_1, Space | Caged/Uncaged and Power Capacity | Contracted for all customers in Noida"
→ [{"id":"op1","type":"column_fetch","label":"Billing Model / Caged / Contracted — Noida","filter":null,"location":["noida"],"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Billing Model | Unnamed: 1_level_1","Space | Caged/Uncaged","Power Capacity | Contracted"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "from Airoli and Kolkata files only, fetch Customer Name | Unnamed: X_level_1, Power Capacity | Contracted and Power Capacity | Consumed"
→ [{"id":"op1","type":"column_fetch","label":"Customer / Contracted / Consumed — Airoli + Kolkata","filter":null,"location":null,"files":["Airoli","Kolkata"],"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Customer Name | Unnamed: X_level_1","Power Capacity | Contracted","Power Capacity | Consumed"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "only the Customer details and Terminated sheets — give me Power Capacity | Rated Load, Power Capacity | Actual Load, Revenue | Monthly"
→ [{"id":"op1","type":"column_fetch","label":"Rated Load / Actual Load / Monthly Revenue — Customer details + Terminated","filter":null,"location":null,"files":null,"sheets":["Customer details","Terminated"],"operation":null,"field_hint":null,"canonical_columns":["Power Capacity | Rated Load","Power Capacity | Actual Load","Revenue | Monthly"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "show Capacity | Total Purchased, Capacity | In Use, Capacity | Surplus and Capacity | Leakage from Rabale T1 and T2"
→ [{"id":"op1","type":"column_fetch","label":"Capacity Matrix — Rabale T1/T2","filter":null,"location":["rabale t1","rabale t2"],"files":["Rabale_T1_T2"],"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Capacity | Total Purchased","Capacity | In Use","Capacity | Reserved","Capacity | Surplus","Capacity | Leakage"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "give me Description | Unnamed: X_level_1 and Remarks | Unnamed: X_level_1 from the Facility details and Summary sheets of Bangalore and Kolkata"
→ [{"id":"op1","type":"column_fetch","label":"Description + Remarks — Facility/Summary — Bangalore + Kolkata","filter":null,"location":["bangalore","kolkata"],"files":["Bangalore","Kolkata"],"sheets":["Facility details","Summary"],"operation":null,"field_hint":null,"canonical_columns":["Description | Unnamed: X_level_1","Remarks | Unnamed: X_level_1"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "fetch Actual PUE | Power Usage and Rated to Consumed | Ratio from all files"
→ [{"id":"op1","type":"column_fetch","label":"Actual PUE + Rated-to-Consumed — All Files","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Actual PUE | Power Usage","Rated to Consumed | Ratio"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "show rows where Billing Model | Rated is set in Noida and also sum Power Capacity | Contracted for them"
→ [
  {"id":"op1","type":"cell_lookup","label":"Rated rows — Noida","filter":{"caged":null,"uncaged":null,"rated":true,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":["noida"],"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Billing Model | Unnamed: 1_level_1","Power Capacity | Contracted"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Rated","target_column_hint":"Billing Model | Rated","return_columns":null,"match_mode":"exact"},
  {"id":"op2","type":"aggregate","label":"Sum Power Capacity | Contracted — Rated, Noida","filter":{"caged":null,"uncaged":null,"rated":true,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":["noida"],"files":null,"sheets":null,"operation":"sum","field_hint":"total capacity purchased","canonical_columns":["Power Capacity | Contracted"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}
]

Query: "show Wipro rows where Power Capacity | Consumed > 0 across all files"
→ [{"id":"op1","type":"cell_lookup","label":"Wipro rows with Consumed > 0","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Power Capacity | Consumed"],"top_n":null,"group_by_location":false,"customer_name":"Wipro","cell_value":">0","target_column_hint":"Power Capacity | Consumed","return_columns":null,"match_mode":"regex"}]

Query: "sum of power purchased"   (legacy, still works)
→ [{"id":"op1","type":"aggregate","label":"Total Power Purchased (KW/KVA)","filter":null,"location":null,"files":null,"sheets":null,"operation":"sum","field_hint":"total capacity purchased","canonical_columns":["Power Capacity | Contracted"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "total revenue by location"
→ [{"id":"op1","type":"aggregate","label":"Total Revenue (₹/month) by Location","filter":null,"location":null,"files":null,"sheets":null,"operation":"sum","field_hint":"total revenue","canonical_columns":["Revenue | Monthly"],"top_n":null,"group_by_location":true,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "list caged customers in noida"
→ [{"id":"op1","type":"list","label":"Caged Customers — Noida","filter":{"caged":true,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":["noida"],"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "show Wipro customer details"
→ [{"id":"op1","type":"list","label":"Wipro Customer Details","filter":{"caged":null,"uncaged":null,"rated":null,"subscribed":null,"bundled":null,"metered":null,"rhs":null,"shs":null},"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":null,"top_n":null,"group_by_location":false,"customer_name":"Wipro","cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "find rows where Caged/Uncaged = Caged in Bangalore"
→ [{"id":"op1","type":"cell_lookup","label":"Caged rows — Bangalore","filter":null,"location":["bangalore"],"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":["Space | Caged/Uncaged"],"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Caged","target_column_hint":"caged/uncaged","return_columns":null,"match_mode":"exact"}]

Query: "any row that contains Infosys anywhere"
→ [{"id":"op1","type":"cell_lookup","label":"Any cell containing 'Infosys'","filter":null,"location":null,"files":null,"sheets":null,"operation":null,"field_hint":null,"canonical_columns":null,"top_n":null,"group_by_location":false,"customer_name":null,"cell_value":"Infosys","target_column_hint":"any","return_columns":null,"match_mode":"contains"}]

Query: "how many customers per location"
→ [{"id":"op1","type":"count","label":"Customer Count by Location","filter":null,"location":null,"files":null,"sheets":null,"operation":"count","field_hint":null,"canonical_columns":null,"top_n":null,"group_by_location":true,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]

Query: "top 5 customers by revenue"
→ [{"id":"op1","type":"aggregate","label":"Top 5 Customers by Revenue","filter":null,"location":null,"files":null,"sheets":null,"operation":"top","field_hint":"total revenue","canonical_columns":["Revenue | Monthly"],"top_n":5,"group_by_location":false,"customer_name":null,"cell_value":null,"target_column_hint":null,"return_columns":null,"match_mode":null}]
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
                return _robust_to_numeric(CUST[col]).sum()
            return None

        def _avg(col):
            if col and col in CUST.columns:
                s = _robust_to_numeric(CUST[col]).dropna()
                return s.mean() if not s.empty else None
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
            k[3].markdown(kpi_html(fmt(tot_cap), cap_c,
                                   "Power Capacity section", GREEN), unsafe_allow_html=True)

        if use_c:
            tot_use = _n(use_c)
            k[4].markdown(kpi_html(fmt(tot_use), use_c,
                                   "Power Capacity section", AMBER), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Billing Model</div>', unsafe_allow_html=True)
        bm_cols = st.columns(4)

        if caged_c:
            cage_vals = CUST[caged_c].astype(str).str.upper().str.strip()
            n_caged   = (cage_vals == "CAGED").sum()
            n_uncaged = (cage_vals == "UNCAGED").sum()
            bm_cols[0].markdown(kpi_html(f"{n_caged}", caged_c,
                                         f"Uncaged: {n_uncaged}", CYAN), unsafe_allow_html=True)

        if own_c:
            own_vals = CUST[own_c].astype(str).str.strip().str.upper()
            n_sify     = int(own_vals.str.contains(r"SIFY", na=False).sum())
            n_customer = int(own_vals.str.contains(r"CUSTOMER|CUST(?!OM)", na=False).sum())
            if n_sify > 0 or n_customer > 0:
                bm_cols[1].markdown(
                    kpi_html(f"Sify: {n_sify}", "Space | Ownership",
                             f"Customer: {n_customer}", LBLUE), unsafe_allow_html=True)
            else:
                rhs_c_cnt = _cnt_val(own_c, "RHS")
                shs_c_cnt = _cnt_val(own_c, "SHS")
                if rhs_c_cnt is not None or shs_c_cnt is not None:
                    bm_cols[1].markdown(
                        kpi_html(f"RHS: {rhs_c_cnt or 0}", "Space | Ownership",
                                 f"SHS: {shs_c_cnt or 0}", LBLUE), unsafe_allow_html=True)

        if pw_sub_c:
            rated = _cnt_val(pw_sub_c, "RATED")
            subsc = _cnt_val(pw_sub_c, "SUBSCRIBED")
            bm_cols[2].markdown(
                kpi_html(f"{rated or 0}", pw_sub_c,
                         f"Subscribed: {subsc or 0}", AMBER), unsafe_allow_html=True)

        if pw_use_m_c:
            bundled = _cnt_val(pw_use_m_c, "BUNDLED")
            metered = _cnt_val(pw_use_m_c, "METERED")
            bm_cols[3].markdown(
                kpi_html(f"{bundled or 0}", pw_use_m_c,
                         f"Metered: {metered or 0}", GREEN), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Space</div>', unsafe_allow_html=True)
        sp_cols = st.columns(5)

        if sub_mode_c:
            sub_vals_upper = CUST[sub_mode_c].astype(str).str.strip().str.upper()
            n_seats = int(sub_vals_upper.str.contains(r"SEAT|NO.*SEAT", na=False).sum())
            n_space = int(sub_vals_upper.str.contains(r"SPACE", na=False).sum())
            if n_seats > 0 or n_space > 0:
                sp_cols[0].markdown(
                    kpi_html(f"Seats: {n_seats}", "Sitting Space | Subscription Model",
                             f"Space: {n_space}", CYAN),
                    unsafe_allow_html=True)
            else:
                rack_m = _cnt_val(sub_mode_c, "RACK")
                u_m    = _cnt_val(sub_mode_c, "U SPACE")
                sqft_m = _cnt_val(sub_mode_c, "SQFT SPACE")
                sp_cols[0].markdown(
                    kpi_html(f"{rack_m or 0}", sub_mode_c,
                             f"U Space: {u_m or 0} | SqFt: {sqft_m or 0}", CYAN),
                    unsafe_allow_html=True)

        if space_sub_c:
            v = _n(space_sub_c)
            if v is not None:
                sp_cols[1].markdown(kpi_html(fmt(v), space_sub_c,
                                             space_sub_c[:25], GREEN), unsafe_allow_html=True)

        if space_inuse_c:
            v = _n(space_inuse_c)
            if v is not None:
                sp_cols[2].markdown(kpi_html(fmt(v), space_inuse_c,
                                             space_inuse_c[:25], AMBER), unsafe_allow_html=True)

        if space_ytbg_c:
            v = _n(space_ytbg_c)
            if v is not None:
                sp_cols[3].markdown(kpi_html(fmt(v), space_ytbg_c,
                                             space_ytbg_c[:25], RED), unsafe_allow_html=True)

        if space_rate_c:
            v = _avg(space_rate_c)
            if v is not None:
                sp_cols[4].markdown(kpi_html(fmt(v), space_rate_c,
                                             space_rate_c[:25], LBLUE), unsafe_allow_html=True)
        elif rack_c:
            v = _n(rack_c)
            if v is not None:
                sp_cols[4].markdown(kpi_html(fmt(v), rack_c,
                                             rack_c[:25], LBLUE), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Power Capacity</div>', unsafe_allow_html=True)
        pc_cols = st.columns(5)

        if cap_c:
            pc_cols[0].markdown(kpi_html(fmt(_n(cap_c)), cap_c,
                                         cap_c[:25], GREEN), unsafe_allow_html=True)
        if use_c:
            pc_cols[1].markdown(kpi_html(fmt(_n(use_c)), use_c,
                                         use_c[:25], AMBER), unsafe_allow_html=True)
        if cap_ytbg_c:
            v = _n(cap_ytbg_c)
            if v is not None:
                pc_cols[2].markdown(kpi_html(fmt(v), cap_ytbg_c,
                                             cap_ytbg_c[:25], RED), unsafe_allow_html=True)
        if sub_kw_c:
            v = _n(sub_kw_c)
            if v is not None:
                pc_cols[3].markdown(kpi_html(fmt(v), sub_kw_c,
                                             sub_kw_c[:25], LBLUE), unsafe_allow_html=True)
        if alloc_kw_c:
            v = _n(alloc_kw_c)
            if v is not None:
                pc_cols[4].markdown(kpi_html(fmt(v), alloc_kw_c,
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
            for col, color in [
                (pu_sub_c,   GREEN),
                (pu_inuse_c, AMBER),
                (pu_ytbg_c,  RED),
            ]:
                if col:
                    v = _n(col)
                    if v is not None:
                        pu_cols[i].markdown(kpi_html(fmt(v), col, col[:25], color),
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
                    ss_cols[0].markdown(kpi_html(fmt(v), seat_sub_c,
                                                 seat_sub_c[:25], CYAN), unsafe_allow_html=True)
            if seat_inuse_c:
                v = _n(seat_inuse_c)
                if v is not None:
                    ss_cols[1].markdown(kpi_html(fmt(v), seat_inuse_c,
                                                 seat_inuse_c[:25], AMBER), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        rev_c = rev_total_c or rev_mrc_c
        st.markdown('<div class="section-title">Revenue (Monthly)</div>', unsafe_allow_html=True)
        rv_cols = st.columns(5)
        rv_items = [
            (rev_space_c,  CYAN),
            (rev_addcap_c, LBLUE),
            (rev_pwuse_c,  GREEN),
            (rev_seat_c,   AMBER),
            (rev_other_c,  MUTED),
        ]
        filled = 0
        for col, color in rv_items:
            if col and filled < 5:
                v = _n(col)
                if v is not None:
                    rv_cols[filled].markdown(kpi_html(fmt(v), col, col[:25], color),
                                             unsafe_allow_html=True)
                    filled += 1

        st.markdown("<br>", unsafe_allow_html=True)
        rv2_cols = st.columns(4)
        rv2_items = [
            (rev_total_c,  GREEN),
            (rev_mrc_c,    LBLUE),
        ]
        filled2 = 0
        for col, color in rv2_items:
            if col and filled2 < 4:
                v = _n(col)
                if v is not None:
                    rv2_cols[filled2].markdown(kpi_html(fmt(v), col, col[:25], color),
                                               unsafe_allow_html=True)
                    filled2 += 1

        if rev_freq_c:
            freq_counts = CUST[rev_freq_c].dropna().value_counts()
            top_freq = freq_counts.index[0] if not freq_counts.empty else "—"
            rv2_cols[min(filled2, 3)].markdown(
                kpi_html(str(top_freq), rev_freq_c,
                         f"{len(freq_counts)} types", AMBER), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown('<div class="section-title">Contract Information</div>', unsafe_allow_html=True)
        ci_cols = st.columns(4)
        ci_i = 0

        if con_start_c:
            non_null = CUST[con_start_c].dropna()
            ci_cols[ci_i].markdown(
                kpi_html(f"{len(non_null):,}", con_start_c,
                         con_start_c[:25], CYAN), unsafe_allow_html=True)
            ci_i += 1

        if con_term_c:
            v = _avg(con_term_c)
            if v is not None:
                ci_cols[ci_i].markdown(kpi_html(f"{v:.1f} yr", con_term_c,
                                                con_term_c[:25], GREEN), unsafe_allow_html=True)
                ci_i += 1

        if con_expiry_c:
            non_null = CUST[con_expiry_c].dropna()
            ci_cols[ci_i].markdown(
                kpi_html(f"{len(non_null):,}", con_expiry_c,
                         con_expiry_c[:25], AMBER), unsafe_allow_html=True)
            ci_i += 1

        if rev_so_c:
            so_count = CUST[rev_so_c].dropna().nunique()
            ci_cols[min(ci_i, 3)].markdown(
                kpi_html(f"{so_count:,}", rev_so_c,
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

            # Space/Rack Utilisation: space_in_use / space_subscription
            t_space_sub  = _n(space_sub_c)  or 0
            t_space_use  = _n(space_inuse_c) or 0
            if t_space_sub > 0:
                rack_pct = min((t_space_use / t_space_sub * 100), 100)
            elif rack_c:
                t_rack_sub = _n(rack_c) or 0
                rack_pct = min((t_space_use / t_rack_sub * 100) if t_rack_sub > 0 else util_pct, 100)
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

            space_util_label = (
                "Space/Rack Utilisation (%)<br>(Space In Use / Space Subscription)"
                if t_space_sub > 0 else "Space/Rack Utilisation (%)"
            )
            g1.plotly_chart(_gauge(util_pct, "Capacity Utilisation (%)", LBLUE),
                            use_container_width=True)
            g2.plotly_chart(_gauge(rack_pct, space_util_label, GREEN),
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

    # ── Unified query input (AI query + customer name-wise full row search) ──────
    st.markdown(
        f'<div style="font-size:.82rem;color:{MUTED};margin-bottom:6px">'
        f'Type an <b style="color:{CYAN}">AI query</b> (e.g. <i>sum of power</i>, '
        f'<i>list caged customers</i>) <b>or</b> a '
        f'<b style="color:{CYAN}">customer name</b> (e.g. <i>Wipro</i>, <i>Oracle</i>) '
        f'to retrieve all column values for every matching row across all 10 Excel files '
        f'and all sheets. Both modes work in the same box.</div>',
        unsafe_allow_html=True,
    )

    query = st.text_area(
        "🔍 Enter your query or customer name",
        placeholder=(
            "AI query: List all caged customers AND sum capacity in use AND total power used\n"
            "Customer search: Wipro   |   Oracle   |   CISCO SYSTEMS   |   YES BANK\n"
            "Combined: power capacity purchased for Oracle"
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
                results = []

                # ── Step A: Customer name-wise full row search (merged) ────────
                # Always search customer names across ALL 10 Excel files (full pool),
                # not just the filtered sq_locs pool, for completeness.
                _sq_merged_pool = combined_df(ALL)   # all 10 files, all sheets
                _sq_cust_name, _sq_field_part = _sq_detect_customer_query(query.strip())

                # Also treat the whole query as a possible customer name
                _sq_name_candidates = [_sq_cust_name] if _sq_cust_name else []
                _pure_name = query.strip()
                # If query has no spaces or is short / not an obvious AI query,
                # try it as a customer name directly
                _ai_keywords = {
                    "sum", "total", "list", "show", "count", "average", "mean",
                    "max", "min", "top", "bottom", "all", "get", "fetch", "display",
                    "caged", "uncaged", "rated", "metered", "bundled", "subscribed",
                }
                _query_words = set(_pure_name.lower().split())
                _looks_like_ai = bool(_query_words & _ai_keywords)

                if not _looks_like_ai and _pure_name not in _sq_name_candidates:
                    _sq_name_candidates.append(_pure_name)

                _cnw_rows_all = pd.DataFrame()
                _cnw_used_name = ""
                for _cand in _sq_name_candidates:
                    if not _cand.strip():
                        continue
                    _try_rows = _sk_find_customers_all(_cand.strip(), _sq_merged_pool)
                    if not _try_rows.empty:
                        _cnw_rows_all = _try_rows
                        _cnw_used_name = _cand.strip()
                        break

                if not _cnw_rows_all.empty:
                    _cnw_disp = _sk_canonical_name(_cnw_rows_all, _cnw_used_name)
                    _cnw_n    = len(_cnw_rows_all)
                    _cnw_files_found  = sorted(_cnw_rows_all["_Location"].unique().tolist()) if "_Location" in _cnw_rows_all.columns else []
                    _cnw_sheets_found = sorted(_cnw_rows_all["_Sheet"].unique().tolist())    if "_Sheet"    in _cnw_rows_all.columns else []

                    # Show summary card
                    st.markdown(
                        f'<div style="background:{DARK2};border:2px solid {CYAN};'
                        f'border-radius:14px;padding:20px 26px;margin:14px 0">'
                        f'<div style="font-size:1.1rem;font-weight:900;color:{WHITE};margin-bottom:6px">'
                        f'✅ Customer Found: {_cnw_disp}</div>'
                        f'<div style="font-size:.85rem;color:{TEXT}">'
                        f'<b style="color:{CYAN}">{_cnw_n}</b> row(s) matched across '
                        f'<b style="color:{CYAN}">{len(_cnw_files_found)}</b> DC file(s) and '
                        f'<b style="color:{CYAN}">{len(_cnw_sheets_found)}</b> sheet(s)</div>'
                        f'<div style="margin-top:10px;font-size:.78rem;color:{MUTED}">Files: '
                        + (", ".join(f'<span style="color:{GREEN}">{l}</span>' for l in _cnw_files_found) or "—")
                        + f'</div><div style="font-size:.78rem;color:{MUTED};margin-top:4px">Sheets: '
                        + (", ".join(f'<span style="color:{AMBER}">{s}</span>' for s in _cnw_sheets_found[:15]) or "—")
                        + f'</div></div>',
                        unsafe_allow_html=True,
                    )

                    # Validation expander
                    with st.expander("🔬 Validation — Row count per file & sheet", expanded=False):
                        if "_Location" in _cnw_rows_all.columns and "_Sheet" in _cnw_rows_all.columns:
                            _cnw_val_df = (
                                _cnw_rows_all.groupby(["_Location", "_Sheet"])
                                .size().reset_index(name="Row Count")
                                .sort_values(["_Location", "_Sheet"])
                            )
                            _cnw_val_df.index = range(1, len(_cnw_val_df) + 1)
                            st.dataframe(_cnw_val_df, use_container_width=True)

                    # Full row display — all columns
                    _cnw_meta_c = [c for c in ["_Location", "_Sheet"] if c in _cnw_rows_all.columns]
                    _cnw_data_c = [c for c in _cnw_rows_all.columns if not c.startswith("_")]
                    _cnw_disp_df = _cnw_rows_all[_cnw_meta_c + _cnw_data_c].copy()
                    _cnw_disp_df.index = range(1, len(_cnw_disp_df) + 1)
                    st.markdown(
                        f'<div style="font-size:.8rem;color:{CYAN};font-weight:700;'
                        f'text-transform:uppercase;letter-spacing:.05em;margin:16px 0 6px">'
                        f'📋 All Columns — {_cnw_n} row(s) for "{_cnw_disp}"</div>',
                        unsafe_allow_html=True,
                    )
                    st.dataframe(_cnw_disp_df, use_container_width=True)
                    st.download_button(
                        f"⬇️ Download All Columns CSV — {_cnw_disp[:30]}",
                        _cnw_disp_df.to_csv(index=False).encode("utf-8"),
                        f"customer_allcols_{_cnw_disp.replace(' ', '_')[:30]}.csv",
                        "text/csv",
                        key="sq_cnw_dl_all",
                    )

                    # Build customer profile result card for the AI results list
                    _cust_profile_res = _sq_build_customer_profile(
                        _cnw_rows_all, _cnw_used_name, _sq_field_part or ""
                    )
                    results = [_cust_profile_res]

                # ── Step B: AI-powered structured query ───────────────────────
                # Always run AI query — it handles aggregations, lists, counts etc.
                with st.spinner("🤖 Parsing and executing AI query…"):
                    ops_raw = parse_query_with_ai(query.strip())

                if isinstance(ops_raw, tuple):
                    err_type, err_msg = ops_raw
                    st.error(f"**{'Config' if err_type == 'config_error' else 'Parse'} Error**: {err_msg}")
                elif not ops_raw:
                    if _cnw_rows_all.empty:
                        st.warning("Could not parse query. Please try rephrasing.")
                else:
                    if ops_raw and not isinstance(ops_raw, tuple):
                        ai_results = _sq_execute_with_schema(ops_raw, pool)
                        results = results + ai_results

                if results:
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
