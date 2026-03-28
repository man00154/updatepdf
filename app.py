import os, re, warnings, tempfile, subprocess
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sify DC – Capacity Tracker",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0e1a,#1a2035,#0d1b2a)!important;}
[data-testid="stSidebar"] *{color:#c9d8f0!important;}
.kcard{border-radius:14px;padding:16px 20px;color:#fff;margin-bottom:10px;
       box-shadow:0 4px 18px rgba(0,0,0,.35);transition:transform .2s;}
.kcard:hover{transform:translateY(-2px);}
.kcard h2{font-size:1.8rem;margin:0;font-weight:800;}
.kcard p{margin:3px 0 0;font-size:.82rem;opacity:.82;}
.kcard-blue  {background:linear-gradient(135deg,#1e3c72,#2a5298);}
.kcard-green {background:linear-gradient(135deg,#0b6e4f,#17a572);}
.kcard-red   {background:linear-gradient(135deg,#7b1a1a,#c0392b);}
.kcard-orange{background:linear-gradient(135deg,#7d4e00,#e67e22);}
.kcard-teal  {background:linear-gradient(135deg,#0f3460,#16213e);}
.kcard-purple{background:linear-gradient(135deg,#4a0072,#7b1fa2);}
.sec-title{font-size:1.15rem;font-weight:700;color:#1e3c72;
    border-left:5px solid #2a5298;padding-left:10px;margin:16px 0 10px;}
.q-user{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff;
    border-radius:18px 18px 4px 18px;padding:10px 16px;
    margin:10px 0 4px auto;max-width:76%;width:fit-content;
    box-shadow:0 3px 12px rgba(30,60,114,.45);float:right;clear:both;}
.ans-box{background:linear-gradient(135deg,#0f2744,#1a4a6b);color:#d0ecff;
    border-radius:12px;padding:14px 18px;margin:8px 0;font-size:.97rem;
    box-shadow:0 3px 14px rgba(0,0,0,.35);white-space:pre-wrap;line-height:1.6;}
.cell-chip{background:#1a2f1a;border-left:4px solid #27ae60;border-radius:6px;
    padding:6px 12px;margin:3px 0;font-family:monospace;font-size:.8rem;color:#b8ffb8;}
.clearfix{clear:both;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# STOP-WORDS
# ─────────────────────────────────────────────────────────────────────────────
_SW = {
    "the","and","for","are","all","any","how","what","show","give","tell","from",
    "this","that","with","get","find","list","much","many","each","every","data",
    "value","values","number","numbers","in","of","a","an","is","at","by","to",
    "do","me","my","about","details","info","please","can","you","per","across",
    "which","where","who","when","does","did","have","has","their","its","our",
    "your","there","these","those","been","will","would","could","should","shall",
    "let","some","just","also","even","only","into","over","under","both","such",
    "than","then","but","not","nor","yet","so","either","neither","versus","vs",
}

# ─────────────────────────────────────────────────────────────────────────────
# PATH HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _app_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path(os.getcwd())

EXCEL_FOLDER = _app_dir() / "excel_files"


def find_excel_files(folder: str) -> list:
    p = Path(folder)
    if not p.is_dir():
        return []
    return sorted(
        f.name for f in p.iterdir()
        if f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~")
    )


# ─────────────────────────────────────────────────────────────────────────────
# LOCATION NAME EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────
def location_from_name(fname: str) -> str:
    n = os.path.basename(fname)
    n = re.sub(r"\.(xlsx?|xls)$", "", n, flags=re.I)
    n = re.sub(r"[Cc]ustomer.?[Aa]nd.?[Cc]apacity.?[Tt]racker.?", "", n)
    n = re.sub(
        r"[_\s]?\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{2,4}.*$",
        "", n, flags=re.I,
    )
    n = re.sub(r"__\d+_*$", "", n)
    n = re.sub(r"[_]+", " ", n).strip()
    return n if n else fname


# ─────────────────────────────────────────────────────────────────────────────
# SAVE UPLOADED FILES
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def save_uploads(file_bytes_tuple: tuple) -> str:
    tmp = tempfile.mkdtemp()
    for name, data in file_bytes_tuple:
        with open(os.path.join(tmp, name), "wb") as fh:
            fh.write(data)
    return tmp


# ─────────────────────────────────────────────────────────────────────────────
# XLS -> XLSX via LibreOffice (using soffice.py wrapper if available)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def ensure_readable(original_path: str) -> str:
    if not original_path.lower().endswith(".xls"):
        return original_path
    # Confirm it's actually a legacy OLE2 .xls
    try:
        with open(original_path, "rb") as fh:
            if fh.read(4) != b"\xd0\xcf\x11\xe0":
                return original_path
    except Exception:
        return original_path
    out_dir = tempfile.mkdtemp()
    soffice_wrapper = "/mnt/skills/public/xlsx/scripts/office/soffice.py"
    try:
        if os.path.exists(soffice_wrapper):
            subprocess.run(
                ["python3", soffice_wrapper, "--convert-to", "xlsx",
                 "--outdir", out_dir, original_path],
                capture_output=True, timeout=60,
            )
        else:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "xlsx",
                 "--outdir", out_dir, original_path],
                capture_output=True, timeout=120,
            )
        base = os.path.splitext(os.path.basename(original_path))[0]
        conv = os.path.join(out_dir, base + ".xlsx")
        if os.path.exists(conv):
            return conv
    except Exception:
        pass
    return original_path


# ─────────────────────────────────────────────────────────────────────────────
# LOAD ONE FILE – all sheets via openpyxl with data_only=True
# Returns dict of sheet_name -> list of (row_idx, col_idx, value)
# This reads EVERY cell with a value, regardless of position.
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_file_raw(original_path: str) -> dict:
    """Load file and return dict: sheet_name -> DataFrame (header=None, dtype=str)"""
    path = ensure_readable(original_path)
    sheets = {}
    try:
        xf = pd.ExcelFile(path, engine="openpyxl")
        for sh in xf.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sh, header=None,
                                   engine="openpyxl", dtype=str)
                sheets[sh] = df
            except Exception:
                pass
    except Exception as e:
        st.sidebar.warning(f"⚠️ {os.path.basename(original_path)}: {e}")
    return sheets


def load_file(original_path: str) -> dict:
    return load_file_raw(original_path)


# ─────────────────────────────────────────────────────────────────────────────
# HEADER DETECTION (for Analytics / Chart tabs)
# ─────────────────────────────────────────────────────────────────────────────
def best_header_row(df: pd.DataFrame) -> int:
    best_row, best_score = 0, -1
    for i in range(min(8, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        filled = (row.str.len() > 0) & (~row.isin(["nan", "None", ""]))
        label = filled & (~row.str.match(r"^-?\d+(\.\d+)?$"))
        score = label.sum() * 2 + filled.sum()
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def smart_header(df: pd.DataFrame) -> pd.DataFrame:
    hr = best_header_row(df)
    hdr = df.iloc[hr].fillna("").astype(str).str.strip()
    seen = {}; cols = []
    for col in hdr:
        col = col if col and col not in ("nan", "None") else f"Col_{len(cols)}"
        if col in seen:
            seen[col] += 1; cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0; cols.append(col)
    data = df.iloc[hr + 1:].copy()
    data.columns = cols
    return data.dropna(how="all").reset_index(drop=True)


def to_numeric(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        out[col] = pd.to_numeric(out[col], errors="ignore")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# MULTI-SECTION HEADER MAP
# Detects ALL header-like rows and assigns each data cell the header above it.
# ─────────────────────────────────────────────────────────────────────────────
def _detect_all_header_rows(df: pd.DataFrame) -> set:
    """
    Detect header rows with strict criteria to avoid misclassifying data rows.
    Header rows are characterized by:
    - High ratio of unique text labels (column names are unique)
    - Low number of repeated values
    - Mostly short strings that look like field names, not data values
    """
    hr_set = set()
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.strip()
        filled_mask = (row.str.len() > 0) & (~row.isin(["nan", "None", ""]))
        filled_vals = row[filled_mask]
        n_filled = filled_vals.shape[0]
        if n_filled < 2:
            continue

        # Count non-numeric (label) cells
        label_mask = filled_mask & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        n_labels = label_mask.sum()
        n_unique = filled_vals.nunique()

        # Ratio checks
        label_ratio = n_labels / max(n_filled, 1)
        unique_ratio = n_unique / max(n_filled, 1)

        # Count repeated text values (data rows have many "SUBSCRIBED", "RACK" etc.)
        value_counts = filled_vals.value_counts()
        n_repeated = (value_counts > 1).sum()  # how many values appear more than once

        # A true header row should have:
        # 1. Very high label ratio (almost all text)
        # 2. Very high unique ratio (column names are different)
        # 3. Low repetition (data rows repeat values like RATED, BUNDLED, RACK)
        is_header = False

        # Strong header: mostly unique labels, very few repeats
        if (label_ratio >= 0.80 and unique_ratio >= 0.75
                and n_repeated <= max(2, n_filled * 0.15) and n_unique >= 3):
            is_header = True

        # Also detect sub-headers (like "Billing Model", "Space", "Power Capacity")
        # These tend to have few filled cells but all unique text
        if (n_filled <= 10 and n_filled >= 2 and label_ratio >= 0.90
                and unique_ratio >= 0.80 and n_repeated <= 1):
            is_header = True

        if is_header:
            hr_set.add(i)
    return hr_set


def _build_cell_col_map(df: pd.DataFrame):
    hr_set = _detect_all_header_rows(df)
    hr_maps = {}
    for hr in hr_set:
        m = {}
        for c in range(df.shape[1]):
            v = str(df.iat[hr, c]).strip()
            if v and v not in ("nan", "None"):
                m[c] = v
        hr_maps[hr] = m
    sorted_hrs = sorted(hr_set)
    cell_map = {}
    for r in range(df.shape[0]):
        prev = [h for h in sorted_hrs if h < r]
        for c in range(df.shape[1]):
            name = f"Col_{c}"
            for h in reversed(prev):
                if c in hr_maps[h]:
                    name = hr_maps[h][c]
                    break
            cell_map[(r, c)] = name
    return cell_map, hr_set


# ─────────────────────────────────────────────────────────────────────────────
# BUILD FULL CORPUS — reads EVERY non-empty cell from EVERY position
# Uses openpyxl data_only=True to get computed values, not formulas.
# For .xls files, converts to .xlsx first.
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def build_corpus(file_list: tuple, folder: str):
    corpus = []
    row_records = defaultdict(dict)

    for fname in file_list:
        full = os.path.join(folder, fname)
        if not os.path.isfile(full):
            continue
        path = ensure_readable(full)
        loc = location_from_name(fname)

        # Use openpyxl with data_only=True to get computed cell values
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, data_only=True)
        except Exception:
            # Fallback to pandas
            try:
                xf = pd.ExcelFile(path, engine="openpyxl")
            except Exception:
                continue
            for sh in xf.sheet_names:
                try:
                    df = pd.read_excel(path, sheet_name=sh, header=None,
                                       engine="openpyxl", dtype=str)
                except Exception:
                    continue
                _index_df(df, fname, loc, sh, corpus, row_records)
            continue

        for sh in wb.sheetnames:
            ws = wb[sh]
            # Read ALL cells into a DataFrame
            rows_data = []
            for row in ws.iter_rows(values_only=True):
                rows_data.append(list(row))
            if not rows_data:
                continue
            max_cols = max(len(r) for r in rows_data)
            for r in rows_data:
                while len(r) < max_cols:
                    r.append(None)
            df = pd.DataFrame(rows_data, dtype=str)
            # Replace 'None' strings with actual NaN for proper detection
            df = df.replace({"None": np.nan, "none": np.nan})
            _index_df(df, fname, loc, sh, corpus, row_records)
        wb.close()

    meta = {
        "total_cells": len(corpus),
        "total_files": len({x["file"] for x in corpus}),
        "total_sheets": len({(x["file"], x["sheet"]) for x in corpus}),
        "total_rows": len(row_records),
        "locations": sorted({x["location"] for x in corpus}),
    }
    return corpus, dict(row_records), meta


def _index_df(df, fname, loc, sh, corpus, row_records):
    """Index every non-empty cell in a DataFrame into the corpus."""
    cell_map, hr_set = _build_cell_col_map(df)

    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            raw = df.iat[r, c]
            if pd.isna(raw):
                continue
            v = str(raw).strip()
            if not v or v in ("nan", "None", "none", ""):
                continue
            ch = cell_map.get((r, c), f"Col_{c}")
            is_hdr = (r in hr_set)
            key = (fname, loc, sh, r)
            corpus.append({
                "file": fname,
                "location": loc,
                "sheet": sh,
                "row": r,
                "col": c,
                "col_header": ch,
                "value": v,
                "is_header": is_hdr,
            })
            if not is_hdr:
                row_records[key][ch] = v


# ─────────────────────────────────────────────────────────────────────────────
# LOCATION MATCHER
# ─────────────────────────────────────────────────────────────────────────────
def find_matching_locations(query: str, all_locs: list) -> list:
    ql = query.lower()
    matches = []
    for loc in all_locs:
        if loc.lower() in ql:
            matches.append(loc)
    if matches:
        return matches
    for loc in all_locs:
        for part in re.findall(r"[a-z]+", loc.lower()):
            if len(part) >= 3 and re.search(r"\b" + re.escape(part) + r"\b", ql):
                if loc not in matches:
                    matches.append(loc)
                break
    return matches


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/fluency/96/data-center.png", width=70)
st.sidebar.title("🏢 Capacity Tracker")
st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Source")

uploaded_files = st.sidebar.file_uploader(
    "Upload Excel files (overrides folder)",
    type=["xlsx", "xls"], accept_multiple_files=True,
)

if uploaded_files:
    file_bytes = tuple((f.name, f.read()) for f in uploaded_files)
    data_dir = save_uploads(file_bytes)
else:
    data_dir = str(EXCEL_FOLDER)

excel_files = find_excel_files(data_dir)

if not excel_files:
    st.error(
        "### ⚠️ No Excel files found\n\n"
        "**Fix:** Create `excel_files/` folder next to `app.py` "
        "and place all 10 Excel files there, then redeploy.\n\n"
        "OR upload files via the sidebar.\n\n"
        f"Looking in: `{data_dir}`"
    )
    st.stop()

loc_map = {f: location_from_name(f) for f in excel_files}
st.sidebar.success(f"✅ {len(excel_files)} file(s) found")

st.sidebar.subheader("🏙️ Location")
selected_file = st.sidebar.selectbox(
    "Location", excel_files, format_func=lambda x: loc_map[x]
)
all_sheets = load_file(os.path.join(data_dir, selected_file))

st.sidebar.subheader("📋 Sheet")
selected_sheet = st.sidebar.selectbox("Sheet", list(all_sheets.keys()))

raw_df = all_sheets[selected_sheet]
df_clean = to_numeric(smart_header(raw_df))
num_cols = df_clean.select_dtypes(include="number").columns.tolist()
cat_cols = [c for c in df_clean.columns if c not in num_cols]

st.sidebar.markdown("---")
st.sidebar.caption(
    f"📊 {len(num_cols)} numeric · {len(df_clean)} rows · {len(excel_files)} file(s)"
)

# ─────────────────────────────────────────────────────────────────────────────
# BUILD CORPUS (once, before all tabs)
# ─────────────────────────────────────────────────────────────────────────────
with st.spinner("🔍 Indexing every cell across all files…"):
    corpus, row_records, meta = build_corpus(tuple(excel_files), data_dir)

if not corpus:
    st.error(
        "⚠️ **No data indexed.**\n\n"
        f"Data dir: `{data_dir}`\n\n"
        "Upload files via the sidebar or place them in `excel_files/`."
    )
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "🏠 Overview", "📋 Raw Data", "📊 Analytics", "📈 Charts",
    "🥧 Distributions", "🔍 Query Engine", "🌍 Multi-Location",
    "🤖 AI Agent", "💬 AI Smart Query",
])
loc_label = loc_map[selected_file]


# ═══════════════════════════════════════════════════════
# TAB 0 – OVERVIEW
# ═══════════════════════════════════════════════════════
with tabs[0]:
    st.title(f"🏢 {loc_label}  ›  {selected_sheet}")
    st.caption(
        f"File: `{selected_file}` | "
        f"Raw {raw_df.shape[0]}×{raw_df.shape[1]} | "
        f"Clean {len(df_clean)}×{len(df_clean.columns)} | "
        f"Corpus: **{meta['total_cells']:,}** cells"
    )
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.markdown(f'<div class="kcard kcard-blue"><h2>{len(df_clean)}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="kcard kcard-green"><h2>{len(df_clean.columns)}</h2><p>Columns</p></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="kcard kcard-purple"><h2>{len(num_cols)}</h2><p>Numeric</p></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="kcard kcard-orange"><h2>{len(excel_files)}</h2><p>Files</p></div>', unsafe_allow_html=True)
    c5.markdown(f'<div class="kcard kcard-teal"><h2>{meta["total_cells"]:,}</h2><p>Cells Indexed</p></div>', unsafe_allow_html=True)
    c6.markdown(f'<div class="kcard kcard-red"><h2>{int(df_clean.isna().sum().sum())}</h2><p>Missing</p></div>', unsafe_allow_html=True)
    st.markdown("---")
    if num_cols:
        st.markdown('<div class="sec-title">📐 Quick Statistics</div>', unsafe_allow_html=True)
        stats = df_clean[num_cols].describe().T
        stats["range"] = stats["max"] - stats["min"]
        st.dataframe(
            stats.style.format("{:.3f}", na_rep="—")
                 .background_gradient(cmap="Blues", subset=["mean", "max"]),
            use_container_width=True)
    st.markdown('<div class="sec-title">🗂️ Column Overview</div>', unsafe_allow_html=True)
    ci = pd.DataFrame({
        "Column": df_clean.columns,
        "Type": df_clean.dtypes.values,
        "Non-Null": df_clean.notna().sum().values,
        "Null%": (df_clean.isna().mean() * 100).round(1).values,
        "Unique": [df_clean[c].nunique() for c in df_clean.columns],
        "Sample": [str(df_clean[c].dropna().iloc[0])[:55]
                   if df_clean[c].dropna().shape[0] > 0 else "—"
                   for c in df_clean.columns],
    })
    st.dataframe(ci, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 1 – RAW DATA
# ═══════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("📋 Data Table")
    srch = st.text_input("🔍 Live search", "", key="rawsrch")
    disp = (df_clean[df_clean.apply(
        lambda col: col.astype(str).str.contains(srch, case=False, na=False)
    ).any(axis=1)] if srch else df_clean)
    st.caption(f"Showing {len(disp):,} / {len(df_clean):,} rows")
    st.dataframe(disp, use_container_width=True, height=500)
    st.download_button("⬇️ Download CSV",
                       disp.to_csv(index=False).encode(), "export.csv", "text/csv")
    st.markdown("---")
    st.subheader("🗃️ Raw Excel (no header processing)")
    st.dataframe(raw_df, use_container_width=True, height=280)


# ═══════════════════════════════════════════════════════
# TAB 2 – ANALYTICS
# ═══════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("📊 Column Analytics")
    if not num_cols:
        st.info("No numeric columns in this sheet.")
    else:
        chosen = st.multiselect("Select columns", num_cols,
                                default=num_cols[:min(6, len(num_cols))])
        if chosen:
            sub = df_clean[chosen].dropna(how="all")
            kc = st.columns(min(len(chosen), 6))
            for i, col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s):
                    kc[i].metric(col[:20], f"{s.sum():,.1f}", f"avg {s.mean():,.1f}")
            st.markdown("---")
            agg_rows = []
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    grand = df_clean[chosen].select_dtypes("number").sum().sum()
                    agg_rows.append({
                        "Column": col, "Count": int(s.count()), "Sum": s.sum(),
                        "Mean": s.mean(), "Median": s.median(), "Min": s.min(),
                        "Max": s.max(), "Std": s.std(),
                        "% Total": f"{s.sum() / grand * 100:.1f}%" if grand else "—"})
            if agg_rows:
                adf = pd.DataFrame(agg_rows).set_index("Column")
                st.dataframe(
                    adf.style.format("{:,.2f}", na_rep="—",
                                     subset=[c for c in adf.columns if c != "% Total"])
                    .background_gradient(cmap="YlOrRd", subset=["Sum", "Max"]),
                    use_container_width=True)
        st.markdown("---")
        st.markdown('<div class="sec-title">🧮 Group-By</div>', unsafe_allow_html=True)
        all_cat = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique() < 60]
        if all_cat and num_cols:
            gc1, gc2, gc3 = st.columns(3)
            gc = gc1.selectbox("Group by", all_cat)
            ac = gc2.selectbox("Aggregate", num_cols)
            af = gc3.selectbox("Function", ["sum", "mean", "count", "min", "max", "median"])
            grp = (df_clean.groupby(gc)[ac].agg(af).reset_index()
                   .rename(columns={ac: f"{af}({ac})"})
                   .sort_values(f"{af}({ac})", ascending=False))
            st.dataframe(grp, use_container_width=True)
            fig = px.bar(grp, x=gc, y=f"{af}({ac})", color=f"{af}({ac})",
                         color_continuous_scale="Viridis", title=f"{af.title()} of {ac} by {gc}")
            fig.update_layout(xaxis_tickangle=-35, height=400)
            st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 3 – CHARTS
# ═══════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("📈 Interactive Charts")
    ctype = st.selectbox("Chart Type", [
        "Bar Chart", "Grouped Bar", "Line Chart", "Scatter Plot", "Area Chart",
        "Bubble Chart", "Heatmap (Correlation)", "Box Plot", "Funnel Chart",
        "Waterfall / Cumulative", "3-D Scatter"])
    if not num_cols:
        st.info("No numeric columns available.")
    else:
        def _s(label, opts, idx=0, key=None):
            return st.selectbox(label, opts, index=min(idx, max(0, len(opts) - 1)), key=key)

        if ctype == "Bar Chart":
            xc = _s("X", cat_cols or df_clean.columns.tolist(), key="bx")
            yc = _s("Y", num_cols, key="by")
            ori = st.radio("Orientation", ["Vertical", "Horizontal"], horizontal=True)
            d = df_clean[[xc, yc]].dropna()
            fig = px.bar(d, x=xc if ori == "Vertical" else yc, y=yc if ori == "Vertical" else xc,
                         color=yc, color_continuous_scale="Turbo",
                         orientation="v" if ori == "Vertical" else "h", title=f"{yc} by {xc}")
            fig.update_layout(height=480)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Grouped Bar":
            xc = _s("X", cat_cols or df_clean.columns.tolist(), key="gbx")
            ycs = st.multiselect("Y", num_cols, default=num_cols[:3])
            if ycs:
                fig = px.bar(df_clean[[xc] + ycs].dropna(subset=ycs, how="all"),
                             x=xc, y=ycs, barmode="group")
                fig.update_layout(height=460)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Line Chart":
            xc = _s("X", df_clean.columns.tolist(), key="lx")
            ycs = st.multiselect("Y", num_cols, default=num_cols[:2])
            if ycs:
                fig = px.line(df_clean[[xc] + ycs].dropna(subset=ycs, how="all"),
                              x=xc, y=ycs, markers=True)
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Scatter Plot":
            xc = _s("X", num_cols, 0, "sc_x")
            yc = _s("Y", num_cols, 1, "sc_y")
            sc = _s("Size", ["None"] + num_cols, key="sc_s")
            cc = _s("Color", ["None"] + cat_cols + num_cols, key="sc_c")
            d = df_clean.dropna(subset=[xc, yc])
            fig = px.scatter(d, x=xc, y=yc, size=sc if sc != "None" else None,
                             color=cc if cc != "None" else None,
                             color_continuous_scale="Rainbow")
            fig.update_layout(height=480)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Area Chart":
            xc = _s("X", df_clean.columns.tolist(), key="ax")
            ycs = st.multiselect("Y", num_cols, default=num_cols[:3])
            if ycs:
                fig = px.area(df_clean[[xc] + ycs].dropna(subset=ycs, how="all"), x=xc, y=ycs)
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Bubble Chart":
            if len(num_cols) >= 3:
                xc = _s("X", num_cols, 0, "bu_x")
                yc = _s("Y", num_cols, 1, "bu_y")
                sz = _s("Size", num_cols, 2, "bu_s")
                lc = _s("Color", ["None"] + cat_cols, key="bu_c")
                d = df_clean[[xc, yc, sz]].dropna()
                if lc != "None":
                    d[lc] = df_clean[lc]
                fig = px.scatter(d, x=xc, y=yc, size=sz,
                                 color=lc if lc != "None" else None,
                                 size_max=65, color_discrete_sequence=px.colors.qualitative.Vivid)
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need >= 3 numeric columns.")

        elif ctype == "Heatmap (Correlation)":
            sel = st.multiselect("Columns", num_cols, default=num_cols[:12])
            if len(sel) >= 2:
                fig = px.imshow(df_clean[sel].corr(), text_auto=".2f",
                                color_continuous_scale="RdBu_r", aspect="auto")
                fig.update_layout(height=540)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Box Plot":
            yc = _s("Value", num_cols, key="bp_v")
            xc = _s("Group", ["None"] + cat_cols, key="bp_g")
            d = df_clean[[yc] + ([xc] if xc != "None" else [])].dropna(subset=[yc])
            fig = px.box(d, y=yc, x=xc if xc != "None" else None,
                         color=xc if xc != "None" else None,
                         points="outliers",
                         color_discrete_sequence=px.colors.qualitative.Pastel)
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Funnel Chart":
            xc = _s("Stage", cat_cols or df_clean.columns.tolist(), key="fn_x")
            yc = _s("Value", num_cols, key="fn_y")
            d = (df_clean[[xc, yc]].dropna()
                 .groupby(xc)[yc].sum().reset_index()
                 .sort_values(yc, ascending=False))
            fig = px.funnel(d, x=yc, y=xc)
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Waterfall / Cumulative":
            yc = _s("Column", num_cols, key="wf_y")
            d = df_clean[yc].dropna().reset_index(drop=True)
            cum = d.cumsum()
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Value", x=d.index, y=d, marker_color="#2a5298"))
            fig.add_trace(go.Scatter(name="Cumulative", x=cum.index, y=cum,
                                     line=dict(color="#f7971e", width=2.5),
                                     mode="lines+markers"))
            fig.update_layout(title=f"Cumulative: {yc}", height=450, barmode="group")
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "3-D Scatter":
            if len(num_cols) >= 3:
                xc = _s("X", num_cols, 0, "3x")
                yc = _s("Y", num_cols, 1, "3y")
                zc = _s("Z", num_cols, 2, "3z")
                cc = _s("Color", ["None"] + cat_cols, key="3c")
                d = df_clean[[xc, yc, zc]].dropna()
                if cc != "None":
                    d[cc] = df_clean[cc]
                fig = px.scatter_3d(d, x=xc, y=yc, z=zc,
                                    color=cc if cc != "None" else None)
                fig.update_layout(height=550)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need >= 3 numeric columns.")


# ═══════════════════════════════════════════════════════
# TAB 4 – DISTRIBUTIONS
# ═══════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("🥧 Distribution Charts")
    if not num_cols:
        st.info("No numeric columns.")
    else:
        r1, r2 = st.columns(2)
        with r1:
            st.markdown('<div class="sec-title">🍕 Pie / Donut</div>', unsafe_allow_html=True)
            pc_list = [c for c in df_clean.columns if c not in num_cols and 1 < df_clean[c].nunique() <= 30]
            if pc_list:
                pc = st.selectbox("Category", pc_list, key="pcat")
                pv = st.selectbox("Value", num_cols, key="pval")
                pd_ = df_clean[[pc, pv]].copy()
                pd_[pv] = pd.to_numeric(pd_[pv], errors="coerce")
                pd_ = pd_.dropna().groupby(pc)[pv].sum().reset_index()
                fig = px.pie(pd_, names=pc, values=pv, hole=.38,
                             color_discrete_sequence=px.colors.qualitative.Vivid)
                st.plotly_chart(fig, use_container_width=True)
        with r2:
            st.markdown('<div class="sec-title">📊 Histogram</div>', unsafe_allow_html=True)
            hc = st.selectbox("Column", num_cols, key="hcol")
            bins = st.slider("Bins", 5, 100, 25)
            fig = px.histogram(df_clean[hc].dropna(), nbins=bins,
                               color_discrete_sequence=["#17a572"])
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown('<div class="sec-title">🗺️ Treemap</div>', unsafe_allow_html=True)
        tmc_list = [c for c in df_clean.columns if c not in num_cols and 1 < df_clean[c].nunique() <= 50]
        if tmc_list and num_cols:
            tmc = st.selectbox("Category", tmc_list, key="tmc")
            tmv = st.selectbox("Value", num_cols, key="tmv")
            tmd = df_clean[[tmc, tmv]].dropna().groupby(tmc)[tmv].sum().reset_index()
            tmd = tmd[tmd[tmv] > 0]
            if len(tmd):
                fig = px.treemap(tmd, path=[tmc], values=tmv, color=tmv,
                                 color_continuous_scale="Turbo")
                fig.update_layout(height=440)
                st.plotly_chart(fig, use_container_width=True)

        st.markdown('<div class="sec-title">🎻 Violin</div>', unsafe_allow_html=True)
        vc = st.selectbox("Column", num_cols, key="vc")
        fig = px.violin(df_clean[vc].dropna(), y=vc, box=True, points="outliers",
                        color_discrete_sequence=["#c0392b"])
        st.plotly_chart(fig, use_container_width=True)

        sun_cats = [c for c in df_clean.columns if c not in num_cols and 1 < df_clean[c].nunique() <= 40]
        if len(sun_cats) >= 2 and num_cols:
            st.markdown('<div class="sec-title">🌡️ Sunburst</div>', unsafe_allow_html=True)
            s1, s2, s3 = st.columns(3)
            sc1 = s1.selectbox("Level 1", sun_cats, key="sc1")
            sc2 = s2.selectbox("Level 2", [c for c in sun_cats if c != sc1], key="sc2")
            sv = s3.selectbox("Value", num_cols, key="sv")
            sd = df_clean[[sc1, sc2, sv]].dropna()
            sd[sv] = pd.to_numeric(sd[sv], errors="coerce")
            sd = sd.dropna()
            if len(sd):
                fig = px.sunburst(sd, path=[sc1, sc2], values=sv, color=sv,
                                  color_continuous_scale="RdYlGn")
                fig.update_layout(height=480)
                st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 5 – QUERY ENGINE (selected sheet)
# ═══════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("🔍 Query Engine  (selected sheet only)")
    st.info("For **cross-file search** use the **💬 AI Smart Query** tab.")
    query = st.text_input("Question",
                          placeholder="e.g. Total subscription / Max capacity / List customers")

    def run_query(q, df, nc):
        ql = q.lower()
        res = []
        if any(w in ql for w in ["sum", "total"]):
            [res.append(f"**SUM `{c}`** = {df[c].sum():,.4f}")
             for c in nc if c.lower() in ql or "all" in ql or len(nc) == 1]
        if any(w in ql for w in ["average", "mean", "avg"]):
            [res.append(f"**MEAN `{c}`** = {df[c].mean():,.4f}")
             for c in nc if c.lower() in ql or len(nc) == 1]
        if any(w in ql for w in ["maximum", "highest", "max"]):
            [res.append(f"**MAX `{c}`** = {df[c].max():,.4f}")
             for c in nc if c.lower() in ql or len(nc) == 1]
        if any(w in ql for w in ["minimum", "lowest", "min"]):
            [res.append(f"**MIN `{c}`** = {df[c].min():,.4f}")
             for c in nc if c.lower() in ql or len(nc) == 1]
        if any(w in ql for w in ["count", "how many"]):
            res.append(f"**Rows** = {len(df):,}")
        if "median" in ql:
            [res.append(f"**MEDIAN `{c}`** = {df[c].median():,.4f}")
             for c in nc if c.lower() in ql or len(nc) == 1]
        if any(w in ql for w in ["describe", "statistics", "summary"]):
            res.append("```\n" + df[nc].describe().to_string() + "\n```")
        if any(w in ql for w in ["missing", "null", "nan"]):
            ni = df.isna().sum()
            ni = ni[ni > 0]
            res.append("**Missing:**\n" + (ni.to_string() if len(ni) else "None 🎉"))
        if "unique" in ql:
            for c in df.columns:
                if c.lower() in ql:
                    u = df[c].dropna().unique()
                    res.append(f"**Unique `{c}`** ({len(u)}): {', '.join(map(str, u[:25]))}")
        if any(w in ql for w in ["customer", "list", "show"]):
            for c in df.columns:
                if "customer" in c.lower() or "name" in c.lower():
                    nm = df[c].dropna().unique()
                    res.append(f"**`{c}`** ({len(nm)}):\n" +
                               "\n".join(f"  • {n}" for n in nm[:30]))
                    break
        if not res:
            res.append("ℹ️ Try: **sum / average / max / min / count / "
                       "median / unique / missing / list**")
        return "\n\n".join(res)

    if query:
        st.markdown(run_query(query, df_clean, num_cols))
    st.markdown("---")
    st.markdown('<div class="sec-title">🧮 Manual Compute</div>', unsafe_allow_html=True)
    if num_cols:
        mc1, mc2, mc3 = st.columns(3)
        op = mc1.selectbox("Op", ["Sum", "Mean", "Max", "Min", "Count", "Median",
                                   "Std Dev", "Variance", "% of Total", "Range", "IQR"])
        sc = mc2.selectbox("Column", num_cols)
        fc = mc3.selectbox("Filter by", ["None"] + [c for c in df_clean.columns if c not in num_cols])
        fv = None
        if fc != "None":
            fv = st.selectbox("Filter value", df_clean[fc].dropna().unique().tolist())
        ds = df_clean.copy()
        if fc != "None" and fv is not None:
            ds = ds[ds[fc] == fv]
        s = ds[sc].dropna()
        ops = {"Sum": s.sum(), "Mean": s.mean(), "Max": s.max(), "Min": s.min(),
               "Count": s.count(), "Median": s.median(), "Std Dev": s.std(),
               "Variance": s.var(),
               "% of Total": f"{s.sum() / max(df_clean[sc].sum(), 1) * 100:.2f}%",
               "Range": s.max() - s.min(), "IQR": s.quantile(.75) - s.quantile(.25)}
        r = ops.get(op, "N/A")
        if isinstance(r, float):
            r = f"{r:,.4f}"
        st.success(f"**{op}** of `{sc}`"
                   f"{f' (where {fc}={fv})' if fv else ''} -> **{r}**")


# ═══════════════════════════════════════════════════════
# TAB 6 – MULTI-LOCATION
# ═══════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("🌍 Cross-Location Comparison")

    @st.cache_data(show_spinner=False)
    def load_all_summ(files, folder):
        summ = {}
        for f in files:
            shd = load_file(os.path.join(folder, f))
            for sh, raw in shd.items():
                dfc = to_numeric(smart_header(raw))
                nc = dfc.select_dtypes(include="number").columns.tolist()
                if nc:
                    summ[f"{loc_map[f]} | {sh}"] = {
                        "df": dfc, "num_cols": nc, "file": f, "sheet": sh}
        return summ

    all_summ = load_all_summ(tuple(excel_files), data_dir)
    if all_summ:
        comp_col = st.selectbox("Compare by column",
                                sorted({c for v in all_summ.values() for c in v["num_cols"]}))
        rows = []
        for lbl, info in all_summ.items():
            if comp_col in info["num_cols"]:
                s = info["df"][comp_col].dropna()
                rows.append({"Location|Sheet": lbl, "Sum": s.sum(), "Mean": s.mean(),
                             "Max": s.max(), "Min": s.min(), "Count": s.count()})
        if rows:
            cmp = pd.DataFrame(rows).set_index("Location|Sheet")
            st.dataframe(cmp.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),
                         use_container_width=True)
            fig = px.bar(cmp.reset_index(), x="Location|Sheet", y="Sum", color="Sum",
                         color_continuous_scale="Viridis",
                         title=f"Sum of '{comp_col}' across locations")
            fig.update_layout(xaxis_tickangle=-30, height=440)
            st.plotly_chart(fig, use_container_width=True)
            fig2 = px.scatter(cmp.reset_index(), x="Mean", y="Max", size="Sum",
                              text="Location|Sheet", color="Count",
                              color_continuous_scale="Turbo", title="Bubble: Mean vs Max")
            fig2.update_traces(textposition="top center")
            fig2.update_layout(height=480)
            st.plotly_chart(fig2, use_container_width=True)
            st.markdown('<div class="sec-title">🕸️ Radar Chart</div>', unsafe_allow_html=True)
            rm = st.radio("Metric", ["Sum", "Mean", "Max"], horizontal=True)
            nrm = cmp[rm] / cmp[rm].max()
            th = nrm.index.tolist()
            rv = nrm.values.tolist()
            fig3 = go.Figure(go.Scatterpolar(
                r=rv + [rv[0]], theta=th + [th[0]], fill="toself",
                fillcolor="rgba(42,82,152,.25)", line=dict(color="#2a5298", width=2)))
            fig3.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 1])), height=500)
            st.plotly_chart(fig3, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 7 – AI AGENT
# ═══════════════════════════════════════════════════════
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights")
    if st.button("🚀 Run Analysis", type="primary"):
        with st.spinner("Analysing all sheets…"):
            for lbl, info in list(all_summ.items())[:10]:
                dfa = info["df"]
                nc = info["num_cols"]
                if not nc:
                    continue
                with st.expander(f"📍 {lbl}", expanded=False):
                    ca, cb = st.columns(2)
                    with ca:
                        st.markdown("**📊 KPIs**")
                        for col in nc[:5]:
                            s = dfa[col].dropna()
                            if len(s):
                                st.metric(col[:26], f"{s.sum():,.1f}", f"avg {s.mean():,.1f}")
                    with cb:
                        st.markdown("**⚠️ Anomalies (Z > 2.5)**")
                        for col in nc[:4]:
                            s = dfa[col].dropna()
                            if len(s) > 3:
                                z = (s - s.mean()) / s.std()
                                o = z[z.abs() > 2.5]
                                (st.warning if len(o) else st.success)(
                                    f"`{col}`: {len(o)} outlier(s)" if len(o)
                                    else f"`{col}`: Clean ✓")
                    if len(nc) >= 2:
                        fig = px.bar(dfa[nc[:3]].dropna().reset_index(), x="index", y=nc[0],
                                     color_discrete_sequence=["#2a5298"], title=nc[0])
                        fig.update_layout(height=240, showlegend=False, margin=dict(t=28, b=0))
                        st.plotly_chart(fig, use_container_width=True)
    st.markdown("---")
    st.markdown('<div class="sec-title">📁 Files Summary</div>', unsafe_allow_html=True)
    fsm = []
    for f in excel_files:
        shd = all_sheets if f == selected_file else load_file(os.path.join(data_dir, f))
        fsm.append({"File": loc_map[f], "Sheets": len(shd),
                     "Rows": sum(len(s) for s in shd.values()),
                     "Cols": sum(len(s.columns) for s in shd.values())})
    fs_df = pd.DataFrame(fsm)
    st.dataframe(fs_df, use_container_width=True)
    fig = px.bar(fs_df, x="File", y="Rows", color="Sheets",
                 color_continuous_scale="Blues", title="Rows per location")
    fig.update_layout(xaxis_tickangle=-30, height=360)
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 8 – AI SMART QUERY
# Searches EVERY cell at ANY row, ANY column, ANYWHERE in ANY sheet.
# ═══════════════════════════════════════════════════════════════════════════
with tabs[8]:
    st.markdown("## 💬 AI Smart Query")
    st.markdown(
        "Type **any question** in plain English. "
        "The engine reads **every row · every column · every sheet · every file** "
        "and returns a direct answer with charts and downloadable results."
    )

    qi1, qi2, qi3, qi4, qi5 = st.columns(5)
    qi1.markdown(f'<div class="kcard kcard-blue"><h2>{meta["total_cells"]:,}</h2><p>Cells Indexed</p></div>', unsafe_allow_html=True)
    qi2.markdown(f'<div class="kcard kcard-green"><h2>{meta["total_files"]}</h2><p>Files</p></div>', unsafe_allow_html=True)
    qi3.markdown(f'<div class="kcard kcard-purple"><h2>{meta["total_sheets"]}</h2><p>Sheets</p></div>', unsafe_allow_html=True)
    qi4.markdown(f'<div class="kcard kcard-orange"><h2>{meta["total_rows"]:,}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
    qi5.markdown(f'<div class="kcard kcard-teal"><h2>{len(meta["locations"])}</h2><p>Locations</p></div>', unsafe_allow_html=True)

    with st.expander("🔧 Optional: Narrow scope before asking", expanded=False):
        scope_locs = st.multiselect("Limit to locations (blank = ALL)",
                                    meta["locations"], default=[], key="slocs")
        scope_sheets = st.multiselect("Limit to sheets (blank = ALL)",
                                      sorted({x["sheet"] for x in corpus}),
                                      default=[], key="ssheets")

    def _is_num(v):
        try:
            float(v)
            return True
        except Exception:
            return False

    # ── Core AI engine ──────────────────────────────────────────────────────
    def ai_smart_query(question: str) -> dict:
        q = question.strip()
        ql = q.lower()
        sig = [w for w in re.findall(r"[a-z0-9]{3,}", ql) if w not in _SW]

        f_sum = any(x in ql for x in ["total", "sum", "aggregate"])
        f_avg = any(x in ql for x in ["average", "mean", "avg"])
        f_max = any(x in ql for x in ["maximum", "highest", "largest", "max"])
        f_min = any(x in ql for x in ["minimum", "lowest", "smallest", "min"])
        f_cnt = any(x in ql for x in ["count", "how many", "number of"])
        f_pct = any(x in ql for x in ["percent", "percentage", "%", "ratio", "share"])
        f_stat = any(x in ql for x in ["statistics", "stats", "describe", "summary"])
        f_uniq = any(x in ql for x in ["unique", "distinct", "different"])
        f_miss = any(x in ql for x in ["missing", "null", "blank", "empty", "nan"])
        f_shts = any(x in ql for x in ["sheet", "sheets", "tab", "tabs"])
        f_cols = any(x in ql for x in ["column", "columns", "field", "header"])
        f_rows = any(x in ql for x in ["row", "rows", "record", "records"])
        f_topn = re.search(r"\btop\s*(\d+)\b", ql)
        f_botn = re.search(r"\bbottom\s*(\d+)\b", ql)
        f_num = f_sum or f_avg or f_max or f_min or f_cnt or f_stat or f_pct

        out = {"answer": "", "table": None, "chart_df": None, "chart_cfg": None,
               "cell_hits": [], "sub_tables": []}

        # Scope + location filtering
        wc = list(corpus)
        if scope_locs:
            wc = [c for c in wc if c["location"] in scope_locs]
        if scope_sheets:
            wc = [c for c in wc if c["sheet"] in scope_sheets]
        matched_locs = find_matching_locations(q, meta["locations"])
        if matched_locs:
            wc = [c for c in wc if c["location"] in matched_locs]

        if not wc:
            out["answer"] = (f"❓ No data found.\n\n"
                             f"**Available locations:** {', '.join(meta['locations'])}\n\n"
                             f"Try: *List all customers* or *Find CISCO*")
            return out

        _OP = {"total", "sum", "avg", "mean", "max", "min", "count", "list", "find", "show",
               "all", "average", "maximum", "minimum", "highest", "lowest", "top", "bottom",
               "describe", "statistics", "stats", "summary", "unique", "distinct", "sheet",
               "column", "row", "missing", "null", "percent", "percentage", "ratio", "share",
               "number", "across", "compare", "location", "locations", "customer", "customers",
               "capacity", "power", "usage", "rack", "space", "subscription", "billing"}

        def best_col_kw(sl):
            cands = [w for w in sl if w not in _OP and len(w) >= 3]
            best, bn = None, 0
            for w in cands:
                n = sum(1 for c in wc if w in c["col_header"].lower())
                if n > bn:
                    bn, best = n, w
            # If no match via col_header, try matching against column values
            if not best:
                for w in cands:
                    n = sum(1 for c in wc if w in c["value"].lower())
                    if n > bn:
                        bn, best = n, w
            return best

        def _match_col_header(kw, header):
            """Flexible column header matching."""
            hl = header.lower()
            kwl = kw.lower()
            if kwl in hl:
                return True
            # Handle common abbreviation/synonyms
            synonyms = {
                "subscription": ["subscription", "subscribed", "subscript"],
                "capacity": ["capacity", "capac"],
                "power": ["power", "kw", "kva"],
                "usage": ["usage", "utilization", "consumption", "consumed"],
                "rack": ["rack", "racks"],
                "space": ["space", "sqft", "sq ft", "sq. ft"],
                "customer": ["customer", "name"],
                "billing": ["billing", "bill"],
            }
            for key, syns in synonyms.items():
                if kwl in syns or kwl == key:
                    for s in syns:
                        if s in hl:
                            return True
            return False

        def num_pairs_for_kw(kw):
            """Find numeric values for cells whose column header matches kw."""
            res = []
            for cell in wc:
                if cell["is_header"]:
                    continue
                if _match_col_header(kw, cell["col_header"]):
                    try:
                        res.append((float(cell["value"]), cell))
                    except ValueError:
                        pass
            return res

        def num_pairs_for_any_sig(sig_words):
            """Try each significant word and return the best-matching set."""
            best_kw = None
            best_pairs = []
            for w in sig_words:
                if w in _OP:
                    continue
                pairs = num_pairs_for_kw(w)
                if len(pairs) > len(best_pairs):
                    best_pairs = pairs
                    best_kw = w
            return best_kw, best_pairs

        def build_rows_df(keys):
            recs = []
            for key in keys:
                rec = row_records.get(key, {})
                if rec:
                    rd = {"📍 Location": key[1], "📋 Sheet": key[2], "Row #": key[3] + 1}
                    rd.update(rec)
                    recs.append(rd)
            return pd.DataFrame(recs) if recs else pd.DataFrame()

        loc_str = " in " + ", ".join(matched_locs) if matched_locs else ""

        # INTENT: Sheet listing
        if f_shts and not f_num:
            seen = set()
            srows = []
            for cell in wc:
                k = (cell["location"], cell["sheet"])
                if k not in seen:
                    seen.add(k)
                    dr = sum(1 for rk in row_records
                             if rk[1] == cell["location"] and rk[2] == cell["sheet"])
                    srows.append({"Location": cell["location"], "Sheet": cell["sheet"],
                                  "File": cell["file"], "Data Rows": dr})
            tbl = pd.DataFrame(srows)
            out["answer"] = f"Found **{len(tbl)}** sheet(s){loc_str}."
            out["table"] = tbl
            return out

        # INTENT: Missing values
        if f_miss:
            mrows = []
            for fname in excel_files:
                loc = location_from_name(fname)
                if matched_locs and loc not in matched_locs:
                    continue
                shd = load_file(os.path.join(data_dir, fname))
                for sh, raw in shd.items():
                    if scope_sheets and sh not in scope_sheets:
                        continue
                    dfc = to_numeric(smart_header(raw))
                    for col in dfc.columns:
                        mc = int(dfc[col].isna().sum())
                        if mc > 0:
                            mrows.append({"Location": loc, "Sheet": sh, "Column": col,
                                          "Missing": mc,
                                          "Missing%": f"{mc / max(len(dfc), 1) * 100:.1f}%"})
            if mrows:
                tbl = pd.DataFrame(mrows).sort_values("Missing", ascending=False)
                out["answer"] = f"Found **{len(tbl)}** column(s) with missing values{loc_str}."
                out["table"] = tbl
            else:
                out["answer"] = "✅ No missing values found."
            return out

        # INTENT: Column listing
        if f_cols and not f_num:
            kw = best_col_kw(sig)
            seen = set()
            crows = []
            for cell in wc:
                if not cell["is_header"]:
                    continue
                ch = cell["col_header"].strip()
                if ch in ("", "nan"):
                    continue
                if kw and not _match_col_header(kw, ch):
                    continue
                k = (cell["location"], cell["sheet"], ch)
                if k not in seen:
                    seen.add(k)
                    crows.append({"Location": cell["location"],
                                  "Sheet": cell["sheet"], "Column": ch})
            tbl = pd.DataFrame(crows) if crows else pd.DataFrame()
            out["answer"] = f"Found **{len(tbl)}** column(s){loc_str}."
            out["table"] = tbl
            return out

        # INTENT: Row count
        if f_rows and f_cnt and not sig:
            crows = []
            total = 0
            for fname in excel_files:
                loc = location_from_name(fname)
                if matched_locs and loc not in matched_locs:
                    continue
                shd = load_file(os.path.join(data_dir, fname))
                for sh, raw in shd.items():
                    if scope_sheets and sh not in scope_sheets:
                        continue
                    dfc = smart_header(raw)
                    crows.append({"Location": loc, "Sheet": sh, "Rows": len(dfc)})
                    total += len(dfc)
            tbl = pd.DataFrame(crows)
            out["answer"] = f"**{total:,}** total data rows across **{len(tbl)}** sheet(s){loc_str}."
            out["table"] = tbl
            out["chart_df"] = tbl.groupby("Location")["Rows"].sum().reset_index()
            out["chart_cfg"] = {"x": "Location", "y": "Rows", "title": "Data Rows per Location"}
            return out

        # INTENT: Numeric aggregation
        if f_num:
            # Try finding by significant keywords first
            kw, pairs = num_pairs_for_any_sig(sig)

            # If no match from sig words, try matching common domain keywords
            if not pairs:
                domain_kws = ["subscription", "capacity", "power", "usage", "rack",
                              "space", "consumption", "kw", "kva", "sqft"]
                for dkw in domain_kws:
                    if dkw in ql:
                        pairs = num_pairs_for_kw(dkw)
                        if pairs:
                            kw = dkw
                            break

            if pairs:
                vals = [v for v, _ in pairs]
                s_all = pd.Series(vals)
                parts = []
                if f_sum or f_stat:
                    parts.append(f"**Total (Sum):**    {s_all.sum():,.4f}")
                if f_avg or f_stat:
                    parts.append(f"**Average (Mean):** {s_all.mean():,.4f}")
                if f_max or f_stat:
                    parts.append(f"**Maximum:**        {s_all.max():,.4f}")
                if f_min or f_stat:
                    parts.append(f"**Minimum:**        {s_all.min():,.4f}")
                if f_cnt or f_stat:
                    parts.append(f"**Count:**          {s_all.count():,}")
                if f_stat:
                    parts.append(f"**Median:** {s_all.median():,.4f}  |  "
                                 f"**Std Dev:** {s_all.std():,.4f}")
                if f_pct:
                    grand = sum(float(c["value"]) for c in wc
                                if not c["is_header"] and _is_num(c["value"]))
                    pct = s_all.sum() / grand * 100 if grand else 0
                    parts.append(f"**% of all numeric:** {pct:.2f}%")
                grp = defaultdict(list)
                for v, cell in pairs:
                    grp[f"{cell['location']} | {cell['sheet']}"].append(v)
                breakdown = []
                for lbl2, vs in grp.items():
                    sv = pd.Series(vs)
                    breakdown.append({"Location | Sheet": lbl2, "Count": sv.count(),
                                      "Sum": round(sv.sum(), 4), "Mean": round(sv.mean(), 4),
                                      "Max": round(sv.max(), 4), "Min": round(sv.min(), 4)})
                tbl = pd.DataFrame(breakdown).sort_values("Sum", ascending=False)
                out["answer"] = (f"Results for columns matching **'{kw}'** "
                                 f"({len(vals):,} values{loc_str}):\n\n" + "\n".join(parts))
                out["table"] = tbl
                out["chart_df"] = tbl
                out["chart_cfg"] = {"x": "Location | Sheet", "y": "Sum",
                                    "title": f"Sum of '{kw}' by Location/Sheet"}
                if f_topn:
                    n = int(f_topn.group(1))
                    top = sorted(pairs, key=lambda x: x[0], reverse=True)[:n]
                    out["sub_tables"].append({
                        "label": f"🏆 Top {n} values for '{kw}'",
                        "df": pd.DataFrame([{"📍 Location": c["location"], "📋 Sheet": c["sheet"],
                                             "Row #": c["row"] + 1, "Col #": c["col"] + 1,
                                             "Column": c["col_header"], "Value": v} for v, c in top])})
                if f_botn:
                    n = int(f_botn.group(1))
                    bot = sorted(pairs, key=lambda x: x[0])[:n]
                    out["sub_tables"].append({
                        "label": f"🔻 Bottom {n}",
                        "df": pd.DataFrame([{"📍 Location": c["location"], "📋 Sheet": c["sheet"],
                                             "Row #": c["row"] + 1, "Col #": c["col"] + 1,
                                             "Column": c["col_header"], "Value": v} for v, c in bot])})
                return out

            # If still no numeric match, try ALL numeric cells
            if f_sum or f_avg or f_max or f_min or f_cnt:
                all_nums = []
                for cell in wc:
                    if cell["is_header"]:
                        continue
                    try:
                        all_nums.append((float(cell["value"]), cell))
                    except ValueError:
                        pass
                if all_nums:
                    vals = [v for v, _ in all_nums]
                    s_all = pd.Series(vals)
                    parts = []
                    if f_sum:
                        parts.append(f"**Total (Sum) of ALL numeric cells:** {s_all.sum():,.4f}")
                    if f_avg:
                        parts.append(f"**Average of ALL numeric cells:** {s_all.mean():,.4f}")
                    if f_max:
                        parts.append(f"**Maximum across ALL numeric cells:** {s_all.max():,.4f}")
                    if f_min:
                        parts.append(f"**Minimum across ALL numeric cells:** {s_all.min():,.4f}")
                    if f_cnt:
                        parts.append(f"**Count of ALL numeric cells:** {s_all.count():,}")
                    out["answer"] = (f"No specific column matched your keywords. "
                                     f"Showing aggregation across **all numeric cells**{loc_str}:\n\n"
                                     + "\n".join(parts))
                    return out

        # INTENT: Unique values
        if f_uniq and sig:
            kw = best_col_kw(sig)
            if kw:
                uvals = set()
                srows = []
                for cell in wc:
                    if cell["is_header"]:
                        continue
                    if _match_col_header(kw, cell["col_header"]):
                        uvals.add(cell["value"])
                        srows.append({"Location": cell["location"], "Sheet": cell["sheet"],
                                      "Column": cell["col_header"], "Value": cell["value"]})
                tbl = (pd.DataFrame(srows).drop_duplicates(subset=["Location", "Sheet", "Value"])
                       if srows else pd.DataFrame())
                out["answer"] = f"Found **{len(uvals)}** unique value(s) in columns matching **'{kw}'**{loc_str}."
                out["table"] = tbl
                return out

        # INTENT: Free-text entity/keyword search
        if sig:
            quoted = re.findall(r'"([^"]+)"', q)
            if quoted:
                terms = [quoted[0].lower()]
            else:
                # Build terms: try each significant word against cell values
                terms = []
                for w in sig:
                    hits = sum(1 for cell in wc if w in cell["value"].lower())
                    if hits > 0:
                        terms.append(w)
                if not terms:
                    terms = sig

            hit_cells = [cell for cell in wc if not cell["is_header"]
                         and any(t in cell["value"].lower() for t in terms)]
            hit_keys = {(c["file"], c["location"], c["sheet"], c["row"]) for c in hit_cells}
            full_df = build_rows_df(hit_keys)

            loc_freq = defaultdict(int)
            sh_freq = defaultdict(int)
            for c in hit_cells:
                loc_freq[c["location"]] += 1
                sh_freq[f"{c['location']} | {c['sheet']}"] += 1

            lf_df = (pd.DataFrame(list(loc_freq.items()), columns=["Location", "Hits"])
                     .sort_values("Hits", ascending=False))
            cell_list = [{"📍 Location": c["location"], "📋 Sheet": c["sheet"],
                          "Row #": c["row"] + 1, "Col #": c["col"] + 1,
                          "Column Header": c["col_header"], "Value": c["value"]}
                         for c in hit_cells[:60]]

            out["answer"] = (f"Found **{len(hit_cells):,}** cell(s) matching "
                             f"**'{', '.join(terms[:4])}'**{loc_str}\n"
                             f"across **{len(loc_freq)}** location(s) and "
                             f"**{len(sh_freq)}** sheet(s).\n"
                             f"**{len(hit_keys):,}** unique data row(s) contain this data.")
            out["table"] = full_df if not full_df.empty else None
            out["cell_hits"] = cell_list
            if len(loc_freq) > 1:
                out["chart_df"] = lf_df
                out["chart_cfg"] = {"x": "Location", "y": "Hits",
                                    "title": f"Hits per location — '{', '.join(terms[:3])}'"}
            if not full_df.empty:
                for col in full_df.columns:
                    if "customer" in col.lower() or "name" in col.lower():
                        cust_df = full_df[["📍 Location", "📋 Sheet", col]].drop_duplicates()
                        out["sub_tables"].append({
                            "label": f"👤 Customer Names ({len(cust_df)} rows)", "df": cust_df})
                        break
            num_hits = []
            for c in hit_cells:
                try:
                    num_hits.append((float(c["value"]), c))
                except Exception:
                    pass
            if num_hits and f_topn:
                n = int(f_topn.group(1))
                top = sorted(num_hits, key=lambda x: x[0], reverse=True)[:n]
                out["sub_tables"].append({
                    "label": f"🏆 Top {n} numeric values",
                    "df": pd.DataFrame([{"Location": c["location"], "Sheet": c["sheet"],
                                         "Row #": c["row"] + 1, "Col": c["col_header"], "Value": v}
                                        for v, c in top])})
            return out

        # Fallback
        out["answer"] = ("❓ No match found.\n\n**Try:**\n"
                         "• *List all customers*  |  *List all customers in Noida*\n"
                         "• *Find CISCO*  |  *Find AT&T*  |  *Find Axis Bank*  |  *Find MOTMOT*\n"
                         "• *Total subscription*  |  *Max capacity Airoli*  |  "
                         "*Average power usage*  |  *Top 10 subscription*\n"
                         "• *Unique billing models*  |  *Show missing values*  |  *Show all sheets*")
        return out

    # ── Render ──────────────────────────────────────────────────────────────
    def render_answer(res: dict):
        st.markdown(f'<div class="ans-box">{res["answer"]}</div>', unsafe_allow_html=True)
        if res.get("table") is not None and not res["table"].empty:
            tbl = res["table"].reset_index(drop=True)
            st.dataframe(tbl, use_container_width=True, height=min(520, 48 + len(tbl) * 36))
            st.download_button("⬇️ Download result CSV", tbl.to_csv(index=False).encode(),
                               "ai_result.csv", "text/csv", key=f"dl_{id(res)}")
        if res.get("chart_cfg") and res.get("chart_df") is not None:
            cfg = res["chart_cfg"]
            cdf = res["chart_df"]
            if cfg["x"] in cdf.columns and cfg["y"] in cdf.columns:
                fig = px.bar(cdf.sort_values(cfg["y"], ascending=False).head(30),
                             x=cfg["x"], y=cfg["y"], color=cfg["y"],
                             color_continuous_scale="Viridis", title=cfg["title"], height=400)
                fig.update_layout(xaxis_tickangle=-30)
                st.plotly_chart(fig, use_container_width=True)
        for si in res.get("sub_tables", []):
            with st.expander(si["label"], expanded=True):
                st.dataframe(si["df"], use_container_width=True)
        if res.get("cell_hits"):
            with st.expander(f"🔬 Cell-level matches — first {len(res['cell_hits'])} shown",
                             expanded=False):
                for ch in res["cell_hits"]:
                    st.markdown(f'<div class="cell-chip">'
                                f'📍 <b>{ch["📍 Location"]}</b> → '
                                f'Sheet: <b>{ch["📋 Sheet"]}</b> | '
                                f'Row <b>{ch["Row #"]}</b> Col <b>{ch["Col #"]}</b> | '
                                f'Header: <i>{ch["Column Header"]}</i> | '
                                f'Value: <b>{ch["Value"]}</b>'
                                f'</div>', unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)

    # ── Chat history ─────────────────────────────────────────────────────────
    if "aisq_hist" not in st.session_state:
        st.session_state.aisq_hist = []
    for turn in st.session_state.aisq_hist:
        st.markdown(f'<div class="q-user">🧑 {turn["q"]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)
        render_answer(turn["res"])
        st.markdown("---")

    # ── Input bar ────────────────────────────────────────────────────────────
    st.markdown("---")
    ic, bc, cc = st.columns([8, 1, 1])
    with ic:
        user_q = st.text_input("Ask:",
                               placeholder="Find CISCO | List all customers in Noida | "
                                           "Total subscription Airoli | Max capacity | Top 10 subscription",
                               label_visibility="collapsed", key="aisq_input")
    with bc:
        ask_btn = st.button("🔍 Ask", use_container_width=True, type="primary")
    with cc:
        if st.button("🗑️ Clear", use_container_width=True):
            st.session_state.aisq_hist = []
            st.rerun()

    # ── Example chips ─────────────────────────────────────────────────────────
    st.markdown("**💡 Click any example to ask instantly:**")
    examples = [
        ["List all customers", "List all customers in Noida",
         "List all customers in Bangalore", "List all customers in Chennai",
         "List all customers in Vashi", "List all customers in Kolkata"],
        ["Total subscription across all locations", "Total capacity in Airoli",
         "Average power usage", "Maximum rack subscription",
         "Minimum capacity in Kolkata", "Top 10 subscription values"],
        ["Find CISCO", "Find Axis Bank", "Find MOTMOT",
         "Find AT&T", "Find Zscaler", "Find Edge Network"],
        ["Show all sheets", "Show missing values", "Count rows per sheet",
         "Unique billing models", "Statistics of subscription", "Average usage in KW"],
    ]
    for row in examples:
        cols = st.columns(len(row))
        for j, ex in enumerate(row):
            if cols[j].button(ex, key=f"chip_{ex}", use_container_width=True):
                user_q = ex
                ask_btn = True

    if ask_btn and user_q.strip():
        with st.spinner(f"Scanning every cell for: **{user_q}** …"):
            answer = ai_smart_query(user_q)
        st.session_state.aisq_hist.append({"q": user_q, "res": answer})
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    f"Sify DC · Customer & Capacity Tracker · "
    f"{meta['total_cells']:,} cells indexed · "
    f"Locations: {', '.join(meta['locations'])} · "
    f"Data: `{data_dir}`"
)
