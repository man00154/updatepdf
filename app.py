import os
import re
import warnings
import tempfile
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats as scipy_stats

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════
# PAGE CONFIG & CSS
# ═══════════════════════════════════════════════════════
st.set_page_config(
    page_title="Sify DC – Capacity Intelligence",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0e1a,#1a2035,#0d1b2a)!important;}
[data-testid="stSidebar"] *{color:#c9d8f0!important;}
[data-testid="stAppViewContainer"]{background:#0a0e1a;}
[data-testid="stHeader"]{background:rgba(10,14,26,0.95);}
.main .block-container{padding-top:1.5rem;}
.kcard{border-radius:14px;padding:18px 22px;color:#fff;margin-bottom:12px;
       box-shadow:0 6px 24px rgba(0,0,0,.45);transition:transform .2s;}
.kcard:hover{transform:translateY(-3px);}
.kcard h2{font-size:2rem;margin:0;font-weight:800;letter-spacing:-0.5px;}
.kcard p{margin:4px 0 0;font-size:.84rem;opacity:.80;letter-spacing:.3px;}
.kcard-blue{background:linear-gradient(135deg,#1e3c72,#2a5298);}
.kcard-green{background:linear-gradient(135deg,#0b6e4f,#17a572);}
.kcard-red{background:linear-gradient(135deg,#7b1a1a,#c0392b);}
.kcard-orange{background:linear-gradient(135deg,#7d4e00,#e67e22);}
.kcard-teal{background:linear-gradient(135deg,#0f3460,#16213e);}
.kcard-purple{background:linear-gradient(135deg,#4a0072,#7b1fa2);}
.kcard-cyan{background:linear-gradient(135deg,#006080,#00a8cc);}
.kcard-pink{background:linear-gradient(135deg,#6b0040,#c0166a);}
.sec-title{font-size:1.18rem;font-weight:700;color:#5ec1f0;
           border-left:5px solid #2a5298;padding-left:12px;margin:18px 0 10px;}
.q-user{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff;
        border-radius:18px 18px 4px 18px;padding:10px 16px;
        margin:10px 0 4px auto;max-width:76%;width:fit-content;
        box-shadow:0 3px 12px rgba(30,60,114,.45);float:right;clear:both;}
.ans-box{background:linear-gradient(135deg,#0f2744,#1a4a6b);color:#d0ecff;
         border-radius:12px;padding:16px 20px;margin:8px 0;font-size:.97rem;
         box-shadow:0 4px 16px rgba(0,0,0,.4);white-space:pre-wrap;line-height:1.7;
         border-left:4px solid #2a5298;}
.cell-chip{background:#0f2010;border-left:4px solid #27ae60;border-radius:6px;
           padding:7px 14px;margin:3px 0;font-family:monospace;font-size:.82rem;color:#9fffac;}
.col-chip{background:#0d1f3c;border-left:4px solid #2a5298;border-radius:6px;
          padding:5px 12px;margin:2px 0;font-family:monospace;font-size:.80rem;color:#a8c8f0;
          display:inline-block;margin-right:4px;margin-bottom:4px;}
.clearfix{clear:both;}
.stat-box{background:#1a2035;border-radius:10px;padding:14px 18px;border:1px solid #2a3a5a;margin:6px 0;}
h1,h2,h3{color:#c9d8f0!important;}
p,li{color:#c9d8f0!important;}
.stTabs [data-baseweb="tab-list"]{background:#1a2035;border-radius:10px;padding:4px;}
.stTabs [data-baseweb="tab"]{color:#7da8d0!important;border-radius:8px;}
.stTabs [aria-selected="true"]{background:#2a5298!important;color:#fff!important;}
.excel-header{background:linear-gradient(135deg,#162a4a,#1e3c72);border:1px solid #2a4a7a;
              border-radius:8px;padding:10px 14px;margin:4px 2px;text-align:center;
              font-weight:700;font-size:.78rem;color:#90c8f0;letter-spacing:.3px;}
.excel-val{background:#0a1020;border:1px solid #1a2a3a;border-radius:4px;padding:4px 8px;
           font-size:.76rem;color:#c9d8f0;text-align:center;}
.loc-badge{background:linear-gradient(135deg,#2a5298,#1e3c72);border-radius:20px;
           padding:4px 12px;font-size:.80rem;font-weight:700;color:#fff;
           display:inline-block;margin:2px;}
</style>
""",
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════
# STOP WORDS
# ═══════════════════════════════════════════════════════
_SW = {
    "the","and","for","are","all","any","how","what","show","give",
    "tell","from","this","that","with","get","find","list","much",
    "many","each","every","data","value","values","number","numbers",
    "in","of","a","an","is","at","by","to","do","me","my","about",
    "details","info","please","can","you","per","across","which",
    "where","who","when","does","did","have","has","their","its",
    "our","your","there","these","those","been","will","would","could",
    "should","shall","let","some","just","also","even","only","into",
    "over","under","both","such","than","then","but","not","nor",
    "yet","so","either","neither","versus","vs",
}

# ═══════════════════════════════════════════════════════
# PATH HELPERS
# ═══════════════════════════════════════════════════════
def _app_dir():
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path(os.getcwd())


EXCEL_FOLDER = _app_dir() / "excel_files"


def find_excel_files(folder):
    p = Path(folder)
    if not p.is_dir():
        return []
    return sorted(
        f.name
        for f in p.iterdir()
        if f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~")
    )


def location_from_name(fname):
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


# ═══════════════════════════════════════════════════════
# UPLOAD HANDLER
# ═══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def save_uploads(file_bytes_tuple):
    tmp = tempfile.mkdtemp()
    for name, data in file_bytes_tuple:
        with open(os.path.join(tmp, name), "wb") as fh:
            fh.write(data)
    return tmp


# ═══════════════════════════════════════════════════════
# READ ONE SHEET – openpyxl data_only + xlrd fallback
# ═══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def _read_sheet_openpyxl(path, sheet_name):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    mr = ws.max_row or 0
    mc = ws.max_column or 0
    if mr == 0:
        wb.close()
        return pd.DataFrame()
    real_mc = 0
    samples = sorted(set(
        list(range(1, min(31, mr + 1))) + list(range(max(1, mr - 9), mr + 1))
    ))
    for r in samples:
        for cell in ws[r]:
            if cell.value is not None:
                real_mc = max(real_mc, cell.column)
    if real_mc == 0:
        wb.close()
        return pd.DataFrame()
    cap = min(real_mc + 2, mc)
    rows = []
    for row in ws.iter_rows(min_row=1, max_row=mr, max_col=cap, values_only=True):
        rows.append(list(row))
    wb.close()
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows, dtype=str)
    df = df.replace({"None": np.nan, "none": np.nan})
    return df


@st.cache_data(show_spinner=False)
def _read_sheet_xlrd(path, sheet_name):
    import xlrd
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet_name)
    rows = []
    for r in range(ws.nrows):
        row = []
        for c in range(ws.ncols):
            try:
                cell = ws.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    row.append(np.nan)
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    v = cell.value
                    row.append(str(int(v)) if v == int(v) else str(v))
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    import xlrd.xldate
                    dt = xlrd.xldate.xldate_as_datetime(cell.value, wb.datemode)
                    row.append(str(dt.date()))
                else:
                    row.append(str(cell.value).strip())
            except Exception:
                row.append(np.nan)
        rows.append(row)
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows, dtype=str)
    df = df.replace({"None": np.nan, "none": np.nan, "nan": np.nan})
    return df


@st.cache_data(show_spinner=False)
def load_file(original_path):
    path = str(original_path)
    sheets = {}
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, data_only=True)
            names = wb.sheetnames
            wb.close()
            for sh in names:
                try:
                    df = _read_sheet_openpyxl(path, sh)
                    if not df.empty:
                        sheets[sh] = df
                except Exception:
                    pass
        except Exception as e:
            st.sidebar.warning(f"⚠️ xlsx error {os.path.basename(path)}: {e}")
    else:
        try:
            import xlrd
            wb = xlrd.open_workbook(path)
            names = wb.sheet_names()
            for sh in names:
                try:
                    df = _read_sheet_xlrd(path, sh)
                    if not df.empty:
                        sheets[sh] = df
                except Exception:
                    pass
        except Exception:
            try:
                from openpyxl import load_workbook
                wb = load_workbook(path, data_only=True)
                names = wb.sheetnames
                wb.close()
                for sh in names:
                    try:
                        df = _read_sheet_openpyxl(path, sh)
                        if not df.empty:
                            sheets[sh] = df
                    except Exception:
                        pass
            except Exception as e2:
                st.sidebar.warning(f"⚠️ {os.path.basename(path)}: {e2}")
    return sheets


# ═══════════════════════════════════════════════════════
# HEADER DETECTION & SMART HEADER
# ═══════════════════════════════════════════════════════
def best_header_row(df):
    best_row, best_score = 0, -1
    for i in range(min(10, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        filled = (row.str.len() > 0) & (~row.isin(["nan", "None", ""]))
        label = filled & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        score = label.sum() * 2 + filled.sum()
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def smart_header(df):
    hr = best_header_row(df)
    hdr = df.iloc[hr].fillna("").astype(str).str.strip()
    seen = {}
    cols = []
    for col in hdr:
        col = col if col and col not in ("nan", "None") else f"Col_{len(cols)}"
        if col in seen:
            seen[col] += 1
            cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            cols.append(col)
    data = df.iloc[hr + 1:].copy()
    data.columns = cols
    return data.dropna(how="all").reset_index(drop=True)


def to_numeric(df):
    out = df.copy()
    for col in out.columns:
        try:
            converted = pd.to_numeric(out[col], errors="coerce")
            if converted.notna().sum() > out[col].notna().sum() * 0.5:
                out[col] = converted
            else:
                pass
        except Exception:
            pass
    return out


# ═══════════════════════════════════════════════════════
# MULTI-SECTION HEADER MAP
# ═══════════════════════════════════════════════════════
def _detect_all_header_rows(df):
    hr_set = set()
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.strip()
        fm = (row.str.len() > 0) & (~row.isin(["nan", "None", ""]))
        fv = row[fm]
        nf = fv.shape[0]
        if nf < 2:
            continue
        lm = fm & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        nl = lm.sum()
        nu = fv.nunique()
        lr = nl / max(nf, 1)
        ur = nu / max(nf, 1)
        vc = fv.value_counts()
        nr = (vc > 1).sum()
        if lr >= 0.80 and ur >= 0.75 and nr <= max(2, nf * 0.15) and nu >= 3:
            hr_set.add(i)
        elif nf <= 10 and nf >= 2 and lr >= 0.90 and ur >= 0.80 and nr <= 1:
            hr_set.add(i)
    return hr_set


def _build_cell_col_map(df):
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


# ═══════════════════════════════════════════════════════
# INDEX A SINGLE SHEET
# ═══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def index_single_sheet(file_path, sheet_name):
    sheets = load_file(file_path)
    if sheet_name not in sheets:
        return [], {}, {"total_cells": 0, "total_rows": 0, "total_data": 0, "total_headers": 0}
    df = sheets[sheet_name]
    cell_map, hr_set = _build_cell_col_map(df)
    cells = []
    row_recs = {}
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            raw = df.iat[r, c]
            if pd.isna(raw):
                continue
            v = str(raw).strip()
            if not v or v in ("nan", "None", "none", ""):
                continue
            ch = cell_map.get((r, c), f"Col_{c}")
            is_hdr = r in hr_set
            cells.append({
                "row": r, "col": c, "col_header": ch,
                "value": v, "is_header": is_hdr,
            })
            if not is_hdr:
                if r not in row_recs:
                    row_recs[r] = {}
                row_recs[r][ch] = v
    meta = {
        "total_cells": len(cells),
        "total_rows": len(row_recs),
        "total_data": sum(1 for x in cells if not x["is_header"]),
        "total_headers": sum(1 for x in cells if x["is_header"]),
    }
    return cells, row_recs, meta


# ═══════════════════════════════════════════════════════
# BUILD FULL CORPUS
# ═══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def build_corpus(file_list, folder):
    corpus = []
    row_records = defaultdict(dict)
    for fname in file_list:
        full = os.path.join(folder, fname)
        if not os.path.isfile(full):
            continue
        loc = location_from_name(fname)
        sheets = load_file(full)
        for sh, df in sheets.items():
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
                    is_hdr = r in hr_set
                    key = (fname, loc, sh, r)
                    corpus.append({
                        "file": fname, "location": loc, "sheet": sh,
                        "row": r, "col": c, "col_header": ch,
                        "value": v, "is_header": is_hdr,
                    })
                    if not is_hdr:
                        row_records[key][ch] = v
    meta = {
        "total_cells": len(corpus),
        "total_files": len({x["file"] for x in corpus}),
        "total_sheets": len({(x["file"], x["sheet"]) for x in corpus}),
        "total_rows": len(row_records),
        "locations": sorted({x["location"] for x in corpus}),
    }
    return corpus, dict(row_records), meta


# ═══════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════
st.sidebar.markdown("## 🏢 Sify DC Intelligence")
st.sidebar.markdown("**Capacity & Customer Tracker**")
st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Source")

uploaded_files = st.sidebar.file_uploader(
    "Upload Excel files (optional)", type=["xlsx", "xls"], accept_multiple_files=True
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
        "The `excel_files/` folder is empty or missing. Upload files via the sidebar."
    )
    st.stop()

loc_map = {f: location_from_name(f) for f in excel_files}
st.sidebar.success(f"✅ {len(excel_files)} file(s) loaded")

st.sidebar.subheader("🏙️ Location")
selected_file = st.sidebar.selectbox(
    "Location", excel_files, format_func=lambda x: loc_map[x]
)
all_sheets = load_file(os.path.join(data_dir, selected_file))

if not all_sheets:
    st.error(f"⚠️ Could not read any sheets from `{selected_file}`.")
    st.stop()

st.sidebar.subheader("📋 Sheet")
selected_sheet = st.sidebar.selectbox("Sheet", list(all_sheets.keys()))

raw_df = all_sheets[selected_sheet]
df_clean = to_numeric(smart_header(raw_df))
num_cols = df_clean.select_dtypes(include="number").columns.tolist()
cat_cols = [c for c in df_clean.columns if c not in num_cols]

st.sidebar.markdown("---")
st.sidebar.caption(
    f"📊 **{len(num_cols)}** numeric · **{len(df_clean)}** rows · **{len(excel_files)}** file(s)"
)

# Build indexes
with st.spinner("🔍 Indexing all data across every file, sheet, row and column…"):
    corpus, row_records, meta = build_corpus(tuple(excel_files), data_dir)

if not corpus:
    st.error("⚠️ **No data indexed.** Upload files via the sidebar.")
    st.stop()

with st.spinner("🔍 Indexing selected sheet…"):
    sq_cells, sq_rows, sq_meta = index_single_sheet(
        os.path.join(data_dir, selected_file), selected_sheet
    )

# ═══════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════
tabs = st.tabs([
    "🏠 Overview",
    "📋 Raw Data",
    "📊 Analytics",
    "📈 Charts",
    "🥧 Distributions",
    "🔍 Query Engine",
    "🌍 Multi-Location",
    "🤖 AI Agent",
    "💬 Smart Query",
])
loc_label = loc_map[selected_file]


# ═══════════════════════════════════════════════════════
# TAB 0 – OVERVIEW  (shows data format exactly like Excel screenshot)
# ═══════════════════════════════════════════════════════
with tabs[0]:
    st.title(f"🏢 {loc_label}  ›  {selected_sheet}")
    st.caption(
        f"File: `{selected_file}` | "
        f"Raw {raw_df.shape[0]}×{raw_df.shape[1]} | "
        f"Clean {len(df_clean)}×{len(df_clean.columns)} | "
        f"Corpus: **{meta['total_cells']:,}** cells across **{meta['total_files']}** files"
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.markdown(f'<div class="kcard kcard-blue"><h2>{len(df_clean)}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="kcard kcard-green"><h2>{len(df_clean.columns)}</h2><p>Columns</p></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="kcard kcard-purple"><h2>{len(num_cols)}</h2><p>Numeric Cols</p></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="kcard kcard-orange"><h2>{len(excel_files)}</h2><p>Excel Files</p></div>', unsafe_allow_html=True)
    c5.markdown(f'<div class="kcard kcard-teal"><h2>{sq_meta["total_cells"]:,}</h2><p>Sheet Cells</p></div>', unsafe_allow_html=True)
    c6.markdown(f'<div class="kcard kcard-red"><h2>{int(df_clean.isna().sum().sum())}</h2><p>Missing</p></div>', unsafe_allow_html=True)

    st.markdown("---")

    # ─── COLUMN FORMAT DISPLAY (exact Excel header structure) ───
    st.markdown('<div class="sec-title">📋 Column Format – All Files</div>', unsafe_allow_html=True)
    st.caption("Shows the exact column header structure from each Excel file as seen in the attached format")

    @st.cache_data(show_spinner=False)
    def get_all_column_headers(files, folder):
        result = {}
        for f in files:
            full = os.path.join(folder, f)
            loc = location_from_name(f)
            shts = load_file(full)
            for sh, raw in shts.items():
                dfc = smart_header(raw)
                cols = [c for c in dfc.columns if c and c not in ("nan","None") and not c.startswith("Col_")]
                num_c = dfc.select_dtypes(include="number").columns.tolist()
                cat_c = [x for x in cols if x not in num_c]
                key = f"{loc} | {sh}"
                result[key] = {
                    "all_cols": cols,
                    "num_cols": num_c,
                    "cat_cols": cat_c,
                    "df": dfc,
                    "file": f,
                    "loc": loc,
                    "sheet": sh,
                }
        return result

    all_col_info = get_all_column_headers(tuple(excel_files), data_dir)

    # Gather all unique column headers across all files
    all_headers_set = []
    seen_h = set()
    for info in all_col_info.values():
        for col in info["all_cols"]:
            if col not in seen_h:
                all_headers_set.append(col)
                seen_h.add(col)

    # Display column format for current file first
    cur_key = f"{loc_label} | {selected_sheet}"
    cur_info = all_col_info.get(cur_key, {})
    if cur_info:
        st.markdown(f'<span class="loc-badge">📍 {loc_label}</span>', unsafe_allow_html=True)
        cols_to_show = cur_info["all_cols"][:40]
        # Show as chips like Excel column headers
        chips_html = ""
        for col in cols_to_show:
            tag = "🔢" if col in cur_info["num_cols"] else "📝"
            chips_html += f'<span class="col-chip">{tag} {col}</span>'
        st.markdown(chips_html, unsafe_allow_html=True)

    st.markdown("---")

    # Show all locations in expandable format
    st.markdown('<div class="sec-title">📊 All Locations – Column Structure</div>', unsafe_allow_html=True)

    for key, info in all_col_info.items():
        with st.expander(f"📍 {key}  ({len(info['all_cols'])} columns · {len(info['df'])} rows)", expanded=False):
            colA, colB = st.columns(2)
            with colA:
                st.markdown("**🔢 Numeric Columns**")
                for c in info["num_cols"]:
                    s = info["df"][c].dropna()
                    if len(s):
                        st.markdown(f'<div class="excel-header">{c}</div>', unsafe_allow_html=True)
            with colB:
                st.markdown("**📝 Text / Category Columns**")
                for c in info["cat_cols"][:20]:
                    st.markdown(f'<div class="excel-header">{c}</div>', unsafe_allow_html=True)

            # Show sample rows in Excel-like format
            st.markdown("**📄 Sample Data (first 5 rows)**")
            disp_cols = info["all_cols"][:20]
            if disp_cols:
                st.dataframe(
                    info["df"][disp_cols].head(5),
                    use_container_width=True,
                )

    st.markdown("---")

    if num_cols:
        st.markdown('<div class="sec-title">📐 Quick Statistics</div>', unsafe_allow_html=True)
        stats = df_clean[num_cols].describe().T
        stats["range"] = stats["max"] - stats["min"]
        st.dataframe(
            stats.style.format("{:.3f}", na_rep="—").background_gradient(cmap="Blues", subset=["mean", "max"]),
            use_container_width=True,
        )

    st.markdown('<div class="sec-title">🗂️ Column Overview</div>', unsafe_allow_html=True)
    ci = pd.DataFrame({
        "Column": df_clean.columns,
        "Type": df_clean.dtypes.values,
        "Non-Null": df_clean.notna().sum().values,
        "Null%": (df_clean.isna().mean() * 100).round(1).values,
        "Unique": [df_clean[c].nunique() for c in df_clean.columns],
        "Sample": [
            str(df_clean[c].dropna().iloc[0])[:55]
            if df_clean[c].dropna().shape[0] > 0 else "—"
            for c in df_clean.columns
        ],
    })
    st.dataframe(ci, use_container_width=True)

    st.markdown("---")
    st.markdown('<div class="sec-title">📄 Complete Raw Sheet (ALL Rows · ALL Columns)</div>', unsafe_allow_html=True)
    st.dataframe(raw_df, use_container_width=True, height=500)
    st.download_button(
        "⬇️ Download Complete Sheet CSV",
        raw_df.to_csv(index=False).encode(),
        f"{loc_label}_{selected_sheet}_complete.csv",
        "text/csv",
        key="ov_dl",
    )


# ═══════════════════════════════════════════════════════
# TAB 1 – RAW DATA
# ═══════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("📋 Data Table – Searchable & Filterable")
    col_a, col_b = st.columns([3, 1])
    srch = col_a.text_input("🔍 Live search across all columns", "", key="rawsrch")
    show_raw = col_b.checkbox("Show raw (no header detection)", False)

    disp_df = raw_df if show_raw else df_clean
    disp = (
        disp_df[disp_df.apply(lambda col: col.astype(str).str.contains(srch, case=False, na=False)).any(axis=1)]
        if srch else disp_df
    )
    st.caption(f"Showing {len(disp):,} / {len(disp_df):,} rows · {len(disp.columns)} columns")
    st.dataframe(disp, use_container_width=True, height=520)
    st.download_button("⬇️ Export as CSV", disp.to_csv(index=False).encode(), "export.csv", "text/csv")


# ═══════════════════════════════════════════════════════
# TAB 2 – ANALYTICS
# ═══════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("📊 Column Analytics & Aggregations")
    if not num_cols:
        st.info("ℹ️ No numeric columns detected in this sheet.")
    else:
        chosen = st.multiselect(
            "Select columns to analyse",
            num_cols,
            default=num_cols[:min(6, len(num_cols))],
        )
        if chosen:
            sub = df_clean[chosen].dropna(how="all")
            kc = st.columns(min(len(chosen), 6))
            for i, col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s):
                    kc[i].metric(col[:22], f"{s.sum():,.1f}", f"avg {s.mean():,.1f}")

            st.markdown("---")
            agg_rows = []
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    grand = df_clean[chosen].select_dtypes("number").sum().sum()
                    agg_rows.append({
                        "Column": col,
                        "Count": int(s.count()),
                        "Sum": round(s.sum(), 4),
                        "Mean": round(s.mean(), 4),
                        "Median": round(s.median(), 4),
                        "Min": round(s.min(), 4),
                        "Max": round(s.max(), 4),
                        "Std Dev": round(s.std(), 4),
                        "Variance": round(s.var(), 4),
                        "Skewness": round(s.skew(), 4) if len(s) >= 3 else "—",
                        "Kurtosis": round(s.kurtosis(), 4) if len(s) >= 4 else "—",
                        "% of Total": f"{s.sum()/grand*100:.1f}%" if grand else "—",
                    })
            if agg_rows:
                adf = pd.DataFrame(agg_rows).set_index("Column")
                fmt_cols = [c for c in adf.columns if c not in ("% of Total", "Skewness", "Kurtosis")]
                st.dataframe(
                    adf.style.format("{:,.3f}", na_rep="—", subset=fmt_cols)
                    .background_gradient(cmap="YlOrRd", subset=["Sum", "Max"]),
                    use_container_width=True,
                )

        st.markdown("---")
        st.markdown('<div class="sec-title">🧮 Group-By Aggregation</div>', unsafe_allow_html=True)
        all_cat = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique() < 80]
        if all_cat and num_cols:
            gc1, gc2, gc3 = st.columns(3)
            gc = gc1.selectbox("Group by column", all_cat)
            ac = gc2.selectbox("Aggregate column", num_cols)
            af = gc3.selectbox("Function", ["sum", "mean", "count", "min", "max", "median", "std", "var"])
            grp = (
                df_clean.groupby(gc)[ac].agg(af).reset_index()
                .rename(columns={ac: f"{af}({ac})"})
                .sort_values(f"{af}({ac})", ascending=False)
            )
            st.dataframe(grp, use_container_width=True)
            fig = px.bar(
                grp, x=gc, y=f"{af}({ac})", color=f"{af}({ac})",
                color_continuous_scale="Viridis",
                title=f"{af.title()} of {ac} by {gc}",
            )
            fig.update_layout(xaxis_tickangle=-35, height=420,
                              paper_bgcolor="#1a2035", plot_bgcolor="#0f1829", font_color="#c9d8f0")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Add categorical columns or more data to use group-by.")


# ═══════════════════════════════════════════════════════
# TAB 3 – CHARTS (all 15 types)
# ═══════════════════════════════════════════════════════
DARK = dict(paper_bgcolor="#1a2035", plot_bgcolor="#0f1829", font_color="#c9d8f0")

with tabs[3]:
    st.subheader("📈 Interactive Charts")

    # Chart type descriptions
    CHART_DESC = {
        "Bar Chart": "Compare a numeric metric across categorical groups (e.g., power per customer)",
        "Grouped Bar": "Side-by-side comparison of multiple numeric columns across groups",
        "Stacked Bar": "Show composition and total simultaneously across groups",
        "Line Chart": "Trend analysis across ordered rows or time-series data",
        "Scatter Plot": "Correlation between two numeric variables; colour-coded by a third",
        "Area Chart": "Cumulative volume trends with filled area for visual emphasis",
        "Bubble Chart": "Three-dimensional numeric relationships (X, Y, size)",
        "Heatmap (Correlation)": "Instantly spot which numeric columns are correlated",
        "Box Plot": "Distribution, spread, median, and outliers for numeric columns",
        "Violin Plot": "Full probability distribution shape for numeric columns",
        "Funnel Chart": "Staged capacity utilisation or sales-pipeline visualisation",
        "Waterfall / Cumulative": "Running total analysis, e.g., cumulative power consumed",
        "3-D Scatter": "Three-axis numeric exploration for high-dimensional data",
        "Radar Chart": "Multi-axis comparison of a single entity across metrics",
        "Histogram": "Frequency distribution of a numeric variable",
    }

    ctype = st.selectbox(
        "Chart Type",
        list(CHART_DESC.keys()),
    )
    st.info(f"ℹ️ **{ctype}:** {CHART_DESC[ctype]}")

    if not num_cols:
        st.info("No numeric columns in this sheet.")
    else:
        def _s(label, opts, idx=0, key=None):
            if not opts:
                return None
            return st.selectbox(label, opts, index=min(idx, max(0, len(opts)-1)), key=key)

        if ctype == "Bar Chart":
            xc = _s("X axis", cat_cols or df_clean.columns.tolist(), key="bx")
            yc = _s("Y axis", num_cols, key="by")
            ori = st.radio("Orientation", ["Vertical", "Horizontal"], horizontal=True)
            d = df_clean[[xc, yc]].dropna()
            fig = px.bar(
                d,
                x=xc if ori == "Vertical" else yc,
                y=yc if ori == "Vertical" else xc,
                color=yc,
                color_continuous_scale="Turbo",
                orientation="v" if ori == "Vertical" else "h",
                title=f"{yc} by {xc}",
            )
            fig.update_layout(height=480, **DARK)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Grouped Bar":
            xc = _s("X axis", cat_cols or df_clean.columns.tolist(), key="gbx")
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:min(4, len(num_cols))])
            if ycs:
                fig = px.bar(
                    df_clean[[xc]+ycs].dropna(subset=ycs, how="all"),
                    x=xc, y=ycs, barmode="group",
                    title=f"Grouped: {', '.join(ycs)} by {xc}",
                )
                fig.update_layout(height=460, **DARK)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Stacked Bar":
            xc = _s("X axis", cat_cols or df_clean.columns.tolist(), key="sbx")
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:min(4, len(num_cols))])
            if ycs:
                fig = px.bar(
                    df_clean[[xc]+ycs].dropna(subset=ycs, how="all"),
                    x=xc, y=ycs, barmode="stack",
                    title=f"Stacked: {', '.join(ycs)} by {xc}",
                )
                fig.update_layout(height=460, **DARK)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Line Chart":
            xc = _s("X axis", df_clean.columns.tolist(), key="lx")
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:min(3, len(num_cols))])
            if ycs:
                fig = px.line(
                    df_clean[[xc]+ycs].dropna(subset=ycs, how="all"),
                    x=xc, y=ycs, markers=True,
                    title=f"Line: {', '.join(ycs)} vs {xc}",
                )
                fig.update_layout(height=450, **DARK)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Scatter Plot":
            xc = _s("X", num_cols, 0, "sc_x")
            yc = _s("Y", num_cols, min(1, len(num_cols)-1), "sc_y")
            sc = _s("Size (optional)", ["None"] + num_cols, key="sc_s")
            cc = _s("Color (optional)", ["None"] + cat_cols + num_cols, key="sc_c")
            if xc and yc:
                d = df_clean.dropna(subset=[xc, yc])
                fig = px.scatter(
                    d, x=xc, y=yc,
                    size=sc if sc != "None" else None,
                    color=cc if cc != "None" else None,
                    color_continuous_scale="Rainbow",
                    title=f"Scatter: {yc} vs {xc}",
                    trendline="ols" if st.checkbox("Add trendline", key="sc_tl") else None,
                )
                fig.update_layout(height=480, **DARK)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Area Chart":
            xc = _s("X axis", df_clean.columns.tolist(), key="ax")
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:min(3, len(num_cols))])
            if ycs:
                fig = px.area(
                    df_clean[[xc]+ycs].dropna(subset=ycs, how="all"),
                    x=xc, y=ycs, title=f"Area: {', '.join(ycs)}",
                )
                fig.update_layout(height=450, **DARK)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Bubble Chart":
            if len(num_cols) >= 2:
                xc = _s("X", num_cols, 0, "bu_x")
                yc = _s("Y", num_cols, min(1, len(num_cols)-1), "bu_y")
                sz = _s("Size", num_cols, min(2, len(num_cols)-1), "bu_s")
                lc = _s("Color", ["None"] + cat_cols, key="bu_c")
                d = df_clean[[xc, yc, sz]].dropna()
                if lc != "None":
                    d[lc] = df_clean[lc]
                fig = px.scatter(
                    d, x=xc, y=yc, size=sz,
                    color=lc if lc != "None" else None,
                    size_max=70, title=f"Bubble: {yc} vs {xc} (size={sz})",
                )
                fig.update_layout(height=500, **DARK)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need ≥ 2 numeric columns for a bubble chart.")

        elif ctype == "Heatmap (Correlation)":
            sel = st.multiselect("Columns", num_cols, default=num_cols[:min(12, len(num_cols))])
            if len(sel) >= 2:
                corr = df_clean[sel].corr()
                fig = px.imshow(
                    corr, text_auto=".2f", color_continuous_scale="RdBu_r",
                    aspect="auto", title="Correlation Matrix",
                )
                fig.update_layout(height=560, **DARK)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Box Plot":
            yc = _s("Value column", num_cols, key="bp_v")
            xc = _s("Group by", ["None"] + cat_cols, key="bp_g")
            d = df_clean[[yc]+([xc] if xc != "None" else [])].dropna(subset=[yc])
            fig = px.box(
                d, y=yc, x=xc if xc != "None" else None,
                color=xc if xc != "None" else None,
                points="outliers", title=f"Box Plot: {yc}",
            )
            fig.update_layout(height=450, **DARK)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Violin Plot":
            yc = _s("Value column", num_cols, key="vp_v")
            xc = _s("Group by", ["None"] + cat_cols, key="vp_g")
            d = df_clean[[yc]+([xc] if xc != "None" else [])].dropna(subset=[yc])
            fig = px.violin(
                d, y=yc, x=xc if xc != "None" else None,
                box=True, points="outliers",
                color=xc if xc != "None" else None,
                title=f"Violin Plot: {yc}",
            )
            fig.update_layout(height=450, **DARK)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Funnel Chart":
            xc = _s("Stage column", cat_cols or df_clean.columns.tolist(), key="fn_x")
            yc = _s("Value column", num_cols, key="fn_y")
            d = (
                df_clean[[xc, yc]].dropna()
                .groupby(xc)[yc].sum().reset_index()
                .sort_values(yc, ascending=False)
            )
            fig = px.funnel(d, x=yc, y=xc, title=f"Funnel: {yc} by {xc}")
            fig.update_layout(height=450, **DARK)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Waterfall / Cumulative":
            yc = _s("Column", num_cols, key="wf_y")
            d = df_clean[yc].dropna().reset_index(drop=True)
            cum = d.cumsum()
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Value", x=d.index, y=d, marker_color="#2a5298"))
            fig.add_trace(go.Scatter(
                name="Cumulative", x=cum.index, y=cum,
                line=dict(color="#f7971e", width=2.5), mode="lines+markers",
            ))
            fig.update_layout(title=f"Waterfall/Cumulative: {yc}", height=450, barmode="group", **DARK)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "3-D Scatter":
            if len(num_cols) >= 3:
                xc = _s("X", num_cols, 0, "3x")
                yc = _s("Y", num_cols, min(1, len(num_cols)-1), "3y")
                zc = _s("Z", num_cols, min(2, len(num_cols)-1), "3z")
                cc = _s("Color", ["None"] + cat_cols, key="3c")
                d = df_clean[[xc, yc, zc]].dropna()
                if cc != "None":
                    d[cc] = df_clean[cc]
                fig = px.scatter_3d(
                    d, x=xc, y=yc, z=zc,
                    color=cc if cc != "None" else None,
                    title=f"3D Scatter: {xc}/{yc}/{zc}",
                )
                fig.update_layout(height=570, **DARK)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need ≥ 3 numeric columns for 3-D scatter.")

        elif ctype == "Radar Chart":
            sel_r = st.multiselect("Columns for radar", num_cols, default=num_cols[:min(6, len(num_cols))])
            if len(sel_r) >= 3 and cat_cols:
                gc = _s("Group by", cat_cols, key="ra_g")
                d = df_clean[[gc]+sel_r].dropna().head(10)
                fig = go.Figure()
                for _, row in d.iterrows():
                    vals = [row[c] for c in sel_r] + [row[sel_r[0]]]
                    fig.add_trace(go.Scatterpolar(
                        r=vals, theta=sel_r+[sel_r[0]],
                        fill="toself", name=str(row[gc])[:30],
                    ))
                fig.update_layout(polar=dict(bgcolor="#0f1829"), height=500, title="Radar Chart", **DARK)
                st.plotly_chart(fig, use_container_width=True)
            elif len(sel_r) >= 3:
                vals = df_clean[sel_r].mean().tolist()
                fig = go.Figure(go.Scatterpolar(
                    r=vals+[vals[0]], theta=sel_r+[sel_r[0]], fill="toself",
                ))
                fig.update_layout(polar=dict(bgcolor="#0f1829"), height=500, title="Radar Chart (Averages)", **DARK)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Select ≥ 3 numeric columns.")

        elif ctype == "Histogram":
            yc = _s("Column", num_cols, key="hi_y")
            bins = st.slider("Bins", 5, 100, 30)
            fig = px.histogram(
                df_clean, x=yc, nbins=bins,
                title=f"Histogram: {yc}",
                color_discrete_sequence=["#2a5298"],
            )
            fig.update_layout(height=450, **DARK)
            st.plotly_chart(fig, use_container_width=True)


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
            pc_list = [
                c for c in df_clean.columns
                if df_clean[c].nunique() < 40 and df_clean[c].notna().sum() > 0
            ]
            if pc_list:
                pc = st.selectbox("Label column", pc_list, key="pc_l")
                pv = st.selectbox("Value column", ["Count"] + num_cols, key="pc_v")
                hole = st.slider("Donut hole", 0.0, 0.8, 0.4, 0.05)
                if pv == "Count":
                    d = df_clean[pc].value_counts().reset_index()
                    d.columns = [pc, "Count"]
                    fig = px.pie(d, names=pc, values="Count", hole=hole,
                                 title=f"Distribution of {pc}",
                                 color_discrete_sequence=px.colors.qualitative.Plotly)
                else:
                    d = df_clean[[pc, pv]].dropna().groupby(pc)[pv].sum().reset_index()
                    fig = px.pie(d, names=pc, values=pv, hole=hole,
                                 title=f"{pv} by {pc}",
                                 color_discrete_sequence=px.colors.qualitative.Bold)
                fig.update_layout(height=420, **DARK)
                st.plotly_chart(fig, use_container_width=True)

        with r2:
            st.markdown('<div class="sec-title">📊 Histogram</div>', unsafe_allow_html=True)
            hc = st.selectbox("Column", num_cols, key="hc_col")
            bins2 = st.slider("Bins", 5, 100, 25, key="hc_bins")
            fig2 = px.histogram(
                df_clean, x=hc, nbins=bins2,
                marginal="box",
                color_discrete_sequence=["#2a5298"],
                title=f"Distribution: {hc}",
            )
            fig2.update_layout(height=420, **DARK)
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("---")
        st.markdown('<div class="sec-title">📦 Box Plot – All Numeric Columns</div>', unsafe_allow_html=True)
        sel_bp = st.multiselect("Columns", num_cols, default=num_cols[:min(8, len(num_cols))], key="bp_all")
        if sel_bp:
            fig3 = go.Figure()
            colors = px.colors.qualitative.Plotly
            for i, col in enumerate(sel_bp):
                fig3.add_trace(go.Box(
                    y=df_clean[col].dropna(), name=col,
                    marker_color=colors[i % len(colors)],
                    boxpoints="outliers",
                ))
            fig3.update_layout(title="Box Plots Comparison", height=450, showlegend=False, **DARK)
            st.plotly_chart(fig3, use_container_width=True)

        st.markdown("---")
        st.markdown('<div class="sec-title">🎻 Sunburst Chart</div>', unsafe_allow_html=True)
        if len(cat_cols) >= 1 and num_cols:
            sb_path = st.multiselect("Hierarchy (select 1-3)", cat_cols, default=cat_cols[:min(2, len(cat_cols))], key="sb_p")
            sb_val = st.selectbox("Value", num_cols, key="sb_v")
            if sb_path and sb_val:
                d_sb = df_clean[sb_path + [sb_val]].dropna()
                if len(d_sb) > 0:
                    fig4 = px.sunburst(
                        d_sb, path=sb_path, values=sb_val,
                        title=f"Sunburst: {sb_val}",
                        color_discrete_sequence=px.colors.qualitative.Plotly,
                    )
                    fig4.update_layout(height=500, **DARK)
                    st.plotly_chart(fig4, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 5 – QUERY ENGINE (No-code with all operations)
# ═══════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("🔍 No-Code Query Engine")
    st.caption("Pick operation + columns; get instant result with chart. Works across ALL locations.")

    # Scope selector
    q_scope = st.radio(
        "Data scope:",
        ["🌍 All Files (entire corpus)", f"📄 Current Sheet ({loc_label} › {selected_sheet})"],
        horizontal=True,
        key="qe_scope",
    )
    use_all_qe = q_scope.startswith("🌍")

    if use_all_qe:
        # Build a combined dataframe from all files
        @st.cache_data(show_spinner=False)
        def build_combined_df(files, folder):
            dfs = []
            for f in files:
                full = os.path.join(folder, f)
                loc = location_from_name(f)
                shts = load_file(full)
                for sh, raw in shts.items():
                    dfc = to_numeric(smart_header(raw))
                    dfc["__Location__"] = loc
                    dfc["__Sheet__"] = sh
                    dfc["__File__"] = f
                    dfs.append(dfc)
            if not dfs:
                return pd.DataFrame()
            combined = pd.concat(dfs, ignore_index=True, sort=False)
            return combined

        combined_df = build_combined_df(tuple(excel_files), data_dir)
        work_df = combined_df
        work_num = [c for c in combined_df.select_dtypes(include="number").columns
                    if not c.startswith("__")]
        work_cat = [c for c in combined_df.columns
                    if c not in combined_df.select_dtypes(include="number").columns
                    and not c.startswith("__")]
    else:
        work_df = df_clean
        work_num = num_cols
        work_cat = cat_cols

    if not work_num:
        st.info("No numeric columns available.")
    else:
        q1, q2, q3, q4 = st.columns(4)
        op = q1.selectbox("Operation", [
            "Sum", "Total", "Average / Mean", "Min", "Max",
            "Count", "Percentage of Total", "Median",
            "Std Deviation", "Variance", "Range (max-min)",
            "Skewness", "Kurtosis", "IQR",
            "Top N rows", "Bottom N rows",
            "All Stats (Full Summary)",
        ], key="qe_op")
        sc_col = q2.selectbox("Column", work_num, key="qe_col")

        # Filter options
        filter_cols = ["—"] + work_cat + (["__Location__", "__Sheet__"] if use_all_qe else [])
        fc = q3.selectbox("Filter by", filter_cols, key="qe_fc")
        fv = q4.text_input("Filter value (partial match)", "", key="qe_fv") if fc != "—" else ""

        n_val = 10
        if op in ("Top N rows", "Bottom N rows"):
            n_val = st.number_input("N", 1, 500, 10)

        # Chart type for result
        result_chart = st.selectbox("Result chart type", [
            "Bar Chart", "Pie Chart", "Line Chart", "Box Plot", "Histogram",
        ], key="qe_chart")

        if st.button("▶ Run Query", type="primary"):
            sub = work_df.copy()
            if fc != "—" and fv:
                sub = sub[sub[fc].astype(str).str.contains(fv, case=False, na=False)]

            vals = sub[sc_col].dropna()
            r = None
            detail_df = None

            if op == "Sum":
                r = vals.sum()
            elif op == "Total":
                r = vals.sum()
            elif op == "Average / Mean":
                r = vals.mean()
            elif op == "Min":
                r = vals.min()
            elif op == "Max":
                r = vals.max()
            elif op == "Count":
                r = int(vals.count())
            elif op == "Percentage of Total":
                grand = work_df[sc_col].dropna().sum()
                r = f"{vals.sum()/grand*100:.2f}%" if grand else "N/A"
            elif op == "Median":
                r = vals.median()
            elif op == "Std Deviation":
                r = vals.std()
            elif op == "Variance":
                r = vals.var()
            elif op == "Range (max-min)":
                r = vals.max() - vals.min()
            elif op == "Skewness":
                r = vals.skew()
            elif op == "Kurtosis":
                r = vals.kurtosis()
            elif op == "IQR":
                r = vals.quantile(0.75) - vals.quantile(0.25)
            elif op == "All Stats (Full Summary)":
                sa = vals
                stat_rows = {
                    "Count": int(sa.count()),
                    "Sum": round(sa.sum(), 4),
                    "Mean": round(sa.mean(), 4),
                    "Median": round(sa.median(), 4),
                    "Min": round(sa.min(), 4),
                    "Max": round(sa.max(), 4),
                    "Std Dev": round(sa.std(), 4),
                    "Variance": round(sa.var(), 4),
                    "Skewness": round(sa.skew(), 4) if len(sa) >= 3 else "—",
                    "Kurtosis": round(sa.kurtosis(), 4) if len(sa) >= 4 else "—",
                    "IQR": round(sa.quantile(0.75) - sa.quantile(0.25), 4),
                    "Range": round(sa.max() - sa.min(), 4),
                    "Q1 (25%)": round(sa.quantile(0.25), 4),
                    "Q3 (75%)": round(sa.quantile(0.75), 4),
                }
                filt_note = f" (where {fc} contains '{fv}')" if fv else ""
                scope_note = "All Files" if use_all_qe else f"{loc_label} › {selected_sheet}"
                st.success(f"**Full Statistics** of `{sc_col}` — {scope_note}{filt_note}")
                sdf = pd.DataFrame(list(stat_rows.items()), columns=["Metric", "Value"])
                st.dataframe(sdf, use_container_width=True, hide_index=True)
                # Show histogram
                fig_h = px.histogram(sub, x=sc_col, nbins=30, title=f"Distribution of {sc_col}{filt_note}",
                                     color_discrete_sequence=["#2a5298"])
                fig_h.update_layout(**DARK)
                st.plotly_chart(fig_h, use_container_width=True)
                r = None

            elif op == "Top N rows":
                tdf = sub.nlargest(int(n_val), sc_col)
                st.success(f"**Top {n_val}** rows by `{sc_col}`")
                st.dataframe(tdf, use_container_width=True)
                detail_df = tdf
                r = None
            elif op == "Bottom N rows":
                bdf = sub.nsmallest(int(n_val), sc_col)
                st.success(f"**Bottom {n_val}** rows by `{sc_col}`")
                st.dataframe(bdf, use_container_width=True)
                detail_df = bdf
                r = None

            if r is not None:
                if isinstance(r, float):
                    r_str = f"{r:,.4f}"
                else:
                    r_str = str(r)
                filt_note = f" (where {fc} contains '{fv}')" if fv else ""
                scope_note = "All Files" if use_all_qe else f"{loc_label} › {selected_sheet}"
                st.success(f"**{op}** of `{sc_col}` — {scope_note}{filt_note} → **{r_str}**")

                # Show result chart
                if use_all_qe and "__Location__" in work_df.columns:
                    loc_grp = sub.groupby("__Location__")[sc_col]
                    if op in ("Sum", "Total"):
                        chart_data = loc_grp.sum().reset_index()
                    elif op in ("Average / Mean"):
                        chart_data = loc_grp.mean().reset_index()
                    elif op == "Count":
                        chart_data = loc_grp.count().reset_index()
                    elif op == "Max":
                        chart_data = loc_grp.max().reset_index()
                    elif op == "Min":
                        chart_data = loc_grp.min().reset_index()
                    else:
                        chart_data = loc_grp.sum().reset_index()
                    chart_data.columns = ["Location", op]

                    if result_chart == "Bar Chart":
                        fig_r = px.bar(chart_data, x="Location", y=op, color=op,
                                       color_continuous_scale="Viridis",
                                       title=f"{op} of {sc_col} by Location{filt_note}")
                    elif result_chart == "Pie Chart":
                        fig_r = px.pie(chart_data, names="Location", values=op, hole=0.4,
                                       title=f"{op} of {sc_col} by Location")
                    elif result_chart == "Line Chart":
                        fig_r = px.line(chart_data, x="Location", y=op, markers=True,
                                        title=f"{op} of {sc_col} by Location")
                    elif result_chart == "Box Plot":
                        fig_r = px.box(sub, y=sc_col, x="__Location__",
                                       title=f"Box Plot of {sc_col} by Location")
                    else:
                        fig_r = px.histogram(sub, x=sc_col, nbins=30, title=f"Histogram of {sc_col}")

                    fig_r.update_layout(**DARK, xaxis_tickangle=-30, height=440)
                    st.plotly_chart(fig_r, use_container_width=True)

            if detail_df is not None:
                st.download_button("⬇️ Download Result CSV", detail_df.to_csv(index=False).encode(),
                                   "query_result.csv", "text/csv")


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
                        "df": dfc, "num_cols": nc, "file": f, "sheet": sh,
                    }
        return summ

    all_summ = load_all_summ(tuple(excel_files), data_dir)

    if all_summ:
        all_num = sorted({c for v in all_summ.values() for c in v["num_cols"]})
        if all_num:
            ml1, ml2 = st.columns(2)
            comp_col = ml1.selectbox("Compare by column", all_num)
            agg_fn = ml2.selectbox("Aggregation", ["sum", "mean", "max", "min", "count", "median", "std"])

            rows_cmp = []
            for lbl, info in all_summ.items():
                if comp_col in info["num_cols"]:
                    s = info["df"][comp_col].dropna()
                    rows_cmp.append({
                        "Location|Sheet": lbl,
                        "Sum": round(s.sum(), 2),
                        "Mean": round(s.mean(), 2),
                        "Max": round(s.max(), 2),
                        "Min": round(s.min(), 2),
                        "Median": round(s.median(), 2),
                        "Std Dev": round(s.std(), 2),
                        "Count": int(s.count()),
                    })
            if rows_cmp:
                cmp = pd.DataFrame(rows_cmp).set_index("Location|Sheet")
                st.dataframe(
                    cmp.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),
                    use_container_width=True,
                )
                col_a, col_b = st.columns(2)
                agg_col = agg_fn.title() if agg_fn.title() in cmp.columns else "Sum"
                with col_a:
                    fig_bar = px.bar(
                        cmp.reset_index(), x="Location|Sheet", y=agg_col,
                        color=agg_col, color_continuous_scale="Viridis",
                        title=f"{agg_col} of '{comp_col}' by Location",
                    )
                    fig_bar.update_layout(xaxis_tickangle=-30, height=440, **DARK)
                    st.plotly_chart(fig_bar, use_container_width=True)
                with col_b:
                    fig_pie = px.pie(
                        cmp.reset_index(), names="Location|Sheet", values=agg_col,
                        title=f"Share of '{comp_col}'",
                        color_discrete_sequence=px.colors.qualitative.Plotly,
                        hole=0.4,
                    )
                    fig_pie.update_layout(height=440, **DARK)
                    st.plotly_chart(fig_pie, use_container_width=True)

                # Multi-metric grouped bar
                st.markdown('<div class="sec-title">📊 Multi-Metric Comparison</div>', unsafe_allow_html=True)
                metrics_sel = st.multiselect("Metrics", ["Sum","Mean","Max","Min","Count"],
                                              default=["Sum","Mean","Max"])
                if metrics_sel:
                    plot_df = cmp.reset_index()[["Location|Sheet"] + metrics_sel]
                    fig_gb = px.bar(
                        plot_df.melt(id_vars="Location|Sheet", var_name="Metric", value_name="Value"),
                        x="Location|Sheet", y="Value", color="Metric", barmode="group",
                        title=f"Multi-metric comparison of '{comp_col}'",
                    )
                    fig_gb.update_layout(xaxis_tickangle=-30, height=460, **DARK)
                    st.plotly_chart(fig_gb, use_container_width=True)

                # Radar comparison
                if len(rows_cmp) >= 3:
                    st.markdown('<div class="sec-title">🕸️ Multi-Location Radar</div>', unsafe_allow_html=True)
                    metrics = ["Sum", "Mean", "Max", "Min"]
                    fig_r = go.Figure()
                    for row_d in rows_cmp[:8]:
                        vals = [row_d[m] for m in metrics] + [row_d[metrics[0]]]
                        fig_r.add_trace(go.Scatterpolar(
                            r=vals, theta=metrics + [metrics[0]],
                            fill="toself", name=row_d["Location|Sheet"][:30],
                        ))
                    fig_r.update_layout(
                        polar=dict(bgcolor="#0f1829"), height=500,
                        title=f"Radar Comparison: {comp_col}", **DARK,
                    )
                    st.plotly_chart(fig_r, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 7 – AI AGENT (Automated Insights)
# ═══════════════════════════════════════════════════════
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights Engine")
    st.caption("Click to run automated analysis across ALL locations, sheets, rows, and columns.")

    col_run, col_opts = st.columns([1, 3])
    run_full = col_run.button("🚀 Run Full Analysis", type="primary")
    max_files = col_opts.slider("Max files to analyse", 1, len(excel_files), min(5, len(excel_files)))

    if run_full:
        with st.spinner("🔄 Analysing all data…"):
            progress = st.progress(0)
            entries = list(all_summ.items())[:max_files * 10]
            for fi, (lbl, info) in enumerate(entries):
                progress.progress(min((fi+1) / max(1, len(entries)), 1.0))
                dfa = info["df"]
                nc = info["num_cols"]
                if not nc:
                    continue
                with st.expander(f"📍 {lbl}", expanded=False):
                    ca, cb, cc_col = st.columns(3)
                    with ca:
                        st.markdown("**📊 Key Metrics**")
                        for col in nc[:5]:
                            s = dfa[col].dropna()
                            if len(s):
                                st.metric(col[:24], f"{s.sum():,.1f}", f"avg {s.mean():,.1f}")
                    with cb:
                        st.markdown("**🎯 Outlier Detection**")
                        for col in nc[:4]:
                            s = dfa[col].dropna()
                            if len(s) > 3:
                                z = (s - s.mean()) / (s.std() + 1e-9)
                                o = z[z.abs() > 2.5]
                                (st.warning if len(o) else st.success)(
                                    f"`{col}`: {len(o)} outlier(s)" if len(o) else f"`{col}`: Clean ✓"
                                )
                    with cc_col:
                        st.markdown("**📈 Trend Summary**")
                        for col in nc[:3]:
                            s = dfa[col].dropna()
                            if len(s) >= 3:
                                trend = "↑ Rising" if s.iloc[-1] > s.mean() else "↓ Below Avg"
                                st.info(f"`{col}`: {trend}")
            progress.empty()

    st.markdown("---")
    st.markdown('<div class="sec-title">📁 All Files Summary</div>', unsafe_allow_html=True)

    fsm = []
    for f in excel_files:
        shd = load_file(os.path.join(data_dir, f))
        total_rows = sum(len(s) for s in shd.values())
        total_cols = max((s.shape[1] for s in shd.values()), default=0)
        fsm.append({
            "Location": loc_map[f],
            "File": f,
            "Sheets": len(shd),
            "Sheet Names": ", ".join(list(shd.keys())[:5]),
            "Total Rows": total_rows,
            "Max Columns": total_cols,
        })
    files_df = pd.DataFrame(fsm)
    st.dataframe(files_df, use_container_width=True)

    if len(fsm) > 1:
        fig_fs = px.bar(
            files_df, x="Location", y="Total Rows",
            color="Sheets", title="Rows per Location",
            color_continuous_scale="Blues",
        )
        fig_fs.update_layout(xaxis_tickangle=-30, height=380, **DARK)
        st.plotly_chart(fig_fs, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 8 – SMART QUERY (Natural Language + Cell Value Search)
# ═══════════════════════════════════════════════════════
with tabs[8]:
    st.markdown("## 💬 AI Smart Query")

    scope_col1, scope_col2 = st.columns([3, 1])
    with scope_col1:
        search_scope = st.radio(
            "Search scope:",
            ["🌍 All Files (entire corpus)", f"📄 Current Sheet ({loc_label} › {selected_sheet})"],
            horizontal=True,
            key="sq_scope",
        )
    with scope_col2:
        st.markdown(
            f'<div class="kcard kcard-blue" style="padding:10px 14px;">'
            f'<h2 style="font-size:1.4rem;">{meta["total_cells"]:,}</h2>'
            f'<p>Total indexed cells</p></div>',
            unsafe_allow_html=True,
        )

    use_corpus = search_scope.startswith("🌍")

    qi1, qi2, qi3, qi4 = st.columns(4)
    if use_corpus:
        qi1.markdown(f'<div class="kcard kcard-blue"><h2>{meta["total_cells"]:,}</h2><p>Total Cells (All Files)</p></div>', unsafe_allow_html=True)
        qi2.markdown(f'<div class="kcard kcard-green"><h2>{meta["total_files"]}</h2><p>Files Indexed</p></div>', unsafe_allow_html=True)
        qi3.markdown(f'<div class="kcard kcard-purple"><h2>{meta["total_sheets"]}</h2><p>Sheets Indexed</p></div>', unsafe_allow_html=True)
        qi4.markdown(f'<div class="kcard kcard-cyan"><h2>{meta["total_rows"]:,}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
    else:
        qi1.markdown(f'<div class="kcard kcard-blue"><h2>{sq_meta["total_cells"]:,}</h2><p>Sheet Cells</p></div>', unsafe_allow_html=True)
        qi2.markdown(f'<div class="kcard kcard-green"><h2>{sq_meta["total_data"]:,}</h2><p>Data Cells</p></div>', unsafe_allow_html=True)
        qi3.markdown(f'<div class="kcard kcard-purple"><h2>{sq_meta["total_rows"]:,}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
        qi4.markdown(f'<div class="kcard kcard-cyan"><h2>{sq_meta["total_headers"]:,}</h2><p>Header Cells</p></div>', unsafe_allow_html=True)

    # ── NATURAL LANGUAGE QUERY ENGINE ──────────────────────────────────────
    _SW2 = {
        "the","and","for","are","all","any","how","what","show","give",
        "tell","from","this","that","with","get","find","list","much",
        "many","each","every","data","value","values","number","numbers",
        "in","of","a","an","is","at","by","to","do","me","my","about",
        "details","info","please","can","you","per","across","which",
        "where","who","when","does","did","have","has","their","its",
        "our","your","there","these","those","been","will","would","could",
        "should","shall","let","some","just","also","even","only","into",
        "over","under","both","such","than","then","but","not","nor",
        "yet","so","either","neither","versus","vs",
    }

    _SQ_SYN = {
        "subscription": ["subscription","subscribed","subscript","subscriptionkw","sub","contracted"],
        "capacity": ["capacity","capac","totalcapacity","installedcapacity","total capacity","installed capacity"],
        "power": ["power","kw","kva","watt","watts","powerload","load","powerallocation"],
        "usage": ["usage","utilization","utilisation","consumption","consumed","actual","actualusage"],
        "rack": ["rack","racks","rackunits","rackunit","noofracks","numberofracks","no of rack"],
        "space": ["space","sqft","sq ft","sq.ft","area","floorspace","cabinets","cabinet"],
        "customer": ["customer","name","customers","client","clients","customername","customer name","company","organization","account"],
        "billing": ["billing","bill","invoice","billed","amount","monthlybilling"],
        "ownership": ["ownership","owned","owner","ownertype","owner type"],
        "revenue": ["revenue","income","earning","earnings","arr"],
        "location": ["location","site","dc","datacenter","data center","city","region"],
        "status": ["status","live","active","inactive","commissioned","customer status","operational"],
        "sno": ["sno","sr no","sr. no","serial","sl no","srno"],
        "sector": ["sector","industry","vertical","segment"],
        "hall": ["hall","floor","zone","wing","block"],
        "ups": ["ups","ups capacity","ups load"],
        "cooling": ["cooling","crac","precision cooling"],
    }

    _TYPE_FILTER_WORDS = [
        "caged","uncaged","metered","unmetered","bundled",
        "un-caged","shared","dedicated","colocation","colo",
        "managed","unmanaged","hosted","virtual",
    ]

    _SQ_OP_VERBS = {
        "total","sum","avg","mean","max","min","count","list","find",
        "show","all","average","maximum","minimum","highest","lowest",
        "top","bottom","describe","statistics","stats","summary","unique",
        "distinct","sheet","column","row","missing","null","percent",
        "percentage","ratio","share","number","across","compare","get",
        "give","what","which","where","how","tell","much","many","each",
        "every","display","view","see","provide",
    }

    def _sq_is_num(v):
        try:
            fv = float(str(v).replace(",", "").replace(" ", "").strip())
            return True, fv
        except Exception:
            return False, None

    def _sq_mcol(kw, hdr):
        hl = hdr.lower().strip()
        kwl = kw.lower().strip()
        if not kwl:
            return False
        if kwl in hl or hl in kwl:
            return True
        _strip = lambda s: re.sub(r"[\s\(\)\[\]\.,:;\-_/]", "", s)
        hl_c = _strip(hl)
        kwl_c = _strip(kwl)
        if kwl_c and (kwl_c in hl_c or hl_c in kwl_c):
            return True
        for key, syns in _SQ_SYN.items():
            if kwl in syns or kwl == key:
                for s in syns:
                    sc = _strip(s)
                    if sc and (sc in hl_c or hl_c in sc):
                        return True
        kw_words = {w for w in re.findall(r"[a-z0-9]+", kwl) if len(w) >= 3}
        hl_words = {w for w in re.findall(r"[a-z0-9]+", hl)}
        if kw_words and kw_words & hl_words:
            return True
        return False

    def _exact_type_match(cell_val, type_val):
        v = cell_val.lower().strip()
        t = type_val.lower().strip()
        if v == t:
            return True
        if " " in v or "/" in v or "," in v:
            tokens = re.split(r"[\s,/;|]+", v)
            if t in tokens:
                return True
        return False

    def _detect_type_vals(ql):
        quoted = re.findall(r'"([^"]+)"', ql)
        if quoted:
            return [t.strip().lower() for t in quoted]
        found = []
        for tw in _TYPE_FILTER_WORDS:
            if re.search(r'(?<![a-z])' + re.escape(tw) + r'(?![a-z])', ql):
                found.append(tw)
        return found

    def smart_corpus_query(question, use_all=True):
        q = question.strip()
        ql = q.lower()
        if use_all:
            wc = corpus
            rr = row_records
        else:
            wc = sq_cells
            rr = sq_rows

        out = {"answer": "", "table": None, "chart_df": None, "chart_cfg": None, "cell_hits": [], "sub_tables": []}
        if not wc:
            out["answer"] = "❓ No data indexed."
            return out

        sig = [w for w in re.findall(r"[a-z0-9]{2,}", ql) if w not in _SW2]

        f_sum  = any(x in ql for x in ["total","sum","aggregate","grand total","add up"])
        f_avg  = any(x in ql for x in ["average","mean","avg"])
        f_max  = any(x in ql for x in ["maximum","highest","largest","max","biggest","top value"])
        f_min  = any(x in ql for x in ["minimum","lowest","smallest","min","least"])
        f_cnt  = any(x in ql for x in ["count","how many","number of","how much"])
        f_stat = any(x in ql for x in ["statistics","stats","describe","summary","all stats","overview"])
        f_std  = any(x in ql for x in ["std","std dev","stddev","standard deviation"])
        f_med  = any(x in ql for x in ["median","middle"])
        f_uniq = any(x in ql for x in ["unique","distinct","different","variety"])
        f_miss = any(x in ql for x in ["missing","null","blank","empty","nan"])
        f_cols = any(x in ql for x in ["column","columns","field","header","headers","fields"])
        f_topn = re.search(r"\btop\s*(\d+)\b", ql)
        f_botn = re.search(r"\bbottom\s*(\d+)\b", ql)
        f_pct  = any(x in ql for x in ["percent","percentage","%","share","proportion"])
        f_num  = (f_sum or f_avg or f_max or f_min or f_cnt or f_stat or f_std or f_med
                  or bool(f_topn) or bool(f_botn))
        f_cust = any(x in ql for x in ["customer","customers","client","clients","name","company","account"])
        f_list = any(x in ql for x in ["list","show","display","all","get all"])

        _early_type_vals = _detect_type_vals(ql)

        location_filter = None
        if use_all:
            for loc_name in meta.get("locations", []):
                if loc_name.lower() in ql:
                    location_filter = loc_name
                    break

        col_kws = [w for w in sig if w not in _SQ_OP_VERBS]
        _type_word_set = set(_TYPE_FILTER_WORDS) | {re.sub(r"[^a-z0-9]","",t) for t in _TYPE_FILTER_WORDS}
        if _early_type_vals:
            col_kws = [w for w in col_kws if w not in _type_word_set]

        def _loc_filter(cells):
            if location_filter and use_all:
                return [c for c in cells if c.get("location","").lower() == location_filter.lower()]
            return cells

        def _rr_key(cell):
            if use_all:
                return (cell["file"], cell["location"], cell["sheet"], cell["row"])
            return cell["row"]

        def _get_row_data(cell):
            return rr.get(_rr_key(cell), {})

        def _npkw(kw, cells):
            res = []
            for cell in cells:
                if cell["is_header"]:
                    continue
                if _sq_mcol(kw, cell["col_header"]):
                    ok, fv2 = _sq_is_num(cell["value"])
                    if ok:
                        res.append((fv2, cell))
            return res

        def _npbest(kws, cells):
            bk, bp = None, []
            for w in kws:
                p = _npkw(w, cells)
                if len(p) > len(bp):
                    bp, bk = p, w
            return bk, bp

        def _build_rows_df(cells_list):
            seen_keys = set()
            recs = []
            for cell in cells_list:
                key = _rr_key(cell)
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                row_data = _get_row_data(cell)
                if not row_data:
                    continue
                rd = {}
                if use_all:
                    rd["Location"] = cell.get("location", "")
                    rd["Sheet"]    = cell.get("sheet", "")
                rd["Row #"] = cell["row"] + 1
                rd.update(row_data)
                recs.append(rd)
            return pd.DataFrame(recs) if recs else pd.DataFrame()

        def _val_match(terms, cells):
            matched = []
            for cell in cells:
                if cell["is_header"]:
                    continue
                val_l = cell["value"].lower()
                if any(t in val_l for t in terms):
                    matched.append(cell)
            return matched

        def _col_match_score(w, cells):
            return sum(1 for c in cells if c["is_header"] and _sq_mcol(w, c["value"]))

        def _get_customer_name_col(row_data):
            cust_kws = ["customer","name","client","company","account","organisation","organization"]
            for col_h, val in row_data.items():
                col_l = col_h.lower()
                if any(kw in col_l for kw in cust_kws):
                    if val and str(val).strip() not in ("nan","None",""):
                        return col_h, val
            return None, None

        # ── INTENT: Missing values ──────────────────────────────────
        if f_miss:
            dfc2 = to_numeric(smart_header(raw_df))
            mr_list = []
            for col in dfc2.columns:
                mc = int(dfc2[col].isna().sum())
                if mc > 0:
                    mr_list.append({"Column": col, "Missing Count": mc,
                                    "Missing %": f"{mc/max(len(dfc2),1)*100:.1f}%",
                                    "Non-Null": len(dfc2)-mc})
            if mr_list:
                tbl = pd.DataFrame(mr_list).sort_values("Missing Count", ascending=False)
                out["answer"] = f"Found **{len(tbl)}** column(s) with missing values in the current sheet."
                out["table"] = tbl
            else:
                out["answer"] = "✅ No missing values found in the current sheet."
            return out

        # ── INTENT: Column listing ──────────────────────────────────
        if f_cols and not f_num:
            src = _loc_filter(wc)
            seen_h, cr = set(), []
            for cell in src:
                if not cell["is_header"]:
                    continue
                ch = cell["value"].strip()
                if ch in ("","nan","None") or ch in seen_h:
                    continue
                seen_h.add(ch)
                entry = {"Column Header": ch}
                if use_all:
                    entry["Location"] = cell.get("location","")
                    entry["Sheet"]    = cell.get("sheet","")
                entry["At Row"] = cell["row"] + 1
                cr.append(entry)
            tbl = pd.DataFrame(cr) if cr else pd.DataFrame()
            scope_txt = "across all files" if use_all else "in this sheet"
            out["answer"] = f"Found **{len(tbl)}** unique column header(s) {scope_txt}."
            out["table"]  = tbl
            return out

        # ── INTENT: Type-filter (caged/metered/bundled/uncaged etc.) ──────────
        if _early_type_vals and not f_num and not f_cols and not f_miss and not f_uniq:
            src = _loc_filter(wc)
            scope_txt = "across all files" if use_all else "in this sheet"
            all_result_rows = []

            for tv in _early_type_vals:
                tv_row_keys = set()
                for cell in src:
                    if cell["is_header"]:
                        continue
                    if _exact_type_match(cell["value"], tv):
                        tv_row_keys.add(_rr_key(cell))

                for key in tv_row_keys:
                    row_data = rr.get(key, {})
                    if not row_data:
                        continue
                    rd = {"Type": tv}
                    if use_all and isinstance(key, tuple):
                        rd["Location"] = key[1]
                        rd["Sheet"]    = key[2]
                        rd["Row #"]    = key[3] + 1
                    else:
                        rd["Row #"] = (key+1) if isinstance(key, int) else (key[-1]+1)
                    cust_col, cust_val = _get_customer_name_col(row_data)
                    if cust_col and cust_val:
                        rd["Customer / Name"] = cust_val
                    rd.update(row_data)
                    all_result_rows.append(rd)

            if all_result_rows:
                result_df = pd.DataFrame(all_result_rows)
                priority_cols = [c for c in ["Type","Customer / Name","Location","Sheet","Row #"] if c in result_df.columns]
                other_cols = [c for c in result_df.columns if c not in priority_cols]
                result_df = result_df[priority_cols + other_cols]
                type_disp = ", ".join(f"**{v}**" for v in _early_type_vals)
                out["answer"] = f"Found **{len(all_result_rows):,}** row(s) matching type(s) {type_disp} {scope_txt}."
                out["table"] = result_df
            else:
                type_disp = ", ".join(f"'{v}'" for v in _early_type_vals)
                out["answer"] = f"No rows found with type value(s) {type_disp} {scope_txt}."
            return out

        # ── INTENT: List customers ──────────────────────────────────
        if (f_cust or f_list) and not f_num and not col_kws:
            src = _loc_filter(wc)
            cust_kws = ["customer","name","client","company","account","organisation","organization"]
            found_cells = [
                c for c in src
                if not c["is_header"]
                and any(x in c["col_header"].lower() for x in cust_kws)
            ]
            if found_cells:
                rows_data = []
                for cell in found_cells:
                    entry = {}
                    if use_all:
                        entry["Location"] = cell.get("location","")
                        entry["Sheet"]    = cell.get("sheet","")
                    entry["Row #"]  = cell["row"] + 1
                    entry["Column"] = cell["col_header"]
                    entry["Value"]  = cell["value"]
                    rows_data.append(entry)
                tbl = pd.DataFrame(rows_data).drop_duplicates(subset=["Value"])
                scope_txt = "across all files" if use_all else "in this sheet"
                out["answer"] = f"Found **{len(tbl)}** unique customer/name entries {scope_txt}."
                out["table"] = tbl
                full_rows_df = _build_rows_df(found_cells)
                if not full_rows_df.empty:
                    out["sub_tables"].append({"label": f"📋 Full Row Data for All Customers", "df": full_rows_df})
            else:
                out["answer"] = "No 'Customer/Name/Client' columns found in the data."
            return out

        # ── INTENT: Numeric aggregation ─────────────────────────────
        if f_num:
            src = _loc_filter(wc)
            kw, pairs = _npbest(col_kws, src)
            if not pairs:
                for dkw in ["subscription","capacity","power","usage","rack","space","revenue","billing","kw","kva","sqft"]:
                    if dkw in ql:
                        pairs = _npkw(dkw, src)
                        if pairs:
                            kw = dkw
                            break
            if pairs:
                vals = [v for v, _ in pairs]
                sa = pd.Series(vals)
                parts = []
                if f_sum  or f_stat: parts.append(f"**Total (Sum):** {sa.sum():,.2f}")
                if f_avg  or f_stat: parts.append(f"**Average (Mean):** {sa.mean():,.2f}")
                if f_max  or f_stat: parts.append(f"**Maximum:** {sa.max():,.2f}")
                if f_min  or f_stat: parts.append(f"**Minimum:** {sa.min():,.2f}")
                if f_cnt  or f_stat: parts.append(f"**Count:** {int(sa.count()):,}")
                if f_std  or f_stat: parts.append(f"**Std Dev:** {sa.std():,.2f}")
                if f_med  or f_stat: parts.append(f"**Median:** {sa.median():,.2f}")
                if f_stat:
                    parts.append(f"**Variance:** {sa.var():,.2f}  |  **Range:** {sa.max()-sa.min():,.2f}  |  **IQR:** {sa.quantile(0.75)-sa.quantile(0.25):,.2f}")
                if f_pct:
                    grand = sum(fv2 for c in src if not c["is_header"] for ok,fv2 in [_sq_is_num(c["value"])] if ok) or 1
                    parts.append(f"**% of All Numeric:** {sa.sum()/grand*100:.2f}%")

                detail = []
                for v, cell in pairs:
                    entry = {}
                    if use_all:
                        entry["Location"]      = cell.get("location","")
                        entry["Sheet"]         = cell.get("sheet","")
                    entry["Row #"]         = cell["row"] + 1
                    entry["Column Header"] = cell["col_header"]
                    entry["Value"]         = v
                    detail.append(entry)
                tbl = pd.DataFrame(detail).sort_values("Value", ascending=False)

                scope_txt = "across all files" if use_all else "in this sheet"
                loc_note  = f" (Location: **{location_filter}**)" if location_filter else ""
                out["answer"] = (
                    f"Results for **'{kw}'** — **{len(vals):,} values** found {scope_txt}{loc_note}:\n\n"
                    + "\n".join(parts)
                )
                out["table"] = tbl

                full_rows_df = _build_rows_df([cell for _,cell in pairs])
                if not full_rows_df.empty:
                    out["sub_tables"].append({"label": f"📋 Full Row Data for '{kw}'", "df": full_rows_df})

                if use_all and "Location" in tbl.columns and tbl["Location"].nunique() > 1:
                    loc_agg = tbl.groupby("Location")["Value"].sum().reset_index().rename(columns={"Value":"Total"})
                    out["chart_df"]  = loc_agg
                    out["chart_cfg"] = {"x":"Location","y":"Total","title":f"Sum of '{kw}' by Location"}

                if f_topn:
                    n = int(f_topn.group(1))
                    top = sorted(pairs, key=lambda x: x[0], reverse=True)[:n]
                    top_df = _build_rows_df([c for _,c in top])
                    out["sub_tables"].append({"label": f"🏆 Top {n} — '{kw}'", "df": top_df})
                if f_botn:
                    n = int(f_botn.group(1))
                    bot = sorted(pairs, key=lambda x: x[0])[:n]
                    bot_df = _build_rows_df([c for _,c in bot])
                    out["sub_tables"].append({"label": f"🔻 Bottom {n} — '{kw}'", "df": bot_df})
                return out

        # ── INTENT: Unique values ──────────────────────────────────
        if f_uniq:
            src = _loc_filter(wc)
            for w in (col_kws or sig):
                uv, sr = set(), []
                for cell in src:
                    if cell["is_header"]:
                        continue
                    if _sq_mcol(w, cell["col_header"]):
                        uv.add(cell["value"])
                        entry = {"Column": cell["col_header"], "Value": cell["value"]}
                        if use_all:
                            entry["Location"] = cell.get("location","")
                            entry["Sheet"]    = cell.get("sheet","")
                        entry["Row #"] = cell["row"] + 1
                        sr.append(entry)
                if uv:
                    tbl = pd.DataFrame(sr).drop_duplicates(subset=["Value"])
                    out["answer"] = f"**{len(uv)}** unique value(s) found for **'{w}'**."
                    out["table"]  = tbl
                    return out

        # ── Cross-column / entity lookup ─────────────────────────────
        src     = _loc_filter(wc)
        tokens  = [w for w in sig if w not in _SQ_OP_VERBS]
        col_scores  = {w: _col_match_score(w, src) for w in tokens}
        val_scores  = {w: sum(1 for c in src if not c["is_header"] and w in c["value"].lower()) for w in tokens}
        attr_toks   = [w for w in tokens if col_scores.get(w, 0) > 0]
        entity_toks = [w for w in tokens if val_scores.get(w, 0) > 0]

        quoted = re.findall(r'"([^"]+)"', q)
        entity_cells = _val_match([quoted[0].lower()], src) if quoted else _val_match(entity_toks, src)

        if attr_toks and entity_cells:
            entity_keys = {_rr_key(c) for c in entity_cells}
            recs = []
            for key in sorted(entity_keys, key=lambda k: (k if isinstance(k, int) else k[-1])):
                row_data = rr.get(key, {})
                if not row_data:
                    continue
                rd = {}
                if use_all and isinstance(key, tuple):
                    rd["Location"] = key[1]
                    rd["Sheet"]    = key[2]
                    rd["Row #"]    = key[3] + 1
                else:
                    rd["Row #"] = (key+1) if isinstance(key,int) else key[-1]+1
                for cell in entity_cells:
                    if _rr_key(cell) == key:
                        rd[f"[Matched] {cell['col_header']}"] = cell["value"]
                for at in attr_toks:
                    for col_h, val in row_data.items():
                        if _sq_mcol(at, col_h):
                            rd[col_h] = val
                recs.append(rd)
            if recs:
                tbl = pd.DataFrame(recs)
                entity_disp = ", ".join(f"'{t}'" for t in (entity_toks if not quoted else [quoted[0]])[:3])
                attr_disp = ", ".join(f"'{t}'" for t in attr_toks[:3])
                out["answer"] = (
                    f"Cross-column lookup — entity **{entity_disp}** found in **{len(entity_keys)}** row(s).\n\n"
                    f"Showing **{attr_disp}** column value(s):"
                )
                out["table"] = tbl
                full_rows_df = _build_rows_df(entity_cells)
                if not full_rows_df.empty:
                    out["sub_tables"].append({"label": f"📋 All Columns for Matching Rows", "df": full_rows_df})
                return out

        if entity_cells:
            full_df = _build_rows_df(entity_cells)
            entity_keys = {_rr_key(c) for c in entity_cells}
            match_disp = f'"{quoted[0]}"' if quoted else ", ".join(f"'{t}'" for t in entity_toks[:4])
            out["answer"] = (
                f"Found **{len(entity_cells):,}** cell(s) matching **{match_disp}** "
                f"across **{len(entity_keys):,}** row(s).\n\nShowing **all column values** for every matching row:"
            )
            out["table"] = full_df if not full_df.empty else None
            out["cell_hits"] = [
                dict(
                    **({} if not use_all else {"Location": c.get("location",""), "Sheet": c.get("sheet","")}),
                    **{"Row #": c["row"]+1, "Col #": c["col"]+1, "Column Header": c["col_header"], "Value": c["value"]},
                )
                for c in entity_cells[:200]
            ]
            return out

        if attr_toks:
            best_attr = max(attr_toks, key=lambda w: col_scores[w])
            sr, seen_keys = [], set()
            for cell in src:
                if cell["is_header"] or not _sq_mcol(best_attr, cell["col_header"]):
                    continue
                key = _rr_key(cell)
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                entry = {}
                if use_all:
                    entry["Location"] = cell.get("location","")
                    entry["Sheet"]    = cell.get("sheet","")
                entry["Row #"]  = cell["row"] + 1
                entry["Column"] = cell["col_header"]
                entry["Value"]  = cell["value"]
                sr.append(entry)
            if sr:
                tbl = pd.DataFrame(sr)
                out["answer"] = f"Found **{len(sr)}** value(s) in column(s) matching **'{best_attr}'**."
                out["table"]  = tbl
                matching_cells = [c for c in src if not c["is_header"] and _sq_mcol(best_attr, c["col_header"])]
                full_rows_df = _build_rows_df(matching_cells)
                if not full_rows_df.empty:
                    out["sub_tables"].append({"label": f"📋 Full Row Data for '{best_attr}'", "df": full_rows_df})
                return out

        if f_list or "all rows" in ql or "show all" in ql:
            seen_keys, recs = set(), []
            for cell in src:
                if cell["is_header"]:
                    continue
                key = _rr_key(cell)
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                row_data = rr.get(key, {})
                if not row_data:
                    continue
                rd = {}
                if use_all and isinstance(key, tuple):
                    rd["Location"] = key[1]
                    rd["Sheet"]    = key[2]
                    rd["Row #"]    = key[3]+1
                else:
                    rd["Row #"] = (key+1) if isinstance(key,int) else key[-1]+1
                rd.update(row_data)
                recs.append(rd)
            if recs:
                tbl = pd.DataFrame(recs)
                scope_txt = "across all files" if use_all else "in this sheet"
                out["answer"] = f"Showing **all {len(recs):,}** data rows {scope_txt}."
                out["table"]  = tbl
                return out

        out["answer"] = "❓ No matching data found. Try: 'total power', 'list caged customers', 'count racks', 'sum capacity Airoli'."
        return out

    # ── Render query results ───────────────────────────────────────
    def sq_render_answer(res, tidx=0):
        st.markdown(f'<div class="ans-box">{res["answer"]}</div>', unsafe_allow_html=True)
        if res.get("table") is not None and not res["table"].empty:
            tbl = res["table"].reset_index(drop=True)
            st.dataframe(tbl, use_container_width=True, height=min(560, 50+len(tbl)*36), key=f"sq_tbl_{tidx}")
            st.download_button("⬇️ Download CSV", tbl.to_csv(index=False).encode(),
                               "smart_query_result.csv", "text/csv", key=f"sq_dl_{tidx}")
        if res.get("chart_cfg") and res.get("chart_df") is not None:
            cfg = res["chart_cfg"]
            cdf = res["chart_df"]
            if cfg["x"] in cdf.columns and cfg["y"] in cdf.columns:
                try:
                    fig = px.bar(
                        cdf.sort_values(cfg["y"], ascending=False).head(30),
                        x=cfg["x"], y=cfg["y"], color=cfg["y"],
                        color_continuous_scale="Viridis",
                        title=cfg["title"], height=400,
                    )
                    fig.update_layout(xaxis_tickangle=-30, **DARK)
                    st.plotly_chart(fig, use_container_width=True, key=f"sq_ch_{tidx}")
                except Exception:
                    pass
        for si, s in enumerate(res.get("sub_tables", [])):
            with st.expander(s["label"], expanded=True):
                st.dataframe(s["df"], use_container_width=True, key=f"sq_sub_{tidx}_{si}")
        if res.get("cell_hits"):
            with st.expander(f"🔬 Matching Cells ({len(res['cell_hits'])})", expanded=False):
                for ch in res["cell_hits"][:100]:
                    loc_info = ""
                    if ch.get("Location"):
                        loc_info = f'{ch["Location"]} › {ch.get("Sheet","")} | '
                    st.markdown(
                        f'<div class="cell-chip">{loc_info}'
                        f'Row {ch["Row #"]} · Col {ch["Col #"]} | '
                        f'<i>{ch["Column Header"]}</i> | '
                        f'<b>{ch["Value"]}</b></div>',
                        unsafe_allow_html=True,
                    )
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)

    # ── Query input ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="sec-title">🧠 Natural Language Query</div>', unsafe_allow_html=True)
    st.caption(
        "Examples: **total power Airoli** · **list caged customers** · "
        "**count racks Noida** · **avg capacity all locations** · "
        "**sum subscription metered** · **statistics power** · **top 5 capacity**"
    )

    q_examples = [
        "— Choose an example —",
        "total power across all locations",
        "list all caged customers",
        "sum of subscription for metered type",
        "count racks across all files",
        "average capacity by location",
        "statistics of power column",
        "top 10 capacity rows",
        "list unique customer names",
        "missing values",
        "show all columns",
    ]
    ex_sel = st.selectbox("Quick examples:", q_examples, key="sq_ex")

    if "sq_history" not in st.session_state:
        st.session_state.sq_history = []

    default_q = ex_sel if ex_sel != "— Choose an example —" else ""
    user_q = st.text_input("💬 Your query:", value=default_q, placeholder="e.g. total power all locations", key="sq_in")

    col_run_q, col_clr = st.columns([1, 1])
    run_q = col_run_q.button("🔍 Run Query", type="primary", key="sq_run")
    clr_q = col_clr.button("🗑 Clear History", key="sq_clr")

    if clr_q:
        st.session_state.sq_history = []
        st.rerun()

    if run_q and user_q.strip():
        with st.spinner("🔍 Querying…"):
            res = smart_corpus_query(user_q.strip(), use_all=use_corpus)
        st.session_state.sq_history.insert(0, (user_q.strip(), res))

    for i, (qtext, res) in enumerate(st.session_state.sq_history[:10]):
        st.markdown(f'<div class="q-user">💬 {qtext}</div>', unsafe_allow_html=True)
        sq_render_answer(res, tidx=i)

    # ── SMART CELL VALUE SEARCH ────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="sec-title">🔎 Smart Cell-Value Search</div>', unsafe_allow_html=True)
    st.caption(
        "Search every file, sheet, row and column by any cell value. "
        "Returns all matching rows with their full column context."
    )

    scs1, scs2, scs3 = st.columns([3, 2, 1])
    cell_search_val   = scs1.text_input("Search value:", placeholder="e.g. Airtel, caged, 1000", key="csv_val")
    cell_match_mode   = scs2.selectbox("Match mode:", [
        "Contains (partial match)", "Exact (case-insensitive)", "Starts with", "Regex"
    ], key="csv_mode")
    cell_search_scope = scs3.checkbox("All files", True, key="csv_scope")

    if st.button("🔎 Search Cell Values", type="primary", key="csv_run"):
        if cell_search_val.strip():
            sv = cell_search_val.strip()
            svl = sv.lower()
            src2 = corpus if cell_search_scope else sq_cells
            hits, row_keys, seen_k = [], [], set()

            for cell in src2:
                if cell["is_header"]:
                    continue
                val = cell["value"]
                vl  = val.lower()
                matched = False
                if cell_match_mode == "Exact (case-insensitive)":
                    matched = (vl == svl)
                elif cell_match_mode == "Contains (partial match)":
                    matched = (svl in vl)
                elif cell_match_mode == "Starts with":
                    matched = vl.startswith(svl)
                elif cell_match_mode == "Regex":
                    try:
                        matched = bool(re.search(sv, val, re.IGNORECASE))
                    except Exception:
                        matched = False
                if matched:
                    h = {}
                    if cell_search_scope and "location" in cell:
                        h["Location"] = cell.get("location","")
                        h["Sheet"]    = cell.get("sheet","")
                    h["Row #"]         = cell["row"] + 1
                    h["Col #"]         = cell["col"] + 1
                    h["Column Header"] = cell["col_header"]
                    h["Matched Value"] = val
                    hits.append(h)
                    if cell_search_scope:
                        key2 = (cell["file"], cell.get("location",""), cell.get("sheet",""), cell["row"])
                    else:
                        key2 = cell["row"]
                    if key2 not in seen_k:
                        seen_k.add(key2)
                        row_keys.append(key2)

            if hits:
                hits_df = pd.DataFrame(hits)
                st.success(f"Found **{len(hits):,}** matching cell(s) across **{len(row_keys):,}** row(s).")
                st.dataframe(hits_df, use_container_width=True, height=400)
                st.download_button("⬇️ Download Matching Cells CSV", hits_df.to_csv(index=False).encode(),
                                   "cell_search_results.csv", "text/csv")

                # Full rows for matching keys
                full_recs = []
                for key2 in row_keys[:500]:
                    rd = row_records.get(key2, {}) if cell_search_scope else sq_rows.get(key2, {})
                    if rd:
                        rec = {}
                        if cell_search_scope and isinstance(key2, tuple):
                            rec["Location"] = key2[1]
                            rec["Sheet"]    = key2[2]
                            rec["Row #"]    = key2[3] + 1
                        elif isinstance(key2, int):
                            rec["Row #"] = key2 + 1
                        rec.update(rd)
                        full_recs.append(rec)
                if full_recs:
                    with st.expander(f"📋 Full Row Data for All Matching Rows ({len(full_recs)})", expanded=True):
                        full_df2 = pd.DataFrame(full_recs)
                        st.dataframe(full_df2, use_container_width=True)
                        st.download_button("⬇️ Download Full Rows CSV", full_df2.to_csv(index=False).encode(),
                                           "cell_search_full_rows.csv", "text/csv")
            else:
                st.warning(f"No cells found matching **'{sv}'** with mode **{cell_match_mode}**.")
        else:
            st.warning("Please enter a search value.")
