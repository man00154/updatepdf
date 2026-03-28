import os, re, warnings, tempfile, subprocess
from collections import defaultdict
 
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
 
warnings.filterwarnings("ignore")
 
# ──────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sify DC – Capacity Tracker",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)
 
# ──────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0e1a,#1a2035,#0d1b2a)!important;}
[data-testid="stSidebar"] *{color:#c9d8f0!important;}
.kcard{border-radius:16px;padding:20px 24px;color:#fff;margin-bottom:12px;
       box-shadow:0 6px 24px rgba(0,0,0,.35);transition:transform .2s;}
.kcard:hover{transform:translateY(-3px);}
.kcard h2{font-size:2.1rem;margin:0;font-weight:800;}
.kcard p{margin:4px 0 0;font-size:.85rem;opacity:.82;letter-spacing:.4px;}
.kcard-blue  {background:linear-gradient(135deg,#1e3c72,#2a5298);}
.kcard-green {background:linear-gradient(135deg,#0b6e4f,#17a572);}
.kcard-red   {background:linear-gradient(135deg,#7b1a1a,#c0392b);}
.kcard-orange{background:linear-gradient(135deg,#7d4e00,#e67e22);}
.kcard-teal  {background:linear-gradient(135deg,#0f3460,#16213e);}
.kcard-purple{background:linear-gradient(135deg,#4a0072,#7b1fa2);}
.sec-title{font-size:1.25rem;font-weight:700;color:#1e3c72;
    border-left:5px solid #2a5298;padding-left:10px;margin:20px 0 12px;}
/* AI chat */
.q-user{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff;
    border-radius:20px 20px 4px 20px;padding:12px 18px;margin:10px 0 4px auto;
    max-width:76%;width:fit-content;font-size:.97rem;
    box-shadow:0 4px 14px rgba(30,60,114,.45);float:right;clear:both;}
.q-bot{background:linear-gradient(135deg,#0d2137,#1a4060);color:#c9e8ff;
    border-radius:20px 20px 20px 4px;padding:12px 18px;margin:4px auto 10px 0;
    max-width:90%;font-size:.95rem;
    box-shadow:0 4px 14px rgba(0,80,160,.30);clear:both;float:left;}
.ans-box{background:linear-gradient(135deg,#0f2744,#1a4a6b);color:#d0ecff;
    border-radius:14px;padding:16px 20px;margin:8px 0;font-size:1rem;
    box-shadow:0 4px 16px rgba(0,0,0,.35);white-space:pre-wrap;line-height:1.65;}
.cell-chip{background:#1a2f1a;border-left:4px solid #27ae60;border-radius:7px;
    padding:7px 13px;margin:4px 0;font-family:monospace;font-size:.82rem;color:#b8ffb8;}
.clearfix{clear:both;}
</style>
""", unsafe_allow_html=True)
 
# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────────────────────
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "excel_files")
 
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
 
# ──────────────────────────────────────────────────────────────────────────────
# XLS → XLSX AUTO-CONVERSION  (LibreOffice, handles true OLE2 .xls)
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def ensure_readable(path: str) -> str:
    """Return xlsx path; converts OLE2 .xls via LibreOffice if needed."""
    if not path.lower().endswith(".xls"):
        return path
    try:
        with open(path, "rb") as fh:
            if fh.read(4) != b"\xd0\xcf\x11\xe0":
                return path
    except Exception:
        return path
    out_dir  = tempfile.mkdtemp()
    basename = os.path.splitext(os.path.basename(path))[0]
    out_path = os.path.join(out_dir, basename + ".xlsx")
    try:
        subprocess.run(
            ["libreoffice","--headless","--convert-to","xlsx","--outdir",out_dir,path],
            capture_output=True, timeout=90, check=True,
        )
        return out_path
    except Exception:
        return path
 
# ──────────────────────────────────────────────────────────────────────────────
# FILE HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def find_excel_files(folder):
    if not os.path.isdir(folder):
        return []
    return sorted(f for f in os.listdir(folder) if f.lower().endswith((".xlsx",".xls")))
 
def location_from_name(fname):
    n = fname.replace("Customer_and_Capacity_Tracker_","").replace(".xlsx","").replace(".xls","")
    n = re.sub(r"_\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{2,4}.*$","",n,flags=re.I)
    return n.replace("_"," ").strip()
 
# ──────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_file(original_path):
    path   = ensure_readable(original_path)
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
        st.sidebar.warning(f"⚠️ Cannot read {os.path.basename(original_path)}: {e}")
    return sheets
 
def best_header_row(df):
    best_row, best_score = 0, -1
    for i in range(min(8, len(df))):
        row    = df.iloc[i].astype(str).str.strip()
        filled = (row.str.len() > 0) & (~row.isin(["nan","None"]))
        label  = filled & (~row.str.match(r"^-?\d+(\.\d+)?$"))
        score  = label.sum() * 2 + filled.sum()
        if score > best_score:
            best_score, best_row = score, i
    return best_row
 
def smart_header(df):
    hr   = best_header_row(df)
    hdr  = df.iloc[hr].fillna("").astype(str).str.strip()
    seen = {}; cols = []
    for col in hdr:
        col = col if col and col not in ("nan","None") else f"Col_{len(cols)}"
        if col in seen:
            seen[col] += 1; cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0; cols.append(col)
    data = df.iloc[hr+1:].copy()
    data.columns = cols
    return data.dropna(how="all").reset_index(drop=True)
 
def to_numeric(df):
    out = df.copy()
    for col in out.columns:
        out[col] = pd.to_numeric(out[col], errors="ignore")
    return out
 
@st.cache_data(show_spinner=False)
def save_uploads(files):
    tmp = tempfile.mkdtemp()
    for f in files:
        with open(os.path.join(tmp, f.name),"wb") as fh:
            fh.write(f.read())
    return tmp
 
# ──────────────────────────────────────────────────────────────────────────────
# BUILD FULL CORPUS — every non-empty cell of every row×col×sheet×file
# Handles multi-section sheets (e.g. Chennai, Vashi) where header rows appear
# multiple times inside the sheet at different positions.
# ──────────────────────────────────────────────────────────────────────────────
 
def _detect_header_rows(df) -> set:
    """
    Find ALL rows that look like column-header rows.
    Criteria: ≥70 % of filled cells are text labels (not pure numbers),
              at least 2 filled cells, min 2 distinct values.
    Used for multi-section sheets where the header repeats mid-sheet.
    """
    hr_set = set()
    for i in range(len(df)):
        row    = df.iloc[i].astype(str).str.strip()
        filled = (row.str.len() > 0) & (~row.isin(["nan", "None", ""]))
        label  = filled & (~row.str.match(r"^-?\d+(\.\d+)?$"))
        if filled.sum() < 2:
            continue
        if label.sum() / filled.sum() >= 0.70 and row[filled].nunique() >= 2:
            hr_set.add(i)
    return hr_set
 
 
def _build_cell_col_map(df) -> dict:
    """
    For every (row, col) in df, return the best column-header name.
    Walks each column top-to-bottom; updates the header name whenever
    it encounters a header row.  This correctly handles sheets that have
    multiple header sections (summary block + customer-list block, etc.).
    """
    hr_set = _detect_header_rows(df)
 
    # Build per-header-row col→name maps
    hr_maps = {}
    for hr in hr_set:
        m = {}
        for c in range(df.shape[1]):
            v = str(df.iat[hr, c]).strip()
            if v and v not in ("nan", "None"):
                m[c] = v
        hr_maps[hr] = m
 
    # For each cell, find the nearest header row above (or equal to) it
    sorted_hrs = sorted(hr_set)
    cell_map   = {}
    for r in range(df.shape[0]):
        prev = [h for h in sorted_hrs if h <= r]
        for c in range(df.shape[1]):
            name = f"Col_{c}"
            for h in reversed(prev):           # most-recent header first
                if c in hr_maps[h]:
                    name = hr_maps[h][c]
                    break
            cell_map[(r, c)] = name
    return cell_map, hr_set
 
 
@st.cache_data(show_spinner=True)
def build_corpus(file_list, folder):
    """
    Reads EVERY non-empty cell from EVERY row × column × sheet × file.
    • Handles unstructured / odd layouts — no row or column is skipped.
    • Handles multi-section sheets (Chennai, Vashi) via _build_cell_col_map.
    • Auto-converts .xls → .xlsx via LibreOffice before reading.
    • Zero Excel formulas used anywhere.
    """
    corpus      = []
    row_records = defaultdict(dict)
 
    for fname in file_list:
        path = ensure_readable(os.path.join(folder, fname))
        loc  = location_from_name(fname)
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
 
            # Smart per-cell column-name map (handles multi-section sheets)
            cell_map, hr_set = _build_cell_col_map(df)
 
            for r in range(df.shape[0]):
                for c in range(df.shape[1]):
                    v = str(df.iat[r, c]).strip()
                    if not v or v in ("nan", "None"):
                        continue
                    ch       = cell_map.get((r, c), f"Col_{c}")
                    is_hdr   = (r in hr_set)
                    key      = (fname, loc, sh, r)
                    corpus.append({
                        "file":       fname,
                        "location":   loc,
                        "sheet":      sh,
                        "row":        r,
                        "col":        c,
                        "col_header": ch,
                        "value":      v,
                        "is_header":  is_hdr,
                    })
                    if not is_hdr:
                        row_records[key][ch] = v
 
    meta = {
        "total_cells":  len(corpus),
        "total_files":  len({x["file"]              for x in corpus}),
        "total_sheets": len({(x["file"], x["sheet"]) for x in corpus}),
        "total_rows":   len(row_records),
        "locations":    sorted({x["location"] for x in corpus}),
    }
    return corpus, dict(row_records), meta
 
# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/fluency/96/data-center.png", width=72)
st.sidebar.title("🏢 Capacity Tracker")
st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Source")
uploaded = st.sidebar.file_uploader(
    "Upload Excel files (overrides folder)",
    type=["xlsx","xls"], accept_multiple_files=True,
)
data_dir    = save_uploads(tuple(uploaded)) if uploaded else UPLOAD_DIR
excel_files = find_excel_files(data_dir)
 
if not excel_files:
    st.warning(
        "⚠️ **No Excel files found.**\n\n"
        "• Upload files via the sidebar, **or**\n"
        "• Place them in `excel_files/` next to `app.py`."
    )
    st.stop()
 
loc_map = {f: location_from_name(f) for f in excel_files}
 
st.sidebar.subheader("🏙️ Location")
selected_file  = st.sidebar.selectbox("Location", excel_files, format_func=lambda x: loc_map[x])
all_sheets     = load_file(os.path.join(data_dir, selected_file))
st.sidebar.subheader("📋 Sheet")
selected_sheet = st.sidebar.selectbox("Sheet", list(all_sheets.keys()))
 
raw_df   = all_sheets[selected_sheet]
df_clean = to_numeric(smart_header(raw_df))
num_cols = df_clean.select_dtypes(include="number").columns.tolist()
cat_cols = [c for c in df_clean.columns if c not in num_cols]
 
st.sidebar.markdown("---")
st.sidebar.caption(f"📊 {len(num_cols)} numeric · {len(df_clean)} rows · {len(excel_files)} files")
 
# ──────────────────────────────────────────────────────────────────────────────
# TABS
# ──────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "🏠 Overview","📋 Raw Data","📊 Analytics","📈 Charts",
    "🥧 Distributions","🔍 Query Engine","🌍 Multi-Location",
    "🤖 AI Agent","💬 AI Smart Query",
])
loc_label = loc_map[selected_file]
 
# ═══════════════════════════════════════════════════════════════════
# TAB 0 – OVERVIEW
# ═══════════════════════════════════════════════════════════════════
with tabs[0]:
    st.title(f"🏢 {loc_label}  ›  {selected_sheet}")
    st.caption(f"File: `{selected_file}`  |  Raw: {raw_df.shape[0]}×{raw_df.shape[1]}  |  Clean: {len(df_clean)}×{len(df_clean.columns)}")
 
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.markdown(f'<div class="kcard kcard-blue"><h2>{len(df_clean)}</h2><p>Data Rows</p></div>',unsafe_allow_html=True)
    c2.markdown(f'<div class="kcard kcard-green"><h2>{len(df_clean.columns)}</h2><p>Columns</p></div>',unsafe_allow_html=True)
    c3.markdown(f'<div class="kcard kcard-purple"><h2>{len(num_cols)}</h2><p>Numeric Cols</p></div>',unsafe_allow_html=True)
    c4.markdown(f'<div class="kcard kcard-orange"><h2>{len(excel_files)}</h2><p>Files Loaded</p></div>',unsafe_allow_html=True)
    c5.markdown(f'<div class="kcard kcard-teal"><h2>{raw_df.shape[1]}</h2><p>Raw Columns</p></div>',unsafe_allow_html=True)
    c6.markdown(f'<div class="kcard kcard-red"><h2>{int(df_clean.isna().sum().sum())}</h2><p>Missing Cells</p></div>',unsafe_allow_html=True)
 
    st.markdown("---")
    if num_cols:
        st.markdown('<div class="sec-title">📐 Quick Statistics</div>',unsafe_allow_html=True)
        stats = df_clean[num_cols].describe().T
        stats["range"] = stats["max"] - stats["min"]
        stats["cv%"]   = (stats["std"]/stats["mean"].replace(0,np.nan)*100).round(1)
        st.dataframe(
            stats.style.format("{:.3f}",na_rep="—")
                 .background_gradient(cmap="Blues",subset=["mean","max"])
                 .background_gradient(cmap="Oranges",subset=["std"]),
            use_container_width=True)
 
    st.markdown('<div class="sec-title">🗂️ Column Overview</div>',unsafe_allow_html=True)
    ci = pd.DataFrame({
        "Column":   df_clean.columns,
        "Type":     df_clean.dtypes.values,
        "Non-Null": df_clean.notna().sum().values,
        "Null%":    (df_clean.isna().mean()*100).round(1).values,
        "Unique":   [df_clean[c].nunique() for c in df_clean.columns],
        "Sample":   [str(df_clean[c].dropna().iloc[0])[:60] if df_clean[c].dropna().shape[0]>0 else "—" for c in df_clean.columns],
    })
    st.dataframe(ci, use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════
# TAB 1 – RAW DATA
# ═══════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("📋 Full Data Table  (smart-header applied)")
    srch = st.text_input("🔍 Live search","",key="rawsrch")
    disp = df_clean[df_clean.apply(lambda col:col.astype(str).str.contains(srch,case=False,na=False)).any(axis=1)] if srch else df_clean
    st.caption(f"Showing {len(disp):,} / {len(df_clean):,} rows")
    st.dataframe(disp, use_container_width=True, height=520)
    st.download_button("⬇️ Download CSV", disp.to_csv(index=False).encode(), "export.csv","text/csv")
    st.markdown("---")
    st.subheader("🗃️ Raw Excel (no header processing)")
    st.dataframe(raw_df, use_container_width=True, height=300)
 
# ═══════════════════════════════════════════════════════════════════
# TAB 2 – ANALYTICS
# ═══════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("📊 Column Analytics")
    if not num_cols:
        st.info("No numeric columns in this sheet.")
    else:
        chosen = st.multiselect("Select columns",num_cols,default=num_cols[:min(6,len(num_cols))])
        if chosen:
            sub = df_clean[chosen].dropna(how="all")
            kc  = st.columns(min(len(chosen),6))
            for i,col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s): kc[i].metric(col[:22],f"{s.sum():,.1f}",f"avg {s.mean():,.1f}")
            st.markdown("---")
            agg_rows=[]
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    grand = df_clean[chosen].select_dtypes("number").sum().sum()
                    agg_rows.append({"Column":col,"Count":int(s.count()),"Sum":s.sum(),"Mean":s.mean(),
                        "Median":s.median(),"Min":s.min(),"Max":s.max(),"Std":s.std(),
                        "Var":s.var(),"25%":s.quantile(.25),"75%":s.quantile(.75),
                        "% Total":f"{s.sum()/grand*100:.1f}%" if grand else "—"})
            if agg_rows:
                adf = pd.DataFrame(agg_rows).set_index("Column")
                st.dataframe(adf.style.format("{:,.2f}",na_rep="—",
                    subset=[c for c in adf.columns if c!="% Total"])
                    .background_gradient(cmap="YlOrRd",subset=["Sum","Max"]),
                    use_container_width=True)
        st.markdown("---")
        st.markdown('<div class="sec-title">🧮 Group-By Aggregation</div>',unsafe_allow_html=True)
        gc1,gc2,gc3 = st.columns(3)
        all_cat = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique()<60]
        if all_cat and num_cols:
            gc = gc1.selectbox("Group by",all_cat)
            ac = gc2.selectbox("Aggregate",num_cols)
            af = gc3.selectbox("Function",["sum","mean","count","min","max","median","std"])
            grp = df_clean.groupby(gc)[ac].agg(af).reset_index()
            grp.columns = [gc,f"{af}({ac})"]
            grp = grp.sort_values(grp.columns[1],ascending=False)
            st.dataframe(grp,use_container_width=True)
            fig=px.bar(grp,x=gc,y=grp.columns[1],color=grp.columns[1],
                color_continuous_scale="Viridis",title=f"{af.title()} of {ac} by {gc}")
            fig.update_layout(xaxis_tickangle=-35,height=420)
            st.plotly_chart(fig,use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════
# TAB 3 – CHARTS
# ═══════════════════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("📈 Interactive Charts")
    ctype = st.selectbox("Chart Type",[
        "Bar Chart","Grouped Bar","Line Chart","Scatter Plot","Area Chart",
        "Bubble Chart","Heatmap (Correlation)","Box Plot","Funnel Chart",
        "Waterfall / Cumulative","3-D Scatter",
    ])
    if not num_cols:
        st.info("No numeric columns available.")
    else:
        def _s(label,opts,idx=0,key=None):
            return st.selectbox(label,opts,index=min(idx,max(0,len(opts)-1)),key=key)
 
        if ctype=="Bar Chart":
            xc=_s("X",cat_cols or df_clean.columns.tolist(),key="bx")
            yc=_s("Y",num_cols,key="by")
            ori=st.radio("Orientation",["Vertical","Horizontal"],horizontal=True)
            d=df_clean[[xc,yc]].dropna()
            fig=px.bar(d,x=xc if ori=="Vertical" else yc,y=yc if ori=="Vertical" else xc,
                color=yc,color_continuous_scale="Turbo",
                orientation="v" if ori=="Vertical" else "h",title=f"{yc} by {xc}")
            fig.update_layout(height=500); st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Grouped Bar":
            xc=_s("X",cat_cols or df_clean.columns.tolist(),key="gbx")
            ycs=st.multiselect("Y columns",num_cols,default=num_cols[:3])
            if ycs:
                fig=px.bar(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs,
                    barmode="group",title=f"Grouped Bar")
                fig.update_layout(height=480); st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Line Chart":
            xc=_s("X",df_clean.columns.tolist(),key="lx")
            ycs=st.multiselect("Y",num_cols,default=num_cols[:2])
            if ycs:
                fig=px.line(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),
                    x=xc,y=ycs,markers=True,title="Line Chart")
                fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Scatter Plot":
            xc=_s("X",num_cols,0,"sc_x"); yc=_s("Y",num_cols,1,"sc_y")
            sc=_s("Size",["None"]+num_cols,key="sc_s"); cc=_s("Color",["None"]+cat_cols+num_cols,key="sc_c")
            d=df_clean.dropna(subset=[xc,yc])
            fig=px.scatter(d,x=xc,y=yc,size=sc if sc!="None" else None,
                color=cc if cc!="None" else None,hover_data=d.columns[:6].tolist(),
                color_continuous_scale="Rainbow",title=f"{yc} vs {xc}")
            fig.update_layout(height=500); st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Area Chart":
            xc=_s("X",df_clean.columns.tolist(),key="ax")
            ycs=st.multiselect("Y",num_cols,default=num_cols[:3])
            if ycs:
                fig=px.area(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs)
                fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Bubble Chart":
            if len(num_cols)>=3:
                xc=_s("X",num_cols,0,"bu_x"); yc=_s("Y",num_cols,1,"bu_y"); sz=_s("Size",num_cols,2,"bu_s")
                lc=_s("Color",["None"]+cat_cols,key="bu_c")
                d=df_clean[[xc,yc,sz]].dropna()
                if lc!="None": d[lc]=df_clean[lc]
                fig=px.scatter(d,x=xc,y=yc,size=sz,color=lc if lc!="None" else None,
                    size_max=70,color_discrete_sequence=px.colors.qualitative.Vivid)
                fig.update_layout(height=520); st.plotly_chart(fig,use_container_width=True)
            else: st.info("Need ≥ 3 numeric columns.")
 
        elif ctype=="Heatmap (Correlation)":
            sel=st.multiselect("Columns",num_cols,default=num_cols[:12])
            if len(sel)>=2:
                fig=px.imshow(df_clean[sel].corr(),text_auto=".2f",
                    color_continuous_scale="RdBu_r",aspect="auto",title="Correlation Heatmap")
                fig.update_layout(height=560); st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Box Plot":
            yc=_s("Value",num_cols,key="bp_v"); xc=_s("Group",["None"]+cat_cols,key="bp_g")
            d=df_clean[[yc]+([xc] if xc!="None" else [])].dropna(subset=[yc])
            fig=px.box(d,y=yc,x=xc if xc!="None" else None,color=xc if xc!="None" else None,
                points="outliers",color_discrete_sequence=px.colors.qualitative.Pastel)
            fig.update_layout(height=460); st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Funnel Chart":
            xc=_s("Stage",cat_cols or df_clean.columns.tolist(),key="fn_x"); yc=_s("Value",num_cols,key="fn_y")
            d=df_clean[[xc,yc]].dropna().groupby(xc)[yc].sum().reset_index().sort_values(yc,ascending=False)
            fig=px.funnel(d,x=yc,y=xc); fig.update_layout(height=460)
            st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="Waterfall / Cumulative":
            yc=_s("Column",num_cols,key="wf_y")
            d=df_clean[yc].dropna().reset_index(drop=True); cum=d.cumsum()
            fig=go.Figure()
            fig.add_trace(go.Bar(name="Value",x=d.index,y=d,marker_color="#2a5298"))
            fig.add_trace(go.Scatter(name="Cumulative",x=cum.index,y=cum,
                line=dict(color="#f7971e",width=2.5),mode="lines+markers"))
            fig.update_layout(title=f"Cumulative: {yc}",height=460,barmode="group")
            st.plotly_chart(fig,use_container_width=True)
 
        elif ctype=="3-D Scatter":
            if len(num_cols)>=3:
                xc=_s("X",num_cols,0,"3x"); yc=_s("Y",num_cols,1,"3y"); zc=_s("Z",num_cols,2,"3z")
                cc=_s("Color",["None"]+cat_cols,key="3c")
                d=df_clean[[xc,yc,zc]].dropna()
                if cc!="None": d[cc]=df_clean[cc]
                fig=px.scatter_3d(d,x=xc,y=yc,z=zc,color=cc if cc!="None" else None)
                fig.update_layout(height=560); st.plotly_chart(fig,use_container_width=True)
            else: st.info("Need ≥ 3 numeric columns.")
 
# ═══════════════════════════════════════════════════════════════════
# TAB 4 – DISTRIBUTIONS
# ═══════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("🥧 Distribution & Composition")
    if not num_cols:
        st.info("No numeric columns.")
    else:
        r1,r2 = st.columns(2)
        with r1:
            st.markdown('<div class="sec-title">🍕 Pie / Donut</div>',unsafe_allow_html=True)
            pie_cats=[c for c in df_clean.columns if c not in num_cols and 1<df_clean[c].nunique()<=30]
            if pie_cats:
                pc=st.selectbox("Category",pie_cats,key="pcat"); pv=st.selectbox("Value",num_cols,key="pval")
                pd_=df_clean[[pc,pv]].copy(); pd_[pv]=pd.to_numeric(pd_[pv],errors="coerce")
                pd_=pd_.dropna().groupby(pc)[pv].sum().reset_index()
                fig=px.pie(pd_,names=pc,values=pv,hole=.38,
                    color_discrete_sequence=px.colors.qualitative.Vivid,title=f"{pv} distribution")
                st.plotly_chart(fig,use_container_width=True)
        with r2:
            st.markdown('<div class="sec-title">📊 Histogram</div>',unsafe_allow_html=True)
            hc=st.selectbox("Column",num_cols,key="hcol"); bins=st.slider("Bins",5,100,25)
            fig=px.histogram(df_clean[hc].dropna(),nbins=bins,color_discrete_sequence=["#17a572"],title=f"Histogram: {hc}")
            fig.update_layout(showlegend=False); st.plotly_chart(fig,use_container_width=True)
 
        st.markdown('<div class="sec-title">🗺️ Treemap</div>',unsafe_allow_html=True)
        tm_cats=[c for c in df_clean.columns if c not in num_cols and 1<df_clean[c].nunique()<=50]
        if tm_cats and num_cols:
            tmc=st.selectbox("Category",tm_cats,key="tmc"); tmv=st.selectbox("Value",num_cols,key="tmv")
            tmd=df_clean[[tmc,tmv]].dropna().groupby(tmc)[tmv].sum().reset_index()
            tmd=tmd[tmd[tmv]>0]
            if len(tmd):
                fig=px.treemap(tmd,path=[tmc],values=tmv,color=tmv,color_continuous_scale="Turbo",title=f"Treemap: {tmv}")
                fig.update_layout(height=460); st.plotly_chart(fig,use_container_width=True)
 
        st.markdown('<div class="sec-title">🎻 Violin</div>',unsafe_allow_html=True)
        vc=st.selectbox("Column",num_cols,key="vc")
        fig=px.violin(df_clean[vc].dropna(),y=vc,box=True,points="outliers",
            color_discrete_sequence=["#c0392b"],title=f"Violin: {vc}")
        st.plotly_chart(fig,use_container_width=True)
 
        sun_cats=[c for c in df_clean.columns if c not in num_cols and 1<df_clean[c].nunique()<=40]
        if len(sun_cats)>=2 and num_cols:
            st.markdown('<div class="sec-title">🌡️ Sunburst</div>',unsafe_allow_html=True)
            s1,s2,s3=st.columns(3)
            sc1=s1.selectbox("Level 1",sun_cats,key="sc1")
            sc2=s2.selectbox("Level 2",[c for c in sun_cats if c!=sc1],key="sc2")
            sv=s3.selectbox("Value",num_cols,key="sv")
            sd=df_clean[[sc1,sc2,sv]].dropna()
            sd[sv]=pd.to_numeric(sd[sv],errors="coerce"); sd=sd.dropna()
            if len(sd):
                fig=px.sunburst(sd,path=[sc1,sc2],values=sv,color=sv,color_continuous_scale="RdYlGn")
                fig.update_layout(height=500); st.plotly_chart(fig,use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════
# TAB 5 – QUERY ENGINE
# ═══════════════════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("🔍 Query Engine  (selected sheet)")
    st.info("Ask in plain English about the **currently selected sheet**. For cross-file search use **💬 AI Smart Query** tab.")
    query=st.text_input("Question",placeholder="e.g. Total subscription / Max capacity / List customers")
 
    def run_query(q,df,nc):
        ql=q.lower(); res=[]
        if any(w in ql for w in ["sum","total","aggregate"]):
            for c in nc:
                if c.lower() in ql or "all" in ql or len(nc)==1: res.append(f"**SUM `{c}`** = {df[c].sum():,.4f}")
        if any(w in ql for w in ["average","mean","avg"]):
            for c in nc:
                if c.lower() in ql or len(nc)==1: res.append(f"**MEAN `{c}`** = {df[c].mean():,.4f}")
        if any(w in ql for w in ["maximum","highest","max"]):
            for c in nc:
                if c.lower() in ql or len(nc)==1: res.append(f"**MAX `{c}`** = {df[c].max():,.4f}")
        if any(w in ql for w in ["minimum","lowest","min"]):
            for c in nc:
                if c.lower() in ql or len(nc)==1: res.append(f"**MIN `{c}`** = {df[c].min():,.4f}")
        if any(w in ql for w in ["count","how many"]): res.append(f"**Rows** = {len(df):,}")
        if "median" in ql:
            for c in nc:
                if c.lower() in ql or len(nc)==1: res.append(f"**MEDIAN `{c}`** = {df[c].median():,.4f}")
        if any(w in ql for w in ["std","deviation"]): [res.append(f"**STD `{c}`** = {df[c].std():,.4f}") for c in nc if c.lower() in ql or len(nc)==1]
        if any(w in ql for w in ["describe","statistics","summary"]): res.append("**Stats:**\n```\n"+df[nc].describe().to_string()+"\n```")
        if any(w in ql for w in ["missing","null","nan"]): ni=df.isna().sum(); ni=ni[ni>0]; res.append("**Missing:**\n"+ni.to_string())
        if "unique" in ql:
            for c in df.columns:
                if c.lower() in ql: u=df[c].dropna().unique(); res.append(f"**Unique `{c}`** ({len(u)}): {', '.join(map(str,u[:25]))}")
        if any(w in ql for w in ["customer","list","show","all"]):
            for c in df.columns:
                if "customer" in c.lower() or "name" in c.lower():
                    names=df[c].dropna().unique(); res.append(f"**`{c}`** ({len(names)} values):\n"+"\n".join(f"  • {n}" for n in names[:30])); break
        if not res: res.append("ℹ️ Try: **sum / average / max / min / count / median / std / unique / missing / describe / customer**")
        return "\n\n".join(res)
 
    if query: st.markdown(run_query(query,df_clean,num_cols))
 
    st.markdown("---")
    st.markdown('<div class="sec-title">🧮 Manual Compute</div>',unsafe_allow_html=True)
    if num_cols:
        mc1,mc2,mc3=st.columns(3)
        op=mc1.selectbox("Op",["Sum","Mean","Max","Min","Count","Median","Std Dev","Variance","% of Total","Cumulative Sum","Range","IQR"])
        sc=mc2.selectbox("Column",num_cols)
        fc=mc3.selectbox("Filter by",["None"]+[c for c in df_clean.columns if c not in num_cols])
        fv=None
        if fc!="None": fv=st.selectbox("Filter value",df_clean[fc].dropna().unique().tolist())
        ds=df_clean.copy()
        if fc!="None" and fv is not None: ds=ds[ds[fc]==fv]
        s=ds[sc].dropna()
        ops={"Sum":s.sum(),"Mean":s.mean(),"Max":s.max(),"Min":s.min(),"Count":s.count(),
             "Median":s.median(),"Std Dev":s.std(),"Variance":s.var(),
             "% of Total":f"{s.sum()/max(df_clean[sc].sum(),1)*100:.2f}%",
             "Cumulative Sum":s.cumsum().iloc[-1] if len(s) else 0,
             "Range":s.max()-s.min(),"IQR":s.quantile(.75)-s.quantile(.25)}
        r=ops.get(op,"N/A")
        if isinstance(r,float): r=f"{r:,.4f}"
        st.success(f"**{op}** of `{sc}`{f' (where {fc}={fv})' if fv else ''} → **{r}**")
        if op=="Cumulative Sum" and len(s):
            st.plotly_chart(px.line(s.cumsum().reset_index(drop=True),title=f"Cumulative: {sc}",markers=True),use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════
# TAB 6 – MULTI-LOCATION
# ═══════════════════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("🌍 Cross-Location Comparison")
 
    @st.cache_data(show_spinner=True)
    def load_all_summ(files, folder):
        summ={}
        for f in files:
            shd=load_file(os.path.join(folder,f))
            for sh,raw in shd.items():
                dfc=to_numeric(smart_header(raw)); nc=dfc.select_dtypes(include="number").columns.tolist()
                if nc: summ[f"{loc_map[f]} | {sh}"]={"df":dfc,"num_cols":nc,"file":f,"sheet":sh}
        return summ
 
    all_summ=load_all_summ(tuple(excel_files),data_dir)
 
    if all_summ:
        comp_col=st.selectbox("Compare by column",sorted({c for v in all_summ.values() for c in v["num_cols"]}))
        rows=[]
        for lbl,info in all_summ.items():
            if comp_col in info["num_cols"]:
                s=info["df"][comp_col].dropna()
                rows.append({"Location|Sheet":lbl,"Sum":s.sum(),"Mean":s.mean(),"Max":s.max(),"Min":s.min(),"Count":s.count()})
        if rows:
            cmp=pd.DataFrame(rows).set_index("Location|Sheet")
            st.dataframe(cmp.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),use_container_width=True)
            fig=px.bar(cmp.reset_index(),x="Location|Sheet",y="Sum",color="Sum",
                color_continuous_scale="Viridis",title=f"Sum of '{comp_col}' across locations")
            fig.update_layout(xaxis_tickangle=-30,height=450); st.plotly_chart(fig,use_container_width=True)
            fig2=px.scatter(cmp.reset_index(),x="Mean",y="Max",size="Sum",text="Location|Sheet",color="Count",
                color_continuous_scale="Turbo",title="Bubble: Mean vs Max")
            fig2.update_traces(textposition="top center"); fig2.update_layout(height=500)
            st.plotly_chart(fig2,use_container_width=True)
            st.markdown('<div class="sec-title">🕸️ Radar / Spider</div>',unsafe_allow_html=True)
            rm=st.radio("Metric",["Sum","Mean","Max"],horizontal=True)
            norm=cmp[rm]/cmp[rm].max(); th=norm.index.tolist(); rv=norm.values.tolist()
            fig3=go.Figure(go.Scatterpolar(r=rv+[rv[0]],theta=th+[th[0]],fill="toself",
                fillcolor="rgba(42,82,152,.25)",line=dict(color="#2a5298",width=2)))
            fig3.update_layout(polar=dict(radialaxis=dict(visible=True,range=[0,1])),height=520,title=f"Radar: {rm}")
            st.plotly_chart(fig3,use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════
# TAB 7 – AI AGENT  (automated insights)
# ═══════════════════════════════════════════════════════════════════
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights")
    st.info("Click Run to scan every sheet and surface KPIs, anomalies and correlations.")
    if st.button("🚀 Run AI Agent Analysis",type="primary"):
        with st.spinner("Analysing…"):
            for lbl,info in list(all_summ.items())[:10]:
                dfa=info["df"]; nc=info["num_cols"]
                if not nc: continue
                with st.expander(f"📍 {lbl}",expanded=False):
                    ca,cb=st.columns(2)
                    with ca:
                        st.markdown("**📊 KPIs**")
                        for col in nc[:5]:
                            s=dfa[col].dropna()
                            if len(s): st.metric(col[:28],f"{s.sum():,.1f}",f"avg {s.mean():,.1f}")
                    with cb:
                        st.markdown("**⚠️ Anomalies (Z-score)**")
                        for col in nc[:4]:
                            s=dfa[col].dropna()
                            if len(s)>3:
                                z=(s-s.mean())/s.std(); o=z[z.abs()>2.5]
                                (st.warning if len(o) else st.success)(f"`{col}`: {len(o)} outlier(s)" if len(o) else f"`{col}`: Clean ✓")
                    if len(nc)>=2:
                        fig=px.bar(dfa[nc[:3]].dropna().reset_index(),x="index",y=nc[0],
                            color_discrete_sequence=["#2a5298"],title=nc[0])
                        fig.update_layout(height=260,showlegend=False,margin=dict(t=30,b=0))
                        st.plotly_chart(fig,use_container_width=True)
                    if len(nc)>=2:
                        cs=(dfa[nc].corr().unstack().drop_duplicates())
                        cs=cs[cs.index.get_level_values(0)!=cs.index.get_level_values(1)].abs().sort_values(ascending=False)
                        if len(cs):
                            p=cs.index[0]; st.info(f"🔗 Top corr: **{p[0]}** ↔ **{p[1]}** ({cs.iloc[0]:.3f})")
 
    st.markdown("---")
    st.markdown('<div class="sec-title">📁 All Files Summary</div>',unsafe_allow_html=True)
    fsm=[]
    for f in excel_files:
        shd=all_sheets if f==selected_file else load_file(os.path.join(data_dir,f))
        fsm.append({"File":loc_map[f],"Sheets":len(shd),
            "Total Rows":sum(len(s) for s in shd.values()),
            "Total Columns":sum(len(s.columns) for s in shd.values())})
    fs_df=pd.DataFrame(fsm)
    st.dataframe(fs_df,use_container_width=True)
    fig=px.bar(fs_df,x="File",y="Total Rows",color="Sheets",color_continuous_scale="Blues",title="Rows per location")
    fig.update_layout(xaxis_tickangle=-30,height=380); st.plotly_chart(fig,use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════
# TAB 8 – 💬 AI SMART QUERY
#
# Searches EVERY cell at ANY row, ANY column, ANYWHERE in the sheet,
# including middle-of-sheet data, merged areas, summary rows, etc.
# Handles unstructured/odd Excel layouts.
# Zero formulas. Zero hard-coded row/column conditions.
# ═══════════════════════════════════════════════════════════════════
with tabs[8]:
 
    # ─── Build corpus ───────────────────────────────────────────────
    with st.spinner("🔍 Indexing every cell in every sheet of every file (including .xls)…"):
        corpus, row_records, meta = build_corpus(tuple(excel_files), data_dir)
 
    # ─── Header ─────────────────────────────────────────────────────
    st.markdown("## 💬 AI Smart Query")
    st.markdown(
        "Type **any question** in plain English. "
        "The engine reads **every row · every column · every sheet · every file** "
        "— including rows in the middle of unstructured sheets, summary rows, "
        "merged-cell areas and anywhere else data exists. "
        "**No formulas. No pre-conditions. No row or column restrictions.**"
    )
 
    qi1,qi2,qi3,qi4,qi5 = st.columns(5)
    qi1.markdown(f'<div class="kcard kcard-blue"><h2>{meta["total_cells"]:,}</h2><p>Cells Indexed</p></div>',unsafe_allow_html=True)
    qi2.markdown(f'<div class="kcard kcard-green"><h2>{meta["total_files"]}</h2><p>Files</p></div>',unsafe_allow_html=True)
    qi3.markdown(f'<div class="kcard kcard-purple"><h2>{meta["total_sheets"]}</h2><p>Sheets</p></div>',unsafe_allow_html=True)
    qi4.markdown(f'<div class="kcard kcard-orange"><h2>{meta["total_rows"]:,}</h2><p>Data Rows</p></div>',unsafe_allow_html=True)
    qi5.markdown(f'<div class="kcard kcard-teal"><h2>{len(meta["locations"])}</h2><p>Locations</p></div>',unsafe_allow_html=True)
 
    st.markdown("")
 
    # ─── Optional scope narrowing ────────────────────────────────────
    with st.expander("🔧 Optional: Narrow scope before asking", expanded=False):
        scope_locs   = st.multiselect("Limit to locations (blank = ALL)", meta["locations"], default=[])
        scope_sheets = st.multiselect("Limit to sheet names (blank = ALL)",
            sorted({x["sheet"] for x in corpus}), default=[])
 
    # ─── Float helper ────────────────────────────────────────────────
    def _try_float(v):
        try: float(v); return True
        except: return False
 
    # ─── AI engine ──────────────────────────────────────────────────
    def ai_smart_query(question: str) -> dict:
        """
        Pure Python/Pandas engine.
        Scans the pre-built corpus that contains EVERY cell from EVERY row
        and EVERY column of EVERY sheet, anywhere in the spreadsheet.
        No Excel formulas. No conditions. Just the question → direct answer.
        """
        q     = question.strip()
        ql    = q.lower()
        sig   = [w for w in re.findall(r"[a-z0-9]{3,}", ql) if w not in _SW]
 
        # Intent detection
        f_sum   = any(x in ql for x in ["total","sum","aggregate","add up"])
        f_avg   = any(x in ql for x in ["average","mean","avg"])
        f_max   = any(x in ql for x in ["maximum","highest","largest","biggest","max","top"])
        f_min   = any(x in ql for x in ["minimum","lowest","smallest","least","min","bottom"])
        f_cnt   = any(x in ql for x in ["count","how many","number of"])
        f_pct   = any(x in ql for x in ["percent","percentage","%","ratio","share","proportion"])
        f_stat  = any(x in ql for x in ["statistics","stats","describe","summary","overview","profile"])
        f_uniq  = any(x in ql for x in ["unique","distinct","different","values of"])
        f_miss  = any(x in ql for x in ["missing","null","blank","empty","nan"])
        f_sheet = any(x in ql for x in ["sheet","sheets","tab","tabs"])
        f_file  = any(x in ql for x in ["file","files","location","locations"])
        f_col   = any(x in ql for x in ["column","columns","field","fields","header","headers"])
        f_row   = any(x in ql for x in ["row","rows","record","records","entry","entries"])
        f_topn  = re.search(r"\btop\s*(\d+)\b", ql)
        f_botn  = re.search(r"\bbottom\s*(\d+)\b", ql)
        f_num   = f_sum or f_avg or f_max or f_min or f_cnt or f_stat or f_pct
 
        out = {"answer":"","table":None,"chart_df":None,"chart_cfg":None,"cell_hits":[],"sub_tables":[]}
 
        if not corpus:
            out["answer"]="⚠️ Corpus empty – no files loaded."; return out
 
        # Apply scope filters
        def scoped(cells):
            if scope_locs:   cells=[c for c in cells if c["location"] in scope_locs]
            if scope_sheets: cells=[c for c in cells if c["sheet"] in scope_sheets]
            return cells
 
        wc = scoped(corpus)   # working corpus
 
        # Location keyword auto-detection
        loc_filt = None
        for loc in meta["locations"]:
            for part in re.findall(r"[a-z0-9]+", loc.lower()):
                if len(part)>=4 and part in ql:
                    loc_filt=loc; break
            if loc_filt: break
 
        def loc_ok(cell_or_str):
            if not loc_filt: return True
            s=cell_or_str if isinstance(cell_or_str,str) else cell_or_str["location"]
            return loc_filt.lower() in s.lower()
 
        wc=[c for c in wc if loc_ok(c)]   # apply location filter
 
        # Best column keyword from question
        _OPS={"total","sum","avg","mean","max","min","count","list","find","show","all","average",
              "maximum","minimum","highest","lowest","top","bottom","describe","statistics","stats",
              "summary","unique","distinct","sheet","column","row","missing","null","percent",
              "percentage","ratio","share","proportion","number","across","compare"}
        def best_col_kw(slist):
            cands=[w for w in slist if w not in _OPS and len(w)>=3]
            best,bn=None,0
            for w in cands:
                n=sum(1 for c in wc if w in c["col_header"].lower())
                if n>bn: bn,best=n,w
            return best
 
        # Get numeric values from any cell whose col_header contains kw
        # (ANY row, ANY column, ANY position in sheet)
        def num_for_kw(kw):
            res=[]
            for cell in wc:
                if cell["is_header"]: continue
                if kw.lower() in cell["col_header"].lower():
                    try: res.append((float(cell["value"]),cell))
                    except ValueError: pass
            return res
 
        # Build full-row DataFrame from row keys
        def rows_df(keys):
            recs=[]
            for key in keys:
                rec=row_records.get(key,{})
                if rec:
                    rd={"📍 Location":key[1],"📋 Sheet":key[2],"Row #":key[3]+1}
                    rd.update(rec); recs.append(rd)
            return pd.DataFrame(recs) if recs else pd.DataFrame()
 
        # ─── INTENT A: Sheet listing ────────────────────────────────
        if f_sheet and not f_num:
            seen=set(); srows=[]
            for cell in wc:
                k=(cell["location"],cell["sheet"])
                if k not in seen:
                    seen.add(k)
                    dr=sum(1 for rk in row_records if rk[1]==cell["location"] and rk[2]==cell["sheet"])
                    srows.append({"Location":cell["location"],"Sheet":cell["sheet"],"File":cell["file"],"Data Rows":dr})
            tbl=pd.DataFrame(srows)
            out["answer"]=f"Found **{len(tbl)}** sheet(s){' in '+loc_filt if loc_filt else ''}."; out["table"]=tbl; return out
 
        # ─── INTENT B: File/location listing ───────────────────────
        if f_file and not f_num:
            seen=set(); frows=[]
            for cell in wc:
                if cell["location"] not in seen:
                    seen.add(cell["location"])
                    cl=[c for c in wc if c["location"]==cell["location"]]
                    frows.append({"Location":cell["location"],"Sheets":len({c["sheet"] for c in cl}),"Cells":len(cl)})
            tbl=pd.DataFrame(frows)
            out["answer"]=f"Found **{len(tbl)}** location(s) in scope."; out["table"]=tbl; return out
 
        # ─── INTENT C: Missing values ───────────────────────────────
        if f_miss:
            mrows=[]
            for fname in excel_files:
                loc=location_from_name(fname)
                if not loc_ok(loc): continue
                shd=load_file(os.path.join(data_dir,fname))
                for sh,raw in shd.items():
                    if scope_sheets and sh not in scope_sheets: continue
                    dfc=to_numeric(smart_header(raw))
                    for col in dfc.columns:
                        mc=int(dfc[col].isna().sum())
                        if mc>0: mrows.append({"Location":loc,"Sheet":sh,"Column":col,"Missing":mc,"Missing%":f"{mc/max(len(dfc),1)*100:.1f}%"})
            if mrows:
                tbl=pd.DataFrame(mrows).sort_values("Missing",ascending=False)
                out["answer"]=f"Found **{len(tbl)}** column(s) with missing values."; out["table"]=tbl
            else:
                out["answer"]="✅ No missing values found in selected scope."
            return out
 
        # ─── INTENT D: Column listing ───────────────────────────────
        if f_col and not f_num:
            kw=best_col_kw(sig); seen=set(); crows=[]
            for cell in wc:
                if not cell["is_header"]: continue
                ch=cell["col_header"].strip()
                if ch in ("","nan"): continue
                if kw and kw.lower() not in ch.lower(): continue
                k=(cell["location"],cell["sheet"],ch)
                if k not in seen: seen.add(k); crows.append({"Location":cell["location"],"Sheet":cell["sheet"],"Column":ch})
            tbl=pd.DataFrame(crows) if crows else pd.DataFrame()
            out["answer"]=f"Found **{len(tbl)}** column(s) matching your query."; out["table"]=tbl; return out
 
        # ─── INTENT E: Row count ────────────────────────────────────
        if f_row and f_cnt and not sig:
            crows=[]; total=0
            for fname in excel_files:
                loc=location_from_name(fname)
                if not loc_ok(loc): continue
                shd=load_file(os.path.join(data_dir,fname))
                for sh,raw in shd.items():
                    if scope_sheets and sh not in scope_sheets: continue
                    dfc=smart_header(raw); crows.append({"Location":loc,"Sheet":sh,"Data Rows":len(dfc)}); total+=len(dfc)
            tbl=pd.DataFrame(crows)
            out["answer"]=f"**{total:,}** total data rows across **{len(tbl)}** sheet(s)."; out["table"]=tbl
            out["chart_df"]=tbl.groupby("Location")["Data Rows"].sum().reset_index()
            out["chart_cfg"]={"x":"Location","y":"Data Rows","title":"Data Rows per Location"}; return out
 
        # ─── INTENT F: Numeric aggregation ─────────────────────────
        # Finds matching cells at ANY row, ANY column position in the sheet
        if f_num and sig:
            kw=best_col_kw(sig)
            if kw:
                pairs=num_for_kw(kw)
                if pairs:
                    vals=[v for v,_ in pairs]; s_all=pd.Series(vals); parts=[]
                    if f_sum  or f_stat: parts.append(f"**Total (Sum):**    {s_all.sum():,.4f}")
                    if f_avg  or f_stat: parts.append(f"**Average (Mean):** {s_all.mean():,.4f}")
                    if f_max  or f_stat: parts.append(f"**Maximum:**        {s_all.max():,.4f}")
                    if f_min  or f_stat: parts.append(f"**Minimum:**        {s_all.min():,.4f}")
                    if f_cnt  or f_stat: parts.append(f"**Count:**          {s_all.count():,}")
                    if f_stat:
                        parts.append(f"**Median:**   {s_all.median():,.4f}\n**Std Dev:**  {s_all.std():,.4f}\n**Variance:** {s_all.var():,.4f}")
                    if f_pct:
                        grand=sum(float(c["value"]) for c in wc if not c["is_header"] and _try_float(c["value"]))
                        pct=s_all.sum()/grand*100 if grand else 0; parts.append(f"**% of all numeric:** {pct:.2f}%")
                    grp=defaultdict(list)
                    for v,cell in pairs: grp[f"{cell['location']} | {cell['sheet']}"].append(v)
                    breakdown=[]
                    for lbl2,vs in grp.items():
                        sv=pd.Series(vs)
                        breakdown.append({"Location | Sheet":lbl2,"Count":sv.count(),"Sum":round(sv.sum(),4),
                            "Mean":round(sv.mean(),4),"Max":round(sv.max(),4),"Min":round(sv.min(),4)})
                    tbl=pd.DataFrame(breakdown).sort_values("Sum",ascending=False)
                    out["answer"]=(f"Results for columns matching **'{kw}'**  "
                        f"({len(vals):,} values from {len(breakdown)} source(s)"
                        f"{', in '+loc_filt if loc_filt else ''}):\n\n"+"\n".join(parts))
                    out["table"]=tbl; out["chart_df"]=tbl
                    out["chart_cfg"]={"x":"Location | Sheet","y":"Sum","title":f"Sum of '{kw}' by Location/Sheet"}
                    if f_topn:
                        n=int(f_topn.group(1)); top=sorted(pairs,key=lambda x:x[0],reverse=True)[:n]
                        out["sub_tables"].append({"label":f"🏆 Top {n} values for '{kw}'",
                            "df":pd.DataFrame([{"📍 Location":c["location"],"📋 Sheet":c["sheet"],
                                "Row #":c["row"]+1,"Col #":c["col"]+1,"Column":c["col_header"],"Value":v} for v,c in top])})
                    if f_botn:
                        n=int(f_botn.group(1)); bot=sorted(pairs,key=lambda x:x[0])[:n]
                        out["sub_tables"].append({"label":f"🔻 Bottom {n} values for '{kw}'",
                            "df":pd.DataFrame([{"📍 Location":c["location"],"📋 Sheet":c["sheet"],
                                "Row #":c["row"]+1,"Col #":c["col"]+1,"Column":c["col_header"],"Value":v} for v,c in bot])})
                    return out
 
        # ─── INTENT G: Unique values ────────────────────────────────
        if f_uniq and sig:
            kw=best_col_kw(sig)
            if kw:
                uvals=set(); srows=[]
                for cell in wc:
                    if cell["is_header"]: continue
                    if kw.lower() in cell["col_header"].lower():
                        uvals.add(cell["value"])
                        srows.append({"Location":cell["location"],"Sheet":cell["sheet"],"Column":cell["col_header"],"Value":cell["value"]})
                tbl=(pd.DataFrame(srows).drop_duplicates(subset=["Location","Sheet","Value"]) if srows else pd.DataFrame())
                out["answer"]=(f"Found **{len(uvals)}** unique value(s) in columns matching **'{kw}'**"
                    f"{' in '+loc_filt if loc_filt else ''}.")
                out["table"]=tbl; return out
 
        # ─── INTENT H: Free-text entity/keyword search ──────────────
        # Searches VALUES anywhere – any row, any column, any position
        # in unstructured sheets, middle of data, summary rows, etc.
        if sig:
            quoted=re.findall(r'"([^"]+)"',q)
            if quoted:
                terms=[quoted[0].lower()]
            else:
                terms=[w for w in sig if any(w in cell["value"].lower() for cell in wc)]
                if not terms: terms=sig
 
            hit_cells=[
                cell for cell in wc
                if not cell["is_header"]
                and any(t in cell["value"].lower() for t in terms)
            ]
            hit_keys={(c["file"],c["location"],c["sheet"],c["row"]) for c in hit_cells}
            full_df=rows_df(hit_keys)
 
            loc_freq=defaultdict(int); sh_freq=defaultdict(int)
            for c in hit_cells:
                loc_freq[c["location"]]+=1; sh_freq[f"{c['location']} | {c['sheet']}"]+=1
 
            lf_df=pd.DataFrame(list(loc_freq.items()),columns=["Location","Hits"]).sort_values("Hits",ascending=False)
 
            cell_list=[{"📍 Location":c["location"],"📋 Sheet":c["sheet"],"Row #":c["row"]+1,
                "Col #":c["col"]+1,"Column Header":c["col_header"],"Value":c["value"]} for c in hit_cells[:40]]
 
            out["answer"]=(f"Found **{len(hit_cells):,}** cell(s) matching **'{', '.join(terms[:4])}'**  "
                f"across **{len(loc_freq)}** location(s) and **{len(sh_freq)}** sheet(s)"
                f"{' in '+loc_filt if loc_filt else ''}.\n\n"
                f"**{len(hit_keys):,}** unique data row(s) contain this data.")
            out["table"]=full_df if not full_df.empty else None
            out["cell_hits"]=cell_list
 
            if len(loc_freq)>1:
                out["chart_df"]=lf_df
                out["chart_cfg"]={"x":"Location","y":"Hits","title":f"Hits per location — '{', '.join(terms[:3])}'"}
 
            if not full_df.empty:
                for col in full_df.columns:
                    if "customer" in col.lower() or "name" in col.lower():
                        cust_df=full_df[["📍 Location","📋 Sheet",col]].drop_duplicates()
                        out["sub_tables"].append({"label":f"👤 Names / Customers ({len(cust_df)} rows)","df":cust_df}); break
 
            num_hits=[]
            for c in hit_cells:
                try: num_hits.append((float(c["value"]),c))
                except: pass
            if num_hits and f_topn:
                n=int(f_topn.group(1)); top=sorted(num_hits,key=lambda x:x[0],reverse=True)[:n]
                out["sub_tables"].append({"label":f"🏆 Top {n} numeric values in results",
                    "df":pd.DataFrame([{"Location":c["location"],"Sheet":c["sheet"],"Row #":c["row"]+1,"Col":c["col_header"],"Value":v} for v,c in top])})
            return out
 
        # ─── FALLBACK ────────────────────────────────────────────────
        out["answer"]=(
            "❓ Could not find a match for your query.\n\n"
            "**Try these patterns:**\n"
            "• Entity:      *Find CISCO*  |  *Axis Bank*  |  *AT&T*  |  *Zscaler*\n"
            "• Numeric:     *Total subscription*  |  *Max capacity Airoli*  |  *Average power usage*  |  *Top 10 subscription*\n"
            "• List:        *List all customers*  |  *Customers in Noida*  |  *All metered customers*\n"
            "• Unique:      *Unique billing models*  |  *Distinct floors*\n"
            "• Meta:        *Show all sheets*  |  *List columns*  |  *Count rows per sheet*\n"
            "• Missing:     *Show missing values*"
        )
        return out
 
    # ─── Render one AI answer ────────────────────────────────────────
    def render_answer(res):
        st.markdown(f'<div class="ans-box">{res["answer"]}</div>',unsafe_allow_html=True)
        if res.get("table") is not None and not res["table"].empty:
            tbl=res["table"].reset_index(drop=True)
            st.dataframe(tbl,use_container_width=True,height=min(560,48+len(tbl)*36))
            st.download_button("⬇️ Download CSV",tbl.to_csv(index=False).encode(),"ai_result.csv","text/csv",key=f"dl_{id(res)}")
        if res.get("chart_cfg") and res.get("chart_df") is not None:
            cfg=res["chart_cfg"]; cdf=res["chart_df"]
            if cfg["x"] in cdf.columns and cfg["y"] in cdf.columns:
                fig=px.bar(cdf.sort_values(cfg["y"],ascending=False).head(30),x=cfg["x"],y=cfg["y"],
                    color=cfg["y"],color_continuous_scale="Viridis",title=cfg["title"],height=420)
                fig.update_layout(xaxis_tickangle=-30); st.plotly_chart(fig,use_container_width=True)
        for st_item in res.get("sub_tables",[]):
            with st.expander(st_item["label"],expanded=True): st.dataframe(st_item["df"],use_container_width=True)
        if res.get("cell_hits"):
            with st.expander(f"🔬 Cell-level matches — first {len(res['cell_hits'])} shown",expanded=False):
                for ch in res["cell_hits"]:
                    st.markdown(
                        f'<div class="cell-chip">'
                        f'📍 <b>{ch["📍 Location"]}</b> → Sheet: <b>{ch["📋 Sheet"]}</b>  |  '
                        f'Row <b>{ch["Row #"]}</b> Col <b>{ch["Col #"]}</b>  |  '
                        f'Header: <i>{ch["Column Header"]}</i>  |  Value: <b>{ch["Value"]}</b>'
                        f'</div>',unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>',unsafe_allow_html=True)
 
    # ─── Chat history ────────────────────────────────────────────────
    if "aisq_hist" not in st.session_state:
        st.session_state.aisq_hist=[]
 
    for turn in st.session_state.aisq_hist:
        st.markdown(f'<div class="q-user">🧑 {turn["q"]}</div>',unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>',unsafe_allow_html=True)
        render_answer(turn["res"]); st.markdown("---")
 
    # ─── Input bar ───────────────────────────────────────────────────
    st.markdown("---")
    ic,bc,cc=st.columns([8,1,1])
    with ic:
        user_q=st.text_input("Ask:",
            placeholder="Find CISCO | List all customers in Noida | Total subscription Airoli | Max capacity | Average power | Top 10 subscription",
            label_visibility="collapsed",key="aisq_input")
    with bc:
        ask_btn=st.button("🔍 Ask",use_container_width=True,type="primary")
    with cc:
        if st.button("🗑️ Clear",use_container_width=True):
            st.session_state.aisq_hist=[]; st.rerun()
 
    # ─── Quick-click example chips ───────────────────────────────────
    st.markdown("**💡 Click any example to ask instantly:**")
    examples=[
        ["List all customers","List all customers in Noida","List all customers in Bangalore",
         "List all customers in Chennai","List all customers in Vashi","List all customers in Kolkata"],
        ["Total subscription across all locations","Total capacity in Airoli","Average power usage",
         "Maximum rack subscription","Minimum capacity in Kolkata","Top 10 subscription values"],
        ["Find CISCO","Find Axis Bank","Find MOTMOT","Find AT&T","Find Zscaler","Find Edge Network"],
        ["Count rows per sheet","Show missing values","Unique billing models",
         "Unique subscription mode","Show all sheets","Statistics of subscription"],
    ]
    for row in examples:
        cols=st.columns(len(row))
        for j,ex in enumerate(row):
            if cols[j].button(ex,key=f"chip_{ex}",use_container_width=True):
                user_q=ex; ask_btn=True
 
    # ─── Execute ─────────────────────────────────────────────────────
    if ask_btn and user_q.strip():
        with st.spinner(f"🤖 Scanning every cell for: **{user_q}** …"):
            answer=ai_smart_query(user_q)
        st.session_state.aisq_hist.append({"q":user_q,"res":answer})
        st.rerun()
 
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Sify DC · Customer & Capacity Tracker · Streamlit + Plotly · Indexes every cell of every sheet")
