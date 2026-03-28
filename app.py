import os, re, warnings, tempfile, subprocess
from collections import defaultdict
from pathlib import Path
 
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
 
warnings.filterwarnings("ignore")
 
st.set_page_config(page_title="Sify DC – Capacity Tracker", page_icon="🏢",
                   layout="wide", initial_sidebar_state="expanded")
 
st.markdown("""
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0e1a,#1a2035,#0d1b2a)!important;}
[data-testid="stSidebar"] *{color:#c9d8f0!important;}
.kcard{border-radius:14px;padding:16px 20px;color:#fff;margin-bottom:10px;box-shadow:0 4px 18px rgba(0,0,0,.35);transition:transform .2s;}
.kcard:hover{transform:translateY(-2px);}
.kcard h2{font-size:1.8rem;margin:0;font-weight:800;}
.kcard p{margin:3px 0 0;font-size:.82rem;opacity:.82;}
.kcard-blue{background:linear-gradient(135deg,#1e3c72,#2a5298);}
.kcard-green{background:linear-gradient(135deg,#0b6e4f,#17a572);}
.kcard-red{background:linear-gradient(135deg,#7b1a1a,#c0392b);}
.kcard-orange{background:linear-gradient(135deg,#7d4e00,#e67e22);}
.kcard-teal{background:linear-gradient(135deg,#0f3460,#16213e);}
.kcard-purple{background:linear-gradient(135deg,#4a0072,#7b1fa2);}
.sec-title{font-size:1.15rem;font-weight:700;color:#1e3c72;border-left:5px solid #2a5298;padding-left:10px;margin:16px 0 10px;}
.q-user{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff;border-radius:18px 18px 4px 18px;padding:10px 16px;margin:10px 0 4px auto;max-width:76%;width:fit-content;box-shadow:0 3px 12px rgba(30,60,114,.45);float:right;clear:both;}
.ans-box{background:linear-gradient(135deg,#0f2744,#1a4a6b);color:#d0ecff;border-radius:12px;padding:14px 18px;margin:8px 0;font-size:.97rem;box-shadow:0 3px 14px rgba(0,0,0,.35);white-space:pre-wrap;line-height:1.6;}
.cell-chip{background:#1a2f1a;border-left:4px solid #27ae60;border-radius:6px;padding:6px 12px;margin:3px 0;font-family:monospace;font-size:.8rem;color:#b8ffb8;}
.clearfix{clear:both;}
</style>
""", unsafe_allow_html=True)
 
_SW = {"the","and","for","are","all","any","how","what","show","give","tell","from","this","that","with","get","find","list","much","many","each","every","data","value","values","number","numbers","in","of","a","an","is","at","by","to","do","me","my","about","details","info","please","can","you","per","across","which","where","who","when","does","did","have","has","their","its","our","your","there","these","those","been","will","would","could","should","shall","let","some","just","also","even","only","into","over","under","both","such","than","then","but","not","nor","yet","so","either","neither","versus","vs"}
 
def _app_dir():
    try: return Path(__file__).resolve().parent
    except NameError: return Path(os.getcwd())
EXCEL_FOLDER = _app_dir() / "excel_files"
 
def find_excel_files(folder):
    p = Path(folder)
    if not p.is_dir(): return []
    return sorted(f.name for f in p.iterdir() if f.suffix.lower() in (".xlsx",".xls") and not f.name.startswith("~"))
 
def location_from_name(fname):
    n = os.path.basename(fname)
    n = re.sub(r"\.(xlsx?|xls)$","",n,flags=re.I)
    n = re.sub(r"[Cc]ustomer.?[Aa]nd.?[Cc]apacity.?[Tt]racker.?","",n)
    n = re.sub(r"[_\s]?\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{2,4}.*$","",n,flags=re.I)
    n = re.sub(r"__\d+_*$","",n)
    n = re.sub(r"[_]+"," ",n).strip()
    return n if n else fname
 
@st.cache_data(show_spinner=False)
def save_uploads(file_bytes_tuple):
    tmp = tempfile.mkdtemp()
    for name, data in file_bytes_tuple:
        with open(os.path.join(tmp, name), "wb") as fh: fh.write(data)
    return tmp
 
@st.cache_data(show_spinner=False)
def ensure_readable(original_path):
    if not original_path.lower().endswith(".xls"): return original_path
    try:
        with open(original_path,"rb") as fh:
            if fh.read(4) != b"\xd0\xcf\x11\xe0": return original_path
    except: return original_path
    out_dir = tempfile.mkdtemp()
    wrapper = "/mnt/skills/public/xlsx/scripts/office/soffice.py"
    try:
        if os.path.exists(wrapper):
            subprocess.run(["python3",wrapper,"--convert-to","xlsx","--outdir",out_dir,original_path], capture_output=True, timeout=60)
        else:
            subprocess.run(["libreoffice","--headless","--convert-to","xlsx","--outdir",out_dir,original_path], capture_output=True, timeout=120)
        base = os.path.splitext(os.path.basename(original_path))[0]
        conv = os.path.join(out_dir, base+".xlsx")
        if os.path.exists(conv): return conv
    except: pass
    return original_path
 
@st.cache_data(show_spinner=False)
def _read_sheet(path, sheet_name):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    mr = ws.max_row or 0; mc = ws.max_column or 0
    if mr == 0: wb.close(); return pd.DataFrame()
    real_mc = 0
    samples = sorted(set(list(range(1, min(31, mr+1))) + list(range(max(1, mr-9), mr+1))))
    for r in samples:
        for cell in ws[r]:
            if cell.value is not None: real_mc = max(real_mc, cell.column)
    if real_mc == 0: wb.close(); return pd.DataFrame()
    cap = min(real_mc + 2, mc)
    rows = []
    for row in ws.iter_rows(min_row=1, max_row=mr, max_col=cap, values_only=True):
        rows.append(list(row))
    wb.close()
    if not rows: return pd.DataFrame()
    df = pd.DataFrame(rows, dtype=str)
    df = df.replace({"None": np.nan, "none": np.nan})
    return df
 
@st.cache_data(show_spinner=False)
def load_file(original_path):
    path = ensure_readable(original_path)
    sheets = {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True); names = wb.sheetnames; wb.close()
        for sh in names:
            try:
                df = _read_sheet(path, sh)
                if not df.empty: sheets[sh] = df
            except: pass
    except Exception as e:
        st.sidebar.warning(f"⚠️ {os.path.basename(original_path)}: {e}")
    return sheets
 
def best_header_row(df):
    best_row, best_score = 0, -1
    for i in range(min(8, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        filled = (row.str.len()>0)&(~row.isin(["nan","None",""]))
        label = filled & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        score = label.sum()*2 + filled.sum()
        if score > best_score: best_score, best_row = score, i
    return best_row
 
def smart_header(df):
    hr = best_header_row(df)
    hdr = df.iloc[hr].fillna("").astype(str).str.strip()
    seen = {}; cols = []
    for col in hdr:
        col = col if col and col not in ("nan","None") else f"Col_{len(cols)}"
        if col in seen: seen[col]+=1; cols.append(f"{col}_{seen[col]}")
        else: seen[col]=0; cols.append(col)
    data = df.iloc[hr+1:].copy(); data.columns = cols
    return data.dropna(how="all").reset_index(drop=True)
 
def to_numeric(df):
    out = df.copy()
    for col in out.columns: out[col] = pd.to_numeric(out[col], errors="ignore")
    return out
 
def _detect_all_header_rows(df):
    hr_set = set()
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.strip()
        fm = (row.str.len()>0) & (~row.isin(["nan","None",""]))
        fv = row[fm]; nf = fv.shape[0]
        if nf < 2: continue
        lm = fm & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        nl = lm.sum(); nu = fv.nunique()
        lr = nl/max(nf,1); ur = nu/max(nf,1)
        vc = fv.value_counts(); nr = (vc>1).sum()
        if lr>=0.80 and ur>=0.75 and nr<=max(2, nf*0.15) and nu>=3: hr_set.add(i)
        elif nf<=10 and nf>=2 and lr>=0.90 and ur>=0.80 and nr<=1: hr_set.add(i)
    return hr_set
 
def _build_cell_col_map(df):
    hr_set = _detect_all_header_rows(df)
    hr_maps = {}
    for hr in hr_set:
        m = {}
        for c in range(df.shape[1]):
            v = str(df.iat[hr,c]).strip()
            if v and v not in ("nan","None"): m[c] = v
        hr_maps[hr] = m
    sorted_hrs = sorted(hr_set)
    cell_map = {}
    for r in range(df.shape[0]):
        prev = [h for h in sorted_hrs if h < r]
        for c in range(df.shape[1]):
            name = f"Col_{c}"
            for h in reversed(prev):
                if c in hr_maps[h]: name = hr_maps[h][c]; break
            cell_map[(r,c)] = name
    return cell_map, hr_set
 
@st.cache_data(show_spinner=False)
def index_single_sheet(file_path, sheet_name):
    sheets = load_file(file_path)
    if sheet_name not in sheets:
        return [], {}, {"total_cells": 0, "total_rows": 0, "total_data": 0, "total_headers": 0}
    df = sheets[sheet_name]
    cell_map, hr_set = _build_cell_col_map(df)
    cells = []; row_recs = {}
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            raw = df.iat[r, c]
            if pd.isna(raw): continue
            v = str(raw).strip()
            if not v or v in ("nan", "None", "none", ""): continue
            ch = cell_map.get((r, c), f"Col_{c}")
            is_hdr = (r in hr_set)
            cells.append({"row": r, "col": c, "col_header": ch, "value": v, "is_header": is_hdr})
            if not is_hdr:
                if r not in row_recs: row_recs[r] = {}
                row_recs[r][ch] = v
    meta = {"total_cells": len(cells), "total_rows": len(row_recs),
            "total_data": sum(1 for x in cells if not x["is_header"]),
            "total_headers": sum(1 for x in cells if x["is_header"])}
    return cells, row_recs, meta
 
@st.cache_data(show_spinner=False)
def build_corpus(file_list, folder):
    corpus = []; row_records = defaultdict(dict)
    for fname in file_list:
        full = os.path.join(folder, fname)
        if not os.path.isfile(full): continue
        loc = location_from_name(fname)
        sheets = load_file(full)
        for sh, df in sheets.items():
            cell_map, hr_set = _build_cell_col_map(df)
            for r in range(df.shape[0]):
                for c in range(df.shape[1]):
                    raw = df.iat[r, c]
                    if pd.isna(raw): continue
                    v = str(raw).strip()
                    if not v or v in ("nan", "None", "none", ""): continue
                    ch = cell_map.get((r, c), f"Col_{c}")
                    is_hdr = (r in hr_set)
                    key = (fname, loc, sh, r)
                    corpus.append({"file": fname, "location": loc, "sheet": sh, "row": r, "col": c, "col_header": ch, "value": v, "is_header": is_hdr})
                    if not is_hdr: row_records[key][ch] = v
    meta = {"total_cells": len(corpus), "total_files": len({x["file"] for x in corpus}), "total_sheets": len({(x["file"], x["sheet"]) for x in corpus}), "total_rows": len(row_records), "locations": sorted({x["location"] for x in corpus})}
    return corpus, dict(row_records), meta
 
# ═══════════════════════ SIDEBAR ═══════════════════════
st.sidebar.image("https://img.icons8.com/fluency/96/data-center.png", width=70)
st.sidebar.title("🏢 Capacity Tracker"); st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Source")
uploaded_files = st.sidebar.file_uploader("Upload Excel files", type=["xlsx","xls"], accept_multiple_files=True)
if uploaded_files:
    file_bytes = tuple((f.name, f.read()) for f in uploaded_files)
    data_dir = save_uploads(file_bytes)
else:
    data_dir = str(EXCEL_FOLDER)
excel_files = find_excel_files(data_dir)
if not excel_files:
    st.error("### ⚠️ No Excel files found\n\nCreate `excel_files/` folder or upload files via the sidebar.")
    st.stop()
loc_map = {f: location_from_name(f) for f in excel_files}
st.sidebar.success(f"✅ {len(excel_files)} file(s) found")
st.sidebar.subheader("🏙️ Location")
selected_file = st.sidebar.selectbox("Location", excel_files, format_func=lambda x: loc_map[x])
all_sheets = load_file(os.path.join(data_dir, selected_file))
st.sidebar.subheader("📋 Sheet")
selected_sheet = st.sidebar.selectbox("Sheet", list(all_sheets.keys()))
raw_df = all_sheets[selected_sheet]
df_clean = to_numeric(smart_header(raw_df))
num_cols = df_clean.select_dtypes(include="number").columns.tolist()
cat_cols = [c for c in df_clean.columns if c not in num_cols]
st.sidebar.markdown("---")
st.sidebar.caption(f"📊 {len(num_cols)} numeric · {len(df_clean)} rows · {len(excel_files)} file(s)")
 
with st.spinner("🔍 Indexing every cell across all files…"):
    corpus, row_records, meta = build_corpus(tuple(excel_files), data_dir)
if not corpus:
    st.error("⚠️ **No data indexed.** Upload files via the sidebar."); st.stop()
 
with st.spinner("🔍 Indexing selected sheet…"):
    sq_cells, sq_rows, sq_meta = index_single_sheet(os.path.join(data_dir, selected_file), selected_sheet)
 
tabs = st.tabs(["🏠 Overview","📋 Raw Data","📊 Analytics","📈 Charts","🥧 Distributions","🔍 Query Engine","🌍 Multi-Location","🤖 AI Agent","💬 AI Smart Query"])
loc_label = loc_map[selected_file]
 
# ═══════ TAB 0 ═══════
with tabs[0]:
    st.title(f"🏢 {loc_label}  ›  {selected_sheet}")
    st.caption(f"File: `{selected_file}` | Raw {raw_df.shape[0]}×{raw_df.shape[1]} | Clean {len(df_clean)}×{len(df_clean.columns)} | Corpus: **{meta['total_cells']:,}** cells")
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.markdown(f'<div class="kcard kcard-blue"><h2>{len(df_clean)}</h2><p>Data Rows</p></div>',unsafe_allow_html=True)
    c2.markdown(f'<div class="kcard kcard-green"><h2>{len(df_clean.columns)}</h2><p>Columns</p></div>',unsafe_allow_html=True)
    c3.markdown(f'<div class="kcard kcard-purple"><h2>{len(num_cols)}</h2><p>Numeric</p></div>',unsafe_allow_html=True)
    c4.markdown(f'<div class="kcard kcard-orange"><h2>{len(excel_files)}</h2><p>Files</p></div>',unsafe_allow_html=True)
    c5.markdown(f'<div class="kcard kcard-teal"><h2>{meta["total_cells"]:,}</h2><p>Cells Indexed</p></div>',unsafe_allow_html=True)
    c6.markdown(f'<div class="kcard kcard-red"><h2>{int(df_clean.isna().sum().sum())}</h2><p>Missing</p></div>',unsafe_allow_html=True)
    st.markdown("---")
    if num_cols:
        st.markdown('<div class="sec-title">📐 Quick Statistics</div>',unsafe_allow_html=True)
        stats = df_clean[num_cols].describe().T; stats["range"] = stats["max"]-stats["min"]
        st.dataframe(stats.style.format("{:.3f}",na_rep="—").background_gradient(cmap="Blues",subset=["mean","max"]),use_container_width=True)
    st.markdown('<div class="sec-title">🗂️ Column Overview</div>',unsafe_allow_html=True)
    ci = pd.DataFrame({"Column":df_clean.columns,"Type":df_clean.dtypes.values,"Non-Null":df_clean.notna().sum().values,"Null%":(df_clean.isna().mean()*100).round(1).values,"Unique":[df_clean[c].nunique() for c in df_clean.columns],"Sample":[str(df_clean[c].dropna().iloc[0])[:55] if df_clean[c].dropna().shape[0]>0 else "—" for c in df_clean.columns]})
    st.dataframe(ci,use_container_width=True)
 
# ═══════ TAB 1 ═══════
with tabs[1]:
    st.subheader("📋 Data Table"); srch = st.text_input("🔍 Live search","",key="rawsrch")
    disp = (df_clean[df_clean.apply(lambda col:col.astype(str).str.contains(srch,case=False,na=False)).any(axis=1)] if srch else df_clean)
    st.caption(f"Showing {len(disp):,} / {len(df_clean):,} rows"); st.dataframe(disp,use_container_width=True,height=500)
    st.download_button("⬇️ CSV",disp.to_csv(index=False).encode(),"export.csv","text/csv")
    st.markdown("---"); st.subheader("🗃️ Raw Excel"); st.dataframe(raw_df,use_container_width=True,height=280)
 
# ═══════ TAB 2 ═══════
with tabs[2]:
    st.subheader("📊 Column Analytics")
    if not num_cols: st.info("No numeric columns.")
    else:
        chosen = st.multiselect("Select columns",num_cols,default=num_cols[:min(6,len(num_cols))])
        if chosen:
            sub = df_clean[chosen].dropna(how="all"); kc = st.columns(min(len(chosen),6))
            for i,col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s): kc[i].metric(col[:20],f"{s.sum():,.1f}",f"avg {s.mean():,.1f}")
            st.markdown("---"); agg_rows=[]
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    grand = df_clean[chosen].select_dtypes("number").sum().sum()
                    agg_rows.append({"Column":col,"Count":int(s.count()),"Sum":s.sum(),"Mean":s.mean(),"Median":s.median(),"Min":s.min(),"Max":s.max(),"Std":s.std(),"% Total":f"{s.sum()/grand*100:.1f}%" if grand else "—"})
            if agg_rows:
                adf = pd.DataFrame(agg_rows).set_index("Column")
                st.dataframe(adf.style.format("{:,.2f}",na_rep="—",subset=[c for c in adf.columns if c!="% Total"]).background_gradient(cmap="YlOrRd",subset=["Sum","Max"]),use_container_width=True)
        st.markdown("---"); st.markdown('<div class="sec-title">🧮 Group-By</div>',unsafe_allow_html=True)
        all_cat = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique()<60]
        if all_cat and num_cols:
            gc1,gc2,gc3 = st.columns(3); gc=gc1.selectbox("Group by",all_cat); ac=gc2.selectbox("Aggregate",num_cols); af=gc3.selectbox("Function",["sum","mean","count","min","max","median"])
            grp = df_clean.groupby(gc)[ac].agg(af).reset_index().rename(columns={ac:f"{af}({ac})"}).sort_values(f"{af}({ac})",ascending=False)
            st.dataframe(grp,use_container_width=True)
            fig = px.bar(grp,x=gc,y=f"{af}({ac})",color=f"{af}({ac})",color_continuous_scale="Viridis",title=f"{af.title()} of {ac} by {gc}"); fig.update_layout(xaxis_tickangle=-35,height=400); st.plotly_chart(fig,use_container_width=True)
 
# ═══════ TAB 3 ═══════
with tabs[3]:
    st.subheader("📈 Interactive Charts")
    ctype = st.selectbox("Chart Type",["Bar Chart","Grouped Bar","Line Chart","Scatter Plot","Area Chart","Bubble Chart","Heatmap (Correlation)","Box Plot","Funnel Chart","Waterfall / Cumulative","3-D Scatter"])
    if not num_cols: st.info("No numeric columns.")
    else:
        def _s(label,opts,idx=0,key=None): return st.selectbox(label,opts,index=min(idx,max(0,len(opts)-1)),key=key)
        if ctype=="Bar Chart":
            xc=_s("X",cat_cols or df_clean.columns.tolist(),key="bx"); yc=_s("Y",num_cols,key="by"); ori=st.radio("Orientation",["Vertical","Horizontal"],horizontal=True)
            d=df_clean[[xc,yc]].dropna(); fig=px.bar(d,x=xc if ori=="Vertical" else yc,y=yc if ori=="Vertical" else xc,color=yc,color_continuous_scale="Turbo",orientation="v" if ori=="Vertical" else "h",title=f"{yc} by {xc}"); fig.update_layout(height=480); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Grouped Bar":
            xc=_s("X",cat_cols or df_clean.columns.tolist(),key="gbx"); ycs=st.multiselect("Y",num_cols,default=num_cols[:3])
            if ycs: fig=px.bar(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs,barmode="group"); fig.update_layout(height=460); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Line Chart":
            xc=_s("X",df_clean.columns.tolist(),key="lx"); ycs=st.multiselect("Y",num_cols,default=num_cols[:2])
            if ycs: fig=px.line(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs,markers=True); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Scatter Plot":
            xc=_s("X",num_cols,0,"sc_x"); yc=_s("Y",num_cols,1,"sc_y"); sc=_s("Size",["None"]+num_cols,key="sc_s"); cc=_s("Color",["None"]+cat_cols+num_cols,key="sc_c")
            d=df_clean.dropna(subset=[xc,yc]); fig=px.scatter(d,x=xc,y=yc,size=sc if sc!="None" else None,color=cc if cc!="None" else None,color_continuous_scale="Rainbow"); fig.update_layout(height=480); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Area Chart":
            xc=_s("X",df_clean.columns.tolist(),key="ax"); ycs=st.multiselect("Y",num_cols,default=num_cols[:3])
            if ycs: fig=px.area(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Bubble Chart":
            if len(num_cols)>=3:
                xc=_s("X",num_cols,0,"bu_x"); yc=_s("Y",num_cols,1,"bu_y"); sz=_s("Size",num_cols,2,"bu_s"); lc=_s("Color",["None"]+cat_cols,key="bu_c"); d=df_clean[[xc,yc,sz]].dropna()
                if lc!="None": d[lc]=df_clean[lc]
                fig=px.scatter(d,x=xc,y=yc,size=sz,color=lc if lc!="None" else None,size_max=65); fig.update_layout(height=500); st.plotly_chart(fig,use_container_width=True)
            else: st.info("Need >= 3 numeric columns.")
        elif ctype=="Heatmap (Correlation)":
            sel=st.multiselect("Columns",num_cols,default=num_cols[:12])
            if len(sel)>=2: fig=px.imshow(df_clean[sel].corr(),text_auto=".2f",color_continuous_scale="RdBu_r",aspect="auto"); fig.update_layout(height=540); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Box Plot":
            yc=_s("Value",num_cols,key="bp_v"); xc=_s("Group",["None"]+cat_cols,key="bp_g"); d=df_clean[[yc]+([xc] if xc!="None" else [])].dropna(subset=[yc])
            fig=px.box(d,y=yc,x=xc if xc!="None" else None,color=xc if xc!="None" else None,points="outliers"); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Funnel Chart":
            xc=_s("Stage",cat_cols or df_clean.columns.tolist(),key="fn_x"); yc=_s("Value",num_cols,key="fn_y")
            d=df_clean[[xc,yc]].dropna().groupby(xc)[yc].sum().reset_index().sort_values(yc,ascending=False); fig=px.funnel(d,x=yc,y=xc); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Waterfall / Cumulative":
            yc=_s("Column",num_cols,key="wf_y"); d=df_clean[yc].dropna().reset_index(drop=True); cum=d.cumsum()
            fig=go.Figure(); fig.add_trace(go.Bar(name="Value",x=d.index,y=d,marker_color="#2a5298")); fig.add_trace(go.Scatter(name="Cumulative",x=cum.index,y=cum,line=dict(color="#f7971e",width=2.5),mode="lines+markers")); fig.update_layout(title=f"Cumulative: {yc}",height=450,barmode="group"); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="3-D Scatter":
            if len(num_cols)>=3:
                xc=_s("X",num_cols,0,"3x"); yc=_s("Y",num_cols,1,"3y"); zc=_s("Z",num_cols,2,"3z"); cc=_s("Color",["None"]+cat_cols,key="3c"); d=df_clean[[xc,yc,zc]].dropna()
                if cc!="None": d[cc]=df_clean[cc]
                fig=px.scatter_3d(d,x=xc,y=yc,z=zc,color=cc if cc!="None" else None); fig.update_layout(height=550); st.plotly_chart(fig,use_container_width=True)
            else: st.info("Need >= 3 numeric columns.")
 
# ═══════ TAB 4 ═══════
with tabs[4]:
    st.subheader("🥧 Distribution Charts")
    if not num_cols: st.info("No numeric columns.")
    else:
        r1,r2 = st.columns(2)
        with r1:
            st.markdown('<div class="sec-title">🍕 Pie / Donut</div>',unsafe_allow_html=True)
            pc_list=[c for c in df_clean.columns if c not in num_cols and 1<df_clean[c].nunique()<=30]
            if pc_list:
                pc=st.selectbox("Category",pc_list,key="pcat"); pv=st.selectbox("Value",num_cols,key="pval")
                pd_=df_clean[[pc,pv]].copy(); pd_[pv]=pd.to_numeric(pd_[pv],errors="coerce"); pd_=pd_.dropna().groupby(pc)[pv].sum().reset_index()
                fig=px.pie(pd_,names=pc,values=pv,hole=.38,color_discrete_sequence=px.colors.qualitative.Vivid); st.plotly_chart(fig,use_container_width=True)
        with r2:
            st.markdown('<div class="sec-title">📊 Histogram</div>',unsafe_allow_html=True)
            hc=st.selectbox("Column",num_cols,key="hcol"); bins=st.slider("Bins",5,100,25)
            fig=px.histogram(df_clean[hc].dropna(),nbins=bins,color_discrete_sequence=["#17a572"]); fig.update_layout(showlegend=False); st.plotly_chart(fig,use_container_width=True)
        st.markdown('<div class="sec-title">🎻 Violin</div>',unsafe_allow_html=True)
        vc=st.selectbox("Column",num_cols,key="vc"); fig=px.violin(df_clean[vc].dropna(),y=vc,box=True,points="outliers",color_discrete_sequence=["#c0392b"]); st.plotly_chart(fig,use_container_width=True)
 
# ═══════ TAB 5 ═══════
with tabs[5]:
    st.subheader("🔍 Query Engine  (selected sheet only)")
    st.info("For **smart search** use the **💬 AI Smart Query** tab.")
    query = st.text_input("Question",placeholder="e.g. Total subscription / Max capacity / List customers")
    def run_query(q,df,nc):
        ql=q.lower(); res=[]
        if any(w in ql for w in ["sum","total"]): [res.append(f"**SUM `{c}`** = {df[c].sum():,.4f}") for c in nc if c.lower() in ql or "all" in ql or len(nc)==1]
        if any(w in ql for w in ["average","mean","avg"]): [res.append(f"**MEAN `{c}`** = {df[c].mean():,.4f}") for c in nc if c.lower() in ql or len(nc)==1]
        if any(w in ql for w in ["maximum","highest","max"]): [res.append(f"**MAX `{c}`** = {df[c].max():,.4f}") for c in nc if c.lower() in ql or len(nc)==1]
        if any(w in ql for w in ["minimum","lowest","min"]): [res.append(f"**MIN `{c}`** = {df[c].min():,.4f}") for c in nc if c.lower() in ql or len(nc)==1]
        if any(w in ql for w in ["count","how many"]): res.append(f"**Rows** = {len(df):,}")
        if any(w in ql for w in ["customer","list","show"]):
            for c in df.columns:
                if "customer" in c.lower() or "name" in c.lower():
                    nm=df[c].dropna().unique(); res.append(f"**`{c}`** ({len(nm)}):\n"+"\n".join(f"  • {n}" for n in nm[:30])); break
        if not res: res.append("ℹ️ Try: **sum / average / max / min / count / list**")
        return "\n\n".join(res)
    if query: st.markdown(run_query(query,df_clean,num_cols))
    st.markdown("---"); st.markdown('<div class="sec-title">🧮 Manual Compute</div>',unsafe_allow_html=True)
    if num_cols:
        mc1,mc2,mc3=st.columns(3); op=mc1.selectbox("Op",["Sum","Mean","Max","Min","Count","Median","Std Dev","Range"]); sc_col=mc2.selectbox("Column",num_cols,key="mc_col"); fc=mc3.selectbox("Filter by",["None"]+[c for c in df_clean.columns if c not in num_cols])
        fv=None
        if fc!="None": fv=st.selectbox("Filter value",df_clean[fc].dropna().unique().tolist())
        ds=df_clean.copy()
        if fc!="None" and fv is not None: ds=ds[ds[fc]==fv]
        s=ds[sc_col].dropna()
        ops={"Sum":s.sum(),"Mean":s.mean(),"Max":s.max(),"Min":s.min(),"Count":s.count(),"Median":s.median(),"Std Dev":s.std(),"Range":s.max()-s.min()}
        r=ops.get(op,"N/A")
        if isinstance(r,float): r=f"{r:,.4f}"
        st.success(f"**{op}** of `{sc_col}`{f' (where {fc}={fv})' if fv else ''} -> **{r}**")
 
# ═══════ TAB 6 ═══════
with tabs[6]:
    st.subheader("🌍 Cross-Location Comparison")
    @st.cache_data(show_spinner=False)
    def load_all_summ(files,folder):
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
                s=info["df"][comp_col].dropna(); rows.append({"Location|Sheet":lbl,"Sum":s.sum(),"Mean":s.mean(),"Max":s.max(),"Min":s.min(),"Count":s.count()})
        if rows:
            cmp=pd.DataFrame(rows).set_index("Location|Sheet"); st.dataframe(cmp.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),use_container_width=True)
            fig=px.bar(cmp.reset_index(),x="Location|Sheet",y="Sum",color="Sum",color_continuous_scale="Viridis",title=f"Sum of '{comp_col}'"); fig.update_layout(xaxis_tickangle=-30,height=440); st.plotly_chart(fig,use_container_width=True)
 
# ═══════ TAB 7 ═══════
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights")
    if st.button("🚀 Run Analysis",type="primary"):
        with st.spinner("Analysing…"):
            for lbl,info in list(all_summ.items())[:10]:
                dfa=info["df"]; nc=info["num_cols"]
                if not nc: continue
                with st.expander(f"📍 {lbl}",expanded=False):
                    ca,cb=st.columns(2)
                    with ca:
                        for col in nc[:5]:
                            s=dfa[col].dropna()
                            if len(s): st.metric(col[:26],f"{s.sum():,.1f}",f"avg {s.mean():,.1f}")
                    with cb:
                        for col in nc[:4]:
                            s=dfa[col].dropna()
                            if len(s)>3:
                                z=(s-s.mean())/s.std(); o=z[z.abs()>2.5]
                                (st.warning if len(o) else st.success)(f"`{col}`: {len(o)} outlier(s)" if len(o) else f"`{col}`: Clean ✓")
    st.markdown("---"); st.markdown('<div class="sec-title">📁 Files Summary</div>',unsafe_allow_html=True)
    fsm=[]
    for f in excel_files:
        shd = all_sheets if f==selected_file else load_file(os.path.join(data_dir,f))
        fsm.append({"File":loc_map[f],"Sheets":len(shd),"Rows":sum(len(s) for s in shd.values())})
    st.dataframe(pd.DataFrame(fsm),use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 8 – AI SMART QUERY  (FIXED: single sheet, no self-asking)
# ═══════════════════════════════════════════════════════════════════════════
with tabs[8]:
    st.markdown("## 💬 AI Smart Query")
    st.markdown(f"Querying: **{loc_label}** › **{selected_sheet}** — change in sidebar.")
 
    qi1, qi2, qi3 = st.columns(3)
    qi1.markdown(f'<div class="kcard kcard-blue"><h2>{sq_meta["total_cells"]:,}</h2><p>Cells in Sheet</p></div>', unsafe_allow_html=True)
    qi2.markdown(f'<div class="kcard kcard-green"><h2>{sq_meta["total_data"]:,}</h2><p>Data Cells</p></div>', unsafe_allow_html=True)
    qi3.markdown(f'<div class="kcard kcard-purple"><h2>{sq_meta["total_rows"]:,}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
 
    def _is_num(v):
        try: float(v); return True
        except: return False
 
    # FIX: _OP only has PURE OPERATOR words. Domain keywords removed so they
    # work as column-matching keywords in npkw/npbest.
    _OP_VERBS = {"total","sum","avg","mean","max","min","count","list","find","show",
                 "all","average","maximum","minimum","highest","lowest","top","bottom",
                 "describe","statistics","stats","summary","unique","distinct",
                 "sheet","column","row","missing","null","percent","percentage",
                 "ratio","share","number","across","compare"}
    _SYN = {
        "subscription": ["subscription","subscribed","subscript"],
        "capacity": ["capacity","capac"],
        "power": ["power","kw","kva"],
        "usage": ["usage","utilization","consumption","consumed"],
        "rack": ["rack","racks"],
        "space": ["space","sqft","sq ft"],
        "customer": ["customer","name","customers"],
        "billing": ["billing","bill"],
        "ownership": ["ownership","owned"],
    }
 
    def _mcol(kw, hdr):
        hl = hdr.lower(); kwl = kw.lower()
        if kwl in hl: return True
        for key, syns in _SYN.items():
            if kwl in syns or kwl == key:
                for s in syns:
                    if s in hl: return True
        return False
 
    def sheet_query(question):
        q = question.strip(); ql = q.lower()
        sig = [w for w in re.findall(r"[a-z0-9]{3,}", ql) if w not in _SW]
 
        f_sum = any(x in ql for x in ["total","sum","aggregate"])
        f_avg = any(x in ql for x in ["average","mean","avg"])
        f_max = any(x in ql for x in ["maximum","highest","largest","max"])
        f_min = any(x in ql for x in ["minimum","lowest","smallest","min"])
        f_cnt = any(x in ql for x in ["count","how many","number of"])
        f_stat = any(x in ql for x in ["statistics","stats","describe","summary"])
        f_uniq = any(x in ql for x in ["unique","distinct","different"])
        f_miss = any(x in ql for x in ["missing","null","blank","empty"])
        f_cols = any(x in ql for x in ["column","columns","field","header"])
        f_topn = re.search(r"\btop\s*(\d+)\b", ql)
        f_botn = re.search(r"\bbottom\s*(\d+)\b", ql)
        f_num = f_sum or f_avg or f_max or f_min or f_cnt or f_stat or f_topn or f_botn
        f_cust = any(x in ql for x in ["customer","customers","client","clients"])
        f_list = any(x in ql for x in ["list","show","name","names"])
 
        out = {"answer": "", "table": None, "chart_df": None, "chart_cfg": None, "cell_hits": [], "sub_tables": []}
        wc = sq_cells; rr = sq_rows
 
        if not wc:
            out["answer"] = "❓ No data in this sheet."; return out
 
        # Keywords for column matching: sig words NOT in pure operator verbs
        col_kws = [w for w in sig if w not in _OP_VERBS]
 
        def npkw(kw):
            res = []
            for cell in wc:
                if cell["is_header"]: continue
                if _mcol(kw, cell["col_header"]):
                    try: res.append((float(cell["value"]), cell))
                    except: pass
            return res
 
        def npbest(kws):
            bk, bp = None, []
            for w in kws:
                p = npkw(w)
                if len(p) > len(bp): bp = p; bk = w
            return bk, bp
 
        def build_rows_df(row_nums):
            recs = []
            for rn in sorted(row_nums):
                rec = rr.get(rn, {})
                if rec:
                    rd = {"Row #": rn + 1}; rd.update(rec); recs.append(rd)
            return pd.DataFrame(recs) if recs else pd.DataFrame()
 
        # ── INTENT: List customers / names ──
        if (f_cust or f_list) and not f_num:
            found = []
            for cell in wc:
                if cell["is_header"]: continue
                ch = cell["col_header"].lower()
                if "customer" in ch or "name" in ch:
                    found.append({"Row #": cell["row"]+1, "Column": cell["col_header"], "Value": cell["value"]})
            if found:
                tbl = pd.DataFrame(found).drop_duplicates(subset=["Value"])
                out["answer"] = f"Found **{len(tbl)}** customer/name entries."
                out["table"] = tbl
            else:
                out["answer"] = "No columns with 'Customer' or 'Name' in header found in this sheet."
            return out
 
        # ── INTENT: Missing values ──
        if f_miss:
            dfc = to_numeric(smart_header(raw_df)); mr = []
            for col in dfc.columns:
                mc = int(dfc[col].isna().sum())
                if mc > 0: mr.append({"Column": col, "Missing": mc, "Missing%": f"{mc/max(len(dfc),1)*100:.1f}%"})
            if mr: tbl = pd.DataFrame(mr).sort_values("Missing", ascending=False); out["answer"] = f"Found **{len(tbl)}** column(s) with missing values."; out["table"] = tbl
            else: out["answer"] = "✅ No missing values found."
            return out
 
        # ── INTENT: Column listing ──
        if f_cols and not f_num:
            seen = set(); cr = []
            for cell in wc:
                if not cell["is_header"]: continue
                ch = cell["value"].strip()
                if ch in ("", "nan"): continue
                if ch not in seen: seen.add(ch); cr.append({"Column": ch, "At Row": cell["row"]+1, "At Col": cell["col"]+1})
            tbl = pd.DataFrame(cr) if cr else pd.DataFrame()
            out["answer"] = f"Found **{len(tbl)}** column header(s)."; out["table"] = tbl; return out
 
        # ── INTENT: Numeric aggregation (sum/avg/max/min/count/stats/top N) ──
        if f_num:
            kw, pairs = npbest(col_kws)
            # Fallback: try domain keywords from query
            if not pairs:
                for dkw in ["subscription","capacity","power","usage","rack","space","consumption","kw","kva","sqft"]:
                    if dkw in ql:
                        pairs = npkw(dkw)
                        if pairs: kw = dkw; break
            if pairs:
                vals = [v for v, _ in pairs]; sa = pd.Series(vals); parts = []
                if f_sum or f_stat: parts.append(f"**Total (Sum):** {sa.sum():,.4f}")
                if f_avg or f_stat: parts.append(f"**Average:** {sa.mean():,.4f}")
                if f_max or f_stat: parts.append(f"**Maximum:** {sa.max():,.4f}")
                if f_min or f_stat: parts.append(f"**Minimum:** {sa.min():,.4f}")
                if f_cnt or f_stat: parts.append(f"**Count:** {sa.count():,}")
                if f_stat: parts.append(f"**Median:** {sa.median():,.4f} | **Std Dev:** {sa.std():,.4f}")
                if (f_topn or f_botn) and not (f_sum or f_avg or f_max or f_min or f_cnt or f_stat):
                    parts.append(f"**Count:** {sa.count():,}")
                    parts.append(f"**Total (Sum):** {sa.sum():,.4f}")
                detail = [{"Row #": c["row"]+1, "Column": c["col_header"], "Value": v} for v, c in pairs]
                tbl = pd.DataFrame(detail).sort_values("Value", ascending=False)
                out["answer"] = f"Results for **'{kw}'** ({len(vals):,} values):\n\n" + "\n".join(parts)
                out["table"] = tbl
                if f_topn:
                    n = int(f_topn.group(1)); top = sorted(pairs, key=lambda x: x[0], reverse=True)[:n]
                    out["sub_tables"].append({"label": f"🏆 Top {n} — {kw}", "df": pd.DataFrame([{"Row": c["row"]+1, "Column": c["col_header"], "Value": v} for v, c in top])})
                if f_botn:
                    n = int(f_botn.group(1)); bot = sorted(pairs, key=lambda x: x[0])[:n]
                    out["sub_tables"].append({"label": f"🔻 Bottom {n} — {kw}", "df": pd.DataFrame([{"Row": c["row"]+1, "Column": c["col_header"], "Value": v} for v, c in bot])})
                return out
            # Fallback: all numeric
            anums = [(float(c["value"]), c) for c in wc if not c["is_header"] and _is_num(c["value"])]
            if anums:
                vals = [v for v, _ in anums]; sa = pd.Series(vals); parts = []
                if f_sum: parts.append(f"**Sum ALL numeric:** {sa.sum():,.4f}")
                if f_avg: parts.append(f"**Avg ALL numeric:** {sa.mean():,.4f}")
                if f_max: parts.append(f"**Max ALL numeric:** {sa.max():,.4f}")
                if f_min: parts.append(f"**Min ALL numeric:** {sa.min():,.4f}")
                if f_cnt: parts.append(f"**Count ALL numeric:** {sa.count():,}")
                out["answer"] = f"No column matched keywords. ALL numeric cells:\n\n" + "\n".join(parts); return out
 
        # ── INTENT: Unique values ──
        if f_uniq:
            for w in col_kws or sig:
                uv = set(); sr = []
                for cell in wc:
                    if cell["is_header"]: continue
                    if _mcol(w, cell["col_header"]):
                        uv.add(cell["value"]); sr.append({"Column": cell["col_header"], "Value": cell["value"], "Row #": cell["row"]+1})
                if uv:
                    tbl = pd.DataFrame(sr).drop_duplicates(subset=["Value"]) if sr else pd.DataFrame()
                    out["answer"] = f"**{len(uv)}** unique value(s) for **'{w}'**."; out["table"] = tbl; return out
 
        # ── INTENT: Free-text entity/keyword search ──
        if sig:
            quoted = re.findall(r'"([^"]+)"', q)
            if quoted:
                terms = [quoted[0].lower()]
            else:
                terms = [w for w in sig if any(w in cell["value"].lower() for cell in wc)]
                if not terms: terms = sig
            hit_cells = [cell for cell in wc if not cell["is_header"] and any(t in cell["value"].lower() for t in terms)]
            hit_rows = sorted({c["row"] for c in hit_cells})
            full_df = build_rows_df(hit_rows)
            cell_list = [{"Row #": c["row"]+1, "Col #": c["col"]+1, "Column Header": c["col_header"], "Value": c["value"]} for c in hit_cells[:80]]
            out["answer"] = f"Found **{len(hit_cells):,}** cell(s) matching **'{', '.join(terms[:4])}'** in **{len(hit_rows):,}** row(s)."
            out["table"] = full_df if not full_df.empty else None; out["cell_hits"] = cell_list
            if not full_df.empty:
                for col in full_df.columns:
                    if "customer" in col.lower() or "name" in col.lower():
                        cdf = full_df[["Row #", col]].drop_duplicates()
                        out["sub_tables"].append({"label": f"👤 Customers ({len(cdf)})", "df": cdf}); break
            return out
 
        out["answer"] = "❓ No match.\n\n**Try:** *List all customers* | *Find CISCO* | *Total subscription* | *Top 10 subscription* | *Show columns*"
        return out
 
    def render_answer(res, tidx=0):
        st.markdown(f'<div class="ans-box">{res["answer"]}</div>', unsafe_allow_html=True)
        if res.get("table") is not None and not res["table"].empty:
            tbl = res["table"].reset_index(drop=True)
            st.dataframe(tbl, use_container_width=True, height=min(520, 48+len(tbl)*36), key=f"tbl_{tidx}")
            st.download_button("⬇️ CSV", tbl.to_csv(index=False).encode(), "result.csv", "text/csv", key=f"dl_{tidx}")
        if res.get("chart_cfg") and res.get("chart_df") is not None:
            cfg = res["chart_cfg"]; cdf = res["chart_df"]
            if cfg["x"] in cdf.columns and cfg["y"] in cdf.columns:
                fig = px.bar(cdf.sort_values(cfg["y"], ascending=False).head(30), x=cfg["x"], y=cfg["y"], color=cfg["y"], color_continuous_scale="Viridis", title=cfg["title"], height=400)
                fig.update_layout(xaxis_tickangle=-30); st.plotly_chart(fig, use_container_width=True, key=f"ch_{tidx}")
        for si, s in enumerate(res.get("sub_tables", [])):
            with st.expander(s["label"], expanded=True):
                st.dataframe(s["df"], use_container_width=True, key=f"sub_{tidx}_{si}")
        if res.get("cell_hits"):
            with st.expander(f"🔬 Cell matches ({len(res['cell_hits'])})", expanded=False):
                for ch in res["cell_hits"]:
                    st.markdown(f'<div class="cell-chip">R{ch["Row #"]} C{ch["Col #"]} | <i>{ch["Column Header"]}</i> | <b>{ch["Value"]}</b></div>', unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)
 
    hist_key = f"sq_hist_{selected_file}_{selected_sheet}"
    if hist_key not in st.session_state: st.session_state[hist_key] = []
    for tidx, turn in enumerate(st.session_state[hist_key]):
        st.markdown(f'<div class="q-user">🧑 {turn["q"]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)
        render_answer(turn["res"], tidx); st.markdown("---")
 
    st.markdown("---")
    ic, bc, cc = st.columns([8, 1, 1])
    with ic: user_q = st.text_input("Ask:", placeholder="Find CISCO | List customers | Total subscription | Top 10 subscription", label_visibility="collapsed", key="sq_input")
    with bc: ask_btn = st.button("🔍 Ask", use_container_width=True, type="primary")
    with cc:
        if st.button("🗑️ Clear", use_container_width=True): st.session_state[hist_key] = []; st.rerun()
 
    st.markdown("**💡 Examples:**")
    examples = [
        ["List all customers", "Find CISCO", "Find AT&T", "Find Axis Bank", "Find MOTMOT", "Find Zscaler"],
        ["Total subscription", "Max capacity", "Average power usage", "Top 10 subscription", "Statistics of subscription", "Count rows"],
        ["Show columns", "Show missing values", "Unique billing models", "Unique ownership"],
    ]
    for row in examples:
        cols = st.columns(len(row))
        for j, ex in enumerate(row):
            if cols[j].button(ex, key=f"sqchip_{ex}", use_container_width=True): user_q = ex; ask_btn = True
 
    if ask_btn and user_q.strip():
        with st.spinner(f"Scanning sheet for: **{user_q}** …"):
            answer = sheet_query(user_q)
        st.session_state[hist_key].append({"q": user_q, "res": answer}); st.rerun()
 
st.markdown("---")
st.caption(f"Sify DC · Capacity Tracker · {meta['total_cells']:,} cells indexed · {', '.join(meta['locations'])}")import os, re, warnings, tempfile, subprocess
from collections import defaultdict
from pathlib import Path
 
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
 
warnings.filterwarnings("ignore")
 
st.set_page_config(page_title="Sify DC – Capacity Tracker", page_icon="🏢",
                   layout="wide", initial_sidebar_state="expanded")
 
st.markdown("""
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0e1a,#1a2035,#0d1b2a)!important;}
[data-testid="stSidebar"] *{color:#c9d8f0!important;}
.kcard{border-radius:14px;padding:16px 20px;color:#fff;margin-bottom:10px;box-shadow:0 4px 18px rgba(0,0,0,.35);transition:transform .2s;}
.kcard:hover{transform:translateY(-2px);}
.kcard h2{font-size:1.8rem;margin:0;font-weight:800;}
.kcard p{margin:3px 0 0;font-size:.82rem;opacity:.82;}
.kcard-blue{background:linear-gradient(135deg,#1e3c72,#2a5298);}
.kcard-green{background:linear-gradient(135deg,#0b6e4f,#17a572);}
.kcard-red{background:linear-gradient(135deg,#7b1a1a,#c0392b);}
.kcard-orange{background:linear-gradient(135deg,#7d4e00,#e67e22);}
.kcard-teal{background:linear-gradient(135deg,#0f3460,#16213e);}
.kcard-purple{background:linear-gradient(135deg,#4a0072,#7b1fa2);}
.sec-title{font-size:1.15rem;font-weight:700;color:#1e3c72;border-left:5px solid #2a5298;padding-left:10px;margin:16px 0 10px;}
.q-user{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff;border-radius:18px 18px 4px 18px;padding:10px 16px;margin:10px 0 4px auto;max-width:76%;width:fit-content;box-shadow:0 3px 12px rgba(30,60,114,.45);float:right;clear:both;}
.ans-box{background:linear-gradient(135deg,#0f2744,#1a4a6b);color:#d0ecff;border-radius:12px;padding:14px 18px;margin:8px 0;font-size:.97rem;box-shadow:0 3px 14px rgba(0,0,0,.35);white-space:pre-wrap;line-height:1.6;}
.cell-chip{background:#1a2f1a;border-left:4px solid #27ae60;border-radius:6px;padding:6px 12px;margin:3px 0;font-family:monospace;font-size:.8rem;color:#b8ffb8;}
.clearfix{clear:both;}
</style>
""", unsafe_allow_html=True)
 
_SW = {"the","and","for","are","all","any","how","what","show","give","tell","from","this","that","with","get","find","list","much","many","each","every","data","value","values","number","numbers","in","of","a","an","is","at","by","to","do","me","my","about","details","info","please","can","you","per","across","which","where","who","when","does","did","have","has","their","its","our","your","there","these","those","been","will","would","could","should","shall","let","some","just","also","even","only","into","over","under","both","such","than","then","but","not","nor","yet","so","either","neither","versus","vs"}
 
def _app_dir():
    try: return Path(__file__).resolve().parent
    except NameError: return Path(os.getcwd())
EXCEL_FOLDER = _app_dir() / "excel_files"
 
def find_excel_files(folder):
    p = Path(folder)
    if not p.is_dir(): return []
    return sorted(f.name for f in p.iterdir() if f.suffix.lower() in (".xlsx",".xls") and not f.name.startswith("~"))
 
def location_from_name(fname):
    n = os.path.basename(fname)
    n = re.sub(r"\.(xlsx?|xls)$","",n,flags=re.I)
    n = re.sub(r"[Cc]ustomer.?[Aa]nd.?[Cc]apacity.?[Tt]racker.?","",n)
    n = re.sub(r"[_\s]?\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{2,4}.*$","",n,flags=re.I)
    n = re.sub(r"__\d+_*$","",n)
    n = re.sub(r"[_]+"," ",n).strip()
    return n if n else fname
 
@st.cache_data(show_spinner=False)
def save_uploads(file_bytes_tuple):
    tmp = tempfile.mkdtemp()
    for name, data in file_bytes_tuple:
        with open(os.path.join(tmp, name), "wb") as fh: fh.write(data)
    return tmp
 
@st.cache_data(show_spinner=False)
def ensure_readable(original_path):
    if not original_path.lower().endswith(".xls"): return original_path
    try:
        with open(original_path,"rb") as fh:
            if fh.read(4) != b"\xd0\xcf\x11\xe0": return original_path
    except: return original_path
    out_dir = tempfile.mkdtemp()
    wrapper = "/mnt/skills/public/xlsx/scripts/office/soffice.py"
    try:
        if os.path.exists(wrapper):
            subprocess.run(["python3",wrapper,"--convert-to","xlsx","--outdir",out_dir,original_path], capture_output=True, timeout=60)
        else:
            subprocess.run(["libreoffice","--headless","--convert-to","xlsx","--outdir",out_dir,original_path], capture_output=True, timeout=120)
        base = os.path.splitext(os.path.basename(original_path))[0]
        conv = os.path.join(out_dir, base+".xlsx")
        if os.path.exists(conv): return conv
    except: pass
    return original_path
 
@st.cache_data(show_spinner=False)
def _read_sheet(path, sheet_name):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    mr = ws.max_row or 0; mc = ws.max_column or 0
    if mr == 0: wb.close(); return pd.DataFrame()
    real_mc = 0
    samples = sorted(set(list(range(1, min(31, mr+1))) + list(range(max(1, mr-9), mr+1))))
    for r in samples:
        for cell in ws[r]:
            if cell.value is not None: real_mc = max(real_mc, cell.column)
    if real_mc == 0: wb.close(); return pd.DataFrame()
    cap = min(real_mc + 2, mc)
    rows = []
    for row in ws.iter_rows(min_row=1, max_row=mr, max_col=cap, values_only=True):
        rows.append(list(row))
    wb.close()
    if not rows: return pd.DataFrame()
    df = pd.DataFrame(rows, dtype=str)
    df = df.replace({"None": np.nan, "none": np.nan})
    return df
 
@st.cache_data(show_spinner=False)
def load_file(original_path):
    path = ensure_readable(original_path)
    sheets = {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True); names = wb.sheetnames; wb.close()
        for sh in names:
            try:
                df = _read_sheet(path, sh)
                if not df.empty: sheets[sh] = df
            except: pass
    except Exception as e:
        st.sidebar.warning(f"⚠️ {os.path.basename(original_path)}: {e}")
    return sheets
 
def best_header_row(df):
    best_row, best_score = 0, -1
    for i in range(min(8, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        filled = (row.str.len()>0)&(~row.isin(["nan","None",""]))
        label = filled & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        score = label.sum()*2 + filled.sum()
        if score > best_score: best_score, best_row = score, i
    return best_row
 
def smart_header(df):
    hr = best_header_row(df)
    hdr = df.iloc[hr].fillna("").astype(str).str.strip()
    seen = {}; cols = []
    for col in hdr:
        col = col if col and col not in ("nan","None") else f"Col_{len(cols)}"
        if col in seen: seen[col]+=1; cols.append(f"{col}_{seen[col]}")
        else: seen[col]=0; cols.append(col)
    data = df.iloc[hr+1:].copy(); data.columns = cols
    return data.dropna(how="all").reset_index(drop=True)
 
def to_numeric(df):
    out = df.copy()
    for col in out.columns: out[col] = pd.to_numeric(out[col], errors="ignore")
    return out
 
def _detect_all_header_rows(df):
    hr_set = set()
    for i in range(len(df)):
        row = df.iloc[i].astype(str).str.strip()
        fm = (row.str.len()>0) & (~row.isin(["nan","None",""]))
        fv = row[fm]; nf = fv.shape[0]
        if nf < 2: continue
        lm = fm & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        nl = lm.sum(); nu = fv.nunique()
        lr = nl/max(nf,1); ur = nu/max(nf,1)
        vc = fv.value_counts(); nr = (vc>1).sum()
        if lr>=0.80 and ur>=0.75 and nr<=max(2, nf*0.15) and nu>=3: hr_set.add(i)
        elif nf<=10 and nf>=2 and lr>=0.90 and ur>=0.80 and nr<=1: hr_set.add(i)
    return hr_set
 
def _build_cell_col_map(df):
    hr_set = _detect_all_header_rows(df)
    hr_maps = {}
    for hr in hr_set:
        m = {}
        for c in range(df.shape[1]):
            v = str(df.iat[hr,c]).strip()
            if v and v not in ("nan","None"): m[c] = v
        hr_maps[hr] = m
    sorted_hrs = sorted(hr_set)
    cell_map = {}
    for r in range(df.shape[0]):
        prev = [h for h in sorted_hrs if h < r]
        for c in range(df.shape[1]):
            name = f"Col_{c}"
            for h in reversed(prev):
                if c in hr_maps[h]: name = hr_maps[h][c]; break
            cell_map[(r,c)] = name
    return cell_map, hr_set
 
# ─── INDEX A SINGLE SHEET into cell list + row_records ───
@st.cache_data(show_spinner=False)
def index_single_sheet(file_path, sheet_name):
    """Read every non-empty cell from one sheet. Returns (cells_list, row_records_dict, meta)."""
    sheets = load_file(file_path)
    if sheet_name not in sheets:
        return [], {}, {"total_cells": 0, "total_rows": 0}
    df = sheets[sheet_name]
    cell_map, hr_set = _build_cell_col_map(df)
    cells = []
    row_recs = {}
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            raw = df.iat[r, c]
            if pd.isna(raw): continue
            v = str(raw).strip()
            if not v or v in ("nan", "None", "none", ""): continue
            ch = cell_map.get((r, c), f"Col_{c}")
            is_hdr = (r in hr_set)
            cells.append({"row": r, "col": c, "col_header": ch, "value": v, "is_header": is_hdr})
            if not is_hdr:
                key = r
                if key not in row_recs: row_recs[key] = {}
                row_recs[key][ch] = v
    meta = {"total_cells": len(cells), "total_rows": len(row_recs),
            "total_data": sum(1 for x in cells if not x["is_header"]),
            "total_headers": sum(1 for x in cells if x["is_header"])}
    return cells, row_recs, meta
 
# ─── BUILD FULL CORPUS (for tabs 0-7 that still need it) ───
@st.cache_data(show_spinner=False)
def build_corpus(file_list, folder):
    corpus = []; row_records = defaultdict(dict)
    for fname in file_list:
        full = os.path.join(folder, fname)
        if not os.path.isfile(full): continue
        loc = location_from_name(fname)
        sheets = load_file(full)
        for sh, df in sheets.items():
            cell_map, hr_set = _build_cell_col_map(df)
            for r in range(df.shape[0]):
                for c in range(df.shape[1]):
                    raw = df.iat[r, c]
                    if pd.isna(raw): continue
                    v = str(raw).strip()
                    if not v or v in ("nan", "None", "none", ""): continue
                    ch = cell_map.get((r, c), f"Col_{c}")
                    is_hdr = (r in hr_set)
                    key = (fname, loc, sh, r)
                    corpus.append({"file": fname, "location": loc, "sheet": sh, "row": r, "col": c, "col_header": ch, "value": v, "is_header": is_hdr})
                    if not is_hdr: row_records[key][ch] = v
    meta = {"total_cells": len(corpus), "total_files": len({x["file"] for x in corpus}), "total_sheets": len({(x["file"], x["sheet"]) for x in corpus}), "total_rows": len(row_records), "locations": sorted({x["location"] for x in corpus})}
    return corpus, dict(row_records), meta
 
# ═══════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════
st.sidebar.image("https://img.icons8.com/fluency/96/data-center.png", width=70)
st.sidebar.title("🏢 Capacity Tracker"); st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Source")
uploaded_files = st.sidebar.file_uploader("Upload Excel files", type=["xlsx","xls"], accept_multiple_files=True)
if uploaded_files:
    file_bytes = tuple((f.name, f.read()) for f in uploaded_files)
    data_dir = save_uploads(file_bytes)
else:
    data_dir = str(EXCEL_FOLDER)
excel_files = find_excel_files(data_dir)
if not excel_files:
    st.error("### ⚠️ No Excel files found\n\nCreate `excel_files/` folder or upload files via the sidebar.")
    st.stop()
loc_map = {f: location_from_name(f) for f in excel_files}
st.sidebar.success(f"✅ {len(excel_files)} file(s) found")
st.sidebar.subheader("🏙️ Location")
selected_file = st.sidebar.selectbox("Location", excel_files, format_func=lambda x: loc_map[x])
all_sheets = load_file(os.path.join(data_dir, selected_file))
st.sidebar.subheader("📋 Sheet")
selected_sheet = st.sidebar.selectbox("Sheet", list(all_sheets.keys()))
raw_df = all_sheets[selected_sheet]
df_clean = to_numeric(smart_header(raw_df))
num_cols = df_clean.select_dtypes(include="number").columns.tolist()
cat_cols = [c for c in df_clean.columns if c not in num_cols]
st.sidebar.markdown("---")
st.sidebar.caption(f"📊 {len(num_cols)} numeric · {len(df_clean)} rows · {len(excel_files)} file(s)")
 
with st.spinner("🔍 Indexing every cell across all files…"):
    corpus, row_records, meta = build_corpus(tuple(excel_files), data_dir)
if not corpus:
    st.error("⚠️ **No data indexed.** Upload files via the sidebar.")
    st.stop()
 
# Index the single selected sheet for Smart Query tab
with st.spinner("🔍 Indexing selected sheet…"):
    sq_cells, sq_rows, sq_meta = index_single_sheet(os.path.join(data_dir, selected_file), selected_sheet)
 
tabs = st.tabs(["🏠 Overview","📋 Raw Data","📊 Analytics","📈 Charts","🥧 Distributions","🔍 Query Engine","🌍 Multi-Location","🤖 AI Agent","💬 AI Smart Query"])
loc_label = loc_map[selected_file]
 
# ═══════════════════════════════════════════════════════
# TAB 0 – OVERVIEW
# ═══════════════════════════════════════════════════════
with tabs[0]:
    st.title(f"🏢 {loc_label}  ›  {selected_sheet}")
    st.caption(f"File: `{selected_file}` | Raw {raw_df.shape[0]}×{raw_df.shape[1]} | Clean {len(df_clean)}×{len(df_clean.columns)} | Corpus: **{meta['total_cells']:,}** cells")
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.markdown(f'<div class="kcard kcard-blue"><h2>{len(df_clean)}</h2><p>Data Rows</p></div>',unsafe_allow_html=True)
    c2.markdown(f'<div class="kcard kcard-green"><h2>{len(df_clean.columns)}</h2><p>Columns</p></div>',unsafe_allow_html=True)
    c3.markdown(f'<div class="kcard kcard-purple"><h2>{len(num_cols)}</h2><p>Numeric</p></div>',unsafe_allow_html=True)
    c4.markdown(f'<div class="kcard kcard-orange"><h2>{len(excel_files)}</h2><p>Files</p></div>',unsafe_allow_html=True)
    c5.markdown(f'<div class="kcard kcard-teal"><h2>{meta["total_cells"]:,}</h2><p>Cells Indexed</p></div>',unsafe_allow_html=True)
    c6.markdown(f'<div class="kcard kcard-red"><h2>{int(df_clean.isna().sum().sum())}</h2><p>Missing</p></div>',unsafe_allow_html=True)
    st.markdown("---")
    if num_cols:
        st.markdown('<div class="sec-title">📐 Quick Statistics</div>',unsafe_allow_html=True)
        stats = df_clean[num_cols].describe().T; stats["range"] = stats["max"]-stats["min"]
        st.dataframe(stats.style.format("{:.3f}",na_rep="—").background_gradient(cmap="Blues",subset=["mean","max"]),use_container_width=True)
    st.markdown('<div class="sec-title">🗂️ Column Overview</div>',unsafe_allow_html=True)
    ci = pd.DataFrame({"Column":df_clean.columns,"Type":df_clean.dtypes.values,"Non-Null":df_clean.notna().sum().values,"Null%":(df_clean.isna().mean()*100).round(1).values,"Unique":[df_clean[c].nunique() for c in df_clean.columns],"Sample":[str(df_clean[c].dropna().iloc[0])[:55] if df_clean[c].dropna().shape[0]>0 else "—" for c in df_clean.columns]})
    st.dataframe(ci,use_container_width=True)
 
# TAB 1 – RAW DATA
with tabs[1]:
    st.subheader("📋 Data Table")
    srch = st.text_input("🔍 Live search","",key="rawsrch")
    disp = (df_clean[df_clean.apply(lambda col:col.astype(str).str.contains(srch,case=False,na=False)).any(axis=1)] if srch else df_clean)
    st.caption(f"Showing {len(disp):,} / {len(df_clean):,} rows"); st.dataframe(disp,use_container_width=True,height=500)
    st.download_button("⬇️ CSV",disp.to_csv(index=False).encode(),"export.csv","text/csv")
    st.markdown("---"); st.subheader("🗃️ Raw Excel"); st.dataframe(raw_df,use_container_width=True,height=280)
 
# TAB 2 – ANALYTICS
with tabs[2]:
    st.subheader("📊 Column Analytics")
    if not num_cols: st.info("No numeric columns.")
    else:
        chosen = st.multiselect("Select columns",num_cols,default=num_cols[:min(6,len(num_cols))])
        if chosen:
            sub = df_clean[chosen].dropna(how="all"); kc = st.columns(min(len(chosen),6))
            for i,col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s): kc[i].metric(col[:20],f"{s.sum():,.1f}",f"avg {s.mean():,.1f}")
            st.markdown("---"); agg_rows=[]
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    grand = df_clean[chosen].select_dtypes("number").sum().sum()
                    agg_rows.append({"Column":col,"Count":int(s.count()),"Sum":s.sum(),"Mean":s.mean(),"Median":s.median(),"Min":s.min(),"Max":s.max(),"Std":s.std(),"% Total":f"{s.sum()/grand*100:.1f}%" if grand else "—"})
            if agg_rows:
                adf = pd.DataFrame(agg_rows).set_index("Column")
                st.dataframe(adf.style.format("{:,.2f}",na_rep="—",subset=[c for c in adf.columns if c!="% Total"]).background_gradient(cmap="YlOrRd",subset=["Sum","Max"]),use_container_width=True)
        st.markdown("---"); st.markdown('<div class="sec-title">🧮 Group-By</div>',unsafe_allow_html=True)
        all_cat = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique()<60]
        if all_cat and num_cols:
            gc1,gc2,gc3 = st.columns(3); gc=gc1.selectbox("Group by",all_cat); ac=gc2.selectbox("Aggregate",num_cols); af=gc3.selectbox("Function",["sum","mean","count","min","max","median"])
            grp = df_clean.groupby(gc)[ac].agg(af).reset_index().rename(columns={ac:f"{af}({ac})"}).sort_values(f"{af}({ac})",ascending=False)
            st.dataframe(grp,use_container_width=True)
            fig = px.bar(grp,x=gc,y=f"{af}({ac})",color=f"{af}({ac})",color_continuous_scale="Viridis",title=f"{af.title()} of {ac} by {gc}"); fig.update_layout(xaxis_tickangle=-35,height=400); st.plotly_chart(fig,use_container_width=True)
 
# TAB 3 – CHARTS
with tabs[3]:
    st.subheader("📈 Interactive Charts")
    ctype = st.selectbox("Chart Type",["Bar Chart","Grouped Bar","Line Chart","Scatter Plot","Area Chart","Bubble Chart","Heatmap (Correlation)","Box Plot","Funnel Chart","Waterfall / Cumulative","3-D Scatter"])
    if not num_cols: st.info("No numeric columns.")
    else:
        def _s(label,opts,idx=0,key=None): return st.selectbox(label,opts,index=min(idx,max(0,len(opts)-1)),key=key)
        if ctype=="Bar Chart":
            xc=_s("X",cat_cols or df_clean.columns.tolist(),key="bx"); yc=_s("Y",num_cols,key="by"); ori=st.radio("Orientation",["Vertical","Horizontal"],horizontal=True)
            d=df_clean[[xc,yc]].dropna(); fig=px.bar(d,x=xc if ori=="Vertical" else yc,y=yc if ori=="Vertical" else xc,color=yc,color_continuous_scale="Turbo",orientation="v" if ori=="Vertical" else "h",title=f"{yc} by {xc}"); fig.update_layout(height=480); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Grouped Bar":
            xc=_s("X",cat_cols or df_clean.columns.tolist(),key="gbx"); ycs=st.multiselect("Y",num_cols,default=num_cols[:3])
            if ycs: fig=px.bar(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs,barmode="group"); fig.update_layout(height=460); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Line Chart":
            xc=_s("X",df_clean.columns.tolist(),key="lx"); ycs=st.multiselect("Y",num_cols,default=num_cols[:2])
            if ycs: fig=px.line(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs,markers=True); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Scatter Plot":
            xc=_s("X",num_cols,0,"sc_x"); yc=_s("Y",num_cols,1,"sc_y"); sc=_s("Size",["None"]+num_cols,key="sc_s"); cc=_s("Color",["None"]+cat_cols+num_cols,key="sc_c")
            d=df_clean.dropna(subset=[xc,yc]); fig=px.scatter(d,x=xc,y=yc,size=sc if sc!="None" else None,color=cc if cc!="None" else None,color_continuous_scale="Rainbow"); fig.update_layout(height=480); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Area Chart":
            xc=_s("X",df_clean.columns.tolist(),key="ax"); ycs=st.multiselect("Y",num_cols,default=num_cols[:3])
            if ycs: fig=px.area(df_clean[[xc]+ycs].dropna(subset=ycs,how="all"),x=xc,y=ycs); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Bubble Chart":
            if len(num_cols)>=3:
                xc=_s("X",num_cols,0,"bu_x"); yc=_s("Y",num_cols,1,"bu_y"); sz=_s("Size",num_cols,2,"bu_s"); lc=_s("Color",["None"]+cat_cols,key="bu_c"); d=df_clean[[xc,yc,sz]].dropna()
                if lc!="None": d[lc]=df_clean[lc]
                fig=px.scatter(d,x=xc,y=yc,size=sz,color=lc if lc!="None" else None,size_max=65); fig.update_layout(height=500); st.plotly_chart(fig,use_container_width=True)
            else: st.info("Need >= 3 numeric columns.")
        elif ctype=="Heatmap (Correlation)":
            sel=st.multiselect("Columns",num_cols,default=num_cols[:12])
            if len(sel)>=2: fig=px.imshow(df_clean[sel].corr(),text_auto=".2f",color_continuous_scale="RdBu_r",aspect="auto"); fig.update_layout(height=540); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Box Plot":
            yc=_s("Value",num_cols,key="bp_v"); xc=_s("Group",["None"]+cat_cols,key="bp_g"); d=df_clean[[yc]+([xc] if xc!="None" else [])].dropna(subset=[yc])
            fig=px.box(d,y=yc,x=xc if xc!="None" else None,color=xc if xc!="None" else None,points="outliers"); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Funnel Chart":
            xc=_s("Stage",cat_cols or df_clean.columns.tolist(),key="fn_x"); yc=_s("Value",num_cols,key="fn_y")
            d=df_clean[[xc,yc]].dropna().groupby(xc)[yc].sum().reset_index().sort_values(yc,ascending=False); fig=px.funnel(d,x=yc,y=xc); fig.update_layout(height=450); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="Waterfall / Cumulative":
            yc=_s("Column",num_cols,key="wf_y"); d=df_clean[yc].dropna().reset_index(drop=True); cum=d.cumsum()
            fig=go.Figure(); fig.add_trace(go.Bar(name="Value",x=d.index,y=d,marker_color="#2a5298")); fig.add_trace(go.Scatter(name="Cumulative",x=cum.index,y=cum,line=dict(color="#f7971e",width=2.5),mode="lines+markers")); fig.update_layout(title=f"Cumulative: {yc}",height=450,barmode="group"); st.plotly_chart(fig,use_container_width=True)
        elif ctype=="3-D Scatter":
            if len(num_cols)>=3:
                xc=_s("X",num_cols,0,"3x"); yc=_s("Y",num_cols,1,"3y"); zc=_s("Z",num_cols,2,"3z"); cc=_s("Color",["None"]+cat_cols,key="3c"); d=df_clean[[xc,yc,zc]].dropna()
                if cc!="None": d[cc]=df_clean[cc]
                fig=px.scatter_3d(d,x=xc,y=yc,z=zc,color=cc if cc!="None" else None); fig.update_layout(height=550); st.plotly_chart(fig,use_container_width=True)
            else: st.info("Need >= 3 numeric columns.")
 
# TAB 4 – DISTRIBUTIONS
with tabs[4]:
    st.subheader("🥧 Distribution Charts")
    if not num_cols: st.info("No numeric columns.")
    else:
        r1,r2 = st.columns(2)
        with r1:
            st.markdown('<div class="sec-title">🍕 Pie / Donut</div>',unsafe_allow_html=True)
            pc_list=[c for c in df_clean.columns if c not in num_cols and 1<df_clean[c].nunique()<=30]
            if pc_list:
                pc=st.selectbox("Category",pc_list,key="pcat"); pv=st.selectbox("Value",num_cols,key="pval")
                pd_=df_clean[[pc,pv]].copy(); pd_[pv]=pd.to_numeric(pd_[pv],errors="coerce"); pd_=pd_.dropna().groupby(pc)[pv].sum().reset_index()
                fig=px.pie(pd_,names=pc,values=pv,hole=.38,color_discrete_sequence=px.colors.qualitative.Vivid); st.plotly_chart(fig,use_container_width=True)
        with r2:
            st.markdown('<div class="sec-title">📊 Histogram</div>',unsafe_allow_html=True)
            hc=st.selectbox("Column",num_cols,key="hcol"); bins=st.slider("Bins",5,100,25)
            fig=px.histogram(df_clean[hc].dropna(),nbins=bins,color_discrete_sequence=["#17a572"]); fig.update_layout(showlegend=False); st.plotly_chart(fig,use_container_width=True)
        st.markdown('<div class="sec-title">🎻 Violin</div>',unsafe_allow_html=True)
        vc=st.selectbox("Column",num_cols,key="vc"); fig=px.violin(df_clean[vc].dropna(),y=vc,box=True,points="outliers",color_discrete_sequence=["#c0392b"]); st.plotly_chart(fig,use_container_width=True)
 
# TAB 5 – QUERY ENGINE
with tabs[5]:
    st.subheader("🔍 Query Engine  (selected sheet only)")
    st.info("For **cross-file search** use the **💬 AI Smart Query** tab.")
    query = st.text_input("Question",placeholder="e.g. Total subscription / Max capacity / List customers")
    def run_query(q,df,nc):
        ql=q.lower(); res=[]
        if any(w in ql for w in ["sum","total"]): [res.append(f"**SUM `{c}`** = {df[c].sum():,.4f}") for c in nc if c.lower() in ql or "all" in ql or len(nc)==1]
        if any(w in ql for w in ["average","mean","avg"]): [res.append(f"**MEAN `{c}`** = {df[c].mean():,.4f}") for c in nc if c.lower() in ql or len(nc)==1]
        if any(w in ql for w in ["maximum","highest","max"]): [res.append(f"**MAX `{c}`** = {df[c].max():,.4f}") for c in nc if c.lower() in ql or len(nc)==1]
        if any(w in ql for w in ["minimum","lowest","min"]): [res.append(f"**MIN `{c}`** = {df[c].min():,.4f}") for c in nc if c.lower() in ql or len(nc)==1]
        if any(w in ql for w in ["count","how many"]): res.append(f"**Rows** = {len(df):,}")
        if any(w in ql for w in ["customer","list","show"]):
            for c in df.columns:
                if "customer" in c.lower() or "name" in c.lower():
                    nm=df[c].dropna().unique(); res.append(f"**`{c}`** ({len(nm)}):\n"+"\n".join(f"  • {n}" for n in nm[:30])); break
        if not res: res.append("ℹ️ Try: **sum / average / max / min / count / list**")
        return "\n\n".join(res)
    if query: st.markdown(run_query(query,df_clean,num_cols))
    st.markdown("---"); st.markdown('<div class="sec-title">🧮 Manual Compute</div>',unsafe_allow_html=True)
    if num_cols:
        mc1,mc2,mc3=st.columns(3); op=mc1.selectbox("Op",["Sum","Mean","Max","Min","Count","Median","Std Dev","Range"]); sc_col=mc2.selectbox("Column",num_cols,key="mc_col"); fc=mc3.selectbox("Filter by",["None"]+[c for c in df_clean.columns if c not in num_cols])
        fv=None
        if fc!="None": fv=st.selectbox("Filter value",df_clean[fc].dropna().unique().tolist())
        ds=df_clean.copy()
        if fc!="None" and fv is not None: ds=ds[ds[fc]==fv]
        s=ds[sc_col].dropna()
        ops={"Sum":s.sum(),"Mean":s.mean(),"Max":s.max(),"Min":s.min(),"Count":s.count(),"Median":s.median(),"Std Dev":s.std(),"Range":s.max()-s.min()}
        r=ops.get(op,"N/A")
        if isinstance(r,float): r=f"{r:,.4f}"
        st.success(f"**{op}** of `{sc_col}`{f' (where {fc}={fv})' if fv else ''} -> **{r}**")
 
# TAB 6 – MULTI-LOCATION
with tabs[6]:
    st.subheader("🌍 Cross-Location Comparison")
    @st.cache_data(show_spinner=False)
    def load_all_summ(files,folder):
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
                s=info["df"][comp_col].dropna(); rows.append({"Location|Sheet":lbl,"Sum":s.sum(),"Mean":s.mean(),"Max":s.max(),"Min":s.min(),"Count":s.count()})
        if rows:
            cmp=pd.DataFrame(rows).set_index("Location|Sheet"); st.dataframe(cmp.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),use_container_width=True)
            fig=px.bar(cmp.reset_index(),x="Location|Sheet",y="Sum",color="Sum",color_continuous_scale="Viridis",title=f"Sum of '{comp_col}'"); fig.update_layout(xaxis_tickangle=-30,height=440); st.plotly_chart(fig,use_container_width=True)
 
# TAB 7 – AI AGENT
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights")
    if st.button("🚀 Run Analysis",type="primary"):
        with st.spinner("Analysing…"):
            for lbl,info in list(all_summ.items())[:10]:
                dfa=info["df"]; nc=info["num_cols"]
                if not nc: continue
                with st.expander(f"📍 {lbl}",expanded=False):
                    ca,cb=st.columns(2)
                    with ca:
                        for col in nc[:5]:
                            s=dfa[col].dropna()
                            if len(s): st.metric(col[:26],f"{s.sum():,.1f}",f"avg {s.mean():,.1f}")
                    with cb:
                        for col in nc[:4]:
                            s=dfa[col].dropna()
                            if len(s)>3:
                                z=(s-s.mean())/s.std(); o=z[z.abs()>2.5]
                                (st.warning if len(o) else st.success)(f"`{col}`: {len(o)} outlier(s)" if len(o) else f"`{col}`: Clean ✓")
    st.markdown("---"); st.markdown('<div class="sec-title">📁 Files Summary</div>',unsafe_allow_html=True)
    fsm=[]
    for f in excel_files:
        shd = all_sheets if f==selected_file else load_file(os.path.join(data_dir,f))
        fsm.append({"File":loc_map[f],"Sheets":len(shd),"Rows":sum(len(s) for s in shd.values())})
    st.dataframe(pd.DataFrame(fsm),use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 8 – AI SMART QUERY  (SINGLE FILE + SINGLE SHEET from sidebar)
# No self-asking. Reads every cell from anywhere in the selected sheet.
# ═══════════════════════════════════════════════════════════════════════════
with tabs[8]:
    st.markdown("## 💬 AI Smart Query")
    st.markdown(
        f"Querying: **{loc_label}** › **{selected_sheet}**  \n"
        f"Change file/sheet in the **sidebar** to query a different one."
    )
 
    qi1, qi2, qi3 = st.columns(3)
    qi1.markdown(f'<div class="kcard kcard-blue"><h2>{sq_meta["total_cells"]:,}</h2><p>Cells in Sheet</p></div>', unsafe_allow_html=True)
    qi2.markdown(f'<div class="kcard kcard-green"><h2>{sq_meta["total_data"]:,}</h2><p>Data Cells</p></div>', unsafe_allow_html=True)
    qi3.markdown(f'<div class="kcard kcard-purple"><h2>{sq_meta["total_rows"]:,}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
 
    def _is_num(v):
        try: float(v); return True
        except: return False
 
    _OP = {"total","sum","avg","mean","max","min","count","list","find","show","all","average","maximum","minimum","highest","lowest","top","bottom","describe","statistics","stats","summary","unique","distinct","sheet","column","row","missing","null","percent","percentage","ratio","share","number","across","compare","location","locations","customer","customers","capacity","power","usage","rack","space","subscription","billing"}
    _SYN = {"subscription":["subscription","subscribed","subscript"],"capacity":["capacity","capac"],"power":["power","kw","kva"],"usage":["usage","utilization","consumption","consumed"],"rack":["rack","racks"],"space":["space","sqft","sq ft"],"customer":["customer","name"],"billing":["billing","bill"]}
 
    def _mcol(kw, hdr):
        hl = hdr.lower(); kwl = kw.lower()
        if kwl in hl: return True
        for key, syns in _SYN.items():
            if kwl in syns or kwl == key:
                for s in syns:
                    if s in hl: return True
        return False
 
    def sheet_query(question):
        q = question.strip(); ql = q.lower()
        sig = [w for w in re.findall(r"[a-z0-9]{3,}", ql) if w not in _SW]
 
        f_sum = any(x in ql for x in ["total","sum","aggregate"])
        f_avg = any(x in ql for x in ["average","mean","avg"])
        f_max = any(x in ql for x in ["maximum","highest","largest","max"])
        f_min = any(x in ql for x in ["minimum","lowest","smallest","min"])
        f_cnt = any(x in ql for x in ["count","how many","number of"])
        f_stat = any(x in ql for x in ["statistics","stats","describe","summary"])
        f_uniq = any(x in ql for x in ["unique","distinct","different"])
        f_miss = any(x in ql for x in ["missing","null","blank","empty"])
        f_cols = any(x in ql for x in ["column","columns","field","header"])
        f_topn = re.search(r"\btop\s*(\d+)\b", ql)
        f_botn = re.search(r"\bbottom\s*(\d+)\b", ql)
        f_num = f_sum or f_avg or f_max or f_min or f_cnt or f_stat
 
        out = {"answer": "", "table": None, "chart_df": None, "chart_cfg": None, "cell_hits": [], "sub_tables": []}
        wc = sq_cells  # single sheet cells
        rr = sq_rows   # single sheet row records
 
        if not wc:
            out["answer"] = "❓ No data in this sheet."
            return out
 
        def npkw(kw):
            res = []
            for cell in wc:
                if cell["is_header"]: continue
                if _mcol(kw, cell["col_header"]):
                    try: res.append((float(cell["value"]), cell))
                    except: pass
            return res
 
        def npbest(sw):
            bk, bp = None, []
            for w in sw:
                if w in _OP: continue
                p = npkw(w)
                if len(p) > len(bp): bp = p; bk = w
            return bk, bp
 
        def build_rows_df(row_nums):
            recs = []
            for rn in row_nums:
                rec = rr.get(rn, {})
                if rec:
                    rd = {"Row #": rn + 1}
                    rd.update(rec)
                    recs.append(rd)
            return pd.DataFrame(recs) if recs else pd.DataFrame()
 
        # INTENT: Missing values
        if f_miss:
            dfc = to_numeric(smart_header(raw_df))
            mr = []
            for col in dfc.columns:
                mc = int(dfc[col].isna().sum())
                if mc > 0: mr.append({"Column": col, "Missing": mc, "Missing%": f"{mc/max(len(dfc),1)*100:.1f}%"})
            if mr:
                tbl = pd.DataFrame(mr).sort_values("Missing", ascending=False)
                out["answer"] = f"Found **{len(tbl)}** column(s) with missing values."
                out["table"] = tbl
            else:
                out["answer"] = "✅ No missing values found."
            return out
 
        # INTENT: Column listing
        if f_cols and not f_num:
            seen = set(); cr = []
            for cell in wc:
                if not cell["is_header"]: continue
                ch = cell["value"].strip()
                if ch in ("", "nan"): continue
                if ch not in seen: seen.add(ch); cr.append({"Column": ch, "At Row": cell["row"]+1, "At Col": cell["col"]+1})
            tbl = pd.DataFrame(cr) if cr else pd.DataFrame()
            out["answer"] = f"Found **{len(tbl)}** column header(s) in this sheet."
            out["table"] = tbl
            return out
 
        # INTENT: Numeric aggregation
        if f_num:
            kw, pairs = npbest(sig)
            if not pairs:
                for dkw in ["subscription","capacity","power","usage","rack","space","consumption","kw","kva"]:
                    if dkw in ql:
                        pairs = npkw(dkw)
                        if pairs: kw = dkw; break
            if pairs:
                vals = [v for v, _ in pairs]; sa = pd.Series(vals); parts = []
                if f_sum or f_stat: parts.append(f"**Total (Sum):** {sa.sum():,.4f}")
                if f_avg or f_stat: parts.append(f"**Average:** {sa.mean():,.4f}")
                if f_max or f_stat: parts.append(f"**Maximum:** {sa.max():,.4f}")
                if f_min or f_stat: parts.append(f"**Minimum:** {sa.min():,.4f}")
                if f_cnt or f_stat: parts.append(f"**Count:** {sa.count():,}")
                if f_stat: parts.append(f"**Median:** {sa.median():,.4f} | **Std Dev:** {sa.std():,.4f}")
                # Build per-row detail
                detail = [{"Row #": c["row"]+1, "Column": c["col_header"], "Value": v} for v, c in pairs]
                tbl = pd.DataFrame(detail).sort_values("Value", ascending=False)
                out["answer"] = f"Results for **'{kw}'** ({len(vals):,} values):\n\n" + "\n".join(parts)
                out["table"] = tbl
                if f_topn:
                    n = int(f_topn.group(1)); top = sorted(pairs, key=lambda x: x[0], reverse=True)[:n]
                    out["sub_tables"].append({"label": f"🏆 Top {n}", "df": pd.DataFrame([{"Row": c["row"]+1, "Column": c["col_header"], "Value": v} for v, c in top])})
                if f_botn:
                    n = int(f_botn.group(1)); bot = sorted(pairs, key=lambda x: x[0])[:n]
                    out["sub_tables"].append({"label": f"🔻 Bottom {n}", "df": pd.DataFrame([{"Row": c["row"]+1, "Column": c["col_header"], "Value": v} for v, c in bot])})
                return out
 
            # Fallback: all numeric cells
            anums = [(float(c["value"]), c) for c in wc if not c["is_header"] and _is_num(c["value"])]
            if anums:
                vals = [v for v, _ in anums]; sa = pd.Series(vals); parts = []
                if f_sum: parts.append(f"**Sum ALL numeric:** {sa.sum():,.4f}")
                if f_avg: parts.append(f"**Avg ALL numeric:** {sa.mean():,.4f}")
                if f_max: parts.append(f"**Max ALL numeric:** {sa.max():,.4f}")
                if f_min: parts.append(f"**Min ALL numeric:** {sa.min():,.4f}")
                if f_cnt: parts.append(f"**Count ALL numeric:** {sa.count():,}")
                out["answer"] = f"No column matched keywords. ALL numeric cells:\n\n" + "\n".join(parts)
                return out
 
        # INTENT: Unique values
        if f_uniq and sig:
            for w in sig:
                if w in _OP: continue
                uv = set(); sr = []
                for cell in wc:
                    if cell["is_header"]: continue
                    if _mcol(w, cell["col_header"]):
                        uv.add(cell["value"])
                        sr.append({"Column": cell["col_header"], "Value": cell["value"], "Row #": cell["row"]+1})
                if uv:
                    tbl = pd.DataFrame(sr).drop_duplicates(subset=["Value"]) if sr else pd.DataFrame()
                    out["answer"] = f"**{len(uv)}** unique value(s) for **'{w}'**."
                    out["table"] = tbl
                    return out
 
        # INTENT: Free-text entity/keyword search
        if sig:
            quoted = re.findall(r'"([^"]+)"', q)
            if quoted:
                terms = [quoted[0].lower()]
            else:
                terms = [w for w in sig if any(w in cell["value"].lower() for cell in wc)]
                if not terms: terms = sig
 
            hit_cells = [cell for cell in wc if not cell["is_header"] and any(t in cell["value"].lower() for t in terms)]
            hit_rows = sorted({c["row"] for c in hit_cells})
            full_df = build_rows_df(hit_rows)
 
            cell_list = [{"Row #": c["row"]+1, "Col #": c["col"]+1, "Column Header": c["col_header"], "Value": c["value"]} for c in hit_cells[:80]]
 
            out["answer"] = (
                f"Found **{len(hit_cells):,}** cell(s) matching **'{', '.join(terms[:4])}'**\n"
                f"in **{len(hit_rows):,}** row(s)."
            )
            out["table"] = full_df if not full_df.empty else None
            out["cell_hits"] = cell_list
 
            # Extract customer names if present
            if not full_df.empty:
                for col in full_df.columns:
                    if "customer" in col.lower() or "name" in col.lower():
                        cdf = full_df[["Row #", col]].drop_duplicates()
                        out["sub_tables"].append({"label": f"👤 Customers ({len(cdf)})", "df": cdf})
                        break
            return out
 
        out["answer"] = (
            "❓ No match found.\n\n**Try:**\n"
            "• *List all customers*  •  *Find CISCO*  •  *Find AT&T*\n"
            "• *Total subscription*  •  *Max capacity*  •  *Top 10 subscription*\n"
            "• *Unique billing models*  •  *Show missing values*  •  *Show columns*"
        )
        return out
 
    # ── Render answer ───────────────────────────────────────────────────
    def render_answer(res, tidx=0):
        st.markdown(f'<div class="ans-box">{res["answer"]}</div>', unsafe_allow_html=True)
        if res.get("table") is not None and not res["table"].empty:
            tbl = res["table"].reset_index(drop=True)
            st.dataframe(tbl, use_container_width=True, height=min(520, 48+len(tbl)*36), key=f"tbl_{tidx}")
            st.download_button("⬇️ CSV", tbl.to_csv(index=False).encode(), "result.csv", "text/csv", key=f"dl_{tidx}")
        if res.get("chart_cfg") and res.get("chart_df") is not None:
            cfg = res["chart_cfg"]; cdf = res["chart_df"]
            if cfg["x"] in cdf.columns and cfg["y"] in cdf.columns:
                fig = px.bar(cdf.sort_values(cfg["y"], ascending=False).head(30), x=cfg["x"], y=cfg["y"], color=cfg["y"], color_continuous_scale="Viridis", title=cfg["title"], height=400)
                fig.update_layout(xaxis_tickangle=-30)
                st.plotly_chart(fig, use_container_width=True, key=f"ch_{tidx}")
        for si, s in enumerate(res.get("sub_tables", [])):
            with st.expander(s["label"], expanded=True):
                st.dataframe(s["df"], use_container_width=True, key=f"sub_{tidx}_{si}")
        if res.get("cell_hits"):
            with st.expander(f"🔬 Cell matches ({len(res['cell_hits'])})", expanded=False):
                for ch in res["cell_hits"]:
                    st.markdown(
                        f'<div class="cell-chip">R{ch["Row #"]} C{ch["Col #"]} | '
                        f'<i>{ch["Column Header"]}</i> | <b>{ch["Value"]}</b></div>',
                        unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)
 
    # ── Chat history (keyed to file+sheet so it resets on change) ─────
    hist_key = f"sq_hist_{selected_file}_{selected_sheet}"
    if hist_key not in st.session_state:
        st.session_state[hist_key] = []
 
    for tidx, turn in enumerate(st.session_state[hist_key]):
        st.markdown(f'<div class="q-user">🧑 {turn["q"]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)
        render_answer(turn["res"], tidx)
        st.markdown("---")
 
    # ── Input bar ─────────────────────────────────────────────────────
    st.markdown("---")
    ic, bc, cc = st.columns([8, 1, 1])
    with ic:
        user_q = st.text_input("Ask:", placeholder="Find CISCO | List customers | Total subscription | Max capacity | Top 10 subscription",
                               label_visibility="collapsed", key="sq_input")
    with bc:
        ask_btn = st.button("🔍 Ask", use_container_width=True, type="primary")
    with cc:
        if st.button("🗑️ Clear", use_container_width=True):
            st.session_state[hist_key] = []
            st.rerun()
 
    # ── Example chips ─────────────────────────────────────────────────
    st.markdown("**💡 Examples:**")
    examples = [
        ["List all customers", "Find CISCO", "Find AT&T", "Find Axis Bank", "Find MOTMOT", "Find Zscaler"],
        ["Total subscription", "Max capacity", "Average power usage", "Top 10 subscription", "Statistics of subscription", "Count rows"],
        ["Show columns", "Show missing values", "Unique billing models", "Unique ownership"],
    ]
    for row in examples:
        cols = st.columns(len(row))
        for j, ex in enumerate(row):
            if cols[j].button(ex, key=f"sqchip_{ex}", use_container_width=True):
                user_q = ex
                ask_btn = True
 
    if ask_btn and user_q.strip():
        with st.spinner(f"Scanning sheet for: **{user_q}** …"):
            answer = sheet_query(user_q)
        st.session_state[hist_key].append({"q": user_q, "res": answer})
        st.rerun()
 
# ─────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(f"Sify DC · Capacity Tracker · {meta['total_cells']:,} cells indexed · {', '.join(meta['locations'])}")
