import os, re, warnings, tempfile
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
 
warnings.filterwarnings("ignore")
 
# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Customer & Capacity Tracker",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)
 
# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#0f2027,#203a43,#2c5364);}
[data-testid="stSidebar"] *{color:#e0f7fa !important;}
.metric-card{background:linear-gradient(135deg,#1e3c72,#2a5298);border-radius:14px;
    padding:18px 22px;color:#fff;margin-bottom:10px;box-shadow:0 4px 18px rgba(0,0,0,.25);}
.metric-card h2{font-size:2rem;margin:0;}
.metric-card p{margin:4px 0 0 0;font-size:.9rem;opacity:.85;}
.section-title{font-size:1.3rem;font-weight:700;color:#1e3c72;
    border-left:5px solid #2a5298;padding-left:10px;margin:18px 0 10px 0;}
.bubble-user{background:linear-gradient(135deg,#1e3c72,#2a5298);color:#fff;
    border-radius:18px 18px 4px 18px;padding:11px 16px;margin:8px 0 4px auto;
    max-width:74%;width:fit-content;font-size:.97rem;
    box-shadow:0 3px 10px rgba(30,60,114,.35);margin-left:26%;}
.bubble-bot{background:linear-gradient(135deg,#0f5460,#17a2b8);color:#e0f7fa;
    border-radius:18px 18px 18px 4px;padding:11px 16px;margin:4px auto 8px 0;
    max-width:92%;font-size:.95rem;box-shadow:0 3px 10px rgba(23,162,184,.25);}
.cell-hit{background:#fff8e1;border-left:4px solid #f7971e;border-radius:6px;
    padding:7px 12px;margin:4px 0;font-family:monospace;font-size:.83rem;}
.answer-box{background:linear-gradient(135deg,#1a3a5c,#1e5f74);color:#e0f7fa;
    border-radius:12px;padding:14px 18px;margin:6px 0;font-size:1rem;
    box-shadow:0 2px 12px rgba(0,0,0,.3);white-space:pre-wrap;}
</style>
""", unsafe_allow_html=True)
 
# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "excel_files")
 
 
def find_excel_files(folder):
    if not os.path.isdir(folder):
        return []
    return sorted(f for f in os.listdir(folder) if f.lower().endswith((".xlsx", ".xls")))
 
 
def location_from_name(fname):
    name = (fname.replace("Customer_and_Capacity_Tracker_", "")
                 .replace(".xlsx", "").replace(".xls", ""))
    name = re.sub(r"_\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{2,4}.*$",
                  "", name, flags=re.I)
    return name.replace("_", " ").strip()
 
 
# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_file(path):
    sheets = {}
    try:
        ext    = os.path.splitext(path)[1].lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        xf     = pd.ExcelFile(path, engine=engine)
        for sh in xf.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sh, header=None,
                                   engine=engine, dtype=str)
                sheets[sh] = df
            except Exception:
                pass
    except Exception as e:
        st.sidebar.warning(f"⚠️ Cannot open {os.path.basename(path)}: {e}")
    return sheets
 
 
def best_header_row(df):
    best_row, best_score = 0, -1
    for i in range(min(6, len(df))):
        score = (df.iloc[i].astype(str).str.strip().str.len() > 0).sum()
        if score > best_score:
            best_score, best_row = score, i
    return best_row
 
 
def smart_header(df):
    hr     = best_header_row(df)
    header = df.iloc[hr].fillna("").astype(str).str.strip()
    seen   = {}
    unique_header = []
    for col in header:
        if col in seen:
            seen[col] += 1
            unique_header.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            unique_header.append(col)
    data = df.iloc[hr + 1:].copy()
    data.columns = unique_header
    return data.dropna(how="all").reset_index(drop=True)
 
 
def to_numeric_cols(df):
    out = df.copy()
    for col in out.columns:
        out[col] = pd.to_numeric(out[col], errors="ignore")
    return out
 
 
@st.cache_data(show_spinner=False)
def save_uploads(files):
    tmp = tempfile.mkdtemp()
    for f in files:
        with open(os.path.join(tmp, f.name), "wb") as fh:
            fh.write(f.read())
    return tmp
 
 
# ─────────────────────────────────────────────────────────────────────────────
# BUILD FULL CORPUS  – every non-empty cell of every sheet of every file
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=True)
def build_corpus(file_list, folder):
    """
    Returns
    -------
    corpus      : list[dict]  – one entry per non-empty cell
    row_records : dict        – (fname, loc, sheet, row_idx) -> {col_header: value}
    meta        : dict        – stats
    """
    corpus      = []
    row_records = defaultdict(dict)
 
    for fname in file_list:
        path   = os.path.join(folder, fname)
        loc    = location_from_name(fname)
        ext    = os.path.splitext(fname)[1].lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
 
        try:
            xf = pd.ExcelFile(path, engine=engine)
        except Exception:
            continue
 
        for sh in xf.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sh, header=None,
                                   engine=engine, dtype=str)
            except Exception:
                continue
 
            # Detect best header row -> build column-name map
            hr = best_header_row(df)
            col_names = {}
            for c in range(df.shape[1]):
                v = str(df.iat[hr, c]).strip()
                if v not in ("nan", "None", ""):
                    col_names[c] = v
 
            # Scan every single cell
            for r in range(df.shape[0]):
                for c in range(df.shape[1]):
                    v = str(df.iat[r, c]).strip()
                    if v in ("nan", "None", ""):
                        continue
                    ch  = col_names.get(c, f"Col_{c}")
                    key = (fname, loc, sh, r)
                    corpus.append({
                        "file":       fname,
                        "location":   loc,
                        "sheet":      sh,
                        "row":        r,
                        "col":        c,
                        "col_header": ch,
                        "value":      v,
                        "is_header":  (r == hr),
                    })
                    if r != hr:
                        row_records[key][ch] = v
 
    meta = {
        "total_cells":  len(corpus),
        "total_files":  len(set(x["file"]  for x in corpus)),
        "total_sheets": len(set((x["file"], x["sheet"]) for x in corpus)),
        "total_rows":   len(row_records),
        "locations":    sorted(set(x["location"] for x in corpus)),
    }
    return corpus, dict(row_records), meta
 
 
# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/fluency/96/data-center.png", width=70)
st.sidebar.title("📊 Capacity Tracker")
st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Source")
 
uploaded = st.sidebar.file_uploader(
    "Upload Excel files (optional – overrides folder)",
    type=["xlsx", "xls"], accept_multiple_files=True
)
 
data_dir    = save_uploads(tuple(uploaded)) if uploaded else UPLOAD_DIR
excel_files = find_excel_files(data_dir)
 
if not excel_files:
    st.warning(
        "⚠️ No Excel files found.\n\n"
        "**Option 1:** Upload via the sidebar uploader.\n\n"
        "**Option 2:** Place Excel files in `excel_files/` next to `app.py`."
    )
    st.stop()
 
loc_map = {f: location_from_name(f) for f in excel_files}
 
st.sidebar.subheader("🏙️ Select Location")
selected_file  = st.sidebar.selectbox("Location", excel_files, format_func=lambda x: loc_map[x])
all_sheets     = load_file(os.path.join(data_dir, selected_file))
st.sidebar.subheader("📋 Select Sheet")
selected_sheet = st.sidebar.selectbox("Sheet", list(all_sheets.keys()))
 
raw_df   = all_sheets[selected_sheet]
df_clean = to_numeric_cols(smart_header(raw_df))
num_cols = df_clean.select_dtypes(include="number").columns.tolist()
 
st.sidebar.markdown("---")
st.sidebar.subheader("🔢 Numeric Columns")
st.sidebar.caption(f"{len(num_cols)} numeric columns detected")
 
# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "🏠 Overview", "📋 Raw Data", "📊 Analytics", "📈 Charts",
    "🥧 Distributions", "🔍 Query Engine", "🌍 Multi-Location",
    "🤖 AI Agent", "💬 AI Smart Query"
])
 
location_label = loc_map[selected_file]
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 1 – OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.title(f"🏢 {location_label} – {selected_sheet}")
    st.caption(f"File: `{selected_file}` | Raw shape: {raw_df.shape[0]} rows × {raw_df.shape[1]} cols")
 
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="metric-card"><h2>{len(df_clean)}</h2><p>Data Rows</p></div>',
                unsafe_allow_html=True)
    c2.markdown(f'<div class="metric-card" style="background:linear-gradient(135deg,#11998e,#38ef7d);color:#003">'
                f'<h2>{len(df_clean.columns)}</h2><p>Columns</p></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="metric-card" style="background:linear-gradient(135deg,#c94b4b,#4b134f);">'
                f'<h2>{len(num_cols)}</h2><p>Numeric Columns</p></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="metric-card" style="background:linear-gradient(135deg,#f7971e,#ffd200);color:#000;">'
                f'<h2>{len(excel_files)}</h2><p>Files Loaded</p></div>', unsafe_allow_html=True)
 
    st.markdown("---")
    if num_cols:
        st.markdown('<div class="section-title">📐 Quick Statistics</div>', unsafe_allow_html=True)
        stats = df_clean[num_cols].describe().T
        stats["range"] = stats["max"] - stats["min"]
        st.dataframe(
            stats.style.format("{:.2f}", na_rep="–")
                       .background_gradient(cmap="Blues", subset=["mean", "max"]),
            use_container_width=True)
 
    st.markdown('<div class="section-title">🗂️ Column Overview</div>', unsafe_allow_html=True)
    col_info = pd.DataFrame({
        "Column":   df_clean.columns,
        "Dtype":    df_clean.dtypes.values,
        "Non-Null": df_clean.notna().sum().values,
        "Null %":   (df_clean.isna().mean() * 100).round(1).values,
        "Sample":   [str(df_clean[c].dropna().iloc[0])
                     if df_clean[c].dropna().shape[0] > 0 else "—"
                     for c in df_clean.columns]
    })
    st.dataframe(col_info, use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 2 – RAW DATA
# ═══════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("📋 Data Table")
    search = st.text_input("🔍 Search any cell value", "")
    disp   = df_clean[
        df_clean.apply(lambda col: col.astype(str)
                                      .str.contains(search, case=False, na=False), axis=0)
               .any(axis=1)
    ] if search else df_clean
    if search:
        st.caption(f"Showing {len(disp)} matching rows")
    st.dataframe(disp, use_container_width=True, height=520)
    st.download_button("⬇️ Download CSV", disp.to_csv(index=False).encode(),
                       "data_export.csv", "text/csv")
    st.markdown("---")
    st.subheader("🗃️ Raw Excel View (no header applied)")
    st.dataframe(raw_df, use_container_width=True, height=300)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 3 – ANALYTICS
# ═══════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("📊 Column-Level Analytics")
    if not num_cols:
        st.info("No numeric columns detected.")
    else:
        chosen = st.multiselect("Select columns to analyse", num_cols, default=num_cols[:6])
        if chosen:
            sub      = df_clean[chosen].dropna(how="all")
            cols_row = st.columns(len(chosen[:6]))
            for i, col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s):
                    cols_row[i % len(cols_row)].metric(col[:25], f"{s.sum():,.2f}", f"μ={s.mean():,.2f}")
 
            st.markdown("---")
            agg_rows = []
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    agg_rows.append({
                        "Column": col, "Count": int(s.count()),
                        "Sum": s.sum(), "Mean": s.mean(), "Median": s.median(),
                        "Min": s.min(), "Max": s.max(), "Std Dev": s.std(),
                        "Variance": s.var(), "25%": s.quantile(.25), "75%": s.quantile(.75),
                        "% Non-Null": f"{s.count()/max(len(df_clean),1)*100:.1f}%",
                        "Sum %": f"{s.sum()/max(df_clean[chosen].sum().sum(),1)*100:.1f}%"
                    })
            if agg_rows:
                agg_df = pd.DataFrame(agg_rows).set_index("Column")
                st.dataframe(
                    agg_df.style.format("{:,.2f}", na_rep="–",
                                        subset=[c for c in agg_df.columns
                                                if c not in ["% Non-Null", "Sum %"]])
                               .background_gradient(cmap="YlOrRd", subset=["Sum", "Max"]),
                    use_container_width=True)
 
        st.markdown("---")
        st.markdown('<div class="section-title">🧮 Group-By Aggregation</div>', unsafe_allow_html=True)
        cat_cols = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique() < 50]
        if cat_cols and num_cols:
            g_col  = st.selectbox("Group by", cat_cols)
            ag_col = st.selectbox("Aggregate column", num_cols)
            ag_fn  = st.radio("Function", ["sum","mean","count","min","max","median"], horizontal=True)
            grp    = df_clean.groupby(g_col)[ag_col].agg(ag_fn).reset_index()
            grp.columns = [g_col, f"{ag_fn}({ag_col})"]
            grp    = grp.sort_values(grp.columns[1], ascending=False)
            st.dataframe(grp, use_container_width=True)
            fig = px.bar(grp, x=g_col, y=grp.columns[1], color=grp.columns[1],
                         color_continuous_scale="Viridis",
                         title=f"{ag_fn.title()} of {ag_col} by {g_col}")
            fig.update_layout(xaxis_tickangle=-35, height=420)
            st.plotly_chart(fig, use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 4 – CHARTS
# ═══════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("📈 Interactive Charts")
    chart_type = st.selectbox("Chart Type", [
        "Bar Chart", "Line Chart", "Scatter Plot", "Area Chart",
        "Bubble Chart", "Heatmap (Correlation)", "Box Plot",
        "Funnel Chart", "Waterfall / Cumulative"
    ])
    if not num_cols:
        st.info("No numeric columns available.")
    else:
        cat_cols2 = [c for c in df_clean.columns if c not in num_cols]
 
        if chart_type == "Bar Chart":
            xc  = st.selectbox("X axis", cat_cols2 or df_clean.columns[:1].tolist())
            yc  = st.selectbox("Y axis", num_cols)
            oc  = st.selectbox("Color by", ["None"] + cat_cols2)
            ori = st.radio("Orientation", ["Vertical", "Horizontal"], horizontal=True)
            d   = df_clean[[xc, yc]].dropna()
            fig = px.bar(d, x=xc if ori=="Vertical" else yc,
                         y=yc if ori=="Vertical" else xc,
                         color=oc if oc != "None" else None,
                         orientation="v" if ori=="Vertical" else "h",
                         color_continuous_scale="Turbo", title=f"{yc} by {xc}")
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
 
        elif chart_type == "Line Chart":
            xc  = st.selectbox("X axis", df_clean.columns.tolist())
            ycs = st.multiselect("Y axis", num_cols, default=num_cols[:2])
            if ycs:
                d   = df_clean[[xc] + ycs].dropna(subset=ycs, how="all")
                fig = px.line(d, x=xc, y=ycs, markers=True, title=f"Line: {', '.join(ycs)}")
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
 
        elif chart_type == "Scatter Plot":
            xc  = st.selectbox("X", num_cols)
            yc  = st.selectbox("Y", num_cols, index=min(1, len(num_cols)-1))
            sc  = st.selectbox("Size", ["None"] + num_cols)
            cc  = st.selectbox("Color", ["None"] + cat_cols2 + num_cols)
            d   = df_clean.dropna(subset=[xc, yc])
            fig = px.scatter(d, x=xc, y=yc,
                             size=sc if sc != "None" else None,
                             color=cc if cc != "None" else None,
                             hover_data=d.columns[:5].tolist(),
                             color_continuous_scale="Rainbow", title=f"{yc} vs {xc}")
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
 
        elif chart_type == "Area Chart":
            xc  = st.selectbox("X", df_clean.columns.tolist())
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:3])
            if ycs:
                d   = df_clean[[xc] + ycs].dropna(subset=ycs, how="all")
                fig = px.area(d, x=xc, y=ycs, title="Area Chart")
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
 
        elif chart_type == "Bubble Chart":
            if len(num_cols) >= 3:
                xc  = st.selectbox("X", num_cols, index=0)
                yc  = st.selectbox("Y", num_cols, index=1)
                sc  = st.selectbox("Bubble Size", num_cols, index=2)
                lc  = st.selectbox("Color", ["None"] + cat_cols2)
                d   = df_clean[[xc, yc, sc]].dropna()
                if lc != "None":
                    d[lc] = df_clean[lc]
                fig = px.scatter(d, x=xc, y=yc, size=sc,
                                 color=lc if lc != "None" else None,
                                 size_max=60,
                                 color_discrete_sequence=px.colors.qualitative.Vivid,
                                 title="Bubble Chart")
                fig.update_layout(height=520)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need ≥ 3 numeric columns.")
 
        elif chart_type == "Heatmap (Correlation)":
            sel = st.multiselect("Columns", num_cols, default=num_cols[:10])
            if len(sel) >= 2:
                fig = px.imshow(df_clean[sel].corr(), text_auto=".2f",
                                color_continuous_scale="RdBu_r", aspect="auto",
                                title="Correlation Heatmap")
                fig.update_layout(height=520)
                st.plotly_chart(fig, use_container_width=True)
 
        elif chart_type == "Box Plot":
            yc  = st.selectbox("Value column", num_cols)
            xc  = st.selectbox("Group by", ["None"] + cat_cols2)
            d   = df_clean[[yc] + ([xc] if xc != "None" else [])].dropna(subset=[yc])
            fig = px.box(d, y=yc, x=xc if xc != "None" else None,
                         color=xc if xc != "None" else None,
                         color_discrete_sequence=px.colors.qualitative.Pastel,
                         points="outliers", title=f"Box: {yc}")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)
 
        elif chart_type == "Funnel Chart":
            xc  = st.selectbox("Stage", cat_cols2 or df_clean.columns[:1].tolist())
            yc  = st.selectbox("Value", num_cols)
            d   = df_clean[[xc, yc]].dropna().groupby(xc)[yc].sum().reset_index()
            d   = d.sort_values(yc, ascending=False)
            fig = px.funnel(d, x=yc, y=xc, title="Funnel Chart")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)
 
        elif chart_type == "Waterfall / Cumulative":
            yc  = st.selectbox("Numeric column", num_cols)
            d   = df_clean[yc].dropna().reset_index(drop=True)
            cum = d.cumsum()
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Value", x=d.index, y=d, marker_color="#2a5298"))
            fig.add_trace(go.Scatter(name="Cumulative", x=cum.index, y=cum,
                                     line=dict(color="#f7971e", width=2.5),
                                     mode="lines+markers"))
            fig.update_layout(title=f"Cumulative: {yc}", height=450, barmode="group")
            st.plotly_chart(fig, use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 5 – DISTRIBUTIONS
# ═══════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("🥧 Distribution Charts")
    if not num_cols:
        st.info("No numeric columns.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            cat_cols3 = [c for c in df_clean.columns
                         if c not in num_cols and df_clean[c].nunique() <= 30]
            if cat_cols3:
                cc  = st.selectbox("Pie – category", cat_cols3, key="pie_cat")
                vc  = st.selectbox("Pie – value", num_cols, key="pie_val")
                pd_ = df_clean[[cc, vc]].dropna()
                pd_[vc] = pd.to_numeric(pd_[vc], errors="coerce")
                pd_ = pd_.dropna().groupby(cc)[vc].sum().reset_index()
                fig = px.pie(pd_, names=cc, values=vc, hole=.35,
                             color_discrete_sequence=px.colors.qualitative.Vivid,
                             title=f"Distribution of {vc}")
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            hc   = st.selectbox("Histogram column", num_cols, key="hist_col")
            bins = st.slider("Bins", 5, 80, 20)
            fig  = px.histogram(df_clean[hc].dropna(), nbins=bins,
                                color_discrete_sequence=["#11998e"],
                                title=f"Histogram: {hc}")
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
 
        st.markdown("---")
        st.markdown('<div class="section-title">🗺️ Treemap</div>', unsafe_allow_html=True)
        cat_cols4 = [c for c in df_clean.columns
                     if c not in num_cols and df_clean[c].nunique() <= 40]
        if cat_cols4 and num_cols:
            tc  = st.selectbox("Treemap – category", cat_cols4)
            tv  = st.selectbox("Treemap – value", num_cols)
            td  = df_clean[[tc, tv]].dropna().groupby(tc)[tv].sum().reset_index()
            td  = td[td[tv] > 0]
            if len(td):
                fig = px.treemap(td, path=[tc], values=tv, color=tv,
                                 color_continuous_scale="Turbo",
                                 title=f"Treemap: {tv} by {tc}")
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
 
        st.markdown("---")
        st.markdown('<div class="section-title">🎻 Violin Plot</div>', unsafe_allow_html=True)
        vc2 = st.selectbox("Violin column", num_cols)
        fig = px.violin(df_clean[vc2].dropna(), y=vc2, box=True, points="outliers",
                        color_discrete_sequence=["#c94b4b"], title=f"Violin: {vc2}")
        st.plotly_chart(fig, use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 6 – QUERY ENGINE
# ═══════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("🔍 Query Engine")
    st.info("Ask questions in plain English. Maps to pandas operations on the selected sheet.")
    query = st.text_input("Your question",
                          placeholder="e.g. What is the total of Subscription?")
 
    def run_query(q_text, df, n_cols):
        q = q_text.lower()
        results = []
        if any(w in q for w in ["sum", "total"]):
            for c in n_cols:
                if c.lower() in q or "all" in q or len(n_cols) == 1:
                    results.append(f"**SUM `{c}`**: {df[c].sum():,.2f}")
        if any(w in q for w in ["mean", "average", "avg"]):
            for c in n_cols:
                if c.lower() in q or len(n_cols) == 1:
                    results.append(f"**MEAN `{c}`**: {df[c].mean():,.2f}")
        if any(w in q for w in ["max", "maximum", "highest"]):
            for c in n_cols:
                if c.lower() in q or len(n_cols) == 1:
                    results.append(f"**MAX `{c}`**: {df[c].max():,.2f}")
        if any(w in q for w in ["min", "minimum", "lowest"]):
            for c in n_cols:
                if c.lower() in q or len(n_cols) == 1:
                    results.append(f"**MIN `{c}`**: {df[c].min():,.2f}")
        if any(w in q for w in ["count", "how many"]):
            results.append(f"**Total rows**: {len(df)}")
        if "median" in q:
            for c in n_cols:
                if c.lower() in q or len(n_cols) == 1:
                    results.append(f"**MEDIAN `{c}`**: {df[c].median():,.2f}")
        if any(w in q for w in ["std", "deviation"]):
            for c in n_cols:
                if c.lower() in q or len(n_cols) == 1:
                    results.append(f"**STD `{c}`**: {df[c].std():,.2f}")
        if any(w in q for w in ["describe", "statistics", "summary"]):
            results.append("**Stats:**\n" + df[n_cols].describe().to_markdown())
        if any(w in q for w in ["missing", "null", "nan"]):
            ni = df.isna().sum()
            ni = ni[ni > 0]
            results.append("**Missing:**\n" + ni.to_string())
        if "unique" in q:
            for c in df.columns:
                if c.lower() in q:
                    u = df[c].dropna().unique()
                    results.append(f"**Unique `{c}`** ({len(u)}): {', '.join(map(str, u[:20]))}")
        if not results:
            results.append("ℹ️ Try: sum, average, max, min, count, median, describe, unique, missing")
        return "\n\n".join(results)
 
    if query:
        st.markdown(run_query(query, df_clean, num_cols))
 
    st.markdown("---")
    st.markdown('<div class="section-title">🧮 Manual Compute</div>', unsafe_allow_html=True)
    if num_cols:
        mc1, mc2, mc3 = st.columns(3)
        op   = mc1.selectbox("Operation", ["Sum","Mean","Max","Min","Count","Median",
                                            "Std Dev","Variance","% of total","Cumulative Sum"])
        sc2  = mc2.selectbox("Column", num_cols)
        fc   = mc3.selectbox("Filter by", ["None"] + [c for c in df_clean.columns
                                                       if c not in num_cols])
        fv   = None
        if fc != "None":
            fv = st.selectbox("Filter value", df_clean[fc].dropna().unique().tolist())
        ds   = df_clean.copy()
        if fc != "None" and fv is not None:
            ds = ds[ds[fc] == fv]
        s    = ds[sc2].dropna()
        ops  = {
            "Sum": s.sum(), "Mean": s.mean(), "Max": s.max(), "Min": s.min(),
            "Count": s.count(), "Median": s.median(), "Std Dev": s.std(),
            "Variance": s.var(),
            "% of total": f"{s.sum()/max(df_clean[sc2].sum(),1)*100:.2f}%",
            "Cumulative Sum": s.cumsum().iloc[-1] if len(s) else 0
        }
        res = ops.get(op, "N/A")
        if isinstance(res, float):
            res = f"{res:,.4f}"
        st.success(f"**{op} of `{sc2}`** "
                   f"{f'(where {fc}={fv})' if fv else ''}: **{res}**")
        if op == "Cumulative Sum":
            st.plotly_chart(
                px.line(s.cumsum().reset_index(drop=True),
                        title=f"Cumulative Sum: {sc2}", markers=True),
                use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 7 – MULTI-LOCATION
# ═══════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("🌍 Cross-Location Comparison")
 
    @st.cache_data(show_spinner=True)
    def load_all_summaries(files, folder):
        summaries = {}
        for f in files:
            sh_dict = load_file(os.path.join(folder, f))
            for sh_name, raw in sh_dict.items():
                df_c = to_numeric_cols(smart_header(raw))
                nc   = df_c.select_dtypes(include="number").columns.tolist()
                if nc:
                    summaries[f"{loc_map[f]} | {sh_name}"] = {
                        "df": df_c, "num_cols": nc, "file": f, "sheet": sh_name
                    }
        return summaries
 
    all_summaries = load_all_summaries(tuple(excel_files), data_dir)
 
    if all_summaries:
        comp_col = st.selectbox(
            "Compare by column",
            sorted(set(c for v in all_summaries.values() for c in v["num_cols"])))
        rows = []
        for label, info in all_summaries.items():
            if comp_col in info["num_cols"]:
                s = info["df"][comp_col].dropna()
                rows.append({"Location|Sheet": label, "Sum": s.sum(),
                             "Mean": s.mean(), "Max": s.max(),
                             "Min": s.min(), "Count": s.count()})
        if rows:
            cmp = pd.DataFrame(rows).set_index("Location|Sheet")
            st.dataframe(
                cmp.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),
                use_container_width=True)
            fig = px.bar(cmp.reset_index(), x="Location|Sheet", y="Sum",
                         color="Sum", color_continuous_scale="Viridis",
                         title=f"Sum of '{comp_col}' across locations")
            fig.update_layout(xaxis_tickangle=-30, height=450)
            st.plotly_chart(fig, use_container_width=True)
 
            fig2 = px.scatter(cmp.reset_index(), x="Mean", y="Max", size="Sum",
                              text="Location|Sheet", color="Count",
                              color_continuous_scale="Turbo", title="Bubble: Mean vs Max")
            fig2.update_traces(textposition="top center")
            fig2.update_layout(height=500)
            st.plotly_chart(fig2, use_container_width=True)
 
            st.markdown('<div class="section-title">🕸️ Radar Chart</div>', unsafe_allow_html=True)
            rm   = st.radio("Radar metric", ["Sum","Mean","Max"], horizontal=True)
            norm = cmp[rm] / cmp[rm].max()
            th   = norm.index.tolist()
            rv   = norm.values.tolist()
            fig3 = go.Figure(go.Scatterpolar(
                r=rv + [rv[0]], theta=th + [th[0]],
                fill="toself", fillcolor="rgba(42,82,152,.3)",
                line=dict(color="#2a5298")))
            fig3.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 1])),
                height=500, title=f"Radar: {rm} of '{comp_col}'")
            st.plotly_chart(fig3, use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 8 – AI AGENT  (automated insights, unchanged)
# ═══════════════════════════════════════════════════════════════════════════
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights")
    st.info("Scans all sheets and surfaces KPIs, anomalies and correlations automatically.")
 
    if st.button("🚀 Run AI Agent Analysis"):
        with st.spinner("Analysing…"):
            for label, info in list(all_summaries.items())[:8]:
                df_a = info["df"]
                nc   = info["num_cols"]
                if not nc:
                    continue
                with st.expander(f"📍 {label}", expanded=False):
                    ca, cb = st.columns(2)
                    with ca:
                        st.markdown("**📊 Key KPIs**")
                        for col in nc[:4]:
                            s = df_a[col].dropna()
                            if len(s):
                                st.metric(col[:30], f"{s.sum():,.1f}", f"Avg {s.mean():,.1f}")
                    with cb:
                        st.markdown("**⚠️ Anomaly Detection (Z-score)**")
                        for col in nc[:3]:
                            s = df_a[col].dropna()
                            if len(s) > 3:
                                z = (s - s.mean()) / s.std()
                                o = z[z.abs() > 2.5]
                                if len(o):
                                    st.warning(f"`{col}`: {len(o)} outlier(s)")
                                else:
                                    st.success(f"`{col}`: No outliers")
                    if len(nc) >= 2:
                        fig = px.bar(df_a[nc[:2]].dropna().reset_index(),
                                     x="index", y=nc[0],
                                     color_discrete_sequence=["#2a5298"],
                                     title=nc[0])
                        fig.update_layout(height=280, showlegend=False,
                                          margin=dict(t=35, b=0))
                        st.plotly_chart(fig, use_container_width=True)
                    if len(nc) >= 2:
                        cs = df_a[nc].corr().unstack().drop_duplicates()
                        cs = cs[cs.index.get_level_values(0) !=
                                cs.index.get_level_values(1)].abs().sort_values(ascending=False)
                        if len(cs):
                            p = cs.index[0]
                            st.info(f"Top correlation: **{p[0]}** ↔ **{p[1]}** ({cs.iloc[0]:.3f})")
 
    st.markdown("---")
    st.markdown('<div class="section-title">📁 All Files Summary</div>', unsafe_allow_html=True)
    fsm = []
    for f in excel_files:
        sh_dict = (all_sheets if f == selected_file
                   else load_file(os.path.join(data_dir, f)))
        fsm.append({
            "File": loc_map[f], "Sheets": len(sh_dict),
            "Total Rows":    sum(len(s) for s in sh_dict.values()),
            "Total Columns": sum(len(s.columns) for s in sh_dict.values())
        })
    fs_df = pd.DataFrame(fsm)
    st.dataframe(fs_df, use_container_width=True)
    fig = px.bar(fs_df, x="File", y="Total Rows", color="Sheets",
                 color_continuous_scale="Blues", title="Total rows per location")
    fig.update_layout(xaxis_tickangle=-30, height=380)
    st.plotly_chart(fig, use_container_width=True)
 
# ═══════════════════════════════════════════════════════════════════════════
# TAB 9 – 💬 AI SMART QUERY  ← NEW
# Scans EVERY cell of EVERY sheet of EVERY file.
# No Excel formulas used. No pre-conditions. Pure free-text → direct answer.
# ═══════════════════════════════════════════════════════════════════════════
with tabs[8]:
 
    # ── Build corpus (cached) ─────────────────────────────────────────────
    with st.spinner("🔍 Indexing every cell across all Excel files…"):
        corpus, row_records, meta = build_corpus(tuple(excel_files), data_dir)
 
    # ── Page header ───────────────────────────────────────────────────────
    st.markdown("## 💬 AI Smart Query — Ask Anything")
    st.markdown(
        "Type **any question** in plain English. "
        "The engine reads **every row · every column · every sheet** of every file "
        "and returns a direct answer — no formulas, no pre-conditions, no restrictions."
    )
 
    qi1, qi2, qi3, qi4 = st.columns(4)
    qi1.markdown(
        f'<div class="metric-card"><h2>{meta["total_cells"]:,}</h2>'
        f'<p>Cells Indexed</p></div>', unsafe_allow_html=True)
    qi2.markdown(
        f'<div class="metric-card" style="background:linear-gradient(135deg,#11998e,#38ef7d);color:#003">'
        f'<h2>{meta["total_files"]}</h2><p>Files</p></div>', unsafe_allow_html=True)
    qi3.markdown(
        f'<div class="metric-card" style="background:linear-gradient(135deg,#c94b4b,#4b134f);">'
        f'<h2>{meta["total_sheets"]}</h2><p>Sheets</p></div>', unsafe_allow_html=True)
    qi4.markdown(
        f'<div class="metric-card" style="background:linear-gradient(135deg,#f7971e,#ffd200);color:#000;">'
        f'<h2>{meta["total_rows"]:,}</h2><p>Data Rows</p></div>', unsafe_allow_html=True)
 
    # ── STOPWORDS ─────────────────────────────────────────────────────────
    _SW = {
        "the","and","for","are","all","any","how","what","show","give","tell","from",
        "this","that","with","get","find","list","much","many","each","every","data",
        "value","values","number","numbers","in","of","a","an","is","at","by","to",
        "do","me","my","about","details","info","please","can","you","per","across",
        "which","where","who","when","does","did","have","has","their","its","our",
        "your","there","these","those","been","will","would","could","should","shall",
        "let","some","just","also","even","only","into","over","under","both","such",
        "than","then","but","not","nor","yet","so","either","neither","give","tell",
    }
 
    # ── CORE AI QUERY ENGINE ───────────────────────────────────────────────
    def ai_smart_query(question: str) -> dict:
        """
        Pure Python / Pandas AI agent.
        Uses the pre-built corpus (every non-empty cell from every sheet).
        Zero Excel formulas. Zero hard-coded column conditions.
        Returns: {answer, table, chart, chart_df, cell_hits, sub_tables}
        """
        q     = question.strip()
        q_low = q.lower()
 
        # significant keywords
        sig   = [w for w in re.findall(r"[a-z0-9]{3,}", q_low) if w not in _SW]
 
        # intent flags
        w_sum    = any(x in q_low for x in ["total","sum","aggregate"])
        w_avg    = any(x in q_low for x in ["average","mean","avg"])
        w_max    = any(x in q_low for x in ["maximum","highest","largest","biggest","max"])
        w_min    = any(x in q_low for x in ["minimum","lowest","smallest","least","min"])
        w_cnt    = any(x in q_low for x in ["count","how many","number of"])
        w_stat   = any(x in q_low for x in ["statistics","stats","describe","summary","overview"])
        w_uniq   = any(x in q_low for x in ["unique","distinct","different"])
        w_miss   = any(x in q_low for x in ["missing","null","blank","empty","nan"])
        w_sheets = any(x in q_low for x in ["sheet","sheets","tab","tabs"])
        w_cols   = any(x in q_low for x in ["column","columns","field","fields","header"])
        w_rows   = any(x in q_low for x in ["row","rows","record","records","entry"])
        w_num_op = w_sum or w_avg or w_max or w_min or w_cnt or w_stat
        w_topn   = re.search(r"\btop\s*(\d+)\b", q_low)
 
        out = {
            "answer":     "",
            "table":      None,
            "chart":      None,
            "chart_df":   None,
            "cell_hits":  [],
            "sub_tables": [],
        }
 
        if not corpus:
            out["answer"] = "⚠️ No data indexed yet."
            return out
 
        # ── location filter extracted from question ───────────────────────
        loc_filter = None
        for loc in meta["locations"]:
            for part in re.findall(r"[a-z]+", loc.lower()):
                if len(part) >= 4 and part in q_low:
                    loc_filter = loc
                    break
            if loc_filter:
                break
 
        # ── helpers ───────────────────────────────────────────────────────
        def loc_ok(cell_or_loc):
            if not loc_filter:
                return True
            loc_str = cell_or_loc if isinstance(cell_or_loc, str) else cell_or_loc["location"]
            return loc_filter.lower() in loc_str.lower()
 
        def find_best_col_keyword(sig_list):
            """Return the sig word that matches the most column headers in corpus."""
            op_excl = {"total","sum","avg","mean","max","min","count","list","find",
                       "show","all","average","maximum","minimum","highest","lowest",
                       "top","bottom","describe","statistics","stats","summary",
                       "unique","distinct","sheet","column","row","missing","null"}
            candidates = [w for w in sig_list if w not in op_excl and len(w) >= 3]
            best, best_n = None, 0
            for w in candidates:
                n = sum(1 for c in corpus if w in c["col_header"].lower())
                if n > best_n:
                    best_n, best = n, w
            return best
 
        def numeric_for_col_kw(col_kw):
            """All (float_value, cell_dict) for cells whose col_header contains col_kw."""
            res = []
            for cell in corpus:
                if cell["is_header"] or not loc_ok(cell):
                    continue
                if col_kw.lower() in cell["col_header"].lower():
                    try:
                        res.append((float(cell["value"]), cell))
                    except ValueError:
                        pass
            return res
 
        def build_rows_df(row_keys):
            """Convert {(fname,loc,sheet,row)} → tidy DataFrame."""
            records = []
            for key in row_keys:
                rec = row_records.get(key, {})
                if rec:
                    row_dict = {
                        "📍 Location": key[1],
                        "📋 Sheet":    key[2],
                        "Row #":       key[3] + 1,
                    }
                    row_dict.update(rec)
                    records.append(row_dict)
            return pd.DataFrame(records) if records else pd.DataFrame()
 
        # ════════════════════════════════════════════════════════════════
        # INTENT ① – Sheet listing
        # ════════════════════════════════════════════════════════════════
        if w_sheets and not w_num_op:
            seen_keys = set()
            sheet_rows = []
            for cell in corpus:
                if not loc_ok(cell):
                    continue
                k = (cell["location"], cell["sheet"])
                if k not in seen_keys:
                    seen_keys.add(k)
                    data_row_cnt = sum(
                        1 for rk in row_records
                        if rk[1] == cell["location"] and rk[2] == cell["sheet"])
                    sheet_rows.append({
                        "Location": cell["location"],
                        "Sheet":    cell["sheet"],
                        "File":     cell["file"],
                        "Data Rows": data_row_cnt,
                    })
            tbl = pd.DataFrame(sheet_rows)
            out["answer"] = (f"Found **{len(tbl)}** sheet(s)"
                             f"{' in ' + loc_filter if loc_filter else ''}.")
            out["table"]  = tbl
            return out
 
        # ════════════════════════════════════════════════════════════════
        # INTENT ② – Missing / null values
        # ════════════════════════════════════════════════════════════════
        if w_miss:
            miss_rows = []
            for fname in excel_files:
                loc = location_from_name(fname)
                if not loc_ok(loc):
                    continue
                sh_d = load_file(os.path.join(data_dir, fname))
                for sh, raw in sh_d.items():
                    df_c = to_numeric_cols(smart_header(raw))
                    for col in df_c.columns:
                        mc = int(df_c[col].isna().sum())
                        if mc > 0:
                            miss_rows.append({
                                "Location": loc, "Sheet": sh, "Column": col,
                                "Missing Count": mc,
                                "Missing %": f"{mc/max(len(df_c),1)*100:.1f}%"
                            })
            if miss_rows:
                tbl = pd.DataFrame(miss_rows).sort_values("Missing Count", ascending=False)
                out["answer"] = f"Found **{len(tbl)}** column(s) with missing values."
                out["table"]  = tbl
            else:
                out["answer"] = "✅ No missing values found."
            return out
 
        # ════════════════════════════════════════════════════════════════
        # INTENT ③ – Column listing
        # ════════════════════════════════════════════════════════════════
        if w_cols and not w_num_op:
            col_kw = find_best_col_keyword(sig)
            seen_cols = set()
            col_rows  = []
            for cell in corpus:
                if not cell["is_header"] or not loc_ok(cell):
                    continue
                ch = cell["col_header"].strip()
                if ch in ("", "nan"):
                    continue
                if col_kw and col_kw.lower() not in ch.lower():
                    continue
                k = (cell["location"], cell["sheet"], ch)
                if k not in seen_cols:
                    seen_cols.add(k)
                    col_rows.append({"Location": cell["location"],
                                     "Sheet": cell["sheet"], "Column": ch})
            tbl = pd.DataFrame(col_rows) if col_rows else pd.DataFrame()
            out["answer"] = f"Found **{len(tbl)}** column(s) matching your query."
            out["table"]  = tbl
            return out
 
        # ════════════════════════════════════════════════════════════════
        # INTENT ④ – Row/record count
        # ════════════════════════════════════════════════════════════════
        if w_rows and w_cnt and not sig:
            cnt_rows = []
            total    = 0
            for fname in excel_files:
                loc = location_from_name(fname)
                if not loc_ok(loc):
                    continue
                sh_d = load_file(os.path.join(data_dir, fname))
                for sh, raw in sh_d.items():
                    df_c = smart_header(raw)
                    cnt_rows.append({"Location": loc, "Sheet": sh, "Data Rows": len(df_c)})
                    total += len(df_c)
            tbl = pd.DataFrame(cnt_rows)
            out["answer"] = (f"**{total:,}** total data rows across "
                             f"**{len(tbl)}** sheet(s).")
            out["table"]  = tbl
            out["chart"]  = {"x": "Location", "y": "Data Rows",
                             "title": "Data Rows per Location", "type": "bar"}
            out["chart_df"] = tbl.groupby("Location")["Data Rows"].sum().reset_index()
            return out
 
        # ════════════════════════════════════════════════════════════════
        # INTENT ⑤ – Numeric aggregation (sum / avg / max / min / stats)
        # ════════════════════════════════════════════════════════════════
        if w_num_op and sig:
            col_kw = find_best_col_keyword(sig)
            if col_kw:
                num_pairs = numeric_for_col_kw(col_kw)
 
                if num_pairs:
                    vals = [v for v, _ in num_pairs]
                    s    = pd.Series(vals)
                    parts = []
                    if w_sum  or w_stat: parts.append(f"**Total (Sum):**   {s.sum():,.4f}")
                    if w_avg  or w_stat: parts.append(f"**Average (Mean):** {s.mean():,.4f}")
                    if w_max  or w_stat: parts.append(f"**Maximum:**       {s.max():,.4f}")
                    if w_min  or w_stat: parts.append(f"**Minimum:**       {s.min():,.4f}")
                    if w_cnt  or w_stat: parts.append(f"**Count:**         {s.count():,}")
                    if w_stat:
                        parts.append(
                            f"**Median:**  {s.median():,.4f}  |  "
                            f"**Std Dev:** {s.std():,.4f}")
 
                    # Per-location/sheet breakdown
                    grp = defaultdict(list)
                    for v, cell in num_pairs:
                        grp[f"{cell['location']} | {cell['sheet']}"].append(v)
                    breakdown = []
                    for lbl, vs in grp.items():
                        sv = pd.Series(vs)
                        breakdown.append({
                            "Location | Sheet": lbl,
                            "Count": sv.count(),
                            "Sum":   round(sv.sum(),  4),
                            "Mean":  round(sv.mean(), 4),
                            "Max":   round(sv.max(),  4),
                            "Min":   round(sv.min(),  4),
                        })
                    tbl = pd.DataFrame(breakdown).sort_values("Sum", ascending=False)
 
                    out["answer"] = (
                        f"Results for columns matching **'{col_kw}'** "
                        f"({len(vals):,} numeric values, "
                        f"{len(breakdown)} source(s)"
                        f"{', in ' + loc_filter if loc_filter else ''}):\n\n"
                        + "\n".join(parts)
                    )
                    out["table"]    = tbl
                    out["chart"]    = {"x": "Location | Sheet", "y": "Sum",
                                       "title": f"Sum of '{col_kw}' by Location/Sheet",
                                       "type": "bar"}
                    out["chart_df"] = tbl
 
                    # Top-N sub-table
                    if w_topn:
                        n    = int(w_topn.group(1))
                        top_ = sorted(num_pairs, key=lambda x: x[0], reverse=True)[:n]
                        top_rows = [{
                            "📍 Location": c["location"], "📋 Sheet": c["sheet"],
                            "Row #": c["row"] + 1, "Column": c["col_header"], "Value": v
                        } for v, c in top_]
                        out["sub_tables"].append({
                            "label": f"🏆 Top {n} values for '{col_kw}'",
                            "df":    pd.DataFrame(top_rows)
                        })
                    return out
 
        # ════════════════════════════════════════════════════════════════
        # INTENT ⑥ – Unique values for a column
        # ════════════════════════════════════════════════════════════════
        if w_uniq and sig:
            col_kw = find_best_col_keyword(sig)
            if col_kw:
                unique_vals = set()
                src_rows    = []
                for cell in corpus:
                    if cell["is_header"] or not loc_ok(cell):
                        continue
                    if col_kw.lower() in cell["col_header"].lower():
                        unique_vals.add(cell["value"])
                        src_rows.append({
                            "Location": cell["location"],
                            "Sheet":    cell["sheet"],
                            "Column":   cell["col_header"],
                            "Value":    cell["value"],
                        })
                tbl = (pd.DataFrame(src_rows)
                         .drop_duplicates(subset=["Location","Sheet","Value"])
                       if src_rows else pd.DataFrame())
                out["answer"] = (
                    f"Found **{len(unique_vals)}** unique value(s) in columns "
                    f"matching **'{col_kw}'**"
                    f"{' in '+loc_filter if loc_filter else ''}.")
                out["table"] = tbl
                return out
 
        # ════════════════════════════════════════════════════════════════
        # INTENT ⑦ – Free-text entity / keyword search
        #   Searches VALUES in every cell across every sheet.
        #   Returns matching rows (full row context) + cell hits.
        # ════════════════════════════════════════════════════════════════
        if sig:
            # Prefer exact quoted entity
            quoted = re.findall(r'"([^"]+)"', q)
            if quoted:
                search_terms = [quoted[0].lower()]
            else:
                # Only keep sig words that appear in at least one cell VALUE
                search_terms = [w for w in sig
                                if any(w in cell["value"].lower() for cell in corpus)]
                if not search_terms:
                    search_terms = sig   # fall back
 
            # --- Find matching cells ---
            hit_cells = [
                cell for cell in corpus
                if not cell["is_header"]
                and loc_ok(cell)
                and any(st in cell["value"].lower() for st in search_terms)
            ]
 
            # --- Get unique row keys for those cells ---
            hit_row_keys = {
                (c["file"], c["location"], c["sheet"], c["row"])
                for c in hit_cells
            }
 
            # --- Build full-row DataFrame ---
            full_rows_df = build_rows_df(hit_row_keys)
 
            # --- Location frequency for chart ---
            loc_freq = defaultdict(int)
            for c in hit_cells:
                loc_freq[c["location"]] += 1
            lf_df = (pd.DataFrame(list(loc_freq.items()), columns=["Location","Hits"])
                       .sort_values("Hits", ascending=False))
 
            # --- Cell-hit summaries (first 30) ---
            cell_hit_list = [{
                "📍 Location":   c["location"],
                "📋 Sheet":      c["sheet"],
                "Row #":        c["row"] + 1,
                "Column Header": c["col_header"],
                "Value":        c["value"],
            } for c in hit_cells[:30]]
 
            out["answer"] = (
                f"Found **{len(hit_cells):,}** matching cell(s) "
                f"across **{len(loc_freq)}** location(s) and "
                f"**{len(set((c['location'],c['sheet']) for c in hit_cells))}** sheet(s)"
                f"{' in '+loc_filter if loc_filter else ''}.\n\n"
                f"**{len(hit_row_keys):,}** unique data row(s) contain this data."
            )
            out["table"]     = full_rows_df if not full_rows_df.empty else None
            out["cell_hits"] = cell_hit_list
 
            if not lf_df.empty and len(loc_freq) > 1:
                out["chart"]    = {"x":"Location","y":"Hits",
                                   "title":f"Hits per location for '{', '.join(search_terms[:3])}'",
                                   "type":"bar"}
                out["chart_df"] = lf_df
 
            # Customer-name sub-table
            if not full_rows_df.empty:
                for col in full_rows_df.columns:
                    if "customer" in col.lower() and "name" in col.lower():
                        cust_df = (full_rows_df[["📍 Location","📋 Sheet", col]]
                                     .drop_duplicates())
                        out["sub_tables"].append({
                            "label": f"👤 Customer names ({len(cust_df)} rows)",
                            "df":    cust_df
                        })
                        break
            return out
 
        # ════════════════════════════════════════════════════════════════
        # FALLBACK
        # ════════════════════════════════════════════════════════════════
        out["answer"] = (
            "❓ Could not find a match.\n\n"
            "**Try these patterns:**\n"
            "• Entity search:   *Find CISCO*  |  *ORACLE details*  |  *Axis Bank*\n"
            "• Numeric ops:     *Total subscription*  |  *Max capacity in Airoli*  |  *Average power usage*\n"
            "• List rows:       *List all customers in Noida*  |  *Show all metered customers*\n"
            "• Top-N:           *Top 10 subscription values*\n"
            "• Unique values:   *Unique billing models*  |  *Distinct floors*\n"
            "• Meta:            *Show all sheets*  |  *List columns*  |  *Count rows per sheet*\n"
            "• Missing data:    *Show missing values*"
        )
        return out
 
    # ── Render one answer ─────────────────────────────────────────────────
    def render_answer(res: dict):
        st.markdown(
            f'<div class="answer-box">{res["answer"]}</div>',
            unsafe_allow_html=True)
 
        if res.get("table") is not None and not res["table"].empty:
            tbl = res["table"].reset_index(drop=True)
            st.dataframe(tbl, use_container_width=True,
                         height=min(560, 48 + len(tbl) * 36))
            st.download_button(
                "⬇️ Download CSV",
                tbl.to_csv(index=False).encode(),
                "ai_result.csv", "text/csv",
                key=f"dl_{id(res)}"
            )
 
        if res.get("chart") and res.get("chart_df") is not None:
            cd  = res["chart"]
            cdf = res["chart_df"]
            if cd["x"] in cdf.columns and cd["y"] in cdf.columns:
                fig = px.bar(
                    cdf.sort_values(cd["y"], ascending=False).head(30),
                    x=cd["x"], y=cd["y"],
                    color=cd["y"], color_continuous_scale="Viridis",
                    title=cd.get("title",""), height=400
                )
                fig.update_layout(xaxis_tickangle=-35)
                st.plotly_chart(fig, use_container_width=True)
 
        for st_item in res.get("sub_tables", []):
            with st.expander(st_item["label"], expanded=True):
                st.dataframe(st_item["df"], use_container_width=True)
 
        if res.get("cell_hits"):
            with st.expander(
                f"🔬 Cell-level matches — first {len(res['cell_hits'])} shown",
                expanded=False
            ):
                for ch in res["cell_hits"]:
                    st.markdown(
                        f'<div class="cell-hit">'
                        f'📍 <b>{ch["📍 Location"]}</b> → '
                        f'Sheet: <b>{ch["📋 Sheet"]}</b> | '
                        f'Row <b>{ch["Row #"]}</b> | '
                        f'Col: <i>{ch["Column Header"]}</i> | '
                        f'Value: <b>{ch["Value"]}</b>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
 
    # ── Chat history ──────────────────────────────────────────────────────
    if "aisq_hist" not in st.session_state:
        st.session_state.aisq_hist = []
 
    for item in st.session_state.aisq_hist:
        st.markdown(
            f'<div class="bubble-user">🧑 {item["q"]}</div>',
            unsafe_allow_html=True)
        with st.container():
            render_answer(item["res"])
        st.markdown("---")
 
    # ── Input bar ─────────────────────────────────────────────────────────
    st.markdown("---")
    inp_c, btn_c, clr_c = st.columns([7, 1, 1])
    with inp_c:
        user_q = st.text_input(
            "Ask:",
            placeholder=(
                "List all customers in Noida  |  Total subscription in Airoli  |  "
                "Max capacity Bangalore  |  Find CISCO  |  Show missing values  |  "
                "Average power usage  |  Unique billing models  |  Count rows per sheet"
            ),
            label_visibility="collapsed",
            key="aisq_input"
        )
    with btn_c:
        ask_btn = st.button("🔍 Ask", use_container_width=True, type="primary")
    with clr_c:
        if st.button("🗑️ Clear", use_container_width=True):
            st.session_state.aisq_hist = []
            st.rerun()
 
    # ── Quick example chips ───────────────────────────────────────────────
    st.markdown("**💡 Quick examples — click any to ask instantly:**")
    examples = [
        "List all customers",
        "List all customers in Noida",
        "List all customers in Bangalore",
        "List all customers in Kolkata",
        "Total subscription across all locations",
        "Total capacity in Airoli",
        "Average power usage",
        "Maximum rack subscription",
        "Minimum capacity",
        "Top 10 subscription values",
        "Count rows per sheet",
        "Show missing values",
        "Find CISCO",
        "Find Axis Bank",
        "Find MOTMOT",
        "Find ORACLE",
        "Unique billing models",
        "Unique subscription mode",
        "Show all sheets",
        "Statistics of subscription",
        "Average usage in KW",
        "List columns",
    ]
    for row_start in range(0, len(examples), 6):
        chunk = examples[row_start:row_start+6]
        cols  = st.columns(len(chunk))
        for j, ex in enumerate(chunk):
            if cols[j].button(ex, key=f"chip_{ex}", use_container_width=True):
                user_q  = ex
                ask_btn = True
 
    # ── Execute query ─────────────────────────────────────────────────────
    if ask_btn and user_q.strip():
        with st.spinner(f"🤖 Scanning every cell for: **{user_q}**…"):
            answer = ai_smart_query(user_q)
        st.session_state.aisq_hist.append({"q": user_q, "res": answer})
        st.rerun()
 
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Customer & Capacity Tracker · Streamlit + Plotly · Sify DC Files")
