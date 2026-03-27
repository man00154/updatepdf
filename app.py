import os, re, warnings
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

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
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stSidebar"] { background: linear-gradient(180deg,#0f2027,#203a43,#2c5364); }
[data-testid="stSidebar"] * { color: #e0f7fa !important; }
.metric-card {
    background: linear-gradient(135deg,#1e3c72,#2a5298);
    border-radius: 14px; padding: 18px 22px;
    color: #fff; margin-bottom: 10px;
    box-shadow: 0 4px 18px rgba(0,0,0,0.25);
}
.metric-card h2 { font-size: 2rem; margin: 0; }
.metric-card p  { margin: 4px 0 0 0; font-size: 0.9rem; opacity: 0.85; }
.section-title { font-size: 1.3rem; font-weight: 700; color: #1e3c72;
    border-left: 5px solid #2a5298; padding-left: 10px; margin: 18px 0 10px 0; }
.stDataFrame { border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# FILE DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "excel_files")

def find_excel_files(folder: str) -> list[str]:
    if not os.path.isdir(folder):
        return []
    return sorted(
        f for f in os.listdir(folder) if f.lower().endswith((".xlsx", ".xls"))
    )

def location_from_name(fname: str) -> str:
    """Extract a human-friendly location label from filename."""
    name = fname.replace("Customer_and_Capacity_Tracker_", "").replace(".xlsx", "").replace(".xls", "")
    name = re.sub(r"_\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{2}$", "", name, flags=re.I)
    name = re.sub(r"_\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{4}.*$", "", name, flags=re.I)
    name = name.replace("_", " ").strip()
    return name

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING (cached)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_file(path: str) -> dict[str, pd.DataFrame]:
    """Load every sheet from an Excel file; returns dict sheet→DataFrame (raw, no header)."""
    sheets = {}
    try:
        ext = os.path.splitext(path)[1].lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        xf = pd.ExcelFile(path, engine=engine)
        for sh in xf.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sh, header=None, engine=engine, dtype=str)
                sheets[sh] = df
            except Exception:
                pass
    except Exception as e:
        st.sidebar.warning(f"⚠️ Could not open {os.path.basename(path)}: {e}")
    return sheets


def smart_header(df: pd.DataFrame) -> pd.DataFrame:
    """Find best header row (most non-null, non-duplicate strings) and return clean df."""
    best_row, best_score = 0, -1
    for i in range(min(6, len(df))):
        row = df.iloc[i].astype(str)
        score = row.str.strip().str.len().gt(0).sum() - row.duplicated().sum()
        if score > best_score:
            best_score, best_row = score, i
    header = df.iloc[best_row].fillna("").astype(str).str.strip()
    # Make unique
    seen = {}
    unique_header = []
    for col in header:
        if col in seen:
            seen[col] += 1
            unique_header.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            unique_header.append(col)
    data = df.iloc[best_row + 1:].copy()
    data.columns = unique_header
    data = data.reset_index(drop=True)
    # Drop fully-empty rows
    data = data.dropna(how="all")
    return data


def to_numeric_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce columns to numeric where possible."""
    out = df.copy()
    for col in out.columns:
        out[col] = pd.to_numeric(out[col], errors="ignore")
    return out

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/fluency/96/data-center.png", width=70)
st.sidebar.title("📊 Capacity Tracker")
st.sidebar.markdown("---")

# Let user either use default upload folder OR upload files in the sidebar
st.sidebar.subheader("📁 Data Source")
uploaded = st.sidebar.file_uploader(
    "Upload Excel files (optional – overrides folder)",
    type=["xlsx", "xls"], accept_multiple_files=True
)

import tempfile, shutil

@st.cache_data(show_spinner=False)
def save_uploads(files) -> str:
    tmp = tempfile.mkdtemp()
    for f in files:
        with open(os.path.join(tmp, f.name), "wb") as fh:
            fh.write(f.read())
    return tmp

if uploaded:
    data_dir = save_uploads(tuple(uploaded))
else:
    data_dir = UPLOAD_DIR

excel_files = find_excel_files(data_dir)

if not excel_files:
    st.warning(
        "⚠️ No Excel files found.\n\n"
        "**Option 1:** Upload files via the sidebar uploader above.\n\n"
        "**Option 2:** Place your Excel files in a folder called `excel_files/` "
        "next to `app.py` and restart the app."
    )
    st.stop()

# Build location labels
loc_map = {f: location_from_name(f) for f in excel_files}

st.sidebar.subheader("🏙️ Select Location")
selected_file = st.sidebar.selectbox(
    "Location", excel_files, format_func=lambda x: loc_map[x]
)

all_sheets = load_file(os.path.join(data_dir, selected_file))

st.sidebar.subheader("📋 Select Sheet")
selected_sheet = st.sidebar.selectbox("Sheet", list(all_sheets.keys()))

raw_df = all_sheets[selected_sheet]
df_clean = to_numeric_cols(smart_header(raw_df))

st.sidebar.markdown("---")
st.sidebar.subheader("🔢 Numeric Columns")
num_cols = df_clean.select_dtypes(include="number").columns.tolist()
st.sidebar.caption(f"{len(num_cols)} numeric columns detected")

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "🏠 Overview", "📋 Raw Data", "📊 Analytics", "📈 Charts",
    "🥧 Distributions", "🔍 Query Engine", "🌍 Multi-Location", "🤖 AI Agent"
])

location_label = loc_map[selected_file]

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1 – OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.title(f"🏢 {location_label} – {selected_sheet}")
    st.caption(f"File: `{selected_file}` | Raw shape: {raw_df.shape[0]} rows × {raw_df.shape[1]} cols")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""<div class="metric-card"><h2>{len(df_clean)}</h2>
        <p>Total Rows (data)</p></div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""<div class="metric-card" style="background:linear-gradient(135deg,#11998e,#38ef7d);"><h2>{len(df_clean.columns)}</h2>
        <p>Columns</p></div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""<div class="metric-card" style="background:linear-gradient(135deg,#c94b4b,#4b134f);"><h2>{len(num_cols)}</h2>
        <p>Numeric Columns</p></div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""<div class="metric-card" style="background:linear-gradient(135deg,#f7971e,#ffd200);color:#000;"><h2>{len(excel_files)}</h2>
        <p>Files Loaded</p></div>""", unsafe_allow_html=True)

    st.markdown("---")

    # Quick stats for numeric columns
    if num_cols:
        st.markdown('<div class="section-title">📐 Quick Statistics</div>', unsafe_allow_html=True)
        stats = df_clean[num_cols].describe().T
        stats["range"] = stats["max"] - stats["min"]
        st.dataframe(stats.style.format("{:.2f}", na_rep="–")
                     .background_gradient(cmap="Blues", subset=["mean", "max"]),
                     use_container_width=True)

    # Column preview
    st.markdown('<div class="section-title">🗂️ Column Overview</div>', unsafe_allow_html=True)
    col_info = pd.DataFrame({
        "Column": df_clean.columns,
        "Dtype": df_clean.dtypes.values,
        "Non-Null": df_clean.notna().sum().values,
        "Null %": (df_clean.isna().mean() * 100).round(1).values,
        "Sample": [str(df_clean[c].dropna().iloc[0]) if df_clean[c].dropna().shape[0] > 0 else "—"
                   for c in df_clean.columns]
    })
    st.dataframe(col_info, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 2 – RAW DATA
# ═══════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("📋 Data Table")

    search = st.text_input("🔍 Search any cell value", "")
    if search:
        mask = df_clean.apply(lambda c: c.astype(str).str.contains(search, case=False, na=False)).any(axis=1)
        display_df = df_clean[mask]
        st.caption(f"Showing {len(display_df)} matching rows")
    else:
        display_df = df_clean

    st.dataframe(display_df, use_container_width=True, height=520)

    csv = display_df.to_csv(index=False).encode()
    st.download_button("⬇️ Download CSV", csv, "data_export.csv", "text/csv")

    st.markdown("---")
    st.subheader("🗃️ Raw Excel View (no header applied)")
    st.dataframe(raw_df, use_container_width=True, height=300)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 3 – ANALYTICS
# ═══════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("📊 Column-Level Analytics")

    if not num_cols:
        st.info("No numeric columns detected in this sheet.")
    else:
        chosen = st.multiselect("Select columns to analyse", num_cols, default=num_cols[:6])
        if chosen:
            sub = df_clean[chosen].dropna(how="all")

            # Summary cards
            cols_row = st.columns(len(chosen[:6]))
            for i, col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s):
                    with cols_row[i % len(cols_row)]:
                        delta_color = "off"
                        st.metric(col[:25], f"{s.sum():,.2f}", f"μ={s.mean():,.2f}")

            st.markdown("---")
            agg_rows = []
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    agg_rows.append({
                        "Column": col, "Count": int(s.count()),
                        "Sum": s.sum(), "Mean": s.mean(), "Median": s.median(),
                        "Min": s.min(), "Max": s.max(), "Std Dev": s.std(),
                        "Variance": s.var(),
                        "25%": s.quantile(0.25), "75%": s.quantile(0.75),
                        "% Non-Null": f"{s.count()/max(len(df_clean),1)*100:.1f}%",
                        "Sum %": f"{s.sum()/max(df_clean[chosen].sum().sum(),1)*100:.1f}%"
                    })
            if agg_rows:
                agg_df = pd.DataFrame(agg_rows).set_index("Column")
                st.dataframe(agg_df.style.format("{:,.2f}", na_rep="–",
                             subset=[c for c in agg_df.columns if c not in ["% Non-Null","Sum %"]])
                             .background_gradient(cmap="YlOrRd", subset=["Sum", "Max"]),
                             use_container_width=True)

        # Group-by
        st.markdown("---")
        st.markdown('<div class="section-title">🧮 Group-By Aggregation</div>', unsafe_allow_html=True)
        cat_cols = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique() < 50]
        if cat_cols and num_cols:
            g_col = st.selectbox("Group by (categorical)", cat_cols)
            agg_col = st.selectbox("Aggregate column (numeric)", num_cols)
            agg_fn = st.radio("Aggregation", ["sum", "mean", "count", "min", "max", "median"], horizontal=True)
            grouped = df_clean.groupby(g_col)[agg_col].agg(agg_fn).reset_index()
            grouped.columns = [g_col, f"{agg_fn}({agg_col})"]
            grouped = grouped.sort_values(grouped.columns[1], ascending=False)
            st.dataframe(grouped, use_container_width=True)

            fig = px.bar(grouped, x=g_col, y=grouped.columns[1],
                         color=grouped.columns[1], color_continuous_scale="Viridis",
                         title=f"{agg_fn.title()} of {agg_col} by {g_col}")
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
            x_col = st.selectbox("X axis (category)", cat_cols2 or df_clean.columns[:1].tolist())
            y_col = st.selectbox("Y axis (numeric)", num_cols)
            color_col = st.selectbox("Color by", ["None"] + cat_cols2)
            orientation = st.radio("Orientation", ["Vertical", "Horizontal"], horizontal=True)
            data = df_clean[[x_col, y_col]].dropna()
            fig = px.bar(data, x=x_col if orientation=="Vertical" else y_col,
                         y=y_col if orientation=="Vertical" else x_col,
                         color=color_col if color_col != "None" else None,
                         orientation="v" if orientation=="Vertical" else "h",
                         color_continuous_scale="Turbo",
                         title=f"{y_col} by {x_col}")
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Line Chart":
            x_col = st.selectbox("X axis", df_clean.columns.tolist())
            y_cols = st.multiselect("Y axis (numeric)", num_cols, default=num_cols[:2])
            if y_cols:
                data = df_clean[[x_col] + y_cols].dropna(subset=y_cols, how="all")
                fig = px.line(data, x=x_col, y=y_cols, markers=True,
                              title=f"Line chart: {', '.join(y_cols)}")
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Scatter Plot":
            x_col = st.selectbox("X axis", num_cols)
            y_col = st.selectbox("Y axis", num_cols, index=min(1, len(num_cols)-1))
            size_col = st.selectbox("Bubble size", ["None"] + num_cols)
            color_col = st.selectbox("Color by", ["None"] + cat_cols2 + num_cols)
            data = df_clean.dropna(subset=[x_col, y_col])
            fig = px.scatter(data, x=x_col, y=y_col,
                             size=size_col if size_col!="None" else None,
                             color=color_col if color_col!="None" else None,
                             hover_data=data.columns[:5].tolist(),
                             color_continuous_scale="Rainbow",
                             title=f"{y_col} vs {x_col}")
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Area Chart":
            x_col = st.selectbox("X axis", df_clean.columns.tolist())
            y_cols = st.multiselect("Y columns", num_cols, default=num_cols[:3])
            if y_cols:
                data = df_clean[[x_col]+y_cols].dropna(subset=y_cols, how="all")
                fig = px.area(data, x=x_col, y=y_cols, title="Area Chart")
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Bubble Chart":
            if len(num_cols) >= 3:
                x_col = st.selectbox("X", num_cols, index=0)
                y_col = st.selectbox("Y", num_cols, index=1)
                s_col = st.selectbox("Bubble Size", num_cols, index=2)
                l_col = st.selectbox("Label/Color", ["None"] + cat_cols2)
                data = df_clean[[x_col, y_col, s_col]].dropna()
                if l_col != "None":
                    data[l_col] = df_clean[l_col]
                fig = px.scatter(data, x=x_col, y=y_col, size=s_col,
                                 color=l_col if l_col!="None" else None,
                                 size_max=60, color_discrete_sequence=px.colors.qualitative.Vivid,
                                 title="Bubble Chart")
                fig.update_layout(height=520)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need at least 3 numeric columns for a bubble chart.")

        elif chart_type == "Heatmap (Correlation)":
            sel = st.multiselect("Columns for correlation", num_cols, default=num_cols[:10])
            if len(sel) >= 2:
                corr = df_clean[sel].corr()
                fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="RdBu_r",
                                aspect="auto", title="Correlation Heatmap")
                fig.update_layout(height=520)
                st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Box Plot":
            y_col = st.selectbox("Value column", num_cols)
            x_col = st.selectbox("Group by", ["None"] + cat_cols2)
            data = df_clean[[y_col] + ([x_col] if x_col!="None" else [])].dropna(subset=[y_col])
            fig = px.box(data, y=y_col, x=x_col if x_col!="None" else None,
                         color=x_col if x_col!="None" else None,
                         color_discrete_sequence=px.colors.qualitative.Pastel,
                         points="outliers", title=f"Box Plot: {y_col}")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Funnel Chart":
            x_col = st.selectbox("Stage (category)", cat_cols2 or df_clean.columns[:1].tolist())
            y_col = st.selectbox("Value", num_cols)
            data = df_clean[[x_col, y_col]].dropna().groupby(x_col)[y_col].sum().reset_index()
            data = data.sort_values(y_col, ascending=False)
            fig = px.funnel(data, x=y_col, y=x_col, title="Funnel Chart")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Waterfall / Cumulative":
            y_col = st.selectbox("Numeric column", num_cols)
            data = df_clean[y_col].dropna().reset_index(drop=True)
            cumulative = data.cumsum()
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Value", x=data.index, y=data, marker_color="#2a5298"))
            fig.add_trace(go.Scatter(name="Cumulative", x=cumulative.index, y=cumulative,
                                     line=dict(color="#f7971e", width=2.5), mode="lines+markers"))
            fig.update_layout(title=f"Cumulative: {y_col}", height=450, barmode="group")
            st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 5 – DISTRIBUTIONS
# ═══════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("🥧 Distribution Charts")

    if not num_cols:
        st.info("No numeric columns in this sheet.")
    else:
        c1, c2 = st.columns(2)

        with c1:
            # Pie chart
            cat_cols3 = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique() <= 30]
            if cat_cols3:
                cat_c = st.selectbox("Pie – category", cat_cols3, key="pie_cat")
                val_c = st.selectbox("Pie – value", num_cols, key="pie_val")
                pie_data = df_clean[[cat_c, val_c]].dropna()
                pie_data[val_c] = pd.to_numeric(pie_data[val_c], errors="coerce")
                pie_data = pie_data.dropna().groupby(cat_c)[val_c].sum().reset_index()
                fig = px.pie(pie_data, names=cat_c, values=val_c,
                             hole=0.35, color_discrete_sequence=px.colors.qualitative.Vivid,
                             title=f"Distribution of {val_c}")
                st.plotly_chart(fig, use_container_width=True)

        with c2:
            # Histogram
            hist_col = st.selectbox("Histogram column", num_cols, key="hist_col")
            bins = st.slider("Bins", 5, 80, 20)
            fig = px.histogram(df_clean[hist_col].dropna(), nbins=bins,
                               color_discrete_sequence=["#11998e"],
                               title=f"Histogram: {hist_col}")
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        # Treemap
        st.markdown("---")
        st.markdown('<div class="section-title">🗺️ Treemap</div>', unsafe_allow_html=True)
        cat_cols4 = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique() <= 40]
        if cat_cols4 and num_cols:
            tm_cat = st.selectbox("Treemap – category", cat_cols4)
            tm_val = st.selectbox("Treemap – value", num_cols)
            tm_data = df_clean[[tm_cat, tm_val]].dropna().groupby(tm_cat)[tm_val].sum().reset_index()
            tm_data = tm_data[tm_data[tm_val] > 0]
            if len(tm_data):
                fig = px.treemap(tm_data, path=[tm_cat], values=tm_val,
                                 color=tm_val, color_continuous_scale="Turbo",
                                 title=f"Treemap: {tm_val} by {tm_cat}")
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)

        # Violin plot
        st.markdown("---")
        st.markdown('<div class="section-title">🎻 Violin Plot</div>', unsafe_allow_html=True)
        if num_cols:
            viol_col = st.selectbox("Violin column", num_cols)
            fig = px.violin(df_clean[viol_col].dropna(), y=viol_col, box=True, points="outliers",
                            color_discrete_sequence=["#c94b4b"],
                            title=f"Violin Plot: {viol_col}")
            st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 6 – QUERY ENGINE
# ═══════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("🔍 AI Query Engine")
    st.info("Ask questions in plain English. The engine maps your request to pandas operations.")

    query = st.text_input("Your question", placeholder="e.g. What is the total of column X? Show max of Y per group Z.")

    def run_query(query_text: str, df: pd.DataFrame, n_cols: list) -> str:
        q = query_text.lower()
        results = []

        # SUM
        if any(w in q for w in ["sum", "total"]):
            for col in n_cols:
                if col.lower() in q or "all" in q or len(n_cols) == 1:
                    results.append(f"**SUM of `{col}`**: {df[col].sum():,.2f}")
        # MEAN / AVERAGE
        if any(w in q for w in ["mean", "average", "avg"]):
            for col in n_cols:
                if col.lower() in q or "all" in q or len(n_cols)==1:
                    results.append(f"**MEAN of `{col}`**: {df[col].mean():,.2f}")
        # MAX
        if "max" in q or "maximum" in q or "highest" in q:
            for col in n_cols:
                if col.lower() in q or "all" in q or len(n_cols)==1:
                    results.append(f"**MAX of `{col}`**: {df[col].max():,.2f}")
        # MIN
        if "min" in q or "minimum" in q or "lowest" in q:
            for col in n_cols:
                if col.lower() in q or "all" in q or len(n_cols)==1:
                    results.append(f"**MIN of `{col}`**: {df[col].min():,.2f}")
        # COUNT
        if "count" in q or "how many" in q:
            results.append(f"**Total rows**: {len(df)}")
            for col in n_cols:
                if col.lower() in q:
                    results.append(f"**Non-null count of `{col}`**: {df[col].count()}")
        # MEDIAN
        if "median" in q:
            for col in n_cols:
                if col.lower() in q or len(n_cols)==1:
                    results.append(f"**MEDIAN of `{col}`**: {df[col].median():,.2f}")
        # STD / VARIANCE
        if "std" in q or "deviation" in q:
            for col in n_cols:
                if col.lower() in q or len(n_cols)==1:
                    results.append(f"**STD DEV of `{col}`**: {df[col].std():,.2f}")
        if "variance" in q or "var" in q:
            for col in n_cols:
                if col.lower() in q or len(n_cols)==1:
                    results.append(f"**VARIANCE of `{col}`**: {df[col].var():,.2f}")
        # PERCENT
        if "percent" in q or "%" in q or "ratio" in q:
            for col in n_cols:
                if col.lower() in q:
                    pct = df[col].sum() / df[n_cols].sum().sum() * 100 if df[n_cols].sum().sum() else 0
                    results.append(f"**% share of `{col}`**: {pct:.2f}%")
        # DESCRIBE
        if any(w in q for w in ["describe", "statistics", "stats", "summary"]):
            results.append("**Descriptive Statistics:**")
            results.append(df[n_cols].describe().to_markdown())
        # COLUMNS
        if "column" in q and "list" in q:
            results.append("**Columns:** " + ", ".join(df.columns.tolist()))
        # ROWS
        if "row" in q and ("first" in q or "top" in q):
            results.append("**First 5 rows:**\n" + df.head().to_markdown())
        # MISSING
        if "missing" in q or "null" in q or "nan" in q:
            null_info = df.isna().sum()
            null_info = null_info[null_info > 0]
            results.append("**Missing values:**\n" + null_info.to_string())
        # UNIQUE
        if "unique" in q:
            for col in df.columns:
                if col.lower() in q:
                    u = df[col].dropna().unique()
                    results.append(f"**Unique values of `{col}`** ({len(u)}): {', '.join(map(str, u[:20]))}")

        if not results:
            results.append("ℹ️ Could not match your query. Try keywords like: **sum, average, max, min, count, median, describe, unique, missing**.")

        return "\n\n".join(results)

    if query:
        answer = run_query(query, df_clean, num_cols)
        st.markdown(answer)

    # Manual compute panel
    st.markdown("---")
    st.markdown('<div class="section-title">🧮 Manual Compute</div>', unsafe_allow_html=True)
    if num_cols:
        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            op = st.selectbox("Operation", ["Sum", "Mean", "Max", "Min", "Count", "Median",
                                             "Std Dev", "Variance", "% of total", "Cumulative Sum"])
        with mc2:
            sel_col = st.selectbox("Column", num_cols)
        with mc3:
            filter_col = st.selectbox("Filter by", ["None"] + [c for c in df_clean.columns if c not in num_cols])

        filter_val = None
        if filter_col != "None":
            unique_vals = df_clean[filter_col].dropna().unique().tolist()
            filter_val = st.selectbox("Filter value", unique_vals)

        data_slice = df_clean.copy()
        if filter_col != "None" and filter_val is not None:
            data_slice = data_slice[data_slice[filter_col] == filter_val]

        s = data_slice[sel_col].dropna()
        ops_map = {
            "Sum": s.sum(), "Mean": s.mean(), "Max": s.max(), "Min": s.min(),
            "Count": s.count(), "Median": s.median(), "Std Dev": s.std(),
            "Variance": s.var(),
            "% of total": f"{s.sum()/max(df_clean[sel_col].sum(),1)*100:.2f}%",
            "Cumulative Sum": s.cumsum().iloc[-1] if len(s) else 0
        }
        result = ops_map.get(op, "N/A")
        if isinstance(result, float):
            result = f"{result:,.4f}"
        st.success(f"**{op} of `{sel_col}`** {f'(where {filter_col}={filter_val})' if filter_val else ''}: **{result}**")

        if op == "Cumulative Sum":
            cs = s.cumsum().reset_index(drop=True)
            fig = px.line(cs, title=f"Cumulative Sum: {sel_col}", markers=True)
            st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 7 – MULTI-LOCATION
# ═══════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("🌍 Cross-Location Comparison")
    st.info("Compares numeric data across all loaded Excel files side by side.")

    @st.cache_data(show_spinner=True)
    def load_all_summaries(files, folder):
        summaries = {}
        for f in files:
            sheets = load_file(os.path.join(folder, f))
            for sh_name, raw in sheets.items():
                df_c = to_numeric_cols(smart_header(raw))
                nc = df_c.select_dtypes(include="number").columns.tolist()
                if nc:
                    summaries[f"{loc_map[f]} | {sh_name}"] = {
                        "df": df_c, "num_cols": nc, "file": f, "sheet": sh_name
                    }
        return summaries

    all_summaries = load_all_summaries(tuple(excel_files), data_dir)

    if all_summaries:
        comp_col = st.selectbox(
            "Compare by column name (across locations)",
            sorted(set(c for v in all_summaries.values() for c in v["num_cols"]))
        )

        rows = []
        for label, info in all_summaries.items():
            if comp_col in info["num_cols"]:
                s = info["df"][comp_col].dropna()
                rows.append({
                    "Location|Sheet": label,
                    "Sum": s.sum(), "Mean": s.mean(),
                    "Max": s.max(), "Min": s.min(), "Count": s.count()
                })

        if rows:
            cmp_df = pd.DataFrame(rows).set_index("Location|Sheet")
            st.dataframe(cmp_df.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),
                         use_container_width=True)

            fig = px.bar(cmp_df.reset_index(), x="Location|Sheet", y="Sum",
                         color="Sum", color_continuous_scale="Viridis",
                         title=f"Sum of '{comp_col}' across all locations")
            fig.update_layout(xaxis_tickangle=-30, height=450)
            st.plotly_chart(fig, use_container_width=True)

            fig2 = px.scatter(cmp_df.reset_index(), x="Mean", y="Max",
                              size="Sum", text="Location|Sheet",
                              color="Count", color_continuous_scale="Turbo",
                              title="Bubble: Mean vs Max (size=Sum)")
            fig2.update_traces(textposition="top center")
            fig2.update_layout(height=500)
            st.plotly_chart(fig2, use_container_width=True)

            # Radar chart
            st.markdown("---")
            st.markdown('<div class="section-title">🕸️ Radar (Spider) Chart</div>', unsafe_allow_html=True)
            radar_metric = st.radio("Radar metric", ["Sum", "Mean", "Max"], horizontal=True)
            norm = cmp_df[radar_metric] / cmp_df[radar_metric].max()
            theta = norm.index.tolist()
            r = norm.values.tolist()
            fig3 = go.Figure(go.Scatterpolar(r=r+[r[0]], theta=theta+[theta[0]],
                                              fill="toself", fillcolor="rgba(42,82,152,0.3)",
                                              line=dict(color="#2a5298")))
            fig3.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0,1])),
                               height=500, title=f"Radar: {radar_metric} of '{comp_col}'")
            st.plotly_chart(fig3, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 8 – AI AGENT
# ═══════════════════════════════════════════════════════════════════════════
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights")
    st.info("The agent scans all sheets and surfaces automated insights, anomalies and KPIs.")

    if st.button("🚀 Run AI Agent Analysis"):
        with st.spinner("Analysing all sheets…"):

            for label, info in list(all_summaries.items())[:8]:  # limit to 8 to avoid overload
                df_a = info["df"]
                nc = info["num_cols"]
                if not nc:
                    continue
                with st.expander(f"📍 {label}", expanded=False):
                    c1, c2 = st.columns(2)

                    # KPI Cards
                    with c1:
                        st.markdown("**📊 Key KPIs**")
                        for col in nc[:4]:
                            s = df_a[col].dropna()
                            if len(s):
                                utilization = s.sum() / (s.max() * s.count()) * 100 if s.max() else 0
                                st.metric(col[:30], f"{s.sum():,.1f}", f"Avg {s.mean():,.1f}")

                    with c2:
                        st.markdown("**⚠️ Anomaly Detection (Z-score)**")
                        for col in nc[:3]:
                            s = df_a[col].dropna()
                            if len(s) > 3:
                                z = (s - s.mean()) / s.std()
                                outliers = z[z.abs() > 2.5]
                                if len(outliers):
                                    st.warning(f"`{col}`: {len(outliers)} potential outlier(s)")
                                else:
                                    st.success(f"`{col}`: No outliers detected")

                    # Mini chart
                    if len(nc) >= 2:
                        fig = px.bar(df_a[nc[:2]].dropna().reset_index(),
                                     x="index", y=nc[0],
                                     color_discrete_sequence=["#2a5298"],
                                     title=f"{nc[0]}")
                        fig.update_layout(height=280, showlegend=False, margin=dict(t=35,b=0))
                        st.plotly_chart(fig, use_container_width=True)

                    # Correlations
                    if len(nc) >= 2:
                        corr_series = df_a[nc].corr().unstack().drop_duplicates()
                        corr_series = corr_series[corr_series.index.get_level_values(0) !=
                                                   corr_series.index.get_level_values(1)]
                        corr_series = corr_series.abs().sort_values(ascending=False)
                        if len(corr_series):
                            top_pair = corr_series.index[0]
                            st.info(f"Highest correlation: **{top_pair[0]}** ↔ **{top_pair[1]}** "
                                    f"({corr_series.iloc[0]:.3f})")

    st.markdown("---")
    st.markdown('<div class="section-title">📁 All Files Summary</div>', unsafe_allow_html=True)
    file_summary = []
    for f in excel_files:
        sheets = all_sheets if f == selected_file else load_file(os.path.join(data_dir, f))
        total_rows = sum(len(s) for s in sheets.values())
        total_cols = sum(len(s.columns) for s in sheets.values())
        file_summary.append({
            "File": loc_map[f], "Sheets": len(sheets),
            "Total Rows": total_rows, "Total Columns": total_cols,
            "Filename": f
        })
    fs_df = pd.DataFrame(file_summary)
    st.dataframe(fs_df, use_container_width=True)

    fig = px.bar(fs_df, x="File", y="Total Rows", color="Sheets",
                 color_continuous_scale="Blues",
                 title="Total rows per location file")
    fig.update_layout(xaxis_tickangle=-30, height=380)
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Customer & Capacity Tracker Dashboard · Built with Streamlit + Plotly · Data: Sify DC Files")
