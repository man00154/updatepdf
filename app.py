import os, re, warnings
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════
# PAGE CONFIG & CSS
# ═══════════════════════════════════════════════════════
st.set_page_config(
    page_title="Sify DC – Capacity Tracker",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

/* ── Hide Streamlit chrome ── */
#MainMenu { visibility: hidden !important; }
header[data-testid="stHeader"] { visibility: hidden !important; height: 0 !important; }
footer { visibility: hidden !important; }
[data-testid="stToolbar"]   { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
[data-testid="stStatusWidget"] { display: none !important; }

/* ── Font everywhere ── */
html, body, button, input, select, textarea {
    font-family: 'Inter', sans-serif !important;
}

/* ════════════════════════════════
   BACKGROUND  — soft light grey
════════════════════════════════ */
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
.main {
    background: #f4f6fb !important;
}
.block-container {
    background: transparent !important;
    padding-top: 0.8rem !important;
    max-width: 100% !important;
    padding-left: 1.2rem !important;
    padding-right: 1.2rem !important;
}

/* ════════════════════════════════
   SIDEBAR — dark navy / teal
════════════════════════════════ */
[data-testid="stSidebar"] {
    background: linear-gradient(175deg, #0b1120 0%, #0f2027 35%, #1a3a4a 65%, #0d4f3c 100%) !important;
    border-right: none !important;
    box-shadow: 4px 0 24px rgba(0,0,0,.35) !important;
}
/* All sidebar text — bright white, bold */
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3,
[data-testid="stSidebar"] .stMarkdown li,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
    color: #ffffff !important;
    font-weight: 700 !important;
}
[data-testid="stSidebar"] label {
    color: #86efac !important;
    font-weight: 700 !important;
    font-size: 0.85rem !important;
}
[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] > div,
[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] > div {
    background: rgba(255,255,255,.1) !important;
    border: 1.5px solid rgba(255,255,255,.3) !important;
    color: #ffffff !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] [data-testid="stMarkdownContainer"],
[data-testid="stSidebar"] [data-baseweb="select"] * {
    color: #ffffff !important;
    font-weight: 600 !important;
}
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.2) !important; }
[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    background: rgba(255,255,255,.07) !important;
    border: 1.5px dashed rgba(255,255,255,.4) !important;
    border-radius: 10px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] label,
[data-testid="stSidebar"] [data-testid="stFileUploader"] small,
[data-testid="stSidebar"] [data-testid="stFileUploader"] p {
    color: rgba(255,255,255,.75) !important;
    font-weight: 600 !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
    background: rgba(255,255,255,.15) !important;
    color: #ffffff !important;
    border: 1px solid rgba(255,255,255,.3) !important;
    border-radius: 6px !important;
}
[data-testid="stSidebar"] small,
[data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
    color: #86efac !important;
    font-weight: 600 !important;
}

/* ════════════════════════════════
   MAIN CONTENT headings & text
════════════════════════════════ */
.main [data-testid="stMarkdownContainer"] h1,
.main .stMarkdown h1 {
    color: #0f172a !important;
    font-weight: 900 !important;
    font-size: 1.9rem !important;
    letter-spacing: -0.3px;
}
.main [data-testid="stMarkdownContainer"] h2,
.main .stMarkdown h2 {
    color: #1e293b !important;
    font-weight: 800 !important;
}
.main [data-testid="stMarkdownContainer"] h3,
.main .stMarkdown h3 {
    color: #1d4ed8 !important;
    font-weight: 700 !important;
}
.main [data-testid="stMarkdownContainer"] p,
.main .stMarkdown p {
    color: #334155 !important;
    font-weight: 500 !important;
    font-size: 0.95rem !important;
    line-height: 1.7;
}
.main [data-testid="stMarkdownContainer"] li,
.main .stMarkdown li {
    color: #334155 !important;
    font-weight: 500 !important;
}
.main label {
    color: #374151 !important;
    font-weight: 700 !important;
    font-size: 0.88rem !important;
}

/* ════════════════════════════════
   TABS — tab bar
════════════════════════════════ */
[data-testid="stTabs"] [role="tablist"] {
    background: #ffffff !important;
    border-bottom: 2px solid #e2e8f0 !important;
    border-radius: 12px 12px 0 0;
    padding: 4px 6px 0;
    gap: 3px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
}
[data-testid="stTabs"] [role="tab"] {
    color: #64748b !important;
    background: #f8fafc !important;
    border-radius: 9px 9px 0 0 !important;
    border: 1px solid #e2e8f0 !important;
    border-bottom: none !important;
    padding: 10px 18px !important;
    font-weight: 700 !important;
    font-size: 0.87rem !important;
    transition: all 0.18s;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: #1e293b !important;
    background: #ffffff !important;
    border-color: #e2e8f0 !important;
    border-bottom: 3px solid #ef4444 !important;
    font-weight: 800 !important;
}
[data-testid="stTabs"] [role="tab"]:hover:not([aria-selected="true"]) {
    color: #1d4ed8 !important;
    background: #eff6ff !important;
}
/* Tab content panel — WHITE, readable */
[data-testid="stTabsContent"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-top: none !important;
    border-radius: 0 0 14px 14px !important;
    padding: 28px !important;
    box-shadow: 0 4px 20px rgba(0,0,0,.08);
}

/* ════════════════════════════════
   ALL TEXT inside tab panels
════════════════════════════════ */
[data-testid="stTabsContent"] [data-testid="stMarkdownContainer"] h1,
[data-testid="stTabsContent"] .stMarkdown h1 {
    color: #0f172a !important; font-weight: 900 !important; font-size: 1.7rem !important;
}
[data-testid="stTabsContent"] [data-testid="stMarkdownContainer"] h2,
[data-testid="stTabsContent"] .stMarkdown h2 {
    color: #1e293b !important; font-weight: 800 !important;
}
[data-testid="stTabsContent"] [data-testid="stMarkdownContainer"] h3,
[data-testid="stTabsContent"] .stMarkdown h3 {
    color: #1d4ed8 !important; font-weight: 700 !important;
}
[data-testid="stTabsContent"] [data-testid="stMarkdownContainer"] p,
[data-testid="stTabsContent"] .stMarkdown p {
    color: #334155 !important; font-weight: 500 !important;
}
[data-testid="stTabsContent"] [data-testid="stMarkdownContainer"] li,
[data-testid="stTabsContent"] .stMarkdown li {
    color: #334155 !important; font-weight: 500 !important;
}
[data-testid="stTabsContent"] label {
    color: #374151 !important; font-weight: 700 !important;
}
[data-testid="stTabsContent"] small,
[data-testid="stTabsContent"] .stCaption,
[data-testid="stTabsContent"] [data-testid="stCaptionContainer"] {
    color: #64748b !important; font-weight: 600 !important;
}
[data-testid="stTabsContent"] code {
    background: #f1f5f9 !important; color: #7c3aed !important;
    border-radius: 5px; padding: 2px 7px; font-weight: 700 !important;
    border: 1px solid #e2e8f0;
}
[data-testid="stTabsContent"] hr { border-color: #e2e8f0 !important; }

/* ════════════════════════════════
   INPUTS & SELECTS (tab area)
════════════════════════════════ */
[data-testid="stTabsContent"] [data-testid="stTextInput"] input,
[data-testid="stTabsContent"] .stTextInput input {
    background: #f8fafc !important;
    color: #1e293b !important;
    border: 1.5px solid #cbd5e1 !important;
    border-radius: 9px !important;
    padding: 10px 14px !important;
    font-size: 0.94rem !important;
    font-weight: 600 !important;
    box-shadow: inset 0 1px 3px rgba(0,0,0,.05);
}
[data-testid="stTabsContent"] [data-testid="stTextInput"] input::placeholder {
    color: #94a3b8 !important;
}
[data-testid="stTabsContent"] [data-testid="stTextInput"] input:focus {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 3px rgba(99,102,241,.18) !important;
    background: #ffffff !important;
}
[data-testid="stTabsContent"] .stSelectbox [data-baseweb="select"] > div,
[data-testid="stTabsContent"] .stMultiSelect [data-baseweb="select"] > div {
    background: #f8fafc !important;
    border: 1.5px solid #cbd5e1 !important;
    border-radius: 9px !important;
    color: #1e293b !important;
}
[data-testid="stTabsContent"] [data-baseweb="select"] * { color: #1e293b !important; font-weight: 600 !important; }
[data-testid="stTabsContent"] [data-baseweb="menu"]    { background: #ffffff !important; border: 1px solid #e2e8f0 !important; box-shadow: 0 8px 24px rgba(0,0,0,.12) !important; }
[data-testid="stTabsContent"] [data-baseweb="option"]  { color: #1e293b !important; font-weight: 600 !important; }
[data-testid="stTabsContent"] [data-baseweb="option"]:hover { background: #f1f5f9 !important; }

/* ════════════════════════════════
   BUTTONS
════════════════════════════════ */
[data-testid="stTabsContent"] [data-testid="baseButton-primary"],
[data-testid="stTabsContent"] button[kind="primary"] {
    background: linear-gradient(135deg, #ef4444, #f97316) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 9px !important;
    font-weight: 800 !important;
    font-size: 0.9rem !important;
    box-shadow: 0 4px 14px rgba(239,68,68,.38) !important;
    transition: all 0.2s !important;
}
[data-testid="stTabsContent"] [data-testid="baseButton-primary"]:hover { 
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(239,68,68,.5) !important;
}
[data-testid="stTabsContent"] [data-testid="baseButton-secondary"],
[data-testid="stTabsContent"] .stButton > button {
    background: #ffffff !important;
    color: #1d4ed8 !important;
    border: 1.5px solid #93c5fd !important;
    border-radius: 9px !important;
    font-weight: 700 !important;
    transition: all 0.2s !important;
}
[data-testid="stTabsContent"] [data-testid="baseButton-secondary"]:hover,
[data-testid="stTabsContent"] .stButton > button:hover {
    background: #eff6ff !important;
    border-color: #2563eb !important;
    color: #1e3a8a !important;
}

/* ════════════════════════════════
   DATAFRAMES — crisp white + readable
════════════════════════════════ */
/* Outer wrapper */
[data-testid="stTabsContent"] .stDataFrame,
[data-testid="stTabsContent"] .dvn-scroller,
.stDataFrame, .dvn-scroller,
[data-testid="stDataFrame"],
[data-testid="stDataFrameResizable"] {
    background: #ffffff !important;
    border: 1.5px solid #e2e8f0 !important;
    border-radius: 10px !important;
    box-shadow: 0 2px 10px rgba(0,0,0,.08);
}

/* Glide Data Grid CSS custom properties — forces dark text on white cells */
.dvn-scroller,
[data-testid="stDataFrame"] .dvn-scroller {
    --gdg-bg-cell:              #ffffff !important;
    --gdg-bg-cell-medium:       #f8fafc !important;
    --gdg-bg-header:            #f1f5f9 !important;
    --gdg-bg-header-hovered:    #e2e8f0 !important;
    --gdg-bg-header-has-focus:  #dbeafe !important;
    --gdg-text-dark:            #0f172a !important;
    --gdg-text-medium:          #374151 !important;
    --gdg-text-light:           #64748b !important;
    --gdg-border-color:         #e2e8f0 !important;
    --gdg-accent-color:         #4f46e5 !important;
    --gdg-accent-light:         rgba(79,70,241,.12) !important;
    --gdg-selection-color:      rgba(79,70,241,.2) !important;
    --gdg-header-font-style:    700 !important;
    color: #0f172a !important;
}

/* Fallback: all child elements of the data grid → dark text */
.dvn-scroller *,
[data-testid="stDataFrame"] *,
[data-testid="stDataFrameResizable"] * {
    color: #0f172a !important;
}

/* Static HTML tables (st.table) */
[data-testid="stTable"] table,
.stTable table {
    width: 100%; border-collapse: collapse;
    background: #ffffff !important;
    border-radius: 10px; overflow: hidden;
    box-shadow: 0 2px 10px rgba(0,0,0,.08);
}
[data-testid="stTable"] thead th,
.stTable thead th {
    background: #f1f5f9 !important;
    color: #0f172a !important;
    font-weight: 800 !important;
    font-size: 0.84rem !important;
    padding: 10px 14px !important;
    border-bottom: 2px solid #e2e8f0 !important;
    text-align: left !important;
    letter-spacing: 0.3px;
}
[data-testid="stTable"] tbody td,
.stTable tbody td {
    color: #1e293b !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    padding: 9px 14px !important;
    border-bottom: 1px solid #f1f5f9 !important;
    background: #ffffff !important;
}
[data-testid="stTable"] tbody tr:hover td,
.stTable tbody tr:hover td {
    background: #f8fafc !important;
    color: #0f172a !important;
}
[data-testid="stTable"] tbody tr:nth-child(even) td,
.stTable tbody tr:nth-child(even) td {
    background: #fafbfc !important;
}

/* ════════════════════════════════
   EXPANDERS
════════════════════════════════ */
[data-testid="stTabsContent"] [data-testid="stExpander"] {
    background: #f8fafc !important;
    border: 1.5px solid #e2e8f0 !important;
    border-radius: 12px !important;
    box-shadow: 0 2px 8px rgba(0,0,0,.06);
}
[data-testid="stTabsContent"] [data-testid="stExpander"] summary {
    color: #1e293b !important;
    font-weight: 800 !important;
    font-size: 0.96rem !important;
}
[data-testid="stTabsContent"] [data-testid="stExpander"] summary:hover { color: #1d4ed8 !important; }
[data-testid="stTabsContent"] [data-testid="stExpander"] [data-testid="stMarkdownContainer"] p,
[data-testid="stTabsContent"] [data-testid="stExpander"] .stMarkdown p {
    color: #374151 !important; font-weight: 500 !important;
}

/* ════════════════════════════════
   ALERTS
════════════════════════════════ */
[data-testid="stTabsContent"] [data-testid="stAlert"],
[data-testid="stTabsContent"] .stAlert { border-radius: 10px !important; }
[data-testid="stTabsContent"] [data-testid="stNotification"],
[data-testid="stTabsContent"] [data-testid="stAlert"] [data-testid="stMarkdownContainer"] p {
    font-weight: 700 !important;
}
[data-testid="stTabsContent"] .element-container div[class*="success"],
[data-testid="stTabsContent"] .stSuccess {
    background: #ecfdf5 !important; border-left: 4px solid #059669 !important;
    border-radius: 10px !important;
}
[data-testid="stTabsContent"] .element-container div[class*="info"],
[data-testid="stTabsContent"] .stInfo {
    background: #eff6ff !important; border-left: 4px solid #2563eb !important;
    border-radius: 10px !important;
}
[data-testid="stTabsContent"] .element-container div[class*="warning"],
[data-testid="stTabsContent"] .stWarning {
    background: #fffbeb !important; border-left: 4px solid #f59e0b !important;
    border-radius: 10px !important;
}
[data-testid="stTabsContent"] .element-container div[class*="error"],
[data-testid="stTabsContent"] .stError {
    background: #fef2f2 !important; border-left: 4px solid #ef4444 !important;
    border-radius: 10px !important;
}

/* Radios & checkboxes */
[data-testid="stTabsContent"] [data-testid="stRadio"] label { color: #374151 !important; font-weight: 600 !important; }
[data-testid="stTabsContent"] .stCheckbox label             { color: #374151 !important; font-weight: 600 !important; }
[data-testid="stTabsContent"] [data-testid="stSlider"] > div > div > div { background: #6366f1 !important; }

/* Metric widgets in white area */
[data-testid="stTabsContent"] [data-testid="metric-container"] {
    background: #f8fafc !important;
    border: 1.5px solid #e2e8f0 !important;
    border-radius: 14px !important;
    padding: 16px !important;
    border-left: 4px solid #6366f1 !important;
    box-shadow: 0 2px 8px rgba(0,0,0,.07) !important;
}
[data-testid="stTabsContent"] [data-testid="metric-container"] label {
    color: #64748b !important; font-weight: 700 !important;
    font-size: 0.78rem !important; text-transform: uppercase; letter-spacing: 0.6px;
}
[data-testid="stTabsContent"] [data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #0f172a !important; font-size: 1.9rem !important; font-weight: 900 !important;
}
[data-testid="stTabsContent"] [data-testid="metric-container"] [data-testid="stMetricDelta"] {
    color: #059669 !important; font-weight: 700 !important;
}

/* ════════════════════════════════
   COLORFUL METRIC CARDS (.kcard)
════════════════════════════════ */
.kcard {
    border-radius: 16px; padding: 22px 26px; color: #fff;
    margin-bottom: 12px;
    box-shadow: 0 8px 28px rgba(0,0,0,.22);
    transition: transform .22s, box-shadow .22s;
    border: 1.5px solid rgba(255,255,255,.18);
}
.kcard:hover { transform: translateY(-5px); box-shadow: 0 16px 40px rgba(0,0,0,.3); }
.kcard h2 {
    font-size: 2.1rem; margin: 0; font-weight: 900;
    color: #ffffff !important; text-shadow: 0 2px 8px rgba(0,0,0,.25);
}
.kcard p {
    margin: 6px 0 0; font-size: 0.87rem;
    color: rgba(255,255,255,.95) !important; font-weight: 700 !important;
}
.kcard-blue   { background: linear-gradient(135deg, #1e40af, #3b82f6); }
.kcard-green  { background: linear-gradient(135deg, #065f46, #059669); }
.kcard-red    { background: linear-gradient(135deg, #7f1d1d, #ef4444); }
.kcard-orange { background: linear-gradient(135deg, #78350f, #f59e0b); }
.kcard-teal   { background: linear-gradient(135deg, #0f4c75, #0ea5e9); }
.kcard-purple { background: linear-gradient(135deg, #4c1d95, #7c3aed); }

/* ════════════════════════════════
   SECTION TITLE STRIP
════════════════════════════════ */
.sec-title {
    font-size: 1.05rem; font-weight: 800;
    color: #1d4ed8 !important;
    border-left: 5px solid #1d4ed8;
    padding: 9px 16px; margin: 20px 0 14px;
    background: linear-gradient(90deg, #eff6ff, #f8faff, transparent);
    border-radius: 0 10px 10px 0;
    letter-spacing: 0.2px;
}

/* ════════════════════════════════
   CHAT BUBBLE
════════════════════════════════ */
.q-user {
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    color: #ffffff !important;
    border-radius: 20px 20px 4px 20px;
    padding: 12px 18px; margin: 10px 0 4px auto;
    max-width: 75%; width: fit-content;
    box-shadow: 0 4px 16px rgba(79,70,229,.35);
    float: right; clear: both;
    font-weight: 600; font-size: 0.95rem;
    border: 1px solid rgba(255,255,255,.2);
}

/* AI answer box */
.ans-box {
    background: #f8fafc !important;
    color: #1e293b !important;
    border-radius: 12px; padding: 18px 22px; margin: 8px 0 14px;
    font-size: 0.96rem; font-weight: 500; line-height: 1.8;
    box-shadow: 0 2px 12px rgba(0,0,0,.08);
    white-space: pre-wrap;
    border: 1.5px solid #e2e8f0;
}

/* Cell chip */
.cell-chip {
    background: #ecfdf5 !important;
    border-left: 4px solid #059669; border-radius: 7px;
    padding: 7px 14px; margin: 4px 0;
    font-family: 'Courier New', monospace; font-size: 0.82rem;
    color: #065f46 !important; font-weight: 700;
    box-shadow: 0 1px 4px rgba(5,150,105,.12);
}

.clearfix { clear: both; }

/* ════════════════════════════════
   SCROLLBAR
════════════════════════════════ */
::-webkit-scrollbar { width: 7px; height: 7px; }
::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #6366f1; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# DATAFRAME DISPLAY PATCH — forces dark text on white
# background regardless of CSS/theme issues.
# Applies pandas Styler cell-level properties which
# Streamlit 1.28+ passes directly to the canvas grid.
# ═══════════════════════════════════════════════════════
import pandas.io.formats.style as _pd_style

_CELL_PROPS = {
    "background-color": "#ffffff",
    "color": "#0f172a",
    "font-size": "13px",
    "font-weight": "500",
}
_HEADER_STYLES = [
    {"selector": "thead th",
     "props": [("background-color", "#1e3a8a"),
               ("color", "#ffffff"),
               ("font-weight", "800"),
               ("font-size", "13px"),
               ("border-bottom", "2px solid #1e40af")]},
    {"selector": "tbody tr:hover td",
     "props": [("background-color", "#f0f9ff")]},
]

def _styled(df):
    """Return a white-background, dark-text Styler for any DataFrame."""
    if isinstance(df, _pd_style.Styler):
        # Already styled (e.g. background_gradient) — just ensure text is dark
        try:
            return df.set_properties(**{"color": "#0f172a", "font-size": "13px"})
        except Exception:
            return df
    try:
        return (pd.DataFrame(df) if not isinstance(df, pd.DataFrame) else df
                ).style.set_properties(**_CELL_PROPS).set_table_styles(_HEADER_STYLES)
    except Exception:
        return df

_orig_st_dataframe = st.dataframe
def _patched_dataframe(data=None, *args, **kwargs):
    """Replacement for st.dataframe that always applies readable styling."""
    _orig_st_dataframe(_styled(data) if data is not None else data, *args, **kwargs)

st.dataframe = _patched_dataframe

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
# PATH & FILE HELPERS
# ═══════════════════════════════════════════════════════
def _app_dir():
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path(os.getcwd())

EXCEL_FOLDER = _app_dir() / "excel_files"


def _location_key(fname):
    """Derive a canonical location key from a filename."""
    n = os.path.basename(fname)
    n = re.sub(r"\.(xlsx?|xls)$", "", n, flags=re.I)
    n = re.sub(r"[Cc]ustomer.?[Aa]nd.?[Cc]apacity.?[Tt]racker.?", "", n)
    n = re.sub(r"[_\s]?\d{2}(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\d{2,4}.*$",
               "", n, flags=re.I)
    n = re.sub(r"_\d{10,}$", "", n)
    n = re.sub(r"__\d+_*$", "", n)
    n = re.sub(r"[_]+", " ", n).strip().lower()
    return n if n else fname


def location_from_name(fname):
    """Human-readable location label."""
    k = _location_key(fname)
    return k.replace("_", " ").title() if k else fname


def find_excel_files(folder):
    """Return deduplicated Excel files — one per location (newest numeric suffix)."""
    p = Path(folder)
    if not p.is_dir():
        return []
    all_files = [
        f.name for f in p.iterdir()
        if f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~")
    ]
    # Group by location key; keep the one with the highest numeric timestamp suffix
    groups: dict[str, str] = {}
    for fname in all_files:
        key = _location_key(fname)
        existing = groups.get(key)
        if existing is None:
            groups[key] = fname
        else:
            # Compare trailing numeric suffixes (higher = newer)
            def _ts(fn):
                m = re.search(r"_(\d{10,})\.", fn)
                return int(m.group(1)) if m else 0
            if _ts(fname) > _ts(existing):
                groups[key] = fname
    return sorted(groups.values())


# ═══════════════════════════════════════════════════════
# EXCEL READING  (openpyxl for .xlsx, xlrd for .xls)
# ═══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def _read_sheet_xlsx(path: str, sheet_name: str) -> pd.DataFrame:
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
        list(range(1, min(31, mr + 1))) +
        list(range(max(1, mr - 9), mr + 1))
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
def _read_sheet_xls(path: str, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name, header=None,
                       engine="xlrd", dtype=str)
    df = df.replace({"None": np.nan, "none": np.nan, "nan": np.nan})
    return df


@st.cache_data(show_spinner=False)
def load_file(original_path: str) -> dict[str, pd.DataFrame]:
    sheets: dict[str, pd.DataFrame] = {}
    is_xls = original_path.lower().endswith(".xls")
    if is_xls:
        try:
            import xlrd
            wb = xlrd.open_workbook(original_path)
            names = wb.sheet_names()
            for sh in names:
                try:
                    df = _read_sheet_xls(original_path, sh)
                    df = df.dropna(how="all").dropna(axis=1, how="all")
                    if not df.empty:
                        sheets[sh] = df
                except Exception:
                    pass
        except Exception as e:
            st.sidebar.warning(f"⚠️ {os.path.basename(original_path)}: {e}")
    else:
        try:
            from openpyxl import load_workbook
            wb = load_workbook(original_path, data_only=True)
            names = wb.sheetnames
            wb.close()
            for sh in names:
                try:
                    df = _read_sheet_xlsx(original_path, sh)
                    df = df.dropna(how="all").dropna(axis=1, how="all")
                    if not df.empty:
                        sheets[sh] = df
                except Exception:
                    pass
        except Exception as e:
            st.sidebar.warning(f"⚠️ {os.path.basename(original_path)}: {e}")
    return sheets


# ═══════════════════════════════════════════════════════
# HEADER DETECTION & SMART HEADER
# ═══════════════════════════════════════════════════════
def best_header_row(df: pd.DataFrame) -> int:
    best_row, best_score = 0, -1
    for i in range(min(8, len(df))):
        row = df.iloc[i].astype(str).str.strip()
        filled = (row.str.len() > 0) & (~row.isin(["nan", "None", ""]))
        label = filled & (~row.str.match(r"^-?\d+\.?\d*[eE]?[+-]?\d*$"))
        score = label.sum() * 2 + filled.sum()
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def smart_header(df: pd.DataFrame) -> pd.DataFrame:
    hr = best_header_row(df)
    hdr = df.iloc[hr].fillna("").astype(str).str.strip()
    seen: dict[str, int] = {}
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


def to_numeric(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        out[col] = pd.to_numeric(out[col], errors="ignore")
    return out


# ═══════════════════════════════════════════════════════
# CELL-LEVEL INDEXING  (Smart Query engine)
# ═══════════════════════════════════════════════════════
def _detect_all_header_rows(df: pd.DataFrame) -> set:
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


def _build_cell_col_map(df: pd.DataFrame):
    hr_set = _detect_all_header_rows(df)
    hr_maps: dict[int, dict[int, str]] = {}
    for hr in hr_set:
        m: dict[int, str] = {}
        for c in range(df.shape[1]):
            v = str(df.iat[hr, c]).strip()
            if v and v not in ("nan", "None"):
                m[c] = v
        hr_maps[hr] = m
    sorted_hrs = sorted(hr_set)
    cell_map: dict[tuple, str] = {}
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


@st.cache_data(show_spinner=False)
def index_sheet(df: pd.DataFrame):
    """Index every cell of a raw DataFrame for the Smart Query engine."""
    cell_map, hr_set = _build_cell_col_map(df)
    cells = []
    row_recs: dict[int, dict] = {}
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
            cells.append({"row": r, "col": c, "col_header": ch,
                           "value": v, "is_header": is_hdr})
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
# CROSS-FILE CORPUS  (for multi-location tab)
# ═══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def build_corpus(file_list: tuple, folder: str):
    corpus = []
    row_records: dict = {}
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
                    corpus.append({"file": fname, "location": loc, "sheet": sh,
                                   "row": r, "col": c, "col_header": ch,
                                   "value": v, "is_header": is_hdr})
                    if not is_hdr:
                        if key not in row_records:
                            row_records[key] = {}
                        row_records[key][ch] = v
    meta = {
        "total_cells": len(corpus),
        "total_files": len({x["file"] for x in corpus}),
        "total_sheets": len({(x["file"], x["sheet"]) for x in corpus}),
        "total_rows": len(row_records),
        "locations": sorted({x["location"] for x in corpus}),
    }
    return corpus, row_records, meta


# ═══════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════
st.sidebar.image("https://img.icons8.com/fluency/96/data-center.png", width=70)
st.sidebar.title("🏢 Capacity Tracker")
st.sidebar.markdown("---")
st.sidebar.subheader("📁 Data Source")

uploaded_files = st.sidebar.file_uploader(
    "Upload Excel files (optional)", type=["xlsx", "xls"], accept_multiple_files=True
)

if uploaded_files:
    import tempfile
    tmp_dir = tempfile.mkdtemp()
    for uf in uploaded_files:
        with open(os.path.join(tmp_dir, uf.name), "wb") as fh:
            fh.write(uf.read())
    data_dir = tmp_dir
else:
    data_dir = str(EXCEL_FOLDER)

excel_files = find_excel_files(data_dir)

if not excel_files:
    st.error("### ⚠️ No Excel files found\n\nPlace files in `attached_assets/` or upload via sidebar.")
    st.stop()

loc_map = {f: location_from_name(f) for f in excel_files}
st.sidebar.success(f"✅ {len(excel_files)} location(s) loaded")

st.sidebar.subheader("🏙️ Location")
selected_file = st.sidebar.selectbox("Location", excel_files,
                                     format_func=lambda x: loc_map[x])

all_sheets = load_file(os.path.join(data_dir, selected_file))
if not all_sheets:
    st.error(f"Could not read any sheets from **{selected_file}**.")
    st.stop()

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

# ── Index selected sheet ──
with st.spinner("🔍 Indexing sheet cells…"):
    sq_cells, sq_rows, sq_meta = index_sheet(raw_df)

# ── Build corpus ──
with st.spinner("🔍 Building cross-file index…"):
    corpus, row_records, meta = build_corpus(tuple(excel_files), data_dir)

loc_label = loc_map[selected_file]

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
    "💬 AI Smart Query",
])

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
    c5.markdown(f'<div class="kcard kcard-teal"><h2>{sq_meta["total_cells"]:,}</h2><p>Sheet Cells</p></div>', unsafe_allow_html=True)
    c6.markdown(f'<div class="kcard kcard-red"><h2>{int(df_clean.isna().sum().sum())}</h2><p>Missing</p></div>', unsafe_allow_html=True)

    st.markdown("---")
    if num_cols:
        st.markdown('<div class="sec-title">📐 Quick Statistics</div>', unsafe_allow_html=True)
        stats_df = df_clean[num_cols].describe().T
        stats_df["range"] = stats_df["max"] - stats_df["min"]
        st.dataframe(
            stats_df.style.format("{:.3f}", na_rep="—").background_gradient(cmap="Blues", subset=["mean", "max"]),
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
            str(df_clean[c].dropna().iloc[0])[:55] if df_clean[c].dropna().shape[0] > 0 else "—"
            for c in df_clean.columns
        ],
    })
    st.dataframe(ci, use_container_width=True)

    st.markdown("---")
    st.markdown('<div class="sec-title">📄 Complete Raw Sheet (all rows · all columns)</div>', unsafe_allow_html=True)
    st.caption(f"Showing all {raw_df.shape[0]} rows × {raw_df.shape[1]} columns from the Excel sheet.")
    st.dataframe(raw_df, use_container_width=True, height=500)
    st.download_button(
        "⬇️ Download Complete Sheet CSV",
        raw_df.to_csv(index=False).encode(),
        f"{loc_label}_{selected_sheet}_complete.csv", "text/csv", key="ov_dl",
    )


# ═══════════════════════════════════════════════════════
# TAB 1 – RAW DATA
# ═══════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("📋 Data Table")
    srch = st.text_input("🔍 Live search", "", key="rawsrch")
    disp = (
        df_clean[df_clean.apply(
            lambda col: col.astype(str).str.contains(srch, case=False, na=False)
        ).any(axis=1)] if srch else df_clean
    )
    st.caption(f"Showing {len(disp):,} / {len(df_clean):,} rows")
    st.dataframe(disp, use_container_width=True, height=500)
    st.download_button("⬇️ CSV", disp.to_csv(index=False).encode(), "export.csv", "text/csv")
    st.markdown("---")
    st.subheader("🗃️ Raw Excel (unprocessed)")
    st.dataframe(raw_df, use_container_width=True, height=280)


# ═══════════════════════════════════════════════════════
# TAB 2 – ANALYTICS
# ═══════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("📊 Column Analytics")
    if not num_cols:
        st.info("No numeric columns detected.")
    else:
        chosen = st.multiselect("Select columns", num_cols, default=num_cols[:min(6, len(num_cols))])
        if chosen:
            sub = df_clean[chosen].dropna(how="all")
            kc = st.columns(min(len(chosen), 6))
            for i, col in enumerate(chosen[:6]):
                s = sub[col].dropna()
                if len(s):
                    kc[i].metric(col[:20], f"{s.sum():,.1f}", f"avg {s.mean():,.1f}")
            st.markdown("---")
            agg_rows = []
            grand = df_clean[chosen].select_dtypes("number").sum().sum()
            for col in chosen:
                s = df_clean[col].dropna()
                if len(s) and pd.api.types.is_numeric_dtype(s):
                    agg_rows.append({
                        "Column": col, "Count": int(s.count()),
                        "Sum": s.sum(), "Mean": s.mean(), "Median": s.median(),
                        "Min": s.min(), "Max": s.max(), "Std": s.std(),
                        "% Total": f"{s.sum()/grand*100:.1f}%" if grand else "—",
                    })
            if agg_rows:
                adf = pd.DataFrame(agg_rows).set_index("Column")
                st.dataframe(
                    adf.style.format("{:,.2f}", na_rep="—",
                                     subset=[c for c in adf.columns if c != "% Total"])
                              .background_gradient(cmap="YlOrRd", subset=["Sum", "Max"]),
                    use_container_width=True,
                )
        st.markdown("---")
        st.markdown('<div class="sec-title">🧮 Group-By</div>', unsafe_allow_html=True)
        all_cat_2 = [c for c in df_clean.columns if c not in num_cols and df_clean[c].nunique() < 60]
        if all_cat_2 and num_cols:
            gc1, gc2, gc3 = st.columns(3)
            gc = gc1.selectbox("Group by", all_cat_2)
            ac = gc2.selectbox("Aggregate", num_cols)
            af = gc3.selectbox("Function", ["sum", "mean", "count", "min", "max", "median"])
            grp = (df_clean.groupby(gc)[ac].agg(af).reset_index()
                   .rename(columns={ac: f"{af}({ac})"})
                   .sort_values(f"{af}({ac})", ascending=False))
            st.dataframe(grp, use_container_width=True)
            fig = px.bar(grp, x=gc, y=f"{af}({ac})", color=f"{af}({ac})",
                         color_continuous_scale="Viridis",
                         title=f"{af.title()} of {ac} by {gc}")
            fig.update_layout(xaxis_tickangle=-35, height=400)
            st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 3 – CHARTS
# ═══════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("📈 Interactive Charts")
    ctype = st.selectbox("Chart Type", [
        "Bar Chart", "Grouped Bar", "Line Chart", "Scatter Plot",
        "Area Chart", "Bubble Chart", "Heatmap (Correlation)",
        "Box Plot", "Funnel Chart", "Waterfall / Cumulative", "3-D Scatter",
    ])
    if not num_cols:
        st.info("No numeric columns.")
    else:
        def _s(label, opts, idx=0, key=None):
            return st.selectbox(label, opts, index=min(idx, max(0, len(opts) - 1)), key=key)

        if ctype == "Bar Chart":
            xc = _s("X", cat_cols or df_clean.columns.tolist(), key="bx")
            yc = _s("Y", num_cols, key="by")
            ori = st.radio("Orientation", ["Vertical", "Horizontal"], horizontal=True)
            d = df_clean[[xc, yc]].dropna()
            fig = px.bar(d, x=xc if ori == "Vertical" else yc,
                         y=yc if ori == "Vertical" else xc,
                         color=yc, color_continuous_scale="Turbo",
                         orientation="v" if ori == "Vertical" else "h",
                         title=f"{yc} by {xc}")
            fig.update_layout(height=480)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Grouped Bar":
            xc = _s("X", cat_cols or df_clean.columns.tolist(), key="gbx")
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:3])
            if ycs:
                fig = px.bar(df_clean[[xc] + ycs].dropna(subset=ycs, how="all"),
                             x=xc, y=ycs, barmode="group")
                fig.update_layout(height=460)
                st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Line Chart":
            xc = _s("X", df_clean.columns.tolist(), key="lx")
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:2])
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
            fig = px.scatter(d, x=xc, y=yc,
                             size=sc if sc != "None" else None,
                             color=cc if cc != "None" else None,
                             color_continuous_scale="Rainbow")
            fig.update_layout(height=480)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Area Chart":
            xc = _s("X", df_clean.columns.tolist(), key="ax")
            ycs = st.multiselect("Y columns", num_cols, default=num_cols[:3])
            if ycs:
                fig = px.area(df_clean[[xc] + ycs].dropna(subset=ycs, how="all"),
                              x=xc, y=ycs)
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
                    d = d.copy()
                    d[lc] = df_clean[lc]
                fig = px.scatter(d, x=xc, y=yc, size=sz,
                                 color=lc if lc != "None" else None, size_max=65)
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need ≥ 3 numeric columns.")

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
                         color=xc if xc != "None" else None, points="outliers")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)

        elif ctype == "Funnel Chart":
            xc = _s("Stage", cat_cols or df_clean.columns.tolist(), key="fn_x")
            yc = _s("Value", num_cols, key="fn_y")
            d = (df_clean[[xc, yc]].dropna().groupby(xc)[yc]
                 .sum().reset_index().sort_values(yc, ascending=False))
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
                    d = d.copy()
                    d[cc] = df_clean[cc]
                fig = px.scatter_3d(d, x=xc, y=yc, z=zc, color=cc if cc != "None" else None)
                fig.update_layout(height=550)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Need ≥ 3 numeric columns.")


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
                fig = px.pie(pd_, names=pc, values=pv, hole=0.38,
                             color_discrete_sequence=px.colors.qualitative.Vivid)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No suitable category columns for pie chart.")
        with r2:
            st.markdown('<div class="sec-title">📊 Histogram</div>', unsafe_allow_html=True)
            hc = st.selectbox("Column", num_cols, key="hcol")
            bins = st.slider("Bins", 5, 100, 25)
            fig = px.histogram(df_clean[hc].dropna(), nbins=bins,
                               color_discrete_sequence=["#17a572"])
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown('<div class="sec-title">🎻 Violin</div>', unsafe_allow_html=True)
        vc = st.selectbox("Column", num_cols, key="vc")
        fig = px.violin(df_clean[vc].dropna(), y=vc, box=True, points="outliers",
                        color_discrete_sequence=["#c0392b"])
        st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 5 – QUERY ENGINE
# ═══════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("🔍 Query Engine  (selected sheet)")
    query = st.text_input("Question",
                          placeholder="e.g. Total subscription / Max capacity / List customers")

    def run_query(q, df, nc):
        ql = q.lower()
        res = []
        if any(w in ql for w in ["sum", "total"]):
            for c in nc:
                if c.lower() in ql or "all" in ql or len(nc) == 1:
                    res.append(f"**SUM `{c}`** = {df[c].sum():,.4f}")
        if any(w in ql for w in ["average", "mean", "avg"]):
            for c in nc:
                if c.lower() in ql or len(nc) == 1:
                    res.append(f"**MEAN `{c}`** = {df[c].mean():,.4f}")
        if any(w in ql for w in ["maximum", "highest", "max"]):
            for c in nc:
                if c.lower() in ql or len(nc) == 1:
                    res.append(f"**MAX `{c}`** = {df[c].max():,.4f}")
        if any(w in ql for w in ["minimum", "lowest", "min"]):
            for c in nc:
                if c.lower() in ql or len(nc) == 1:
                    res.append(f"**MIN `{c}`** = {df[c].min():,.4f}")
        if any(w in ql for w in ["count", "how many"]):
            res.append(f"**Rows** = {len(df):,}")
        if any(w in ql for w in ["customer", "list", "show", "name"]):
            for c in df.columns:
                if "customer" in c.lower() or "name" in c.lower():
                    nm = df[c].dropna().unique()
                    res.append(f"**`{c}`** ({len(nm)}):\n" +
                               "\n".join(f"  • {n}" for n in nm[:30]))
                    break
        if not res:
            res.append("ℹ️ Try: **sum / average / max / min / count / list**")
        return "\n\n".join(res)

    if query:
        st.markdown(run_query(query, df_clean, num_cols))

    st.markdown("---")
    st.markdown('<div class="sec-title">🧮 Manual Compute</div>', unsafe_allow_html=True)
    if num_cols:
        mc1, mc2, mc3 = st.columns(3)
        op = mc1.selectbox("Op", ["Sum", "Mean", "Max", "Min", "Count", "Median", "Std Dev", "Range"])
        sc_col = mc2.selectbox("Column", num_cols, key="mc_col")
        fc = mc3.selectbox("Filter by", ["None"] + [c for c in df_clean.columns if c not in num_cols])
        fv = None
        if fc != "None":
            fv = st.selectbox("Filter value", df_clean[fc].dropna().unique().tolist())
        ds = df_clean.copy()
        if fc != "None" and fv is not None:
            ds = ds[ds[fc] == fv]
        s = ds[sc_col].dropna()
        ops = {"Sum": s.sum(), "Mean": s.mean(), "Max": s.max(), "Min": s.min(),
               "Count": s.count(), "Median": s.median(), "Std Dev": s.std(),
               "Range": s.max() - s.min()}
        r_val = ops.get(op, "N/A")
        if isinstance(r_val, float):
            r_val = f"{r_val:,.4f}"
        st.success(f"**{op}** of `{sc_col}`{f' (where {fc}={fv})' if fv else ''} → **{r_val}**")


# ═══════════════════════════════════════════════════════
# TAB 6 – MULTI-LOCATION
# ═══════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("🌍 Cross-Location Comparison")

    @st.cache_data(show_spinner=False)
    def load_all_summ(files: tuple, folder: str):
        summ = {}
        for f in files:
            shd = load_file(os.path.join(folder, f))
            lbl = location_from_name(f)
            for sh, raw in shd.items():
                dfc = to_numeric(smart_header(raw))
                nc = dfc.select_dtypes(include="number").columns.tolist()
                if nc:
                    summ[f"{lbl} | {sh}"] = {"df": dfc, "num_cols": nc, "file": f, "sheet": sh}
        return summ

    all_summ = load_all_summ(tuple(excel_files), data_dir)
    if all_summ:
        all_num_cols = sorted({c for v in all_summ.values() for c in v["num_cols"]})
        if all_num_cols:
            comp_col = st.selectbox("Compare by column", all_num_cols)
            rows_cmp = []
            for lbl, info in all_summ.items():
                if comp_col in info["num_cols"]:
                    s = info["df"][comp_col].dropna()
                    rows_cmp.append({"Location|Sheet": lbl, "Sum": s.sum(),
                                     "Mean": s.mean(), "Max": s.max(),
                                     "Min": s.min(), "Count": s.count()})
            if rows_cmp:
                cmp = pd.DataFrame(rows_cmp).set_index("Location|Sheet")
                st.dataframe(
                    cmp.style.format("{:,.2f}").background_gradient(cmap="YlOrRd"),
                    use_container_width=True,
                )
                fig = px.bar(cmp.reset_index(), x="Location|Sheet", y="Sum", color="Sum",
                             color_continuous_scale="Viridis",
                             title=f"Sum of '{comp_col}' across locations")
                fig.update_layout(xaxis_tickangle=-30, height=440)
                st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No numeric data found across files.")


# ═══════════════════════════════════════════════════════
# TAB 7 – AI AGENT
# ═══════════════════════════════════════════════════════
with tabs[7]:
    st.subheader("🤖 AI Agent – Automated Insights")
    if st.button("🚀 Run Full Analysis", type="primary"):
        with st.spinner("Analysing all files…"):
            for lbl, info in list(all_summ.items())[:12]:
                dfa = info["df"]
                nc = info["num_cols"]
                if not nc:
                    continue
                with st.expander(f"📍 {lbl}", expanded=False):
                    ca, cb = st.columns(2)
                    with ca:
                        for col in nc[:5]:
                            s = dfa[col].dropna()
                            if len(s):
                                st.metric(col[:26], f"{s.sum():,.1f}", f"avg {s.mean():,.1f}")
                    with cb:
                        for col in nc[:4]:
                            s = dfa[col].dropna()
                            if len(s) > 3:
                                z = (s - s.mean()) / s.std()
                                o = z[z.abs() > 2.5]
                                (st.warning if len(o) else st.success)(
                                    f"`{col}`: {len(o)} outlier(s)" if len(o) else f"`{col}`: Clean ✓"
                                )

    st.markdown("---")
    st.markdown('<div class="sec-title">📁 Files Summary</div>', unsafe_allow_html=True)
    fsm = []
    for f in excel_files:
        shd = load_file(os.path.join(data_dir, f))
        total_r = sum(smart_header(raw).shape[0] for raw in shd.values())
        fsm.append({"Location": loc_map[f], "File": f, "Sheets": len(shd), "Data Rows": total_r})
    st.dataframe(pd.DataFrame(fsm), use_container_width=True)


# ═══════════════════════════════════════════════════════
# TAB 8 – AI SMART QUERY
# Cell-level — no formulas, direct match from raw data
# ═══════════════════════════════════════════════════════
with tabs[8]:
    st.markdown("## 💬 AI Smart Query")
    st.markdown(
        f"Querying: **{loc_label}** › **{selected_sheet}** — change location/sheet in the sidebar."
    )

    qi1, qi2, qi3 = st.columns(3)
    qi1.markdown(
        f'<div class="kcard kcard-blue"><h2>{sq_meta["total_cells"]:,}</h2>'
        f'<p>Cells in Sheet</p></div>', unsafe_allow_html=True,
    )
    qi2.markdown(
        f'<div class="kcard kcard-green"><h2>{sq_meta["total_data"]:,}</h2>'
        f'<p>Data Cells</p></div>', unsafe_allow_html=True,
    )
    qi3.markdown(
        f'<div class="kcard kcard-purple"><h2>{sq_meta["total_rows"]:,}</h2>'
        f'<p>Data Rows</p></div>', unsafe_allow_html=True,
    )

    # ── synonyms & helpers ──
    _OP_VERBS = {
        "total","sum","avg","mean","max","min","count","list","find",
        "show","all","average","maximum","minimum","highest","lowest",
        "top","bottom","describe","statistics","stats","summary","unique",
        "distinct","sheet","column","row","missing","null","percent",
        "percentage","ratio","share","number","across","compare",
    }

    _SYN = {
        "subscription": ["subscription","subscribed","subscript"],
        "capacity":     ["capacity","capac"],
        "power":        ["power","kw","kva"],
        "usage":        ["usage","utilization","consumption","consumed"],
        "rack":         ["rack","racks"],
        "space":        ["space","sqft","sq ft"],
        "customer":     ["customer","name","customers","client","clients"],
        "billing":      ["billing","bill","invoice"],
        "ownership":    ["ownership","owned","owner"],
        "revenue":      ["revenue","income","earning"],
    }

    def _is_num(v):
        try:
            float(v); return True
        except Exception:
            return False

    def _mcol(kw: str, hdr: str) -> bool:
        hl, kwl = hdr.lower(), kw.lower()
        if kwl in hl:
            return True
        for key, syns in _SYN.items():
            if kwl in syns or kwl == key:
                for s in syns:
                    if s in hl:
                        return True
        return False

    def sheet_query(question: str) -> dict:
        q  = question.strip()
        ql = q.lower()
        sig = [w for w in re.findall(r"[a-z0-9]{2,}", ql) if w not in _SW]

        f_sum  = any(x in ql for x in ["total","sum","aggregate"])
        f_avg  = any(x in ql for x in ["average","mean","avg"])
        f_max  = any(x in ql for x in ["maximum","highest","largest","max"])
        f_min  = any(x in ql for x in ["minimum","lowest","smallest","min"])
        f_cnt  = any(x in ql for x in ["count","how many","number of"])
        f_stat = any(x in ql for x in ["statistics","stats","describe","summary","all stats"])
        f_uniq = any(x in ql for x in ["unique","distinct","different"])
        f_miss = any(x in ql for x in ["missing","null","blank","empty"])
        f_cols = any(x in ql for x in ["column","columns","field","header","headers"])
        f_topn = re.search(r"\btop\s*(\d+)\b", ql)
        f_botn = re.search(r"\bbottom\s*(\d+)\b", ql)
        f_num  = f_sum or f_avg or f_max or f_min or f_cnt or f_stat or bool(f_topn) or bool(f_botn)
        f_cust = any(x in ql for x in ["customer","customers","client","clients","name","names"])
        f_list = any(x in ql for x in ["list","show","display"])

        out = {"answer":"","table":None,"chart_df":None,"chart_cfg":None,
               "cell_hits":[],"sub_tables":[]}
        wc = sq_cells
        rr = sq_rows

        if not wc:
            out["answer"] = "❓ No data indexed in this sheet."
            return out

        col_kws = [w for w in sig if w not in _OP_VERBS]

        # ── numeric helpers ──
        def npkw(kw):
            res = []
            for cell in wc:
                if cell["is_header"]: continue
                if _mcol(kw, cell["col_header"]):
                    try:
                        res.append((float(cell["value"]), cell))
                    except Exception:
                        pass
            return res

        def npbest(kws):
            bk, bp = None, []
            for w in kws:
                p = npkw(w)
                if len(p) > len(bp):
                    bp = p; bk = w
            return bk, bp

        def build_rows_df(row_nums):
            recs = []
            for rn in sorted(row_nums):
                rec = rr.get(rn, {})
                if rec:
                    rd = {"Row #": rn + 1}
                    rd.update(rec)
                    recs.append(rd)
            return pd.DataFrame(recs) if recs else pd.DataFrame()

        # ── INTENT: List customers / names ──
        if (f_cust or f_list) and not f_num:
            found = []
            for cell in wc:
                if cell["is_header"]: continue
                ch = cell["col_header"].lower()
                if any(x in ch for x in ["customer","name","client","company"]):
                    found.append({"Row #": cell["row"]+1,
                                  "Column": cell["col_header"],
                                  "Value": cell["value"]})
            if found:
                tbl = pd.DataFrame(found).drop_duplicates(subset=["Value"])
                out["answer"] = f"Found **{len(tbl)}** customer/name entries."
                out["table"] = tbl
                name_rows = sorted({r["Row #"]-1 for r in found})
                full_df = build_rows_df(name_rows)
                if not full_df.empty:
                    out["sub_tables"].append({
                        "label": f"📋 Full Row Data ({len(name_rows)} rows)",
                        "df": full_df,
                    })
            else:
                out["answer"] = ("No 'Customer/Name/Client' columns found in this sheet.\n"
                                 "Try: *Find CISCO* or *Show all columns*")
            return out

        # ── INTENT: Missing values ──
        if f_miss:
            dfc_miss = to_numeric(smart_header(raw_df))
            mr_list = []
            for col in dfc_miss.columns:
                mc_cnt = int(dfc_miss[col].isna().sum())
                if mc_cnt > 0:
                    mr_list.append({"Column": col, "Missing": mc_cnt,
                                    "Missing%": f"{mc_cnt/max(len(dfc_miss),1)*100:.1f}%"})
            if mr_list:
                tbl = pd.DataFrame(mr_list).sort_values("Missing", ascending=False)
                out["answer"] = f"Found **{len(tbl)}** column(s) with missing values."
                out["table"] = tbl
            else:
                out["answer"] = "✅ No missing values found in this sheet."
            return out

        # ── INTENT: Column listing ──
        if f_cols and not f_num:
            seen_cols = set()
            cr = []
            for cell in wc:
                if not cell["is_header"]: continue
                ch = cell["value"].strip()
                if ch in ("", "nan"): continue
                if ch not in seen_cols:
                    seen_cols.add(ch)
                    cr.append({"Column": ch, "At Row": cell["row"]+1, "At Col": cell["col"]+1})
            # also collect col_headers from data cells
            for cell in wc:
                if cell["is_header"]: continue
                ch = cell["col_header"]
                if ch not in seen_cols and ch and not ch.startswith("Col_"):
                    seen_cols.add(ch)
                    cr.append({"Column": ch, "At Row": "data", "At Col": cell["col"]+1})
            tbl = pd.DataFrame(cr) if cr else pd.DataFrame()
            out["answer"] = f"Found **{len(tbl)}** column header(s) in this sheet."
            out["table"] = tbl
            return out

        # ── INTENT: Numeric aggregation ──
        if f_num:
            kw, pairs = npbest(col_kws)
            # fallback: try domain keywords from query
            if not pairs:
                for dkw in ["subscription","capacity","power","usage","rack",
                            "space","consumption","kw","kva","sqft"]:
                    if dkw in ql:
                        pairs = npkw(dkw)
                        if pairs:
                            kw = dkw
                            break

            if pairs:
                vals = [v for v, _ in pairs]
                sa = pd.Series(vals)
                parts = []
                if f_sum or f_stat:
                    parts.append(f"**Total (Sum):** {sa.sum():,.4f}")
                if f_avg or f_stat:
                    parts.append(f"**Average:** {sa.mean():,.4f}")
                if f_max or f_stat:
                    parts.append(f"**Maximum:** {sa.max():,.4f}")
                if f_min or f_stat:
                    parts.append(f"**Minimum:** {sa.min():,.4f}")
                if f_cnt or f_stat:
                    parts.append(f"**Count:** {sa.count():,}")
                if f_stat:
                    parts.append(f"**Median:** {sa.median():,.4f}  |  **Std Dev:** {sa.std():,.4f}")
                if (f_topn or f_botn) and not (f_sum or f_avg or f_max or f_min or f_cnt or f_stat):
                    parts.append(f"**Count:** {sa.count():,}")
                    parts.append(f"**Total (Sum):** {sa.sum():,.4f}")

                detail = [{"Row #": c["row"]+1, "Column": c["col_header"], "Value": v}
                          for v, c in pairs]
                tbl = pd.DataFrame(detail).sort_values("Value", ascending=False)
                out["answer"] = (f"Results for **'{kw}'** ({len(vals):,} values):\n\n"
                                 + "\n".join(parts))
                out["table"] = tbl

                if f_topn:
                    n = int(f_topn.group(1))
                    top = sorted(pairs, key=lambda x: x[0], reverse=True)[:n]
                    top_rows = build_rows_df(sorted({c["row"] for _, c in top}))
                    out["sub_tables"].append({
                        "label": f"🏆 Top {n} — {kw}",
                        "df": top_rows if not top_rows.empty else pd.DataFrame(
                            [{"Row": c["row"]+1, "Column": c["col_header"], "Value": v}
                             for v, c in top]),
                    })
                if f_botn:
                    n = int(f_botn.group(1))
                    bot = sorted(pairs, key=lambda x: x[0])[:n]
                    bot_rows = build_rows_df(sorted({c["row"] for _, c in bot}))
                    out["sub_tables"].append({
                        "label": f"🔻 Bottom {n} — {kw}",
                        "df": bot_rows if not bot_rows.empty else pd.DataFrame(
                            [{"Row": c["row"]+1, "Column": c["col_header"], "Value": v}
                             for v, c in bot]),
                    })
                return out

            # fallback: all numeric cells
            anums = [(float(c["value"]), c) for c in wc
                     if not c["is_header"] and _is_num(c["value"])]
            if anums:
                vals = [v for v, _ in anums]
                sa = pd.Series(vals)
                parts = []
                if f_sum:  parts.append(f"**Sum ALL numeric:** {sa.sum():,.4f}")
                if f_avg:  parts.append(f"**Avg ALL numeric:** {sa.mean():,.4f}")
                if f_max:  parts.append(f"**Max ALL numeric:** {sa.max():,.4f}")
                if f_min:  parts.append(f"**Min ALL numeric:** {sa.min():,.4f}")
                if f_cnt:  parts.append(f"**Count ALL numeric:** {sa.count():,}")
                out["answer"] = ("No column matched keywords — results from ALL numeric cells:\n\n"
                                 + "\n".join(parts))
                return out

        # ── INTENT: Unique values ──
        if f_uniq:
            for w in col_kws or sig:
                uv, sr = set(), []
                for cell in wc:
                    if cell["is_header"]: continue
                    if _mcol(w, cell["col_header"]):
                        uv.add(cell["value"])
                        sr.append({"Column": cell["col_header"],
                                   "Value": cell["value"], "Row #": cell["row"]+1})
                if uv:
                    tbl = pd.DataFrame(sr).drop_duplicates(subset=["Value"]) if sr else pd.DataFrame()
                    out["answer"] = f"**{len(uv)}** unique value(s) for **'{w}'**."
                    out["table"] = tbl
                    return out

        # ── INTENT: Free-text entity / keyword search ──
        if sig:
            quoted = re.findall(r'"([^"]+)"', q)
            if quoted:
                terms = [quoted[0].lower()]
            else:
                terms = [w for w in sig if any(w in cell["value"].lower() for cell in wc)]
                if not terms:
                    terms = sig

            hit_cells = [
                cell for cell in wc
                if not cell["is_header"]
                and any(t in cell["value"].lower() for t in terms)
            ]
            hit_rows = sorted({c["row"] for c in hit_cells})
            full_df = build_rows_df(hit_rows)
            cell_list = [
                {"Row #": c["row"]+1, "Col #": c["col"]+1,
                 "Column Header": c["col_header"], "Value": c["value"]}
                for c in hit_cells[:100]
            ]
            out["answer"] = (
                f"Found **{len(hit_cells):,}** cell(s) matching "
                f"**'{', '.join(terms[:4])}'** in **{len(hit_rows):,}** row(s)."
            )
            out["table"] = full_df if not full_df.empty else None
            out["cell_hits"] = cell_list
            if not full_df.empty:
                for col in full_df.columns:
                    if any(x in col.lower() for x in ["customer","name","client"]):
                        cdf = full_df[["Row #", col]].drop_duplicates()
                        out["sub_tables"].append({"label": f"👤 Names ({len(cdf)})", "df": cdf})
                        break
            return out

        out["answer"] = (
            "❓ No match found.\n\n"
            "**Try:** *List all customers* | *Find CISCO* | *Total subscription* | "
            "*Top 10 subscription* | *Show columns* | *Statistics of capacity*"
        )
        return out

    # ── Render answer ──
    def render_answer(res: dict, tidx: int = 0):
        st.markdown(
            f'<div class="ans-box">{res["answer"]}</div>', unsafe_allow_html=True
        )
        if res.get("table") is not None and not res["table"].empty:
            tbl = res["table"].reset_index(drop=True)
            st.dataframe(tbl, use_container_width=True,
                         height=min(520, 48 + len(tbl) * 36),
                         key=f"tbl_{tidx}")
            st.download_button("⬇️ CSV", tbl.to_csv(index=False).encode(),
                               "result.csv", "text/csv", key=f"dl_{tidx}")
        if res.get("chart_cfg") and res.get("chart_df") is not None:
            cfg = res["chart_cfg"]
            cdf = res["chart_df"]
            if cfg["x"] in cdf.columns and cfg["y"] in cdf.columns:
                fig = px.bar(cdf.sort_values(cfg["y"], ascending=False).head(30),
                             x=cfg["x"], y=cfg["y"], color=cfg["y"],
                             color_continuous_scale="Viridis",
                             title=cfg["title"], height=400)
                fig.update_layout(xaxis_tickangle=-30)
                st.plotly_chart(fig, use_container_width=True, key=f"ch_{tidx}")
        for si, s in enumerate(res.get("sub_tables", [])):
            with st.expander(s["label"], expanded=True):
                st.dataframe(s["df"], use_container_width=True, key=f"sub_{tidx}_{si}")
        if res.get("cell_hits"):
            with st.expander(f"🔬 Cell matches ({len(res['cell_hits'])})", expanded=False):
                for ch in res["cell_hits"]:
                    st.markdown(
                        f'<div class="cell-chip">'
                        f'R{ch["Row #"]} C{ch["Col #"]} | '
                        f'<i>{ch["Column Header"]}</i> | '
                        f'<b>{ch["Value"]}</b></div>',
                        unsafe_allow_html=True,
                    )
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)

    # ── Chat history ──
    hist_key = f"sq_hist_{selected_file}_{selected_sheet}"
    if hist_key not in st.session_state:
        st.session_state[hist_key] = []

    for tidx, turn in enumerate(st.session_state[hist_key]):
        st.markdown(f'<div class="q-user">🧑 {turn["q"]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="clearfix"></div>', unsafe_allow_html=True)
        render_answer(turn["res"], tidx)
        st.markdown("---")

    # ── Input bar ──
    st.markdown("---")
    ic, bc, cc = st.columns([8, 1, 1])
    with ic:
        user_q = st.text_input(
            "Ask:", key="sq_input",
            placeholder="Find CISCO | List customers | Total subscription | Top 10 subscription",
            label_visibility="collapsed",
        )
    with bc:
        ask_btn = st.button("🔍 Ask", use_container_width=True, type="primary")
    with cc:
        if st.button("🗑️ Clear", use_container_width=True):
            st.session_state[hist_key] = []
            st.rerun()

    # ── Example chips ──
    st.markdown("**💡 Quick Examples — click any:**")
    examples = [
        ["List all customers", "Find CISCO", "Find AT&T", "Find Axis Bank", "Find TCS", "Find Wipro"],
        ["Total subscription", "Max capacity", "Average power",
         "Top 10 subscription", "Statistics of subscription", "Count rows"],
        ["Show columns", "Show missing values", "Unique billing models", "Unique ownership"],
    ]
    for row in examples:
        ex_cols = st.columns(len(row))
        for j, ex in enumerate(row):
            if ex_cols[j].button(ex, key=f"sqchip_{ex}", use_container_width=True):
                user_q = ex
                ask_btn = True

    if ask_btn and user_q.strip():
        with st.spinner(f"Scanning sheet for: **{user_q}** …"):
            answer = sheet_query(user_q)
        st.session_state[hist_key].append({"q": user_q, "res": answer})
        st.rerun()


# ═══════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════
st.markdown("---")
st.caption(
    f"Sify DC · Capacity Tracker · {meta['total_cells']:,} cells indexed · "
    f"{', '.join(meta['locations'])}"
)
