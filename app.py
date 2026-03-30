import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import re
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="Customer & Capacity Tracker Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_DIR = "attached_assets"

LOCATION_MAP = {
    "airoli": "Airoli",
    "bangalore": "Bangalore",
    "chennai": "Chennai",
    "kolkata": "Kolkata",
    "noida_01": "Noida 01",
    "noida_02": "Noida 02",
    "noida01": "Noida 01",
    "noida02": "Noida 02",
    "rabale_t1_t2": "Rabale T1-T2",
    "rabale_tower_4": "Rabale Tower 4",
    "rabale_tower_5": "Rabale Tower 5",
    "vashi": "Vashi"
}

@st.cache_data(show_spinner=False)
def load_all_excel_files():
    all_data = {}
    excel_files = []
    for f in os.listdir(EXCEL_DIR):
        if f.endswith(".xlsx") or f.endswith(".xls"):
            excel_files.append(f)

    for filename in excel_files:
        filepath = os.path.join(EXCEL_DIR, filename)
        location_key = extract_location_from_filename(filename)
        try:
            engine = "xlrd" if filename.endswith(".xls") else "openpyxl"
            xl = pd.ExcelFile(filepath, engine=engine)
            sheets_data = {}
            for sheet_name in xl.sheet_names:
                try:
                    df_raw = pd.read_excel(
                        filepath,
                        sheet_name=sheet_name,
                        header=None,
                        engine=engine
                    )
                    if df_raw.empty:
                        continue
                    df_raw = df_raw.dropna(how="all").dropna(axis=1, how="all")
                    if df_raw.empty:
                        continue
                    sheets_data[sheet_name] = {
                        "raw": df_raw,
                        "structured": try_structure_dataframe(df_raw)
                    }
                except Exception:
                    pass
            if sheets_data:
                all_data[location_key] = {
                    "filename": filename,
                    "sheets": sheets_data
                }
        except Exception:
            pass
    return all_data

def extract_location_from_filename(filename):
    name = filename.lower()
    name = re.sub(r"customer_and_capacity_tracker_", "", name)
    name = re.sub(r"_\d{8}_\d+", "", name)
    name = re.sub(r"_\d{10,}", "", name)
    name = re.sub(r"\.(xlsx|xls)$", "", name)
    name = re.sub(r"_15(mar|feb|jan)26", "", name, flags=re.IGNORECASE)
    name = re.sub(r"_\(2\)", "", name)
    name = name.strip("_").strip()
    name_title = name.replace("_", " ").title()
    return name_title

def try_structure_dataframe(df_raw):
    for header_row in range(min(10, len(df_raw))):
        row = df_raw.iloc[header_row]
        non_null = row.count()
        if non_null >= max(2, len(df_raw.columns) * 0.3):
            header = [str(c).strip() if pd.notna(c) and str(c).strip() != "nan" else f"Col_{i}" for i, c in enumerate(row)]
            data_rows = df_raw.iloc[header_row + 1:].copy()
            data_rows.columns = header
            data_rows = data_rows.dropna(how="all")
            data_rows.reset_index(drop=True, inplace=True)
            return data_rows
    df_copy = df_raw.copy()
    df_copy.columns = [f"Col_{i}" for i in range(len(df_copy.columns))]
    return df_copy

# ─── Cell-level indexing for Smart Query ──────────────────────────────────────

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

@st.cache_data(show_spinner=False)
def index_raw_df(df):
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


def get_summary_stats(all_data):
    stats = {}
    total_sheets = 0
    total_rows = 0
    total_cols = 0
    for location, loc_data in all_data.items():
        for sheet_name, sheet_data in loc_data["sheets"].items():
            df = sheet_data["structured"]
            total_sheets += 1
            total_rows += len(df)
            total_cols += len([c for c in df.columns if not c.startswith("__")])
    stats["files"] = len(all_data)
    stats["sheets"] = total_sheets
    stats["rows"] = total_rows
    stats["columns"] = total_cols
    return stats

def get_numeric_columns(df):
    numeric_cols = []
    for col in df.columns:
        if col.startswith("__"):
            continue
        numeric_col = pd.to_numeric(df[col], errors="coerce")
        if numeric_col.notna().sum() > 0:
            numeric_cols.append(col)
    return numeric_cols

def get_categorical_columns(df):
    cat_cols = []
    for col in df.columns:
        if col.startswith("__"):
            continue
        numeric_col = pd.to_numeric(df[col], errors="coerce")
        if numeric_col.isna().sum() > len(df) * 0.5:
            unique_vals = df[col].dropna().unique()
            if 2 <= len(unique_vals) <= 50:
                cat_cols.append(col)
    return cat_cols


# ─── Main App ───────────────────────────────────────────────────────────────

st.title("📊 Customer & Capacity Tracker — AI Dashboard")
st.markdown("*Smart analytics across all locations: Airoli, Bangalore, Chennai, Kolkata, Noida, Rabale, Vashi*")

with st.spinner("Loading all Excel files..."):
    all_data = load_all_excel_files()

if not all_data:
    st.error("No Excel files could be loaded. Please check the attached_assets folder.")
    st.stop()

stats = get_summary_stats(all_data)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📁 Files Loaded", stats["files"])
with col2:
    st.metric("📋 Total Sheets", stats["sheets"])
with col3:
    st.metric("📝 Total Rows", f"{stats['rows']:,}")
with col4:
    st.metric("🔢 Total Columns", stats["columns"])

tabs = st.tabs(["🤖 Smart Query", "📊 Dashboard", "🔍 Data Explorer", "📈 Analytics", "📋 Raw Data"])

# ─── Tab 1: Smart Query ──────────────────────────────────────────────────────
with tabs[0]:
    st.header("💬 Smart Query — Ask Anything About Your Data")
    st.caption("Cell-level search: matches exact values, column headers, and cross-column lookups across every row in the selected sheet.")

    # ── Location & Sheet selector ──
    sq_c1, sq_c2 = st.columns(2)
    with sq_c1:
        sq_locations = list(all_data.keys())
        sq_sel_loc = st.selectbox("📍 Location", sq_locations, key="sq_loc")
    with sq_c2:
        sq_sheets = list(all_data[sq_sel_loc]["sheets"].keys()) if sq_sel_loc else []
        sq_sel_sheet = st.selectbox("📋 Sheet", sq_sheets, key="sq_sheet")

    sq_raw = all_data[sq_sel_loc]["sheets"][sq_sel_sheet]["raw"] if (sq_sel_loc and sq_sel_sheet) else None

    if sq_raw is not None:
        with st.spinner("Indexing sheet — reading every cell, row and column…"):
            sq_cells, sq_rows, sq_meta = index_raw_df(sq_raw)
    else:
        sq_cells, sq_rows, sq_meta = [], {}, {"total_cells": 0, "total_data": 0, "total_rows": 0, "total_headers": 0}
        st.info("Select a location and sheet above to start querying.")

    # ── Stats row ──
    qi1, qi2, qi3, qi4 = st.columns(4)
    qi1.metric("Total Cells",   f"{sq_meta['total_cells']:,}")
    qi2.metric("Data Cells",    f"{sq_meta['total_data']:,}")
    qi3.metric("Data Rows",     f"{sq_meta['total_rows']:,}")
    qi4.metric("Header Cells",  f"{sq_meta['total_headers']:,}")

    # ── Internal helpers & query engine ──────────────────────────────────────
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
        "space":        ["space","sqft","sq ft","area"],
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

    def _mcol(kw, hdr):
        hl  = hdr.lower()
        kwl = kw.lower()
        if kwl in hl:
            return True
        for key, syns in _SYN.items():
            if kwl in syns or kwl == key:
                for s in syns:
                    if s in hl:
                        return True
        return False

    def sheet_query(question):
        q  = question.strip()
        ql = q.lower()
        sig = [w for w in re.findall(r"[a-z0-9]{2,}", ql) if w not in _SW]

        f_sum  = any(x in ql for x in ["total","sum","aggregate"])
        f_avg  = any(x in ql for x in ["average","mean","avg"])
        f_max  = any(x in ql for x in ["maximum","highest","largest","max","top value"])
        f_min  = any(x in ql for x in ["minimum","lowest","smallest","min"])
        f_cnt  = any(x in ql for x in ["count","how many","number of"])
        f_stat = any(x in ql for x in ["statistics","stats","describe","summary","all stats"])
        f_uniq = any(x in ql for x in ["unique","distinct","different"])
        f_miss = any(x in ql for x in ["missing","null","blank","empty"])
        f_cols = any(x in ql for x in ["column","columns","field","header","headers"])
        f_topn = re.search(r"\btop\s*(\d+)\b", ql)
        f_botn = re.search(r"\bbottom\s*(\d+)\b", ql)
        f_num  = (f_sum or f_avg or f_max or f_min or f_cnt or f_stat
                  or bool(f_topn) or bool(f_botn))
        f_cust = any(x in ql for x in ["customer","customers","client","clients","name","names"])
        f_list = any(x in ql for x in ["list","show","display","all"])

        out = {"answer": "", "table": None, "chart_df": None,
               "chart_cfg": None, "cell_hits": [], "sub_tables": []}
        wc = sq_cells
        rr = sq_rows

        if not wc:
            out["answer"] = "No data indexed in this sheet."
            return out

        col_kws = [w for w in sig if w not in _OP_VERBS]

        # ── helpers ──
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

        def _col_match_score(w):
            return sum(1 for c in wc if c["is_header"] and _mcol(w, c["value"]))

        def _val_match_rows(terms):
            rows_s, cells_s = set(), []
            for cell in wc:
                if cell["is_header"]: continue
                if any(t in cell["value"].lower() for t in terms):
                    rows_s.add(cell["row"])
                    cells_s.append(cell)
            return sorted(rows_s), cells_s

        # ── INTENT: List all columns ──────────────────────────────────────────
        if f_cols and not f_num:
            hdrs = sorted({c["col_header"] for c in wc if c["is_header"]})
            data_hdrs = sorted({c["col_header"] for c in wc if not c["is_header"]})
            all_hdrs = sorted(set(hdrs) | set(data_hdrs))
            out["answer"] = f"**{len(all_hdrs)}** unique column headers detected in this sheet."
            out["table"]  = pd.DataFrame({"Column Header": all_hdrs})
            return out

        # ── INTENT: Missing / null ────────────────────────────────────────────
        if f_miss:
            total_possible = sq_raw.shape[0] * sq_raw.shape[1]
            non_null = len(wc)
            missing  = total_possible - non_null
            out["answer"] = (
                f"**{missing:,}** missing (null/empty) cells out of "
                f"**{total_possible:,}** total cell positions "
                f"({missing/max(total_possible,1)*100:.1f}% missing)."
            )
            return out

        # ── INTENT: Numeric operations ────────────────────────────────────────
        if f_num:
            kw, pairs = npbest(col_kws or sig)
            if pairs:
                vals = [v for v, _ in pairs]
                sa   = pd.Series(vals)
                parts = []
                if f_sum:  parts.append(f"**Sum of '{kw}':** {sa.sum():,.4f}")
                if f_avg:  parts.append(f"**Avg of '{kw}':** {sa.mean():,.4f}")
                if f_max:  parts.append(f"**Max of '{kw}':** {sa.max():,.4f}")
                if f_min:  parts.append(f"**Min of '{kw}':** {sa.min():,.4f}")
                if f_cnt:  parts.append(f"**Count of '{kw}':** {sa.count():,}")
                if f_stat:
                    parts.append(
                        f"**Stats for '{kw}'**\n"
                        f"- Sum: {sa.sum():,.4f}\n"
                        f"- Avg: {sa.mean():,.4f}\n"
                        f"- Max: {sa.max():,.4f}\n"
                        f"- Min: {sa.min():,.4f}\n"
                        f"- Median: {sa.median():,.4f}\n"
                        f"- Std Dev: {sa.std():,.4f}"
                    )
                if not parts:
                    parts.append(f"**{len(vals):,}** numeric values found under **'{kw}'**")

                out["answer"] = "\n\n".join(parts)
                tbl_rows = sorted({c["row"] for _, c in pairs})
                full_df  = build_rows_df(tbl_rows)
                if not full_df.empty:
                    out["sub_tables"].append({
                        "label": f"📋 Full Row Data for '{kw}' ({len(tbl_rows)} rows)",
                        "df": full_df,
                    })

                if f_topn:
                    n   = int(f_topn.group(1))
                    top = sorted(pairs, key=lambda x: x[0], reverse=True)[:n]
                    top_df = build_rows_df(sorted({c["row"] for _, c in top}))
                    out["sub_tables"].append({
                        "label": f"🏆 Top {n} rows by '{kw}'",
                        "df": top_df if not top_df.empty else pd.DataFrame(
                            [{"Row": c["row"]+1, "Column": c["col_header"], "Value": v} for v,c in top]),
                    })
                if f_botn:
                    n   = int(f_botn.group(1))
                    bot = sorted(pairs, key=lambda x: x[0])[:n]
                    bot_df = build_rows_df(sorted({c["row"] for _, c in bot}))
                    out["sub_tables"].append({
                        "label": f"🔻 Bottom {n} rows by '{kw}'",
                        "df": bot_df if not bot_df.empty else pd.DataFrame(
                            [{"Row": c["row"]+1, "Column": c["col_header"], "Value": v} for v,c in bot]),
                    })
                return out

            # fallback – all numeric cells
            anums = [(float(c["value"]), c) for c in wc
                     if not c["is_header"] and _is_num(c["value"])]
            if anums:
                vals = [v for v, _ in anums]
                sa   = pd.Series(vals)
                parts = []
                if f_sum:  parts.append(f"**Sum ALL numeric cells:** {sa.sum():,.4f}")
                if f_avg:  parts.append(f"**Avg ALL numeric cells:** {sa.mean():,.4f}")
                if f_max:  parts.append(f"**Max ALL numeric cells:** {sa.max():,.4f}")
                if f_min:  parts.append(f"**Min ALL numeric cells:** {sa.min():,.4f}")
                if f_cnt:  parts.append(f"**Count ALL numeric cells:** {sa.count():,}")
                if f_stat:
                    parts.append(
                        f"**All Numeric Cells Stats**\n"
                        f"- Sum: {sa.sum():,.4f}\n- Avg: {sa.mean():,.4f}\n"
                        f"- Max: {sa.max():,.4f}\n- Min: {sa.min():,.4f}\n"
                        f"- Median: {sa.median():,.4f}\n- Std Dev: {sa.std():,.4f}"
                    )
                out["answer"] = ("No column matched your keywords. Results from **ALL numeric cells**:\n\n"
                                 + "\n".join(parts))
                return out

        # ── INTENT: Unique values ─────────────────────────────────────────────
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
                    tbl = pd.DataFrame(sr).drop_duplicates(subset=["Value"])
                    out["answer"] = f"**{len(uv)}** unique value(s) for **'{w}'**."
                    out["table"] = tbl
                    return out

        # ── INTENT: Cross-column relational lookup ────────────────────────────
        # e.g. "power of CISCO" / "subscription for HDFC" / "WIPRO capacity"
        if not f_num:
            tokens = [w for w in sig if w not in _OP_VERBS]

            col_scores = {w: _col_match_score(w) for w in tokens}
            val_scores = {w: sum(1 for c in wc
                                 if not c["is_header"] and w in c["value"].lower())
                          for w in tokens}

            attr_toks   = [w for w in tokens if col_scores.get(w, 0) > 0]
            entity_toks = [w for w in tokens if val_scores.get(w, 0) > 0]

            quoted = re.findall(r'"([^"]+)"', q)
            if quoted:
                ep = [quoted[0].lower()]
                er, ec = _val_match_rows(ep)
                if er:
                    entity_toks = ep
                    entity_rows, entity_cells = er, ec
                else:
                    entity_rows, entity_cells = [], []
            else:
                entity_rows, entity_cells = _val_match_rows(entity_toks)

            # Case 1: both attribute (column) and entity (value) found
            if attr_toks and entity_rows:
                recs = []
                for rn in sorted(entity_rows):
                    row_data = rr.get(rn, {})
                    if not row_data: continue
                    rd = {"Row #": rn + 1}
                    for c in entity_cells:
                        if c["row"] == rn:
                            rd[f"[Matched] {c['col_header']}"] = c["value"]
                    for at in attr_toks:
                        for col_h, val in row_data.items():
                            if _mcol(at, col_h):
                                rd[col_h] = val
                    recs.append(rd)
                if recs:
                    tbl = pd.DataFrame(recs)
                    ed  = ", ".join(f"'{t}'" for t in entity_toks[:3])
                    ad  = ", ".join(f"'{t}'" for t in attr_toks[:3])
                    out["answer"] = (
                        f"Cross-column lookup — entity **{ed}** found in **{len(entity_rows)}** row(s).\n\n"
                        f"Showing **{ad}** column value(s) for those rows:"
                    )
                    out["table"] = tbl
                    full_df = build_rows_df(entity_rows)
                    if not full_df.empty:
                        out["sub_tables"].append({
                            "label": f"📋 All Column Values for Matching Rows ({len(entity_rows)})",
                            "df": full_df,
                        })
                    return out

            # Case 2: only entity found → show all columns of matching rows
            if entity_rows:
                full_df = build_rows_df(entity_rows)
                md = ", ".join(f"'{t}'" for t in entity_toks[:4])
                out["answer"] = (
                    f"Found **{len(entity_cells):,}** cell(s) matching **{md}** "
                    f"across **{len(entity_rows):,}** row(s).\n\n"
                    f"Showing **all column values** for every matching row:"
                )
                out["table"] = full_df if not full_df.empty else None
                out["cell_hits"] = [
                    {"Row #": c["row"]+1, "Col #": c["col"]+1,
                     "Column Header": c["col_header"], "Value": c["value"]}
                    for c in entity_cells[:100]
                ]
                return out

            # Case 3: only attribute/column token → list all values in that column
            if attr_toks:
                best = max(attr_toks, key=lambda w: col_scores[w])
                sr, seen = [], set()
                for cell in wc:
                    if cell["is_header"]: continue
                    if _mcol(best, cell["col_header"]) and cell["row"] not in seen:
                        seen.add(cell["row"])
                        sr.append({"Row #": cell["row"]+1,
                                   "Column": cell["col_header"], "Value": cell["value"]})
                if sr:
                    out["answer"] = (
                        f"Found **{len(sr)}** value(s) in column(s) matching **'{best}'**."
                    )
                    out["table"] = pd.DataFrame(sr)
                    full_df = build_rows_df(list(seen))
                    if not full_df.empty:
                        out["sub_tables"].append({
                            "label": f"📋 Full Row Data for '{best}' ({len(seen)} rows)",
                            "df": full_df,
                        })
                    return out

        # ── INTENT: List customers / names ────────────────────────────────────
        if (f_cust or f_list) and not f_num:
            found = []
            for cell in wc:
                if cell["is_header"]: continue
                ch = cell["col_header"].lower()
                if any(x in ch for x in ["customer","name","client","company"]):
                    found.append({"Row #": cell["row"]+1,
                                  "Column": cell["col_header"], "Value": cell["value"]})
            if found:
                tbl = pd.DataFrame(found).drop_duplicates(subset=["Value"])
                out["answer"] = f"Found **{len(tbl)}** customer/name entries."
                out["table"]  = tbl
                name_rows = sorted({r["Row #"]-1 for r in found})
                full_df   = build_rows_df(name_rows)
                if not full_df.empty:
                    out["sub_tables"].append({
                        "label": f"📋 Full Row Data ({len(name_rows)} rows)",
                        "df": full_df,
                    })
            else:
                out["answer"] = ("No 'Customer/Name/Client' columns found. "
                                 "Try: *Find CISCO* or *Show all rows*")
            return out

        out["answer"] = (
            "Could not match your query.\n\n"
            "**Try relational queries like:**\n"
            "- *Power of CISCO* · *Subscription for HDFC* · *WIPRO capacity*\n\n"
            "**Or standard queries:**\n"
            "- *List all customers* · *Find CISCO* · *Total subscription*\n"
            "- *Top 10 power* · *Average capacity* · *Statistics of usage*\n"
            "- *Show all columns* · *How many missing values*"
        )
        return out

    def render_answer(res, tidx=0):
        st.markdown(res["answer"])
        if res.get("table") is not None and not res["table"].empty:
            tbl = res["table"].reset_index(drop=True)
            st.dataframe(tbl, use_container_width=True,
                         height=min(540, 50 + len(tbl) * 36),
                         key=f"tbl_{tidx}")
            st.download_button(
                "⬇️ Download CSV",
                tbl.to_csv(index=False).encode(),
                "query_result.csv", "text/csv",
                key=f"dl_{tidx}",
            )
        for si, s in enumerate(res.get("sub_tables", [])):
            with st.expander(s["label"], expanded=True):
                st.dataframe(s["df"], use_container_width=True,
                             key=f"sub_{tidx}_{si}")
        if res.get("cell_hits"):
            with st.expander(f"🔬 Matching Cells ({len(res['cell_hits'])})", expanded=False):
                for ch in res["cell_hits"]:
                    st.code(
                        f"Row {ch['Row #']} · Col {ch['Col #']} | "
                        f"{ch['Column Header']} | {ch['Value']}"
                    )

    # ── Suggested queries ──
    with st.expander("💡 Suggested Queries — click to expand", expanded=False):
        st.markdown("""
**🔗 Cross-Column Relational Lookups:**
- `Power of CISCO` — find CISCO in any column, show its power value
- `Subscription for HDFC` — find HDFC rows, show subscription column
- `WIPRO capacity` — find WIPRO, show capacity from same rows
- `Rack of TCS` — find TCS, show rack details
- `Show power where customer is CISCO`

**📊 Numeric Operations:**
- `Total subscription` · `Average capacity` · `Max power` · `Min rack`
- `Count customers` · `Top 10 power` · `Bottom 5 usage`
- `Statistics of capacity`

**🔍 Search & List:**
- `Find CISCO` · `List all customers` · `Show HDFC`

**ℹ️ Info:**
- `Show all columns` · `How many missing values` · `Unique values of rack`
""")

    # ── Chat history ──
    hist_key = f"sq_hist_{sq_sel_loc}_{sq_sel_sheet}"
    if hist_key not in st.session_state:
        st.session_state[hist_key] = []

    for tidx, turn in enumerate(st.session_state[hist_key]):
        st.markdown(f"**You:** {turn['q']}")
        render_answer(turn["res"], tidx)
        st.divider()

    # ── Input bar ──
    st.divider()
    ic, bc, cc = st.columns([8, 1, 1])
    with ic:
        user_q = st.text_input(
            "Ask anything about this sheet:",
            placeholder="e.g. Total subscription · Find CISCO · Top 10 power · Average capacity",
            key=f"user_q_{hist_key}",
            label_visibility="collapsed",
        )
    with bc:
        ask = st.button("Ask ▶", type="primary", use_container_width=True)
    with cc:
        if st.button("Clear 🗑", use_container_width=True):
            st.session_state[hist_key] = []
            st.rerun()

    if ask and user_q.strip():
        with st.spinner("🔍 Searching every row and column…"):
            res = sheet_query(user_q)
        st.session_state[hist_key].append({"q": user_q, "res": res})
        st.rerun()

# ─── Tab 2: Dashboard ────────────────────────────────────────────────────────
with tabs[1]:
    st.header("📊 Interactive Dashboard")

    all_structured = []
    for location, loc_data in all_data.items():
        for sheet_name, sheet_data in loc_data["sheets"].items():
            df = sheet_data["structured"].copy()
            df["Location"] = location
            df["Sheet"] = sheet_name
            all_structured.append(df)

    if all_structured:
        location_counts = {loc: 0 for loc in all_data.keys()}
        location_rows = {loc: 0 for loc in all_data.keys()}

        for df in all_structured:
            if "Location" in df.columns:
                for loc, count in df["Location"].value_counts().items():
                    location_rows[loc] = location_rows.get(loc, 0) + count

        location_sheets = {}
        for location, loc_data in all_data.items():
            location_sheets[location] = len(loc_data["sheets"])

        dc1, dc2 = st.columns(2)

        with dc1:
            st.subheader("📍 Sheets per Location")
            loc_sheet_df = pd.DataFrame({
                "Location": list(location_sheets.keys()),
                "Sheets": list(location_sheets.values())
            })
            fig_pie = px.pie(
                loc_sheet_df, values="Sheets", names="Location",
                title="Distribution of Sheets by Location",
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            fig_pie.update_traces(textposition="inside", textinfo="percent+label")
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)

        with dc2:
            st.subheader("📋 Data Rows per Location")
            loc_rows_df = pd.DataFrame({
                "Location": list(location_rows.keys()),
                "Rows": list(location_rows.values())
            }).sort_values("Rows", ascending=False)
            fig_bar = px.bar(
                loc_rows_df, x="Location", y="Rows",
                title="Number of Data Rows per Location",
                color="Rows",
                color_continuous_scale="Blues",
                text="Rows"
            )
            fig_bar.update_traces(textposition="outside")
            fig_bar.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig_bar, use_container_width=True)

        st.divider()
        st.subheader("🔍 Location-wise Sheet Analytics")

        sel_location = st.selectbox("Select Location", list(all_data.keys()), key="dash_location")
        if sel_location and sel_location in all_data:
            sel_sheet = st.selectbox(
                "Select Sheet",
                list(all_data[sel_location]["sheets"].keys()),
                key="dash_sheet"
            )
            if sel_sheet:
                df_sel = all_data[sel_location]["sheets"][sel_sheet]["structured"].copy()
                numeric_cols = get_numeric_columns(df_sel)
                cat_cols = get_categorical_columns(df_sel)

                st.markdown(f"**{sel_location} → {sel_sheet}** | Rows: {len(df_sel)} | Columns: {len(df_sel.columns)}")

                if numeric_cols:
                    nc1, nc2 = st.columns(2)
                    with nc1:
                        y_col = st.selectbox("Y-axis (numeric)", numeric_cols, key="y_col_dash")
                    with nc2:
                        x_col = st.selectbox("X-axis", [c for c in df_sel.columns if c not in ["__location__", "__sheet__"]], key="x_col_dash")

                    chart_type = st.radio("Chart Type", ["Bar", "Line", "Scatter", "Box", "Histogram"], horizontal=True)

                    df_plot = df_sel.copy()
                    df_plot[y_col] = pd.to_numeric(df_plot[y_col], errors="coerce")
                    df_plot = df_plot.dropna(subset=[y_col])

                    if not df_plot.empty:
                        if chart_type == "Bar":
                            fig = px.bar(df_plot, x=x_col, y=y_col, title=f"{y_col} by {x_col}",
                                        color=y_col, color_continuous_scale="Viridis")
                        elif chart_type == "Line":
                            fig = px.line(df_plot, x=x_col, y=y_col, title=f"{y_col} over {x_col}",
                                         markers=True)
                        elif chart_type == "Scatter":
                            fig = px.scatter(df_plot, x=x_col, y=y_col, title=f"{y_col} vs {x_col}",
                                            color=y_col, size=y_col, color_continuous_scale="Rainbow")
                        elif chart_type == "Box":
                            fig = px.box(df_plot, y=y_col, title=f"Box Plot: {y_col}")
                        else:
                            fig = px.histogram(df_plot, x=y_col, title=f"Distribution of {y_col}",
                                              color_discrete_sequence=["#636EFA"])

                        fig.update_layout(height=400)
                        st.plotly_chart(fig, use_container_width=True)

                if cat_cols:
                    st.subheader("📊 Category Distribution")
                    cat_sel = st.selectbox("Category column", cat_cols, key="cat_col_dash")
                    val_counts = df_sel[cat_sel].value_counts().reset_index()
                    val_counts.columns = [cat_sel, "Count"]

                    vc1, vc2 = st.columns(2)
                    with vc1:
                        fig_donut = px.pie(val_counts, values="Count", names=cat_sel,
                                          title=f"Distribution of {cat_sel}", hole=0.4,
                                          color_discrete_sequence=px.colors.qualitative.Pastel)
                        st.plotly_chart(fig_donut, use_container_width=True)
                    with vc2:
                        fig_hbar = px.bar(val_counts.head(20), x="Count", y=cat_sel,
                                         orientation="h",
                                         title=f"Top values in {cat_sel}",
                                         color="Count", color_continuous_scale="Teal")
                        fig_hbar.update_layout(height=400)
                        st.plotly_chart(fig_hbar, use_container_width=True)

        st.divider()
        st.subheader("🌐 Cross-Location Comparison")

        all_numeric_cols = set()
        for df in all_structured:
            for col in get_numeric_columns(df):
                all_numeric_cols.add(col)

        if all_numeric_cols:
            compare_col = st.selectbox("Select numeric column to compare", sorted(all_numeric_cols), key="compare_col")
            loc_compare = []
            for df in all_structured:
                if compare_col in df.columns:
                    numeric_series = pd.to_numeric(df[compare_col], errors="coerce").dropna()
                    if len(numeric_series) > 0:
                        loc_compare.append({
                            "Location": df["Location"].iloc[0] if "Location" in df.columns else "Unknown",
                            "Sheet": df["Sheet"].iloc[0] if "Sheet" in df.columns else "Unknown",
                            "Sum": numeric_series.sum(),
                            "Mean": numeric_series.mean(),
                            "Max": numeric_series.max(),
                            "Min": numeric_series.min(),
                            "Count": len(numeric_series)
                        })

            if loc_compare:
                cmp_df = pd.DataFrame(loc_compare)
                fig_cmp = px.bar(
                    cmp_df, x="Location", y="Sum", color="Sheet",
                    title=f"Total '{compare_col}' by Location",
                    barmode="group",
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                fig_cmp.update_layout(height=450)
                st.plotly_chart(fig_cmp, use_container_width=True)
                st.dataframe(cmp_df, use_container_width=True)

# ─── Tab 3: Data Explorer ────────────────────────────────────────────────────
with tabs[2]:
    st.header("🔍 Data Explorer — Browse Any Sheet")

    exp_loc = st.selectbox("📍 Select Location", list(all_data.keys()), key="exp_loc")
    if exp_loc:
        exp_sheet = st.selectbox(
            "📋 Select Sheet",
            list(all_data[exp_loc]["sheets"].keys()),
            key="exp_sheet"
        )
        if exp_sheet:
            df_exp = all_data[exp_loc]["sheets"][exp_sheet]["structured"].copy()
            raw_df = all_data[exp_loc]["sheets"][exp_sheet]["raw"].copy()

            view_mode = st.radio("View Mode", ["Structured (auto-detected headers)", "Raw (original)"], horizontal=True)

            if view_mode.startswith("Raw"):
                display_data = raw_df
            else:
                display_data = df_exp

            st.markdown(f"**Shape:** {display_data.shape[0]} rows × {display_data.shape[1]} columns")

            search_filter = st.text_input("🔎 Filter rows (search any cell value):", key="exp_search")
            if search_filter:
                mask = display_data.apply(
                    lambda col: col.astype(str).str.contains(search_filter, case=False, na=False)
                ).any(axis=1)
                display_data = display_data[mask]
                st.markdown(f"Showing {len(display_data)} matching rows")

            st.dataframe(display_data, use_container_width=True, height=500)

            ec1, ec2 = st.columns(2)
            with ec1:
                csv_data = display_data.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Download as CSV",
                    data=csv_data,
                    file_name=f"{exp_loc}_{exp_sheet}.csv",
                    mime="text/csv"
                )
            with ec2:
                numeric_cols_exp = get_numeric_columns(df_exp)
                if numeric_cols_exp:
                    st.markdown(f"**Numeric columns:** {', '.join(numeric_cols_exp[:5])}")

            st.subheader("📊 Quick Column Stats")
            stats_cols = [c for c in df_exp.columns if not c.startswith("__")]
            if stats_cols:
                stat_col = st.selectbox("Select column", stats_cols, key="stat_col")
                col_data = df_exp[stat_col]
                numeric_col_data = pd.to_numeric(col_data, errors="coerce")

                scol1, scol2, scol3, scol4, scol5 = st.columns(5)
                scol1.metric("Count", col_data.dropna().count())
                scol2.metric("Unique", col_data.nunique())
                if numeric_col_data.notna().sum() > 0:
                    scol3.metric("Sum", f"{numeric_col_data.sum():,.2f}")
                    scol4.metric("Max", f"{numeric_col_data.max():,.2f}")
                    scol5.metric("Min", f"{numeric_col_data.min():,.2f}")

                    fig_hist = px.histogram(
                        df_exp, x=stat_col,
                        title=f"Distribution of {stat_col}",
                        color_discrete_sequence=["#00CC96"]
                    )
                    st.plotly_chart(fig_hist, use_container_width=True)
                else:
                    top_vals = col_data.value_counts().head(15).reset_index()
                    top_vals.columns = [stat_col, "Count"]
                    fig_top = px.bar(top_vals, x=stat_col, y="Count",
                                    title=f"Top values in {stat_col}",
                                    color="Count", color_continuous_scale="Sunset")
                    st.plotly_chart(fig_top, use_container_width=True)

# ─── Tab 4: Analytics ────────────────────────────────────────────────────────
with tabs[3]:
    st.header("📈 Advanced Analytics & Operations")

    an_loc = st.selectbox("📍 Location", list(all_data.keys()), key="an_loc")
    an_sheet = st.selectbox(
        "📋 Sheet",
        list(all_data[an_loc]["sheets"].keys()) if an_loc else [],
        key="an_sheet"
    )

    if an_loc and an_sheet:
        df_an = all_data[an_loc]["sheets"][an_sheet]["structured"].copy()
        numeric_cols_an = get_numeric_columns(df_an)
        cat_cols_an = get_categorical_columns(df_an)

        if numeric_cols_an:
            st.subheader("🔢 Numeric Operations")
            op_col = st.selectbox("Select column for operations", numeric_cols_an, key="op_col")
            series = pd.to_numeric(df_an[op_col], errors="coerce").dropna()

            oc1, oc2, oc3, oc4, oc5, oc6 = st.columns(6)
            oc1.metric("Sum", f"{series.sum():,.2f}")
            oc2.metric("Count", f"{len(series):,}")
            oc3.metric("Average", f"{series.mean():,.2f}")
            oc4.metric("Max", f"{series.max():,.2f}")
            oc5.metric("Min", f"{series.min():,.2f}")
            oc6.metric("Std Dev", f"{series.std():,.2f}")

            p25, p50, p75 = series.quantile([0.25, 0.5, 0.75])
            pc1, pc2, pc3, pc4 = st.columns(4)
            pc1.metric("Median", f"{p50:,.2f}")
            pc2.metric("25th Pct", f"{p25:,.2f}")
            pc3.metric("75th Pct", f"{p75:,.2f}")
            pc4.metric("Range", f"{series.max() - series.min():,.2f}")

            st.divider()

            if cat_cols_an:
                st.subheader("📊 Group-by Analysis")
                group_col = st.selectbox("Group by (category)", cat_cols_an, key="group_col")
                agg_func = st.selectbox("Aggregation", ["sum", "mean", "count", "max", "min"], key="agg_func")

                df_an_copy = df_an.copy()
                df_an_copy[op_col] = pd.to_numeric(df_an_copy[op_col], errors="coerce")

                grouped = df_an_copy.groupby(group_col)[op_col].agg(agg_func).reset_index()
                grouped.columns = [group_col, f"{agg_func.title()} of {op_col}"]
                grouped = grouped.dropna().sort_values(f"{agg_func.title()} of {op_col}", ascending=False)

                gc1, gc2 = st.columns(2)
                with gc1:
                    fig_grp = px.bar(
                        grouped.head(20), x=group_col, y=f"{agg_func.title()} of {op_col}",
                        title=f"{agg_func.title()} of {op_col} by {group_col}",
                        color=f"{agg_func.title()} of {op_col}",
                        color_continuous_scale="Plasma"
                    )
                    fig_grp.update_xaxes(tickangle=45)
                    fig_grp.update_layout(height=400)
                    st.plotly_chart(fig_grp, use_container_width=True)

                with gc2:
                    fig_grp_pie = px.pie(
                        grouped.head(15), values=f"{agg_func.title()} of {op_col}", names=group_col,
                        title=f"{agg_func.title()} Distribution by {group_col}",
                        color_discrete_sequence=px.colors.qualitative.G10
                    )
                    fig_grp_pie.update_traces(textposition="inside", textinfo="percent+label")
                    fig_grp_pie.update_layout(height=400)
                    st.plotly_chart(fig_grp_pie, use_container_width=True)

                st.dataframe(grouped, use_container_width=True)

            st.divider()
            st.subheader("🫧 Bubble & Scatter Analysis")
            if len(numeric_cols_an) >= 2:
                bc1, bc2, bc3 = st.columns(3)
                with bc1:
                    bub_x = st.selectbox("X axis", numeric_cols_an, key="bub_x")
                with bc2:
                    bub_y = st.selectbox("Y axis", numeric_cols_an, index=min(1, len(numeric_cols_an)-1), key="bub_y")
                with bc3:
                    bub_size = st.selectbox("Bubble size", numeric_cols_an, index=min(2, len(numeric_cols_an)-1), key="bub_size")

                df_bub = df_an.copy()
                df_bub[bub_x] = pd.to_numeric(df_bub[bub_x], errors="coerce")
                df_bub[bub_y] = pd.to_numeric(df_bub[bub_y], errors="coerce")
                df_bub[bub_size] = pd.to_numeric(df_bub[bub_size], errors="coerce").abs()
                df_bub = df_bub.dropna(subset=[bub_x, bub_y, bub_size])

                if not df_bub.empty:
                    color_col = cat_cols_an[0] if cat_cols_an else None
                    fig_bub = px.scatter(
                        df_bub, x=bub_x, y=bub_y, size=bub_size,
                        color=color_col,
                        title=f"Bubble Chart: {bub_x} vs {bub_y} (size={bub_size})",
                        color_discrete_sequence=px.colors.qualitative.Vivid,
                        size_max=60
                    )
                    fig_bub.update_layout(height=500)
                    st.plotly_chart(fig_bub, use_container_width=True)

            st.divider()
            st.subheader("📊 Correlation Heatmap")
            if len(numeric_cols_an) >= 2:
                df_corr = df_an[numeric_cols_an].apply(pd.to_numeric, errors="coerce")
                corr_matrix = df_corr.corr()
                fig_heat = px.imshow(
                    corr_matrix, title="Correlation Matrix",
                    color_continuous_scale="RdBu_r", zmin=-1, zmax=1,
                    text_auto=True
                )
                fig_heat.update_layout(height=400)
                st.plotly_chart(fig_heat, use_container_width=True)

        else:
            st.info("No numeric columns detected in this sheet. Try another sheet.")

# ─── Tab 5: Raw Data ─────────────────────────────────────────────────────────
with tabs[4]:
    st.header("📋 All Files & Sheets Overview")

    for location, loc_data in all_data.items():
        with st.expander(f"📁 {location} — {loc_data['filename']}", expanded=False):
            st.markdown(f"**Sheets ({len(loc_data['sheets'])}):**")
            for sheet_name, sheet_data in loc_data["sheets"].items():
                raw = sheet_data["raw"]
                structured = sheet_data["structured"]
                st.markdown(f"- **{sheet_name}** — Raw: {raw.shape[0]}×{raw.shape[1]} | Structured: {structured.shape[0]}×{len([c for c in structured.columns if not c.startswith('__')])} cols")

            if st.button(f"View all sheets for {location}", key=f"view_{location}"):
                for sheet_name, sheet_data in loc_data["sheets"].items():
                    st.subheader(f"📋 {sheet_name}")
                    st.dataframe(sheet_data["structured"], use_container_width=True, height=300)

st.divider()
st.markdown(
    "<div style='text-align:center; color:#888'>Customer & Capacity Tracker AI Dashboard | "
    "Data from: Airoli · Bangalore · Chennai · Kolkata · Noida · Rabale · Vashi</div>",
    unsafe_allow_html=True
)
