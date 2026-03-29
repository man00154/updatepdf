# Sify DC – Capacity Intelligence Dashboard

An AI-powered, multi-tab Streamlit dashboard for analysing Sify Data Centre Customer & Capacity Tracker Excel files — including unstructured, odd-format multi-header layouts.

---

## Features

| Tab | Description |
|-----|-------------|
| 🏠 Overview | KPI cards · column stats · complete raw sheet (all rows & columns) |
| 📋 Raw Data | Live search · filtered export · raw vs. clean toggle |
| 📊 Analytics | Aggregations (sum, mean, median, std, variance, % total) · group-by |
| 📈 Charts | 15 chart types (bar, line, scatter, bubble, heatmap, radar, 3D, …) |
| 🥧 Distributions | Pie/donut · histogram · box plot · sunburst |
| 🔍 Query Engine | No-code: sum · avg · min · max · count · % · top/bottom N |
| 🌍 Multi-Location | Cross-location comparison with bar, pie, and radar charts |
| 🤖 AI Agent | Automated insights – outlier detection, trends, file summaries |
| 💬 Smart Query | Natural-language chat – fetches any row/column/sheet without formulas |

---

## Supported File Formats

- `.xlsx` — read via **openpyxl** (handles complex merged/multi-header layouts)
- `.xls` — read via **xlrd** (legacy OLE2 Excel format)

Files are auto-detected from the `excel_files/` folder **or** uploaded live via the sidebar.

---

## Repository Structure

```
your-repo/
│
├── app.py                    ← Main Streamlit application
├── requirements.txt          ← Python dependencies
├── packages.txt              ← System-level dependencies (Streamlit Cloud)
│
├── .streamlit/
│   └── config.toml           ← Server config (port 5000, headless)
│
└── excel_files/              ← Place your Excel files here
    ├── Customer_and_Capacity_Tracker_Airoli_15Mar26.xlsx
    ├── Customer_and_Capacity_Tracker_Bangalore_01_15Feb26.xlsx
    ├── Customer_and_Capacity_Tracker_Chennai_01_15Feb26.xls
    ├── ... (all your .xlsx / .xls files)
    └── .gitkeep              ← Keeps the folder in git even when empty
```

---

## Running Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/YOUR_REPO.git
cd YOUR_REPO

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add your Excel files
#    Copy your .xlsx / .xls files into the excel_files/ folder

# 4. Run
streamlit run app.py
```

App opens at **http://localhost:5000**

---

## Deploy on Streamlit Cloud (Free)

1. Push this repo to GitHub (include `excel_files/` with your Excel files)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select your GitHub repo
4. Set **Main file path** → `app.py`
5. Click **Deploy** — done ✅

> **Tip:** If your Excel files are large (> 100 MB), use Streamlit's sidebar file uploader instead of committing them to git.

---

## Smart Query Examples

Type any natural-language question in the **💬 Smart Query** tab:

```
Total subscription
Average capacity
Max power usage
Top 10 rack
Bottom 5 subscription
Find CISCO
List all customers
Show HDFC
Statistics of capacity
How many missing values
Unique values of rack
Percentage of subscription
```

The agent searches **every row, every column, every sheet** — no formulas, no configuration required.

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| streamlit | 1.55.0 | Web dashboard framework |
| pandas | 2.3.3 | Data manipulation |
| numpy | 2.4.3 | Numeric operations |
| plotly | 6.6.0 | Interactive charts |
| openpyxl | 3.1.5 | Read .xlsx files |
| xlrd | 2.0.2 | Read legacy .xls files |
