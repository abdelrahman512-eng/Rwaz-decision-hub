import streamlit as st
import pandas as pd
import numpy as np
import scipy.optimize as opt
import numpy_financial as npf
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
import os
from datetime import datetime
from zoneinfo import ZoneInfo

# ==============================================================================
# GLOBAL BRANDING & MODERN EXECUTIVE THEME SYSTEM (RWAZ VIEW THEME)
# ==============================================================================
st.set_page_config(
    page_title="RWAZ VIEW — Executive Decision Support Platform",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium RWAZ Executive Theme — brand-led, Arabic-first, compact
st.markdown("""
<style>
    :root {
        --rwaz-primary: #684929;
        --rwaz-dark: #3F2D1E;
        --rwaz-mid: #8D765E;
        --rwaz-accent: #C5A477;
        --rwaz-bg: #F7F5F2;
        --rwaz-card: #FFFFFF;
        --rwaz-border: #E3DDD5;
        --rwaz-text: #1F2937;
        --rwaz-muted: #667085;
        --rwaz-green: #1F7A55;
        --rwaz-amber: #B7791F;
        --rwaz-red: #C53030;
    }

    /* Clean shell */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp {
        background: var(--rwaz-bg);
        color: var(--rwaz-text);
        font-family: Tahoma, "Segoe UI", Arial, sans-serif;
    }
    [data-testid="stHeader"] { display:none !important; height:0 !important; min-height:0 !important; }
    [data-testid="stToolbar"] { display:none !important; }
    [data-testid="stAppViewContainer"] > .main { padding-top:0 !important; margin-top:0 !important; }

    /* Global outer page margins — increased from the current compact values */
    .main .block-container,
    [data-testid="stMainBlockContainer"],
    [data-testid="stAppViewBlockContainer"] {
        /* Current top value was 0; use a small explicit executive breathing space. */
        padding-top: .90rem !important;
        padding-bottom: .5rem !important;
        /* .09rem × 1.60 = .144rem */
        padding-left: .9rem !important;
        padding-right: .9rem !important;
        margin-top: 0 !important;
        max-width: 100% !important;
    }
    div[data-testid="stVerticalBlock"] { gap: .24rem !important; }
    .stMarkdown, div[data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }

    /* Arabic-first direction without reversing numeric fields */
    div[data-testid="stMarkdownContainer"],
    .page-title, .page-subtitle, .section-title, .term-card,
    [data-testid="stAlert"], [data-testid="stExpander"],
    [data-testid="stSelectbox"], [data-testid="stNumberInput"] label,
    [data-testid="stTabs"] { direction: rtl; text-align: right; }
    input, .ltr-num, .term-code { direction: ltr !important; unicode-bidi: isolate; }

    /* Titles */
    .page-title {
        font-size: 19px; font-weight: 800; color: #FFF !important;
        background: linear-gradient(90deg, var(--rwaz-dark), var(--rwaz-primary));
        padding: 7px 13px; border-radius: 9px; letter-spacing: -.2px;
        box-shadow: 0 3px 10px rgba(63,45,30,.10);
    }
    .page-subtitle {
        font-size: 10.5px; color: var(--rwaz-muted); font-weight: 600;
        padding: 3px 2px 5px; border-bottom: 1px solid var(--rwaz-border);
        margin-bottom: 3px;
    }
    .section-title {
        font-size: 12.5px; font-weight: 800; color: var(--rwaz-primary);
        margin-top: 6px; margin-bottom: 4px;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #322318 0%, var(--rwaz-dark) 55%, #2D2017 100%) !important;
        border-right: 1px solid #503A27 !important;
    }
    [data-testid="stSidebar"] > div:first-child { padding-top: .65rem !important; }
    .rwaz-logo-card {
        background:#FFF; border-radius:12px; padding:8px 10px 6px; margin:2px 8px 8px;
        box-shadow:0 4px 15px rgba(0,0,0,.16); text-align:center;
    }
    .rwaz-logo-card img { max-width:132px; width:74%; height:auto; display:block; margin:auto; }
    .sidebar-tagline {
        color:#E9DED2; font-size:10px; font-weight:600; text-align:center;
        padding:0 8px 7px; direction:rtl;
    }
    [data-testid="stSidebar"] .stRadio > label {
        font-size:12px !important; font-weight:800 !important; color:#FFF !important;
        direction:rtl; text-align:right;
    }
    [data-testid="stSidebar"] .stCaption,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span { color:#F4EEE8 !important; }
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        padding:7px 10px !important; border-radius:8px !important;
        margin-bottom:3px !important; width:100% !important;
        background:transparent !important; transition:.15s ease;
        direction:rtl; text-align:right;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover { background:#553B27 !important; }
    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked),
    [data-testid="stSidebar"] div[role="radiogroup"] label[aria-checked="true"],
    [data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] {
        background:#FFF !important; border:1px solid #E7DCCF !important;
        box-shadow:0 3px 8px rgba(0,0,0,.15) !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) span,
    [data-testid="stSidebar"] div[role="radiogroup"] label[aria-checked="true"] span,
    [data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] span {
        color:var(--rwaz-dark) !important; font-weight:900 !important;
    }

    /* KPI / cards */
    .kpi-container, .combined-card {
        background:var(--rwaz-card); border:1px solid var(--rwaz-border);
        border-radius:9px; padding:7px 10px; text-align:right;
        box-shadow:0 2px 8px rgba(63,45,30,.045); height:100%;
    }
    .kpi-container { min-height:78px; }
    .kpi-title { font-size:9.5px; color:var(--rwaz-mid); font-weight:800; text-transform:none; }
    .kpi-value {
        font-size:17px; color:var(--rwaz-text); font-weight:900;
        margin:2px 0; font-variant-numeric:tabular-nums;
    }
    .kpi-sub { font-size:9.5px; font-weight:700; }
    .kpi-sub-positive { color:var(--rwaz-green); }
    .kpi-sub-warning { color:var(--rwaz-amber); }
    .kpi-sub-danger { color:var(--rwaz-red); }
    /* Compact Executive cards — Page 1 */
    .combined-card {
        padding:5px 8px !important;
        min-height:0 !important;
    }
    .combined-header {
        border-bottom:1px solid #EFE9E2;
        padding-bottom:3px;
        margin-bottom:4px;
    }
    .combined-title { font-size:9px; color:var(--rwaz-mid); font-weight:800; }
    .combined-value {
        font-size:17px; line-height:1.05; color:var(--rwaz-text); font-weight:900;
        font-variant-numeric:tabular-nums;
    }
    .combined-sub {
        font-size:8.3px; line-height:1.15; color:var(--rwaz-primary); font-weight:700;
    }
    .mini-cell { background:#FBF9F7; border:1px solid #EEE6DD; border-radius:7px; padding:5px 7px; text-align:right; }
    .combined-card .mini-cell {
        border-radius:6px; padding:3px 5px; min-height:0; line-height:1.15;
    }
    .gauge-bar-bg {
        background:#ECE6DF; border-radius:8px; height:6px; width:100%; overflow:hidden; margin-top:4px;
    }
    .gauge-bar-fill {
        background:linear-gradient(90deg, var(--rwaz-primary), var(--rwaz-accent));
        height:100%; border-radius:8px;
    }

    /* Decision summary */
    .decision-box {
        border-radius:9px; padding:9px 12px; margin:5px 0 2px; direction:rtl; text-align:right;
        border:1px solid var(--rwaz-border); background:#FFF;
    }
    .decision-pass { border-right:4px solid var(--rwaz-green); background:#F5FBF8; }
    .decision-watch { border-right:4px solid var(--rwaz-amber); background:#FFFBF2; }
    .decision-fail { border-right:4px solid var(--rwaz-red); background:#FFF6F5; }
    .decision-title { font-size:11px; font-weight:900; color:var(--rwaz-text); margin-bottom:2px; }
    .decision-text { font-size:10.5px; font-weight:600; color:#4B5563; line-height:1.65; }

    /* Status tags */
    .status-pass, .status-watch, .status-fail { font-weight:900; padding:3px 9px; border-radius:6px; display:inline-block; font-size:11px; direction:ltr; }
    .status-pass { background:#E6F5EE; color:#166A49; border:1px solid #73B99B; }
    .status-watch { background:#FFF5D9; color:#9B6518; border:1px solid #D8AA57; }
    .status-fail { background:#FDE9E7; color:#A82D2D; border:1px solid #DD8580; }

    /* Inputs: compact, clearer and aligned */
    [data-testid="stNumberInput"] label,
    [data-testid="stSelectbox"] label {
        color:var(--rwaz-primary) !important; font-size:10.5px !important; font-weight:800 !important;
        min-height:24px; display:flex; align-items:flex-end;
    }
    [data-testid="stNumberInput"] div[data-baseweb="input"],
    [data-testid="stSelectbox"] div[data-baseweb="select"] {
        background:#EDE8E2 !important; border:1px solid #D7CDC2 !important; border-radius:8px !important;
    }
    [data-testid="stNumberInput"] input {
        color:#1F2937 !important; font-weight:800 !important; font-size:12px !important;
        font-variant-numeric:tabular-nums;
    }
    [data-testid="stNumberInput"] div[data-baseweb="input"]:focus-within,
    [data-testid="stSelectbox"] div[data-baseweb="select"]:focus-within {
        border-color:var(--rwaz-primary) !important; box-shadow:0 0 0 1px rgba(104,73,41,.12) !important;
    }
    div[data-baseweb="select"] span, div[data-baseweb="menu"] div,
    li[role="option"], div[role="combobox"] { color:#1F2937 !important; font-weight:700 !important; font-size:11px !important; }

    /* Tabs / expander */
    button[data-baseweb="tab"] { font-size:11px !important; font-weight:800 !important; }
    button[data-baseweb="tab"][aria-selected="true"] { color:var(--rwaz-primary) !important; }
    [data-testid="stExpander"] { background:#FBF9F7; border:1px solid var(--rwaz-border) !important; border-radius:9px !important; }
    [data-testid="stExpander"] summary { font-weight:800; color:var(--rwaz-text); font-size:11.5px; }

    /* Tables: full-height HTML rendering — no internal vertical scroll */
    .rwaz-table-wrap {
        width:100%; overflow:visible; background:#FFF;
        border:1px solid var(--rwaz-border); border-radius:9px;
        box-shadow:0 2px 7px rgba(63,45,30,.035); margin:2px 0 5px;
    }
    .rwaz-html-table {
        width:100% !important; table-layout:fixed; border-collapse:separate !important;
        border-spacing:0 !important; margin:0 !important; background:#FFF;
        font-variant-numeric:tabular-nums;
    }
    .rwaz-html-table th {
        background:#3F2D1E !important; color:#FFF !important; font-weight:900 !important;
        text-align:center !important; padding:2px 3px !important; line-height:1.15 !important; height:27px !important;
        border-right:1px solid #6E5238 !important; border-bottom:1px solid #6E5238 !important;
        white-space:normal !important; overflow-wrap:anywhere; unicode-bidi:plaintext;
    }
    .rwaz-html-table td {
        padding:1px 3px !important; line-height:1.15 !important; height:25px !important; text-align:center !important;
        border-right:1px solid #ECE7E1 !important; border-bottom:1px solid #ECE7E1 !important;
        white-space:normal !important; overflow-wrap:anywhere; unicode-bidi:plaintext;
    }
    .rwaz-html-table tr:last-child td { border-bottom:0 !important; }
    .rwaz-table-wrap table tr:first-child th:first-child { border-top-left-radius:8px; }
    .rwaz-table-wrap table tr:first-child th:last-child { border-top-right-radius:8px; }

    /* Compact executive alerts */
    .exec-alert {
        direction:rtl; text-align:right; display:flex; align-items:center; gap:6px;
        min-height:58px; height:100%; padding:6px 7px; margin:2px 0; border-radius:8px;
        font-size:9px; font-weight:800; line-height:1.35; border:1px solid transparent;
        box-sizing:border-box;
    }
    .exec-alert .alert-dot { width:7px; height:7px; border-radius:50%; flex:0 0 7px; }
    .exec-alert-danger { background:#FCE8E6; color:#9F231F; border-color:#F4C9C5; }
    .exec-alert-danger .alert-dot { background:#C53030; }
    .exec-alert-warning { background:#FFF6DE; color:#805315; border-color:#F0DFB0; }
    .exec-alert-warning .alert-dot { background:#B7791F; }
    .exec-alert-success { background:#EAF6F0; color:#155F43; border-color:#C7E7D8; }
    .exec-alert-success .alert-dot { background:#1F7A55; }
    [data-testid="stAlert"] { border-radius:8px !important; padding:6px 9px !important; min-height:0 !important; }
    [data-testid="stAlert"] > div { padding:0 !important; min-height:0 !important; }
    [data-testid="stAlert"] p { font-size:10.5px !important; line-height:1.4 !important; font-weight:700 !important; margin:0 !important; }

    /* Plotly charts as executive cards */
    div[data-testid="stPlotlyChart"] {
        background:#FFF; border:1px solid var(--rwaz-border); border-radius:10px;
        box-shadow:0 2px 8px rgba(63,45,30,.04); padding:3px 5px 1px;
    }

    .negative-value { color:#C53030 !important; }
    .revenue-insight {
        background:#FBF6EF; border:1px solid #E8D8C5; border-radius:8px;
        padding:6px 9px; margin-top:4px; direction:rtl; text-align:right;
        color:#684929; font-size:10px; font-weight:800;
    }

    /* Approved Executive Revenue Mix card */
    .revenue-mix-card {
        background:#FFF; border:1px solid var(--rwaz-border); border-radius:10px;
        padding:8px 10px; box-shadow:0 2px 8px rgba(63,45,30,.045); direction:rtl;
    }
    .revenue-mix-title { font-size:12px; font-weight:900; color:var(--rwaz-dark); text-align:right; margin-bottom:1px; }
    .revenue-mix-subtitle { font-size:8.8px; font-weight:650; color:#8A8178; text-align:right; margin-bottom:6px; }
    .revenue-mix-kpis { display:grid; grid-template-columns:repeat(3,1fr); gap:5px; margin-bottom:7px; }
    .revenue-mix-kpi { text-align:center; border-left:1px solid #EEE7DF; padding:2px 4px; min-width:0; }
    .revenue-mix-kpi:last-child { border-left:0; }
    .revenue-mix-kpi-label { color:#756A60; font-size:8.2px; font-weight:750; white-space:nowrap; }
    .revenue-mix-kpi-value { color:#3F2D1E; font-size:14px; font-weight:950; margin-top:1px; white-space:nowrap; }
    .revenue-mix-kpi-value.main-source { color:#1F7A55; }
    .revenue-mix-stack { display:flex; direction:rtl; width:100%; height:20px; overflow:hidden; border-radius:6px; background:#F1ECE6; border:1px solid #E6DDD4; margin:3px 0 5px; }
    .revenue-mix-stack-seg { height:100%; display:flex; align-items:center; justify-content:center; color:#FFF; font-size:8px; font-weight:900; min-width:2px; overflow:hidden; white-space:nowrap; }
    .revenue-mix-legend { display:flex; flex-wrap:wrap; gap:4px 9px; justify-content:flex-start; margin-bottom:5px; }
    .revenue-mix-legend-item { display:flex; align-items:center; gap:4px; color:#5E554D; font-size:7.8px; font-weight:700; }
    .revenue-mix-dot { width:7px; height:7px; border-radius:2px; flex:0 0 7px; }
    .revenue-mix-detail-head { display:grid; grid-template-columns:1.25fr 2.2fr .9fr .65fr; gap:5px; color:#7A7066; font-size:7.8px; font-weight:800; padding:2px 2px 3px; border-bottom:1px solid #EEE7DF; }
    .revenue-mix-row { display:grid; grid-template-columns:1.25fr 2.2fr .9fr .65fr; gap:5px; align-items:center; min-height:23px; border-bottom:1px solid #F1ECE7; font-size:8.3px; }
    .revenue-mix-row:last-child { border-bottom:0; }
    .revenue-mix-name { font-weight:850; color:#3F2D1E; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .revenue-mix-bar-bg { height:7px; border-radius:5px; background:#F0EBE5; overflow:hidden; direction:rtl; }
    .revenue-mix-bar { height:100%; border-radius:5px; min-width:2px; }
    .revenue-mix-amount { direction:ltr; unicode-bidi:isolate; text-align:center; font-weight:800; color:#3F2D1E; white-space:nowrap; }
    .revenue-mix-share { direction:ltr; unicode-bidi:isolate; text-align:center; font-weight:900; white-space:nowrap; }

    /* Development portfolio summary — fills the space below the projects table */
    .dev-summary-wrap {
        margin-top:5px;
        direction:rtl;
        text-align:right;
    }
    .dev-summary-kpis {
        display:grid;
        grid-template-columns:repeat(2, minmax(0, 1fr));
        gap:6px;
        margin-bottom:6px;
    }
    .dev-summary-card {
        background:#FFF;
        border:1px solid var(--rwaz-border);
        border-radius:8px;
        padding:7px 9px;
        min-height:58px;
        box-shadow:0 2px 7px rgba(63,45,30,.035);
        display:flex;
        flex-direction:column;
        justify-content:center;
    }
    .dev-summary-label {
        font-size:8.5px;
        font-weight:800;
        color:#756A60;
        margin-bottom:3px;
    }
    .dev-summary-value {
        font-size:14px;
        line-height:1.1;
        font-weight:950;
        color:#3F2D1E;
        font-variant-numeric:tabular-nums;
    }
    .dev-summary-note {
        margin-top:2px;
        font-size:8px;
        font-weight:750;
        color:#8D765E;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
    }
    .dev-cost-mix-card {
        background:#FFF;
        border:1px solid var(--rwaz-border);
        border-radius:8px;
        padding:7px 9px 6px;
        box-shadow:0 2px 7px rgba(63,45,30,.035);
    }
    .dev-cost-mix-title {
        font-size:9px;
        font-weight:900;
        color:#684929;
        margin-bottom:5px;
    }
    .dev-cost-stack {
        display:flex;
        width:100%;
        height:20px;
        overflow:hidden;
        border-radius:6px;
        background:#F1ECE6;
        border:1px solid #E6DDD4;
        direction:rtl;
    }
    .dev-cost-seg {
        height:100%;
        min-width:2px;
        display:flex;
        align-items:center;
        justify-content:center;
        font-size:7.6px;
        font-weight:900;
        white-space:nowrap;
        overflow:hidden;
    }
    .dev-cost-legend {
        display:flex;
        flex-wrap:wrap;
        justify-content:flex-start;
        gap:3px 8px;
        margin-top:5px;
    }
    .dev-cost-legend-item {
        display:flex;
        align-items:center;
        gap:4px;
        color:#625B54;
        font-size:7.6px;
        font-weight:750;
    }
    .dev-cost-dot {
        width:7px;
        height:7px;
        border-radius:2px;
        flex:0 0 7px;
    }

    /* Guidance cards */
    .term-card {
        background:#FFF; border:1px solid var(--rwaz-border); border-radius:9px;
        padding:11px 13px; margin-bottom:7px; text-align:right;
        box-shadow:0 2px 7px rgba(63,45,30,.035); direction:rtl;
    }
    .term-title { font-size:13px; font-weight:900; color:var(--rwaz-primary); border-bottom:1px solid #EEE7DF; padding-bottom:4px; }
    .term-en { font-size:9.5px; color:#8A8178; margin-top:2px; direction:ltr; text-align:right; }
    .term-desc { font-size:10.5px; color:#374151; margin-top:6px; line-height:1.75; }
    .term-code { font-weight:900; color:var(--rwaz-primary); }

    /* Buttons */
    .stButton > button {
        border-radius:8px; border:1px solid #D9CDBF; color:var(--rwaz-primary);
        font-weight:800; background:#FFF; font-size:10.5px;
    }
    .stButton > button:hover { border-color:var(--rwaz-primary); color:var(--rwaz-dark); }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CENTRALIZED ACCOUNTING FORMATTING & RED NEGATIVES STYLING
# ==============================================================================
def fmt_num(val):
    if pd.isna(val) or val is None or val == "":
        return ""
    try:
        val = float(val)
        if val < 0:
            return f"({abs(val):,.0f})"
        return f"{val:,.0f}"
    except ValueError:
        return str(val)

def fmt_currency(val, show_symbol=True):
    if pd.isna(val) or val is None or val == "":
        return ""
    try:
        val = float(val)
        prefix = "SAR " if show_symbol else ""
        if val < 0:
            return f"{prefix}({abs(val):,.0f})"
        return f"{prefix}{val:,.0f}"
    except ValueError:
        return str(val)

def fmt_currency_m(val, show_symbol=True):
    if pd.isna(val) or val is None or val == "":
        return ""
    try:
        val = float(val)
        prefix = "SAR " if show_symbol else ""
        m_val = val / 1e6
        if val < 0:
            return f"{prefix}({abs(m_val):,.2f}M)"
        return f"{prefix}{m_val:,.2f}M"
    except ValueError:
        return str(val)

def fmt_pct(val, decimals=1):
    if pd.isna(val) or val is None or val == "":
        return ""
    try:
        val = float(val)
        if val < 0:
            return f"({abs(val)*100:.{decimals}f}%)"
        return f"{val*100:.{decimals}f}%"
    except ValueError:
        return str(val)

def fmt_multiple(val):
    if pd.isna(val) or val is None or val == "":
        return ""
    try:
        val = float(val)
        if val < 0:
            return f"({abs(val):.2f}x)"
        return f"{val:.2f}x"
    except ValueError:
        return str(val)

def fmt_currency_compact(val, decimals=2):
    """Compact executive currency labels for charts while preserving accounting negatives."""
    if pd.isna(val) or val is None or val == "":
        return ""
    try:
        val = float(val)
        abs_val = abs(val)
        if abs_val >= 1_000_000:
            num = f"{abs_val/1_000_000:.{decimals}f}M"
        elif abs_val >= 1_000:
            num = f"{abs_val/1_000:.1f}K"
        else:
            num = f"{abs_val:,.0f}"
        return f"SAR ({num})" if val < 0 else f"SAR {num}"
    except (TypeError, ValueError):
        return str(val)


def fmt_share(value):
    try:
        value = float(value)
        if value > 0 and value < 0.001:
            return "<0.1%"
        return f"{value*100:.1f}%"
    except (TypeError, ValueError):
        return ""


def value_color_style(value):
    """Display-only color for standalone numeric values; negatives are always red."""
    try:
        return "color:#C53030 !important;" if float(value) < 0 else "color:#1F2937 !important;"
    except (TypeError, ValueError):
        return "color:#1F2937 !important;"


def recalculate_installment_days(df):
    """Recalculate days remaining using Saudi local date and due date.
    Paid installments display blank days. Source Excel data is never mutated.
    """
    if df is None or df.empty:
        return df.copy() if df is not None else pd.DataFrame()
    out = df.copy()
    due_col = 'تاريخ الاستحقاق' if 'تاريخ الاستحقاق' in out.columns else None
    if due_col is None or 'الأيام المتبقية' not in out.columns:
        return out
    today_riyadh = pd.Timestamp(datetime.now(ZoneInfo('Asia/Riyadh')).date())
    due_dates = pd.to_datetime(out[due_col], errors='coerce').dt.normalize()
    if 'المتبقي للدفعة' in out.columns:
        remaining = pd.to_numeric(out['المتبقي للدفعة'], errors='coerce').fillna(0)
    else:
        remaining = pd.Series(1, index=out.index, dtype=float)
    days = (due_dates - today_riyadh).dt.days
    out['الأيام المتبقية'] = days.where((remaining > 0) & due_dates.notna(), np.nan)
    return out


def render_compact_alert(kind, message):
    cls = 'exec-alert-danger' if kind == 'error' else ('exec-alert-warning' if kind == 'warning' else 'exec-alert-success')
    st.markdown(
        f"<div class='exec-alert {cls}'><span class='alert-dot'></span><span>{message}</span></div>",
        unsafe_allow_html=True
    )

def style_df_accounting(df):
    if df is None or df.empty:
        return pd.DataFrame()
    
    df_clean = df.copy()
    df_clean.dropna(how='all', axis=0, inplace=True)
    df_clean.dropna(how='all', axis=1, inplace=True)
    df_clean = df_clean.fillna("")
    
    for col in df_clean.columns:
        def auto_fmt(v):
            if isinstance(v, (int, float, np.number)):
                if abs(v) <= 1.0 and v != 0 and 'code' not in str(col).lower() and 'كود' not in str(col):
                    return fmt_pct(v)
                return fmt_num(v)
            return str(v) if v is not None else ""
        df_clean[col] = df_clean[col].apply(auto_fmt)
            
    df_clean = df_clean.astype(str).replace({"None": "", "nan": "", "NaN": "", "<NA>": ""})
    return df_clean.reset_index(drop=True)

def render_styled_dataframe(df, max_height=None, table_kind=None):
    """Executive HTML table renderer with zero internal vertical scrolling.
    It preserves the original DataFrame values and only changes presentation.
    Font size adapts gently to wide tables to keep the full table on the page.
    """
    if df is None or df.empty:
        return

    df_fmt = style_df_accounting(df)
    col_count = max(1, len(df_fmt.columns))
    body_font = 9.0 if col_count <= 10 else (8.8 if col_count <= 14 else 8.5)
    head_font = 9.8 if col_count <= 10 else (9.3 if col_count <= 14 else 8.9)

    def highlight_executive_rows_and_negatives(row):
        styles = [''] * len(row)
        cat_str = str(row.iloc[0]).lower()
        bg_color = ''
        font_weight = 'font-weight:700;'
        text_color = '#1F2937;'

        pnl_base_rows = {'total units', 'occupied units', 'occupancy rate', 'improvements assets'}
        if table_kind == 'pnl' and cat_str.strip() in pnl_base_rows:
            # The four operating-base rows immediately before Net Revenue form one visual block.
            bg_color = 'background-color:#E2D6CA;'
            text_color = '#3F2D1E;'
            font_weight = 'font-weight:850;'
        elif 'cash beginnin' in cat_str or 'units' in cat_str or 'occupancy' in cat_str:
            bg_color = 'background-color:#FAF8F5;'
            text_color = '#625B54;'
        elif cat_str.strip() == 'net revenue':
            bg_color = 'background-color:#DCE7EA;'
            text_color = '#274C59;'
            font_weight = 'font-weight:900;'
        elif 'cash in' in cat_str or 'revenue' in cat_str:
            bg_color = 'background-color:#F3ECE5;'
            text_color = '#684929;'
            font_weight = 'font-weight:800;'
        elif 'cash out' in cat_str or 'total opex' in cat_str:
            bg_color = 'background-color:#FDECEA;'
            text_color = '#A83232;'
            font_weight = 'font-weight:800;'
        elif 'operating income' in cat_str or 'noi' in cat_str:
            bg_color = 'background-color:#EAF6F0;'
            text_color = '#166A49;'
            font_weight = 'font-weight:800;'
        elif 'net profit' in cat_str or 'end of period' in cat_str:
            bg_color = 'background-color:#3F2D1E;'
            text_color = '#FFFFFF;'
            font_weight = 'font-weight:900;'

        for i, val in enumerate(row):
            val_str = str(val)
            if '(' in val_str or val_str.startswith('-'):
                neg_color = '#F6B1AA' if bg_color == 'background-color:#3F2D1E;' else '#C53030'
                styles[i] = f'{bg_color} color:{neg_color}; {font_weight}'
            else:
                styles[i] = f'{bg_color} color:{text_color}; {font_weight}'
        return styles

    try:
        styler = df_fmt.style.apply(highlight_executive_rows_and_negatives, axis=1)
        styler = styler.set_table_attributes('class="rwaz-html-table"')
        try:
            styler = styler.hide(axis='index')
        except Exception:
            pass
        styler = styler.set_table_styles([
            {'selector': 'th', 'props': [
                ('background-color', '#3F2D1E'), ('color', '#FFFFFF'),
                ('font-weight', '900'), ('font-size', f'{head_font}px'),
                ('text-align', 'center'), ('padding', '6px 4px')
            ]},
            {'selector': 'td', 'props': [
                ('font-size', f'{body_font}px'), ('padding', '5px 4px'),
                ('font-variant-numeric', 'tabular-nums')
            ]}
        ], overwrite=False)
        html = styler.to_html()
        st.markdown(f"<div class='rwaz-table-wrap'>{html}</div>", unsafe_allow_html=True)
    except Exception:
        # Plain HTML fallback still avoids an internal Streamlit grid scrollbar.
        html = df_fmt.to_html(index=False, escape=True, classes='rwaz-html-table', border=0)
        st.markdown(f"<div class='rwaz-table-wrap'>{html}</div>", unsafe_allow_html=True)

def render_kpi(title, value, sub_text, sub_type="positive"):
    sub_class = f"kpi-sub-{sub_type}"
    is_negative = str(value).startswith("(") or str(value).startswith("SAR (")
    val_color = "color: #C53030 !important;" if is_negative else "color: #1F2937 !important;"
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value ltr-num" style="{val_color}">{value}</div>
        <div class="kpi-sub {sub_class}">{sub_text}</div>
    </div>
    """, unsafe_allow_html=True)


def render_decision_summary(status, text):
    cls = 'decision-pass' if status == 'PASS' else ('decision-watch' if status == 'WATCH' else 'decision-fail')
    label = 'قرار مناسب' if status == 'PASS' else ('يحتاج مراجعة' if status == 'WATCH' else 'غير مناسب بالشروط الحالية')
    st.markdown(f"""
    <div class="decision-box {cls}">
        <div class="decision-title">ماذا يعني هذا القرار؟ — {label}</div>
        <div class="decision-text">{text}</div>
    </div>
    """, unsafe_allow_html=True)


def extract_property_metrics(df_pl):
    """Read project KPIs from the existing P&L layout without altering source data."""
    cols = ['Project', 'Units', 'Occupied', 'Occupancy', 'Revenue', 'NOI', 'NetProfit', 'Margin']
    if df_pl is None or df_pl.empty or len(df_pl) < 2:
        return pd.DataFrame(columns=cols)

    project_names = []
    for c in range(1, max(1, df_pl.shape[1] - 1)):
        name = str(df_pl.iloc[1, c]).strip()
        if name not in ['', 'nan', 'TOTAL', 'Category'] and not name.startswith('Unnamed'):
            project_names.append((c, name))

    def row_match(term):
        return df_pl[df_pl.iloc[:, 0].astype(str).str.contains(term, case=False, na=False, regex=False)]

    units_r = row_match('Total Units')
    occ_r = row_match('Occupied units')
    rev_r = row_match('Net Revenue')
    noi_r = row_match('Operating Income')
    profit_r = row_match('Net Profit')
    margin_r = row_match('Net Profit Margin')

    rows = []
    for c, name in project_names:
        units = pd.to_numeric(units_r.iloc[0, c], errors='coerce') if not units_r.empty else np.nan
        occupied = pd.to_numeric(occ_r.iloc[0, c], errors='coerce') if not occ_r.empty else np.nan
        revenue = pd.to_numeric(rev_r.iloc[0, c], errors='coerce') if not rev_r.empty else np.nan
        noi = pd.to_numeric(noi_r.iloc[0, c], errors='coerce') if not noi_r.empty else np.nan
        profit = pd.to_numeric(profit_r.iloc[0, c], errors='coerce') if not profit_r.empty else np.nan
        margin = pd.to_numeric(margin_r.iloc[0, c], errors='coerce') if not margin_r.empty else np.nan
        occupancy = occupied / units if pd.notna(units) and units > 0 and pd.notna(occupied) else np.nan
        rows.append({'Project': name, 'Units': units, 'Occupied': occupied, 'Occupancy': occupancy,
                     'Revenue': revenue, 'NOI': noi, 'NetProfit': profit, 'Margin': margin})
    return pd.DataFrame(rows, columns=cols)


def apply_rwaz_plot_layout(fig, height=240, showlegend=False):
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#374151', size=10, family='Tahoma'),
        margin=dict(t=18, b=25, l=20, r=20), height=height,
        showlegend=showlegend, hoverlabel=dict(font_size=11)
    )
    return fig


def render_revenue_mix_card(df_revenues):
    """Approved RWAZ Revenue Mix visualization. Presentation only; source data is not mutated."""
    if df_revenues is None or df_revenues.empty or 'المبلغ' not in df_revenues.columns or 'نوع الايراد' not in df_revenues.columns:
        return
    work = df_revenues[['نوع الايراد', 'المبلغ']].copy()
    work['amount'] = pd.to_numeric(work['المبلغ'], errors='coerce').fillna(0.0)
    work = work.groupby('نوع الايراد', as_index=False)['amount'].sum().sort_values('amount', ascending=False).reset_index(drop=True)
    total = float(work['amount'].sum())
    if total <= 0:
        return
    work['share'] = work['amount'] / total
    palette = {
        'الإيجارات': RWAZ_GREEN,
        'التطوير العقاري': RWAZ_PRIMARY,
        'المقاولات': RWAZ_ACCENT,
        'إيرادات أخري': '#D9D0C5',
        'إيرادات أخرى': '#D9D0C5'
    }
    work['color'] = [palette.get(str(v).strip(), RWAZ_MID) for v in work['نوع الايراد']]
    top = work.iloc[0]
    max_amount = max(float(work['amount'].max()), 1.0)

    stack_parts = []
    legend_parts = []
    detail_rows = []
    for _, r in work.iterrows():
        name = str(r['نوع الايراد'])
        amount = float(r['amount'])
        share = float(r['share'])
        color = str(r['color'])
        width_pct = max(share * 100, 0.18)
        stack_label = fmt_share(share) if share >= 0.035 else ''
        stack_text_color = '#3F2D1E' if color.upper() in {'#D9D0C5', '#C5A477'} else '#FFFFFF'
        stack_parts.append(
            f"<div class='revenue-mix-stack-seg' style='width:{width_pct:.4f}%;background:{color};color:{stack_text_color};'>{stack_label}</div>"
        )
        legend_parts.append(
            f"<div class='revenue-mix-legend-item'><span class='revenue-mix-dot' style='background:{color};'></span><span>{name}</span></div>"
        )
        rel_width = max(0.8, amount / max_amount * 100)
        share_color = RWAZ_GREEN if name == str(top['نوع الايراد']) else color
        detail_rows.append(
            f"<div class='revenue-mix-row'>"
            f"<div class='revenue-mix-name'>{name}</div>"
            f"<div class='revenue-mix-bar-bg'><div class='revenue-mix-bar' style='width:{rel_width:.3f}%;background:{color};'></div></div>"
            f"<div class='revenue-mix-amount'>{fmt_currency_compact(amount)}</div>"
            f"<div class='revenue-mix-share' style='color:{share_color};'>{fmt_share(share)}</div>"
            f"</div>"
        )

    card = f"""
    <div class='revenue-mix-card'>
      <div class='revenue-mix-title'>مزيج الإيرادات</div>
      <div class='revenue-mix-subtitle'>توزيع الإيرادات حسب المصدر</div>
      <div class='revenue-mix-kpis'>
        <div class='revenue-mix-kpi'><div class='revenue-mix-kpi-label'>إجمالي الإيرادات</div><div class='revenue-mix-kpi-value ltr-num'>{fmt_currency_compact(total)}</div></div>
        <div class='revenue-mix-kpi'><div class='revenue-mix-kpi-label'>المصدر الرئيسي</div><div class='revenue-mix-kpi-value main-source'>{top['نوع الايراد']}</div></div>
        <div class='revenue-mix-kpi'><div class='revenue-mix-kpi-label'>نسبة المصدر الرئيسي</div><div class='revenue-mix-kpi-value main-source ltr-num'>{fmt_share(top['share'])}</div></div>
      </div>
      <div class='revenue-mix-stack'>{''.join(stack_parts)}</div>
      <div class='revenue-mix-legend'>{''.join(legend_parts)}</div>
      <div class='revenue-mix-detail-head'><div>المصدر</div><div></div><div>القيمة</div><div>النسبة</div></div>
      {''.join(detail_rows)}
      <div class='revenue-insight'>المصدر الرئيسي للإيرادات هو <b>{top['نوع الايراد']}</b> ويمثل <span class='ltr-num'>{fmt_share(top['share'])}</span> من إجمالي الإيرادات.</div>
    </div>
    """
    st.markdown(card, unsafe_allow_html=True)


def render_dev_project_summary(df_dev_projects):
    """Executive summary below the projects table. Presentation only; source data is never mutated."""
    if df_dev_projects is None or df_dev_projects.empty:
        return

    work = df_dev_projects.copy()

    project_candidates = ['اسم المشروع', 'المشروع', 'Project', 'Project Name']
    cost_candidates = ['إجمالي التكلفة', 'اجمالي التكلفة', 'إجمالي التكلفه', 'Total Cost', 'Total cost']

    project_col = next((c for c in project_candidates if c in work.columns), None)
    cost_col = next((c for c in cost_candidates if c in work.columns), None)

    if project_col is None or cost_col is None:
        return

    work['_project_name'] = work[project_col].astype(str).str.strip()
    work['_cost'] = pd.to_numeric(work[cost_col], errors='coerce').fillna(0.0)
    work = work[(work['_project_name'] != '') & (work['_project_name'].str.lower() != 'nan')].copy()

    if work.empty:
        return

    total_cost = float(work['_cost'].sum())
    if total_cost <= 0:
        return

    highest_idx = work['_cost'].idxmax()
    highest_name = str(work.loc[highest_idx, '_project_name'])
    highest_cost = float(work.loc[highest_idx, '_cost'])

    work = work.sort_values('_cost', ascending=False).reset_index(drop=True)
    work['_share'] = work['_cost'] / total_cost

    # RWAZ executive palette; cycles safely if future projects are added.
    palette = [RWAZ_GREEN, RWAZ_PRIMARY, RWAZ_ACCENT, '#E5C99E', '#9B7A56', '#D9D0C5']
    work['_color'] = [palette[i % len(palette)] for i in range(len(work))]

    stack_parts = []
    legend_parts = []

    for _, row in work.iterrows():
        name = str(row['_project_name'])
        share = float(row['_share'])
        color = str(row['_color'])
        width_pct = max(share * 100, 0.20)
        label = f"{share*100:.1f}%" if share >= 0.055 else ''
        text_color = '#3F2D1E' if color.upper() in {'#E5C99E', '#D9D0C5', '#C5A477'} else '#FFFFFF'

        stack_parts.append(
            f"<div class='dev-cost-seg' style='width:{width_pct:.4f}%;background:{color};color:{text_color};'>{label}</div>"
        )
        legend_parts.append(
            f"<div class='dev-cost-legend-item'><span class='dev-cost-dot' style='background:{color};'></span><span>{name}</span></div>"
        )

    # Build one continuous HTML string so Streamlit never interprets the
    # distribution card as an indented Markdown code block.
    html = (
        "<div class='dev-summary-wrap'>"
            "<div class='dev-summary-kpis'>"
                "<div class='dev-summary-card'>"
                    "<div class='dev-summary-label'>إجمالي تكلفة المشاريع</div>"
                    f"<div class='dev-summary-value ltr-num'>{fmt_currency_compact(total_cost)}</div>"
                    "<div class='dev-summary-note'>إجمالي محفظة المشاريع تحت الإنشاء</div>"
                "</div>"
                "<div class='dev-summary-card'>"
                    "<div class='dev-summary-label'>أعلى مشروع تكلفة</div>"
                    f"<div class='dev-summary-value'>{highest_name}</div>"
                    f"<div class='dev-summary-note ltr-num'>{fmt_currency_compact(highest_cost)}</div>"
                "</div>"
            "</div>"
            "<div class='dev-cost-mix-card'>"
                "<div class='dev-cost-mix-title'>توزيع تكلفة المشاريع</div>"
                f"<div class='dev-cost-stack'>{''.join(stack_parts)}</div>"
                f"<div class='dev-cost-legend'>{''.join(legend_parts)}</div>"
            "</div>"
        "</div>"
    )
    st.markdown(html, unsafe_allow_html=True)


# ==============================================================================
# LAYER 1: DYNAMIC EXCEL DATA INGESTION ENGINE
# ==============================================================================
@st.cache_data
def load_and_validate_source_data():
    def find_file(candidates):
        for c in candidates:
            if os.path.exists(c):
                return c
        return None

    master_f = find_file(['Master_Financial_Data_F.xlsx', 'Master_Financial_Data_F (1).xlsx', 'Master_Financial_Data_F_2.xlsx', 'Master_Financial_Data_F_3.xlsx'])
    cf_f = find_file(['Cash Flow 12 Month.xlsx', 'Cash Flow 12 Month_2.xlsx', 'Cash Flow 24 Month.xlsx'])
    pl_f = find_file(['P&L_Rent_Projects_F.xlsx', 'P&L_Rent_Projects_F_2.xlsx', 'P&L_Rent_Projects_F_3.xlsx'])

    missing = []
    if not master_f: missing.append('Master_Financial_Data_F.xlsx')
    if not cf_f: missing.append('Cash Flow File')
    if not pl_f: missing.append('P&L_Rent_Projects_F.xlsx')

    if missing:
        st.error(f"❌ تعذر العثور على ملفات المصدر التالية في المجلد: {missing}. يرجى التثبت من وجودها.")
        st.stop()
        
    try:
        wb_m = openpyxl.load_workbook(master_f, data_only=True)
        ws_m = wb_m['Loans_&_Installments']
        
        def parse_named_table(ws, table_name):
            if table_name not in ws.tables:
                return pd.DataFrame()
            tbl = ws.tables[table_name]
            min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(tbl.ref)
            data = []
            for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
                data.append(list(row))
            headers = data[0]
            rows = data[1:]
            return pd.DataFrame(rows, columns=headers)

        df_loans = parse_named_table(ws_m, 'القروض')
        df_installments = parse_named_table(ws_m, 'الاقساط')
        df_revenues = parse_named_table(ws_m, 'الايردات')
        df_dev_projects = parse_named_table(ws_m, 'Units_Under_Construction')
        df_banks = parse_named_table(ws_m, 'البنوك')
        df_collections = parse_named_table(ws_m, 'تحصيلات_الايجار')
        df_partners = parse_named_table(ws_m, 'حساب_الشركاء')
        
        # Parse CF safely without DateParseError
        excel_cf = pd.ExcelFile(cf_f)
        df_cf_raw = excel_cf.parse(excel_cf.sheet_names[0])
        
        category_col = df_cf_raw.columns[0]
        time_cols_raw = [c for c in df_cf_raw.columns if c != category_col and 'unnamed' not in str(c).lower()]
        
        time_cols = []
        for c in time_cols_raw:
            try:
                dt = pd.to_datetime(c)
                time_cols.append(dt.strftime('%b-%y'))
            except Exception:
                time_cols.append(str(c))
                
        df_cf_clean = df_cf_raw.dropna(how='all').copy()
        df_cf_clean.rename(columns={category_col: 'Category'}, inplace=True)
        
        rename_map = dict(zip(time_cols_raw, time_cols))
        df_cf_clean.rename(columns=rename_map, inplace=True)
        
        excel_pl = pd.ExcelFile(pl_f)
        pl_sheet = 'Sheet3' if 'Sheet3' in excel_pl.sheet_names else excel_pl.sheet_names[0]
        df_pl_raw = excel_pl.parse(pl_sheet)
        
        return {
            'df_loans': df_loans,
            'df_installments': df_installments,
            'df_revenues': df_revenues,
            'df_dev_projects': df_dev_projects,
            'df_banks': df_banks,
            'df_collections': df_collections,
            'df_partners': df_partners,
            'df_cf': df_cf_clean,
            'time_cols': time_cols,
            'df_pl': df_pl_raw
        }
    except Exception as e:
        st.error(f"❌ خطأ في معالجة بيانات الإكسل: {str(e)}")
        st.stop()

store = load_and_validate_source_data()

# ==============================================================================
# LAYER 2: PURE 100% EQUITY FEASIBILITY ENGINES
# ==============================================================================
def run_dev_engine(land_price, rett_rate, dev_cost_per_sqm, sellable_area,
                   selling_price_per_sqm, dev_months, sales_months,
                   cost_of_equity, target_equity_irr, min_npv_threshold=0, watch_buffer=0.03):
    if dev_months <= 0 or sales_months <= 0 or sellable_area <= 0 or selling_price_per_sqm <= 0:
        return {'decision': 'FAIL', 'error': 'مدخلات غير صالحة'}
        
    land_total = land_price * (1 + rett_rate)
    dev_cost_total = dev_cost_per_sqm * sellable_area
    total_cost = land_total + dev_cost_total
    total_rev = selling_price_per_sqm * sellable_area
    net_profit = total_rev - total_cost
    profit_margin = net_profit / total_rev if total_rev > 0 else 0
    
    total_timeline = int(dev_months + sales_months)
    monthly_cf = np.zeros(total_timeline + 1)
    monthly_cf[0] = -land_total
    dev_m_out = dev_cost_total / dev_months
    for m in range(1, int(dev_months) + 1):
        monthly_cf[m] -= dev_m_out
    sales_m_in = total_rev / sales_months
    for m in range(int(dev_months) + 1, total_timeline + 1):
        monthly_cf[m] += sales_m_in
        
    try:
        r_m = npf.irr(monthly_cf)
        equity_irr = (1 + r_m)**12 - 1 if not np.isnan(r_m) and r_m > -1 else np.nan
    except Exception:
        equity_irr = np.nan
        
    r_disc_m = (1 + cost_of_equity)**(1/12) - 1
    equity_npv = sum([cf / ((1 + r_disc_m)**t) for t, cf in enumerate(monthly_cf)])
    
    peak_equity = abs(min(np.cumsum(monthly_cf)))
    total_inflows = sum([max(0, cf) for cf in monthly_cf])
    total_outflows = sum([abs(min(0, cf)) for cf in monthly_cf])
    equity_moic = total_inflows / total_outflows if total_outflows > 0 else 0
    
    cum_cf = np.cumsum(monthly_cf)
    payback_m = np.nan
    for t in range(len(cum_cf)):
        if cum_cf[t] >= 0:
            payback_m = (t - 1) + (-cum_cf[t-1] / (cum_cf[t] - cum_cf[t-1])) if t > 0 else 0
            break
            
    accounting_be = total_cost
    
    def sim_cf(r_val):
        cf = np.zeros(total_timeline + 1)
        cf[0] = -land_total
        for m_idx in range(1, int(dev_months) + 1):
            cf[m_idx] -= dev_m_out
        in_m = r_val / sales_months
        for m_idx in range(int(dev_months) + 1, total_timeline + 1):
            cf[m_idx] += in_m
        return cf

    def npv_func(r_val):
        cf = sim_cf(r_val)
        return sum([c / ((1 + r_disc_m)**t) for t, c in enumerate(cf)])
        
    try:
        npv_zero_rev = opt.brentq(npv_func, total_cost * 0.1, total_cost * 5.0)
    except Exception:
        npv_zero_rev = np.nan
    
    def target_irr_func(r_val):
        cf = sim_cf(r_val)
        rm = npf.irr(cf)
        irr_ann = (1 + rm)**12 - 1 if not np.isnan(rm) and rm > -1 else -0.99
        return irr_ann - target_equity_irr
        
    try:
        target_irr_rev = opt.brentq(target_irr_func, total_cost * 0.1, total_cost * 5.0)
    except Exception:
        target_irr_rev = np.nan
    
    if not np.isnan(equity_irr) and equity_irr >= target_equity_irr and equity_npv >= min_npv_threshold:
        decision = "PASS"
    elif not np.isnan(equity_irr) and equity_irr >= (target_equity_irr - watch_buffer) and equity_npv >= min_npv_threshold:
        decision = "WATCH"
    else:
        decision = "FAIL"

    return {
        'total_rev': total_rev, 'total_cost': total_cost, 'net_profit': net_profit,
        'profit_margin': profit_margin, 'peak_equity': peak_equity,
        'equity_irr': equity_irr, 'equity_npv': equity_npv, 'equity_moic': equity_moic,
        'payback_m': payback_m, 'accounting_be': accounting_be,
        'npv_zero_rev': npv_zero_rev, 'target_irr_rev': target_irr_rev,
        'req_price_npv_zero': npv_zero_rev / sellable_area if not np.isnan(npv_zero_rev) else np.nan,
        'req_price_target_irr': target_irr_rev / sellable_area if not np.isnan(target_irr_rev) else np.nan,
        'decision': decision
    }

def run_rental_engine(head_lease_rent, lease_term_yrs, rent_escalation_pct, escalation_freq_yrs,
                      grace_period_m, total_units, sub_rent_per_unit, target_occupancy,
                      opex_ratio, fitout_capex, cost_of_equity, target_equity_irr,
                      min_npv_threshold=0, watch_buffer=0.03):
    gross_potential_rev = total_units * sub_rent_per_unit
    active_m_yr1 = max(0, 12 - grace_period_m)
    actual_rev_yr1 = gross_potential_rev * target_occupancy * (active_m_yr1 / 12)
    head_rent_yr1 = head_lease_rent * (active_m_yr1 / 12)
    opex_yr1 = actual_rev_yr1 * opex_ratio
    noi_yr1 = actual_rev_yr1 - head_rent_yr1 - opex_yr1
    
    annual_pnl = []
    cf_list = [-fitout_capex]
    
    total_life_revenue = 0.0
    
    for yr in range(1, int(lease_term_yrs) + 1):
        esc_factor = (1 + rent_escalation_pct) ** ((yr - 1) // escalation_freq_yrs)
        curr_head_rent = head_lease_rent * esc_factor
        
        if yr == 1:
            rev = actual_rev_yr1
            h_rent = head_rent_yr1
        else:
            rev = gross_potential_rev * target_occupancy
            h_rent = curr_head_rent
            
        total_life_revenue += rev
        opex = rev * opex_ratio
        noi = rev - h_rent - opex
        noi_m = noi / rev if rev > 0 else 0
        
        annual_pnl.append({
            'السنة': f"السنة {yr}", 'إيراد الإيجار': rev, 'إيجار المالك': h_rent, 'التكاليف التشغيلية': opex, 'صافي الدخل NOI': noi, 'هامش NOI': noi_m
        })
        cf_list.append(noi)
        
    try:
        r_ann = npf.irr(cf_list)
        equity_irr = r_ann if not np.isnan(r_ann) else np.nan
    except Exception:
        equity_irr = np.nan
        
    equity_npv = sum([cf / ((1 + cost_of_equity)**t) for t, cf in enumerate(cf_list)])
    
    total_inflows = sum([max(0, cf) for cf in cf_list[1:]])
    total_outflows = abs(cf_list[0]) + sum([abs(min(0, cf)) for cf in cf_list[1:]])
    equity_moic = total_inflows / total_outflows if total_outflows > 0 else 0
    
    cum_cf = np.cumsum(cf_list)
    payback_yrs = np.nan
    for t in range(len(cum_cf)):
        if cum_cf[t] >= 0:
            payback_yrs = (t - 1) + (-cum_cf[t-1] / (cum_cf[t] - cum_cf[t-1])) if t > 0 else 0
            break
            
    be_occupancy = head_lease_rent / (gross_potential_rev * (1 - opex_ratio)) if gross_potential_rev > 0 else 0
    
    def occ_irr_solver(occ_val):
        test_cfs = [-fitout_capex]
        for yr in range(1, int(lease_term_yrs) + 1):
            esc = (1 + rent_escalation_pct) ** ((yr - 1) // escalation_freq_yrs)
            hr = head_lease_rent * esc
            if yr == 1:
                r_val = gross_potential_rev * occ_val * (active_m_yr1 / 12)
                hr = hr * (active_m_yr1 / 12)
            else:
                r_val = gross_potential_rev * occ_val
            op = r_val * opex_ratio
            test_cfs.append(r_val - hr - op)
        r_calc = npf.irr(test_cfs)
        if np.isnan(r_calc):
            r_calc = -0.99
        return r_calc - target_equity_irr
        
    try:
        occ_for_target_irr = opt.brentq(occ_irr_solver, 0.05, 1.0)
    except Exception:
        occ_for_target_irr = np.nan
        
    if not np.isnan(equity_irr) and equity_irr >= target_equity_irr and equity_npv >= min_npv_threshold:
        decision = "PASS"
    elif not np.isnan(equity_irr) and equity_irr >= (target_equity_irr - watch_buffer) and equity_npv >= min_npv_threshold:
        decision = "WATCH"
    else:
        decision = "FAIL"

    return {
        'gross_potential_rev': gross_potential_rev,
        'actual_rev_yr1': actual_rev_yr1,
        'total_life_revenue': total_life_revenue,
        'opex_yr1': opex_yr1,
        'noi_yr1': noi_yr1,
        'fitout_capex': fitout_capex,
        'equity_irr': equity_irr,
        'equity_npv': equity_npv,
        'equity_moic': equity_moic,
        'payback_yrs': payback_yrs,
        'be_occupancy': be_occupancy,
        'occ_for_target_irr': occ_for_target_irr,
        'decision': decision,
        'annual_pnl': pd.DataFrame(annual_pnl)
    }

# ==============================================================================
# SIDEBAR NAVIGATION — RWAZ BRAND
# ==============================================================================
st.sidebar.markdown("""
<div class="rwaz-logo-card">
    <img src="data:image/png;base64,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" alt="RWAZ Logo">
</div>
<div class="sidebar-tagline">مركز القرار الاستثماري والمالي</div>
""", unsafe_allow_html=True)

page = st.sidebar.radio("القائمة الرئيسية:", [
    "الملخص التنفيذي والمركز المالي",
    "جدول التمويلات والاقساط",
    "السيولة والتدفقات النقدية",
    "مشاريع الايجار",
    "موديل التطوير العقاري",
    "موديل الايجارات",
    "ارشادات"
])

# Shared chart palette
RWAZ_PRIMARY = "#684929"
RWAZ_DARK = "#3F2D1E"
RWAZ_MID = "#8D765E"
RWAZ_ACCENT = "#C5A477"
RWAZ_GREEN = "#1F7A55"
RWAZ_AMBER = "#B7791F"
RWAZ_RED = "#C53030"
RWAZ_PALETTE = [RWAZ_PRIMARY, RWAZ_ACCENT, RWAZ_GREEN, RWAZ_MID, "#A98D70", "#6B7280"]

# ==============================================================================
# PAGE 1: EXECUTIVE DASHBOARD
# ==============================================================================
if page == "الملخص التنفيذي والمركز المالي":
    st.markdown("<div class='page-title'>رواز | لوحة الإدارة التنفيذية والمركز المالي</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>نظرة تنفيذية مركزة على السيولة، التحصيل، الالتزامات، أداء العقارات والمشاريع تحت الإنشاء.</div>", unsafe_allow_html=True)

    df_b = store['df_banks'].copy()
    total_cash = pd.to_numeric(df_b['الرصيد'], errors='coerce').sum() if 'الرصيد' in df_b.columns else 0.0
    rajhi_cash = pd.to_numeric(df_b[df_b['البنك'].astype(str).str.contains('الراجحي', na=False)]['الرصيد'], errors='coerce').sum() if 'البنك' in df_b.columns else 0.0
    snb_cash = pd.to_numeric(df_b[df_b['البنك'].astype(str).str.contains('الأهلي', na=False)]['الرصيد'], errors='coerce').sum() if 'البنك' in df_b.columns else 0.0
    total_revenue = pd.to_numeric(store['df_revenues']['المبلغ'], errors='coerce').sum() if 'المبلغ' in store['df_revenues'].columns else 0.0

    coll_rate = pd.to_numeric(store['df_collections']['كفاءة التحصيل %'], errors='coerce').values[0] if 'كفاءة التحصيل %' in store['df_collections'].columns and not store['df_collections'].empty else 0.0
    due_coll = pd.to_numeric(store['df_collections']['المستحق للتحصيل'], errors='coerce').values[0] if 'المستحق للتحصيل' in store['df_collections'].columns and not store['df_collections'].empty else 0.0
    act_coll = pd.to_numeric(store['df_collections']['المحصل الفعلي'], errors='coerce').values[0] if 'المحصل الفعلي' in store['df_collections'].columns and not store['df_collections'].empty else 0.0
    partners_net = pd.to_numeric(store['df_partners']['الرصيد'], errors='coerce').sum() if 'الرصيد' in store['df_partners'].columns else 0.0

    col_top1, col_top2, col_top3 = st.columns([1.05, .9, 1.05])
    with col_top1:
        st.markdown(f"""
        <div class="combined-card">
            <div class="combined-header">
                <div class="combined-title">إجمالي رصيد البنوك</div>
                <div class="combined-value ltr-num" style="{value_color_style(total_cash)}">{fmt_currency(total_cash)}</div>
                <div class="combined-sub">النقدية المتاحة بالحسابات</div>
            </div>
            <div style="display:flex;gap:5px;">
                <div class="mini-cell" style="flex:1;"><div style="font-size:9px;color:#7A7066;font-weight:700;">مصرف الراجحي</div><div class="ltr-num" style="font-size:11px;font-weight:900;{value_color_style(rajhi_cash)}">{fmt_currency(rajhi_cash)}</div></div>
                <div class="mini-cell" style="flex:1;"><div style="font-size:9px;color:#7A7066;font-weight:700;">البنك الأهلي</div><div class="ltr-num" style="font-size:11px;font-weight:900;{value_color_style(snb_cash)}">{fmt_currency(snb_cash)}</div></div>
            </div>
        </div>""", unsafe_allow_html=True)

    with col_top2:
        coll_pct_str = f"{coll_rate*100:.1f}%"
        st.markdown(f"""
        <div class="combined-card">
            <div style="font-size:9px;color:#684929;font-weight:900;text-align:right;margin-bottom:3px;">أداء التحصيل</div>
            <div style="display:flex;align-items:center;gap:5px;">
                <div style="flex:1;display:flex;flex-direction:column;gap:4px;">
                    <div class="mini-cell" style="text-align:center;"><div class="ltr-num" style="font-size:11px;font-weight:900;{value_color_style(due_coll)}">{fmt_currency(due_coll)}</div><div style="font-size:8.5px;color:#7A7066;">المستحق</div></div>
                    <div class="mini-cell" style="text-align:center;"><div class="ltr-num" style="font-size:11px;font-weight:900;{value_color_style(act_coll)}">{fmt_currency(act_coll)}</div><div style="font-size:8.5px;color:#7A7066;">المحصل</div></div>
                </div>
                <div style="flex:1.05;text-align:center;"><div class="ltr-num" style="font-size:22px;font-weight:900;line-height:1;">{coll_pct_str}</div><div style="font-size:8.3px;color:#7A7066;font-weight:700;margin-top:3px;">كفاءة التحصيل</div><div class="gauge-bar-bg"><div class="gauge-bar-fill" style="width:{min(100,max(0,coll_rate*100)):.1f}%;"></div></div></div>
            </div>
        </div>""", unsafe_allow_html=True)

    with col_top3:
        df_part = store['df_partners'].copy()
        p1_name, p1_bal, p2_name, p2_bal = "شريك 1", 0.0, "شريك 2", 0.0
        if not df_part.empty:
            if len(df_part) >= 1:
                p1_name = df_part.iloc[0]['الشريك'] if 'الشريك' in df_part.columns else p1_name
                p1_bal = pd.to_numeric(df_part.iloc[0]['الرصيد'], errors='coerce') if 'الرصيد' in df_part.columns else 0.0
            if len(df_part) >= 2:
                p2_name = df_part.iloc[1]['الشريك'] if 'الشريك' in df_part.columns else p2_name
                p2_bal = pd.to_numeric(df_part.iloc[1]['الرصيد'], errors='coerce') if 'الرصيد' in df_part.columns else 0.0
        st.markdown(f"""
        <div class="combined-card">
            <div class="combined-header"><div class="combined-title">صافي أرصدة الشركاء</div><div class="combined-value ltr-num" style="{value_color_style(partners_net)}">{fmt_currency(partners_net)}</div><div class="combined-sub">إجمالي صافي حسابات الشركاء</div></div>
            <div style="display:flex;gap:5px;">
                <div class="mini-cell" style="flex:1;"><div style="font-size:9px;color:#7A7066;font-weight:700;">{p1_name}</div><div class="ltr-num" style="font-size:11px;font-weight:900;{value_color_style(p1_bal)}">{fmt_currency(p1_bal)}</div></div>
                <div class="mini-cell" style="flex:1;"><div style="font-size:9px;color:#7A7066;font-weight:700;">{p2_name}</div><div class="ltr-num" style="font-size:11px;font-weight:900;{value_color_style(p2_bal)}">{fmt_currency(p2_bal)}</div></div>
            </div>
        </div>""", unsafe_allow_html=True)

    # Executive alerts — concise, combined and prioritized
    alerts = []
    df_inst = recalculate_installment_days(store['df_installments'])
    if 'الأيام المتبقية' in df_inst.columns and 'المتبقي للدفعة' in df_inst.columns:
        days = pd.to_numeric(df_inst['الأيام المتبقية'], errors='coerce')
        rem = pd.to_numeric(df_inst['المتبقي للدفعة'], errors='coerce').fillna(0)
        overdue_df = df_inst[(days < 0) & (rem > 0)].copy()
        if not overdue_df.empty:
            overdue_total = pd.to_numeric(overdue_df['المتبقي للدفعة'], errors='coerce').sum()
            if len(overdue_df) == 1:
                desc = overdue_df.iloc[0]['بيان الدفعة'] if 'بيان الدفعة' in overdue_df.columns else 'دفعة تمويل'
                alerts.append((1, 'error', f"تنبيه تمويلي حرج: {desc} متأخر السداد — المتبقي {fmt_currency(overdue_total)}."))
            else:
                alerts.append((1, 'error', f"تنبيه تمويلي حرج: {len(overdue_df)} دفعات متأخرة بإجمالي متبقٍ {fmt_currency(overdue_total)}."))

    if total_cash < 500000:
        alerts.append((1, 'error', f"السيولة الحالية {fmt_currency(total_cash)} أقل من حد الأمان التشغيلي {fmt_currency(500000)}."))

    df_cf_alert = store['df_cf']
    tcols_alert = store['time_cols']
    if tcols_alert:
        ending_alert = df_cf_alert[df_cf_alert['Category'].astype(str).str.contains('end of period', case=False, na=False)]
        vals_alert = ending_alert[tcols_alert].values.flatten() if not ending_alert.empty else df_cf_alert[tcols_alert].iloc[-1].values.flatten()
        vals_alert = np.nan_to_num(pd.to_numeric(vals_alert, errors='coerce'), nan=0.0)
        if len(vals_alert):
            min_i = int(np.argmin(vals_alert))
            if vals_alert[min_i] < 0:
                alerts.append((1, 'error', f"عجز نقدي متوقع: يصل الرصيد إلى {fmt_currency(vals_alert[min_i])} في {tcols_alert[min_i]}."))

    if coll_rate < 0.60:
        alerts.append((1, 'error', f"كفاءة التحصيل منخفضة بشدة عند {fmt_pct(coll_rate)} مقابل المستوى الإداري المستهدف."))
    elif coll_rate < 0.80:
        alerts.append((2, 'warning', f"كفاءة التحصيل منخفضة عند {fmt_pct(coll_rate)} وتحتاج متابعة."))

    prop_metrics_alert = extract_property_metrics(store['df_pl'])
    if not prop_metrics_alert.empty:
        neg_margin = prop_metrics_alert[pd.to_numeric(prop_metrics_alert['Margin'], errors='coerce') < 0].sort_values('Margin')
        if not neg_margin.empty:
            worst = neg_margin.iloc[0]
            alerts.append((2, 'warning', f"{len(neg_margin)} عقار/مشروع يحقق صافي هامش ربح سالب؛ الأدنى {worst['Project']} عند {fmt_pct(worst['Margin'])}."))
        low_occ = prop_metrics_alert[pd.to_numeric(prop_metrics_alert['Occupancy'], errors='coerce') < 0.97].sort_values('Occupancy')
        if not low_occ.empty:
            worst_o = low_occ.iloc[0]
            alerts.append((2, 'warning', f"{len(low_occ)} عقار/مشروع أقل من مستهدف الإشغال 97%؛ الأدنى {worst_o['Project']} عند {fmt_pct(worst_o['Occupancy'])}."))

    st.markdown("<div class='section-title'>التنبيهات والإجراءات الإدارية المباشرة</div>", unsafe_allow_html=True)
    alerts = sorted(alerts, key=lambda x: x[0])

    if not alerts:
        render_compact_alert('success', "لا توجد تنبيهات حرجة وفق مؤشرات الصفحة الحالية.")
    else:
        visible_alerts = alerts[:5]

        # Executive alert row: show the most important alerts side-by-side
        alert_cols = st.columns(len(visible_alerts), gap="small")
        for col, (_, kind, msg) in zip(alert_cols, visible_alerts):
            with col:
                render_compact_alert(kind, msg)

        if len(alerts) > 5:
            with st.expander(f"عرض جميع التنبيهات ({len(alerts)})", expanded=False):
                for _, kind, msg in alerts[5:]:
                    render_compact_alert(kind, msg)

    st.markdown("<div class='section-title'>المشاريع تحت الإنشاء ومزيج الإيرادات</div>", unsafe_allow_html=True)
    c1_page1, c2_page1 = st.columns([1.0, 1.0])
    with c1_page1:
        render_styled_dataframe(store['df_dev_projects'].copy(), max_height=245)
        render_dev_project_summary(store['df_dev_projects'].copy())
    with c2_page1:
        render_revenue_mix_card(store['df_revenues'].copy())

# ==============================================================================
# PAGE 2: FINANCING & INSTALLMENTS
# ==============================================================================
elif page == "جدول التمويلات والاقساط":
    st.markdown("<div class='page-title'>التمويلات والالتزامات وجدول الأقساط</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>متابعة أصل التمويلات، المدفوع والمتبقي مع الاحتفاظ بالتفاصيل المحاسبية الكاملة.</div>", unsafe_allow_html=True)

    df_loans_disp = store['df_loans'].copy()
    total_debt_orig = pd.to_numeric(df_loans_disp['أصل التمويل'], errors='coerce').sum() if 'أصل التمويل' in df_loans_disp.columns else 0.0
    total_debt_rem = pd.to_numeric(df_loans_disp['المتبقي للقرض'], errors='coerce').sum() if 'المتبقي للقرض' in df_loans_disp.columns else 0.0
    total_paid = pd.to_numeric(df_loans_disp['إجمالي المدفوع'], errors='coerce').sum() if 'إجمالي المدفوع' in df_loans_disp.columns else 0.0
    active_count = int((pd.to_numeric(df_loans_disp['المتبقي للقرض'], errors='coerce').fillna(0) > 0).sum()) if 'المتبقي للقرض' in df_loans_disp.columns else 0

    d1, d2, d3, d4 = st.columns(4)
    with d1: render_kpi("أصل التمويلات", fmt_currency(total_debt_orig), "إجمالي أصل التمويلات المسجلة", "positive")
    with d2: render_kpi("إجمالي المدفوع", fmt_currency(total_paid), "المدفوع حتى الآن", "positive")
    with d3: render_kpi("المتبقي للتمويلات", fmt_currency(total_debt_rem), "الرصيد المتبقي حالياً", "warning" if total_debt_rem > 0 else "positive")
    with d4: render_kpi("التمويلات النشطة", f"{active_count}", f"من إجمالي {len(df_loans_disp)}", "warning" if active_count else "positive")

    tab_over, tab_loans, tab_inst = st.tabs(["نظرة تنفيذية", "القروض والتسهيلات", "جدول الأقساط"])
    with tab_over:
        if 'المتبقي للقرض' in df_loans_disp.columns and 'جهة التمويل' in df_loans_disp.columns:
            chart_debt = df_loans_disp.copy()
            chart_debt['المتبقي_الرقمي'] = pd.to_numeric(chart_debt['المتبقي للقرض'], errors='coerce').fillna(0)
            chart_debt['أصل_رقمي'] = pd.to_numeric(chart_debt['أصل التمويل'], errors='coerce').fillna(0) if 'أصل التمويل' in chart_debt.columns else 0
            chart_debt = chart_debt.sort_values('المتبقي_الرقمي', ascending=False).reset_index(drop=True)
            debt_colors = [RWAZ_GREEN if v <= 0 else RWAZ_PRIMARY for v in chart_debt['المتبقي_الرقمي']]
            debt_text = ['تم السداد بالكامل' if v <= 0 else fmt_currency_compact(v) for v in chart_debt['المتبقي_الرقمي']]
            fig_debt = go.Figure(go.Bar(
                x=chart_debt['المتبقي_الرقمي'], y=chart_debt['جهة التمويل'], orientation='h',
                marker_color=debt_colors, text=debt_text, textposition='outside', cliponaxis=False,
                customdata=np.stack([chart_debt['أصل_رقمي']], axis=-1),
                hovertemplate='%{y}<br>المتبقي: SAR %{x:,.0f}<br>أصل التمويل: SAR %{customdata[0]:,.0f}<extra></extra>'
            ))
            max_debt = max(float(chart_debt['المتبقي_الرقمي'].max()), 1.0)
            apply_rwaz_plot_layout(fig_debt, height=max(210, 55 + 42*len(chart_debt)))
            fig_debt.update_layout(
                xaxis=dict(title='', range=[0, max_debt*1.28], showgrid=False, zeroline=False, showticklabels=False),
                yaxis=dict(title='', autorange='reversed', tickfont=dict(size=10)),
                margin=dict(t=8,b=8,l=20,r=100), bargap=.34
            )
            st.markdown("<div class='section-title'>الرصيد المتبقي حسب جهة التمويل</div>", unsafe_allow_html=True)
            st.markdown(f"<div style='direction:rtl;text-align:right;font-size:10px;color:#684929;font-weight:800;margin-bottom:3px;'>إجمالي الدين المتبقي: <span class='ltr-num'>{fmt_currency_compact(total_debt_rem)}</span></div>", unsafe_allow_html=True)
            st.plotly_chart(fig_debt, use_container_width=True, config={'displayModeBar': False})
    with tab_loans:
        render_styled_dataframe(df_loans_disp, max_height=500)
    with tab_inst:
        df_inst_display = recalculate_installment_days(store['df_installments'])
        render_styled_dataframe(df_inst_display, max_height=620)

# ==============================================================================
# PAGE 3: CASH FLOW & LIQUIDITY
# ==============================================================================
elif page == "السيولة والتدفقات النقدية":
    st.markdown("<div class='page-title'>التدفقات النقدية والسيولة</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>مسار السيولة المتوقعة، أدنى وأعلى نقطة نقدية والتدفقات الداخلة والخارجة.</div>", unsafe_allow_html=True)

    df_cf = store['df_cf'].copy()
    time_cols = store['time_cols']
    ending_row = df_cf[df_cf['Category'].astype(str).str.contains('end of period', case=False, na=False)]
    ending_cash_vals = ending_row[time_cols].values.flatten() if not ending_row.empty else df_cf[time_cols].iloc[-1].values.flatten()
    ending_cash_vals = np.nan_to_num(pd.to_numeric(ending_cash_vals, errors='coerce'), nan=0.0)
    min_idx = int(np.argmin(ending_cash_vals)) if len(ending_cash_vals) else 0
    max_idx = int(np.argmax(ending_cash_vals)) if len(ending_cash_vals) else 0
    min_cash_val = ending_cash_vals[min_idx] if len(ending_cash_vals) else 0.0
    max_cash_val = ending_cash_vals[max_idx] if len(ending_cash_vals) else 0.0
    min_cash_month = time_cols[min_idx] if time_cols else ""
    max_cash_month = time_cols[max_idx] if time_cols else ""

    outflow_row = df_cf[df_cf['Category'].astype(str).str.contains('cash out', case=False, na=False)]
    if outflow_row.empty:
        outflow_row = df_cf[df_cf['Category'].astype(str).str.fullmatch('out', case=False, na=False)]
    outflow_90d = abs(pd.to_numeric(outflow_row[time_cols[:min(3, len(time_cols))]].values.flatten(), errors='coerce').sum()) if not outflow_row.empty else 0.0

    c1, c2, c3, c4 = st.columns(4)
    current_cash = pd.to_numeric(store['df_banks']['الرصيد'], errors='coerce').sum() if 'الرصيد' in store['df_banks'].columns else 0.0
    with c1: render_kpi("السيولة الحالية", fmt_currency(current_cash), "النقدية المتاحة بالبنوك", "positive" if current_cash >= 500000 else "danger")
    with c2: render_kpi("أدنى نقطة سيولة", fmt_currency(min_cash_val), f"{min_cash_month}", "danger" if min_cash_val < 0 else ("warning" if min_cash_val < 500000 else "positive"))
    with c3: render_kpi("أعلى نقطة سيولة", fmt_currency(max_cash_val), f"{max_cash_month}", "positive")
    with c4: render_kpi("التدفقات الخارجة لـ 90 يوماً", fmt_currency(outflow_90d), "وفق جدول التدفقات الحالي", "warning")

    tab_cf_over, tab_cf_detail = st.tabs(["نظرة السيولة", "جدول التدفقات التفصيلي"])
    with tab_cf_over:
        cash_in_row = df_cf[df_cf['Category'].astype(str).str.contains('cash in', case=False, na=False)]
        cash_out_row = df_cf[df_cf['Category'].astype(str).str.contains('cash out', case=False, na=False)]
        cash_in_vals = np.nan_to_num(pd.to_numeric(cash_in_row[time_cols].values.flatten()[:len(time_cols)], errors='coerce'), nan=0.0) if not cash_in_row.empty else np.zeros(len(time_cols))
        cash_out_vals = np.nan_to_num(pd.to_numeric(cash_out_row[time_cols].values.flatten()[:len(time_cols)], errors='coerce'), nan=0.0) if not cash_out_row.empty else np.zeros(len(time_cols))
        cash_out_vals = np.abs(cash_out_vals)
        net_cash_vals = cash_in_vals - cash_out_vals
        total_cash_in = float(np.sum(cash_in_vals))
        total_cash_out = float(np.sum(cash_out_vals))
        total_net_cash = float(np.sum(net_cash_vals))

        chart_left, chart_right = st.columns(2)

        with chart_left:
            st.markdown("<div class='section-title'>Cash In مقابل Cash Out وصافي التدفق</div>", unsafe_allow_html=True)
            fig_io = go.Figure()
            fig_io.add_trace(go.Bar(name='Cash In', x=time_cols, y=cash_in_vals, marker_color=RWAZ_GREEN,
                                    hovertemplate='%{x}<br>Cash In: SAR %{y:,.0f}<extra></extra>'))
            fig_io.add_trace(go.Bar(name='Cash Out', x=time_cols, y=cash_out_vals, marker_color=RWAZ_ACCENT,
                                    hovertemplate='%{x}<br>Cash Out: SAR %{y:,.0f}<extra></extra>'))
            net_labels = [f"({abs(v)/1e6:.1f}M)" if v < 0 else f"{v/1e6:.1f}M" for v in net_cash_vals]
            net_marker_colors = [RWAZ_RED if v < 0 else RWAZ_DARK for v in net_cash_vals]
            fig_io.add_trace(go.Scatter(
                name='Net Cash Flow', x=time_cols, y=net_cash_vals, mode='lines+markers+text',
                line=dict(color=RWAZ_DARK, width=2.2), marker=dict(size=5.5, color=net_marker_colors),
                text=net_labels, textposition='top center', textfont=dict(size=8.5, color=RWAZ_DARK),
                hovertemplate='%{x}<br>Net Cash Flow: SAR %{y:,.0f}<extra></extra>'
            ))
            fig_io.add_hline(y=0, line_color='#B9B1A9', line_width=1)
            apply_rwaz_plot_layout(fig_io, height=260, showlegend=True)
            fig_io.update_layout(
                barmode='group', legend=dict(orientation='h', y=1.11, x=.5, xanchor='center', font=dict(size=9)),
                xaxis=dict(title='', tickfont=dict(size=8.5), showgrid=False),
                yaxis=dict(title='SAR', tickformat='~s', gridcolor='#EEE9E3', zeroline=False),
                margin=dict(t=30,b=24,l=25,r=15), bargap=.28
            )
            st.plotly_chart(fig_io, use_container_width=True, config={'displayModeBar': False})
            st.markdown(
                f"<div style='display:grid;grid-template-columns:repeat(3,1fr);gap:5px;'>"
                f"<div class='mini-cell'><div style='font-size:8px;color:#7A7066;'>إجمالي Cash In</div><div class='ltr-num' style='font-size:9.5px;font-weight:900;color:#1F7A55;'>{fmt_currency_compact(total_cash_in)}</div></div>"
                f"<div class='mini-cell'><div style='font-size:8px;color:#7A7066;'>إجمالي Cash Out</div><div class='ltr-num' style='font-size:9.5px;font-weight:900;color:#8D622F;'>{fmt_currency_compact(total_cash_out)}</div></div>"
                f"<div class='mini-cell'><div style='font-size:8px;color:#7A7066;'>صافي التدفق</div><div class='ltr-num' style='font-size:9.5px;font-weight:900;{value_color_style(total_net_cash)}'>{fmt_currency_compact(total_net_cash)}</div></div>"
                f"</div>", unsafe_allow_html=True
            )

        with chart_right:
            st.markdown("<div class='section-title'>مسار السيولة وحاجز الأمان</div>", unsafe_allow_html=True)
            df_chart = pd.DataFrame({'التاريخ': time_cols, 'النقدية المتبقية': ending_cash_vals})
            point_colors = [RWAZ_RED if v < 0 else RWAZ_PRIMARY for v in ending_cash_vals]
            fig_cf = go.Figure()
            fig_cf.add_trace(go.Scatter(
                x=df_chart['التاريخ'], y=df_chart['النقدية المتبقية'], mode='lines+markers', fill='tozeroy',
                fillcolor='rgba(104,73,41,0.13)', line=dict(color=RWAZ_PRIMARY, width=2.4),
                marker=dict(size=6, color=point_colors, line=dict(color='#FFFFFF', width=.7)),
                hovertemplate='%{x}<br>SAR %{y:,.0f}<extra></extra>', name='الرصيد المتوقع'
            ))
            fig_cf.add_hline(y=500000, line_dash='dash', line_color=RWAZ_RED, line_width=1.5,
                             annotation_text='حد الأمان SAR 500,000', annotation_position='top right')
            if len(ending_cash_vals):
                fig_cf.add_annotation(x=min_cash_month, y=min_cash_val, text=fmt_currency_compact(min_cash_val),
                                      showarrow=True, arrowhead=0, ay=28, font=dict(size=9.5, color=RWAZ_RED if min_cash_val < 0 else RWAZ_DARK), bgcolor='rgba(255,255,255,.88)')
                fig_cf.add_annotation(x=max_cash_month, y=max_cash_val, text=fmt_currency_compact(max_cash_val),
                                      showarrow=True, arrowhead=0, ay=-28, font=dict(size=9.5, color=RWAZ_GREEN), bgcolor='rgba(255,255,255,.88)')
            apply_rwaz_plot_layout(fig_cf, height=260, showlegend=False)
            fig_cf.update_layout(
                xaxis=dict(title='', tickfont=dict(size=8.5), showgrid=False),
                yaxis=dict(title='SAR', tickformat='~s', gridcolor='#EEE9E3', zeroline=False),
                margin=dict(t=30,b=24,l=25,r=15)
            )
            st.plotly_chart(fig_cf, use_container_width=True, config={'displayModeBar': False})
            st.markdown(
                f"<div style='display:grid;grid-template-columns:repeat(3,1fr);gap:5px;'>"
                f"<div class='mini-cell'><div style='font-size:8px;color:#7A7066;'>الرصيد الحالي</div><div class='ltr-num' style='font-size:9.5px;font-weight:900;{value_color_style(current_cash)}'>{fmt_currency_compact(current_cash)}</div></div>"
                f"<div class='mini-cell'><div style='font-size:8px;color:#7A7066;'>أعلى رصيد</div><div class='ltr-num' style='font-size:9.5px;font-weight:900;color:#1F7A55;'>{fmt_currency_compact(max_cash_val)} · {max_cash_month}</div></div>"
                f"<div class='mini-cell'><div style='font-size:8px;color:#7A7066;'>أدنى رصيد</div><div class='ltr-num' style='font-size:9.5px;font-weight:900;{value_color_style(min_cash_val)}'>{fmt_currency_compact(min_cash_val)} · {min_cash_month}</div></div>"
                f"</div>", unsafe_allow_html=True
            )

    with tab_cf_detail:
        render_styled_dataframe(df_cf, max_height=650)

# ==============================================================================
# PAGE 4: RENTAL PROJECTS P&L
# ==============================================================================
elif page == "مشاريع الايجار":
    st.markdown("<div class='page-title'>قائمة الأرباح والخسائر للعقارات</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>أداء العقارات المكتملة، نسب الإشغال، صافي الدخل التشغيلي وصافي الربح.</div>", unsafe_allow_html=True)

    df_pl = store['df_pl'].copy()
    prop_metrics = extract_property_metrics(df_pl)
    proj_columns = prop_metrics['Project'].tolist() if not prop_metrics.empty else []
    selected_project = st.selectbox("اختر العقار / المشروع", ["جميع العقارات (All)"] + proj_columns)

    if selected_project != "جميع العقارات (All)" and not prop_metrics.empty:
        sel = prop_metrics[prop_metrics['Project'] == selected_project].iloc[0]
        t_units = float(sel['Units']) if pd.notna(sel['Units']) else 0.0
        o_units = float(sel['Occupied']) if pd.notna(sel['Occupied']) else 0.0
        p_occ = float(sel['Occupancy']) if pd.notna(sel['Occupancy']) else 0.0
        p_rev = float(sel['Revenue']) if pd.notna(sel['Revenue']) else 0.0
        p_noi = float(sel['NOI']) if pd.notna(sel['NOI']) else 0.0
        loss_count = 1 if pd.notna(sel['Margin']) and sel['Margin'] < 0 else 0
    else:
        t_units = pd.to_numeric(prop_metrics['Units'], errors='coerce').sum() if not prop_metrics.empty else 0.0
        o_units = pd.to_numeric(prop_metrics['Occupied'], errors='coerce').sum() if not prop_metrics.empty else 0.0
        p_occ = o_units / t_units if t_units > 0 else 0.0
        p_rev = pd.to_numeric(prop_metrics['Revenue'], errors='coerce').sum() if not prop_metrics.empty else 0.0
        p_noi = pd.to_numeric(prop_metrics['NOI'], errors='coerce').sum() if not prop_metrics.empty else 0.0
        loss_count = int((pd.to_numeric(prop_metrics['Margin'], errors='coerce') < 0).sum()) if not prop_metrics.empty else 0

    p1, p2, p3, p4, p5, p6 = st.columns(6)
    with p1: render_kpi("الوحدات", f"{t_units:.0f}", "إجمالي الوحدات", "positive")
    with p2: render_kpi("المؤجرة", f"{o_units:.0f}", f"الشاغر {max(0,t_units-o_units):.0f}", "positive")
    with p3: render_kpi("الإشغال", fmt_pct(p_occ), "المستهدف 97%", "positive" if p_occ >= .97 else "warning")
    with p4: render_kpi("الإيرادات", fmt_currency(p_rev), "صافي الإيرادات", "positive")
    with p5: render_kpi("صافي NOI", fmt_currency(p_noi), "الدخل التشغيلي", "positive" if p_noi >= 0 else "danger")
    with p6: render_kpi("مشاريع بخسارة", f"{loss_count}", "صافي هامش ربح سالب", "danger" if loss_count else "positive")

    tab_port, tab_perf, tab_pnl = st.tabs(["نظرة المحفظة", "تحليل الأداء", "P&L التفصيلي"])
    with tab_port:
        if not prop_metrics.empty:
            ranked = prop_metrics.dropna(subset=['Margin']).sort_values('Margin', ascending=False)
            if not ranked.empty:
                best_p = ranked.iloc[0]
                worst_p = ranked.iloc[-1]
                r1, r2 = st.columns(2)
                with r1: render_kpi("أعلى هامش صافي ربح", str(best_p['Project']), f"{fmt_pct(best_p['Margin'])}", "positive")
                with r2: render_kpi("أقل هامش صافي ربح", str(worst_p['Project']), f"{fmt_pct(worst_p['Margin'])}", "danger" if worst_p['Margin'] < 0 else "warning")
    with tab_perf:
        if not prop_metrics.empty:
            chart_pm = prop_metrics.copy()
            left, right = st.columns(2)

            with left:
                occ_df = chart_pm[['Project','Occupancy']].copy()
                occ_df['OccPct'] = pd.to_numeric(occ_df['Occupancy'], errors='coerce').fillna(0) * 100
                occ_df = occ_df.sort_values('OccPct', ascending=True).reset_index(drop=True)
                occ_colors = [RWAZ_RED if v < 80 else (RWAZ_AMBER if v < 97 else RWAZ_GREEN) for v in occ_df['OccPct']]
                occ_h = max(300, 72 + 27*len(occ_df))
                fig_occ = go.Figure()
                fig_occ.add_trace(go.Bar(
                    x=[100]*len(occ_df), y=occ_df['Project'], orientation='h',
                    marker_color='#ECE7E1', hoverinfo='skip', showlegend=False, width=.36
                ))
                fig_occ.add_trace(go.Bar(
                    x=occ_df['OccPct'], y=occ_df['Project'], orientation='h',
                    marker_color=occ_colors, text=[f"{v:.1f}%" for v in occ_df['OccPct']],
                    textposition='outside', cliponaxis=False, width=.20,
                    textfont=dict(size=12.5, color='#3F2D1E', family='Arial Black, Tahoma, Segoe UI'),
                    hovertemplate='%{y}<br>الإشغال %{x:.1f}%<extra></extra>', showlegend=False
                ))
                fig_occ.add_vline(x=97, line_dash='dash', line_color=RWAZ_PRIMARY, line_width=1.5,
                                  annotation_text='المستهدف 97%', annotation_position='top')
                apply_rwaz_plot_layout(fig_occ, height=occ_h)
                fig_occ.update_layout(
                    barmode='overlay',
                    xaxis=dict(range=[0,115], showgrid=False, title='', ticksuffix='%', tickfont=dict(size=10.5)),
                    yaxis=dict(title='', autorange='reversed', tickfont=dict(size=10.5)),
                    margin=dict(t=30,b=18,l=20,r=72), bargap=.48,
                    uniformtext_minsize=12, uniformtext_mode='show'
                )
                fig_occ.update_annotations(font=dict(size=11.5, color=RWAZ_PRIMARY, family='Tahoma, Segoe UI'))
                st.markdown("<div class='section-title'>نسبة الإشغال حسب العقار</div>", unsafe_allow_html=True)
                st.plotly_chart(fig_occ, use_container_width=True, config={'displayModeBar': False})

            with right:
                noi_df = chart_pm[['Project','NOI']].copy()
                noi_df['NOI_num'] = pd.to_numeric(noi_df['NOI'], errors='coerce').fillna(0)
                noi_df = noi_df.sort_values('NOI_num', ascending=False).reset_index(drop=True)
                noi_colors = [RWAZ_GREEN if v >= 0 else RWAZ_RED for v in noi_df['NOI_num']]
                noi_text = [fmt_currency_compact(v, decimals=1) for v in noi_df['NOI_num']]
                noi_h = max(300, 72 + 27*len(noi_df))
                fig_noi = go.Figure(go.Bar(
                    x=noi_df['NOI_num'], y=noi_df['Project'], orientation='h',
                    marker_color=noi_colors, text=noi_text, textposition='outside', cliponaxis=False,
                    hovertemplate='%{y}<br>SAR %{x:,.0f}<extra></extra>', width=.55
                ))
                fig_noi.add_vline(x=0, line_color='#B7AEA5', line_width=1)
                apply_rwaz_plot_layout(fig_noi, height=noi_h)
                max_abs_noi = max(float(np.abs(noi_df['NOI_num']).max()), 1.0)
                fig_noi.update_layout(
                    xaxis=dict(range=[-max_abs_noi*1.28, max_abs_noi*1.35], showgrid=False, zeroline=False, title='', showticklabels=False),
                    yaxis=dict(title='', autorange='reversed', tickfont=dict(size=9.5)),
                    margin=dict(t=18,b=18,l=20,r=72), bargap=.38
                )
                st.markdown("<div class='section-title'>NOI حسب العقار</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='direction:rtl;text-align:right;font-size:9.5px;color:#684929;font-weight:800;margin-bottom:2px;'>إجمالي NOI للمحفظة: <span class='ltr-num'>{fmt_currency_compact(pd.to_numeric(chart_pm['NOI'], errors='coerce').sum())}</span></div>", unsafe_allow_html=True)
                st.plotly_chart(fig_noi, use_container_width=True, config={'displayModeBar': False})

            margin_df = chart_pm[['Project','Margin']].copy()
            margin_df['MarginPct'] = pd.to_numeric(margin_df['Margin'], errors='coerce').fillna(0) * 100
            margin_df = margin_df.sort_values('MarginPct', ascending=True).reset_index(drop=True)
            margin_colors = [RWAZ_GREEN if v >= 0 else RWAZ_RED for v in margin_df['MarginPct']]
            fig_margin = go.Figure(go.Bar(
                x=margin_df['MarginPct'], y=margin_df['Project'], orientation='h',
                marker_color=margin_colors, text=[f"{v:.1f}%" for v in margin_df['MarginPct']],
                textposition='outside', cliponaxis=False, width=.55,
                hovertemplate='%{y}<br>صافي هامش الربح %{x:.1f}%<extra></extra>'
            ))
            fig_margin.add_vline(x=0, line_color='#8C837A', line_width=1.2)
            max_abs_margin = max(float(np.abs(margin_df['MarginPct']).max()), 1.0)
            apply_rwaz_plot_layout(fig_margin, height=max(310, 72 + 28*len(margin_df)))
            fig_margin.update_layout(
                xaxis=dict(range=[-max_abs_margin*1.28, max_abs_margin*1.28], title='', ticksuffix='%', gridcolor='#EEE9E3', zeroline=False),
                yaxis=dict(title='', autorange='reversed', tickfont=dict(size=9.5)),
                margin=dict(t=18,b=24,l=20,r=55), bargap=.38
            )
            st.markdown("<div class='section-title'>صافي هامش الربح حسب العقار</div>", unsafe_allow_html=True)
            st.plotly_chart(fig_margin, use_container_width=True, config={'displayModeBar': False})
    with tab_pnl:
        if len(df_pl) > 1:
            headers_row = df_pl.iloc[1].fillna("").values
            df_pl_formatted = df_pl.iloc[3:].copy()
            df_pl_formatted.columns = headers_row
            first_col_name = df_pl_formatted.columns[0]
            if first_col_name == "" or str(first_col_name).startswith("Unnamed"):
                df_pl_formatted.rename(columns={first_col_name: 'Category'}, inplace=True)
            render_styled_dataframe(df_pl_formatted, max_height=650, table_kind='pnl')
        else:
            render_styled_dataframe(df_pl, max_height=650, table_kind='pnl')

# ==============================================================================
# PAGE 5: DEVELOPMENT MODEL — financial engine preserved
# ==============================================================================
elif page == "موديل التطوير العقاري":
    st.markdown("<div class='page-title'>موديل دراسة جدوى التطوير العقاري (100% Equity)</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>شراء أرض، تطوير وبيع بالكامل بالتمويل الذاتي — معزول عن ديون الشركة.</div>", unsafe_allow_html=True)

    dev_defaults = {
        'dev_land_price': 12000000, 'dev_cost_sqm': 2200, 'dev_sellable_area': 8000, 'dev_selling_price_sqm': 6500,
        'dev_months_input': 14, 'dev_sales_months_input': 10, 'dev_ke_input': 14.0, 'dev_target_irr_input': 18.0
    }
    for k, v in dev_defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    with st.expander("افتراضات التطوير | Development Assumptions", expanded=True):
        if st.button("إعادة ضبط الافتراضات", key='reset_dev_assumptions'):
            for k, v in dev_defaults.items(): st.session_state[k] = v
        i1, i2, i3, i4 = st.columns(4)
        land_price = i1.number_input("سعر الأرض | Land Price", step=500000, key='dev_land_price')
        dev_cost_sqm = i2.number_input("تكلفة التطوير / م²", step=100, key='dev_cost_sqm')
        sellable_area = i3.number_input("المساحة البيعية | Sellable Area", step=500, key='dev_sellable_area')
        selling_price_sqm = i4.number_input("سعر البيع / م²", step=250, key='dev_selling_price_sqm')
        i5, i6, i7, i8 = st.columns(4)
        dev_months = i5.number_input("مدة التطوير | Months", step=1, key='dev_months_input')
        sales_months = i6.number_input("مدة البيع | Months", step=1, key='dev_sales_months_input')
        cost_of_equity = i7.number_input("تكلفة الملكية | Ke %", step=0.5, key='dev_ke_input') / 100.0
        target_equity_irr = i8.number_input("العائد المستهدف | IRR %", step=0.5, key='dev_target_irr_input') / 100.0

    res = run_dev_engine(land_price, 0.05, dev_cost_sqm, sellable_area, selling_price_sqm, dev_months, sales_months, cost_of_equity, target_equity_irr)
    tag_class = "status-pass" if res['decision'] == "PASS" else ("status-watch" if res['decision'] == "WATCH" else "status-fail")
    st.markdown(f"<div style='direction:rtl;text-align:right;font-size:11px;font-weight:800;margin:3px 0 5px;'>حالة القرار الاستثماري: <span class='{tag_class}'>{res['decision']}</span></div>", unsafe_allow_html=True)

    # Derived decision-support metrics use the existing engine as-is
    price_safety_margin = np.nan
    if selling_price_sqm > 0 and not np.isnan(res['req_price_npv_zero']):
        price_safety_margin = (selling_price_sqm - res['req_price_npv_zero']) / selling_price_sqm

    max_dev_cost_sqm = np.nan
    try:
        def _dev_cost_gap(c):
            rr = run_dev_engine(land_price, 0.05, c, sellable_area, selling_price_sqm, dev_months, sales_months, cost_of_equity, target_equity_irr)
            irr = rr['equity_irr']
            return (-1.0 if np.isnan(irr) else irr) - target_equity_irr
        lo_c, hi_c = max(1.0, dev_cost_sqm * 0.05), max(dev_cost_sqm * 5.0, dev_cost_sqm + 1000)
        if _dev_cost_gap(lo_c) * _dev_cost_gap(hi_c) <= 0:
            max_dev_cost_sqm = opt.brentq(_dev_cost_gap, lo_c, hi_c)
    except Exception:
        max_dev_cost_sqm = np.nan
    cost_overrun_sar = (max_dev_cost_sqm - dev_cost_sqm) * sellable_area if not np.isnan(max_dev_cost_sqm) else np.nan

    max_land_price = np.nan
    try:
        def _land_gap(lp):
            rr = run_dev_engine(lp, 0.05, dev_cost_sqm, sellable_area, selling_price_sqm, dev_months, sales_months, cost_of_equity, target_equity_irr)
            irr = rr['equity_irr']
            return (-1.0 if np.isnan(irr) else irr) - target_equity_irr
        lo_l, hi_l = 1.0, max(land_price * 5.0, land_price + 1000000)
        if _land_gap(lo_l) * _land_gap(hi_l) <= 0:
            max_land_price = opt.brentq(_land_gap, lo_l, hi_l)
    except Exception:
        max_land_price = np.nan

    delayed_res = run_dev_engine(land_price, 0.05, dev_cost_sqm, sellable_area, selling_price_sqm, dev_months + 6, sales_months, cost_of_equity, target_equity_irr)
    delay_npv_impact = delayed_res['equity_npv'] - res['equity_npv']
    required_monthly_sales = res['total_rev'] / sales_months if sales_months > 0 else np.nan

    tab_dev_res, tab_dev_be, tab_dev_sens = st.tabs(["نتائج الاستثمار", "نقاط وحدود القرار", "تحليل الحساسية"])
    with tab_dev_res:
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        with m1: render_kpi("إجمالي الإيرادات", fmt_currency(res['total_rev']), "إجمالي مبيعات المشروع", "positive")
        with m2: render_kpi("Equity IRR", fmt_pct(res['equity_irr']), "العائد الاستثماري", "positive" if res['equity_irr'] >= target_equity_irr else "danger")
        with m3: render_kpi("Equity NPV", fmt_currency(res['equity_npv']), f"Ke {fmt_pct(cost_of_equity)}", "positive" if res['equity_npv'] >= 0 else "danger")
        with m4: render_kpi("Equity MOIC", fmt_multiple(res['equity_moic']).replace('x','×'), "مضاعف الاستثمار", "positive")
        with m5: render_kpi("فترة الاسترداد", f"{res['payback_m']:.1f} شهر" if not np.isnan(res['payback_m']) else "N/A", "Payback", "positive")
        with m6: render_kpi("أعلى احتياج تمويلي", fmt_currency(res['peak_equity']), "Peak Equity", "warning")

        if res['decision'] == 'PASS':
            decision_text = f"العائد المتوقع {fmt_pct(res['equity_irr'])} أعلى من المستهدف {fmt_pct(target_equity_irr)} والقيمة الحالية موجبة {fmt_currency(res['equity_npv'])}."
        elif res['decision'] == 'WATCH':
            decision_text = "المشروع قريب من حدود القبول ويحتاج مراجعة الافتراضات الرئيسية قبل الاعتماد النهائي."
        else:
            decision_text = "المشروع لا يحقق شروط القبول الحالية؛ راجع سعر البيع أو تكلفة الأرض والتطوير أو مدة التنفيذ قبل الموافقة."
        render_decision_summary(res['decision'], decision_text)

    with tab_dev_be:
        st.markdown("<div class='section-title'>نقاط التعادل</div>", unsafe_allow_html=True)
        b1, b2, b3 = st.columns(3)
        with b1: render_kpi("التعادل المحاسبي", fmt_currency(res['accounting_be']), f"{fmt_currency(res['total_cost']/sellable_area)} /م²", "warning")
        with b2: render_kpi("إيراد NPV = 0", fmt_currency(res['npv_zero_rev']), f"{fmt_currency(res['req_price_npv_zero'])} /م²", "positive")
        with b3: render_kpi("إيراد Target IRR", fmt_currency(res['target_irr_rev']), f"{fmt_currency(res['req_price_target_irr'])} /م²", "positive")

        st.markdown("<div class='section-title'>حدود التحمل وصناعة القرار</div>", unsafe_allow_html=True)
        q1, q2, q3, q4, q5 = st.columns(5)
        with q1: render_kpi("هامش أمان السعر", fmt_pct(price_safety_margin) if not np.isnan(price_safety_margin) else "N/A", "قبل وصول NPV إلى صفر", "positive" if not np.isnan(price_safety_margin) and price_safety_margin > 0 else "danger")
        with q2: render_kpi("تحمل زيادة التكلفة", fmt_currency(cost_overrun_sar) if not np.isnan(cost_overrun_sar) else "N/A", "قبل فقد Target IRR", "positive" if not np.isnan(cost_overrun_sar) and cost_overrun_sar >= 0 else "danger")
        with q3: render_kpi("أقصى سعر أرض", fmt_currency(max_land_price) if not np.isnan(max_land_price) else "N/A", "مع الحفاظ على Target IRR", "positive")
        with q4: render_kpi("أثر تأخير 6 أشهر", fmt_currency(delay_npv_impact), "التغير في NPV", "danger" if delay_npv_impact < 0 else "positive")
        with q5: render_kpi("المبيعات الشهرية المطلوبة", fmt_currency(required_monthly_sales), "خلال فترة البيع", "warning")

    with tab_dev_sens:
        price_range = [selling_price_sqm * factor for factor in [0.85, 1.00, 1.15]]
        cost_range = [dev_cost_sqm * factor for factor in [0.85, 1.00, 1.15]]
        irr_matrix, npv_matrix, text_matrix = [], [], []
        for p_val in price_range:
            irr_row, npv_row, txt_row = [], [], []
            for c_val in cost_range:
                rr = run_dev_engine(land_price, 0.05, c_val, sellable_area, p_val, dev_months, sales_months, cost_of_equity, target_equity_irr)
                irr_pct = rr['equity_irr'] * 100 if not np.isnan(rr['equity_irr']) else np.nan
                irr_row.append(irr_pct); npv_row.append(rr['equity_npv'])
                txt_row.append(f"IRR {fmt_pct(rr['equity_irr'])}<br>NPV {fmt_currency_m(rr['equity_npv'])}")
            irr_matrix.append(irr_row); npv_matrix.append(npv_row); text_matrix.append(txt_row)
        x_labels = [f"تكلفة {c:,.0f}/م²" for c in cost_range]
        y_labels = [f"بيع {p:,.0f}/م²" for p in price_range]
        fig_sens = go.Figure(go.Heatmap(z=irr_matrix, x=x_labels, y=y_labels, colorscale=[[0,'#F6DAD7'], [.5,'#F2E6CF'], [1,'#DDEEE6']], colorbar=dict(title='IRR %'), hovertemplate='%{y}<br>%{x}<br>IRR %{z:.1f}%<extra></extra>'))
        for yi, yv in enumerate(y_labels):
            for xi, xv in enumerate(x_labels):
                fig_sens.add_annotation(x=xv, y=yv, text=text_matrix[yi][xi], showarrow=False, font=dict(size=10, color=RWAZ_DARK))
        apply_rwaz_plot_layout(fig_sens, height=330)
        fig_sens.update_layout(xaxis_title="", yaxis_title="")
        st.plotly_chart(fig_sens, use_container_width=True, config={'displayModeBar': False})

# ==============================================================================
# PAGE 6: SUB-LEASE MODEL — financial engine preserved
# ==============================================================================
elif page == "موديل الايجارات":
    st.markdown("<div class='page-title'>موديل إعادة التأجير Sub-Lease (100% Equity)</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>تقييم فرص الاستئجار وإعادة التأجير بالتمويل الذاتي مع فترة السماح وتصاعد إيجار المالك.</div>", unsafe_allow_html=True)

    rent_defaults = {
        'r_head_lease': 1200000, 'r_term': 10, 'r_escalation': 5.0, 'r_grace': 6, 'r_units': 40,
        'r_sub_rent': 45000, 'r_occ': 85.0, 'r_opex': 15.0, 'r_capex': 2000000, 'r_ke': 10.0, 'r_target': 15.0
    }
    for k, v in rent_defaults.items():
        if k not in st.session_state: st.session_state[k] = v

    with st.expander("افتراضات الإدارة | Rental Assumptions", expanded=True):
        if st.button("إعادة ضبط الافتراضات", key='reset_rent_assumptions'):
            for k, v in rent_defaults.items(): st.session_state[k] = v
        r1, r2, r3, r4 = st.columns(4)
        head_lease_rent = r1.number_input("إيجار المالك | Head Lease", step=100000, key='r_head_lease')
        lease_term_yrs = r2.number_input("مدة العقد | Years", step=1, key='r_term')
        rent_escalation = r3.number_input("زيادة إيجار المالك %", step=1.0, key='r_escalation') / 100.0
        grace_period_m = r4.number_input("فترة السماح | Months", step=1, key='r_grace')
        r5, r6, r7, r8 = st.columns(4)
        total_units = r5.number_input("إجمالي الوحدات", step=5, key='r_units')
        sub_rent_unit = r6.number_input("إيجار الوحدة | SAR", step=2500, key='r_sub_rent')
        target_occ = r7.number_input("الإشغال المستهدف %", step=5.0, key='r_occ') / 100.0
        opex_ratio = r8.number_input("التكاليف التشغيلية %", step=1.0, key='r_opex') / 100.0
        r9, r10, r11 = st.columns(3)
        fitout_capex = r9.number_input("التجهيز و CapEx | SAR", step=250000, key='r_capex')
        cost_of_equity = r10.number_input("تكلفة الملكية | Ke %", step=0.5, key='r_ke') / 100.0
        target_equity_irr = r11.number_input("العائد المستهدف | IRR %", step=0.5, key='r_target') / 100.0

    res_r = run_rental_engine(head_lease_rent, lease_term_yrs, rent_escalation, 3, grace_period_m, total_units, sub_rent_unit, target_occ, opex_ratio, fitout_capex, cost_of_equity, target_equity_irr)
    tag_class = "status-pass" if res_r['decision'] == "PASS" else ("status-watch" if res_r['decision'] == "WATCH" else "status-fail")
    st.markdown(f"<div style='direction:rtl;text-align:right;font-size:11px;font-weight:800;margin:3px 0 5px;'>حالة القرار الاستثماري: <span class='{tag_class}'>{res_r['decision']}</span></div>", unsafe_allow_html=True)

    occupancy_safety_margin = target_occ - res_r['be_occupancy']
    pnl_r = res_r['annual_pnl'].copy()
    first_loss_year = None
    if not pnl_r.empty and 'صافي الدخل NOI' in pnl_r.columns:
        neg_rows = pnl_r[pd.to_numeric(pnl_r['صافي الدخل NOI'], errors='coerce') < 0]
        if not neg_rows.empty: first_loss_year = neg_rows.iloc[0]['السنة']
    break_even_rent_unit = head_lease_rent / (total_units * target_occ * (1 - opex_ratio)) if total_units > 0 and target_occ > 0 and opex_ratio < 1 else np.nan

    tab_r_res, tab_r_be, tab_r_pnl = st.tabs(["نتائج الاستثمار", "التعادل وحدود القرار", "P&L لعمر العقد"])
    with tab_r_res:
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        with m1: render_kpi("إجمالي الإيرادات", fmt_currency(res_r['total_life_revenue']), f"خلال {lease_term_yrs} سنوات", "positive")
        with m2: render_kpi("Equity IRR", fmt_pct(res_r['equity_irr']), "العائد الاستثماري", "positive" if res_r['equity_irr'] >= target_equity_irr else "danger")
        with m3: render_kpi("Equity NPV", fmt_currency(res_r['equity_npv']), f"Ke {fmt_pct(cost_of_equity)}", "positive" if res_r['equity_npv'] >= 0 else "danger")
        with m4: render_kpi("Equity MOIC", fmt_multiple(res_r['equity_moic']).replace('x','×'), "مضاعف الاستثمار", "positive")
        with m5: render_kpi("فترة الاسترداد", f"{res_r['payback_yrs']:.1f} سنة" if not np.isnan(res_r['payback_yrs']) else "N/A", "Payback", "positive" if not np.isnan(res_r['payback_yrs']) else "danger")
        with m6: render_kpi("Fit-out CapEx", fmt_currency(res_r['fitout_capex']), "رأس المال المستثمر", "warning")

        if res_r['decision'] == 'PASS':
            r_text = f"المشروع يحقق العائد المستهدف بقيمة حالية {fmt_currency(res_r['equity_npv'])} وهامش إشغال فوق نقطة التعادل."
        elif res_r['decision'] == 'WATCH':
            r_text = "الصفقة قريبة من حدود القبول؛ راجع الإيجار الرئيسي والإشغال والتكاليف قبل الاعتماد."
        else:
            r_text = "الصفقة غير مناسبة بالشروط الحالية؛ يلزم تحسين إيجار الوحدات أو خفض إيجار المالك/التكاليف أو تعديل الافتراضات."
        render_decision_summary(res_r['decision'], r_text)

        if not pnl_r.empty:
            fig_noi_trend = go.Figure(go.Scatter(x=pnl_r['السنة'], y=pd.to_numeric(pnl_r['صافي الدخل NOI'], errors='coerce'), mode='lines+markers', line=dict(color=RWAZ_PRIMARY, width=2.5), marker=dict(size=6), hovertemplate='%{x}<br>SAR %{y:,.0f}<extra></extra>'))
            fig_noi_trend.add_hline(y=0, line_color=RWAZ_RED, line_dash='dash')
            apply_rwaz_plot_layout(fig_noi_trend, height=235)
            fig_noi_trend.update_layout(xaxis_title="", yaxis_title="NOI (SAR)")
            st.markdown("<div class='section-title'>مسار NOI عبر عمر العقد</div>", unsafe_allow_html=True)
            st.plotly_chart(fig_noi_trend, use_container_width=True, config={'displayModeBar': False})

    with tab_r_be:
        b1, b2, b3, b4 = st.columns(4)
        with b1: render_kpi("إشغال التعادل", fmt_pct(res_r['be_occupancy']), "NOI = 0", "warning")
        with b2: render_kpi("هامش أمان الإشغال", fmt_pct(occupancy_safety_margin), "المستهدف - التعادل", "positive" if occupancy_safety_margin > 0 else "danger")
        with b3: render_kpi("إشغال Target IRR", fmt_pct(res_r['occ_for_target_irr']) if not np.isnan(res_r['occ_for_target_irr']) else "N/A", "لتحقيق العائد المستهدف", "positive" if not np.isnan(res_r['occ_for_target_irr']) and res_r['occ_for_target_irr'] <= 1 else "danger")
        with b4: render_kpi("إيجار التعادل / وحدة", fmt_currency(break_even_rent_unit) if not np.isnan(break_even_rent_unit) else "N/A", "عند الإشغال المستهدف", "warning")
        if first_loss_year:
            render_compact_alert('warning', f"أول سنة يتحول فيها NOI إلى قيمة سالبة وفق الافتراضات الحالية: {first_loss_year}.")
        else:
            render_compact_alert('success', "لا يظهر NOI سلبي خلال مدة العقد وفق الافتراضات الحالية.")

    with tab_r_pnl:
        render_styled_dataframe(pnl_r, max_height=520)

# ==============================================================================
# PAGE 7: GUIDANCE — Arabic first / bidirectional-safe
# ==============================================================================
elif page == "ارشادات":
    st.markdown("<div class='page-title'>إرشادات المنصة ودليل المصطلحات المالية</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>شرح مبسط للإدارة لفهم أهم مؤشرات القرار الاستثماري المستخدمة في المنصة.</div>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>دليل المصطلحات والنسب المالية</div>", unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        st.markdown("""
        <div class="term-card"><div class="term-title">صافي القيمة الحالية | <span class="term-code">NPV</span></div><div class="term-en">Net Present Value</div><div class="term-desc">القيمة الحالية للتدفقات النقدية المستقبلية بعد خصمها بتكلفة الملكية. القيمة الموجبة تعني أن المشروع يضيف قيمة فوق تكلفة الفرصة البديلة.</div></div>
        <div class="term-card"><div class="term-title">معدل العائد الداخلي | <span class="term-code">IRR</span></div><div class="term-en">Internal Rate of Return</div><div class="term-desc">العائد السنوي المركب المتوقع من المشروع، ويُقارن بالعائد المستهدف لاتخاذ قرار القبول أو المراجعة أو الرفض.</div></div>
        <div class="term-card"><div class="term-title">مضاعف رأس المال | <span class="term-code">MOIC</span></div><div class="term-en">Multiple on Invested Capital</div><div class="term-desc">إجمالي التدفقات النقدية المحصلة مقارنة برأس المال المستثمر. مثال: <span class="term-code">1.72×</span> يعني استرداد رأس المال إضافة إلى 72% فوقه.</div></div>
        """, unsafe_allow_html=True)
    with g2:
        st.markdown("""
        <div class="term-card"><div class="term-title">فترة استرداد رأس المال | <span class="term-code">Payback</span></div><div class="term-en">Payback Period</div><div class="term-desc">المدة اللازمة حتى تسترد التدفقات النقدية المتراكمة أصل رأس المال المستثمر بالكامل.</div></div>
        <div class="term-card"><div class="term-title">نقاط تعادل الإيراد | <span class="term-code">Breakeven</span></div><div class="term-en">Revenue Breakeven Levels</div><div class="term-desc"><b>التعادل المحاسبي:</b> الإيراد الذي يجعل صافي الربح = 0.<br><b><span class="term-code">NPV = 0</span>:</b> الإيراد الذي يغطي تكلفة الملكية بالكامل.<br><b><span class="term-code">Target IRR</span>:</b> الإيراد اللازم لتحقيق العائد المستهدف.</div></div>
        <div class="term-card"><div class="term-title">نسبة إشغال التعادل | <span class="term-code">Break-even Occupancy</span></div><div class="term-en">Break-even Occupancy</div><div class="term-desc">أدنى إشغال مطلوب لتغطية إيجار المالك والتكاليف التشغيلية بحيث يكون صافي الدخل التشغيلي <span class="term-code">NOI = 0</span>.</div></div>
        """, unsafe_allow_html=True)
