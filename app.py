import streamlit as st
import pandas as pd
import numpy as np
import scipy.optimize as opt
import numpy_financial as npf
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
import os

# ==============================================================================
# GLOBAL BRANDING & MODERN EXECUTIVE THEME SYSTEM (RWAZ VIEW THEME)
# ==============================================================================
st.set_page_config(
    page_title="RWAZ VIEW — Executive Decision Support Platform",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium Executive Light Theme CSS System
st.markdown("""
<style>
    /* Hide Default Streamlit Elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Core Layout & Premium Light Gray Background */
    .stApp { background-color: #F1F5F9; color: #0F172A; font-family: 'Segoe UI', -apple-system, sans-serif; }
    .main .block-container { padding-top: 0.2rem !important; padding-bottom: 0.2rem !important; padding-left: 0.4rem !important; padding-right: 0.4rem !important; max-width: 100% !important; }
    
    /* Zero Unnecessary Space / Tight Spacing Rules */
    .stMarkdown, div[data-testid="stMarkdownContainer"] { margin-bottom: 0px !important; }
    div[data-testid="stVerticalBlock"] > div { gap: 0.3rem !important; }
    
    /* Typography Hierarchy */
    .page-title { font-size: 20px; font-weight: 800; color: #FFFFFF !important; letter-spacing: -0.3px; margin-bottom: 1px; text-align: right; background: #0F172A; padding: 6px 12px; border-radius: 6px; }
    .page-subtitle { font-size: 11px; color: #475569; font-weight: 600; margin-bottom: 6px; border-bottom: 1px solid #CBD5E1; padding-bottom: 4px; text-align: right; }
    .section-title { font-size: 13px; font-weight: 800; color: #0284C7; margin-top: 4px; margin-bottom: 4px; text-align: right; }
    
    /* Sidebar RWAZ VIEW Brand Header */
    .sidebar-brand { font-size: 22px; font-weight: 900; color: #38BDF8; letter-spacing: 0.5px; text-align: right; margin-bottom: 2px; display: flex; align-items: center; justify-content: flex-end; gap: 8px; }
    .sidebar-brand-icon { font-size: 26px; }
    
    /* Filter Controls Styling & Ultra High Contrast Dropdown Text */
    label { color: #0369A1 !important; font-weight: 800 !important; font-size: 12px !important; }
    div[data-baseweb="input"] { background-color: #FFFFFF !important; border: 1px solid #CBD5E1 !important; border-radius: 6px !important; }
    div[data-baseweb="input"] input { color: #0F172A !important; font-weight: 800 !important; }
    
    /* High Contrast Black Selectbox Text across all states */
    div[data-baseweb="select"] { background-color: #FFFFFF !important; border: 1px solid #CBD5E1 !important; border-radius: 6px !important; }
    div[data-baseweb="select"] span, div[data-baseweb="menu"] div, li[role="option"], div[role="combobox"] { color: #0F172A !important; font-weight: 800 !important; font-size: 12px !important; }
    
    /* Full Page Width DataFrame & Elegant Table Visibility */
    .stDataFrame, div[data-testid="stTable"] { background-color: #FFFFFF !important; border: 1px solid #CBD5E1 !important; border-radius: 6px !important; width: 100% !important; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    div[data-testid="stDataFrame"] > div { background-color: #FFFFFF !important; width: 100% !important; }
    
    /* Executive KPI Cards System */
    .kpi-container { background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 6px 10px; text-align: right; box-shadow: 0 2px 5px rgba(0,0,0,0.04); }
    .kpi-title { font-size: 10px; color: #64748B; font-weight: 700; text-transform: uppercase; }
    .kpi-value { font-size: 17px; color: #0F172A; font-weight: 800; margin-top: 2px; margin-bottom: 2px; font-variant-numeric: tabular-nums; }
    .kpi-sub { font-size: 10px; font-weight: 600; }
    .kpi-sub-positive { color: #16A34A; }
    .kpi-sub-warning { color: #D97706; }
    .kpi-sub-danger { color: #DC2626; }
    
    /* Status Tags */
    .status-pass { background-color: #DCFCE7; color: #15803D; border: 1px solid #16A34A; font-weight: 800; padding: 3px 10px; border-radius: 4px; display: inline-block; font-size: 12px; }
    .status-watch { background-color: #FEF9C3; color: #A16207; border: 1px solid #D97706; font-weight: 800; padding: 3px 10px; border-radius: 4px; display: inline-block; font-size: 12px; }
    .status-fail { background-color: #FEE2E2; color: #B91C1C; border: 1px solid #DC2626; font-weight: 800; padding: 3px 10px; border-radius: 4px; display: inline-block; font-size: 12px; }

    /* 1. SIDEBAR DARK BLUE THEME & WHITE TEXT */
    [data-testid="stSidebar"] { background-color: #0B132B !important; border-right: 1px solid #1C2541 !important; }
    [data-testid="stSidebar"] .stRadio > label { font-size: 13px !important; font-weight: 700 !important; color: #FFFFFF !important; text-align: right; }
    [data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span { color: #E2E8F0 !important; }
    
    /* Sidebar Radio Options Formatting */
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        padding: 8px 12px !important;
        border-radius: 6px !important;
        margin-bottom: 4px !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
        background-color: transparent !important;
    }

    /* Hover State for Sidebar Options */
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background-color: #1C2541 !important;
        cursor: pointer !important;
    }

    /* SELECTED / ACTIVE SIDEBAR PAGE HIGHLIGHT (WHITE BACKGROUND + DARK TEXT) */
    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked),
    [data-testid="stSidebar"] div[role="radiogroup"] label[aria-checked="true"],
    [data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15) !important;
    }

    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) span,
    [data-testid="stSidebar"] div[role="radiogroup"] label[aria-checked="true"] span,
    [data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] span {
        color: #0B132B !important;
        font-weight: 800 !important;
    }

    /* Combined Multi-Section Card Styles */
    .combined-card { background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 8px; padding: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.04); margin-bottom: 8px; }
    .combined-header { border-bottom: 1px solid #E2E8F0; padding-bottom: 6px; margin-bottom: 8px; text-align: right; }
    .combined-title { font-size: 11px; color: #64748B; font-weight: 800; text-transform: uppercase; }
    .combined-value { font-size: 20px; color: #0F172A; font-weight: 900; }
    .combined-sub { font-size: 10px; color: #0284C7; font-weight: 700; }
    
    /* Collection Card Progress Bar */
    .gauge-bar-bg { background-color: #E2E8F0; border-radius: 10px; height: 12px; width: 100%; overflow: hidden; margin-top: 8px; }
    .gauge-bar-fill { background: linear-gradient(90deg, #0284C7 0%, #38BDF8 100%); height: 100%; border-radius: 10px; }

    /* Expanded Term Card for Page 6 Guidance */
    .term-card { background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 14px 16px; margin-bottom: 8px; text-align: right; box-shadow: 0 2px 4px rgba(0,0,0,0.03); }
    .term-title { font-size: 15px; font-weight: 800; color: #0284C7; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px; }
    .term-desc { font-size: 12px; color: #334155; margin-top: 6px; line-height: 1.6; }
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

def render_kpi(title, value, sub_text, sub_type="positive"):
    sub_class = f"kpi-sub-{sub_type}"
    is_negative = str(value).startswith("(") or str(value).startswith("SAR (")
    val_color = "color: #DC2626 !important;" if is_negative else "color: #0F172A !important;"
    
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value" style="{val_color}">{value}</div>
        <div class="kpi-sub {sub_class}">{sub_text}</div>
    </div>
    """, unsafe_allow_html=True)

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

    master_f = find_file(['Master_Financial_Data_F.xlsx', 'Master_Financial_Data_F_2.xlsx', 'Master_Financial_Data_F_3.xlsx'])
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
# SIDEBAR NAVIGATION (RWAZ VIEW BRAND HEADER)
# ==============================================================================
st.sidebar.markdown("""
<div class="sidebar-brand">
    <span class="sidebar-brand-icon">🏢</span>
    <span>RWAZ VIEW</span>
</div>
""", unsafe_allow_html=True)
st.sidebar.caption("مركز القرار الاستثماري والمالي")

page = st.sidebar.radio("القائمة الرئيسية:", [
    "الملخص التنفيذي والمركز المالي",
    "السيولة والتدفقات النقدية",
    "مشاريع الايجار",
    "موديل التطوير العقاري",
    "موديل الايجارات",
    "ارشادات"
])

# ==============================================================================
# PAGE 1: الملخص التنفيذي والمركز المالي (MERGED EXECUTIVE DASHBOARD)
# ==============================================================================
if page == "الملخص التنفيذي والمركز المالي":
    st.markdown("<div class='page-title'>رواز | لوحة الإدارة التنفيذية والمركز المالي</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>نظرة شاملة رفيعة المستوى على السيولة، الإيرادات، الديون، محفظة الشركاء والمشاريع تحت الإنشاء.</div>", unsafe_allow_html=True)
    
    # Dynamic Math Calculations from Source Tables (Zero Hardcoding)
    df_b = store['df_banks']
    total_cash = pd.to_numeric(df_b['الرصيد'], errors='coerce').sum() if 'الرصيد' in df_b.columns else 0.0
    rajhi_cash = pd.to_numeric(df_b[df_b['البنك'].astype(str).str.contains('الراجحي', na=False)]['الرصيد'], errors='coerce').sum() if 'البنك' in df_b.columns else 0.0
    snb_cash = pd.to_numeric(df_b[df_b['البنك'].astype(str).str.contains('الأهلي', na=False)]['الرصيد'], errors='coerce').sum() if 'البنك' in df_b.columns else 0.0

    total_dev_val = pd.to_numeric(store['df_dev_projects']['إجمالي التكلفة'], errors='coerce').sum() if 'إجمالي التكلفة' in store['df_dev_projects'].columns else 0.0
    total_debt_orig = pd.to_numeric(store['df_loans']['أصل التمويل'], errors='coerce').sum() if 'أصل التمويل' in store['df_loans'].columns else 0.0
    total_debt_rem = pd.to_numeric(store['df_loans']['المتبقي للقرض'], errors='coerce').sum() if 'المتبقي للقرض' in store['df_loans'].columns else 0.0
    total_revenue = pd.to_numeric(store['df_revenues']['المبلغ'], errors='coerce').sum() if 'المبلغ' in store['df_revenues'].columns else 0.0
    
    coll_rate = pd.to_numeric(store['df_collections']['كفاءة التحصيل %'], errors='coerce').values[0] if 'كفاءة التحصيل %' in store['df_collections'].columns else 0.0
    due_coll = pd.to_numeric(store['df_collections']['المستحق للتحصيل'], errors='coerce').values[0] if 'المستحق للتحصيل' in store['df_collections'].columns else 0.0
    act_coll = pd.to_numeric(store['df_collections']['المحصل الفعلي'], errors='coerce').values[0] if 'المحصل الفعلي' in store['df_collections'].columns else 0.0
    
    partners_net = pd.to_numeric(store['df_partners']['الرصيد'], errors='coerce').sum() if 'الرصيد' in store['df_partners'].columns else 0.0

    # Upper Executive Row: Liquidity Card (Double Height Split) + Revenue + Debt
    col_top1, col_top2, col_top3 = st.columns([1.1, 0.9, 1.0])
    
    with col_top1:
        # Request 3: Double Height Cash Card (Upper = Total Bank Cash, Lower = 50/50 Al Rajhi & SNB)
        st.markdown(f"""
        <div class="combined-card">
            <div class="combined-header">
                <div class="combined-title">إجمالي رصيد البنوك</div>
                <div class="combined-value">{fmt_currency(total_cash)}</div>
                <div class="combined-sub">النقدية المتاحة بالحسابات</div>
            </div>
            <div style="display: flex; gap: 8px; justify-content: space-between;">
                <div style="flex: 1; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px; text-align: right;">
                    <div style="font-size: 10px; color: #64748B; font-weight: 700;">مصرف الراجحي</div>
                    <div style="font-size: 14px; font-weight: 800; color: #0F172A;">{fmt_currency(rajhi_cash)}</div>
                </div>
                <div style="flex: 1; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px; text-align: right;">
                    <div style="font-size: 10px; color: #64748B; font-weight: 700;">البنك الأهلي</div>
                    <div style="font-size: 14px; font-weight: 800; color: #0F172A;">{fmt_currency(snb_cash)}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with col_top2:
        # Request 4: Collection Performance Card matching image 92270e.png
        coll_pct_str = f"{coll_rate*100:.1f}%"
        st.markdown(f"""
        <div class="combined-card">
            <div style="font-size: 11px; color: #0F172A; font-weight: 800; text-align: right; margin-bottom: 6px;">أداء التحصيل</div>
            <div style="display: flex; align-items: center; justify-content: space-between; gap: 10px;">
                <div style="flex: 1; display: flex; flex-direction: column; gap: 6px;">
                    <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 4px 8px; text-align: center;">
                        <div style="font-size: 13px; font-weight: 800; color: #0F172A;">{fmt_currency_m(due_coll)}</div>
                        <div style="font-size: 9px; color: #64748B; font-weight: 700;">المستحق للتحصيل</div>
                    </div>
                    <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 4px 8px; text-align: center;">
                        <div style="font-size: 13px; font-weight: 800; color: #0F172A;">{fmt_currency_m(act_coll)}</div>
                        <div style="font-size: 9px; color: #64748B; font-weight: 700;">المحصل الفعلي</div>
                    </div>
                </div>
                <div style="flex: 1.1; text-align: center;">
                    <div style="font-size: 32px; font-weight: 900; color: #0F172A; line-height: 1;">{coll_pct_str}</div>
                    <div style="font-size: 10px; color: #64748B; font-weight: 700; margin-top: 4px;">كفاءة التحصيل</div>
                    <div class="gauge-bar-bg"><div class="gauge-bar-fill" style="width: {min(100, coll_rate*100):.1f}%;"></div></div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with col_top3:
        # Request 5: Merged Partner Accounts Card (Upper Net Balance, Lower Split 50/50 for each Partner)
        df_part = store['df_partners'].copy()
        p1_name, p1_bal, p2_name, p2_bal = "شريك 1", 0.0, "شريك 2", 0.0
        if not df_part.empty:
            if len(df_part) >= 1:
                p1_name = df_part.iloc[0]['الشريك'] if 'الشريك' in df_part.columns else "شريك 1"
                p1_bal = pd.to_numeric(df_part.iloc[0]['الرصيد'], errors='coerce') if 'الرصيد' in df_part.columns else 0.0
            if len(df_part) >= 2:
                p2_name = df_part.iloc[1]['الشريك'] if 'الشريك' in df_part.columns else "شريك 2"
                p2_bal = pd.to_numeric(df_part.iloc[1]['الرصيد'], errors='coerce') if 'الرصيد' in df_part.columns else 0.0

        st.markdown(f"""
        <div class="combined-card">
            <div class="combined-header">
                <div class="combined-title">صافي أرصدة الشركاء</div>
                <div class="combined-value">{fmt_currency(partners_net)}</div>
                <div class="combined-sub">إجمالي صافي حسابات الشركاء</div>
            </div>
            <div style="display: flex; gap: 8px; justify-content: space-between;">
                <div style="flex: 1; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px; text-align: right;">
                    <div style="font-size: 10px; color: #64748B; font-weight: 700;">{p1_name}</div>
                    <div style="font-size: 13px; font-weight: 800; color: #0F172A;">{fmt_currency(p1_bal)}</div>
                </div>
                <div style="flex: 1; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px; text-align: right;">
                    <div style="font-size: 10px; color: #64748B; font-weight: 700;">{p2_name}</div>
                    <div style="font-size: 13px; font-weight: 800; color: #0F172A;">{fmt_currency(p2_bal)}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Dynamic Management Alerts Executed at the Top/Middle Executive Area
    st.markdown("<div class='section-title'>🚨 التنبيهات والإجراءات الإدارية المباشرة</div>", unsafe_allow_html=True)
    
    df_inst = store['df_installments'].copy()
    overdue_found = False
    if 'الأيام المتبقية' in df_inst.columns and 'المتبقي للدفعة' in df_inst.columns:
        df_inst['days_rem'] = pd.to_numeric(df_inst['الأيام المتبقية'], errors='coerce')
        df_inst['bal_rem'] = pd.to_numeric(df_inst['المتبقي للدفعة'], errors='coerce')
        overdue_df = df_inst[(df_inst['days_rem'] < 0) & (df_inst['bal_rem'] > 0)]
        if not overdue_df.empty:
            overdue_found = True
            for _, o_row in overdue_df.iterrows():
                b_desc = o_row['بيان الدفعة'] if 'بيان الدفعة' in o_row else "دفعة مستحقة"
                b_val = o_row['المتبقي للدفعة']
                st.error(f"🔴 تنبيه تمويلي حرج: {b_desc} متأخر السداد (المتبقي المستحق {fmt_currency(b_val)}).")
    
    if not overdue_found:
        st.success("🟢 جميع أقساط التسهيلات والتمويلات مسددة في مواعيدها ومستقرة.")
        
    if coll_rate < 0.8:
        st.warning(f"🟡 كفاءة التحصيل الحالية ({fmt_pct(coll_rate)}) أقل من المستهدف الموصى به (80.0%).")

    # Section 5 — Re-organized Page 1 Layout:
    # Column Left: Debt KPIs next to Debt Table
    # Column Right: Revenue Mix Chart next to Projects Under Construction Table
    c_left, c_right = st.columns([1.1, 0.9])
    
    with c_left:
        st.markdown("<div class='section-title'>التمويلات والالتزامات الحالية</div>", unsafe_allow_html=True)
        # Debt KPIs placed right above/next to Debt Table
        dk1, dk2 = st.columns(2)
        with dk1: render_kpi("أصل التمويلات", fmt_currency(total_debt_orig), "إجمالي المسجل", "positive")
        with dk2: render_kpi("المتبقي للتمويلات", fmt_currency(total_debt_rem), "رصيد الدين الحالي", "warning" if total_debt_rem>0 else "positive")
        
        st.write("")
        df_loans_disp = store['df_loans'].copy()
        if 'أصل التمويل' in df_loans_disp.columns and 'المتبقي للقرض' in df_loans_disp.columns:
            rem_vals = pd.to_numeric(df_loans_disp['المتبقي للقرض'], errors='coerce').fillna(0)
            orig_vals = pd.to_numeric(df_loans_disp['أصل التمويل'], errors='coerce').fillna(1)
            df_loans_disp['نسبة السداد'] = ((orig_vals - rem_vals) / orig_vals).apply(lambda x: fmt_pct(x))
            df_loans_disp['الحالة'] = rem_vals.apply(lambda x: "تم السداد بالكامل" if x<=0 else "جاري السداد")
            
        st.dataframe(style_df_accounting(df_loans_disp), use_container_width=True)

    with c_right:
        st.markdown("<div class='section-title'>محفظة المشاريع تحت الإنشاء ومزيج الإيرادات</div>", unsafe_allow_html=True)
        # Revenue Mix Chart placed right next to Dev Projects
        df_rev_chart = store['df_revenues'].copy()
        if 'المبلغ' in df_rev_chart.columns and 'نوع الايراد' in df_rev_chart.columns:
            df_rev_chart['المبلغ_الرقدي'] = pd.to_numeric(df_rev_chart['المبلغ'], errors='coerce')
            df_rev_chart.sort_values(by='المبلغ_الرقدي', ascending=True, inplace=True)
            df_rev_chart['Formatted_SAR'] = df_rev_chart['المبلغ_الرقدي'].apply(lambda x: fmt_currency(x))
            
            fig_rev = px.bar(df_rev_chart, x='المبلغ_الرقدي', y='نوع الايراد', orientation='h', text='Formatted_SAR',
                             color='نوع الايراد', color_discrete_sequence=['#0284C7', '#6366F1', '#16A34A', '#D97706'])
            fig_rev.update_traces(textposition='outside', textfont=dict(color='#0F172A', size=11, family='Segoe UI'))
            
            max_val = df_rev_chart['المبلغ_الرقدي'].max() if not df_rev_chart.empty else 1000000
            fig_rev.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                                  font=dict(color='#0F172A', size=11),
                                  margin=dict(t=10, b=10, l=10, r=140), height=170, showlegend=False,
                                  xaxis=dict(range=[0, max_val * 1.45], title=""), yaxis_title="")
            st.plotly_chart(fig_rev, use_container_width=True)

        df_dev_disp = store['df_dev_projects'].copy()
        st.dataframe(style_df_accounting(df_dev_disp), use_container_width=True)

    # Request 6: Add Installments Table (جدول الاقساط) at the end of Page 1
    st.markdown("<div class='section-title'>جدول استحقاقات الأقساط (جدول الاقساط)</div>", unsafe_allow_html=True)
    df_inst_disp = store['df_installments'].copy()
    st.dataframe(style_df_accounting(df_inst_disp), use_container_width=True)

# ==============================================================================
# PAGE 2: السيولة والتدفقات النقدية
# ==============================================================================
elif page == "السيولة والتدفقات النقدية":
    st.markdown("<div class='page-title'>التدفقات النقدية والسيولة</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>مسار السيولة المتوقعة، أدنى وأعلى نقطة نقدية مع الشهر المرتبط، والتزامات الـ 90 يوماً القادمة.</div>", unsafe_allow_html=True)
    
    df_cf = store['df_cf']
    time_cols = store['time_cols']
    
    ending_row = df_cf[df_cf['Category'].astype(str).str.contains('end of period', case=False, na=False)]
    if ending_row.empty:
        ending_cash_vals = df_cf[time_cols].iloc[-1].values.flatten()
    else:
        ending_cash_vals = ending_row[time_cols].values.flatten()
        
    ending_cash_vals = np.nan_to_num(pd.to_numeric(ending_cash_vals, errors='coerce'), nan=0.0)
    
    min_idx = np.argmin(ending_cash_vals)
    max_idx = np.argmax(ending_cash_vals)
    
    min_cash_val = ending_cash_vals[min_idx]
    min_cash_month = time_cols[min_idx] if min_idx < len(time_cols) else ""
    
    max_cash_val = ending_cash_vals[max_idx]
    max_cash_month = time_cols[max_idx] if max_idx < len(time_cols) else ""
    
    outflow_row = df_cf[df_cf['Category'].astype(str).str.contains('out', case=False, na=False)]
    if not outflow_row.empty:
        outflow_90d = abs(pd.to_numeric(outflow_row[time_cols[:min(3, len(time_cols))]].values.flatten(), errors='coerce').sum())
    else:
        outflow_90d = 0.0
    
    m1, m2, m3, m4 = st.columns(4)
    with m1: render_kpi("السيولة الحالية", fmt_currency(store['df_banks']['الرصيد'].sum()), "النقدية المتاحة بالبنوك", "positive")
    with m2: render_kpi("أدنى نقطة سيولة", fmt_currency(min_cash_val), f"أدنى مستوى: {min_cash_month}", "warning" if min_cash_val<500000 else "positive")
    with m3: render_kpi("أعلى نقطة سيولة", fmt_currency(max_cash_val), f"أعلى مستوى: {max_cash_month}", "positive")
    with m4: render_kpi("التزامات الـ 90 يوماً القادمة", fmt_currency(outflow_90d), "إجمالي التدفقات الخارجة", "danger")

    st.markdown("<div class='section-title'>مسار السيولة النقدية وحاجز الأمان الأدنـى</div>", unsafe_allow_html=True)
    df_chart = pd.DataFrame({'التاريخ': time_cols, 'النقدية المتبقية': ending_cash_vals})
    
    fig_cf = px.area(df_chart, x='التاريخ', y='النقدية المتبقية', markers=True, color_discrete_sequence=['#0284C7'])
    fig_cf.add_hline(y=500000, line_dash="dash", line_color="#DC2626", annotation_text="حاجز أمان السيولة الأدنـى (SAR 500,000)")
    fig_cf.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#0F172A', size=11),
                         margin=dict(t=10, b=10, l=10, r=10), height=220, xaxis_title="الشهر", yaxis_title="النقدية (SAR)")
    st.plotly_chart(fig_cf, use_container_width=True)

    st.markdown("<div class='section-title'>جدول التدفقات النقدية الشامل</div>", unsafe_allow_html=True)
    st.dataframe(style_df_accounting(df_cf), use_container_width=True)

# ==============================================================================
# PAGE 3: مشاريع الايجار (P&L MATCHING ORIGINAL EXCEL SHEET3 STRUCTURE)
# ==============================================================================
elif page == "مشاريع الايجار":
    st.markdown("<div class='page-title'>قائمة الأرباح والخسائر للعقارات</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>أداء العقارات المكتملة، نسب الإشغال، وصافي الدخل التشغيلي NOI مطابق لهيكلة Excel المصدرية.</div>", unsafe_allow_html=True)
    
    df_pl = store['df_pl'].copy()
    
    # Extract Real Project Names dynamically from Row 1 Header
    raw_headers = df_pl.iloc[1].dropna().values.tolist()
    proj_columns = [str(p).strip() for p in raw_headers if str(p).strip() not in ['Category', 'TOTAL', 'Unnamed: 0', 'Unnamed: 1', 'Unnamed: 11', 'nan'] and not str(p).startswith('Unnamed')]
    
    # High Contrast Project Filter
    selected_project = st.selectbox("🎯 اختر العقار/المشروع لمتابعة المؤشرات الرئيسية:", ["جميع العقارات (All)"] + proj_columns)
    
    # Dynamic KPI Calculations depending on Selected Filter Choice
    units_row = df_pl[df_pl.iloc[:, 0].astype(str).str.contains('Total Units', case=False, na=False)]
    occ_units_row = df_pl[df_pl.iloc[:, 0].astype(str).str.contains('Occupied units', case=False, na=False)]
    rev_row = df_pl[df_pl.iloc[:, 0].astype(str).str.contains('Net Revenue', case=False, na=False)]
    noi_row = df_pl[df_pl.iloc[:, 0].astype(str).str.contains('Operating Income', case=False, na=False)]
    
    proj_col_idx = None
    if selected_project != "جميع العقارات (All)":
        for c in range(df_pl.shape[1]):
            if str(df_pl.iloc[1, c]).strip() == selected_project:
                proj_col_idx = c
                break
                
    if proj_col_idx is not None:
        t_units = pd.to_numeric(units_row.iloc[0, proj_col_idx], errors='coerce') if not units_row.empty else 0
        o_units = pd.to_numeric(occ_units_row.iloc[0, proj_col_idx], errors='coerce') if not occ_units_row.empty else 0
        p_occ = o_units / t_units if t_units > 0 else 0.0
        p_rev = pd.to_numeric(rev_row.iloc[0, proj_col_idx], errors='coerce') if not rev_row.empty else 0.0
        p_noi = pd.to_numeric(noi_row.iloc[0, proj_col_idx], errors='coerce') if not noi_row.empty else 0.0
    else:
        t_units = pd.to_numeric(units_row.iloc[0, 1:-1], errors='coerce').sum() if not units_row.empty else 263
        o_units = pd.to_numeric(occ_units_row.iloc[0, 1:-1], errors='coerce').sum() if not occ_units_row.empty else 241
        p_occ = o_units / t_units if t_units > 0 else 0.9164
        p_rev = pd.to_numeric(rev_row.iloc[0, 1:-1], errors='coerce').sum() if not rev_row.empty else 4848824
        p_noi = pd.to_numeric(noi_row.iloc[0, 1:-1], errors='coerce').sum() if not noi_row.empty else 1428887

    p1, p2, p3, p4, p5 = st.columns(5)
    with p1: render_kpi("الوحدات", f"{t_units:.0f} وحدة", f"المشروع: {selected_project}", "positive")
    with p2: render_kpi("المؤجرة", f"{o_units:.0f} وحدة", f"الشاغر: {t_units - o_units:.0f}", "positive")
    with p3: render_kpi("نسبة الإشغال", fmt_pct(p_occ), "المستهدف > 90%", "positive" if p_occ>=0.9 else "warning")
    with p4: render_kpi("الإيرادات", fmt_currency(p_rev), "إجمالي الإيرادات", "positive")
    with p5: render_kpi("صافي NOI", fmt_currency(p_noi), "الدخل التشغيلي", "positive" if p_noi>=0 else "danger")

    # Section: NOI Property Performance Ranking Cards (Dynamic)
    st.markdown("<div class='section-title'>🏆 ترتيب أداء العقارات (Rental Property Performance Ranking)</div>", unsafe_allow_html=True)
    if not noi_row.empty and len(proj_columns) > 0:
        noi_vals = []
        for p in proj_columns:
            for c in range(df_pl.shape[1]):
                if str(df_pl.iloc[1, c]).strip() == p:
                    val = pd.to_numeric(noi_row.iloc[0, c], errors='coerce')
                    noi_vals.append({'Project': p, 'NOI': val if not np.isnan(val) else 0.0})
                    break
        df_noi_rank = pd.DataFrame(noi_vals).sort_values(by='NOI', ascending=False)
        if not df_noi_rank.empty:
            best_p = df_noi_rank.iloc[0]
            worst_p = df_noi_rank.iloc[-1]
            rk1, rk2 = st.columns(2)
            with rk1: render_kpi("أعلى العقارات أداءً (Best NOI)", f"{best_p['Project']}", f"صافي NOI: {fmt_currency(best_p['NOI'])}", "positive")
            with rk2: render_kpi("أقل العقارات أداءً (Worst NOI)", f"{worst_p['Project']}", f"صافي NOI: {fmt_currency(worst_p['NOI'])}", "danger" if worst_p['NOI']<0 else "warning")

    st.markdown("<div class='section-title'>قائمة الأرباح والخسائر للعقارات (المصدر الرسمي - Sheet3)</div>", unsafe_allow_html=True)
    
    if len(df_pl) > 1:
        headers_row = df_pl.iloc[1].fillna("").values
        df_pl_formatted = df_pl.iloc[3:].copy()
        df_pl_formatted.columns = headers_row
        first_col_name = df_pl_formatted.columns[0]
        if first_col_name == "" or str(first_col_name).startswith("Unnamed"):
            df_pl_formatted.rename(columns={first_col_name: 'Category'}, inplace=True)
        st.dataframe(style_df_accounting(df_pl_formatted), use_container_width=True)
    else:
        st.dataframe(style_df_accounting(df_pl), use_container_width=True)

# ==============================================================================
# PAGE 4: موديل التطوير العقاري (100% EQUITY)
# ==============================================================================
elif page == "موديل التطوير العقاري":
    st.markdown("<div class='page-title'>موديل دراسة جدوى التطوير العقاري (100% Equity)</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>دراسة جدوى فرص التطوير بالكامل بالتمويل الذاتي — معزول تماماً عن ديون الشركة.</div>", unsafe_allow_html=True)

    with st.expander("🛠️ افتراضات التطوير (Development Assumptions)", expanded=True):
        i1, i2, i3, i4 = st.columns(4)
        land_price = i1.number_input("سعر الأرض Land Price (SAR)", value=12000000, step=500000)
        dev_cost_sqm = i2.number_input("تكلفة البناء Dev Cost / Sqm (SAR)", value=2200, step=100)
        sellable_area = i3.number_input("المساحة البيعية Sellable Area (Sqm)", value=8000, step=500)
        selling_price_sqm = i4.number_input("سعر البيع/متر Selling Price / Sqm (SAR)", value=6500, step=250)

        i5, i6, i7, i8 = st.columns(4)
        dev_months = i5.number_input("مدة البناء Dev Duration (Months)", value=14, step=1)
        sales_months = i6.number_input("مدة البيع Sales Horizon (Months)", value=10, step=1)
        cost_of_equity = i7.number_input("Cost of Equity Ke (%)", value=14.0, step=0.5) / 100.0
        target_equity_irr = i8.number_input("Target Equity IRR (%)", value=18.0, step=0.5) / 100.0

    res = run_dev_engine(land_price, 0.05, dev_cost_sqm, sellable_area, selling_price_sqm, dev_months, sales_months, cost_of_equity, target_equity_irr)

    st.markdown("<div class='section-title'>مؤشرات الاستثمار</div>", unsafe_allow_html=True)
    
    tag_class = "status-pass" if res['decision'] == "PASS" else ("status-watch" if res['decision'] == "WATCH" else "status-fail")
    st.markdown(f"**حالة القرار الاستثماري:** <span class='{tag_class}'>{res['decision']}</span>", unsafe_allow_html=True)
    st.write("")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    with m1: render_kpi("إجمالي الإيرادات", fmt_currency(res['total_rev']), "إجمالي مبيعات المشروع", "positive")
    with m2: render_kpi("Equity IRR", fmt_pct(res['equity_irr']), "العائد الاستثماري", "positive" if res['equity_irr']>=target_equity_irr else "danger")
    with m3: render_kpi("Equity NPV", fmt_currency(res['equity_npv']), f"خصم {fmt_pct(cost_of_equity)} Ke", "positive" if res['equity_npv']>=0 else "danger")
    with m4: render_kpi("Equity MOIC", fmt_multiple(res['equity_moic']), "مضاعف الاستثمار", "positive")
    with m5: render_kpi("Payback Period", f"{res['payback_m']:.1f} شهراً", "فترة الاسترداد", "positive")
    with m6: render_kpi("Peak Equity Req.", fmt_currency(res['peak_equity']), "أعلى احتياج سيولة", "warning")

    st.markdown("<div class='section-title'>نقاط التعادل</div>", unsafe_allow_html=True)
    s1, s2, s3 = st.columns(3)
    with s1: render_kpi("1. Accounting Breakeven", fmt_currency(res['accounting_be']), f"سعر المتر: {fmt_currency(res['total_cost']/sellable_area)}/Sqm", "warning")
    with s2: render_kpi("2. NPV = 0 Revenue (at Ke)", fmt_currency(res['npv_zero_rev']), f"سعر المتر: {fmt_currency(res['req_price_npv_zero'])}/Sqm", "positive")
    with s3: render_kpi("3. Target IRR Revenue (at Target)", fmt_currency(res['target_irr_rev']), f"سعر المتر: {fmt_currency(res['req_price_target_irr'])}/Sqm", "positive")

    st.markdown("<div class='section-title'>تحليل الحساسية (سعر البيع مقابل تكلفة التطوير)</div>", unsafe_allow_html=True)
    price_range = [selling_price_sqm * factor for factor in [0.85, 1.00, 1.15]]
    cost_range = [dev_cost_sqm * factor for factor in [0.85, 1.00, 1.15]]
    
    matrix_data = []
    for p in price_range:
        row = []
        for c in cost_range:
            r = run_dev_engine(land_price, 0.05, c, sellable_area, p, dev_months, sales_months, cost_of_equity, target_equity_irr)
            row.append(f"IRR: {fmt_pct(r['equity_irr'])} | NPV: {fmt_currency(r['equity_npv'])}")
        matrix_data.append(row)
        
    df_sens = pd.DataFrame(matrix_data, index=[f"سعر البيع: {fmt_currency(p)}/متر" for p in price_range], columns=[f"تكلفة التطوير: {fmt_currency(c)}/متر" for c in cost_range])
    st.dataframe(style_df_accounting(df_sens), use_container_width=True)

# ==============================================================================
# PAGE 5: موديل الايجارات (HEAD LEASE ESCALATION)
# ==============================================================================
elif page == "موديل الايجارات":
    st.markdown("<div class='page-title'>موديل إعادة التأجير Sub-Lease (100% Equity)</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>دراسة جدوى فرص الاستئجار وإعادة التأجير بالكامل بالتمويل الذاتي مع زيادة الإيجار وفترة السماح Grace Period.</div>", unsafe_allow_html=True)

    with st.expander("🛠️ افتراضات الإدارة", expanded=True):
        r1, r2, r3, r4 = st.columns(4)
        head_lease_rent = r1.number_input("إيجار المالك الرئيسي Head Lease (SAR)", value=1200000, step=100000)
        lease_term_yrs = r2.number_input("مدة العقد (سنوات)", value=10, step=1)
        rent_escalation = r3.number_input("نسبة زيادة إيجار المالك Head Lease Escalation (%)", value=5.0, step=1.0) / 100.0
        grace_period_m = r4.number_input("فترة السماح (أشهر)", value=6, step=1)

        r5, r6, r7, r8 = st.columns(4)
        total_units = r5.number_input("إجمالي عدد الوحدات", value=40, step=5)
        sub_rent_unit = r6.number_input("إيجار الوحدة / التجهيز (SAR)", value=45000, step=2500)
        target_occ = r7.number_input("نسبة الإشغال المستهدفة (%)", value=85.0, step=5.0) / 100.0
        opex_ratio = r8.number_input("نسبة التكاليف التشغيلية OPEX (%)", value=15.0, step=1.0) / 100.0

        r9, r10, r11 = st.columns(3)
        fitout_capex = r9.number_input("تكلفة التجهيز والـ CapEx (SAR)", value=2000000, step=250000)
        cost_of_equity = r10.number_input("Cost of Equity Ke (%)", value=10.0, step=0.5) / 100.0
        target_equity_irr = r11.number_input("Target Equity IRR (%)", value=15.0, step=0.5) / 100.0

    res_r = run_rental_engine(head_lease_rent, lease_term_yrs, rent_escalation, 3, grace_period_m, total_units, sub_rent_unit, target_occ, opex_ratio, fitout_capex, cost_of_equity, target_equity_irr)

    st.markdown("<div class='section-title'>مؤشرات الاستثمار</div>", unsafe_allow_html=True)
    
    tag_class = "status-pass" if res_r['decision'] == "PASS" else ("status-watch" if res_r['decision'] == "WATCH" else "status-fail")
    st.markdown(f"**حالة القرار الاستثماري:** <span class='{tag_class}'>{res_r['decision']}</span>", unsafe_allow_html=True)
    st.write("")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    with m1: render_kpi("إجمالي الإيرادات", fmt_currency(res_r['total_life_revenue']), f"خلال {lease_term_yrs} سنوات", "positive")
    with m2: render_kpi("Equity IRR", fmt_pct(res_r['equity_irr']), "العائد الاستثماري", "positive" if res_r['equity_irr']>=target_equity_irr else "danger")
    with m3: render_kpi("Equity NPV", fmt_currency(res_r['equity_npv']), f"خصم {fmt_pct(cost_of_equity)} Ke", "positive" if res_r['equity_npv']>=0 else "danger")
    with m4: render_kpi("Equity MOIC", fmt_multiple(res_r['equity_moic']), "مضاعف الاستثمار", "positive")
    with m5: render_kpi("Payback Period", f"{res_r['payback_yrs']:.1f} سنوات" if not np.isnan(res_r['payback_yrs']) else "N/A", "استرداد رأس المال", "positive")
    with m6: render_kpi("Fit-out CapEx Equity", fmt_currency(res_r['fitout_capex']), "رأس المال المستثمر", "warning")

    st.markdown("<div class='section-title'>تحليل التعادل</div>", unsafe_allow_html=True)
    b1, b2 = st.columns(2)
    with b1: render_kpi("Stabilized Break-even Occupancy %", fmt_pct(res_r['be_occupancy']), "لتحقيق صافي دخل NOI = 0", "warning")
    with b2: render_kpi("Occupancy for Target 15% IRR", fmt_pct(res_r['occ_for_target_irr']) if not np.isnan(res_r['occ_for_target_irr']) else "N/A", "لتحقيق العائد المستهدف", "positive")

    st.markdown("<div class='section-title'>القوائم المالية المتوقعة</div>", unsafe_allow_html=True)
    df_pnl_disp = res_r['annual_pnl'].copy()
    st.dataframe(style_df_accounting(df_pnl_disp), use_container_width=True)

# ==============================================================================
# PAGE 6: ارشادات ودليل المصطلحات المالي
# ==============================================================================
elif page == "ارشادات":
    st.markdown("<div class='page-title'>إرشادات المنصة ودليل المصطلحات المالية</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>دليل شرح مالي مبسط ومفصل لمساعدة الإدارة التنفيذية في فهم النسب والمصطلحات الاستثمارية.</div>", unsafe_allow_html=True)
    
    st.markdown("<div class='section-title'>📖 دليل التعريف بالمصطلحات والنسب المالية (Glossary)</div>", unsafe_allow_html=True)
    
    # Fully Expanded Cards for Maximum Screen Utilization & Comfort
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("""
        <div class="term-card">
            <div class="term-title">NPV (Net Present Value - صافي القيمة الحالية)</div>
            <div class="term-desc">القيمة الحالية للتدفقات النقدية المستقبلية للمشروع مخصومة بسعر خصم يمثل تكلفة الملكية (Cost of Equity Ke). النتيجة الإيجابية تعني أن المشروع يحقق عائداً أعلى من تكلفة الفرصة البديلة ويضيف قيمة حقيقية للشركة.</div>
        </div>
        <div class="term-card">
            <div class="term-title">IRR (Internal Rate of Return - معدل العائد الداخلي)</div>
            <div class="term-desc">نسبة العائد السنوي المركب المتوقع تحقيقه من أصل المشروع. يُقارن بمعدل العائد المستهدف (Target IRR) للشركة للبت في قبول المشروع PASS أو رفعه للرقابة WATCH أو رفضه FAIL.</div>
        </div>
        <div class="term-card">
            <div class="term-title">MOIC (Multiple on Invested Capital - مضاعف رأس المال)</div>
            <div class="term-desc">إجمالي التدفقات النقدية المحصلة مقسومة على إجمالي رأس المال الذاتي المستثمر. يُظهر كم مرة استرد المشروع النقدية المستثمرة (مثلاً 1.72x تعني استرداد رأس المال بالكامل + 72% أرباح صافية).</div>
        </div>
        """, unsafe_allow_html=True)
        
    with g2:
        st.markdown("""
        <div class="term-card">
            <div class="term-title">Payback Period (فترة استرداد رأس المال)</div>
            <div class="term-desc">المدة الزمنية بالأشهر أو السنوات المطلوبة حتى تسترد التدفقات النقدية المحصلة من المشروع أصل المبلغ المستثمر بالكامل.</div>
        </div>
        <div class="term-card">
            <div class="term-title">Breakeven Revenue Solvers (نقاط التعادل المستقلة)</div>
            <div class="term-desc">
            • <b>Accounting Breakeven:</b> الإيراد الذي يجعل صافي الربح المحاسبي = 0.<br>
            • <b>NPV = 0 Revenue:</b> الإيراد المطلوب لتغطية تكلفة الملكية Ke بالكامل دون أرباح إضافية.<br>
            • <b>Target IRR Revenue:</b> الإيراد المطلوب لتحقيق العائد المستهدف للإدارة بالكامل.
            </div>
        </div>
        <div class="term-card">
            <div class="term-title">Break-even Occupancy (نسبة إشغال التعادل)</div>
            <div class="term-desc">أدنى نسبة إشغال مطلوبة في مشاريع الإيجار لتغطية المصاريف التشغيلية وإيجار المالك الرئيسي بحيث يكون صافي الدخل التشغيلي NOI = 0.</div>
        </div>
        """, unsafe_allow_html=True)
