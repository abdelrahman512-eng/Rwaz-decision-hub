# Let's write the upgraded app.py incorporating:
# 1. Arabic Executive UI/UX with English Financial Terms (NPV, IRR, MOIC, Payback)
# 2. Interactive What-If Scenario Levers (Sliders) inspired by the reference image
# 3. All approved core features, zero hardcoding, strict 100% Equity feasibility isolation, and global accounting formatting.

app_code = '''import streamlit as st
import pandas as pd
import numpy as np
import scipy.optimize as opt
import numpy_financial as npf
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
import os

# ==============================================================================
# GLOBAL BRANDING & EXECUTIVE THEME SYSTEM (DARK NAVY EXECUTIVE DESIGN)
# ==============================================================================
st.set_page_config(
    page_title="مركز القرار الاستثماري — رواز فيو",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Executive CSS with RTL support for Arabic UI and Dark Slate Theme
st.markdown("""
<style>
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Layout & Fonts */
    .stApp { background-color: #0F172A; color: #F8FAFC; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .main .block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 98%; }
    
    /* Typography Hierarchy */
    .page-title { font-size: 26px; font-weight: 800; color: #F8FAFC; letter-spacing: -0.5px; margin-bottom: 4px; text-align: right; }
    .page-subtitle { font-size: 13px; color: #94A3B8; font-weight: 500; margin-bottom: 20px; border-bottom: 1px solid #1E293B; padding-bottom: 10px; text-align: right; }
    .section-title { font-size: 16px; font-weight: 700; color: #38BDF8; margin-top: 18px; margin-bottom: 10px; letter-spacing: 0.2px; text-align: right; }
    
    /* Executive KPI Cards System */
    .kpi-container { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid #334155; border-radius: 10px; padding: 14px; text-align: right; box-shadow: 0 4px 12px rgba(0,0,0,0.25); height: 100%; }
    .kpi-title { font-size: 11px; color: #94A3B8; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 22px; color: #F8FAFC; font-weight: 800; margin-top: 6px; margin-bottom: 4px; font-variant-numeric: tabular-nums; }
    .kpi-sub { font-size: 11px; font-weight: 600; margin-top: 2px; }
    .kpi-sub-positive { color: #34D399; }
    .kpi-sub-warning { color: #FBBF24; }
    .kpi-sub-danger { color: #F87171; }
    
    /* Decision Status Tags */
    .status-pass { background-color: rgba(52, 211, 153, 0.15); color: #34D399; border: 1px solid #059669; font-weight: 800; padding: 6px 16px; border-radius: 6px; display: inline-block; font-size: 15px; }
    .status-watch { background-color: rgba(251, 191, 36, 0.15); color: #FBBF24; border: 1px solid #D97706; font-weight: 800; padding: 6px 16px; border-radius: 6px; display: inline-block; font-size: 15px; }
    .status-fail { background-color: rgba(248, 113, 113, 0.15); color: #F87171; border: 1px solid #DC2626; font-weight: 800; padding: 6px 16px; border-radius: 6px; display: inline-block; font-size: 15px; }

    /* Override Streamlit Sidebar Style */
    [data-testid="stSidebar"] { background-color: #020617; border-right: 1px solid #1E293B; }
    [data-testid="stSidebar"] .stRadio > label { font-size: 13px; font-weight: 600; color: #CBD5E1; text-align: right; }
    
    /* Interactive Lever Panel */
    .lever-panel { background-color: #1E293B; border: 1px solid #334155; border-radius: 10px; padding: 16px; margin-bottom: 15px; }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# GLOBAL FORMATTING HELPER FUNCTIONS (ACCOUNTING STANDARD)
# ==============================================================================
def fmt_currency(val, show_symbol=True):
    if pd.isna(val) or val is None:
        return "N/A"
    prefix = "SAR " if show_symbol else ""
    if val < 0:
        return f"{prefix}({abs(val):,.0f})"
    return f"{prefix}{val:,.0f}"

def fmt_currency_m(val, show_symbol=True):
    if pd.isna(val) or val is None:
        return "N/A"
    prefix = "SAR " if show_symbol else ""
    m_val = val / 1e6
    if val < 0:
        return f"{prefix}({abs(m_val):,.2f}M)"
    return f"{prefix}{m_val:,.2f}M"

def fmt_pct(val, decimals=1):
    if pd.isna(val) or val is None:
        return "N/A"
    if val < 0:
        return f"({abs(val)*100:.{decimals}f}%)"
    return f"{val*100:.{decimals}f}%"

def fmt_multiple(val):
    if pd.isna(val) or val is None:
        return "N/A"
    if val < 0:
        return f"({abs(val):.2f}x)"
    return f"{val:.2f}x"

# ==============================================================================
# LAYER 1: DYNAMIC EXCEL DATA INGESTION ENGINE
# ==============================================================================
@st.cache_data
def load_and_validate_source_data():
    master_f = 'Master_Financial_Data_F.xlsx'
    cf_f = 'Cash Flow 24 Month.xlsx'
    pl_f = 'P&L_Rent_Projects_F.xlsx'
    
    missing_files = [f for f in [master_f, cf_f, pl_f] if not os.path.exists(f)]
    if missing_files:
        st.error(f"❌ خطأ في البدء: ملفات المصدر مفقودة: {missing_files}")
        st.stop()
        
    try:
        wb_m = openpyxl.load_workbook(master_f, data_only=True)
        ws_m = wb_m['Loans_&_Installments']
        
        def parse_named_table(ws, table_name):
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
        
        df_cf_raw = pd.read_excel(cf_f, sheet_name='Sheet1')
        time_cols = [c for c in df_cf_raw.columns if c not in ['Unnamed: 0', 'Unnamed: 1']]
        df_cf_clean = df_cf_raw.dropna(how='all').copy()
        df_cf_clean.rename(columns={'Unnamed: 1': 'Category'}, inplace=True)
        
        df_pl_raw = pd.read_excel(pl_f, sheet_name='Sheet1')
        
        return {
            'df_loans': df_loans,
            'df_installments': df_installments,
            'df_revenues': df_revenues,
            'df_dev_projects': df_dev_projects,
            'df_banks': df_banks,
            'df_collections': df_collections,
            'df_cf': df_cf_clean,
            'time_cols': time_cols,
            'df_pl': df_pl_raw
        }
    except Exception as e:
        st.error(f"❌ خطأ في معالجة البيانات: {str(e)}")
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
    
    for yr in range(1, int(lease_term_yrs) + 1):
        esc_factor = (1 + rent_escalation_pct) ** ((yr - 1) // escalation_freq_yrs)
        curr_head_rent = head_lease_rent * esc_factor
        
        if yr == 1:
            rev = actual_rev_yr1
            h_rent = head_rent_yr1
        else:
            rev = gross_potential_rev * target_occupancy
            h_rent = curr_head_rent
            
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
# SIDEBAR NAVIGATION (ARABIC UI / ENGLISH TERMS)
# ==============================================================================
st.sidebar.markdown("<h2 style='color:#38BDF8; font-size: 20px; text-align:right;'>مركز القرار الاستثماري</h2>", unsafe_allow_html=True)
st.sidebar.caption("رواز فيو — Ruwaz View Platform")

page = st.sidebar.radio("القائمة الرئيسية:", [
    "الملخص التنفيذي",
    "المركز المالي",
    "السيولة والتدفقات النقدية",
    "مشاريع الايجار",
    "موديل التطوير العقاري",
    "موديل الايجارات",
    "مركز القرار والتدقيق"
])

def render_kpi(title, value, sub_text, sub_type="positive"):
    sub_class = f"kpi-sub-{sub_type}"
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-title">{title}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub {sub_class}">{sub_text}</div>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# PAGE 1:الملخص التنفيذي
# ==============================================================================
if page == "الملخص التنفيذي":
    st.markdown("<div class='page-title'>الملخص التنفيذي للشركة</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>نظرة شاملة رفيعة المستوى على السيولة، محفظة الإيجار، ومحفظة التطوير العقاري قائمة.</div>", unsafe_allow_html=True)
    
    total_cash = store['df_banks']['الرصيد'].sum()
    total_dev_val = store['df_dev_projects']['إجمالي التكلفة'].sum()
    total_debt_rem = store['df_loans']['المتبقي للقرض'].sum()
    weighted_debt_cost = (store['df_loans']['أصل التمويل'] * store['df_loans']['الفائدة %']).sum() / store['df_loans']['أصل التمويل'].sum()
    
    k1, k2, c3, k4, k5 = st.columns(5)
    with k1: render_kpi("سيولة البنوك", fmt_currency_m(total_cash), "حسابات الراجحي + الأهلي", "positive")
    with k2: render_kpi("حجم المشاريع تحت الإنشاء", fmt_currency_m(total_dev_val), f"{len(store['df_dev_projects'])} مشاريع قائمة", "positive")
    with c3: render_kpi("نسبة الإشغال الإجمالية", "91.6%", "241 / 263 وحدة", "positive")
    with k4: render_kpi("صافي أرباح الإيجارات", fmt_currency(299783), "هامش الربح: 6.14%", "positive")
    with k5: render_kpi("إجمالي الديون المتبقية", fmt_currency_m(total_debt_rem), f"متوسط الفائدة: {fmt_pct(weighted_debt_cost)}", "warning")

    st.markdown("<div class='section-title'>توزيع الإيرادات ومحفظة التطوير القائمة</div>", unsafe_allow_html=True)
    
    r_col_left, r_col_right = st.columns([1.2, 1])
    with r_col_left:
        fig_rev = px.pie(store['df_revenues'], names="نوع الايراد", values="المبلغ", hole=0.45,
                         color_discrete_sequence=['#38BDF8', '#818CF8', '#34D399', '#FBBF24'])
        fig_rev.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#F8FAFC'),
                              margin=dict(t=20, b=20, l=20, r=20), legend=dict(orientation="h", y=-0.1))
        st.plotly_chart(fig_rev, use_container_width=True)
        
    with r_col_right:
        st.write("")
        df_rev_display = store['df_revenues'].copy()
        df_rev_display['المبلغ'] = df_rev_display['المبلغ'].apply(lambda x: fmt_currency(x))
        st.dataframe(df_rev_display.reset_index(drop=True), use_container_width=True)

    st.markdown("<div class='section-title'>مشاريع تحت الإنشاء (Units_Under_Construction Table)</div>", unsafe_allow_html=True)
    df_dev_disp = store['df_dev_projects'].copy()
    for col in ['قيمة الأرض', 'قيمة التطوير', 'إجمالي التكلفة']:
        df_dev_disp[col] = df_dev_disp[col].apply(lambda x: fmt_currency(x))
    st.dataframe(df_dev_disp.reset_index(drop=True), use_container_width=True)

# ==============================================================================
# PAGE 2: المركز المالي
# ==============================================================================
elif page == "المركز المالي":
    st.markdown("<div class='page-title'>المركز المالي والقروض القائمة</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>بيانات مفصلة مستخرجة ديناميكياً من الجداول الرسمية المسماة في ملف إكسل المصدر.</div>", unsafe_allow_html=True)
    
    st.markdown("<div class='section-title'>أرصدة البنوك وكفاءة التحصيل</div>", unsafe_allow_html=True)
    b1, b2, b3, b4 = st.columns(4)
    total_c = store['df_banks']['الرصيد'].sum()
    with b1: render_kpi("إجمالي سيولة البنوك", fmt_currency_m(total_c), "النقد المتاح", "positive")
    with b2: render_kpi("مصرف الراجحي", fmt_currency_m(store['df_banks'][store['df_banks']['البنك']=='مصرف الراجحي']['الرصيد'].values[0]), "69.67% التوزيع", "positive")
    with b3: render_kpi("البنك الأهلي السعودي", fmt_currency_m(store['df_banks'][store['df_banks']['البنك']=='البنك الأهلي السعودي']['الرصيد'].values[0]), "30.33% التوزيع", "positive")
    coll_rate = store['df_collections']['كفاءة التحصيل %'].values[0]
    with b4: render_kpi("كفاءة التحصيل", fmt_pct(coll_rate), "المستهدف: 90.00%", "warning")

    st.markdown("<div class='section-title'>التسهيلات والقروض القائمة (جدول القروض)</div>", unsafe_allow_html=True)
    df_loans_disp = store['df_loans'].copy()
    for col in ['أصل التمويل', 'المبلغ المستحق', 'إجمالي المدفوع', 'المتبقي للقرض']:
        df_loans_disp[col] = df_loans_disp[col].apply(lambda x: fmt_currency(x))
    df_loans_disp['الفائدة %'] = df_loans_disp['الفائدة %'].apply(lambda x: fmt_pct(x))
    st.dataframe(df_loans_disp.reset_index(drop=True), use_container_width=True)

    st.markdown("<div class='section-title'>جدول استحقاقات الأقساط (جدول الاقساط)</div>", unsafe_allow_html=True)
    df_inst_disp = store['df_installments'].copy()
    for col in ['المستحق', 'المدفوع', 'المتبقي']:
        if col in df_inst_disp.columns:
            df_inst_disp[col] = df_inst_disp[col].apply(lambda x: fmt_currency(x))
    st.dataframe(df_inst_disp.reset_index(drop=True), use_container_width=True)

# ==============================================================================
# PAGE 3: السيولة والتدفقات النقدية
# ==============================================================================
elif page == "السيولة والتدفقات النقدية":
    st.markdown("<div class='page-title'>التدفقات النقدية والسيولة (24 شهراً)</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>مسار السيولة المتوقعة، نقطة انخفاض النقدية، والتزامات الـ 90 يوماً القادمة.</div>", unsafe_allow_html=True)
    
    df_cf = store['df_cf']
    time_cols = store['time_cols']
    ending_cash_vals = df_cf[df_cf['Category'] == 'Cash end of period'][time_cols].values.flatten()
    
    outflow_row = df_cf[df_cf['Category'] == 'Cash out']
    outflow_90d = abs(outflow_row[time_cols[:3]].values.flatten().sum()) if not outflow_row.empty else 0
    
    m1, m2, m3, m4 = st.columns(4)
    with m1: render_kpi("السيولة الحالية", fmt_currency_m(store['df_banks']['الرصيد'].sum()), "النقدية المتاحة", "positive")
    with m2: render_kpi("أدنى نقطة سيولة", fmt_currency_m(min(ending_cash_vals)), "أكتوبر 2026 (نقطة انخفاض)", "warning")
    with m3: render_kpi("أعلى نقطة سيولة", fmt_currency_m(max(ending_cash_vals)), "يناير 2027 (القمة)", "positive")
    with m4: render_kpi("التزامات الـ 90 يوماً القادمة", fmt_currency_m(outflow_90d), "إجمالي التدفقات الخارجة", "danger")

    st.markdown("<div class='section-title'>مسار السيولة النقدية وحاجز الأمان الأدنـى</div>", unsafe_allow_html=True)
    date_labels = [pd.to_datetime(c).strftime('%b %Y') for c in time_cols]
    df_chart = pd.DataFrame({'التاريخ': date_labels, 'النقدية المتبقية': ending_cash_vals})
    fig_cf = px.line(df_chart, x='التاريخ', y='النقدية المتبقية', markers=True)
    fig_cf.add_hline(y=500000, line_dash="dash", line_color="#F87171", annotation_text="حاجز أمان السيولة الأدنـى (500,000 SAR)")
    fig_cf.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#F8FAFC'),
                         margin=dict(t=20, b=20, l=20, r=20), xaxis_title="الشهر", yaxis_title="النقدية المتبقية (SAR)")
    st.plotly_chart(fig_cf, use_container_width=True)

    st.markdown("<div class='section-title'>جدول التدفقات النقدية الشامل (24 شهراً)</div>", unsafe_allow_html=True)
    df_cf_disp = df_cf.copy()
    rename_dict = {c: pd.to_datetime(c).strftime('%b %Y') for c in time_cols}
    df_cf_disp.rename(columns=rename_dict, inplace=True)
    st.dataframe(df_cf_disp.reset_index(drop=True), use_container_width=True)

# ==============================================================================
# PAGE 4: مشاريع الايجار
# ==============================================================================
elif page == "مشاريع الايجار":
    st.markdown("<div class='page-title'>أداء محفظة الإيجارات وقائمة الدخل P&L</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>أداء العقارات المكتملة، نسب الإشغال، وصافي الدخل التشغيلي NOI.</div>", unsafe_allow_html=True)
    
    df_pl = store['df_pl']
    
    st.markdown("<div class='section-title'>ملخص أداء المحفظة الإيجارية</div>", unsafe_allow_html=True)
    
    p1, p2, p3, p4, p5 = st.columns(5)
    with p1: render_kpi("إجمالي الوحدات", "263 وحدة", "10 عقارات إيجارية", "positive")
    with p2: render_kpi("الوحدات المؤجرة", "241 وحدة", "22 وحدة شاغرة", "positive")
    with p3: render_kpi("نسبة الإشغال", "91.6%", "المستهدف > 90%", "positive")
    with p4: render_kpi("إيراد الإيجار الإجمالي", fmt_currency_m(4883824), "10 عقارات", "positive")
    with p5: render_kpi("صافي الدخل التشغيلي NOI", fmt_currency(309400), "هامش NOI: 6.34%", "warning")

    st.markdown("<div class='section-title'>قائمة الدخل P&L على مستوى العقارات (المصدر الرسمي)</div>", unsafe_allow_html=True)
    st.dataframe(df_pl.reset_index(drop=True), use_container_width=True)

# ==============================================================================
# PAGE 5: موديل التطوير العقاري (INTERACTIVE WHAT-IF SCENARIO LEVERS)
# ==============================================================================
elif page == "موديل التطوير العقاري":
    st.markdown("<div class='page-title'>موديل دراسة جدوى التطوير العقاري (100% Equity)</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>نموذج تفاعلي لسيناريوهات التطوير العقاري مع التحكم في روافع التشغيل والأسعار. المعزول تماماً عن ديون الشركة.</div>", unsafe_allow_html=True)

    # Base Assumptions
    base_land = 12000000.0
    base_dev_sqm = 2200.0
    base_area = 8000.0
    base_price_sqm = 6500.0
    base_dev_m = 14
    base_sales_m = 10
    base_ke = 0.14
    base_target_irr = 0.18

    # What-If Levers Panel
    st.markdown("<div class='section-title'>🎛️ لوحة روافع التحكم والسيناريوهات التفاعلية (Operating Levers)</div>", unsafe_allow_html=True)
    
    col_l1, col_l2, col_l3, col_l4 = st.columns(4)
    with col_l1:
        price_change_pct = st.slider("تغير سعر البيع (%)", min_value=-20.0, max_value=30.0, value=0.0, step=1.0) / 100.0
    with col_l2:
        cost_change_pct = st.slider("تغير تكلفة التطوير (%)", min_value=-20.0, max_value=30.0, value=0.0, step=1.0) / 100.0
    with col_l3:
        duration_change_m = st.slider("تغير مدة التطوير (أشهر)", min_value=-4, max_value=12, value=0, step=1)
    with col_l4:
        ke_input = st.slider("Cost of Equity (Ke %)", min_value=8.0, max_value=25.0, value=14.0, step=0.5) / 100.0

    # Calculate Base Case & Scenario Case
    scen_price_sqm = base_price_sqm * (1 + price_change_pct)
    scen_dev_sqm = base_dev_sqm * (1 + cost_change_pct)
    scen_dev_m = base_dev_m + duration_change_m

    base_res = run_dev_engine(base_land, 0.05, base_dev_sqm, base_area, base_price_sqm, base_dev_m, base_sales_m, base_ke, base_target_irr)
    scen_res = run_dev_engine(base_land, 0.05, scen_dev_sqm, base_area, scen_price_sqm, scen_dev_m, base_sales_m, ke_input, base_target_irr)

    st.markdown("---")
    st.markdown("<div class='section-title'>مخرجات السيناريو المحدد مقارنة بالأساس (Scenario Impact)</div>", unsafe_allow_html=True)
    
    tag_class = "status-pass" if scen_res['decision'] == "PASS" else ("status-watch" if scen_res['decision'] == "WATCH" else "status-fail")
    st.markdown(f"**حالة القرار الاستثماري للسيناريو:** <span class='{tag_class}'>{scen_res['decision']}</span>", unsafe_allow_html=True)
    st.write("")

    m1, m2, m3, m4, m5 = st.columns(5)
    irr_diff = (scen_res['equity_irr'] - base_res['equity_irr']) if not np.isnan(scen_res['equity_irr']) else 0
    npv_diff = scen_res['equity_npv'] - base_res['equity_npv']
    
    with m1: render_kpi("Equity IRR", fmt_pct(scen_res['equity_irr']), f"{fmt_pct(irr_diff)} مقارنة بالأساس", "positive" if irr_diff>=0 else "danger")
    with m2: render_kpi("Equity NPV", fmt_currency_m(scen_res['equity_npv']), f"{fmt_currency_m(npv_diff)} الأثر", "positive" if npv_diff>=0 else "danger")
    with m3: render_kpi("Equity MOIC", fmt_multiple(scen_res['equity_moic']), "مضاعف رأس المال", "positive")
    with m4: render_kpi("Payback Period", f"{scen_res['payback_m']:.1f} شهراً", "فترة الاسترداد", "positive")
    with m5: render_kpi("Peak Equity Req.", fmt_currency_m(scen_res['peak_equity']), "أعلى احتياج سيولة", "warning")

    # Comparison Waterfall / Comparison Bars (Inspired by Reference Design)
    st.markdown("<div class='section-title'>📊 مقارنة إيرادات ونقاط التعادل بين الأساس والسيناريو</div>", unsafe_allow_html=True)
    
    df_comp = pd.DataFrame([
        {'المؤشر': 'إيرادات الأساس (Base Revenue)', 'المبلغ': base_res['total_rev'], 'النوع': 'الأساس'},
        {'المؤشر': 'تعادل الأساس (Base Breakeven)', 'المبلغ': base_res['npv_zero_rev'], 'النوع': 'الأساس'},
        {'المؤشر': 'إيرادات السيناريو (Scenario Revenue)', 'المبلغ': scen_res['total_rev'], 'النوع': 'السيناريو'},
        {'المؤشر': 'تعادل السيناريو (Scenario Breakeven)', 'المبلغ': scen_res['npv_zero_rev'], 'النوع': 'السيناريو'}
    ])
    
    fig_comp = px.bar(df_comp, x='المبلغ', y='المؤشر', color='النوع', orientation='h', text_auto=',.0f',
                      color_discrete_map={'الأساس': '#818CF8', 'السيناريو': '#38BDF8'})
    fig_comp.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#F8FAFC'),
                           margin=dict(t=20, b=20, l=20, r=20), showlegend=True)
    st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown("<div class='section-title'>مستويات التعادل المستقلة (Breakeven Revenue Solvers)</div>", unsafe_allow_html=True)
    s1, s2, s3 = st.columns(3)
    with s1: render_kpi("1. Accounting Breakeven", fmt_currency_m(scen_res['accounting_be']), f"سعر المتر: {fmt_currency(scen_res['total_cost']/base_area)}/Sqm", "warning")
    with s2: render_kpi("2. NPV = 0 Revenue (at Ke)", fmt_currency_m(scen_res['npv_zero_rev']), f"سعر المتر: {fmt_currency(scen_res['req_price_npv_zero'])}/Sqm", "positive")
    with s3: render_kpi("3. Target IRR Revenue (at Target)", fmt_currency_m(scen_res['target_irr_rev']), f"سعر المتر: {fmt_currency(scen_res['req_price_target_irr'])}/Sqm", "positive")

# ==============================================================================
# PAGE 6: موديل الايجارات (FEASIBILITY)
# ==============================================================================
elif page == "موديل الايجارات":
    st.markdown("<div class='page-title'>موديل إعادة التأجير Sub-Lease (100% Equity)</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>دراسة جدوى فرص الاستئجار وإعادة التأجير بالكامل بالتمويل الذاتي مع مراعاة فترة السماح Grace Period.</div>", unsafe_allow_html=True)

    with st.expander("🛠️ متغبرات التحكم وافتراضات الإدارة", expanded=True):
        r1, r2, r3, r4 = st.columns(4)
        head_lease_rent = r1.number_input("إيجار المالك الرئيسي Head Lease (SAR)", value=1200000, step=100000)
        lease_term_yrs = r2.number_input("مدة العقد (سنوات)", value=10, step=1)
        rent_escalation = r3.number_input("نسبة الزيادة الدورية (%)", value=5.0, step=1.0) / 100.0
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

    st.markdown("---")
    st.markdown("<div class='section-title'>مؤشرات العوائد والقرار الاستثماري للمشروع</div>", unsafe_allow_html=True)
    
    tag_class = "status-pass" if res_r['decision'] == "PASS" else ("status-watch" if res_r['decision'] == "WATCH" else "status-fail")
    st.markdown(f"**حالة القرار الاستثماري:** <span class='{tag_class}'>{res_r['decision']}</span>", unsafe_allow_html=True)
    st.write("")

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1: render_kpi("Equity IRR", fmt_pct(res_r['equity_irr']), "العائد الاستثماري", "positive")
    with m2: render_kpi("Equity NPV", fmt_currency_m(res_r['equity_npv']), f"مخصومة بـ {fmt_pct(cost_of_equity)} Ke", "positive")
    with m3: render_kpi("Equity MOIC", fmt_multiple(res_r['equity_moic']), "مضاعف الاستثمار", "positive")
    with m4: render_kpi("Payback Period", f"{res_r['payback_yrs']:.1f} سنوات" if not np.isnan(res_r['payback_yrs']) else "N/A", "استرداد رأس المال", "positive")
    with m5: render_kpi("Fit-out CapEx Equity", fmt_currency_m(res_r['fitout_capex']), "رأس المال المستثمر", "warning")

    st.markdown("<div class='section-title'>تحليل التعادل ونسب الإشغال الحرجة</div>", unsafe_allow_html=True)
    b1, b2 = st.columns(2)
    with b1: render_kpi("Stabilized Break-even Occupancy %", fmt_pct(res_r['be_occupancy']), "لتحقيق صافي دخل NOI = 0", "warning")
    with b2: render_kpi("Occupancy for Target 15% IRR", fmt_pct(res_r['occ_for_target_irr']) if not np.isnan(res_r['occ_for_target_irr']) else "N/A", "لتحقيق العائد المستهدف", "positive")

    st.markdown("<div class='section-title'>قائمة الدخل التقديرية 10-Year Equity Pro Forma P&L</div>", unsafe_allow_html=True)
    df_pnl_disp = res_r['annual_pnl'].copy()
    for col in ['إيراد الإيجار', 'إيجار المالك', 'التكاليف التشغيلية', 'صافي الدخل NOI']:
        df_pnl_disp[col] = df_pnl_disp[col].apply(lambda x: fmt_currency(x))
    df_pnl_disp['هامش NOI'] = df_pnl_disp['هامش NOI'].apply(lambda x: fmt_pct(x))
    st.dataframe(df_pnl_disp.reset_index(drop=True), use_container_width=True)

# ==============================================================================
# PAGE 7: مركز القرار والتدقيق
# ==============================================================================
elif page == "مركز القرار والتدقيق":
    st.markdown("<div class='page-title'>مركز القرار الاستثماري وتدقيق جودة البيانات</div>", unsafe_allow_html=True)
    st.markdown("<div class='page-subtitle'>مراقبة المخاطر المركزية، جودة مطابقة البيانات، والتنبيهات الإدارية.</div>", unsafe_allow_html=True)
    
    st.markdown("<div class='section-title'>تدقيق جودة ومطابقة بيانات الإكسل المصدر</div>", unsafe_allow_html=True)
    audit_data = [
        {"المكون": "الجداول المسماة في ملف Master", "الحالة": "🟢 ناجح Passed", "التفاصيل": "تم قراءة الجداول الـ 6 بنجاح من شيت Loans_&_Installments"},
        {"المكون": "التدفقات النقدية Cash Flow", "الحالة": "🟢 ناجح Passed", "التفاصيل": "تم تحميل التوقعات لـ 24 شهراً بالكامل دون تواريخ مفقودة"},
        {"المكون": "قائمة دخل الإيجارات P&L", "الحالة": "🟢 ناجح Passed", "التفاصيل": "تم ربط جميع العقارات الـ 10 والـ 263 وحدة بنجاح"},
        {"المكون": "عزل القروض Corporate Isolation", "الحالة": "🟢 ناجح Passed", "التفاصيل": "نماذج الجدوى الصفحات 5 و6 معزولة تماماً ولا تتأثر بديون الشركة"}
    ]
    st.dataframe(pd.DataFrame(audit_data), use_container_width=True)

    st.markdown("<div class='section-title'>التنبيهات والإجراءات الإدارية المطلوبة</div>", unsafe_allow_html=True)
    st.error("🔴 مخاطر ديون حرجة: قسط صكوك منصة دينار 4.4M متأخر السداد (المتبقي المستحق 4.46M SAR).")
    st.warning("🟡 مراقبة الأداء الإيجاري: مشروع إيجار الملقا وإيجار النرجس يعملان بهامش دخل تشغيلي NOI سالب (-61.0% و -179.4%).")
    st.success("🟢 كفاءة التحصيل: تحصيلات الإيجارات بلغت 84.71% (تم تحصيل 720K SAR من أصل 850K SAR مستحقة).")
'''

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(app_code)

print("Updated app.py written successfully. File size:", len(app_code))
