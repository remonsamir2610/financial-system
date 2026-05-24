"""
النظام المالي والزكوي الشامل - الإصدار المستقر والنهائي (v12.0)
تصميم: Navy & Emerald | Remon Samir Saad
متوافق مع هيئة الزكاة والضريبة والجمارك (ZATCA) - اللائحة التنفيذية 1445هـ
"""

OWNER_NAME_EN = "Remon Samir Saad"
OWNER_NAME_AR = "ريمون سمير سعد"
OWNER_INITIALS = "RS"
OWNER_ROLE_AR  = "مدير النظام"
OWNER_ROLE_EN  = "System Admin"
APP_VERSION    = "v12.0"

import streamlit as st
import pandas as pd
import io
import hashlib
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
import warnings

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title=f"النظام المالي | {OWNER_NAME_EN} | ZATCA {APP_VERSION}",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800;900&display=swap');
* { font-family: 'Tajawal', sans-serif !important; box-sizing: border-box; }
.stApp { direction: rtl; background: #E8EDF2; }
.hbanner { background: #0F2942; border-radius: 12px; padding: 18px 28px; margin-bottom: 18px; display: flex; align-items: center; justify-content: space-between; border-bottom: 3px solid #10B981; }
.hbanner h1 { color:#F0F9FF; font-size:1.5rem; font-weight:700; margin:0 0 3px; }
.hbanner p  { color:#475569; font-size:0.82rem; margin:0; }
.hbanner-right { display:flex; flex-direction:column; align-items:flex-end; gap:6px; }
.zatca-pill { background:#10B981; color:#064E3B; padding:3px 12px; border-radius:20px; font-size:0.72rem; font-weight:700; }
.user-chip { display:flex; align-items:center; gap:8px; background:rgba(255,255,255,0.06); border:0.5px solid rgba(255,255,255,0.1); border-radius:8px; padding:5px 10px; }
.user-avatar { width:28px; height:28px; background:#10B981; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:700; color:#064E3B; }
.user-name  { color:#CBD5E1; font-size:12px; font-weight:500; }
.user-role  { color:#475569; font-size:10px; }
.sec-title { background:#0F2942; color:#F0F9FF; padding:9px 16px; border-radius:7px; font-weight:600; margin:16px 0 10px; border-right:4px solid #10B981; font-size:0.88rem; }
.kpi { background:#fff; border:0.5px solid #E2E8F0; border-radius:10px; padding:14px 12px; text-align:center; }
.kpi .val { font-size:1.3rem; font-weight:700; color:#0F2942; }
.kpi .lbl { font-size:0.75rem; color:#64748B; margin-bottom:4px; }
.zatca-badge { background:#10B981; color:#064E3B; padding:4px 12px; border-radius:20px; font-size:0.75rem; font-weight:700; display:inline-block; margin-bottom:8px; }
.footer-wm { background:#0A1E30; border-top:0.5px solid rgba(255,255,255,0.06); padding:8px 20px; border-radius:0 0 10px 10px; display:flex; align-items:center; justify-content:space-between; margin-top:18px; }
.footer-wm .ft-left { display:flex; align-items:center; gap:8px; }
.footer-wm .ft-av { width:24px; height:24px; background:#10B981; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:9px; font-weight:700; color:#064E3B; }
.footer-wm .ft-name { color:#94A3B8; font-size:10px; }
.footer-wm .ft-name span { color:#CBD5E1; font-weight:500; }
.footer-wm .ft-right { font-size:9px; color:#334155; }
</style>
""", unsafe_allow_html=True)

# ── نظام المستخدمين ──
USERS = {
    "remon":  {"hash": hashlib.sha256("remon123".encode()).hexdigest(), "role": OWNER_ROLE_AR, "name_en": OWNER_NAME_EN, "name_ar": OWNER_NAME_AR, "initials": OWNER_INITIALS},
}

def show_login():
    st.markdown("<style>.stApp{background:#0A1E30!important;}</style>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown(f"""
        <div style="background:#132D45;border:0.5px solid rgba(16,185,129,0.25); border-radius:14px;padding:32px 28px;margin-top:36px;">
          <div style="text-align:center;margin-bottom:22px;">
            <div style="width:56px;height:56px;background:#10B981;border-radius:12px; display:inline-flex;align-items:center;justify-content:center; font-size:28px;margin-bottom:10px;">📊</div>
            <div style="color:#F0F9FF;font-size:15px;font-weight:700;">النظام المالي والزكوي</div>
            <div style="color:#475569;font-size:11px;margin-top:3px;">ZATCA Financial System {APP_VERSION}</div>
            <div style="margin-top:14px;padding-top:12px;border-top:0.5px solid rgba(255,255,255,0.08);">
              <div style="display:inline-flex;align-items:center;gap:8px;background:rgba(16,185,129,0.1);border:0.5px solid rgba(16,185,129,0.3);border-radius:8px;padding:6px 14px;">
                <div style="width:30px;height:30px;background:#10B981;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;color:#064E3B;">{OWNER_INITIALS}</div>
                <div style="text-align:right;">
                  <div style="color:#CBD5E1;font-size:11px;font-weight:600;">تطوير وبرمجة</div>
                  <div style="color:#10B981;font-size:12px;font-weight:700;">{OWNER_NAME_AR} | {OWNER_NAME_EN}</div>
                </div>
              </div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        with st.form("login_form"):
            user = st.text_input("اسم المستخدم", value="remon")
            pwd  = st.text_input("كلمة المرور", type="password")
            if st.form_submit_button("دخول ←", use_container_width=True):
                u = USERS.get(user)
                if u and u["hash"] == hashlib.sha256(pwd.encode()).hexdigest():
                    st.session_state.user = u
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("❌ بيانات غير صحيحة")

def fmt(v, zero=True):
    if pd.isna(v): return "–"
    if v == 0: return "–" if zero else "0.00"
    return f"{v:,.2f}"

def fmt_pct(v):
    if pd.isna(v) or v == 0: return "–"
    return f"{v:.1%}"

COMPANY_TYPES = {
    "مؤسسة فردية": {"zatca_rate": 0.025}, "شركة ذات مسؤولية محدودة": {"zatca_rate": 0.025},
    "شركة أشخاص (تضامن)": {"zatca_rate": 0.025}, "شركة مساهمة": {"zatca_rate": 0.025},
}

# ── شجرة الحسابات والهيكل المالي ──
_main_cats = ["قائمة المركز المالي", "قائمة الدخل"]
_sub_cats  = ["الأصول المتداولة", "الأصول غير المتداولة", "الخصوم المتداولة", "الخصوم غير المتداولة", "حقوق الملكية", "الإيرادات", "تكاليف البضاعة المباعة", "المصاريف", "غير مبوب"]
_acc_cats = ["النقدية وما في حكمها", "مدينون تجاريون", "مخزون بضاعة", "مصاريف مقدمة وأرصدة مدينة أخرى", "أطراف ذات علاقة (مدين)", "الممتلكات والآلات والمعدات (صافي)", "الاستثمارات في شركات تابعة", "أصول حق الاستخدام", "الشهرة والأصول غير الملموسة", "موردون تجاريون", "مصاريف مستحقة وأرصدة دائنة أخرى", "أطراف ذات علاقة (دائن)", "مخصص الزكاة الشرعية", "تسهيلات ائتمانية قصيرة الأجل", "الجزء المتداول من القروض طويلة الأجل", "مخصص مكافأة نهاية الخدمة", "قروض طويلة الأجل", "التزامات عقود الإيجار", "مخصص الضمان والصيانة", "رأس المال", "احتياطي نظامي", "الأرباح المبقاة", "مسحوبات الشركاء", "احتياطيات أخرى وتسويات", "الإيرادات والمبيعات", "مردودات ومسموحات المبيعات", "الحسم الممنوح", "إيرادات أخرى", "تكلفة الإيرادات", "الرواتب والأجور", "عمولات البيع", "الإيجارات", "التأمينات الاجتماعية (GOSI)", "مصاريف المستودع والتالف", "ديون معدومة ومشكوك فيها", "النقل والتوزيع والمحروقات", "الكهرباء والمياه والاتصالات", "الاستهلاك والإهلاك", "أتعاب مراجعة واستشارات", "رسوم حكومية وتراخيص وتأشيرات", "مصاريف تسويق وإعلان", "تكاليف تمويل وفوائد بنكية", "مصاريف عمومية وإدارية أخرى"]
_cost_centers = ["المركز الرئيسي", "الرياض", "جدة", "الدمام", "مكة المكرمة", "المدينة المنورة", "أبها", "تبوك", "حائل", "القصيم", "نجران", "جازان", "الجوف", "الباحة", "عرعر", "إدارة عامة", "فرع 1", "فرع 2", "مستودع رئيسي", "مشروع خاص"]
_hierarchy_map = [("قائمة المركز المالي", "الأصول المتداولة", "النقدية وما في حكمها"), ("قائمة المركز المالي", "الأصول المتداولة", "مدينون تجاريون"), ("قائمة المركز المالي", "الأصول المتداولة", "مخزون بضاعة"), ("قائمة المركز المالي", "الأصول المتداولة", "مصاريف مقدمة وأرصدة مدينة أخرى"), ("قائمة المركز المالي", "الأصول المتداولة", "أطراف ذات علاقة (مدين)"), ("قائمة المركز المالي", "الأصول غير المتداولة", "الممتلكات والآلات والمعدات (صافي)"), ("قائمة المركز المالي", "الأصول غير المتداولة", "الاستثمارات في شركات تابعة"), ("قائمة المركز المالي", "الخصوم المتداولة", "موردون تجاريون"), ("قائمة المركز المالي", "الخصوم المتداولة", "مصاريف مستحقة وأرصدة دائنة أخرى"), ("قائمة المركز المالي", "الخصوم المتداولة", "أطراف ذات علاقة (دائن)"), ("قائمة المركز المالي", "الخصوم المتداولة", "مخصص الزكاة الشرعية"), ("قائمة المركز المالي", "الخصوم غير المتداولة", "مخصص مكافأة نهاية الخدمة"), ("قائمة المركز المالي", "الخصوم غير المتداولة", "قروض طويلة الأجل"), ("قائمة المركز المالي", "حقوق الملكية", "رأس المال"), ("قائمة المركز المالي", "حقوق الملكية", "احتياطي نظامي"), ("قائمة المركز المالي", "حقوق الملكية", "الأرباح المبقاة"), ("قائمة المركز المالي", "حقوق الملكية", "مسحوبات الشركاء"), ("قائمة الدخل", "الإيرادات", "الإيرادات والمبيعات"), ("قائمة الدخل", "الإيرادات", "مردودات ومسموحات المبيعات"), ("قائمة الدخل", "تكاليف البضاعة المباعة", "تكلفة الإيرادات"), ("قائمة الدخل", "المصاريف", "الرواتب والأجور"), ("قائمة الدخل", "المصاريف", "الإيجارات"), ("قائمة الدخل", "المصاريف", "التأمينات الاجتماعية (GOSI)"), ("قائمة الدخل", "المصاريف", "الاستهلاك والإهلاك"), ("قائمة الدخل", "المصاريف", "مصاريف عمومية وإدارية أخرى")]

# ── دالة قراءة وتفسير ميزان المراجعة ──
def parse_tb(file) -> pd.DataFrame:
    xl = pd.ExcelFile(file)
    tb_sheets = [s for s in xl.sheet_names if any(k in s for k in ["ميزان","مراجع","TB","trial"])]
    sheet = tb_sheets[0] if tb_sheets else xl.sheet_names[0]

    df_raw = pd.read_excel(file, sheet_name=sheet, header=None, nrows=15)
    header_row = 1
    for i, row in df_raw.iterrows():
        if sum(1 for kw in ["مدين","دائن","حساب","تبويب","تفصيل"] if kw in str(row.values).lower()) >= 2:
            header_row = i; break

    df = pd.read_excel(file, sheet_name=sheet, skiprows=header_row, header=0)
    df = df.dropna(axis=1, how='all').dropna(how='all')

    assigned_cols = {}
    for c in df.columns:
        lc = str(c).lower().replace("\n"," ").strip()
        t = None
        if "مدين" in lc and any(k in lc for k in ["سابق","prev"]): t = "dr_prev"
        elif "دائن" in lc and any(k in lc for k in ["سابق","prev"]): t = "cr_prev"
        elif "مدين" in lc: t = "dr_curr"
        elif "دائن" in lc: t = "cr_curr"
        elif "رئيس" in lc or "main" in lc: t = "main_cat"
        elif "تبويب" in lc or "تصنيف" in lc: t = "cat"
        elif "تفصيلي" in lc or "detail" in lc: t = "detail"
        elif "مركز" in lc or "cost" in lc: t = "cost_center"
        elif "اسم" in lc or "بيان" in lc or "account" in lc or "حساب" in lc: 
            t = "name" if "name" not in assigned_cols.values() else "detail"
        
        if t and t not in assigned_cols.values(): assigned_cols[c] = t

    # 🟢 تم الإصلاح هنا ليعمل التوجيه بشكل صحيح ومباشر 100% بدون قلب المصفوفة
    df = df.rename(columns=assigned_cols)
    
    text_cols = ["name", "main_cat", "cat", "detail", "cost_center"]
    for c in text_cols:
        if c not in df.columns: df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
        df.loc[df[c] == "nan", c] = ""

    num_cols = ["dr_curr", "cr_curr", "dr_prev", "cr_prev"]
    for c in num_cols:
        if c not in df.columns: df[c] = 0.0
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", "").str.replace("–", "0").str.replace("−", "0").str.strip(), errors="coerce").fillna(0.0)

    df = df[~df["name"].str.contains("الإجمالي|فحص|مثال", case=False, na=False)]
    
    df["net_curr"] = df["dr_curr"] - df["cr_curr"]
    df["net_prev"] = df["dr_prev"] - df["cr_prev"]

    return df[df["name"].str.len() > 1].reset_index(drop=True)

# ── المحرك المالي والزكوي الذكي ──
def calc_financials(df: pd.DataFrame, net_col: str, zakat_adj: float, company_type: str) -> dict:
    
    def get_amt(details, fallbacks=None):
        mask = df["detail"].isin(details)
        if fallbacks:
            pat = "|".join(fallbacks)
            mask_fall = df["name"].str.contains(pat, case=False, regex=True, na=False) & (df["detail"] == "")
            mask = mask | mask_fall
        return df.loc[mask, net_col].sum()

    # الأصول
    cash         = get_amt(["النقدية وما في حكمها"], ["نقد", "صندوق", "بنك", "cash"])
    receivables  = get_amt(["مدينون تجاريون"], ["عملاء", "مدينون", "ذمم مدينة", "receivable"])
    inventory    = get_amt(["مخزون بضاعة"], ["مخزون", "بضاعة", "inventory"])
    prepaid      = get_amt(["مصاريف مقدمة وأرصدة مدينة أخرى"], ["مقدمة", "عهد", "سلف", "prepaid"])
    
    m_rel_rec_fall = df["name"].str.contains("أطراف ذات علاقة|related|شريك", case=False, na=False) & ~df["name"].str.contains("دائن|التزام|pay", case=False, na=False) & (df["detail"] == "")
    related_rec  = df.loc[df["detail"] == "أطراف ذات علاقة (مدين)", net_col].sum() + df.loc[m_rel_rec_fall, net_col].sum()
    
    total_current = cash + receivables + inventory + prepaid + related_rec

    fixed_assets = get_amt(["الممتلكات والآلات والمعدات (صافي)", "أصول حق الاستخدام", "الشهرة والأصول غير الملموسة"], ["أصول ثابتة", "ممتلكات", "معدات", "آلات", "سيارات", "furniture", "equipment"])
    investments  = get_amt(["الاستثمارات في شركات تابعة"], ["استثمار", "invest"])
    total_nca    = fixed_assets + investments
    total_assets = total_current + total_nca

    # الخصوم وحقوق الملكية
    payables     = -get_amt(["موردون تجاريون"], ["موردون", "دائنون", "payable"])
    accruals     = -get_amt(["مصاريف مستحقة وأرصدة دائنة أخرى", "مخصص الزكاة الشرعية"], ["مستحقة", "accrued", "أمانات"])
    short_loans  = -get_amt(["تسهيلات ائتمانية قصيرة الأجل", "الجزء المتداول من القروض طويلة الأجل"], ["تسهيلات", "قرض قصير"])
    
    m_rel_pay_fall = df["name"].str.contains("أطراف ذات علاقة|related|شريك", case=False, na=False) & df["name"].str.contains("دائن|التزام|pay", case=False, na=False) & (df["detail"] == "")
    related_pay  = -(df.loc[df["detail"] == "أطراف ذات علاقة (دائن)", net_col].sum() + df.loc[m_rel_pay_fall, net_col].sum())
    total_cl     = payables + accruals + short_loans + related_pay

    eosb         = -get_amt(["مخصص مكافأة نهاية الخدمة"], ["نهاية الخدمة", "eosb"])
    long_loans   = -get_amt(["قروض طويلة الأجل", "التزامات عقود الإيجار", "مخصص الضمان والصيانة"], ["قرض طويل", "التزام ايجار"])
    total_ncl    = eosb + long_loans

    capital      = -get_amt(["رأس المال"], ["رأس المال", "capital"])
    reserves     = -get_amt(["احتياطي نظامي", "احتياطيات أخرى وتسويات"], ["احتياطي", "reserve"])
    retained     = -get_amt(["الأرباح المبقاة"], ["مبقاة", "أرباح سابقة", "retained"])
    withdrawals  = get_amt(["مسحوبات الشركاء"], ["مسحوبات", "سحب", "جاري"])

    # قائمة الدخل
    gross_sales  = -get_amt(["الإيرادات والمبيعات"], ["إيراد", "مبيعات", "sales"])
    returns      = get_amt(["مردودات ومسموحات المبيعات", "الحسم الممنوح"], ["مردود", "خصم", "return"])
    net_sales    = gross_sales - returns
    cogs         = get_amt(["تكلفة الإيرادات"], ["تكلفة", "cogs", "بضاعة مباعة"])
    gross_profit = net_sales - cogs

    salaries     = get_amt(["الرواتب والأجور"], ["رواتب", "أجور", "salary"])
    rent         = get_amt(["الإيجارات"], ["ايجار", "إيجار"])
    gosi         = get_amt(["التأمينات الاجتماعية (GOSI)"], ["تأمينات", "gosi"])
    marketing    = get_amt(["مصاريف تسويق وإعلان", "عمولات البيع"], ["تسويق", "إعلان", "عمول"])
    transport    = get_amt(["النقل والتوزيع والمحروقات"], ["نقل", "شحن", "محروقات"])
    utilities    = get_amt(["الكهرباء والمياه والاتصالات"], ["كهرباء", "مياه", "اتصالات"])
    depreciation = get_amt(["الاستهلاك والإهلاك"], ["إهلاك", "استهلاك", "depreciation"])
    
    misc_mapped  = get_amt(["مصاريف المستودع والتالف", "ديون معدومة ومشكوك فيها", "أتعاب مراجعة واستشارات", "رسوم حكومية وتراخيص وتأشيرات", "مصاريف عمومية وإدارية أخرى"])
    m_exp_fall   = df["cat"].str.contains("مصاريف", na=False) & (df["detail"] == "")
    m_exp_name   = df["name"].str.contains("رسوم|صيانة|تالف|متنوع|مصاريف|expense", case=False, regex=True, na=False) & (df["detail"] == "")
    misc_exp     = misc_mapped + df.loc[m_exp_fall, net_col].sum() + df.loc[m_exp_name, net_col].sum()

    total_opex   = salaries + rent + marketing + gosi + depreciation + transport + utilities + misc_exp
    op_profit    = gross_profit - total_opex
    
    other_inc    = -get_amt(["إيرادات أخرى"], ["دخل آخر", "other income"])
    finance_cost = get_amt(["تكاليف تمويل وفوائد بنكية"], ["فوائد", "تمويل", "finance"])
    net_bz       = op_profit + other_inc - finance_cost

    # حساب الوعاء الزكوي
    net_adj = net_bz + zakat_adj
    additions = capital + reserves + retained - withdrawals + eosb + long_loans
    if net_adj > 0: additions += net_adj
    deductions = total_nca

    zakat_base_calc = additions - deductions
    min_base = total_current + max(zakat_adj, 0)
    equity_raw = capital + reserves + retained - withdrawals + net_adj
    max_base = equity_raw + max(zakat_adj, 0)

    zakat_base = max(zakat_base_calc, min_base if net_adj > 0 else 0)
    zakat_base = min(zakat_base, max_base) if max_base > 0 else zakat_base
    z_rate = 0.025
    
    if net_adj > 0: zakat_due = (net_adj * z_rate) + (max(zakat_base - net_adj, 0) * (z_rate * 365/354))
    else: zakat_due = zakat_base * (z_rate * 365/354)
    zakat_due = max(zakat_due, 0)
    net_az = net_bz - zakat_due

    equity = capital + reserves + retained - withdrawals + net_az
    total_eq_lb = equity + total_cl + total_ncl + zakat_due

    return {
        "net_sales": net_sales, "gross_sales": gross_sales, "returns": returns,
        "gross_profit": gross_profit, "op_profit": op_profit, "net_bz": net_bz, "zakat_due": zakat_due, "net_az": net_az,
        "salaries": salaries, "rent": rent, "marketing": marketing, "gosi": gosi, "depreciation": depreciation,
        "transport": transport, "utilities": utilities, "misc_exp": misc_exp, "total_opex": total_opex,
        "cogs": cogs, "other_inc": other_inc, "finance_cost": finance_cost,
        "cash": cash, "receivables": receivables, "inventory": inventory, "prepaid": prepaid, "related_rec": related_rec,
        "total_current": total_current, "fixed_assets": fixed_assets, "investments": investments, "total_nca": total_nca,
        "total_assets": total_assets, "payables": payables, "accruals": accruals, "short_loans": short_loans,
        "related_pay": related_pay, "total_cl": total_cl, "eosb": eosb, "long_loans": long_loans, "total_ncl": total_ncl,
        "capital": capital, "reserves": reserves, "retained": retained, "withdrawals": withdrawals, "equity": equity,
        "total_eq_lb": total_eq_lb, "net_adj": net_adj, "additions": additions, "deductions": deductions,
        "zakat_base": zakat_base, "min_base": min_base, "max_base": max_base,
    }

# ── شاشات عرض واجهة المستخدم وعناصر التحليل ──
def show_auto_expenses(df, year):
    st.markdown('<div class="sec-title">💸 تحليل المصروفات التفصيلي (مقارنة السنتين)</div>', unsafe_allow_html=True)
    exp_df = df[df["cat"].str.contains("مصاريف|مصروفات", na=False) | df["detail"].str.contains("مصاريف", na=False) | df["main_cat"].str.contains("الدخل", na=False)].copy()
    exp_df = exp_df[~exp_df["cat"].str.contains("إيراد|تكلف", na=False) & ~exp_df["detail"].str.contains("إيراد|مبيعات|تكلفة", na=False)]
    
    if exp_df.empty:
        st.warning("⚠️ لا توجد حسابات تم تصنيفها كمصروفات.")
        return

    exp_df["الفرق"] = exp_df["net_curr"] - exp_df["net_prev"]
    exp_df["% النمو"] = (exp_df["الفرق"] / exp_df["net_prev"].replace(0, 1).abs()) * 100

    c1, c2, c3 = st.columns(3)
    total_c = exp_df["net_curr"].sum()
    total_p = exp_df["net_prev"].sum()
    var = total_c - total_p
    c1.metric(f"إجمالي المصروفات {year}", f"{total_c:,.0f} ر.س")
    c2.metric(f"إجمالي المصروفات {year-1}", f"{total_p:,.0f} ر.س")
    c3.metric("الانحراف الإجمالي", f"{var:,.0f} ر.س", delta="تجاوز الميزانية ❌" if var > 0 else "وفر في الميزانية ✅", delta_color="inverse")

    disp = exp_df.groupby("name", as_index=False)[["net_curr","net_prev","الفرق","% النمو"]].sum()
    disp.columns = ["اسم الحساب", f"العام {year}", f"العام {year-1}", "الفرق", "% النمو"]
    disp = disp[disp[f"العام {year}"].abs() > 0].sort_values(f"العام {year}", ascending=False).reset_index(drop=True)

    disp_display = disp.copy()
    for c in [f"العام {year}", f"العام {year-1}", "الفرق"]: disp_display[c] = disp_display[c].apply(lambda x: f"{x:,.2f}")
    disp_display["% النمو"] = disp_display["% النمو"].apply(lambda x: f"{x:.1f}%")
    st.dataframe(disp_display, use_container_width=True)

    if len(disp) > 0:
        top_10 = disp.head(10).copy()
        fig = go.Figure(go.Bar(x=top_10[f"العام {year}"], y=top_10["اسم الحساب"], orientation='h', marker=dict(color=top_10[f"العام {year}"], colorscale="Viridis", showscale=True)))
        fig.update_layout(title=f"أعلى 10 مصروفات - {year}", yaxis={'categoryorder':'total ascending'}, height=400, paper_bgcolor="#fff", plot_bgcolor="#F8FAFC")
        st.plotly_chart(fig, use_container_width=True)

def show_zakat_report(r, zakat_adj, year):
    st.markdown('<div class="sec-title">🕌 تقرير الزكاة (مطابق للائحة ZATCA 1445هـ)</div>', unsafe_allow_html=True)
    st.markdown('<span class="zatca-badge">✅ متوافق مع هيئة الزكاة والضريبة والجمارك</span>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**أولاً: صافي الربح المعدل**")
        st.table(pd.DataFrame({
            "البند": ["الربح قبل الزكاة", f"التعديلات الزكوية (مادة 4)", "صافي الربح المعدل"],
            "المبلغ (ر.س)": [fmt(r["net_bz"]), fmt(zakat_adj), fmt(r["net_adj"])]
        }))

    with col2:
        st.markdown("**ثانياً: وعاء الزكاة**")
        st.table(pd.DataFrame({
            "البند": ["الإضافات", "الحسميات (الأصول غير المتداولة)", "الوعاء المحسوب", "الحد الأدنى (الأصول المتداولة)", "الحد الأعلى (حقوق الملكية)", "الوعاء النهائي"],
            "المبلغ (ر.س)": [fmt(r["additions"]), fmt(-r["deductions"]), fmt(r["zakat_base"]), fmt(r["min_base"]), fmt(r["max_base"]), fmt(r["zakat_base"])]
        }))

    st.success(f"✅ **الزكاة الشرعية المستحقة: {r['zakat_due']:,.2f} ريال سعودي**")

def show_cash_flow(r, r_prev, report_year):
    st.markdown(f'<div class="sec-title">💵 قائمة التدفقات النقدية | عن السنة المنتهية 31 ديسمبر {report_year}</div>', unsafe_allow_html=True)

    # حساب التدفقات
    op_cf   = r["net_bz"] + r["depreciation"] \
              + (-r["receivables"] + r_prev["receivables"]) \
              + (-r["inventory"]   + r_prev["inventory"]) \
              + (-r["prepaid"]     + r_prev["prepaid"]) \
              + (r["payables"]  - r_prev["payables"]) \
              + (r["accruals"]  - r_prev["accruals"]) \
              - r["zakat_due"]
    op_cf_p = r_prev["net_bz"] + r_prev["depreciation"] - r_prev["zakat_due"]
    inv_cf  = -(r["fixed_assets"] - r_prev["fixed_assets"]) - (r["investments"] - r_prev["investments"])
    inv_cf_p= 0
    fin_cf  = -r["withdrawals"] + (r["long_loans"] - r_prev["long_loans"]) + (r["short_loans"] - r_prev["short_loans"])
    fin_cf_p= -r_prev["withdrawals"]
    net_chg = op_cf + inv_cf + fin_cf
    net_chg_p = op_cf_p + inv_cf_p + fin_cf_p

    # بطاقات ملخص
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📥 التشغيل", f"{op_cf:,.0f} ر.س", delta="إيجابي ✅" if op_cf >= 0 else "سلبي ⚠️", delta_color="normal" if op_cf >= 0 else "inverse")
    c2.metric("🏗 الاستثمار", f"{inv_cf:,.0f} ر.س")
    c3.metric("🏦 التمويل", f"{fin_cf:,.0f} ر.س")
    c4.metric("📊 صافي التغير", f"{net_chg:,.0f} ر.س", delta="زيادة ✅" if net_chg >= 0 else "انخفاض ⚠️", delta_color="normal" if net_chg >= 0 else "inverse")

    st.markdown("---")
    cf_data = {
        "البيـــان": [
            "🔵 أولاً: الأنشطة التشغيلية",
            "صافي الربح قبل الزكاة",
            "يضاف: الاستهلاك والإهلاك (بنود غير نقدية)",
            "(زيادة)/نقص في المدينين التجاريين",
            "(زيادة)/نقص في المخزون السلعي",
            "(زيادة)/نقص في الأرصدة المدينة الأخرى",
            "زيادة/(نقص) في الدائنين والموردين",
            "زيادة/(نقص) في المصاريف المستحقة",
            "الزكاة المدفوعة",
            "✅ صافي النقد من الأنشطة التشغيلية",
            "",
            "🟠 ثانياً: الأنشطة الاستثمارية",
            "التغير في الأصول الثابتة (صافي)",
            "التغير في الاستثمارات",
            "✅ صافي النقد من الأنشطة الاستثمارية",
            "",
            "🟣 ثالثاً: الأنشطة التمويلية",
            "مسحوبات الشركاء",
            "التغير في القروض والتسهيلات",
            "✅ صافي النقد من الأنشطة التمويلية",
            "",
            "📌 صافي التغير في النقد وما في حكمه",
            f"النقد وما في حكمه - بداية الفترة ({report_year-1})",
            f"🏁 النقد وما في حكمه - نهاية الفترة ({report_year})",
        ],
        f"{report_year} (ر.س)": [
            "", fmt(r["net_bz"]), fmt(r["depreciation"]),
            fmt(-r["receivables"]+r_prev["receivables"]),
            fmt(-r["inventory"]+r_prev["inventory"]),
            fmt(-r["prepaid"]+r_prev["prepaid"]),
            fmt(r["payables"]-r_prev["payables"]),
            fmt(r["accruals"]-r_prev["accruals"]),
            fmt(-r["zakat_due"]),
            fmt(op_cf),
            "",
            "", fmt(-(r["fixed_assets"]-r_prev["fixed_assets"])),
            fmt(-(r["investments"]-r_prev["investments"])),
            fmt(inv_cf),
            "",
            "", fmt(-r["withdrawals"]),
            fmt((r["long_loans"]-r_prev["long_loans"])+(r["short_loans"]-r_prev["short_loans"])),
            fmt(fin_cf),
            "",
            fmt(net_chg), fmt(r_prev["cash"]), fmt(r["cash"]),
        ],
        f"{report_year-1} (ر.س)": [
            "", fmt(r_prev["net_bz"]), fmt(r_prev["depreciation"]),
            "–","–","–","–","–",
            fmt(-r_prev["zakat_due"]),
            fmt(op_cf_p),
            "",
            "", "–", "–", fmt(inv_cf_p),
            "",
            "", fmt(-r_prev["withdrawals"]), "–", fmt(fin_cf_p),
            "",
            fmt(net_chg_p), "–", fmt(r_prev["cash"]),
        ],
    }
    st.table(pd.DataFrame(cf_data))

    # رسم بياني للمقارنة
    fig = go.Figure()
    cats_cf = ["التشغيل", "الاستثمار", "التمويل", "صافي التغير"]
    vals_c  = [op_cf, inv_cf, fin_cf, net_chg]
    vals_p  = [op_cf_p, inv_cf_p, fin_cf_p, net_chg_p]
    colors_c = ["#10B981" if v >= 0 else "#EF4444" for v in vals_c]
    fig.add_trace(go.Bar(name=str(report_year), x=cats_cf, y=vals_c, marker_color=colors_c))
    fig.add_trace(go.Bar(name=str(report_year-1), x=cats_cf, y=vals_p, marker_color="#94A3B8"))
    fig.update_layout(barmode="group", title="مقارنة التدفقات النقدية", height=360, paper_bgcolor="#fff", plot_bgcolor="#F8FAFC")
    st.plotly_chart(fig, use_container_width=True)


def show_equity_changes(r, r_prev, report_year):
    st.markdown(f'<div class="sec-title">📋 قائمة التغيرات في حقوق الملكية | عن السنة المنتهية 31 ديسمبر {report_year}</div>', unsafe_allow_html=True)

    eq_open  = r_prev["capital"] + r_prev["reserves"] + r_prev["retained"] - r_prev["withdrawals"]
    eq_close = r["capital"] + r["reserves"] + r["retained"] - r["withdrawals"]
    eq_change = eq_close - eq_open

    # بطاقات ملخص
    c1, c2, c3 = st.columns(3)
    c1.metric("📂 حقوق الملكية - بداية الفترة", f"{eq_open:,.0f} ر.س")
    c2.metric("📈 صافي التغير", f"{eq_change:,.0f} ر.س", delta="ارتفاع ✅" if eq_change >= 0 else "انخفاض ⚠️", delta_color="normal" if eq_change >= 0 else "inverse")
    c3.metric("📂 حقوق الملكية - نهاية الفترة", f"{eq_close:,.0f} ر.س")

    st.markdown("---")
    headers = ["البيـــان", "رأس المال", "الاحتياطيات", "الأرباح المبقاة", "(ناقص) المسحوبات", "مجموع حقوق الملكية"]
    rows = [
        (f"الرصيد أول الفترة ({report_year-1})",
         r_prev["capital"], r_prev["reserves"], r_prev["retained"], -r_prev["withdrawals"],
         r_prev["capital"]+r_prev["reserves"]+r_prev["retained"]-r_prev["withdrawals"]),
        ("صافي ربح الفترة السابقة", 0, 0, r_prev["net_az"], 0, r_prev["net_az"]),
        (f"الرصيد بداية الفترة الحالية ({report_year})",
         r_prev["capital"], r_prev["reserves"], r_prev["retained"]+r_prev["net_az"], -r_prev["withdrawals"],
         eq_open + r_prev["net_az"]),
        ("صافي ربح الفترة الحالية", 0, 0, r["net_az"], 0, r["net_az"]),
        ("مسحوبات الشركاء / توزيعات", 0, 0, 0, -r["withdrawals"], -r["withdrawals"]),
        ("تغير في الاحتياطيات", 0, r["reserves"]-r_prev["reserves"], 0, 0, r["reserves"]-r_prev["reserves"]),
        (f"الرصيد نهاية الفترة ({report_year})",
         r["capital"], r["reserves"], r["retained"], -r["withdrawals"],
         r["capital"]+r["reserves"]+r["retained"]-r["withdrawals"]),
    ]
    totals_rows = {0, 2, 6}

    table_data = {h: [] for h in headers}
    for i, row in enumerate(rows):
        table_data["البيـــان"].append(("🔷 " if i in totals_rows else "    ") + row[0])
        for h, val in zip(headers[1:], row[1:]):
            table_data[h].append(fmt(val))

    st.table(pd.DataFrame(table_data))

    # رسم بياني للتغيرات
    comp_labels = [headers[1], headers[2], headers[3], headers[4]]
    open_vals  = [r_prev["capital"], r_prev["reserves"], r_prev["retained"], -r_prev["withdrawals"]]
    close_vals = [r["capital"],      r["reserves"],      r["retained"],      -r["withdrawals"]]

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(name=f"بداية {report_year}", x=comp_labels, y=open_vals,  marker_color="#0F2942"))
    fig2.add_trace(go.Bar(name=f"نهاية {report_year}",  x=comp_labels, y=close_vals, marker_color="#10B981"))
    fig2.update_layout(barmode="group", title="مكونات حقوق الملكية - مقارنة", height=360, paper_bgcolor="#fff", plot_bgcolor="#F8FAFC")
    st.plotly_chart(fig2, use_container_width=True)


def show_dashboard(r, r_prev, report_year, company_name):
    """لوحة التحكم التنفيذية - Executive Dashboard"""
    st.markdown(f'<div class="sec-title">🏠 لوحة التحكم التنفيذية | {company_name} | {report_year}</div>', unsafe_allow_html=True)

    # ── حساب المتغيرات ──
    ns   = r["net_sales"];       ns_p  = r_prev["net_sales"]
    gp   = r["gross_profit"];    gp_p  = r_prev["gross_profit"]
    op   = r["op_profit"];       op_p  = r_prev["op_profit"]
    net  = r["net_az"];          net_p = r_prev["net_az"]
    ta   = r["total_assets"];    ta_p  = r_prev["total_assets"]
    eq   = r["equity"];          eq_p  = r_prev["equity"]
    cash = r["cash"];            cash_p= r_prev["cash"]
    zk   = r["zakat_due"]

    def delta_pct(curr, prev):
        if prev == 0: return None
        return (curr - prev) / abs(prev) * 100

    def kpi_card(label, curr, prev, icon, is_pct=False, currency=True):
        d = delta_pct(curr, prev)
        if is_pct:
            val_str = f"{curr:.1%}"
        elif currency:
            val_str = f"{curr:,.0f} ر.س"
        else:
            val_str = f"{curr:,.2f}"
        arrow = ""
        color = "#10B981"
        if d is not None:
            if d >= 0:
                arrow = f'<span style="color:#10B981;font-size:11px;">▲ {d:.1f}%</span>'
            else:
                arrow = f'<span style="color:#EF4444;font-size:11px;">▼ {abs(d):.1f}%</span>'
                color = "#EF4444"
        st.markdown(f"""
        <div style="background:#fff;border:0.5px solid #E2E8F0;border-radius:10px;padding:14px 12px;
                    border-top:3px solid {color};text-align:center;">
          <div style="font-size:1.5rem;margin-bottom:4px;">{icon}</div>
          <div style="font-size:0.72rem;color:#64748B;margin-bottom:4px;">{label}</div>
          <div style="font-size:1.1rem;font-weight:700;color:#0F2942;">{val_str}</div>
          <div style="margin-top:4px;">{arrow}</div>
        </div>""", unsafe_allow_html=True)

    # ── صف KPIs ──
    st.markdown("#### 📌 المؤشرات الرئيسية")
    c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
    with c1: kpi_card("المبيعات الصافية",   ns,   ns_p,  "🏪")
    with c2: kpi_card("مجمل الربح",         gp,   gp_p,  "📈")
    with c3: kpi_card("الربح التشغيلي",     op,   op_p,  "⚙️")
    with c4: kpi_card("صافي الربح",         net,  net_p, "💰")
    with c5: kpi_card("إجمالي الأصول",      ta,   ta_p,  "🏦")
    with c6: kpi_card("حقوق الملكية",       eq,   eq_p,  "🏛️")
    with c7: kpi_card("النقدية",            cash, cash_p,"💵")
    with c8: kpi_card("الزكاة المستحقة",    zk,   0,     "🕌")

    st.markdown("---")

    # ── صف الرسوم البيانية: جزءان ──
    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.markdown("#### 📊 مسار الأداء المالي (مقارنة السنتين)")
        categories = ["المبيعات الصافية", "مجمل الربح", "الربح التشغيلي", "صافي الربح"]
        curr_vals  = [ns, gp, op, net]
        prev_vals  = [ns_p, gp_p, op_p, net_p]
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(name=str(report_year),   x=categories, y=curr_vals,
                                  marker_color="#10B981", text=[f"{v/1e6:.1f}M" if v >= 1e6 else f"{v:,.0f}" for v in curr_vals],
                                  textposition="outside"))
        fig_bar.add_trace(go.Bar(name=str(report_year-1), x=categories, y=prev_vals,
                                  marker_color="#94A3B8", text=[f"{v/1e6:.1f}M" if v >= 1e6 else f"{v:,.0f}" for v in prev_vals],
                                  textposition="outside"))
        fig_bar.update_layout(barmode="group", height=320, paper_bgcolor="#fff",
                               plot_bgcolor="#F8FAFC", margin=dict(t=20,b=20,l=10,r=10),
                               legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_right:
        st.markdown("#### 🍩 هيكل التكاليف")
        cost_labels = ["تكلفة المبيعات", "الرواتب", "الإيجارات", "التسويق", "الاستهلاك", "أخرى"]
        cost_vals   = [r["cogs"], r["salaries"], r["rent"], r["marketing"], r["depreciation"], r["misc_exp"]]
        cost_vals   = [max(v, 0) for v in cost_vals]
        colors_pie  = ["#0F2942","#10B981","#3B82F6","#F59E0B","#8B5CF6","#94A3B8"]
        fig_pie = go.Figure(go.Pie(labels=cost_labels, values=cost_vals, hole=0.5,
                                    marker_colors=colors_pie, textinfo="percent+label",
                                    textfont_size=10))
        fig_pie.update_layout(height=320, paper_bgcolor="#fff",
                               margin=dict(t=20,b=20,l=10,r=10),
                               showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True)

    # ── صف ثانٍ من الرسوم ──
    col2_left, col2_right = st.columns([2, 3])

    with col2_left:
        st.markdown("#### 🏗️ هيكل الأصول")
        asset_labels = ["نقدية", "مدينون", "مخزون", "أرصدة أخرى", "أصول ثابتة", "استثمارات"]
        asset_vals   = [r["cash"], r["receivables"], r["inventory"], r["prepaid"], r["fixed_assets"], r["investments"]]
        asset_vals   = [max(v, 0) for v in asset_vals]
        colors_assets= ["#10B981","#3B82F6","#F59E0B","#A78BFA","#0F2942","#64748B"]
        fig_assets = go.Figure(go.Pie(labels=asset_labels, values=asset_vals, hole=0.4,
                                       marker_colors=colors_assets, textinfo="percent+label",
                                       textfont_size=10))
        fig_assets.update_layout(height=300, paper_bgcolor="#fff",
                                  margin=dict(t=20,b=20,l=10,r=10), showlegend=False)
        st.plotly_chart(fig_assets, use_container_width=True)

    with col2_right:
        st.markdown("#### 📉 تحليل الهوامش (%)")
        margin_cats  = ["هامش مجمل الربح", "هامش الربح التشغيلي", "هامش الربح الصافي"]
        gpm_c  = gp / ns  if ns else 0
        opm_c  = op / ns  if ns else 0
        npm_c  = net / ns if ns else 0
        gpm_p  = gp_p / ns_p  if ns_p else 0
        opm_p  = op_p / ns_p  if ns_p else 0
        npm_p  = net_p / ns_p if ns_p else 0
        margin_curr = [gpm_c*100, opm_c*100, npm_c*100]
        margin_prev = [gpm_p*100, opm_p*100, npm_p*100]
        fig_margin = go.Figure()
        fig_margin.add_trace(go.Bar(name=str(report_year),   x=margin_cats, y=margin_curr,
                                     marker_color="#10B981",
                                     text=[f"{v:.1f}%" for v in margin_curr], textposition="outside"))
        fig_margin.add_trace(go.Bar(name=str(report_year-1), x=margin_cats, y=margin_prev,
                                     marker_color="#94A3B8",
                                     text=[f"{v:.1f}%" for v in margin_prev], textposition="outside"))
        fig_margin.update_layout(barmode="group", height=300, paper_bgcolor="#fff",
                                  plot_bgcolor="#F8FAFC", yaxis_ticksuffix="%",
                                  margin=dict(t=20,b=20,l=10,r=10),
                                  legend=dict(orientation="h", y=-0.25))
        st.plotly_chart(fig_margin, use_container_width=True)

    # ── ملخص نصي تنفيذي ──
    st.markdown("---")
    st.markdown("#### 📝 الملخص التنفيذي التلقائي")
    ns_chg  = delta_pct(ns,  ns_p)
    net_chg = delta_pct(net, net_p)
    gpm_lbl = f"{gpm_c:.1%}"
    npm_lbl = f"{npm_c:.1%}"
    trend_ns  = f"ارتفعت بنسبة {ns_chg:.1f}%" if ns_chg and ns_chg > 0 else (f"انخفضت بنسبة {abs(ns_chg):.1f}%" if ns_chg else "لم تتغير")
    trend_net = f"ارتفع بنسبة {net_chg:.1f}%" if net_chg and net_chg > 0 else (f"انخفض بنسبة {abs(net_chg):.1f}%" if net_chg else "لم يتغير")
    health = "🟢 ممتاز" if npm_c > 0.15 else ("🟡 جيد" if npm_c > 0.05 else "🔴 يحتاج مراجعة")
    st.info(f"""
**{company_name}** — السنة المالية {report_year}

- المبيعات الصافية **{trend_ns}** مقارنةً بالعام السابق، لتبلغ **{ns:,.0f} ر.س**
- هامش مجمل الربح **{gpm_lbl}** وهامش الربح الصافي **{npm_lbl}**
- صافي الربح بعد الزكاة **{trend_net}** ليبلغ **{net:,.0f} ر.س**
- الزكاة الشرعية المستحقة: **{zk:,.0f} ر.س**
- **التقييم العام للأداء المالي: {health}**
    """)


def show_financial_ratios(r, r_prev, report_year):
    """تحليل النسب المالية الموسّع"""
    st.markdown(f'<div class="sec-title">📐 تحليل النسب المالية الموسّع | {report_year}</div>', unsafe_allow_html=True)

    ns   = r["net_sales"];      ns_p  = r_prev["net_sales"]
    gp   = r["gross_profit"];   gp_p  = r_prev["gross_profit"]
    op   = r["op_profit"];      op_p  = r_prev["op_profit"]
    net  = r["net_az"];         net_p = r_prev["net_az"]
    ta   = r["total_assets"];   ta_p  = r_prev["total_assets"]
    eq   = r["equity"];         eq_p  = r_prev["equity"]
    tl   = r["total_cl"] + r["total_ncl"] + r["zakat_due"]
    tl_p = r_prev["total_cl"] + r_prev["total_ncl"] + r_prev["zakat_due"]
    ca   = r["total_current"];  ca_p  = r_prev["total_current"]
    cl   = r["total_cl"];       cl_p  = r_prev["total_cl"]
    cash = r["cash"]
    rec  = r["receivables"];    rec_p = r_prev["receivables"]
    inv  = r["inventory"];      inv_p = r_prev["inventory"]
    cogs = r["cogs"]
    pay  = r["payables"];       pay_p = r_prev["payables"]

    def safe(num, den, fmt_type="x"):
        if den == 0: return "–"
        v = num / den
        if fmt_type == "x": return f"{v:.2f}x"
        if fmt_type == "%": return f"{v:.1%}"
        if fmt_type == "d": return f"{v:.0f} يوم"
        return f"{v:.2f}"

    def score(val, thresholds, labels):
        """إعطاء تقييم نصي بناءً على العتبات"""
        if val is None: return "–"
        for t, lbl in zip(thresholds, labels):
            if val >= t: return lbl
        return labels[-1]

    # ── 1. نسب السيولة ──
    st.markdown("### 💧 أولاً: نسب السيولة")
    cr_c   = ca / cl   if cl   else None
    cr_p   = ca_p / cl_p if cl_p else None
    qr_c   = (ca - inv) / cl       if cl   else None
    qr_p   = (ca_p - inv_p) / cl_p if cl_p else None
    casr_c = cash / cl  if cl else None

    liq_data = {
        "النسبة": ["نسبة التداول (Current Ratio)", "النسبة السريعة (Quick Ratio)", "نسبة النقدية (Cash Ratio)"],
        "المعادلة": ["الأصول المتداولة ÷ الالتزامات المتداولة", "(الأصول المتداولة - المخزون) ÷ الالتزامات المتداولة", "النقدية ÷ الالتزامات المتداولة"],
        f"{report_year}": [
            safe(ca, cl), safe(ca - inv, cl), safe(cash, cl)
        ],
        f"{report_year-1}": [
            safe(ca_p, cl_p), safe(ca_p - inv_p, cl_p), "–"
        ],
        "المعيار المثالي": ["≥ 2.0x", "≥ 1.0x", "≥ 0.5x"],
        "التقييم": [
            score(cr_c,   [2.0, 1.5, 1.0], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
            score(qr_c,   [1.5, 1.0, 0.5], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
            score(casr_c, [1.0, 0.5, 0.2], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
        ]
    }
    st.table(pd.DataFrame(liq_data))

    # ── 2. نسب الربحية ──
    st.markdown("### 💰 ثانياً: نسب الربحية")
    gpm_c = gp / ns   if ns else None
    opm_c = op / ns   if ns else None
    npm_c = net / ns  if ns else None
    roa_c = net / ta  if ta else None
    roe_c = net / eq  if eq else None
    gpm_p = gp_p / ns_p  if ns_p else None
    opm_p = op_p / ns_p  if ns_p else None
    npm_p = net_p / ns_p if ns_p else None
    roa_p = net_p / ta_p if ta_p else None
    roe_p = net_p / eq_p if eq_p else None

    prof_data = {
        "النسبة": [
            "هامش مجمل الربح (GPM)", "هامش الربح التشغيلي (OPM)",
            "هامش الربح الصافي (NPM)", "العائد على الأصول (ROA)", "العائد على حقوق الملكية (ROE)"
        ],
        "المعادلة": [
            "مجمل الربح ÷ صافي المبيعات", "الربح التشغيلي ÷ صافي المبيعات",
            "صافي الربح ÷ صافي المبيعات", "صافي الربح ÷ إجمالي الأصول", "صافي الربح ÷ حقوق الملكية"
        ],
        f"{report_year}": [safe(gp,ns,"%"), safe(op,ns,"%"), safe(net,ns,"%"), safe(net,ta,"%"), safe(net,eq,"%")],
        f"{report_year-1}": [safe(gp_p,ns_p,"%"), safe(op_p,ns_p,"%"), safe(net_p,ns_p,"%"), safe(net_p,ta_p,"%"), safe(net_p,eq_p,"%")],
        "المعيار المثالي": ["≥ 30%", "≥ 15%", "≥ 10%", "≥ 5%", "≥ 15%"],
        "التقييم": [
            score(gpm_c, [0.30, 0.20, 0.10], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
            score(opm_c, [0.15, 0.08, 0.03], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
            score(npm_c, [0.10, 0.05, 0.02], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
            score(roa_c, [0.10, 0.05, 0.02], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
            score(roe_c, [0.20, 0.12, 0.05], ["🟢 ممتاز", "🟡 جيد", "🟠 مقبول", "🔴 ضعيف"]),
        ]
    }
    st.table(pd.DataFrame(prof_data))

    # ── 3. نسب النشاط / الكفاءة ──
    st.markdown("### ⚙️ ثالثاً: نسب النشاط والكفاءة التشغيلية")
    avg_rec = (rec + rec_p) / 2  if rec_p else rec
    avg_inv = (inv + inv_p) / 2  if inv_p else inv
    avg_pay = (pay + pay_p) / 2  if pay_p else pay
    avg_ta  = (ta  + ta_p)  / 2  if ta_p  else ta

    dso = (avg_rec / ns * 365)   if ns   else None   # أيام التحصيل
    dio = (avg_inv / cogs * 365) if cogs else None   # أيام المخزون
    dpo = (avg_pay / cogs * 365) if cogs else None   # أيام الدفع
    ccc = (dso or 0) + (dio or 0) - (dpo or 0)       # دورة التحويل النقدي
    ato = ns / avg_ta            if avg_ta else None  # معدل دوران الأصول

    act_data = {
        "النسبة": [
            "معدل دوران المدينين", "أيام التحصيل (DSO)",
            "معدل دوران المخزون", "أيام المخزون (DIO)",
            "أيام الدفع للدائنين (DPO)", "دورة التحويل النقدي (CCC)",
            "معدل دوران الأصول (ATO)"
        ],
        "المعادلة": [
            "صافي المبيعات ÷ متوسط المدينين", "متوسط المدينين ÷ المبيعات × 365",
            "تكلفة المبيعات ÷ متوسط المخزون", "متوسط المخزون ÷ تكلفة المبيعات × 365",
            "متوسط الموردين ÷ تكلفة المبيعات × 365", "DSO + DIO - DPO",
            "صافي المبيعات ÷ متوسط الأصول"
        ],
        f"{report_year}": [
            safe(ns, avg_rec),
            f"{dso:.0f} يوم" if dso else "–",
            safe(cogs, avg_inv),
            f"{dio:.0f} يوم" if dio else "–",
            f"{dpo:.0f} يوم" if dpo else "–",
            f"{ccc:.0f} يوم" if ccc else "–",
            safe(ns, avg_ta),
        ],
        "التقييم": [
            score(ns/avg_rec if avg_rec else None, [8,5,3], ["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 بطيء"]),
            score(1/dso*365 if dso else None,      [8,5,3], ["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 بطيء"]),
            score(cogs/avg_inv if avg_inv else None,[8,5,3],["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 بطيء"]),
            "–","–",
            score(-ccc if ccc else None, [0,-30,-60], ["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 ضعيف"]),
            score(ato, [1.5,1.0,0.5], ["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 ضعيف"]),
        ]
    }
    st.table(pd.DataFrame(act_data))

    # ── 4. نسب الرفع المالي ──
    st.markdown("### 🏗️ رابعاً: نسب الرفع المالي والملاءة")
    der_c  = tl  / eq   if eq   else None
    der_p  = tl_p / eq_p if eq_p else None
    dar_c  = tl  / ta   if ta   else None
    dar_p  = tl_p / ta_p if ta_p else None
    eq_r_c = eq  / ta   if ta   else None
    eq_r_p = eq_p / ta_p if ta_p else None

    lev_data = {
        "النسبة": [
            "نسبة الدين إلى حقوق الملكية (D/E)",
            "نسبة الدين إلى الأصول (D/A)",
            "نسبة حقوق الملكية إلى الأصول",
            "مضاعف حقوق الملكية (Equity Multiplier)"
        ],
        "المعادلة": [
            "إجمالي الالتزامات ÷ حقوق الملكية",
            "إجمالي الالتزامات ÷ إجمالي الأصول",
            "حقوق الملكية ÷ إجمالي الأصول",
            "إجمالي الأصول ÷ حقوق الملكية"
        ],
        f"{report_year}": [safe(tl,eq), safe(tl,ta,"%"), safe(eq,ta,"%"), safe(ta,eq)],
        f"{report_year-1}": [safe(tl_p,eq_p), safe(tl_p,ta_p,"%"), safe(eq_p,ta_p,"%"), safe(ta_p,eq_p)],
        "المعيار المثالي": ["≤ 1.0x", "≤ 50%", "≥ 50%", "≤ 2.0x"],
        "التقييم": [
            score(1/der_c if der_c else None,  [1.0, 0.5, 0.2], ["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 مرتفع"]),
            score(1-dar_c  if dar_c else None,  [0.5, 0.3, 0.1], ["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 مرتفع"]),
            score(eq_r_c,   [0.5, 0.35, 0.2],  ["🟢 ممتاز","🟡 جيد","🟠 مقبول","🔴 ضعيف"]),
            "–"
        ]
    }
    st.table(pd.DataFrame(lev_data))

    # ── رسم بياني راداري للنسب ──
    st.markdown("---")
    st.markdown("#### 🕸️ الرسم الراداري للأداء المالي الشامل")
    radar_cats   = ["السيولة", "الربحية", "الكفاءة التشغيلية", "الرفع المالي", "العائد على الأصول"]
    def norm(val, lo, hi):
        if val is None: return 0
        return max(0, min(100, (val - lo) / (hi - lo) * 100))
    radar_vals_c = [
        norm(cr_c,  0, 3),
        norm(npm_c, 0, 0.25),
        norm(ato,   0, 2),
        norm(1 - (dar_c or 0), 0, 1),
        norm(roa_c, 0, 0.20),
    ]
    radar_vals_p = [
        norm(cr_p,  0, 3),
        norm(npm_p, 0, 0.25),
        norm(ns_p / ((ta_p + ta_p)/2) if ta_p else 0, 0, 2),
        norm(1 - (dar_p or 0), 0, 1),
        norm(roa_p, 0, 0.20),
    ]
    fig_radar = go.Figure()
    fig_radar.add_trace(go.Scatterpolar(r=radar_vals_c + [radar_vals_c[0]], theta=radar_cats + [radar_cats[0]],
                                         fill="toself", name=str(report_year),
                                         line_color="#10B981", fillcolor="rgba(16,185,129,0.15)"))
    fig_radar.add_trace(go.Scatterpolar(r=radar_vals_p + [radar_vals_p[0]], theta=radar_cats + [radar_cats[0]],
                                         fill="toself", name=str(report_year-1),
                                         line_color="#94A3B8", fillcolor="rgba(148,163,184,0.1)"))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        showlegend=True, height=450, paper_bgcolor="#fff",
        legend=dict(orientation="h", y=-0.1)
    )
    st.plotly_chart(fig_radar, use_container_width=True)

    # ── تقييم شامل ──
    st.markdown("---")
    st.markdown("#### 🏅 التقييم الإجمالي")
    scores = [
        ("💧 السيولة",          cr_c,   [(2.0,"🟢 ممتاز"),(1.5,"🟡 جيد"),(1.0,"🟠 مقبول"),(0,"🔴 ضعيف")]),
        ("💰 الربحية",          npm_c,  [(0.10,"🟢 ممتاز"),(0.05,"🟡 جيد"),(0.02,"🟠 مقبول"),(0,"🔴 ضعيف")]),
        ("⚙️ الكفاءة",         ato,    [(1.5,"🟢 ممتاز"),(1.0,"🟡 جيد"),(0.5,"🟠 مقبول"),(0,"🔴 ضعيف")]),
        ("🏗️ الرفع المالي",    1-(dar_c or 1), [(0.5,"🟢 ممتاز"),(0.3,"🟡 جيد"),(0.2,"🟠 مقبول"),(0,"🔴 مرتفع")]),
        ("📈 العائد على الأصول",roa_c, [(0.10,"🟢 ممتاز"),(0.05,"🟡 جيد"),(0.02,"🟠 مقبول"),(0,"🔴 ضعيف")]),
    ]
    sc_cols = st.columns(5)
    for col, (lbl, val, thresholds) in zip(sc_cols, scores):
        rating = "–"
        if val is not None:
            for t, r_lbl in thresholds:
                if val >= t: rating = r_lbl; break
        with col:
            st.markdown(f"""
            <div style="background:#fff;border:0.5px solid #E2E8F0;border-radius:10px;
                        padding:14px 10px;text-align:center;">
              <div style="font-size:0.75rem;color:#64748B;margin-bottom:6px;">{lbl}</div>
              <div style="font-size:1rem;font-weight:700;">{rating}</div>
            </div>""", unsafe_allow_html=True)


# ── تصدير Excel الشامل (7 شيتات مالية متكاملة) ──
def export_excel(r, r_prev, df, company_name, company_type_label, year):
    wb = Workbook()
    wb.remove(wb.active)
    DARK, W = "1A3A5C", "FFFFFF"
    bd = Border(left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"), top=Side(style="thin", color="CCCCCC"), bottom=Side(style="thin", color="CCCCCC"))
    bd_thick = Border(left=Side(style="medium", color=DARK), right=Side(style="medium", color=DARK), top=Side(style="medium", color=DARK), bottom=Side(style="medium", color=DARK))

    def f(sz=11, bold=False, color="000000"): return Font(name="Arial", size=sz, bold=bold, color=color)
    def al(h="center", v="center"): return Alignment(horizontal=h, vertical=v, wrap_text=True)
    def fill(clr): return PatternFill("solid", fgColor=clr)
    def header_row_style(ws, row, cols, text_list, bg=DARK):
        for j, txt in enumerate(text_list, 1): c = ws.cell(row, j, txt); c.font = f(11, True, W); c.fill = fill(bg); c.alignment = al(); c.border = bd
    def total_style(ws, row, ncols, bg="E8F0FE"):
        for j in range(1, ncols+1): ws.cell(row, j).fill = fill(bg); ws.cell(row, j).font = f(11, True); ws.cell(row, j).border = bd
    def section_style(ws, row, ncols):
        for j in range(1, ncols+1): ws.cell(row, j).fill = fill("DBEAFE"); ws.cell(row, j).font = f(11, True, "1A3A5C"); ws.cell(row, j).border = bd
    def num_cell(ws, row, col, val, bold=False):
        c = ws.cell(row, col, val if val != "" else None); c.number_format = '#,##0.00;(#,##0.00);"-"'; c.border = bd
        if bold: c.font = f(11, True)
    def company_header(ws, title, period, cols="A:E"):
        ws.merge_cells(f'A1:{cols[-1]}1'); ws['A1'] = company_name; ws['A1'].font = f(14, True, DARK); ws['A1'].alignment = al()
        ws.merge_cells(f'A2:{cols[-1]}2'); ws['A2'] = f'({company_type_label})'; ws['A2'].font = f(11, False); ws['A2'].alignment = al()
        ws.merge_cells(f'A3:{cols[-1]}3'); ws['A3'] = title; ws['A3'].font = f(13, True); ws['A3'].alignment = al()
        ws.merge_cells(f'A4:{cols[-1]}4'); ws['A4'] = period; ws['A4'].font = f(11, False); ws['A4'].alignment = al()
        ws.merge_cells(f'A5:{cols[-1]}5')
        ws['A5'] = f"تطوير وبرمجة: {OWNER_NAME_AR}  |  {OWNER_NAME_EN}  |  {OWNER_ROLE_EN}  |  {APP_VERSION}"
        ws['A5'].font = Font(name="Arial", size=9, italic=True, color="10B981")
        ws['A5'].alignment = al(); ws['A5'].fill = PatternFill("solid", fgColor="0A1E30")

    # 1. المركز المالي
    ws_bs = wb.create_sheet("قائمة المركز المالي")
    ws_bs.sheet_view.rightToLeft = True
    for col, w in zip(['A','B','C','D','E'], [36, 10, 22, 22, 28]): ws_bs.column_dimensions[col].width = w
    company_header(ws_bs, "قائمة المركز المالي", f"كما في 31 ديسمبر {year}")
    header_row_style(ws_bs, 6, 5, ["البيـــان", "إيضاح", f"31 ديسمبر {year}", "التعديل", "إعادة تصنيف للزكاة"])
    bs_rows = [
        ("الأصول", None, True, "section"), ("الأصول غير المتداولة", None, True, "sub"),
        ("الممتلكات والآلات والمعدات (صافي)", r["fixed_assets"], False, ""), ("الاستثمارات في شركات تابعة", r["investments"], False, ""),
        ("مجموع الأصول غير المتداولة", r["total_nca"], True, "total"), ("الأصول المتداولة", None, True, "sub"),
        ("نقد وما في حكمه", r["cash"], False, ""), ("مدينون تجاريون (صافي)", r["receivables"], False, ""),
        ("مخزون بضاعة", r["inventory"], False, ""), ("أطراف ذات علاقة (مدين)", r["related_rec"], False, ""),
        ("أرصدة مدينة أخرى (سلف، ضريبة)", r["prepaid"], False, ""), ("مجموع الأصول المتداولة", r["total_current"], True, "total"),
        ("مجموع الأصول", r["total_assets"], True, "grand"), ("", None, False, ""),
        ("حقوق الملكية والالتزامات", None, True, "section"), ("حقوق الملكية", None, True, "sub"),
        ("رأس المال", r["capital"], False, ""), ("احتياطي نظامي", r["reserves"], False, ""),
        ("الأرباح المبقاة", r["retained"], False, ""), ("(ناقص) مسحوبات الشركاء", -r["withdrawals"], False, ""),
        ("صافي ربح الفترة", r["net_az"], False, ""), ("مجموع حقوق الملكية", r["equity"], True, "total"),
        ("الالتزامات غير المتداولة", None, True, "sub"), ("مخصص مكافأة نهاية الخدمة", r["eosb"], False, ""),
        ("قروض طويلة الأجل", r["long_loans"], False, ""), ("مجموع الالتزامات غير المتداولة", r["total_ncl"], True, "total"),
        ("الالتزامات المتداولة", None, True, "sub"), ("دائنون تجاريون (موردون)", r["payables"], False, ""),
        ("مصاريف مستحقة وأرصدة دائنة أخرى", r["accruals"], False, ""), ("أطراف ذات علاقة (دائن)", r["related_pay"], False, ""),
        ("تسهيلات ائتمانية قصيرة الأجل", r["short_loans"], False, ""), ("مخصص الزكاة الشرعية", r["zakat_due"], False, ""),
        ("مجموع الالتزامات المتداولة", r["total_cl"] + r["zakat_due"], True, "total"), ("مجموع حقوق الملكية والالتزامات", r["total_eq_lb"], True, "grand")
    ]
    ri = 7
    for lbl, val, bold, style in bs_rows:
        c_lbl = ws_bs.cell(ri, 1, lbl); c_lbl.border = bd
        ws_bs.cell(ri, 2).border = bd; ws_bs.cell(ri, 4).border = bd; ws_bs.cell(ri, 5).border = bd
        if val is not None and val != 0: num_cell(ws_bs, ri, 3, val, bold)
        else: ws_bs.cell(ri, 3).border = bd
        if style == "section": section_style(ws_bs, ri, 5)
        elif style == "sub":
            for j in range(1,6): ws_bs.cell(ri,j).fill = fill("EFF6FF"); ws_bs.cell(ri,j).font = f(11,True,"1E40AF")
        elif style == "total": total_style(ws_bs, ri, 5)
        elif style == "grand":
            for j in range(1,6): ws_bs.cell(ri,j).fill = fill(DARK); ws_bs.cell(ri,j).font = f(11,True,W); ws_bs.cell(ri,j).border = bd_thick
        else: c_lbl.font = f(11, bold)
        ri += 1

    # 2. قائمة الدخل
    ws_inc = wb.create_sheet("قائمة الدخل الشامل")
    ws_inc.sheet_view.rightToLeft = True
    for col, w in zip(['A','B','C','D'], [36, 10, 25, 25]): ws_inc.column_dimensions[col].width = w
    company_header(ws_inc, "قائمة الدخل الشامل", f"عن السنة المنتهية في 31 ديسمبر {year}", cols="A:D")
    header_row_style(ws_inc, 6, 4, ["البيـــان", "إيضاح", f"سنة {year}", f"سنة {year-1}"])
    inc_rows = [
        ("إجمالي الإيرادات والمبيعات", r["gross_sales"], r_prev["gross_sales"], False, ""),
        ("يخصم: مردودات ومسموحات المبيعات", -r["returns"], -r_prev["returns"], False, ""),
        ("صافي إيرادات المبيعات", r["net_sales"], r_prev["net_sales"], True, "total"),
        ("يخصم: تكلفة المبيعات (النشاط)", -r["cogs"], -r_prev["cogs"], False, ""),
        ("إجمالي الفائض / الربح (Gross Profit)", r["gross_profit"], r_prev["gross_profit"], True, "total"),
        ("", None, None, False, ""), ("المصاريف التشغيلية والإدارية", None, None, True, "section"),
        ("الرواتب والأجور البديلة للموظفين", -r["salaries"], -r_prev["salaries"], False, ""),
        ("مصاريف الإيجارات العقارية والمستودعات", -r["rent"], -r_prev["rent"], False, ""),
        ("الحصة في التأمينات الاجتماعية (GOSI)", -r["gosi"], -r_prev["gosi"], False, ""),
        ("مصاريف التسويق والدعاية والإعلان", -r["marketing"], -r_prev["marketing"], False, ""),
        ("مصاريف شحن البضائع والنقل والتوزيع", -r["transport"], -r_prev["transport"], False, ""),
        ("مصاريف المنافع والخدمات والمرافق", -r["utilities"], -r_prev["utilities"], False, ""),
        ("الاستهلاك والإهلاك للأصول الثابتة", -r["depreciation"], -r_prev["depreciation"], False, ""),
        ("مصاريف عمومية وإدارية أخرى متنوعة", -r["misc_exp"], -r_prev["misc_exp"], False, ""),
        ("إجمالي المصاريف التشغيلية", -r["total_opex"], -r_prev["total_opex"], True, "total"),
        ("الربح التشغيلي للفترة (EBIT)", r["op_profit"], r_prev["op_profit"], True, "total"),
        ("", None, None, False, ""), ("إيرادات مكاسب أخرى", r["other_inc"], r_prev["other_inc"], False, ""),
        ("تكاليف التمويل والفوائد البنكية", -r["finance_cost"], -r_prev["finance_cost"], False, ""),
        ("الربح قبل احتساب الزكاة الشرعية", r["net_bz"], r_prev["net_bz"], True, "total"),
        ("الزكاة الشرعية المستحقة لهيئة الزكاة", -r["zakat_due"], -r_prev["zakat_due"], False, ""),
        ("صافي ربح الفترة بعد التعديل والزكاة", r["net_az"], r_prev["net_az"], True, "grand"),
    ]
    ri = 7
    for lbl, v_c, v_p, bold, style in inc_rows:
        ws_inc.cell(ri, 1, lbl).border = bd; ws_inc.cell(ri, 2).border = bd
        if v_c is not None: num_cell(ws_inc, ri, 3, v_c, bold)
        else: ws_inc.cell(ri, 3).border = bd
        if v_p is not None: num_cell(ws_inc, ri, 4, v_p, bold)
        else: ws_inc.cell(ri, 4).border = bd
        if style == "section": section_style(ws_inc, ri, 4)
        elif style == "total": total_style(ws_inc, ri, 4)
        elif style == "grand":
            for j in range(1,5): ws_inc.cell(ri,j).fill = fill(DARK); ws_inc.cell(ri,j).font = f(11,True,W); ws_inc.cell(ri,j).border = bd_thick
        else: ws_inc.cell(ri,1).font = f(11, bold)
        ri += 1

    # 3. قائمة التغيرات في حقوق الملكية
    ws_eq = wb.create_sheet("قائمة التغيرات في حقوق الملكية")
    ws_eq.sheet_view.rightToLeft = True
    for col, w in zip(['A','B','C','D','E','F'], [30, 20, 20, 20, 20, 20]): ws_eq.column_dimensions[col].width = w
    company_header(ws_eq, "قائمة التغيرات في حقوق الملكية", f"عن السنة المنتهية في 31 ديسمبر {year}", cols="A:F")
    header_row_style(ws_eq, 6, 6, ["البيـــان", "رأس المال", "الاحتياطيات", "الأرباح المبقاة", "(ناقص) المسحوبات", "مجموع حقوق الملكية"])
    ws_eq.cell(7, 1, f"الرصيد أول الفترة {year-1}").font = f(11, True)
    for j in range(1,7): ws_eq.cell(7,j).border = bd

    for ri, (lbl, cap, res, ret, wd) in enumerate([
        (f"الرصيد أول الفترة {year}", r_prev["capital"], r_prev["reserves"], r_prev["retained"], r_prev["withdrawals"]),
        ("إضافات خلال الفترة", 0, 0, 0, 0), ("صافي ربح الفترة", 0, 0, r["net_az"], 0),
        ("مسحوبات الشركاء والملاك", 0, 0, 0, -r["withdrawals"]),
        (f"الرصيد نهاية الفترة {year}", r["capital"], r["reserves"], r["retained"], -r["withdrawals"]),
    ], start=8):
        ws_eq.cell(ri, 1, lbl).font = f(11, ri in [8, 12])
        for j in range(1,7): ws_eq.cell(ri,j).border = bd
        vals = [cap, res, ret, wd, cap+res+ret+wd]
        for j, v in enumerate(vals, 2): num_cell(ws_eq, ri, j, v)
        if ri in [8, 12]: total_style(ws_eq, ri, 6)

    # 4. التدفقات النقدية
    ws_cf = wb.create_sheet("قائمة التدفقات النقدية")
    ws_cf.sheet_view.rightToLeft = True
    for col, w in zip(['A','B','C','D'], [36, 10, 25, 25]): ws_cf.column_dimensions[col].width = w
    company_header(ws_cf, "قائمة التدفقات النقدية", f"عن السنة المنتهية في 31 ديسمبر {year}", cols="A:D")
    header_row_style(ws_cf, 6, 4, ["البيـــان", "إيضاح", f"سنة {year}", f"سنة {year-1}"])
    cf_rows = [
        ("الأنشطة التشغيلية", None, None, True, "section"),
        ("صافي الربح قبل الزكاة", r["net_bz"], r_prev["net_bz"], False, ""),
        ("تعديلات بنود غير نقدية:", None, None, True, "sub"),
        ("الاستهلاك والإهلاك للأصول", r["depreciation"], r_prev["depreciation"], False, ""),
        ("تغيرات في رأس المال العامل:", None, None, True, "sub"),
        ("(زيادة)/نقص في المدينين التجاريين", -r["receivables"]+r_prev["receivables"], 0, False, ""),
        ("(زيادة)/نقص في المخزون السلعي", -r["inventory"]+r_prev["inventory"], 0, False, ""),
        ("(زيادة)/نقص في الأرصدة المدينة الأخرى", -r["prepaid"]+r_prev["prepaid"], 0, False, ""),
        ("زيادة/(نقص) في الدائنين والموردين", r["payables"]-r_prev["payables"], 0, False, ""),
        ("زيادة/(نقص) في المصاريف المستحقة", r["accruals"]-r_prev["accruals"], 0, False, ""),
        ("الزكاة المدفوعة فعلياً", -r["zakat_due"], -r_prev["zakat_due"], False, ""),
        ("صافي النقد من الأنشطة التشغيلية", r["net_bz"] + r["depreciation"], r_prev["net_bz"] + r_prev["depreciation"], True, "total"),
        ("", None, None, False, ""), ("الأنشطة الاستثمارية", None, None, True, "section"),
        ("شراء ممتلكات وآلات ومعدات جديدة", 0, 0, False, ""),
        ("صافي النقد من الأنشطة الاستثمارية", 0, 0, True, "total"),
        ("", None, None, False, ""), ("الأنشطة التمويلية", None, None, True, "section"),
        ("مسحوبات الشركاء النقدية", -r["withdrawals"], -r_prev["withdrawals"], False, ""),
        ("صافي النقد من الأنشطة التمويلية", -r["withdrawals"], -r_prev["withdrawals"], True, "total"),
        ("", None, None, False, ""),
        ("صافي التغير في النقد وما في حكمه", r["cash"] - r_prev["cash"], 0, True, "total"),
        ("النقد وما في حكمه - بداية الفترة", r_prev["cash"], 0, False, ""),
        ("النقد وما في حكمه - نهاية الفترة", r["cash"], r_prev["cash"], True, "grand"),
    ]
    ri = 7
    for lbl, v_c, v_p, bold, style in cf_rows:
        ws_cf.cell(ri, 1, lbl).border = bd; ws_cf.cell(ri, 2).border = bd
        if v_c is not None: num_cell(ws_cf, ri, 3, v_c, bold)
        else: ws_cf.cell(ri, 3).border = bd
        if v_p is not None: num_cell(ws_cf, ri, 4, v_p, bold)
        else: ws_cf.cell(ri, 4).border = bd
        if style == "section": section_style(ws_cf, ri, 4)
        elif style == "sub":
            for j in range(1,5): ws_cf.cell(ri,j).fill = fill("F0F9FF"); ws_cf.cell(ri,j).font = f(11, True, "374151")
        elif style == "total": total_style(ws_cf, ri, 4)
        elif style == "grand":
            for j in range(1,5): ws_cf.cell(ri,j).fill = fill(DARK); ws_cf.cell(ri,j).font = f(11,True,W); ws_cf.cell(ri,j).border = bd_thick
        else: ws_cf.cell(ri,1).font = f(11, bold)
        ri += 1

    # 5. مخصص الزكاة الشرعية (ZATCA)
    ws_z = wb.create_sheet("مخصص الزكاة الشرعية")
    ws_z.sheet_view.rightToLeft = True
    for col, w in zip(['A','B','C'], [40, 30, 25]): ws_z.column_dimensions[col].width = w
    company_header(ws_z, "إحتساب مخصص الزكاة الشرعية", f"كما في 31 ديسمبر {year}", cols="A:C")
    ws_z.merge_cells('A6:C6'); ws_z['A6'] = "⚖️ مطابق للائحة التنفيذية لجباية الزكاة - إصدار 1445هـ"
    ws_z['A6'].font = f(11, True, "198754"); ws_z['A6'].alignment = al()
    header_row_style(ws_z, 7, 3, ["البيـــان", "المرجع / الملاحظة", f"كما في 31 ديسمبر {year}"])
    z_rows = [
        ("أولاً: احتساب صافي الربح المعدل", None, None, "section"),
        ("الربح قبل الزكاة (من قائمة الدخل)", "قائمة الدخل", r["net_bz"], ""),
        ("إضافة: التعديلات الزكوية (مادة 4)", "لائحة جباية الزكاة", r["net_adj"] - r["net_bz"], ""),
        ("صافي الربح المعدل (رقم 1)", None, r["net_adj"], "total"),
        ("", None, None, ""), ("ثانياً: الإضافات لوعاء الزكاة", None, None, "section"),
        ("أ - حقوق الملكية والخصوم طويلة المدى", None, r["additions"], ""),
        ("إجمالي إضافات الوعاء الزكوي (رقم 2)", None, r["additions"], "total"),
        ("", None, None, ""), ("ثالثاً: حسميات الوعاء الزكوي", None, None, "section"),
        ("جـ - الأصول غير المتداولة (تحسم)", None, -r["deductions"], ""),
        ("إجمالي الحسميات (رقم 3)", None, -r["deductions"], "total"),
        ("", None, None, ""), ("رابعاً: الوعاء الزكوي", None, None, "section"),
        ("صافي الإضافات والحسميات (وعاء محسوب)", "رقم 2 - رقم 3", r["zakat_base"], ""),
        ("الحد الأدنى للوعاء (الأصول المتداولة المعدلة)", "المادة 8", r["min_base"], ""),
        ("الحد الأعلى للوعاء (حقوق الملكية المعدلة)", "المادة 9", r["max_base"], ""),
        ("وعاء الزكاة النهائي المعتمد", None, r["zakat_base"], "grand"),
        ("", None, None, ""), ("خامساً: الزكاة المستحقة", None, None, "section"),
        ("نسبة الزكاة الشرعية المعتمدة", "2.5% × 365/354 يوم هجري", None, ""),
        ("الزكاة الشرعية المستحقة للفترة", "وعاء الزكاة × نسبة الزكاة", r["zakat_due"], "grand"),
    ]
    ri = 8
    for lbl, ref, val, style in z_rows:
        ws_z.cell(ri, 1, lbl).border = bd; ws_z.cell(ri, 2).border = bd
        if val is not None and val != 0: num_cell(ws_z, ri, 3, val, style in ["total","grand"])
        else: ws_z.cell(ri, 3).border = bd
        if style == "section": section_style(ws_z, ri, 3)
        elif style == "total": total_style(ws_z, ri, 3)
        elif style == "grand":
            for j in range(1,4): ws_z.cell(ri,j).fill = fill("065F46"); ws_z.cell(ri,j).font = f(11,True,W); ws_z.cell(ri,j).border = bd_thick
        else: ws_z.cell(ri,1).font = f(11)
        ri += 1

    # 6. الإعدادات وميزان المراجعة
    ws_settings = wb.create_sheet("الإعدادات")
    ws_settings.sheet_view.rightToLeft = True
    hdr_colors = ["1B5E20", "4A148C", "880E4F"]
    hdr_labels = ["الحساب الرئيسي", "التبويب", "الحساب التفصيلي"]
    for j, (lbl, clr) in enumerate(zip(hdr_labels, hdr_colors), 1):
        c = ws_settings.cell(1, j, lbl); c.font = Font(bold=True, color=W, name="Arial", size=11); c.fill = fill(clr); c.alignment = al()
        ws_settings.column_dimensions[get_column_letter(j)].width = 38
    for i, row_data in enumerate(_hierarchy_map, 2):
        ws_settings.cell(i, 1, row_data[0]); ws_settings.cell(i, 2, row_data[1]); ws_settings.cell(i, 3, row_data[2])
    for i, v in enumerate(_main_cats, 2):    ws_settings.cell(i, 27, v)
    for i, v in enumerate(_sub_cats, 2):     ws_settings.cell(i, 28, v)
    for i, v in enumerate(_acc_cats, 2):     ws_settings.cell(i, 29, v)
    for i, v in enumerate(_cost_centers, 2): ws_settings.cell(i, 30, v)
    for col_idx in [27, 28, 29, 30]: ws_settings.column_dimensions[get_column_letter(col_idx)].hidden = True
    ws_settings.sheet_state = 'hidden'

    ws_tb = wb.create_sheet("ميزان المراجعة")
    ws_tb.sheet_view.rightToLeft = True
    _col_colors_tb = {1:"1A3A5C",2:"0D47A1",3:"0D47A1",4:"37474F",5:"37474F",6:"1B5E20",7:"4A148C",8:"880E4F",9:"BF360C"}
    tb_headers = ["اسم الحساب","رصيد مدين (العام الحالي)", "رصيد دائن (العام الحالي)","رصيد مدين (العام السابق)", "رصيد دائن (العام السابق)","الحساب الرئيسي ◄ اختر", "التبويب ◄ اختر","الحساب التفصيلي ◄ اختر", "مركز التكلفة"]
    _tb_col_w = [38, 22, 22, 22, 22, 28, 30, 38, 20]
    for j, (h, w) in enumerate(zip(tb_headers, _tb_col_w), 1):
        c = ws_tb.cell(1, j, h); c.font = Font(bold=True, color=W, name="Arial", size=11); c.fill = fill(_col_colors_tb[j]); c.alignment = al(); c.border = bd
        ws_tb.column_dimensions[get_column_letter(j)].width = w
    
    n_main = len(_main_cats) + 1; n_sub = len(_sub_cats) + 1; n_acc = len(_acc_cats) + 1; n_cc = len(_cost_centers) + 1
    dv_main = DataValidation(type="list", formula1=f"'الإعدادات'!$AA$2:$AA${n_main}", allow_blank=True, showDropDown=False)
    dv_sub  = DataValidation(type="list", formula1=f"'الإعدادات'!$AB$2:$AB${n_sub}", allow_blank=True, showDropDown=False)
    dv_acc  = DataValidation(type="list", formula1=f"'الإعدادات'!$AC$2:$AC${n_acc}", allow_blank=True, showDropDown=False)
    dv_cc   = DataValidation(type="list", formula1=f"'الإعدادات'!$AD$2:$AD${n_cc}", allow_blank=True, showDropDown=False)
    
    ws_tb.add_data_validation(dv_main); ws_tb.add_data_validation(dv_sub); ws_tb.add_data_validation(dv_acc); ws_tb.add_data_validation(dv_cc)
    dv_main.add("F3:F5000"); dv_sub.add("G3:G5000"); dv_acc.add("H3:H5000"); dv_cc.add("I3:I5000")

    _real_data = []
    for _, row in df.iterrows():
        _real_data.append((row.get("name", ""), row.get("dr_curr", 0), row.get("cr_curr", 0), row.get("dr_prev", 0), row.get("cr_prev", 0), row.get("main_cat", ""), row.get("cat", ""), row.get("detail", row.get("name", "")), row.get("cost_center", "")))

    for i, row_data in enumerate(_real_data, 3):
        bg = "F8FBFF" if i % 2 == 0 else "FFFFFF"
        for j, val in enumerate(row_data, 1):
            c = ws_tb.cell(i, j, val); c.border = bd; c.fill = fill(bg); c.font = Font(name="Arial", size=10)
            if j in (2,3,4,5):
                try: c.value = float(val)
                except: c.value = 0.0
                c.number_format = '#,##0.00;(#,##0.00);"-"'

    # ── 7. لوحة التحكم التنفيذية (رسوم بيانية) ──
    ws_dash = wb.create_sheet("لوحة التحكم التنفيذية", 0)
    ws_dash.sheet_view.rightToLeft = True
    for col, w in zip(list("ABCDEFGHI"), [18]*9):
        ws_dash.column_dimensions[col].width = w
    company_header(ws_dash, "لوحة التحكم التنفيذية", f"السنة المالية {year}", cols="A:I")

    def pct_chg2(c, p): return (c-p)/abs(p)*100 if p else 0
    def rating_kpi2(pct): return "\U0001f7e2 ارتفاع" if pct > 0 else ("\U0001f534 انخفاض" if pct < 0 else "\u25fc لا تغيير")

    dash_hdr = ["المؤشر", f"{year} (ر.س)", f"{year-1} (ر.س)", "التغير المطلق", "نسبة التغير %", "التقييم"]
    header_row_style(ws_dash, 6, 6, dash_hdr)
    kpi_rows2 = [
        ("المبيعات الصافية",  r["net_sales"],    r_prev["net_sales"]),
        ("مجمل الربح",        r["gross_profit"], r_prev["gross_profit"]),
        ("الربح التشغيلي",    r["op_profit"],    r_prev["op_profit"]),
        ("صافي الربح",        r["net_az"],       r_prev["net_az"]),
        ("إجمالي الأصول",     r["total_assets"], r_prev["total_assets"]),
        ("حقوق الملكية",      r["equity"],       r_prev["equity"]),
        ("النقدية",           r["cash"],         r_prev["cash"]),
        ("الزكاة المستحقة",   r["zakat_due"],    r_prev["zakat_due"]),
    ]
    ri_d = 7
    for lbl2, curr_v, prev_v in kpi_rows2:
        chg2 = curr_v - prev_v
        pct2 = pct_chg2(curr_v, prev_v)
        ws_dash.cell(ri_d,1,lbl2).border=bd; ws_dash.cell(ri_d,1).font=f(11)
        for jj,v in zip([2,3,4],[curr_v,prev_v,chg2]):
            c2=ws_dash.cell(ri_d,jj,v); c2.border=bd; c2.number_format='#,##0.00;(#,##0.00);"-"'
        cp=ws_dash.cell(ri_d,5,round(pct2,2)); cp.border=bd; cp.number_format='0.00"%"'
        cp.font=f(11,True,"10B981" if pct2>=0 else "EF4444")
        ws_dash.cell(ri_d,6,rating_kpi2(pct2)).border=bd
        if ri_d%2==0:
            for jj in range(1,7): ws_dash.cell(ri_d,jj).fill=fill("F0FDF4")
        ri_d += 1

    chart_start2 = ri_d + 2
    ws_dash.cell(chart_start2-1,1,"بيانات الرسوم البيانية").font=f(11,True,DARK)
    chart_labels2=["المبيعات","مجمل الربح","الربح التشغيلي","صافي الربح"]
    chart_curr2=[r["net_sales"],r["gross_profit"],r["op_profit"],r["net_az"]]
    chart_prev2=[r_prev["net_sales"],r_prev["gross_profit"],r_prev["op_profit"],r_prev["net_az"]]
    header_row_style(ws_dash,chart_start2,3,["البيان",str(year),str(year-1)])
    for i,(lbl2,cv,pv) in enumerate(zip(chart_labels2,chart_curr2,chart_prev2)):
        ws_dash.cell(chart_start2+1+i,1,lbl2)
        ws_dash.cell(chart_start2+1+i,2,cv)
        ws_dash.cell(chart_start2+1+i,3,pv)
    bar1=BarChart(); bar1.type="col"; bar1.grouping="clustered"
    bar1.title=f"مقارنة الأداء المالي {year} vs {year-1}"; bar1.style=10; bar1.width=22; bar1.height=14
    bar1.y_axis.numFmt='#,##0'; bar1.y_axis.title="ر.س"
    d1=Reference(ws_dash,min_col=2,max_col=3,min_row=chart_start2,max_row=chart_start2+4)
    c1=Reference(ws_dash,min_col=1,min_row=chart_start2+1,max_row=chart_start2+4)
    bar1.add_data(d1,titles_from_data=True); bar1.set_categories(c1)
    bar1.series[0].graphicalProperties.solidFill="10B981"
    bar1.series[1].graphicalProperties.solidFill="94A3B8"
    ws_dash.add_chart(bar1,"D6")

    cost_start2=chart_start2+7
    ws_dash.cell(cost_start2,1,"هيكل التكاليف").font=f(11,True,DARK)
    cost_labels2=["تكلفة المبيعات","الرواتب","الإيجارات","التسويق","الاستهلاك","أخرى"]
    cost_vals2=[r["cogs"],r["salaries"],r["rent"],r["marketing"],r["depreciation"],r["misc_exp"]]
    for i,(lbl2,val2) in enumerate(zip(cost_labels2,cost_vals2)):
        ws_dash.cell(cost_start2+1+i,1,lbl2); ws_dash.cell(cost_start2+1+i,2,max(val2,0))
    pie1=PieChart(); pie1.title="هيكل التكاليف"; pie1.style=10; pie1.width=16; pie1.height=14
    pd1=Reference(ws_dash,min_col=2,min_row=cost_start2+1,max_row=cost_start2+6)
    pc1=Reference(ws_dash,min_col=1,min_row=cost_start2+1,max_row=cost_start2+6)
    pie1.add_data(pd1); pie1.set_categories(pc1)
    pie1.dataLabels=DataLabelList(); pie1.dataLabels.showPercent=True
    ws_dash.add_chart(pie1,"D22")

    asset_start2=cost_start2+9
    ws_dash.cell(asset_start2,1,"هيكل الأصول").font=f(11,True,DARK)
    asset_labels2=["نقدية","مدينون","مخزون","أرصدة أخرى","أصول ثابتة","استثمارات"]
    asset_vals2=[r["cash"],r["receivables"],r["inventory"],r["prepaid"],r["fixed_assets"],r["investments"]]
    for i,(lbl2,val2) in enumerate(zip(asset_labels2,asset_vals2)):
        ws_dash.cell(asset_start2+1+i,1,lbl2); ws_dash.cell(asset_start2+1+i,2,max(val2,0))
    pie2=PieChart(); pie2.title="هيكل الأصول"; pie2.style=26; pie2.width=16; pie2.height=14
    pd2=Reference(ws_dash,min_col=2,min_row=asset_start2+1,max_row=asset_start2+6)
    pc2=Reference(ws_dash,min_col=1,min_row=asset_start2+1,max_row=asset_start2+6)
    pie2.add_data(pd2); pie2.set_categories(pc2)
    pie2.dataLabels=DataLabelList(); pie2.dataLabels.showPercent=True
    ws_dash.add_chart(pie2,"G22")

    # ── 8. النسب المالية ──
    ws_rat=wb.create_sheet("النسب المالية الموسّعة")
    ws_rat.sheet_view.rightToLeft=True
    for col,w in zip(list("ABCDEF"),[34,30,20,20,18,16]):
        ws_rat.column_dimensions[col].width=w
    company_header(ws_rat,"تقرير النسب المالية الموسّعة",f"عن السنة المنتهية 31 ديسمبر {year}",cols="A:F")
    header_row_style(ws_rat,6,6,["النسبة","المعادلة",f"{year}",f"{year-1}","المعيار","التقييم"])

    def safe_r(num,den,fmt="x"):
        if not den: return "–"
        v=num/den
        if fmt=="x": return round(v,2)
        if fmt=="%": return round(v*100,2)
        if fmt=="d": return round(v,0)
        return round(v,2)
    def score_r(val,thresholds,labels):
        if val is None or val=="–": return "–"
        try: v=float(val)
        except: return "–"
        for t,lbl in zip(thresholds,labels):
            if v>=t: return lbl
        return labels[-1]

    ns_r=r["net_sales"]; gp_r=r["gross_profit"]; op_r=r["op_profit"]; net_r=r["net_az"]
    ta_r=r["total_assets"]; eq_r=r["equity"]; ca_r=r["total_current"]; cl_r=r["total_cl"]
    cash_r=r["cash"]; inv_r=r["inventory"]; rec_r=r["receivables"]; cogs_r=r["cogs"]
    tl_r=r["total_cl"]+r["total_ncl"]+r["zakat_due"]
    avg_rec_r=(rec_r+r_prev["receivables"])/2
    avg_inv_r=(inv_r+r_prev["inventory"])/2
    avg_ta_r=(ta_r+r_prev["total_assets"])/2
    gpm_r=safe_r(gp_r,ns_r,"%"); gpm_rp=safe_r(r_prev["gross_profit"],r_prev["net_sales"],"%")
    opm_r=safe_r(op_r,ns_r,"%"); opm_rp=safe_r(r_prev["op_profit"],r_prev["net_sales"],"%")
    npm_r=safe_r(net_r,ns_r,"%"); npm_rp=safe_r(r_prev["net_az"],r_prev["net_sales"],"%")
    roa_r=safe_r(net_r,ta_r,"%"); roa_rp=safe_r(r_prev["net_az"],r_prev["total_assets"],"%")
    roe_r=safe_r(net_r,eq_r,"%"); roe_rp=safe_r(r_prev["net_az"],r_prev["equity"],"%")
    cr_r=safe_r(ca_r,cl_r); cr_rp=safe_r(r_prev["total_current"],r_prev["total_cl"])
    qr_r=safe_r(ca_r-inv_r,cl_r); qr_rp=safe_r(r_prev["total_current"]-r_prev["inventory"],r_prev["total_cl"])
    der_r=safe_r(tl_r,eq_r); dar_r=safe_r(tl_r*100,ta_r,"%")
    dso_r=safe_r(avg_rec_r*365,ns_r,"d") if ns_r else "–"
    dio_r=safe_r(avg_inv_r*365,cogs_r,"d") if cogs_r else "–"
    ato_r=safe_r(ns_r,avg_ta_r)

    ratio_rows2=[
        ("أولاً: نسب السيولة",None,None,None,None,True),
        ("نسبة التداول",cr_r,cr_rp,"≥ 2.0x",score_r(cr_r,[2.0,1.5,1.0],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("النسبة السريعة",qr_r,qr_rp,"≥ 1.0x",score_r(qr_r,[1.5,1.0,0.5],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("نسبة النقدية",safe_r(cash_r,cl_r),"–","≥ 0.5x","–",False),
        ("ثانياً: نسب الربحية",None,None,None,None,True),
        ("هامش مجمل الربح %",gpm_r,gpm_rp,"≥ 30%",score_r(gpm_r,[30,20,10],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("هامش الربح التشغيلي %",opm_r,opm_rp,"≥ 15%",score_r(opm_r,[15,8,3],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("هامش الربح الصافي %",npm_r,npm_rp,"≥ 10%",score_r(npm_r,[10,5,2],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("العائد على الأصول ROA %",roa_r,roa_rp,"≥ 5%",score_r(roa_r,[10,5,2],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("العائد على الملكية ROE %",roe_r,roe_rp,"≥ 15%",score_r(roe_r,[20,12,5],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("ثالثاً: نسب الكفاءة",None,None,None,None,True),
        ("أيام التحصيل DSO",dso_r,"–","< 45 يوم","–",False),
        ("أيام المخزون DIO",dio_r,"–","< 60 يوم","–",False),
        ("معدل دوران الأصول",ato_r,"–","≥ 1.0x",score_r(ato_r,[1.5,1.0,0.5],["\U0001f7e2 ممتاز","\U0001f7e1 جيد","\U0001f7e0 مقبول","\U0001f534 ضعيف"]),False),
        ("رابعاً: نسب الرفع المالي",None,None,None,None,True),
        ("نسبة الدين للملكية D/E",der_r,"–","≤ 1.0x","–",False),
        ("نسبة الدين للأصول %",dar_r,"–","≤ 50%","–",False),
        ("نسبة حقوق الملكية %",safe_r(eq_r,ta_r,"%"),"–","≥ 50%","–",False),
    ]
    ri_rat=7
    for lbl2,v_c2,v_p2,std2,rating2,is_sec2 in ratio_rows2:
        ws_rat.cell(ri_rat,1,lbl2).border=bd; ws_rat.cell(ri_rat,1).font=f(11,is_sec2)
        for jj,v in zip([2,3,4,5],[v_c2,v_p2,std2,rating2]):
            c2=ws_rat.cell(ri_rat,jj,v if v is not None else ""); c2.border=bd
            if isinstance(v,(int,float)): c2.number_format='#,##0.00'
        ws_rat.cell(ri_rat,6,"").border=bd
        if is_sec2: section_style(ws_rat,ri_rat,6)
        elif ri_rat%2==0:
            for jj in range(1,7): ws_rat.cell(ri_rat,jj).fill=fill("F0FDF4")
        ri_rat+=1
    rc_start=ri_rat+2
    header_row_style(ws_rat,rc_start,3,["الهامش",str(year),str(year-1)])
    margins2=[("هامش مجمل الربح",gpm_r,gpm_rp),("هامش التشغيلي",opm_r,opm_rp),("هامش صافي",npm_r,npm_rp),("ROA",roa_r,roa_rp),("ROE",roe_r,roe_rp)]
    for i,(lbl2,vc,vp) in enumerate(margins2):
        ws_rat.cell(rc_start+1+i,1,lbl2)
        ws_rat.cell(rc_start+1+i,2,vc if isinstance(vc,(int,float)) else 0)
        ws_rat.cell(rc_start+1+i,3,vp if isinstance(vp,(int,float)) else 0)
    bar2=BarChart(); bar2.type="col"; bar2.grouping="clustered"; bar2.title="مقارنة النسب الربحية %"
    bar2.style=10; bar2.width=20; bar2.height=14; bar2.y_axis.title="%"
    d2=Reference(ws_rat,min_col=2,max_col=3,min_row=rc_start,max_row=rc_start+len(margins2))
    c2r=Reference(ws_rat,min_col=1,min_row=rc_start+1,max_row=rc_start+len(margins2))
    bar2.add_data(d2,titles_from_data=True); bar2.set_categories(c2r)
    bar2.series[0].graphicalProperties.solidFill="10B981"
    bar2.series[1].graphicalProperties.solidFill="94A3B8"
    ws_rat.add_chart(bar2,"D6")

    # ── 9. الميزانية التقديرية ──
    ws_bud=wb.create_sheet("الميزانية التقديرية vs الفعلي")
    ws_bud.sheet_view.rightToLeft=True
    for col,w in zip(list("ABCDEFG"),[32,22,22,22,18,18,22]):
        ws_bud.column_dimensions[col].width=w
    company_header(ws_bud,"الميزانية التقديرية مقارنةً بالأداء الفعلي",f"عن السنة المنتهية 31 ديسمبر {year}",cols="A:G")
    ws_bud.merge_cells("A5:G5")
    ws_bud["A5"]="\U0001f4a1 أدخل قيم الميزانية التقديرية في العمود C (باللون الأخضر) — تُحسب الانحرافات والرسوم البيانية تلقائياً"
    ws_bud["A5"].font=f(10,True,"065F46"); ws_bud["A5"].fill=fill("DCFCE7"); ws_bud["A5"].alignment=al()
    header_row_style(ws_bud,6,7,["البيـــان",f"الفعلي {year} (ر.س)",f"التقديري {year} (ر.س)","الانحراف (ر.س)","الانحراف %","نسبة التحقق %","الحكم"])
    bud_fill_pat=PatternFill("solid",fgColor="DCFCE7")
    bud_items2=[
        ("إيرادات المبيعات الصافية",r["net_sales"],"item"),
        ("تكلفة المبيعات",-r["cogs"],"item"),
        ("إجمالي الربح",r["gross_profit"],"total"),
        ("",None,""),
        ("المصاريف التشغيلية",None,"section"),
        ("الرواتب والأجور",-r["salaries"],"item"),
        ("الإيجارات",-r["rent"],"item"),
        ("GOSI",-r["gosi"],"item"),
        ("التسويق والإعلان",-r["marketing"],"item"),
        ("النقل والشحن",-r["transport"],"item"),
        ("الخدمات والمرافق",-r["utilities"],"item"),
        ("الاستهلاك والإهلاك",-r["depreciation"],"item"),
        ("مصاريف أخرى",-r["misc_exp"],"item"),
        ("إجمالي المصاريف",-r["total_opex"],"total"),
        ("الربح التشغيلي EBIT",r["op_profit"],"total"),
        ("",None,""),
        ("إيرادات أخرى",r["other_inc"],"item"),
        ("تكاليف التمويل",-r["finance_cost"],"item"),
        ("الربح قبل الزكاة",r["net_bz"],"total"),
        ("الزكاة المستحقة",-r["zakat_due"],"item"),
        ("صافي الربح",r["net_az"],"grand"),
    ]
    ri_b=7
    for lbl2,actual_val2,style2 in bud_items2:
        is_sec3=(style2=="section"); is_tot3=(style2 in ["total","grand"])
        ws_bud.cell(ri_b,1,lbl2).border=bd; ws_bud.cell(ri_b,1).font=f(11,is_tot3 or is_sec3)
        if actual_val2 is not None:
            ca3=ws_bud.cell(ri_b,2,actual_val2); ca3.border=bd; ca3.number_format='#,##0.00;(#,##0.00);"-"'; ca3.font=f(11,is_tot3)
            cb3=ws_bud.cell(ri_b,3,""); cb3.border=bd; cb3.fill=bud_fill_pat; cb3.number_format='#,##0.00;(#,##0.00);"-"'; cb3.font=f(11,False,"065F46")
            dev3=ws_bud.cell(ri_b,4,f'=IF(C{ri_b}=0,"-",B{ri_b}-C{ri_b})'); dev3.border=bd; dev3.number_format='#,##0.00;(#,##0.00);"-"'; dev3.font=f(11,is_tot3)
            pct3=ws_bud.cell(ri_b,5,f'=IF(OR(C{ri_b}=0,B{ri_b}=0),"-",(B{ri_b}-C{ri_b})/C{ri_b}*100)'); pct3.border=bd; pct3.number_format='0.00;(0.00);"-"'
            ach3=ws_bud.cell(ri_b,6,f'=IF(C{ri_b}=0,"-",B{ri_b}/C{ri_b}*100)'); ach3.border=bd; ach3.number_format='0.0"%"'
            jdg3=ws_bud.cell(ri_b,7,f'=IF(C{ri_b}=0,"أدخل التقديري",IF(B{ri_b}/C{ri_b}>1.05,"تجاوز التقدير",IF(B{ri_b}/C{ri_b}>=0.95,"ضمن النطاق",IF(B{ri_b}/C{ri_b}>=0.80,"دون التقدير","فجوة كبيرة"))))'); jdg3.border=bd
        else:
            for jj in range(1,8): ws_bud.cell(ri_b,jj).border=bd
        if style2=="section": section_style(ws_bud,ri_b,7)
        elif style2=="total": total_style(ws_bud,ri_b,7)
        elif style2=="grand":
            for jj in range(1,8): ws_bud.cell(ri_b,jj).fill=fill(DARK); ws_bud.cell(ri_b,jj).font=f(11,True,W); ws_bud.cell(ri_b,jj).border=bd_thick
        elif ri_b%2==0:
            for jj in [1,2]: ws_bud.cell(ri_b,jj).fill=fill("F8FAFC")
        ri_b+=1
    bc_start2=ri_b+2
    bc_labels2=["المبيعات","مجمل الربح","الربح التشغيلي","صافي الربح"]
    bc_actual2=[r["net_sales"],r["gross_profit"],r["op_profit"],r["net_az"]]
    header_row_style(ws_bud,bc_start2,3,["البيان","الفعلي","التقديري"])
    for i,(lbl2,av2) in enumerate(zip(bc_labels2,bc_actual2)):
        ws_bud.cell(bc_start2+1+i,1,lbl2); ws_bud.cell(bc_start2+1+i,2,av2)
        cb4=ws_bud.cell(bc_start2+1+i,3,""); cb4.fill=bud_fill_pat
    bar3=BarChart(); bar3.type="col"; bar3.grouping="clustered"; bar3.title="الفعلي vs التقديري"
    bar3.style=10; bar3.width=22; bar3.height=14; bar3.y_axis.numFmt='#,##0'; bar3.y_axis.title="ر.س"
    d3=Reference(ws_bud,min_col=2,max_col=3,min_row=bc_start2,max_row=bc_start2+4)
    c3r=Reference(ws_bud,min_col=1,min_row=bc_start2+1,max_row=bc_start2+4)
    bar3.add_data(d3,titles_from_data=True); bar3.set_categories(c3r)
    bar3.series[0].graphicalProperties.solidFill="0F2942"
    bar3.series[1].graphicalProperties.solidFill="10B981"
    ws_bud.add_chart(bar3,"D"+str(bc_start2))

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def create_tb_template():
    wb = Workbook()
    DARK, W = "1A3A5C", "FFFFFF"
    bd = Border(
        left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),  bottom=Side(style="thin", color="CCCCCC")
    )
    def f(sz=11, bold=False, color="000000"): return Font(name="Arial", size=sz, bold=bold, color=color)
    def al(h="center", v="center"): return Alignment(horizontal=h, vertical=v, wrap_text=True)
    def fill(clr): return PatternFill("solid", fgColor=clr)

    # ── شيت ميزان المراجعة ──
    ws = wb.active
    ws.title = "ميزان المراجعة"
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "A3"

    # عنوان
    ws.merge_cells("A1:I1")
    ws["A1"] = "📊 قالب ميزان المراجعة — النظام المالي والزكوي الشامل"
    ws["A1"].font = Font(name="Arial", size=13, bold=True, color=W)
    ws["A1"].fill = fill(DARK); ws["A1"].alignment = al()
    ws.row_dimensions[1].height = 28

    # رؤوس الأعمدة
    headers = [
        ("اسم الحساب",                      38, "1A3A5C"),
        ("رصيد مدين — العام الحالي",         22, "0D47A1"),
        ("رصيد دائن — العام الحالي",         22, "0D47A1"),
        ("رصيد مدين — العام السابق",         22, "37474F"),
        ("رصيد دائن — العام السابق",         22, "37474F"),
        ("الحساب الرئيسي ◄ اختر",           28, "1B5E20"),
        ("التبويب ◄ اختر",                   30, "4A148C"),
        ("الحساب التفصيلي ◄ اختر",           38, "880E4F"),
        ("مركز التكلفة",                     20, "BF360C"),
    ]
    for j, (hdr, w, clr) in enumerate(headers, 1):
        c = ws.cell(2, j, hdr)
        c.font = Font(name="Arial", size=10, bold=True, color=W)
        c.fill = fill(clr); c.alignment = al(); c.border = bd
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[2].height = 32

    # قوائم التحقق من البيانات
    ws_set = wb.create_sheet("الإعدادات")
    ws_set.sheet_state = "hidden"

    main_cats   = ["قائمة المركز المالي", "قائمة الدخل"]
    sub_cats    = ["الأصول المتداولة", "الأصول غير المتداولة", "الخصوم المتداولة",
                   "الخصوم غير المتداولة", "حقوق الملكية", "الإيرادات",
                   "تكاليف البضاعة المباعة", "المصاريف", "غير مبوب"]
    acc_cats    = ["النقدية وما في حكمها", "مدينون تجاريون", "مخزون بضاعة",
                   "مصاريف مقدمة وأرصدة مدينة أخرى", "أطراف ذات علاقة (مدين)",
                   "الممتلكات والآلات والمعدات (صافي)", "الاستثمارات في شركات تابعة",
                   "موردون تجاريون", "مصاريف مستحقة وأرصدة دائنة أخرى",
                   "أطراف ذات علاقة (دائن)", "مخصص الزكاة الشرعية",
                   "تسهيلات ائتمانية قصيرة الأجل", "مخصص مكافأة نهاية الخدمة",
                   "قروض طويلة الأجل", "رأس المال", "احتياطي نظامي",
                   "الأرباح المبقاة", "مسحوبات الشركاء", "الإيرادات والمبيعات",
                   "مردودات ومسموحات المبيعات", "تكلفة الإيرادات",
                   "الرواتب والأجور", "الإيجارات", "التأمينات الاجتماعية (GOSI)",
                   "مصاريف التسويق والإعلان", "النقل والتوزيع", "الكهرباء والمياه",
                   "الاستهلاك والإهلاك", "تكاليف تمويل وفوائد بنكية",
                   "إيرادات أخرى", "مصاريف عمومية وإدارية أخرى"]
    cost_centers = ["المركز الرئيسي", "الرياض", "جدة", "الدمام", "مكة المكرمة",
                    "المدينة المنورة", "أبها", "تبوك", "القصيم", "إدارة عامة",
                    "فرع 1", "فرع 2", "مستودع رئيسي", "مشروع خاص"]

    for i, v in enumerate(main_cats,   2): ws_set.cell(i, 1, v)
    for i, v in enumerate(sub_cats,    2): ws_set.cell(i, 2, v)
    for i, v in enumerate(acc_cats,    2): ws_set.cell(i, 3, v)
    for i, v in enumerate(cost_centers,2): ws_set.cell(i, 4, v)

    dv_main = DataValidation(type="list", formula1=f"'الإعدادات'!$A$2:$A${len(main_cats)+1}",   allow_blank=True)
    dv_sub  = DataValidation(type="list", formula1=f"'الإعدادات'!$B$2:$B${len(sub_cats)+1}",    allow_blank=True)
    dv_acc  = DataValidation(type="list", formula1=f"'الإعدادات'!$C$2:$C${len(acc_cats)+1}",   allow_blank=True)
    dv_cc   = DataValidation(type="list", formula1=f"'الإعدادات'!$D$2:$D${len(cost_centers)+1}",allow_blank=True)
    ws.add_data_validation(dv_main); ws.add_data_validation(dv_sub)
    ws.add_data_validation(dv_acc);  ws.add_data_validation(dv_cc)
    dv_main.add("F3:F2000"); dv_sub.add("G3:G2000")
    dv_acc.add("H3:H2000");  dv_cc.add("I3:I2000")

    # بيانات نموذجية (أمثلة)
    sample_rows = [
        # (اسم الحساب, مدين_حالي, دائن_حالي, مدين_سابق, دائن_سابق, رئيسي, تبويب, تفصيلي, مركز_تكلفة)
        ("الصندوق والبنوك",        250000, 0,      180000, 0,      "قائمة المركز المالي", "الأصول المتداولة",      "النقدية وما في حكمها",              "المركز الرئيسي"),
        ("العملاء والمدينون",       420000, 0,      310000, 0,      "قائمة المركز المالي", "الأصول المتداولة",      "مدينون تجاريون",                    "المركز الرئيسي"),
        ("مخزون البضاعة",           380000, 0,      290000, 0,      "قائمة المركز المالي", "الأصول المتداولة",      "مخزون بضاعة",                       "مستودع رئيسي"),
        ("سلف ومصاريف مقدمة",       85000,  0,      62000,  0,      "قائمة المركز المالي", "الأصول المتداولة",      "مصاريف مقدمة وأرصدة مدينة أخرى",   "المركز الرئيسي"),
        ("أثاث وأجهزة (صافي)",      150000, 0,      175000, 0,      "قائمة المركز المالي", "الأصول غير المتداولة",  "الممتلكات والآلات والمعدات (صافي)", "المركز الرئيسي"),
        ("الموردون والدائنون",       0,      195000, 0,      140000, "قائمة المركز المالي", "الخصوم المتداولة",      "موردون تجاريون",                    "المركز الرئيسي"),
        ("مصاريف مستحقة الدفع",     0,      68000,  0,      52000,  "قائمة المركز المالي", "الخصوم المتداولة",      "مصاريف مستحقة وأرصدة دائنة أخرى",  "المركز الرئيسي"),
        ("رأس المال",               0,      500000, 0,      500000, "قائمة المركز المالي", "حقوق الملكية",          "رأس المال",                         "المركز الرئيسي"),
        ("الأرباح المبقاة",          0,      220000, 0,      158000, "قائمة المركز المالي", "حقوق الملكية",          "الأرباح المبقاة",                   "المركز الرئيسي"),
        ("مسحوبات الشركاء",          120000, 0,      85000,  0,      "قائمة المركز المالي", "حقوق الملكية",          "مسحوبات الشركاء",                   "المركز الرئيسي"),
        ("إيرادات المبيعات",         0,      3609526,0,      3200000,"قائمة الدخل",         "الإيرادات",             "الإيرادات والمبيعات",               "المركز الرئيسي"),
        ("مردودات المبيعات",         45000,  0,      38000,  0,      "قائمة الدخل",         "الإيرادات",             "مردودات ومسموحات المبيعات",         "المركز الرئيسي"),
        ("تكلفة البضاعة المباعة",    2200000,0,      1980000,0,      "قائمة الدخل",         "تكاليف البضاعة المباعة","تكلفة الإيرادات",                   "المركز الرئيسي"),
        ("رواتب وأجور",              380000, 0,      340000, 0,      "قائمة الدخل",         "المصاريف",              "الرواتب والأجور",                   "المركز الرئيسي"),
        ("إيجارات",                  120000, 0,      110000, 0,      "قائمة الدخل",         "المصاريف",              "الإيجارات",                         "المركز الرئيسي"),
        ("تأمينات اجتماعية GOSI",    42000,  0,      37000,  0,      "قائمة الدخل",         "المصاريف",              "التأمينات الاجتماعية (GOSI)",        "المركز الرئيسي"),
        ("مصاريف تسويق وإعلان",      55000,  0,      48000,  0,      "قائمة الدخل",         "المصاريف",              "مصاريف التسويق والإعلان",           "المركز الرئيسي"),
        ("نقل وتوزيع",               38000,  0,      32000,  0,      "قائمة الدخل",         "المصاريف",              "النقل والتوزيع",                    "المركز الرئيسي"),
        ("كهرباء ومياه",             24000,  0,      21000,  0,      "قائمة الدخل",         "المصاريف",              "الكهرباء والمياه",                  "المركز الرئيسي"),
        ("استهلاك وإهلاك",           45000,  0,      40000,  0,      "قائمة الدخل",         "المصاريف",              "الاستهلاك والإهلاك",                "المركز الرئيسي"),
        ("مصاريف إدارية أخرى",       62000,  0,      55000,  0,      "قائمة الدخل",         "المصاريف",              "مصاريف عمومية وإدارية أخرى",        "المركز الرئيسي"),
        ("فوائد بنكية",              18000,  0,      15000,  0,      "قائمة الدخل",         "المصاريف",              "تكاليف تمويل وفوائد بنكية",         "المركز الرئيسي"),
        ("إيرادات أخرى",             0,      12000,  0,      8000,   "قائمة الدخل",         "الإيرادات",             "إيرادات أخرى",                      "المركز الرئيسي"),
    ]

    num_fmt = '#,##0.00;(#,##0.00);"-"'
    for row_i, row_data in enumerate(sample_rows, 3):
        bg = "F8FBFF" if row_i % 2 == 0 else "FFFFFF"
        for col_j, val in enumerate(row_data, 1):
            c = ws.cell(row_i, col_j, val)
            c.border = bd
            c.fill = fill(bg)
            c.font = Font(name="Arial", size=10, color="374151" if col_j == 1 else "000000")
            c.alignment = Alignment(horizontal="right" if col_j == 1 else "center", vertical="center")
            if col_j in (2, 3, 4, 5):
                c.number_format = num_fmt

    # صف فارغ للبداية (بعد الأمثلة)
    blank_start = len(sample_rows) + 3
    for extra_row in range(blank_start, blank_start + 100):
        bg = "F8FBFF" if extra_row % 2 == 0 else "FFFFFF"
        for col_j in range(1, 10):
            c = ws.cell(extra_row, col_j, "")
            c.border = bd; c.fill = fill(bg)
            if col_j in (2, 3, 4, 5): c.number_format = num_fmt

    # ── شيت التعليمات ──
    ws_help = wb.create_sheet("📋 تعليمات الاستخدام")
    ws_help.sheet_view.rightToLeft = True
    ws_help.column_dimensions["A"].width = 20
    ws_help.column_dimensions["B"].width = 70

    ws_help.merge_cells("A1:B1")
    ws_help["A1"] = "📋 تعليمات استخدام قالب ميزان المراجعة"
    ws_help["A1"].font = Font(name="Arial", size=13, bold=True, color=W)
    ws_help["A1"].fill = fill(DARK); ws_help["A1"].alignment = al()
    ws_help.row_dimensions[1].height = 28

    instructions = [
        ("العمود A",   "اسم الحساب — أدخل اسم الحساب المحاسبي كما هو في دفاترك"),
        ("العمود B",   "رصيد مدين للعام الحالي — أدخل أرقاماً موجبة فقط، الحسابات ذات رصيد مدين"),
        ("العمود C",   "رصيد دائن للعام الحالي — أدخل أرقاماً موجبة فقط، الحسابات ذات رصيد دائن"),
        ("العمود D",   "رصيد مدين للعام السابق — نفس القاعدة للسنة الماضية"),
        ("العمود E",   "رصيد دائن للعام السابق — نفس القاعدة للسنة الماضية"),
        ("العمود F",   "الحساب الرئيسي — اختر من القائمة: قائمة المركز المالي أو قائمة الدخل"),
        ("العمود G",   "التبويب — اختر التصنيف الفرعي من القائمة المنسدلة"),
        ("العمود H",   "الحساب التفصيلي — اختر التصنيف التفصيلي الأدق من القائمة"),
        ("العمود I",   "مركز التكلفة — اختر الفرع أو المركز من القائمة"),
        ("", ""),
        ("⚠️ تنبيه 1", "لا تترك أعمدة F و G و H فارغة — بدونها لن تُصنَّف الأرقام في القوائم المالية"),
        ("⚠️ تنبيه 2", "لا تدخل الرقم في عمودي المدين والدائن معاً لنفس الحساب — اختر الأنسب"),
        ("⚠️ تنبيه 3", "لا تحذف أو تعيد تسمية شيت 'ميزان المراجعة' — النظام يقرأ هذا الاسم تحديداً"),
        ("⚠️ تنبيه 4", "الصفوف الملونة بالأزرق الفاتح هي أمثلة توضيحية — احذفها أو استبدلها ببياناتك"),
        ("", ""),
        ("✅ نصيحة",   "ابدأ من الصف 3 وأضف حساباتك — الصفوف الفارغة الموجودة قابلة للتعديل مباشرة"),
        ("✅ نصيحة",   "يمكنك إضافة صفوف جديدة بنسخ صف موجود للحفاظ على التنسيق وقوائم الاختيار"),
        ("✅ نصيحة",   "عند الانتهاء، احفظ الملف وارفعه من الشريط الجانبي في التطبيق"),
    ]
    for row_i, (col_a, col_b) in enumerate(instructions, 2):
        c_a = ws_help.cell(row_i, 1, col_a)
        c_b = ws_help.cell(row_i, 2, col_b)
        is_warn = "⚠️" in col_a; is_tip = "✅" in col_a; is_col = col_a.startswith("العمود")
        bg_clr = "FFF3CD" if is_warn else ("DCFCE7" if is_tip else ("EFF6FF" if is_col else "FFFFFF"))
        for c in [c_a, c_b]:
            c.fill = fill(bg_clr); c.border = bd
            c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            c.font = Font(name="Arial", size=10, bold=is_warn or is_tip)
        ws_help.row_dimensions[row_i].height = 22

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def show_budget_tab(r_curr, report_year):
    st.markdown(f'<div class="sec-title">🎯 الميزانية التقديرية مقارنةً بالأداء الفعلي | {report_year}</div>', unsafe_allow_html=True)

    # ── مفتاح session_state للتقديري ──
    if "budget" not in st.session_state:
        st.session_state.budget = {}

    bud = st.session_state.budget

    # ── بنود الميزانية ──
    bud_items = [
        ("rev",   "💰 إيرادات المبيعات الصافية",  r_curr["net_sales"]),
        ("cogs",  "📦 تكلفة المبيعات",             r_curr["cogs"]),
        ("sal",   "👥 الرواتب والأجور",             r_curr["salaries"]),
        ("rent",  "🏢 الإيجارات",                  r_curr["rent"]),
        ("gosi",  "🛡️ التأمينات (GOSI)",           r_curr["gosi"]),
        ("mkt",   "📣 التسويق والإعلان",            r_curr["marketing"]),
        ("trn",   "🚚 النقل والتوزيع",              r_curr["transport"]),
        ("util",  "💡 الكهرباء والمياه",            r_curr["utilities"]),
        ("dep",   "🔧 الاستهلاك والإهلاك",          r_curr["depreciation"]),
        ("misc",  "📋 مصاريف أخرى",                r_curr["misc_exp"]),
        ("fin",   "🏦 تكاليف التمويل",              r_curr["finance_cost"]),
        ("oth",   "➕ إيرادات أخرى",               r_curr["other_inc"]),
    ]

    # ── طريقة الإدخال ──
    st.markdown("#### اختر طريقة إدخال الميزانية التقديرية")
    method = st.radio("", ["✏️ إدخال يدوي", "📂 رفع ملف Excel"], horizontal=True, label_visibility="collapsed")

    if method == "📂 رفع ملف Excel":
        st.info("ارفع ملف Excel يحتوي على عمودين: **البند** و**المبلغ التقديري** — يمكنك تحميل النموذج أدناه")
        bud_template_data = _make_budget_excel_template(bud_items, report_year)
        st.download_button("⬇️ تحميل نموذج Excel للتقديري", data=bud_template_data,
                           file_name=f"نموذج_الميزانية_التقديرية_{report_year}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        uploaded_bud = st.file_uploader("📂 ارفع ملف Excel التقديري", type=["xlsx"], key="bud_upload")
        if uploaded_bud:
            try:
                df_bud = pd.read_excel(uploaded_bud, sheet_name=0)
                df_bud.columns = [str(c).strip() for c in df_bud.columns]
                key_col = df_bud.columns[0]; val_col = df_bud.columns[1]
                key_map = {row[key_col]: row[val_col] for _, row in df_bud.iterrows()
                           if pd.notna(row[val_col]) and str(row[val_col]).strip() != ""}
                code_map = {lbl: code for code, lbl, _ in bud_items}
                loaded = 0
                for lbl, val in key_map.items():
                    for item_lbl, code in code_map.items():
                        if str(lbl).strip() in item_lbl or item_lbl in str(lbl).strip():
                            try:
                                bud[code] = float(val)
                                loaded += 1
                            except: pass
                st.session_state.budget = bud
                st.success(f"✅ تم تحميل {loaded} بنداً من الملف")
            except Exception as e:
                st.error(f"خطأ في قراءة الملف: {e}")

    else:  # إدخال يدوي
        st.markdown("#### أدخل الأرقام التقديرية لكل بند")
        cols_per_row = 3
        rows = [bud_items[i:i+cols_per_row] for i in range(0, len(bud_items), cols_per_row)]
        for row_group in rows:
            cols = st.columns(cols_per_row)
            for col, (code, label, actual) in zip(cols, row_group):
                with col:
                    default_val = float(bud.get(code, actual))
                    entered = st.number_input(label, min_value=0.0, value=default_val,
                                              step=1000.0, format="%.0f", key=f"bud_{code}")
                    bud[code] = entered
        if st.button("💾 حفظ الميزانية التقديرية", type="primary", use_container_width=True):
            st.session_state.budget = bud
            st.success("✅ تم حفظ الميزانية التقديرية")

    st.markdown("---")

    # ── حساب الفعلي من r_curr ──
    actual_map = {
        "rev":  r_curr["net_sales"],  "cogs": r_curr["cogs"],
        "sal":  r_curr["salaries"],   "rent": r_curr["rent"],
        "gosi": r_curr["gosi"],       "mkt":  r_curr["marketing"],
        "trn":  r_curr["transport"],  "util": r_curr["utilities"],
        "dep":  r_curr["depreciation"],"misc": r_curr["misc_exp"],
        "fin":  r_curr["finance_cost"],"oth":  r_curr["other_inc"],
    }

    # ── جدول المقارنة ──
    st.markdown("#### 📊 جدول المقارنة التفصيلي")
    rows_data = {"البيان": [], "الفعلي (ر.س)": [], "التقديري (ر.س)": [],
                 "الانحراف (ر.س)": [], "الانحراف %": [], "نسبة التحقق %": [], "الحكم": []}

    for code, label, _ in bud_items:
        actual_v = actual_map[code]
        budget_v = bud.get(code, actual_v)
        dev      = actual_v - budget_v
        dev_pct  = (dev / budget_v * 100) if budget_v else 0
        ach_pct  = (actual_v / budget_v * 100) if budget_v else 0
        if   ach_pct > 105: verdict = "🟢 تجاوز التقدير"
        elif ach_pct >= 95: verdict = "🟡 ضمن النطاق"
        elif ach_pct >= 80: verdict = "🟠 دون التقدير"
        else:               verdict = "🔴 فجوة كبيرة"

        rows_data["البيان"].append(label)
        rows_data["الفعلي (ر.س)"].append(f"{actual_v:,.0f}")
        rows_data["التقديري (ر.س)"].append(f"{budget_v:,.0f}")
        rows_data["الانحراف (ر.س)"].append(f"{dev:+,.0f}")
        rows_data["الانحراف %"].append(f"{dev_pct:+.1f}%")
        rows_data["نسبة التحقق %"].append(f"{ach_pct:.1f}%")
        rows_data["الحكم"].append(verdict)

    # ── بنود مشتقة: إجمالي الربح / صافي الربح ──
    rev_a  = actual_map["rev"];  rev_b  = bud.get("rev",  rev_a)
    cogs_a = actual_map["cogs"]; cogs_b = bud.get("cogs", cogs_a)
    gp_a   = rev_a - cogs_a;    gp_b   = rev_b - cogs_b
    opex_a = sum(actual_map[k] for k in ["sal","rent","gosi","mkt","trn","util","dep","misc"])
    opex_b = sum(bud.get(k, actual_map[k]) for k in ["sal","rent","gosi","mkt","trn","util","dep","misc"])
    op_a   = gp_a - opex_a;     op_b   = gp_b - opex_b
    oth_a  = actual_map["oth"];  oth_b  = bud.get("oth", oth_a)
    fin_a  = actual_map["fin"];  fin_b  = bud.get("fin", fin_a)
    net_a  = op_a + oth_a - fin_a; net_b = op_b + oth_b - fin_b

    for lbl_d, a_v, b_v in [("📈 إجمالي الربح", gp_a, gp_b), ("⚙️ الربح التشغيلي", op_a, op_b), ("💰 صافي الربح", net_a, net_b)]:
        dev = a_v - b_v; dev_pct = (dev/b_v*100) if b_v else 0
        ach_pct = (a_v/b_v*100) if b_v else 0
        if   ach_pct > 105: verdict = "🟢 تجاوز التقدير"
        elif ach_pct >= 95: verdict = "🟡 ضمن النطاق"
        elif ach_pct >= 80: verdict = "🟠 دون التقدير"
        else:               verdict = "🔴 فجوة كبيرة"
        rows_data["البيان"].append(f"**{lbl_d}**")
        rows_data["الفعلي (ر.س)"].append(f"{a_v:,.0f}")
        rows_data["التقديري (ر.س)"].append(f"{b_v:,.0f}")
        rows_data["الانحراف (ر.س)"].append(f"{dev:+,.0f}")
        rows_data["الانحراف %"].append(f"{dev_pct:+.1f}%")
        rows_data["نسبة التحقق %"].append(f"{ach_pct:.1f}%")
        rows_data["الحكم"].append(verdict)

    st.table(pd.DataFrame(rows_data))

    # ── بطاقات ملخص ──
    st.markdown("#### 🏅 ملخص الأداء مقابل التقديري")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("المبيعات الفعلية",   f"{rev_a:,.0f} ر.س",  delta=f"{(rev_a-rev_b):+,.0f}")
    k2.metric("إجمالي الربح الفعلي",f"{gp_a:,.0f} ر.س",   delta=f"{(gp_a-gp_b):+,.0f}")
    k3.metric("الربح التشغيلي",     f"{op_a:,.0f} ر.س",   delta=f"{(op_a-op_b):+,.0f}")
    k4.metric("صافي الربح الفعلي",  f"{net_a:,.0f} ر.س",  delta=f"{(net_a-net_b):+,.0f}")

    # ── رسم بياني مقارنة ──
    st.markdown("#### 📊 رسم بياني: الفعلي vs التقديري")
    chart_labels = ["المبيعات", "إجمالي الربح", "الربح التشغيلي", "صافي الربح"]
    chart_actual = [rev_a, gp_a, op_a, net_a]
    chart_budget = [rev_b, gp_b, op_b, net_b]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="الفعلي",     x=chart_labels, y=chart_actual, marker_color="#0F2942",
                          text=[f"{v:,.0f}" for v in chart_actual], textposition="outside"))
    fig.add_trace(go.Bar(name="التقديري",   x=chart_labels, y=chart_budget, marker_color="#10B981",
                          text=[f"{v:,.0f}" for v in chart_budget], textposition="outside"))
    fig.update_layout(barmode="group", height=400, paper_bgcolor="#fff", plot_bgcolor="#F8FAFC",
                      legend=dict(orientation="h", y=-0.2),
                      yaxis=dict(tickformat=",.0f"))
    st.plotly_chart(fig, use_container_width=True)

    # ── رسم انحرافات ──
    st.markdown("#### 📉 رسم الانحرافات لكل بند")
    dev_labels = [lbl.split(" ", 1)[-1] for _, lbl, _ in bud_items]
    dev_vals   = [actual_map[code] - bud.get(code, actual_map[code]) for code, _, _ in bud_items]
    colors     = ["#10B981" if v >= 0 else "#EF4444" for v in dev_vals]
    fig2 = go.Figure(go.Bar(x=dev_labels, y=dev_vals, marker_color=colors,
                             text=[f"{v:+,.0f}" for v in dev_vals], textposition="outside"))
    fig2.update_layout(title="الانحراف (فعلي − تقديري) لكل بند", height=380,
                        paper_bgcolor="#fff", plot_bgcolor="#F8FAFC",
                        yaxis=dict(tickformat=",.0f"),
                        shapes=[dict(type="line", x0=-0.5, x1=len(dev_labels)-0.5, y0=0, y1=0,
                                     line=dict(color="gray", width=1, dash="dash"))])
    st.plotly_chart(fig2, use_container_width=True)


def _make_budget_excel_template(bud_items, report_year):
    """نموذج Excel فارغ لإدخال التقديري"""
    wb = Workbook()
    ws = wb.active; ws.title = "الميزانية التقديرية"
    ws.sheet_view.rightToLeft = True
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 24
    DARK = "1A3A5C"; W = "FFFFFF"
    bd = Border(left=Side(style="thin",color="CCCCCC"), right=Side(style="thin",color="CCCCCC"),
                top=Side(style="thin",color="CCCCCC"),  bottom=Side(style="thin",color="CCCCCC"))
    ws.merge_cells("A1:B1")
    ws["A1"] = f"نموذج الميزانية التقديرية — {report_year}"
    ws["A1"].font = Font(name="Arial", size=13, bold=True, color=W)
    ws["A1"].fill = PatternFill("solid", fgColor=DARK)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    for j, hdr in enumerate(["البند", "المبلغ التقديري (ر.س)"], 1):
        c = ws.cell(2, j, hdr)
        c.font = Font(name="Arial", size=11, bold=True, color=W)
        c.fill = PatternFill("solid", fgColor="10B981")
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = bd
    for i, (_, label, actual) in enumerate(bud_items, 3):
        clean_lbl = label.split(" ", 1)[-1]
        ws.cell(i, 1, clean_lbl).border = bd
        ws.cell(i, 1).font = Font(name="Arial", size=10)
        ws.cell(i, 1).alignment = Alignment(horizontal="right", vertical="center")
        c = ws.cell(i, 2, actual)
        c.border = bd; c.number_format = '#,##0.00'; c.fill = PatternFill("solid", fgColor="DCFCE7")
        c.font = Font(name="Arial", size=10, color="065F46")
        c.alignment = Alignment(horizontal="center", vertical="center")
        if i % 2 == 0: ws.cell(i, 1).fill = PatternFill("solid", fgColor="F8FAFC")
    out = io.BytesIO(); wb.save(out); return out.getvalue()


# ── واجهة التطبيق التشغيلية الرئيسية ──
def main():
    if not st.session_state.get("logged_in", False):
        show_login()
        return

    user = st.session_state.user
    st.markdown(f"""
    <div class="hbanner">
      <div>
        <h1>النظام المالي والزكوي الشامل</h1>
        <p>ZATCA-Compliant Financial Reporting System | الإصدار {APP_VERSION}</p>
      </div>
      <div class="hbanner-right">
        <div class="zatca-pill">ZATCA 1445هـ ✓</div>
        <div class="user-chip">
          <div class="user-avatar">{OWNER_INITIALS}</div>
          <div>
            <div class="user-name">تطوير: {OWNER_NAME_AR}</div>
            <div class="user-role">{OWNER_NAME_EN} | {OWNER_ROLE_EN}</div>
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### 🏢 إعدادات التقارير")
        company_name  = st.text_input("اسم الشركة", value="مؤسسة إمداد التغذية")
        company_type  = st.selectbox("نوع الشركة", list(COMPANY_TYPES.keys()))
        report_year   = st.number_input("السنة المالية", value=2025, step=1)
        zakat_adj     = st.number_input("التعديلات الزكوية (مادة 4)", value=380000.0)
        
        st.markdown("---")
        st.markdown("### 📥 ميزان المراجعة")
        uploaded = st.file_uploader("رفع ملف Excel", type=["xlsx"])
        st.download_button("⬇️ تحميل قالب ميزان المراجعة", data=create_tb_template(), file_name="قالب_ميزان_المراجعة.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.markdown("---")
        if st.button("🚪 تسجيل الخروج"):
            st.session_state.clear(); st.rerun()

    if not uploaded:
        st.info("👈 يرجى رفع ملف ميزان المراجعة المعبأ من الشريط الجانبي لعرض كافة التقارير.")
        return

    with st.spinner("🔄 جاري التحليل المتقدم للموازين..."):
        try:
            df = parse_tb(uploaded)
            r_curr = calc_financials(df, "net_curr", zakat_adj, company_type)
            r_prev = calc_financials(df, "net_prev", 0, company_type)
        except Exception as e:
            st.error(f"❌ خطأ أثناء المعالجة: {e}")
            return

    st.success(f"✅ تم سحب بيانات {company_name} بنجاح | تم مطابقة {len(df)} حساب بناءً على تصنيفك التفصيلي الشجري.")

    # 🟢 إظهار صف الكروت القياسية (KPIs) بعد القراءة الناجحة
    gp_margin = r_curr["gross_profit"] / r_curr["net_sales"] if r_curr["net_sales"] else 0
    np_margin = r_curr["net_az"] / r_curr["net_sales"] if r_curr["net_sales"] else 0
    kpis = [
        ("المبيعات الصافية", r_curr["net_sales"], "🏪"),
        ("مجمل الربح", r_curr["gross_profit"], "📈"),
        ("صافي الربح", r_curr["net_az"], "💰"),
        ("إجمالي الأصول", r_curr["total_assets"], "🏦"),
        ("الزكاة المستحقة", r_curr["zakat_due"], "🕌"),
        ("هامش الربح الصافي", np_margin, "📊"),
    ]
    cols = st.columns(6)
    for i, (lbl, val, icon) in enumerate(kpis):
        with cols[i]:
            disp = fmt_pct(val) if i == 5 else fmt(val)
            st.markdown(f'<div class="kpi"><div class="lbl">{icon} {lbl}</div><div class="val">{disp}</div></div>', unsafe_allow_html=True)

    # 🟢 إنشاء جميع صفحات التقارير ببياناتها الفعلية بالكامل
    tabs = st.tabs(["🏠 لوحة التحكم", "📑 قائمة الدخل", "🏦 المركز المالي", "💵 التدفقات النقدية", "📋 حقوق الملكية", "💸 تحليل المصروفات", "📐 النسب المالية", "🕌 تقرير الزكاة", "📊 مؤشرات", "🎯 الميزانية التقديرية", "📥 تصدير Excel"])
    
    with tabs[0]:
        show_dashboard(r_curr, r_prev, report_year, company_name)

    with tabs[1]:
        st.markdown(f'<div class="sec-title">قائمة الدخل الشامل | مقارنة {report_year} و {report_year-1}</div>', unsafe_allow_html=True)
        inc_data = {
            "البيـــان": ["إجمالي المبيعات", "صافي إيرادات المبيعات", "يخصم: تكلفة الإيرادات", "إجمالي الربح (Gross Profit)", "الرواتب والأجور", "الإيجارات", "الاستهلاك والإهلاك", "مصاريف إدارية أخرى", "إجمالي المصاريف التشغيلية", "الربح التشغيلي (EBIT)", "الربح قبل الزكاة", "الزكاة الشرعية", "صافي ربح الفترة"],
            f"{report_year} (ر.س)": [fmt(r_curr["gross_sales"]), fmt(r_curr["net_sales"]), fmt(-r_curr["cogs"]), fmt(r_curr["gross_profit"]), fmt(-r_curr["salaries"]), fmt(-r_curr["rent"]), fmt(-r_curr["depreciation"]), fmt(-r_curr["misc_exp"]), fmt(-r_curr["total_opex"]), fmt(r_curr["op_profit"]), fmt(r_curr["net_bz"]), fmt(-r_curr["zakat_due"]), fmt(r_curr["net_az"])],
            f"{report_year-1} (ر.س)": [fmt(r_prev["gross_sales"]), fmt(r_prev["net_sales"]), fmt(-r_prev["cogs"]), fmt(r_prev["gross_profit"]), fmt(-r_prev["salaries"]), fmt(-r_prev["rent"]), fmt(-r_prev["depreciation"]), fmt(-r_prev["misc_exp"]), fmt(-r_prev["total_opex"]), fmt(r_prev["op_profit"]), fmt(r_prev["net_bz"]), fmt(-r_prev["zakat_due"]), fmt(r_prev["net_az"])],
        }
        st.table(pd.DataFrame(inc_data))

    with tabs[2]:
        st.markdown(f'<div class="sec-title">قائمة المركز المالي | كما في 31 ديسمبر {report_year}</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**الأصول**")
            assets_data = {
                "البيـــان": ["الممتلكات والآلات (صافي)", "مجموع الأصول غير المتداولة", "نقد وما في حكمه", "مدينون تجاريون", "مخزون بضاعة", "أرصدة مدينة أخرى", "مجموع الأصول المتداولة", "مجموع الأصول"],
                "المبلغ (ر.س)": [fmt(r_curr["fixed_assets"]), fmt(r_curr["total_nca"]), fmt(r_curr["cash"]), fmt(r_curr["receivables"]), fmt(r_curr["inventory"]), fmt(r_curr["prepaid"]), fmt(r_curr["total_current"]), fmt(r_curr["total_assets"])]
            }
            st.table(pd.DataFrame(assets_data))
        with c2:
            st.markdown("**الالتزامات وحقوق الملكية**")
            eq_data = {
                "البيـــان": ["رأس المال", "الأرباح المبقاة", "مجموع حقوق الملكية", "دائنون تجاريون", "مصاريف مستحقة", "مخصص الزكاة", "مجموع الالتزامات المتداولة", "مجموع حقوق الملكية والالتزامات"],
                "المبلغ (ر.س)": [fmt(r_curr["capital"]), fmt(r_curr["retained"]), fmt(r_curr["equity"]), fmt(r_curr["payables"]), fmt(r_curr["accruals"]), fmt(r_curr["zakat_due"]), fmt(r_curr["total_cl"] + r_curr["zakat_due"]), fmt(r_curr["total_eq_lb"])]
            }
            st.table(pd.DataFrame(eq_data))
        diff = abs(r_curr["total_assets"] - r_curr["total_eq_lb"])
        if diff < 1: st.success("✅ الميزانية متوازنة تماماً")
        else: st.warning(f"⚠️ فرق في الميزانية: {diff:,.2f} ر.س")

    with tabs[3]:
        show_cash_flow(r_curr, r_prev, report_year)

    with tabs[4]:
        show_equity_changes(r_curr, r_prev, report_year)

    with tabs[5]:
        show_auto_expenses(df, report_year)

    with tabs[6]:
        show_financial_ratios(r_curr, r_prev, report_year)

    with tabs[7]:
        show_zakat_report(r_curr, zakat_adj, report_year)

    with tabs[8]:
        st.markdown('<div class="sec-title">📊 المؤشرات المالية والرسم البياني</div>', unsafe_allow_html=True)
        ratios = {
            "هامش إجمالي الربح": fmt_pct(r_curr["gross_profit"]/r_curr["net_sales"]) if r_curr["net_sales"] else "–",
            "هامش الربح التشغيلي": fmt_pct(r_curr["op_profit"]/r_curr["net_sales"]) if r_curr["net_sales"] else "–",
            "العائد على الأصول (ROA)": fmt_pct(r_curr["net_az"]/r_curr["total_assets"]) if r_curr["total_assets"] else "–",
            "نسبة التداول": f"{r_curr['total_current']/r_curr['total_cl']:.2f}x" if r_curr["total_cl"] else "–",
        }
        st.table(pd.DataFrame({"المؤشر": list(ratios.keys()), "القيمة": list(ratios.values())}))
        fig = go.Figure()
        cats = ["المبيعات", "مجمل الربح", "الربح التشغيلي", "صافي الربح"]
        vals_c = [r_curr["net_sales"], r_curr["gross_profit"], r_curr["op_profit"], r_curr["net_az"]]
        vals_p = [r_prev["net_sales"], r_prev["gross_profit"], r_prev["op_profit"], r_prev["net_az"]]
        fig.add_trace(go.Bar(name=str(report_year), x=cats, y=vals_c, marker_color="#10B981"))
        fig.add_trace(go.Bar(name=str(report_year-1), x=cats, y=vals_p, marker_color="#0F2942"))
        fig.update_layout(barmode="group", title="مقارنة الأداء المالي", height=380, paper_bgcolor="#fff", plot_bgcolor="#F8FAFC")
        st.plotly_chart(fig, use_container_width=True)

    with tabs[9]:
        show_budget_tab(r_curr, report_year)

    with tabs[10]:
        st.markdown('<div class="sec-title">📥 تصدير القوائم المالية الرسمية</div>', unsafe_allow_html=True)
        excel_bytes = export_excel(r_curr, r_prev, df, company_name, company_type, report_year)
        st.download_button("⬇️ تحميل التقرير المالي الكامل (Excel)", data=excel_bytes, file_name=f"القوائم_المالية_{report_year}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

    st.markdown(f"""
    <div class="footer-wm">
      <div class="ft-left">
        <div class="ft-av">{OWNER_INITIALS}</div>
        <div class="ft-name">أعدّه: <span>{OWNER_NAME_EN} — {OWNER_NAME_AR}</span></div>
      </div>
      <div class="ft-right">النظام المالي والزكوي {APP_VERSION} · ZATCA 1445هـ</div>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
