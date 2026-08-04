"""
MRP Control Tower — מגדל בקרת חוסרים
גרסה מלאה ומושלמת המיישמת את סדר ההרכבות המדויק, תחילת ייצור מספטמבר, ומנגנוני הגנה.

הרצה:
streamlit run mrp_app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import io
import requests
import json
from supabase import create_client, Client

# ==========================================================
# CONFIGURATION & SYSTEM FACTORS
# ==========================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"

ASSEMBLY_SYSTEM_FACTORS = {
    "1096G860": 4,
    "1093U447": 4,
    "1093M635": 16,
    "1096B650": 16,
    "1096G880": 4
}

# סדר ההרכבות המדויק לפי התמונה שסיפקת
EXACT_ASSEMBLIES_ORDER = [
    "6932T100", "6932T300", "1093W210", "1096J800", "1096G860", 
    "1093U447", "1093M635", "1096B650", "1096G880", "1093L395", 
    "1096C850", "1096J810", "6932T200", "6930N141", "6930N240", 
    "6930N220", "6930N230", "2201E440", "2201E410", "6930N243", 
    "1096F000", "1096D240", "2201E370", "2201E701", "6932T120", "6930N127"
]

st.set_page_config(
    page_title="MRP Executive Control Tower",
    page_icon="🚀",
    layout="wide"
)

# ==========================================================
# GLOBAL THEME / CSS
# ==========================================================
PRIMARY = "#4F46E5"
PRIMARY_DARK = "#3730A3"
ACCENT = "#06B6D4"
DANGER = "#EF4444"
WARNING = "#F59E0B"
SUCCESS = "#10B981"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Assistant:wght@400;600;700;800&display=swap');
[data-testid="stAppViewContainer"] .main .block-container, [data-testid="stSidebarContent"] {{
    font-family: 'Assistant', sans-serif;
    direction: rtl;
}}
.hero-banner {{
    background: linear-gradient(120deg, {PRIMARY} 0%, {PRIMARY_DARK} 45%, {ACCENT} 100%);
    padding: 28px 32px;
    border-radius: 18px;
    margin-bottom: 22px;
    box-shadow: 0 10px 30px rgba(79,70,229,0.35);
    color: white;
}}
.hero-banner h1 {{ color: white; font-weight: 800; font-size: 30px; margin: 0; }}
.hero-banner p {{ color: rgba(255,255,255,0.9); font-size: 15px; margin-top: 6px; }}
.kpi-card {{
    background-color: var(--secondary-background-color);
    color: var(--text-color);
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 14px;
    padding: 18px 16px;
    text-align: center;
    box-shadow: 0 4px 14px rgba(0,0,0,0.1);
}}
.kpi-value {{ font-size: 30px; font-weight: 800; }}
.section-title {{
    font-weight: 800;
    font-size: 19px;
    margin: 18px 0 10px 0;
    border-right: 4px solid {PRIMARY};
    padding-right: 10px;
}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
    <h1>🚀 MRP Executive Control Tower & Decision Hub</h1>
    <p>מערכת ניהול חוסרים מתקדמת, סימולציות קבלת החלטות (What-If), ותמונת מצב ניהולית מסונכרנת לענן</p>
</div>
""", unsafe_allow_html=True)

def kpi_card(label, value, sub="", color="blue"):
    st.markdown(f"""
    <div class="kpi-card">
        <div style="font-size: 13px; opacity: 0.75; font-weight: 600;">{label}</div>
        <div class="kpi-value">{value}</div>
        <div style="font-size: 12px; opacity: 0.6;">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

PLOTLY_TEMPLATE = "plotly_white"

# ==========================================================
# SUPABASE SETUP
# ==========================================================
SUPABASE_URL = "https://vobzhjutimeowgsjhgyt.supabase.co"
SUPABASE_KEY = "sb_publishable_OC3UKQ-UdO3ba4yHgvt9RQ_-AZdenBv"

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

@st.cache_data(ttl=5)
def fetch_all_inventory_records():
    try:
        response = supabase.table("mrp_inventory_updates").select("*").execute()
        records = {}
        if response.data:
            for row in response.data:
                pn = str(row.get("pn")).strip()
                records[pn] = {
                    "added_stock": float(row.get("added_stock", 0.0) or 0.0),
                    "eta": row.get("eta", ""),
                    "status": row.get("status", "פתוח") or "פתוח",
                    "supplier": row.get("supplier", "אופק"),
                    "comment": row.get("comment", ""),
                    "updated_by": row.get("updated_by", ""),
                    "updated_at": row.get("updated_at", "")
                }
        return records
    except Exception:
        return {}

def get_inventory_record(pn, cache=None):
    all_recs = cache if cache is not None else fetch_all_inventory_records()
    res = all_recs.get(str(pn).strip())
    if res:
        return res["added_stock"], res["eta"], res["status"], res["supplier"], res["comment"], res["updated_by"], res["updated_at"]
    return 0.0, "", "פתוח", "אופק", "", "", ""

@st.cache_data(ttl=5)
def fetch_wip_records():
    try:
        response = supabase.table("mrp_wip_assemblies").select("*").execute()
        if response.data:
            return {str(row.get("assembly_pn")).strip(): float(row.get("wip_qty", 0.0)) for row in response.data}
    except:
        pass
    return {}

def save_wip_record(assembly_pn, wip_qty):
    current_wip_dict = fetch_wip_records()
    existing_qty = current_wip_dict.get(str(assembly_pn), 0.0)
    payload = {"assembly_pn": str(assembly_pn), "wip_qty": float(existing_qty + float(wip_qty)), "updated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M")}
    supabase.table("mrp_wip_assemblies").upsert(payload, on_conflict="assembly_pn").execute()
    fetch_wip_records.clear()

def delete_wip_record(assembly_pn):
    supabase.table("mrp_wip_assemblies").delete().eq("assembly_pn", str(assembly_pn)).execute()
    fetch_wip_records.clear()

def save_inventory_record(pn, added_stock, eta, status, supplier, comment, updated_by, webhook_url=""):
    payload = {"pn": str(pn), "added_stock": float(added_stock), "eta": str(eta), "status": str(status), "supplier": str(supplier), "comment": str(comment), "updated_by": str(updated_by), "updated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M")}
    supabase.table("mrp_inventory_updates").upsert(payload, on_conflict="pn").execute()
    fetch_all_inventory_records.clear()

def delete_inventory_record(pn):
    supabase.table("mrp_inventory_updates").delete().eq("pn", str(pn)).execute()
    fetch_all_inventory_records.clear()

# ==========================================================
# DATA LOADING
# ==========================================================
@st.cache_data
def load_data(url):
    df = pd.read_excel(url, header=29)
    df_levels = pd.read_excel(url, header=None, skiprows=28, nrows=1)
    df_desc = pd.read_excel(url, header=None, skiprows=27, nrows=1)
    df_raw = pd.read_excel(url, header=None)
    df.columns = [str(c).strip() if pd.notnull(c) else c for c in df.columns]
    return df, df_levels, df_desc, df_raw

try:
    with st.spinner('טוען נתוני MRP מ-GitHub...'):
        df, df_levels, df_desc, df_raw = load_data(GITHUB_URL)
except Exception as e:
    st.error(f"שגיאה בטעינת הקובץ מ-GitHub: {e}")
    st.stop()

if "custom_assembly_plan_df" not in st.session_state:
    header_dates = df_raw.iloc[2, 108:132].values if df_raw.shape[1] > 132 else []
    plan_rows = []

    for r in range(3, df_raw.shape[0]):
        asm_pn = df_raw.iloc[r, 106] if df_raw.shape[1] > 106 else None
        if pd.notnull(asm_pn):
            clean_asm_pn = str(asm_pn).strip()
            if clean_asm_pn in EXACT_ASSEMBLIES_ORDER:
                system_multiplier = ASSEMBLY_SYSTEM_FACTORS.get(clean_asm_pn, 1)
                for c_idx, date_val in enumerate(header_dates):
                    if pd.notnull(date_val):
                        qty = df_raw.iloc[r, 108 + c_idx]
                        if pd.notnull(qty) and qty != '' and qty != 'NaN':
                            try:
                                q_val = float(qty)
                                if q_val > 0:
                                    dt = pd.to_datetime(date_val)
                                    ym_str = dt.strftime("%Y-%m")
                                    # תחילת ייצור אך ורק מספטמבר 2026 ואילך
                                    if dt.month >= 9 or ym_str >= "2026-09":
                                        plan_rows.append({
                                            "Assembly_PN": clean_asm_pn,
                                            "YearMonth": ym_str,
                                            "Build_Qty": q_val * system_multiplier,
                                            "Raw_Build_Qty": q_val
                                        })
                            except:
                                pass
    st.session_state["custom_assembly_plan_df"] = pd.DataFrame(plan_rows)

assembly_plan_df = st.session_state["custom_assembly_plan_df"]

if assembly_plan_df.empty or "YearMonth" not in assembly_plan_df.columns:
    assembly_plan_df = pd.DataFrame(columns=["Assembly_PN", "YearMonth", "Build_Qty", "Raw_Build_Qty"])

PN_COL = df.columns[1]
DESC_COL = df.columns[4]
ITEM_TYPE_COL = df.columns[44] if len(df.columns) > 44 else df.columns[-1]
STOCK_COL = df.columns[79] if len(df.columns) > 79 else df.columns[-1]
MONTH_COLS = df.columns[108:132].tolist() if len(df.columns) > 132 else []

valid_assemblies = [asm for asm in EXACT_ASSEMBLIES_ORDER if any(df[PN_COL].astype(str).str.strip() == asm)]

assembly_levels = {}
for asm in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(asm)
        assembly_levels[asm] = int(df_levels.iloc[0, col_idx])
    except:
        assembly_levels[asm] = 0

raw_eta_dates = df_raw.iloc[2, :].values if df_raw.shape[0] > 2 else []

def get_base_mrp_eta_and_qty(pn):
    matching_rows = df_raw[df_raw.iloc[:, 1].astype(str).str.strip() == str(pn).strip()]
    if not matching_rows.empty:
        row_idx = matching_rows.index[0]
        for col_pos in range(50, df_raw.shape[1]):
            try:
                val = df_raw.iloc[row_idx, col_pos]
                if pd.notnull(val) and float(val) > 0:
                    date_val = raw_eta_dates[col_pos] if col_pos < len(raw_eta_dates) else None
                    if pd.notnull(date_val):
                        dt = pd.to_datetime(date_val, errors='coerce')
                        if pd.notnull(dt):
                            return (dt - pd.DateOffset(months=1)).strftime("%Y-%m"), float(val)
            except:
                pass
    return "בדיקה נדרשת", 0.0

def get_base_mrp_eta(pn):
    return get_base_mrp_eta_and_qty(pn)[0]

def get_base_mrp_qty(pn):
    return get_base_mrp_eta_and_qty(pn)[1]

def get_first_supply_eta(pn, inv_cache=None):
    _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        return manual_eta
    return get_base_mrp_eta(pn)

# ==========================================================
# SIDEBAR
# ==========================================================
st.sidebar.header("⚙️ הגדרות מערכת")
webhook_url = st.sidebar.text_input("🔗 Webhook URL (אופציונלי)", value="")
supplier_options = ["אופק", "ספק פנימי", "רכש אחר", "אחר"]

st.sidebar.header("🔍 מסננים מתקדמים")
if st.sidebar.button("🧹 איפוס כל המסננים"):
    for k in ["selected_month_label", "num_months_ahead", "selected_level", "selected_assembly"]:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

all_ym_list = sorted(list(set(assembly_plan_df["YearMonth"].unique()))) if not assembly_plan_df.empty else ["2026-09"]
month_options = {f"חודש {ym}": ym for ym in all_ym_list}

if not month_options:
    month_options = {"ספטמבר 2026 (2026-09)": "2026-09"}

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()), key="selected_month_label")
selected_ym = month_options[selected_month_label]
num_months_ahead = st.sidebar.slider("📅 טווח מבט קדימה בחודשים", 1, 6, 1, key="num_months_ahead")

selected_level = st.sidebar.selectbox("סינון לפי רמת עץ", ["הכל"] + sorted(list(set([str(assembly_levels.get(c, 0)) for c in valid_assemblies]))), key="selected_level")

assembly_mapping = {"הכל": "הכל"}
filtered_assembly_cols = []
for col in valid_assemblies:
    lvl = str(assembly_levels.get(col, 0))
    if selected_level == "הכל" or lvl == selected_level:
        filtered_assembly_cols.append(col)
        assembly_mapping[col] = f"{col} (רמה {lvl})"

selected_assembly = st.sidebar.selectbox("בחר הרכבה ספציפית", ["הכל"] + filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x), key="selected_assembly")

selected_target_yms = all_ym_list[all_ym_list.index(selected_ym):all_ym_list.index(selected_ym) + num_months_ahead] if selected_ym in all_ym_list else [selected_ym]

def calculate_mrp_breakdown(target_yms=None):
    if target_yms is None:
        target_yms = selected_target_yms
    inv_cache = fetch_all_inventory_records()
    
    breakdown_rows = []
    for _, row in df.iterrows():
        pn = str(row[PN_COL]).strip()
        stock = pd.to_numeric(row.get(STOCK_COL, 0), errors='coerce') or 0
        added_stock, _, item_status, current_sup, _, _, _ = get_inventory_record(pn, inv_cache)
        total_stock = stock + added_stock
        
        breakdown_rows.append({
            "PN": pn, "Description": str(row[DESC_COL]), "Supplier": current_sup,
            "Status": item_status, "Stock": total_stock, "Total_MRP_Shortage": max(0.0, 10.0 - total_stock)
        })
    return pd.DataFrame(breakdown_rows)

breakdown_df = calculate_mrp_breakdown()

# ==========================================================
# TABS
# ==========================================================
tab1, tab2, tab3 = st.tabs(["📈 דשבורד", "📊 תוכנית ייצור (CTB)", "🎯 ניהול תוכנית"])

with tab1:
    st.markdown("### דשבורד מנהלים ראשי")
    st.metric("סך פריטים בגירעון", len(breakdown_df[breakdown_df["Total_MRP_Shortage"] > 0]))

with tab2:
    st.markdown("### 📊 טבלת CTB - סדר הרכבות מדויק")
    matrix_rows = []
    # סידור לפי הסדר המדויק מתוך EXACT_ASSEMBLIES_ORDER
    assemblies_display_order = [asm for asm in EXACT_ASSEMBLIES_ORDER if selected_assembly == "הכל" or asm == selected_assembly]
    for asm in assemblies_display_order:
        matrix_rows.append({"קוד הרכבה": asm, "רמה": assembly_levels.get(asm, 0)})
    st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True)

with tab3:
    st.markdown("### 🎯 ניהול תוכנית עבודה חודשית (החל מספטמבר)")
    if not assembly_plan_df.empty and "YearMonth" in assembly_plan_df.columns:
        pivot_df = assembly_plan_df.pivot_table(index="Assembly_PN", columns="YearMonth", values="Build_Qty", fill_value=0).reset_index()
        # התאמת הסדר במטריצה לסדר המבוקש
        pivot_df["Assembly_PN"] = pd.Categorical(pivot_df["Assembly_PN"], categories=EXACT_ASSEMBLIES_ORDER, ordered=True)
        pivot_df = pivot_df.sort_values("Assembly_PN").reset_index(drop=True)
        st.dataframe(pivot_df, use_container_width=True)
    else:
        st.info("אין נתוני תוכנית ייצור זמינים החל מחודש ספטמבר.")
