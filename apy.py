"""
MRP Control Tower — מגדל בקרת חוסרים (גרסה מלאה עם ניתוח רגישות פרטני לפי הרכבה וחודש)
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
    "1096G860-002": 4,
    "1093U447-001": 4,
    "1093M635-003": 16,
    "1096B650-003": 16,
    "1096G880-003": 4
}

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

[data-testid="stAppViewContainer"] .main .block-container,
[data-testid="stSidebarContent"] {{
    font-family: 'Assistant', sans-serif;
    direction: rtl;
}}
[data-testid="stAppViewContainer"] .main .block-container * ,
[data-testid="stSidebarContent"] * {{
    font-family: 'Assistant', sans-serif;
}}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

.hero-banner {{
    background: linear-gradient(120deg, {PRIMARY} 0%, {PRIMARY_DARK} 45%, {ACCENT} 100%);
    padding: 28px 32px;
    border-radius: 18px;
    margin-bottom: 22px;
    box-shadow: 0 10px 30px rgba(79,70,229,0.35);
}}
.hero-banner h1 {{
    color: white;
    font-weight: 800;
    font-size: 30px;
    margin: 0;
}}
.hero-banner p {{
    color: rgba(255,255,255,0.9);
    font-size: 15px;
    margin-top: 6px;
}}

.kpi-card {{
    background-color: var(--secondary-background-color);
    color: var(--text-color);
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 14px;
    padding: 18px 16px;
    text-align: center;
    box-shadow: 0 4px 14px rgba(0,0,0,0.1);
    transition: transform 0.15s ease;
}}
.kpi-card:hover {{ transform: translateY(-3px); }}
.kpi-label {{
    font-size: 13px;
    opacity: 0.75;
    margin-bottom: 6px;
    font-weight: 600;
}}
.kpi-value {{
    font-size: 30px;
    font-weight: 800;
}}
.kpi-sub {{
    font-size: 12px;
    opacity: 0.6;
    margin-top: 4px;
}}

.kpi-green {{ border-top: 4px solid {SUCCESS}; }}
.kpi-red {{ border-top: 4px solid {DANGER}; }}
.kpi-orange {{ border-top: 4px solid {WARNING}; }}
.kpi-blue {{ border-top: 4px solid {ACCENT}; }}

@media (prefers-color-scheme: light) {{
    .kpi-card, .kanban-card {{
        background-color: #ffffff !important;
        color: #111827 !important;
        border-color: #e5e7eb !important;
    }}
}}

@media (prefers-color-scheme: dark) {{
    .kpi-card, .kanban-card {{
        background-color: #1f2937 !important;
        color: #f9fafb !important;
        border-color: #374151 !important;
    }}
}}

.section-title {{
    font-weight: 800;
    font-size: 19px;
    margin: 18px 0 10px 0;
    border-right: 4px solid {PRIMARY};
    padding-right: 10px;
    color: var(--text-color, inherit);
}}

.kanban-col-header {{
    font-weight: 800;
    font-size: 15px;
    padding: 8px 12px;
    border-radius: 10px;
    margin-bottom: 8px;
    text-align: center;
}}
.kanban-card {{
    border: 1px solid rgba(128,128,128,0.2);
    border-radius: 10px;
    padding: 10px 12px;
    margin-bottom: 8px;
    border-right: 3px solid {PRIMARY};
    font-size: 13px;
}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
    <h1>🚀 MRP Executive Control Tower & Decision Hub</h1>
    <p>מערכת ניהול חוסרים מתקדמת, סימולציות קבלת החלטות (What-If), ותמונת מצב ניהולית מסונכרנת לענן במהירות שיא</p>
</div>
""", unsafe_allow_html=True)

def kpi_card(label, value, sub="", color="blue"):
    st.markdown(f"""
    <div class="kpi-card kpi-{color}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

try:
    _theme_base = st.get_option("theme.base")
except Exception:
    _theme_base = None

PLOTLY_TEMPLATE = "plotly_white" if _theme_base == "light" else "plotly_dark"
COLOR_SEQ = [PRIMARY, ACCENT, WARNING, DANGER, SUCCESS, "#A78BFA", "#F472B6", "#34D399"]

# ==========================================================
# SUPABASE SETUP & FAST CACHED STORAGE
# ==========================================================
SUPABASE_URL = "https://vobzhjutimeowgsjhgyt.supabase.co"
SUPABASE_KEY = "sb_publishable_OC3UKQ-UdO3ba4yHgvt9RQ_-AZdenBv"

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

@st.cache_data(ttl=60)
def fetch_all_inventory_records():
    try:
        response = supabase.table("mrp_inventory_updates").select("*").execute()
        records = {}
        if response.data:
            for row in response.data:
                pn = str(row.get("pn")).strip()
                eta_val = row.get("eta", "")
                if not eta_val or str(eta_val).strip() in ["", "None", "NaT", "nan"]:
                    eta_val = ""
                status_val = row.get("status", "פתוח") or "פתוח"
                records[pn] = {
                    "added_stock": float(row.get("added_stock", 0.0) or 0.0),
                    "eta": eta_val,
                    "status": status_val,
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
        return (
            res["added_stock"],
            res["eta"],
            res["status"],
            res["supplier"],
            res["comment"],
            res["updated_by"],
            res["updated_at"]
        )
    return 0.0, "", "פתוח", "אופק", "", "", ""

@st.cache_data(ttl=60)
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
    total_new_qty = existing_qty + float(wip_qty)

    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {
        "assembly_pn": str(assembly_pn),
        "wip_qty": float(total_new_qty),
        "updated_at": now_str
    }
    try:
        supabase.table("mrp_wip_assemblies").upsert(payload, on_conflict="assembly_pn").execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירת WIP ל-Supabase: {e}")

def delete_wip_record(assembly_pn):
    try:
        supabase.table("mrp_wip_assemblies").delete().eq("assembly_pn", str(assembly_pn)).execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקת WIP מ-Supabase: {e}")

def save_inventory_record(pn, added_stock, eta, status, supplier, comment, updated_by, webhook_url=""):
    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {
        "pn": str(pn),
        "added_stock": float(added_stock),
        "eta": str(eta),
        "status": str(status),
        "supplier": str(supplier),
        "comment": str(comment),
        "updated_by": str(updated_by),
        "updated_at": now_str
    }
    try:
        supabase.table("mrp_inventory_updates").upsert(payload, on_conflict="pn").execute()
        supabase.table("mrp_inventory_history").insert(payload).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירה ל-Supabase: {e}")

    if webhook_url:
        msg = "🔔 עדכון מלאי/ETA למוצר!\nמק'ט: " + str(pn) + "\nתוספת מלאי: " + str(added_stock) + "\nסטטוס: " + str(status) + "\nETA: " + str(eta)
        try:
            requests.post(webhook_url, data=json.dumps({"text": msg}), headers={'Content-Type': 'application/json'})
        except:
            pass

def delete_inventory_record(pn):
    try:
        supabase.table("mrp_inventory_updates").delete().eq("pn", str(pn)).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקה מ-Supabase: {e}")

# ==========================================================
# DATA LOADING FROM GITHUB & SESSION STATE
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
    st.error(f"שגיאה בטעינת הקובץ מ-GitHub. פירוט השגיאה: {e}")
    st.stop()

PN_COL = df.columns[1]
DESC_COL = df.columns[4]
ITEM_TYPE_COL = df.columns[44] if len(df.columns) > 44 else df.columns[-1]
STOCK_COL = df.columns[79] if len(df.columns) > 79 else df.columns[-1]
ASSEMBLY_COLS = df.columns[10:36].tolist()
MONTH_COLS = df.columns[108:132].tolist() if len(df.columns) > 132 else []

valid_assemblies = []
for col in ASSEMBLY_COLS:
    valid_assemblies.append(col)

assembly_levels = {}
for col in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(col)
        lvl_val = int(df_levels.iloc[0, col_idx])
        assembly_levels[col] = lvl_val
    except:
        assembly_levels[col] = 0

valid_assemblies = sorted(valid_assemblies, key=lambda x: (assembly_levels.get(x, 0), str(x)))

if "custom_assembly_plan_df" not in st.session_state:
    header_dates = df_raw.iloc[2, 108:132].values if df_raw.shape[1] > 132 else []
    plan_rows = []

    for r in range(3, df_raw.shape[0]):
        asm_pn = df_raw.iloc[r, 106] if df_raw.shape[1] > 106 else None
        if pd.notnull(asm_pn):
            clean_asm_pn = str(asm_pn).strip()
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
                                if dt.month >= 9 or ym_str >= "2026-09":
                                    displayed_build_qty = q_val * system_multiplier
                                    plan_rows.append({
                                        "Assembly_PN": clean_asm_pn,
                                        "YearMonth": ym_str,
                                        "Build_Qty": displayed_build_qty,
                                        "Raw_Build_Qty": q_val
                                    })
                        except:
                            pass
    st.session_state["custom_assembly_plan_df"] = pd.DataFrame(plan_rows)

assembly_plan_df = st.session_state["custom_assembly_plan_df"]
raw_eta_dates = df_raw.iloc[2, :].values if df_raw.shape[0] > 2 else []

def get_base_mrp_eta_and_qty(pn):
    matching_rows = df_raw[df_raw.iloc[:, 1].astype(str).str.strip() == str(pn).strip()]
    if not matching_rows.empty:
        row_idx = matching_rows.index[0]
        max_cols = df_raw.shape[1]

        for col_pos in range(50, max_cols):
            try:
                val = df_raw.iloc[row_idx, col_pos]
                if pd.notnull(val) and val != '' and val != 'NaN':
                    q = float(val)
                    if q > 0:
                        date_val = raw_eta_dates[col_pos] if col_pos < len(raw_eta_dates) else None
                        if pd.notnull(date_val):
                            if isinstance(date_val, datetime):
                                corrected_dt = date_val - pd.DateOffset(months=1)
                                return corrected_dt.strftime("%Y-%m"), q
                            dt = pd.to_datetime(date_val, errors='coerce', dayfirst=False)
                            if pd.notnull(dt) and dt.year >= 2024:
                                corrected_dt = dt - pd.DateOffset(months=1)
                                return corrected_dt.strftime("%Y-%m"), q
            except:
                pass
    return "בדיקה נדרשת", 0.0

def get_base_mrp_eta(pn):
    eta, _ = get_base_mrp_eta_and_qty(pn)
    return eta

def get_base_mrp_qty(pn):
    _, qty = get_base_mrp_eta_and_qty(pn)
    return qty

def get_first_supply_eta(pn, inv_cache=None):
    _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        return manual_eta
    return get_base_mrp_eta(pn)

# ==========================================================
# SIDEBAR FILTERS & WHAT-IF CONTROLS
# ==========================================================
st.sidebar.header("⚙️ הגדרות מערכת וחיבור")
webhook_url = st.sidebar.text_input("🔗 Teams / Slack Webhook URL (אופציונלי)", value="")
supplier_options = ["אופק", "ספק פנימי", "רכש אחר", "אחר"]

st.sidebar.header("🔍 מסננים מתקדמים")

if st.sidebar.button("🧹 איפוס כל המסננים (Clear All)"):
    keys_to_clear = ["selected_month_label", "num_months_ahead", "selected_level", "selected_assembly", "selected_item_type", "selected_search_item"]
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

month_options = {}
for m in MONTH_COLS:
    if pd.notnull(m):
        try:
            dt = pd.to_datetime(m)
            m_ym = dt.strftime("%Y-%m")
            if m_ym >= "2026-09":
                month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
        except:
            pass

if not month_options:
    for m in MONTH_COLS:
        if pd.notnull(m):
            try:
                dt = pd.to_datetime(m)
                if dt.month >= 9 or dt.strftime("%Y-%m") >= "2026-09":
                    month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
            except:
                pass

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()), key="selected_month_label")
selected_month_col = month_options[selected_month_label]

try:
    selected_ym = pd.to_datetime(selected_month_col).strftime("%Y-%m")
except:
    selected_ym = str(selected_month_col)[:7]

num_months_ahead = st.sidebar.slider("📅 טווח מבט קדימה במספר חודשים", min_value=1, max_value=6, value=1, key="num_months_ahead")

level_options = ["הכל"] + sorted(list(set(str(assembly_levels[c]) for c in valid_assemblies)), key=lambda x: int(x) if x.isdigit() else 0)
selected_level = st.sidebar.selectbox("סינון לפי רמת עץ (BOM Level)", level_options, key="selected_level")

assembly_mapping = {"הכל": "הכל"}
filtered_assembly_cols = []
for col in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(col)
        lvl = str(assembly_levels.get(col, 0))
        desc = df_desc.iloc[0, col_idx]
        if selected_level == "הכל" or lvl == selected_level:
            filtered_assembly_cols.append(col)
            assembly_mapping[col] = f"[רמה {lvl}] {str(col)} - {str(desc)}"
    except:
        filtered_assembly_cols.append(col)
        assembly_mapping[col] = col

selected_assembly = st.sidebar.selectbox(
    "בחר הרכבה ספציפית לדשבורד",
    ["הכל"] + filtered_assembly_cols,
    format_func=lambda x: assembly_mapping.get(x, x),
    key="selected_assembly"
)

item_types = df[ITEM_TYPE_COL].dropna().unique().tolist() if ITEM_TYPE_COL in df.columns else []
selected_item_type = st.sidebar.selectbox("בחר סוג פריט (עמודה AS)", ["הכל"] + item_types, key="selected_item_type")

item_choices = ["הכל"] + sorted([f"{str(r[PN_COL]).strip()} - {str(r[DESC_COL])}" for _, r in df.iterrows() if pd.notnull(r[PN_COL])])
selected_search_item = st.sidebar.selectbox("🔎 חיפוש מהיר (בחר או הקלד מק'ט/תיאור)", item_choices, key="selected_search_item")
search_pn = selected_search_item.split(" - ")[0] if selected_search_item != "הכל" else "הכל"

# ==========================================================
# FILE UPLOADS & TEMPLATES
# ==========================================================
st.sidebar.divider()
st.sidebar.markdown("##### 📥 עדכון ETA וכמות אספקה מקובץ ספק")

eta_template_df = pd.DataFrame(columns=["PN", "ETA", "Qty"])
eta_template_output = io.BytesIO()
with pd.ExcelWriter(eta_template_output, engine='openpyxl') as writer:
    eta_template_df.to_excel(writer, index=False, sheet_name='ETA_Template')
st.sidebar.download_button(
    label="📄 הורד תבנית Excel לעדכון ETA",
    data=eta_template_output.getvalue(),
    file_name="ETA_Update_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded_eta_file = st.sidebar.file_uploader("העלה קובץ ETA (עמודות: PN, ETA, Qty)", type=["xlsx", "xls"], key="eta_uploader")
if uploaded_eta_file is not None:
    try:
        eta_df_sup = pd.read_excel(uploaded_eta_file)
        if st.sidebar.button("⚡ עדכן ETA וכמות אספקה"):
            eta_count = 0
            for _, s_row in eta_df_sup.iterrows():
                p_code = str(s_row.iloc[0]).strip()
                new_eta = str(s_row.iloc[1]).strip() if len(s_row) > 1 and pd.notnull(s_row.iloc[1]) else ""
                new_supply_qty = float(s_row.iloc[2]) if len(s_row) > 2 and pd.notnull(s_row.iloc[2]) else 0.0

                if p_code and p_code != 'nan' and new_eta and new_eta not in ["nan", "NaT", "None"]:
                    curr_stock, _, curr_status, curr_sup, curr_comm, _, _ = get_inventory_record(p_code)
                    updated_total_stock = curr_stock + new_supply_qty if new_supply_qty > 0 else curr_stock

                    save_inventory_record(
                        pn=p_code,
                        added_stock=updated_total_stock,
                        eta=new_eta,
                        status=curr_status if curr_status != "פתוח" else "הוזמן",
                        supplier=curr_sup,
                        comment=f"{curr_comm} | אספקה בסך {new_supply_qty} בתאריך ETA {new_eta} מקובץ ספק",
                        updated_by="ETA & Qty File Upload",
                        webhook_url=webhook_url
                    )
                    eta_count += 1
            st.sidebar.success(f"עודכנו בהצלחה ETA וכמויות אספקה עבור {eta_count} שורות!")
    except Exception as e:
        st.sidebar.error(f"שגיאה בקריאת קובץ ה-ETA: {e}")

# ==========================================================
# OPTIMIZED SHORTAGE CALCULATION (CACHED & FIXED)
# ==========================================================
all_ym_list = sorted(list(set(assembly_plan_df["YearMonth"].unique())))
start_idx = 0
for idx, ym in enumerate(all_ym_list):
    if ym >= selected_ym:
        start_idx = idx
        break
selected_target_yms = all_ym_list[start_idx:start_idx + num_months_ahead]
if not selected_target_yms:
    selected_target_yms = [selected_ym]

@st.cache_data(ttl=60)
def calculate_mrp_breakdown_cached(target_yms_tuple, sim_extra_stock_items_tuple, active_plan_df):
    sim_extra_stock_dict = dict(sim_extra_stock_items_tuple)
    inv_cache = fetch_all_inventory_records()
    wip_cache = fetch_wip_records()

    target_month_cols_map = {}
    for m_c in MONTH_COLS:
        if pd.notnull(m_c):
            try:
                m_dt_ym = pd.to_datetime(m_c).strftime("%Y-%m")
                if m_dt_ym in target_yms_tuple:
                    target_month_cols_map[m_dt_ym] = m_c
            except:
                pass

    temp_df = df.copy()
    shortage_records = {}

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        sim_val = sim_extra_stock_dict.get(pn, 0.0)
        total_added_stock = saved_stock_add + sim_val

        max_shortage_val = 0.0
        is_short_or = False

        for ym in target_yms_tuple:
            col_name = target_month_cols_map.get(ym)
            if col_name and col_name in temp_df.columns:
                mrp_val = pd.to_numeric(row[col_name], errors='coerce') or 0
                effective_mrp_val = mrp_val + total_added_stock if mrp_val < 0 else mrp_val

                if effective_mrp_val < 0:
                    is_short_or = True
                    sh_qty = abs(effective_mrp_val)
                    if sh_qty > max_shortage_val:
                        max_shortage_val = sh_qty

        if is_short_or:
            shortage_records[idx] = max_shortage_val

    temp_df['Monthly_Balance'] = temp_df.index.map(lambda i: -shortage_records[i] if i in shortage_records else 1.0)

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
            try:
                eta_ym = pd.to_datetime(manual_eta).strftime("%Y-%m")
                if eta_ym > max(target_yms_tuple):
                    if idx in shortage_records:
                        temp_df.at[idx, 'Monthly_Balance'] = -abs(temp_df.at[idx, 'Monthly_Balance'])
            except:
                pass

    mrp_shortages = temp_df[temp_df['Monthly_Balance'] < 0].copy()
    mrp_shortages['Total_MRP_Shortage'] = mrp_shortages['Monthly_Balance'].abs()

    month_plan = active_plan_df[active_plan_df["YearMonth"].isin(target_yms_tuple)]
    plan_dict = month_plan.groupby("Assembly_PN")["Raw_Build_Qty"].sum().to_dict()

    for asm_wip, wip_qty in wip_cache.items():
        if wip_qty > 0 and asm_wip in plan_dict:
            sys_factor = ASSEMBLY_SYSTEM_FACTORS.get(asm_wip, 1)
            raw_wip_qty = wip_qty / sys_factor
            plan_dict[asm_wip] = max(0.0, plan_dict[asm_wip] - raw_wip_qty)

    breakdown_rows = []
    for idx, row in mrp_shortages.iterrows():
        pn = str(row[PN_COL]).strip()
        desc = str(row[DESC_COL])
        item_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in temp_df.columns else ""

        base_stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
        saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        stock = base_stock + saved_stock_add + sim_extra_stock_dict.get(pn, 0.0)

        total_mrp_shortage = row['Total_MRP_Shortage']
        _, _, item_status, current_sup, _, _, _ = get_inventory_record(pn, inv_cache)

        mouser_link = f"https://www.mouser.co.il/c/?q={pn}"
        digikey_link = f"https://www.digikey.com/en/products/result?keywords={pn}"
        findchips_link = f"https://www.findchips.com/search/{pn}"

        added_for_this_pn = False
        for asm in filtered_assembly_cols:
            qty_per_asm = pd.to_numeric(row[asm], errors='coerce') or 0
            if qty_per_asm > 0:
                added_for_this_pn = True
                asm_build_qty = plan_dict.get(asm, 0.0)
                required_demand = qty_per_asm * asm_build_qty
                asm_desc = assembly_mapping.get(asm, asm)

                breakdown_rows.append({
                    "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                    "Status": item_status, "Assembly": asm, "Assembly_Desc": asm_desc, "Qty_Per_Assembly": qty_per_asm,
                    "Assembly_Monthly_Build": asm_build_qty * ASSEMBLY_SYSTEM_FACTORS.get(asm, 1),
                    "Required_Demand": required_demand,
                    "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage,
                    "חיפוש במאוזר": mouser_link, "חיפוש בדיגיקי": digikey_link, "חיפוש ב-Findchips": findchips_link
                })

        if not added_for_this_pn:
            breakdown_rows.append({
                "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                "Status": item_status, "Assembly": "ללא שיוך", "Assembly_Desc": "ללא שיוך להרכבה", "Qty_Per_Assembly": 0,
                "Assembly_Monthly_Build": 0, "Required_Demand": 0, "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage,
                "חיפוש במאוזר": mouser_link, "חיפוש בדיגיקי": digikey_link, "חיפוש ב-Findchips": findchips_link
            })

    res_df = pd.DataFrame(breakdown_rows)
    return res_df

def calculate_mrp_breakdown(sim_extra_stock=None, target_yms=None, plan_df_override=None):
    if sim_extra_stock is None:
        sim_extra_stock = {}
    if target_yms is None:
        target_yms = selected_target_yms
    active_plan = plan_df_override if plan_df_override is not None else assembly_plan_df
    
    res = calculate_mrp_breakdown_cached(tuple(target_yms), tuple(sorted(sim_extra_stock.items())), active_plan)
    res_df = res.copy()

    if not res_df.empty:
        if selected_item_type != "הכל":
            res_df = res_df[res_df["Item_Type"] == selected_item_type]
        if selected_assembly != "הכל":
            res_df = res_df[res_df["Assembly"] == selected_assembly]
        if search_pn != "הכל":
            res_df = res_df[res_df["PN"] == search_pn]

    return res_df

breakdown_df = calculate_mrp_breakdown(target_yms=selected_target_yms)

# ==========================================================
# TABS DEFINITION (ALL 10 TABS FULLY PRESERVED)
# ==========================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    "📈 Executive Dashboard",
    "📊 תוכנית ייצור (Smart CTB)",
    "💡 סימולציית What-If",
    "📌 לוח סטטוסים (Kanban)",
    "🏭 ניהול WIP (בייצור)",
    "📅 עדכון מלאי וספקים",
    "📅 מעקב ETA ודחיות",
    "↩️ ניהול UNDO",
    "📦 ניהול מלאי מעודכן",
    "🎯 ניתוח רגישות ותוכנית"
])

with tab1:
    israel_time = datetime.utcnow() + timedelta(hours=3)
    current_time_str = israel_time.strftime("%d/%m/%Y | %H:%M:%S")
    st.markdown(f"""
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; opacity: 0.85; font-weight: 600;">
        <div>🎯 תמונת מצב ניהולית לטווח חודשים: {', '.join(selected_target_yms)}</div>
        <div>🕒 שעון ישראל (עדכני): {current_time_str}</div>
    </div>
    """, unsafe_allow_html=True)

    dash_df = breakdown_df.copy()
    if selected_assembly != "הכל":
        dash_df = dash_df[dash_df["Assembly"] == selected_assembly]

    wip_cache_dash = fetch_wip_records()
    inv_cache_dash = fetch_all_inventory_records()

    total_planned_qty = 0.0
    total_executable_qty = 0.0
    total_planned_assemblies_count = 0
    blocked_assemblies = len(dash_df['Assembly'].unique()) if not dash_df.empty else 0

    assemblies_to_evaluate = [a for a in valid_assemblies if selected_assembly == "הכל" or a == selected_assembly]

    for asm_col in assemblies_to_evaluate:
        for target_m in selected_target_yms:
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan_df["Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache_dash.get(asm_col, 0.0)

            if raw_build > 0 or current_wip_qty > 0:
                total_planned_assemblies_count += 1
                total_planned_qty += raw_build

                month_breakdown = calculate_mrp_breakdown(target_yms=[target_m])
                asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

                max_possible_build = raw_build
                if not asm_shortages.empty and raw_build > 0:
                    for _, s_row in asm_shortages.iterrows():
                        req_per = s_row["Qty_Per_Assembly"]
                        if req_per > 0:
                            comp_pn = str(s_row["PN"]).strip()
                            comp_match = df[df[PN_COL].astype(str).str.strip() == comp_pn]
                            if not comp_match.empty:
                                c_row = comp_match.iloc[0]
                                base_stk = pd.to_numeric(c_row.get(STOCK_COL, 0), errors='coerce') or 0
                                saved_stk, _, _, _, _, _, _ = get_inventory_record(comp_pn, inv_cache_dash)
                                total_comp_stock = base_stk + saved_stk
                                possible_from_this = total_comp_stock / req_per
                                max_possible_build = min(max_possible_build, possible_from_this)
                    gross_executable = max(0.0, min(raw_build, max_possible_build))
                else:
                    gross_executable = raw_build

                net_executable_qty = max(0.0, gross_executable - current_wip_qty)
                total_executable_qty += net_executable_qty

    readiness_pct = (total_executable_qty / total_planned_qty * 100) if total_planned_qty > 0 else 100
    unique_shortage_count = len(dash_df['PN'].unique()) if not dash_df.empty else 0
    active_wip_list = [(w, q) for w, q in wip_cache_dash.items() if q > 0]
    total_wip_active_count = len(active_wip_list)

    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    with col_k1:
        kpi_card("🟢 מוכנות ייצור משוקללת", f"{readiness_pct:.1f}%", f"{total_executable_qty:,.0f} / {total_planned_qty:,.0f} יחידות ניתן לייצור", "green")
    with col_k2:
        kpi_card("🔴 הרכבות חסומות", blocked_assemblies, "בטווח הנבחר", "red")
    with col_k3:
        kpi_card("🏭 פעילים ב-WIP", total_wip_active_count, "הודעות ייצור פעילות", "blue")
    with col_k4:
        kpi_card("📦 מק'טים בגירעון", unique_shortage_count, "פריטים ייחודיים", "orange")
    with col_k5:
        kpi_card("📊 גירעון מצטברת", f"{dash_df['Total_MRP_Shortage'].sum():,.0f}" if not dash_df.empty else "0", "יחידות", "blue")

    with st.expander("🔍 הצג פירוט כרטיסי הרכבות פעילים ב-WIP (לחץ לפתיחה)", expanded=False):
        if active_wip_list:
            wip_detail_rows = [{"קוד הרכבה": asm_pn, "תיאור הרכבה": df_desc.iloc[0, df.columns.get_loc(asm_pn)] if asm_pn in df.columns else "", "כמות ב-WIP": asm_qty, "רמה בעץ": assembly_levels.get(asm_pn, 0)} for asm_pn, asm_qty in active_wip_list]
            st.dataframe(pd.DataFrame(wip_detail_rows), use_container_width=True)
        else:
            st.info("אין כרגע הרכבות פעילות ב-WIP.")

    st.divider()

    if not dash_df.empty and len(dash_df) > 0:
        col_g0, col_g1, col_g2 = st.columns([1, 1.2, 1.2])

        with col_g0:
            st.markdown("##### 🎯 מד מוכנות ייצור")
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=readiness_pct,
                number={'suffix': "%", 'font': {'size': 34}},
                gauge={
                    'axis': {'range': [0, 100]},
                    'bar': {'color': PRIMARY},
                    'steps': [
                        {'range': [0, 50], 'color': '#3B1F1F'},
                        {'range': [50, 80], 'color': '#3B2F1F'},
                        {'range': [80, 100], 'color': '#1F3B2A'},
                    ],
                }
            ))
            fig_gauge.update_layout(template=PLOTLY_TEMPLATE, height=260, margin=dict(t=10, b=10, l=20, r=20))
            st.plotly_chart(fig_gauge, use_container_width=True)

        with col_g1:
            st.markdown("##### 🥧 התפלגות חוסרים לפי סוג פריט")
            fig_pie = px.pie(dash_df, names="Item_Type", values="Total_MRP_Shortage", hole=0.5, color_discrete_sequence=COLOR_SEQ)
            fig_pie.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_g2:
            st.markdown("##### 🏭 התפלגות חוסרים לפי ספק")
            fig_sup = px.pie(dash_df, names="Supplier", values="Total_MRP_Shortage", hole=0.5, color_discrete_sequence=COLOR_SEQ)
            fig_sup.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_sup, use_container_width=True)

        st.markdown('<div class="section-title">📋 טבלת פירוט ניהולית עם אפשרות ייצוא וקישורי חיפוש מלאי</div>', unsafe_allow_html=True)
        display_df = dash_df[[
            "PN", "Description", "Item_Type", "Supplier", "Status", "Assembly", "Assembly_Desc",
            "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage",
            "חיפוש במאוזר", "חיפוש בדיגיקי", "חיפוש ב-Findchips"
        ]].rename(columns={
            "PN": "מק'ט", "Description": "תיאור פריט", "Item_Type": "סוג פריט", "Supplier": "ספק",
            "Status": "סטטוס טיפול", "Assembly": "קוד הרכבה", "Assembly_Desc": "תיאור הרכבה",
            "Qty_Per_Assembly": "כמות נדרשת", "Assembly_Monthly_Build": "ת. ייצור",
            "Required_Demand": "ביקוש מדויק", "Stock": "מלאי", "Total_MRP_Shortage": "סך חוסר"
        })

        def _shortage_color(val, vmax):
            if vmax <= 0:
                return ""
            ratio = min(1.0, float(val) / vmax)
            return f"background-color: rgba(239,{int(180 - ratio * 140)},{int(120 - ratio * 100)},0.55); color: white;"

        sorted_display_df = display_df.sort_values(by="סך חוסר", ascending=False)
        max_shortage = sorted_display_df["סך חוסר"].max() if not sorted_display_df.empty else 0

        styled = sorted_display_df.style.map(lambda v: _shortage_color(v, max_shortage), subset=["סך חוסר"])
        st.dataframe(styled, column_config={
            "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
            "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
            "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
        }, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='Executive_Shortages')
        st.download_button(label="📥 הורד דו'ח מנהלים מלא ל-Excel", data=output.getvalue(), file_name=f"MRP_Executive_Report_{selected_ym}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.success("🎉 אין חוסרים ב-MRP עבור ההגדרות והסינונים שנבחרו!")

with tab2:
    st.markdown(f'<div class="section-title">📊 סימולציית Clear To Build (CTB) מטריציונית עם השוואת כמויות וגרף הרכבות מפורט</div>', unsafe_allow_html=True)
    inv_cache_ctb = fetch_all_inventory_records()
    wip_cache_ctb = fetch_wip_records()
    matrix_rows, chart_assembly_data = [], []
    assemblies_to_check = [asm for asm in valid_assemblies if selected_assembly == "הכל" or asm == selected_assembly]

    for asm_col in assemblies_to_check:
        asm_desc = df_desc.iloc[0, df.columns.get_loc(asm_col)] if asm_col in df.columns else ""
        row_data = {"קוד הרכבה": asm_col, "תיאור הרכבה": asm_desc, "רמה בעץ": assembly_levels.get(asm_col, 0)}
        has_any_build = False

        for target_m in selected_target_yms:
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan_df["Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache_ctb.get(asm_col, 0.0)

            if raw_build > 0 or current_wip_qty > 0:
                has_any_build = True

            month_breakdown = calculate_mrp_breakdown(target_yms=[target_m])
            asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

            max_possible_build = raw_build
            if not asm_shortages.empty and raw_build > 0:
                for _, s_row in asm_shortages.iterrows():
                    req_per = s_row["Qty_Per_Assembly"]
                    if req_per > 0:
                        comp_pn = str(s_row["PN"]).strip()
                        comp_match = df[df[PN_COL].astype(str).str.strip() == comp_pn]
                        if not comp_match.empty:
                            c_row = comp_match.iloc[0]
                            base_stk = pd.to_numeric(c_row.get(STOCK_COL, 0), errors='coerce') or 0
                            saved_stk, _, _, _, _, _, _ = get_inventory_record(comp_pn, inv_cache_ctb)
                            max_possible_build = min(max_possible_build, (base_stk + saved_stk) / req_per)
                gross_executable = max(0.0, min(raw_build, max_possible_build))
            else:
                gross_executable = raw_build

            net_executable_qty = max(0.0, gross_executable - current_wip_qty)
            row_data[f"תכנית ייצור ({target_m})"] = raw_build
            row_data[f"ניתן לייצור ({target_m})"] = net_executable_qty
            row_data[f"WIP ({target_m})"] = current_wip_qty

            if raw_build > 0 or current_wip_qty > 0:
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "תכנית ייצור", "כמות": raw_build})
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "ניתן לייצור בפועל", "כמות": net_executable_qty})
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "WIP", "כמות": current_wip_qty})

        for target_m in selected_target_yms:
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan_df["Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache_ctb.get(asm_col, 0.0)
            net_build = max(0.0, raw_build - current_wip_qty)

            month_breakdown = calculate_mrp_breakdown(target_yms=[target_m])
            asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

            missing_items_details = []
            for _, s_row in asm_shortages.iterrows():
                c_pn, c_desc, s_qty = str(s_row["PN"]).strip(), str(s_row["Description"]).strip(), s_row["Total_MRP_Shortage"]
                raw_eta = get_first_supply_eta(c_pn, inv_cache_ctb)
                missing_items_details.append((c_pn, c_desc, s_qty, raw_eta))

            if missing_items_details:
                formatted_missing = [f"{c_pn} ({c_desc[:10]}) - חסר: {m_qty:g} [ETA: {raw_eta}]" for c_pn, c_desc, m_qty, raw_eta in missing_items_details]
                row_data[f"סטטוס וחוסרים ({target_m})"] = "❌ חסר: " + " | ".join(formatted_missing)
            else:
                row_data[f"סטטוס וחוסרים ({target_m})"] = "✅ מוכן לייצור מלא" if net_build > 0 else "💤 ללא תוכנית ייצור"

        if has_any_build:
            matrix_rows.append(row_data)

    if matrix_rows:
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, height=420)
        if chart_assembly_data:
            fig_bar_asm = px.bar(pd.DataFrame(chart_assembly_data), x="הרכבה ותיאור", y="כמות", color="מדד", barmode="group", color_discrete_sequence=[PRIMARY, SUCCESS, ACCENT])
            fig_bar_asm.update_layout(template=PLOTLY_TEMPLATE, height=400, margin=dict(t=20, b=40, l=20, r=20), xaxis_tickangle=-25)
            st.plotly_chart(fig_bar_asm, use_container_width=True)

with tab3:
    st.markdown('<div class="section-title">💡 סימולציית What-If (מה יקרה אם...)</div>', unsafe_allow_html=True)
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        sim_pn = st.selectbox("בחר מק'ט לסימולציה", sorted(df[PN_COL].dropna().astype(str).unique()), key="sim_pn")
    with col_w2:
        sim_extra_stock = st.number_input("תוספת כמות מדומיינת למלאי", min_value=0.0, value=10.0, step=1.0)

    if st.button("🔮 הרץ סימולציית שחרור צוואר בקבוק"):
        sim_df = calculate_mrp_breakdown({sim_pn: sim_extra_stock}, target_yms=selected_target_yms)
        orig_blocked = set(breakdown_df['Assembly'].unique()) if not breakdown_df.empty else set()
        sim_blocked = set(sim_df['Assembly'].unique()) if not sim_df.empty else set()
        st.success(f"סימולציה הופעלה בהצלחה עבור מק'ט `{sim_pn}`.")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            kpi_card("🟢 הרכבות שהשתחררו", len(orig_blocked - sim_blocked), "", "green")
        with col_m2:
            kpi_card("🔴 עדיין חסום", len(sim_blocked), "", "red")
        with col_m3:
            before_after_delta = (breakdown_df['Total_MRP_Shortage'].sum() if not breakdown_df.empty else 0) - (sim_df['Total_MRP_Shortage'].sum() if not sim_df.empty else 0)
            kpi_card("📉 צמצום גירעון", f"{before_after_delta:,.0f}", "יחידות", "blue")

with tab4:
    st.markdown('<div class="section-title">📌 לוח מעקב סטטוסים (Kanban Pipeline)</div>', unsafe_allow_html=True)
    statuses = [("פתוח", "📝 פתוח לטיפול", "#3B1F1F", DANGER), ("הוזמן", "🛒 הוזמן / בטיפול רכש", "#3B2F1F", WARNING), ("בדרך", "🚚 בדרך לקו", "#1F2A3B", ACCENT), ("התקבל", "✅ התקבל / סגור", "#1F3B2A", SUCCESS)]
    dedup_all = breakdown_df.drop_duplicates(subset=["PN"]) if not breakdown_df.empty else pd.DataFrame()
    kcols = st.columns(4)
    for (status_key, title, bg, accent_color), kcol in zip(statuses, kcols):
        with kcol:
            items = dedup_all[dedup_all["Status"] == status_key] if not dedup_all.empty else pd.DataFrame()
            st.markdown(f'<div class="kanban-col-header" style="background:{bg}; color:{accent_color};">{title} ({len(items)})</div>', unsafe_allow_html=True)
            for _, r in items.head(6).iterrows():
                st.markdown(f'<div class="kanban-card" style="border-color:{accent_color};"><b>{r["PN"]}</b><br><span style="opacity:0.75;">{str(r["Description"])[:24]}</span></div>', unsafe_allow_html=True)

with tab5:
    st.markdown(f'<div class="section-title">🏭 ניהול WIP חכם (כולל סגירת מחזור ייצור ואימות היררכיה)</div>', unsafe_allow_html=True)
    wip_current = fetch_wip_records()
    if wip_current:
        with st.form("close_wip_form"):
            wip_to_close = st.selectbox("בחר הרכבה שסיימה ייצור לחודש זה", list(wip_current.keys()), format_func=lambda x: f"{x} (כמות ב-WIP: {wip_current[x]})")
            is_finished = st.checkbox("האם ההרכבה הסתיימה לחלוטין והושלמה בהצלחה?")
            if st.form_submit_button("סגור WIP והוסף למלאי הזמין"):
                if is_finished:
                    closing_qty = wip_current[wip_to_close]
                    curr_stk, curr_eta, curr_stat, curr_sup, curr_comm, _, _ = get_inventory_record(wip_to_close)
                    save_inventory_record(wip_to_close, curr_stk + closing_qty, curr_eta, "התקבל", curr_sup, f"{curr_comm} | הושלם מייצור WIP בסך {closing_qty}", "WIP Close", webhook_url)
                    delete_wip_record(wip_to_close)
                    st.rerun()

    with st.form("wip_form"):
        wip_asm_choice = st.selectbox("בחר הרכבה חדשה לצירוף ל-WIP", filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x))
        wip_qty_input = st.number_input("כמות יחידות הרכבה להוספה לייצור (WIP)", min_value=0.0, value=1.0, step=1.0)
        if st.form_submit_button("בדיקת זמינות היררכית מלאה ושמור WIP"):
            save_wip_record(wip_asm_choice, wip_qty_input)
            st.success("ההרכבה נוספה בהצלחה ל-WIP!")
            st.rerun()

with tab6:
    st.markdown('<div class="section-title">📅 עדכון מלאי, סטטוס ודחיית ספקים (ETA)</div>', unsafe_allow_html=True)
    selected_pn = search_pn if search_pn != "הכל" else st.selectbox("בחר מק'ט מכלל הפריטים לעדכון", sorted(df[PN_COL].dropna().astype(str).unique()), key="update_pn_select")
    if selected_pn != "הכל":
        saved_stock, saved_eta, saved_status, saved_supplier, saved_comment, saved_by, _ = get_inventory_record(selected_pn)
        with st.form("inventory_form"):
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                added_stock_input = st.number_input("תוספת למלאי זמין", min_value=0.0, value=float(saved_stock), step=1.0)
            with col_f2:
                try: parsed_eta = pd.to_datetime(saved_eta).date() if saved_eta else date.today()
                except: parsed_eta = date.today()
                eta_date = st.date_input("תאריך הגעה מעודכן (ETA)", value=parsed_eta)
            with col_f3:
                status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
                status = st.selectbox("סטטוס טיפול", status_options, index=status_options.index(saved_status) if saved_status in status_options else 0)
            supplier = st.selectbox("ספק", supplier_options, index=supplier_options.index(saved_supplier) if saved_supplier in supplier_options else 0)
            comment = st.text_area("הערות", value=saved_comment)
            if st.form_submit_button("שמור עדכון קבוע בענן"):
                save_inventory_record(selected_pn, added_stock_input, str(eta_date), status, supplier, comment, "User", webhook_url)
                st.success("העדכון נשמר!")
                st.rerun()

with tab7:
    st.markdown('<div class="section-title">📅 מעקב ETA, דחיות, כמויות וקישורים למפיצים</div>', unsafe_allow_html=True)
    inv_cache_all = fetch_all_inventory_records()
    eta_table_rows = []

    for _, row in df.iterrows():
        p_num = str(row[PN_COL]).strip()
        if not p_num or p_num == 'nan':
            continue
        p_desc = str(row[DESC_COL])
        p_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in df.columns else ""

        orig_eta = get_base_mrp_eta(p_num)
        orig_qty = get_base_mrp_qty(p_num)
        
        saved_rec = inv_cache_all.get(p_num, {})
        current_eta_raw = saved_rec.get("eta", "")
        current_added_stock = saved_rec.get("added_stock", 0.0)
        curr_eta_fmt = pd.to_datetime(current_eta_raw).strftime("%Y-%m") if current_eta_raw else orig_eta

        eta_table_rows.append({
            "מק'ט": p_num,
            "תיאור פריט": p_desc,
            "סוג פריט": p_type,
            "ETA מקורי (MRP)": orig_eta,
            "כמות מקורית": orig_qty,
            "ETA מעודכן": curr_eta_fmt,
            "כמות מעודכנת": current_added_stock,
            "ספק": saved_rec.get("supplier", "אופק"),
            "חיפוש במאוזר": f"https://www.mouser.co.il/c/?q={p_num}",
            "חיפוש בדיגיקי": f"https://www.digikey.com/en/products/result?keywords={p_num}",
            "חיפוש ב-Findchips": f"https://www.findchips.com/search/{p_num}"
        })

    eta_df = pd.DataFrame(eta_table_rows)
    if not eta_df.empty:
        st.dataframe(eta_df, column_config={
            "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
            "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
            "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
        }, use_container_width=True, height=450)

with tab8:
    st.markdown('<div class="section-title">↩️ חזרה לאחור וניהול היסטוריה (UNDO)</div>', unsafe_allow_html=True)
    try: updated_items = supabase.table("mrp_inventory_updates").select("*").order("updated_at", desc=True).execute().data or []
    except: updated_items = []
    for item in updated_items:
        col_u1, col_u2, col_u3 = st.columns([3, 4, 1])
        with col_u1: st.markdown(f"**מק'ט:** `{item.get('pn')}`")
        with col_u2: st.text(f"תוספת: {item.get('added_stock')} | ETA: {item.get('eta')}")
        with col_u3:
            if st.button("🔄 UNDO", key=f"undo_{item.get('pn')}"):
                delete_inventory_record(item.get('pn'))
                st.rerun()

with tab9:
    st.markdown('<div class="section-title">📦 ניהול מלאי מעודכן (עריכה וגריעת כמויות)</div>', unsafe_allow_html=True)
    active_stock_items = {k: v for k, v in fetch_all_inventory_records().items() if float(v.get("added_stock", 0.0)) > 0}
    if active_stock_items:
        st.dataframe(pd.DataFrame([{"מק'ט": k, "כמות": v["added_stock"], "ETA": v["eta"]} for k, v in active_stock_items.items()]), use_container_width=True)
        selected_mgmt_pn = st.selectbox("בחר מק'ט לעריכה או גריעה", list(active_stock_items.keys()), key="mgmt_pn_select")
        if selected_mgmt_pn:
            with st.form("edit_mgmt_form"):
                new_qty = st.number_input("עדכן כמות", min_value=0.0, value=float(active_stock_items[selected_mgmt_pn]["added_stock"]), step=1.0)
                if st.form_submit_button("🗑️ אפס או עדכן"):
                    delete_inventory_record(selected_mgmt_pn)
                    if new_qty > 0:
                        save_inventory_record(selected_mgmt_pn, new_qty, "", "פתוח", "אופק", "", "Tab 9", webhook_url)
                    st.rerun()

with tab10:
    st.markdown('<div class="section-title">🎯 ניתוח רגישות וניהול תוכנית הייצור (עריכה פרטנית לפי הרכבה וחודש)</div>', unsafe_allow_html=True)
    if not assembly_plan_df.empty:
        orig_pivot_plan = assembly_plan_df.pivot_table(index=["Assembly_PN"], columns="YearMonth", values="Build_Qty", fill_value=0.0).reset_index()
        orig_pivot_plan.insert(1, "רמה", orig_pivot_plan["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        orig_pivot_plan.insert(2, "תיאור הרכבה", orig_pivot_plan["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        orig_pivot_plan = orig_pivot_plan.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(orig_pivot_plan, use_container_width=True, height=280)

    st.divider()
    st.markdown("##### ⚙️ הגדרת שינוי רגישות: גורף או חודש ספציפי")
    
    col_mode_choice = st.columns(2)
    with col_mode_choice[0]:
        sens_scope = st.radio("היקף השינוי", ["שינוי גורף לכל החודשים", "שינוי לחודש ספציפי בלבד"], horizontal=True, key="sens_scope")
    with col_mode_choice[1]:
        if sens_scope == "שינוי לחודש ספציפי בלבד":
            available_yms = sorted(assembly_plan_df["YearMonth"].unique())
            target_sens_month = st.selectbox("בחר חודש ספציפי לעדכון", available_yms, key="target_sens_month")

    col_sens1, col_sens2, col_sens3 = st.columns([1.2, 1, 1])
    with col_sens1:
        sens_assembly_target = st.selectbox("בחר הרכבה לניתוח רגישות", ["הכל (כלל ההרכבות)"] + filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x), key="sens_assembly_target")
    with col_sens2:
        sens_mode = st.radio("סוג שינוי", ["אחוזים (%)", "מספרי (יחידות)"], horizontal=True, key="sens_mode")
    with col_sens3:
        if sens_mode == "אחוזים (%)":
            sensitivity_val = st.slider("שינוי אחוז תוכנית הייצור (%)", -50, 100, 0, 5, key="sens_slider")
        else:
            sensitivity_val = st.number_input("תוספת/הפחתה מספרית (יחידות)", -500, 500, 0, 1, key="sens_num")

    if st.button("🚀 הרץ ניתוח רגישות לתוכנית", key="run_sensitivity"):
        simulated_plan_df = assembly_plan_df.copy()
        
        if sens_scope == "שינוי גורף לכל החודשים":
            if sens_mode == "אחוזים (%)" and sensitivity_val != 0:
                multiplier = 1.0 + (sensitivity_val / 100.0)
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    simulated_plan_df["Raw_Build_Qty"] *= multiplier
                    simulated_plan_df["Build_Qty"] *= multiplier
                else:
                    mask = simulated_plan_df["Assembly_PN"] == sens_assembly_target
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
            elif sens_mode == "מספרי (יחידות)" and sensitivity_val != 0:
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    simulated_plan_df["Raw_Build_Qty"] = (simulated_plan_df["Raw_Build_Qty"] + sensitivity_val).clip(lower=0)
                    simulated_plan_df["Build_Qty"] = (simulated_plan_df["Build_Qty"] + sensitivity_val).clip(lower=0)
                else:
                    mask = simulated_plan_df["Assembly_PN"] == sens_assembly_target
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] = (simulated_plan_df.loc[mask, "Raw_Build_Qty"] + sensitivity_val).clip(lower=0)
                    simulated_plan_df.loc[mask, "Build_Qty"] = (simulated_plan_df.loc[mask, "Build_Qty"] + sensitivity_val).clip(lower=0)
        else:
            # שינוי לחודש ספציפי בלבד
            sys_factor_map = ASSEMBLY_SYSTEM_FACTORS
            if sens_mode == "אחוזים (%)" and sensitivity_val != 0:
                multiplier = 1.0 + (sensitivity_val / 100.0)
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    mask = simulated_plan_df["YearMonth"] == target_sens_month
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
                else:
                    mask = (simulated_plan_df["Assembly_PN"] == sens_assembly_target) & (simulated_plan_df["YearMonth"] == target_sens_month)
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
            elif sens_mode == "מספרי (יחידות)" and sensitivity_val != 0:
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    mask = simulated_plan_df["YearMonth"] == target_sens_month
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] = (simulated_plan_df.loc[mask, "Raw_Build_Qty"] + sensitivity_val).clip(lower=0)
                    simulated_plan_df.loc[mask, "Build_Qty"] = (simulated_plan_df.loc[mask, "Build_Qty"] + sensitivity_val).clip(lower=0)
                else:
                    mask = (simulated_plan_df["Assembly_PN"] == sens_assembly_target) & (simulated_plan_df["YearMonth"] == target_sens_month)
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] = (simulated_plan_df.loc[mask, "Raw_Build_Qty"] + sensitivity_val).clip(lower=0)
                    simulated_plan_df.loc[mask, "Build_Qty"] = (simulated_plan_df.loc[mask, "Build_Qty"] + sensitivity_val).clip(lower=0)

        st.session_state["temp_simulated_plan"] = simulated_plan_df
        st.success("ניתוח הרגישות בוצע בהצלחה! צפה בתוצאות למטה ומאשר לשמור במידת הצורך.")

    if "temp_simulated_plan" in st.session_state:
        st.divider()
        st.markdown("##### 📋 תצוגה מקדימה של התוכנית הסימולטיבית (לאחר ניתוח רגישות):")
        preview_pivot = st.session_state["temp_simulated_plan"].pivot_table(index=["Assembly_PN"], columns="YearMonth", values="Build_Qty", fill_value=0.0).reset_index()
        preview_pivot.insert(1, "רמה", preview_pivot["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        preview_pivot.insert(2, "תיאור הרכבה", preview_pivot["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        preview_pivot = preview_pivot.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(preview_pivot, use_container_width=True, height=240)

        with st.form("update_plan_form"):
            update_confirmation = st.checkbox("❓ האם אתה מאשר לשמור את השינויים ולהחיל את תוכנית הייצור החדשה על כלל המערכת?")
            if st.form_submit_button("💾 שמור שינויים ועדכן את תוכנית העבודה"):
                if update_confirmation:
                    st.session_state["previous_approved_plan"] = assembly_plan_df.copy()
                    st.session_state["custom_assembly_plan_df"] = st.session_state["temp_simulated_plan"]
                    del st.session_state["temp_simulated_plan"]
                    st.success("תוכנית הייצור עודכנה ונשמרה בהצלחה במערכת!")
                    st.rerun()
                else:
                    st.warning("יש לסמן את תיבת האישור כדי לשמור את השינויים.")

    if "previous_approved_plan" in st.session_state:
        st.divider()
        st.markdown("##### 📊 השוואה בין תוכנית הייצור הקודמת לחדשה")
        orig_plan_pivot = st.session_state["previous_approved_plan"].pivot_table(index="Assembly_PN", columns="YearMonth", values="Build_Qty", fill_value=0.0)
        new_plan_pivot = assembly_plan_df.pivot_table(index="Assembly_PN", columns="YearMonth", values="Build_Qty", fill_value=0.0)
        comparison_diff = new_plan_pivot.sub(orig_plan_pivot, fill_value=0.0).reset_index()
        comparison_diff.insert(1, "רמה", comparison_diff["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        comparison_diff.insert(2, "תיאור הרכבה", comparison_diff["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        comparison_diff = comparison_diff.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(comparison_diff, use_container_width=True)
