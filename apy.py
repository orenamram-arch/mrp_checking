"""
MRP Control Tower — מגדל בקרת חוסרים
גרסה מלאה ומושלמת הכוללת:
1. טבלת CTB מטריציונית עם הפרדת עמודות לתוכנית ייצור חודשית, WIP חודשי וחוסרים.
2. קישורים ישירים לחיפוש מלאי (Mouser, DigiKey, Findchips) בכל הטבלאות.
3. איפוס מסננים ויזואלי מלא בסיידבר (Clear All).
4. ניהול סגירת WIP והמרתו למלאי, סימולציות What-If ותוכנית מרובת חודשים.

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
# CONFIGURATION
# ==========================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"

st.set_page_config(
    page_title="MRP Executive Control Tower",
    page_icon="🚀",
    layout="wide"
)

# ==========================================================
# GLOBAL THEME / CSS
# ==========================================================
PRIMARY = "#4F46E5"      # indigo
PRIMARY_DARK = "#3730A3"
ACCENT = "#06B6D4"       # cyan
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

@media (max-width: 640px) {{
    .hero-banner {{ padding: 18px 16px; border-radius: 14px; }}
    .hero-banner h1 {{ font-size: 21px; }}
    .hero-banner p {{ font-size: 13px; }}
    .kpi-value {{ font-size: 22px; }}
    .kpi-label {{ font-size: 12px; }}
    .kpi-card {{ padding: 14px 10px; }}
    .section-title {{ font-size: 16px; }}
    [data-testid="stHorizontalBlock"] {{ flex-wrap: wrap !important; }}
    [data-testid="column"] {{ min-width: 100% !important; flex: 1 1 100% !important; }}
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

@st.cache_data(ttl=5)
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
    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {
        "assembly_pn": str(assembly_pn),
        "wip_qty": float(wip_qty),
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
# DATA LOADING FROM GITHUB
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

# ==========================================================
# EXTRACT ASSEMBLY MONTHLY BUILD PLAN & BOM LEVELS
# ==========================================================
header_dates = df_raw.iloc[2, 108:132].values if df_raw.shape[1] > 132 else []
plan_rows = []

for r in range(3, min(24, df_raw.shape[0])):
    asm_pn = df_raw.iloc[r, 106] if df_raw.shape[1] > 106 else None
    if pd.notnull(asm_pn):
        for c_idx, date_val in enumerate(header_dates):
            if pd.notnull(date_val):
                qty = df_raw.iloc[r, 108 + c_idx]
                if pd.notnull(qty) and qty != '' and qty != 'NaN':
                    try:
                        q_val = float(qty)
                        if q_val > 0:
                            dt = pd.to_datetime(date_val)
                            plan_rows.append({
                                "Assembly_PN": str(asm_pn).strip(),
                                "YearMonth": dt.strftime("%Y-%m"),
                                "Build_Qty": q_val
                            })
                    except:
                        pass

assembly_plan_df = pd.DataFrame(plan_rows)

PN_COL = df.columns[1]
DESC_COL = df.columns[4]
ITEM_TYPE_COL = df.columns[44] if len(df.columns) > 44 else df.columns[-1]
STOCK_COL = df.columns[79] if len(df.columns) > 79 else df.columns[-1]
ASSEMBLY_COLS = df.columns[10:36].tolist()
MONTH_COLS = df.columns[108:132].tolist() if len(df.columns) > 132 else []

valid_assemblies = []
for col in ASSEMBLY_COLS:
    try:
        col_type = df.loc[df[PN_COL] == col, ITEM_TYPE_COL].values
        if len(col_type) > 0 and str(col_type[0]) != 'nan':
            valid_assemblies.append(col)
        else:
            valid_assemblies.append(col)
    except:
        pass

assembly_levels = {}
for col in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(col)
        lvl_val = int(df_levels.iloc[0, col_idx])
        assembly_levels[col] = lvl_val
    except:
        assembly_levels[col] = 0

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
# SIDEBAR FILTERS & WHAT-IF CONTROLS (WITH CLEAR ALL RESET)
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

current_ym_str = datetime.now().strftime("%Y-%m")
month_options = {}
for m in MONTH_COLS:
    if pd.notnull(m):
        try:
            dt = pd.to_datetime(m)
            m_ym = dt.strftime("%Y-%m")
            if m_ym >= current_ym_str:
                month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
        except:
            pass

if not month_options:
    for m in MONTH_COLS:
        if pd.notnull(m):
            try:
                dt = pd.to_datetime(m)
                month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
            except:
                month_options[str(m)] = m

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()), key="selected_month_label")
selected_month_col = month_options[selected_month_label]

try:
    selected_ym = pd.to_datetime(selected_month_col).strftime("%Y-%m")
except:
    selected_ym = str(selected_month_col)[:7]

num_months_ahead = st.sidebar.slider("📅 טווח מבט קדימה במספר חודשים", min_value=1, max_value=6, value=1, key="num_months_ahead")

level_options = ["הכל"] + sorted([str(df_levels.iloc[0, df.columns.get_loc(c)]) for c in valid_assemblies if pd.notnull(df_levels.iloc[0, df.columns.get_loc(c)])])
selected_level = st.sidebar.selectbox("סינון לפי רמת עץ (BOM Level)", level_options, key="selected_level")

assembly_mapping = {"הכל": "הכל"}
filtered_assembly_cols = []
for col in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(col)
        lvl = str(df_levels.iloc[0, col_idx])
        desc = df_desc.iloc[0, col_idx]
        if selected_level == "הכל" or lvl == selected_level:
            filtered_assembly_cols.append(col)
            assembly_mapping[col] = str(col) + " - " + str(desc) + " (רמה " + str(lvl) + ")"
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
# EXTERNAL ETA FILE UPLOAD (SUPPLIER FILE)
# ==========================================================
st.sidebar.divider()
st.sidebar.markdown("##### 📥 ייבוא קובץ ETA מעודכן מהספק")
uploaded_eta_file = st.sidebar.file_uploader("העלה קובץ Excel (PN, ETA, Qty)", type=["xlsx", "xls"])
if uploaded_eta_file is not None:
    try:
        sup_df = pd.read_excel(uploaded_eta_file)
        if st.sidebar.button("⚡ עדכן אוטומטית לפי קובץ ספק"):
            updated_count = 0
            for _, s_row in sup_df.iterrows():
                p_code = str(s_row.iloc[0]).strip()
                new_eta = str(s_row.iloc[1]).strip()
                new_qty = float(s_row.iloc[2]) if len(s_row) > 2 and pd.notnull(s_row.iloc[2]) else 0.0

                if p_code and p_code != 'nan':
                    curr_stock, _, curr_status, curr_sup, curr_comm, _, _ = get_inventory_record(p_code)
                    save_inventory_record(
                        pn=p_code,
                        added_stock=curr_stock + new_qty if new_qty > 0 else curr_stock,
                        eta=new_eta,
                        status=curr_status if curr_status != "פתוח" else "הוזמן",
                        supplier=curr_sup,
                        comment=f"{curr_comm} | עודכן מקובץ ספק חיצוני",
                        updated_by="Supplier File",
                        webhook_url=webhook_url
                    )
                    updated_count += 1
            st.sidebar.success(f"עודכנו בהצלחה {updated_count} פריטים מקובץ הספק!")
    except Exception as e:
        st.sidebar.error(f"שגיאה בקריאת קובץ הספק: {e}")

# ==========================================================
# CORE LOGIC FOR SHORTAGES (Multi-Month support)
# ==========================================================
def calculate_mrp_breakdown(sim_extra_stock=None, target_yms=None):
    if sim_extra_stock is None:
        sim_extra_stock = {}
    if target_yms is None:
        target_yms = [selected_ym]

    inv_cache = fetch_all_inventory_records()
    wip_cache = fetch_wip_records()

    active_month_cols = []
    for m_c in MONTH_COLS:
        if pd.notnull(m_c):
            try:
                m_dt_ym = pd.to_datetime(m_c).strftime("%Y-%m")
                if m_dt_ym in target_yms:
                    active_month_cols.append(m_c)
            except:
                pass
    if not active_month_cols:
        active_month_cols = [selected_month_col]

    temp_df = df.copy()
    temp_df['Monthly_Balance'] = temp_df[active_month_cols].sum(axis=1, numeric_only=True)

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        sim_val = sim_extra_stock.get(pn, 0.0)

        total_added_stock = saved_stock_add + sim_val

        if total_added_stock > 0:
            current_bal = temp_df.at[idx, 'Monthly_Balance']
            if current_bal < 0:
                temp_df.at[idx, 'Monthly_Balance'] = current_bal + total_added_stock

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
            try:
                eta_ym = pd.to_datetime(manual_eta).strftime("%Y-%m")
                if eta_ym > selected_ym:
                    temp_df.at[idx, 'Monthly_Balance'] = -abs(temp_df.at[idx, 'Monthly_Balance']) if temp_df.at[idx, 'Monthly_Balance'] < 0 else -1.0
            except:
                pass

    mrp_shortages = temp_df[temp_df['Monthly_Balance'] < 0].copy()
    mrp_shortages['Total_MRP_Shortage'] = mrp_shortages['Monthly_Balance'].abs()

    month_plan = assembly_plan_df[assembly_plan_df["YearMonth"].isin(target_yms)]
    plan_dict = month_plan.groupby("Assembly_PN")["Build_Qty"].sum().to_dict()

    for asm_wip, wip_qty in wip_cache.items():
        if wip_qty > 0 and asm_wip in plan_dict:
            plan_dict[asm_wip] = max(0.0, plan_dict[asm_wip] - wip_qty)

    breakdown_rows = []
    for idx, row in mrp_shortages.iterrows():
        pn = str(row[PN_COL]).strip()
        desc = str(row[DESC_COL])
        item_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in temp_df.columns else ""

        base_stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
        saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        stock = base_stock + saved_stock_add + sim_extra_stock.get(pn, 0.0)

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
                    "Assembly_Monthly_Build": asm_build_qty, "Required_Demand": required_demand,
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

    if not res_df.empty:
        if selected_item_type != "הכל":
            res_df = res_df[res_df["Item_Type"] == selected_item_type]
        if selected_assembly != "הכל":
            res_df = res_df[res_df["Assembly"] == selected_assembly]
        if search_pn != "הכל":
            res_df = res_df[res_df["PN"] == search_pn]

    return res_df

all_ym_list = sorted(list(set(assembly_plan_df["YearMonth"].unique())))
start_idx = 0
for idx, ym in enumerate(all_ym_list):
    if ym >= selected_ym:
        start_idx = idx
        break
selected_target_yms = all_ym_list[start_idx:start_idx + num_months_ahead]
if not selected_target_yms:
    selected_target_yms = [selected_ym]

breakdown_df = calculate_mrp_breakdown(target_yms=selected_target_yms)

# ==========================================================
# TABS DEFINITION
# ==========================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📈 Executive Dashboard",
    "📊 תוכנית ייצור (Smart CTB)",
    "💡 סימולציית What-If",
    "📌 לוח סטטוסים (Kanban)",
    "🏭 ניהול WIP (בייצור)",
    "📅 עדכון מלאי וספקים",
    "📅 מעקב ETA ודחיות",
    "↩️ ניהול UNDO"
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
    total_planned_assemblies = len([a for a in valid_assemblies if (assembly_plan_df[(assembly_plan_df["YearMonth"].isin(selected_target_yms)) & (assembly_plan_df["Assembly_PN"] == a)]["Build_Qty"].sum() - wip_cache_dash.get(a, 0)) > 0])
    blocked_assemblies = len(dash_df['Assembly'].unique()) if not dash_df.empty else 0
    ready_assemblies = max(0, total_planned_assemblies - blocked_assemblies)
    readiness_pct = (ready_assemblies / total_planned_assemblies * 100) if total_planned_assemblies > 0 else 100

    unique_shortage_count = len(dash_df['PN'].unique()) if not dash_df.empty else 0
    total_wip_active_count = len([w for w, q in wip_cache_dash.items() if q > 0])

    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    with col_k1:
        kpi_card("🟢 מוכנות קווי ייצור", f"{readiness_pct:.1f}%", f"{ready_assemblies}/{total_planned_assemblies} הרכבות מוכנות", "green")
    with col_k2:
        kpi_card("🔴 הרכבות חסומות", blocked_assemblies, "בטווח הנבחר", "red")
    with col_k3:
        kpi_card("🏭 פעילים ב-WIP", total_wip_active_count, "הודעות ייצור פעילות", "blue")
    with col_k4:
        kpi_card("📦 מק'טים בגירעון", unique_shortage_count, "פריטים ייחודיים", "orange")
    with col_k5:
        kpi_card("📊 גירעון מצטברת", f"{dash_df['Total_MRP_Shortage'].sum():,.0f}" if not dash_df.empty else "0", "יחידות", "blue")

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
            fig_pie = px.pie(dash_df, names="Item_Type", values="Total_MRP_Shortage", hole=0.5,
                             color_discrete_sequence=COLOR_SEQ)
            fig_pie.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_g2:
            st.markdown("##### 🏭 התפלגות חוסרים לפי ספק")
            fig_sup = px.pie(dash_df, names="Supplier", values="Total_MRP_Shortage", hole=0.5,
                             color_discrete_sequence=COLOR_SEQ)
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
            r = 239
            g = int(180 - ratio * 140)
            b = int(120 - ratio * 100)
            return f"background-color: rgba({r},{max(g,0)},{max(b,0)},0.55); color: white;"

        sorted_display_df = display_df.sort_values(by="סך חוסר", ascending=False)
        max_shortage = sorted_display_df["סך חוסר"].max() if not sorted_display_df.empty else 0

        styled = sorted_display_df.style.map(
            lambda v: _shortage_color(v, max_shortage), subset=["סך חוסר"]
        )
        st.dataframe(
            styled,
            column_config={
                "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
                "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
                "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
            },
            use_container_width=True
        )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='Executive_Shortages')
        processed_data = output.getvalue()

        st.download_button(
            label="📥 הורד דו'ח מנהלים מלא ל-Excel",
            data=processed_data,
            file_name=f"MRP_Executive_Report_{selected_ym}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("🎉 אין חוסרים ב-MRP עבור ההגדרות והסינונים שנבחרו!")

with tab2:
    st.markdown(f'<div class="section-title">📊 סימולציית Clear To Build (CTB) מטריציונית עם הפרדת עמודות תוכנית, WIP וחוסרים חודשיים</div>', unsafe_allow_html=True)
    st.markdown("טבלה זו מציגה לכל הרכבה שורה אחת אחידה, כאשר עבור כל חודש בטווח הנבחר מוצגות עמודות ייעודיות נפרדות: **כמות התוכנית**, **כמות ה-WIP** ו**סטטוס החוסרים והרכיבים הקריטיים**.")

    inv_cache_ctb = fetch_all_inventory_records()
    wip_cache_ctb = fetch_wip_records()

    matrix_rows = []
    assemblies_to_check = [asm for asm in valid_assemblies if selected_assembly == "הכל" or asm == selected_assembly]

    for asm_col in assemblies_to_check:
        try:
            asm_desc = df_desc.iloc[0, df.columns.get_loc(asm_col)]
        except:
            asm_desc = ""

        row_data = {
            "קוד הרכבה": asm_col,
            "תיאור הרכבה": asm_desc,
            "רמה בעץ": assembly_levels.get(asm_col, 0)
        }

        has_any_build = False

        for target_m in selected_target_yms:
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan_df["Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache_ctb.get(asm_col, 0.0)
            net_build = max(0.0, raw_build - current_wip_qty)

            if raw_build > 0 or current_wip_qty > 0:
                has_any_build = True

            # הכנסת נתונים לעמודות הנפרדות לחודש נתון
            row_data[f"תכנית ייצור ({target_m})"] = raw_build
            row_data[f"WIP ({target_m})"] = current_wip_qty

            # בדיקת חוסרים להרכבה זו בחודש הספציפי
            month_breakdown = calculate_mrp_breakdown(target_yms=[target_m])
            asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

            missing_items_details = []
            for _, s_row in asm_shortages.iterrows():
                c_pn = str(s_row["PN"]).strip()
                c_desc = str(s_row["Description"]).strip()
                s_qty = s_row["Total_MRP_Shortage"]
                eta_display_str = get_first_supply_eta(c_pn, inv_cache_ctb)
                try:
                    eta_dt = pd.to_datetime(eta_display_str).date() if eta_display_str != "בדיקה נדרשת" else date(2099, 12, 31)
                except:
                    eta_dt = date(2099, 12, 31)
                missing_items_details.append((c_pn, c_desc, s_qty, eta_dt, eta_display_str))

            if missing_items_details:
                missing_items_details.sort(key=lambda x: x[3], reverse=True)
                most_critical_pn = missing_items_details[0][0]

                formatted_missing = []
                for c_pn, c_desc, m_qty, _, raw_eta in missing_items_details:
                    eta_str = f" [ETA: {raw_eta}]" if raw_eta != "בדיקה נדרשת" else ""
                    item_text = f"{c_pn} ({c_desc[:10]}) - חסר: {m_qty:g}{eta_str}"
                    if c_pn == most_critical_pn:
                        formatted_missing.append(f"**{item_text}**")
                    else:
                        formatted_missing.append(item_text)
                cell_text = f"❌ חסר: " + " | ".join(formatted_missing)
            else:
                if net_build > 0:
                    cell_text = "✅ מוכן לייצור מלא"
                else:
                    cell_text = "💤 ללא תוכנית ייצור"

            row_data[f"סטטוס וחוסרים ({target_m})"] = cell_text

        if has_any_build:
            matrix_rows.append(row_data)

    if matrix_rows:
        matrix_df = pd.DataFrame(matrix_rows)
        st.dataframe(matrix_df, use_container_width=True, height=450)
    else:
        st.info("לא נמצאו הרכבות מתוכננות לייצור בטווח החודשים שנבחר.")

with tab3:
    st.markdown('<div class="section-title">💡 סימולציית What-If (מה יקרה אם...)</div>', unsafe_allow_html=True)
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        sim_pn = st.selectbox("בחר מק'ט לסימולציה", sorted(df[PN_COL].dropna().astype(str).unique()), key="sim_pn")
    with col_w2:
        sim_extra_stock = st.number_input("תוספת כמות מדומיינת למלאי לצורך סימולציה", min_value=0.0, value=10.0, step=1.0)

    if st.button("🔮 הרץ סימולציית שחרור צוואר בקבוק"):
        sim_df = calculate_mrp_breakdown({sim_pn: sim_extra_stock}, target_yms=selected_target_yms)
        orig_blocked = set(breakdown_df['Assembly'].unique()) if not breakdown_df.empty else set()
        sim_blocked = set(sim_df['Assembly'].unique()) if not sim_df.empty else set()
        freed_assemblies = orig_blocked - sim_blocked

        st.success(f"סימולציה הופעלה בהצלחה עבור מק'ט `{sim_pn}` עם תוספת מדומיינת של {sim_extra_stock} יחידות.")

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            kpi_card("🟢 הרכבות שהשתחררו", len(freed_assemblies), "מוכנות לייצור מלא", "green")
        with col_m2:
            kpi_card("🔴 עדיין חסום", len(sim_blocked), "הרכבות שנותרו חסומות", "red")
        with col_m3:
            before_after_delta = (breakdown_df['Total_MRP_Shortage'].sum() if not breakdown_df.empty else 0) - (sim_df['Total_MRP_Shortage'].sum() if not sim_df.empty else 0)
            kpi_card("📉 צמצום גירעון כולל", f"{before_after_delta:,.0f}", "יחידות", "blue")

with tab4:
    st.markdown('<div class="section-title">📌 לוח מעקב סטטוסים (Kanban Pipeline)</div>', unsafe_allow_html=True)
    statuses = [
        ("פתוח", "📝 פתוח לטיפול", "#3B1F1F", DANGER),
        ("הוזמן", "🛒 הוזמן / בטיפול רכש", "#3B2F1F", WARNING),
        ("בדרך", "🚚 בדרך לקו", "#1F2A3B", ACCENT),
        ("התקבל", "✅ התקבל / סגור", "#1F3B2A", SUCCESS),
    ]

    dedup_all = breakdown_df.drop_duplicates(subset=["PN"]) if not breakdown_df.empty else pd.DataFrame()
    total_items = len(dedup_all) if not dedup_all.empty else 0

    kcols = st.columns(4)
    for (status_key, title, bg, accent_color), kcol in zip(statuses, kcols):
        with kcol:
            items = dedup_all[dedup_all["Status"] == status_key] if not dedup_all.empty else pd.DataFrame()
            count = len(items)
            pct = (count / total_items * 100) if total_items > 0 else 0
            st.markdown(f"""
            <div class="kanban-col-header" style="background:{bg}; color:{accent_color};">
                {title} ({count})
            </div>
            """, unsafe_allow_html=True)
            st.progress(min(1.0, pct / 100))
            for _, r in items.head(6).iterrows():
                st.markdown(f"""
                <div class="kanban-card" style="border-color:{accent_color};">
                    <b>{r['PN']}</b><br>
                    <span style="opacity:0.75;">{str(r['Description'])[:24]}</span><br>
                    <span style="opacity:0.6; font-size:11px;">חוסר: {r['Total_MRP_Shortage']:g}</span>
                </div>
                """, unsafe_allow_html=True)
            if count > 6:
                st.caption(f"+ עוד {count - 6} פריטים נוספים")

with tab5:
    st.markdown(f'<div class="section-title">🏭 ניהול WIP חכם (כולל סגירת מחזור ייצור ואימות היררכיה)</div>', unsafe_allow_html=True)
    st.markdown("כאן ניתן להוריד הרכבות לייצור לאחר בדיקת עץ קפדנית, או לדווח על סיום תהליך ייצור בסוף חודש (המרה אוטומטית למלאי זמין).")

    wip_current = fetch_wip_records()

    if wip_current:
        st.markdown("##### 🏁 סגירת הרכבות שהיו ב-WIP (המרתן למלאי זמין)")
        with st.form("close_wip_form"):
            wip_to_close = st.selectbox("בחר הרכבה שסיימה ייצור לחודש זה", list(wip_current.keys()), format_func=lambda x: f"{x} (כמות ב-WIP: {wip_current[x]})")
            is_finished = st.checkbox("האם ההרכבה הסתיימה לחלוטין והושלמה בהצלחה?")
            
            if st.form_submit_button("סגור WIP והוסף למלאי הזמין"):
                if is_finished:
                    closing_qty = wip_current[wip_to_close]
                    curr_stk, curr_eta, curr_stat, curr_sup, curr_comm, _, _ = get_inventory_record(wip_to_close)
                    
                    save_inventory_record(
                        pn=wip_to_close,
                        added_stock=curr_stk + closing_qty,
                        eta=curr_eta,
                        status="התקבל",
                        supplier=curr_sup,
                        comment=f"{curr_comm} | הושלם מייצור WIP בסך {closing_qty} יחידות",
                        updated_by="WIP Close Module",
                        webhook_url=webhook_url
                    )
                    delete_wip_record(wip_to_close)
                    st.success(f"ההרכבה `{wip_to_close}` טופלה בהצלחה! ה-WIP אופס והכמות ({closing_qty}) נוספה למלאי ההרכבה.")
                    st.rerun()
                else:
                    st.warning("יש לסמן את תיבת הסימון המאשרת שההרכבה אכן הסתיימה.")

    st.divider()

    with st.form("wip_form"):
        wip_asm_choice = st.selectbox("בחר הרכבה חדשה לצירוף ל-WIP", filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x))
        current_wip_val = wip_current.get(wip_asm_choice, 0.0)
        wip_qty_input = st.number_input("כמות יחידות הרכבה להורדה לייצור", min_value=0.0, value=float(current_wip_val if current_wip_val > 0 else 1.0), step=1.0)

        submitted_wip = st.form_submit_button("בדיקת זמינות היררכית מלאה ושמור WIP")

        if submitted_wip:
            inv_cache_wip = fetch_all_inventory_records()
            
            def get_all_descendants(parent_pn, all_rows_df):
                descendants = set()
                queue = [parent_pn]
                visited = set()

                while queue:
                    current_parent = queue.pop(0)
                    if current_parent in visited:
                        continue
                    visited.add(current_parent)

                    for idx, row in all_rows_df.iterrows():
                        p_num = str(row[PN_COL]).strip()
                        if p_num == current_parent:
                            continue
                        
                        if current_parent in all_rows_df.columns:
                            val = pd.to_numeric(row.get(current_parent, 0), errors='coerce')
                            if pd.notnull(val) and val > 0:
                                if p_num not in descendants:
                                    descendants.add(p_num)
                                    queue.append(p_num)
                return descendants

            def get_recursive_tree_issues(target_asm):
                issues = []
                all_child_pns = get_all_descendants(target_asm, df)

                if not all_child_pns and target_asm in df.columns:
                    direct_rows = df[df[target_asm].notnull() & (pd.to_numeric(df[target_asm], errors='coerce') > 0)]
                    for _, d_row in direct_rows.iterrows():
                        all_child_pns.add(str(d_row[PN_COL]).strip())

                for c_pn in all_child_pns:
                    if c_pn == target_asm:
                        continue
                    
                    match_row = df[df[PN_COL].astype(str).str.strip() == c_pn]
                    if match_row.empty:
                        continue
                    
                    r_data = match_row.iloc[0]
                    monthly_bal = pd.to_numeric(r_data.get(selected_month_col, 0), errors='coerce') or 0
                    added_stk, man_eta, _, _, _, _, _ = get_inventory_record(c_pn, inv_cache_wip)
                    eff_bal = monthly_bal + added_stk

                    is_late = False
                    if man_eta and str(man_eta).strip() not in ["", "None", "NaT", "nan"]:
                        try:
                            if pd.to_datetime(man_eta).strftime("%Y-%m") > selected_ym:
                                is_late = True
                        except:
                            pass

                    if eff_bal < 0 or is_late:
                        q_missing = abs(eff_bal) if eff_bal < 0 else 1.0
                        desc_text = str(r_data.get(DESC_COL, ""))
                        issues.append({
                            "PN": c_pn,
                            "Description": desc_text,
                            "Reason": f"גירעון של {q_missing:g} יח'" if eff_bal < 0 else f"עיכוב ספק (ETA: {man_eta})"
                        })

                return issues

            recursive_issues_list = get_recursive_tree_issues(wip_asm_choice)

            if recursive_issues_list:
                st.error(f"❌ שגיאה היררכית קפדנית: לא ניתן להכניס את ההרכבה `{wip_asm_choice}` ל-WIP מכיוון שחלק מתתי-ההרכבות או רכיבי הבן בשרשרת העץ שלה חסרים או באיחור!")
                st.markdown("**רכיבים ותתי-הרכבות חסרים בשרשרת העץ:**")
                for item in recursive_issues_list:
                    st.markdown(f"- מק'ט: `{item['PN']}` | תיאור: {item['Description']} | סיבה: {item['Reason']}")
                st.warning("💡 עדכן את המלאי או ה-ETA של הרכיבים הללו בלשונית 'עדכון מלאי וספקים', ורק לאחר מכן נסה שוב.")
            else:
                save_wip_record(wip_asm_choice, wip_qty_input)
                st.success(f"✅ בדיקת העץ הרקורסיבית עברה בהצלחה! כל שרשרת הבנים ותתי-ההרכבות של `{wip_asm_choice}` זמינות. ה-WIP נשמר!")
                st.rerun()

    st.markdown("##### 📋 רשימת ההרכבות הפעילות ב-WIP כרגע:")
    if wip_current:
        wip_display_rows = []
        for asm_k, qty_v in wip_current.items():
            if qty_v > 0:
                wip_display_rows.append({
                    "קוד הרכבה": asm_k,
                    "תיאור הרכבה": assembly_mapping.get(asm_k, ""),
                    "כמות ב-WIP": qty_v
                })
        if wip_display_rows:
            st.dataframe(pd.DataFrame(wip_display_rows), use_container_width=True)
            if st.button("🗑️ איפוס כל ה-WIP"):
                for asm_k in list(wip_current.keys()):
                    delete_wip_record(asm_k)
                st.success("כל נתוני ה-WIP אופסו.")
                st.rerun()
        else:
            st.info("אין כרגע הרכבות פעילות ב-WIP.")
    else:
        st.info("אין כרגע הרכבות פעילות ב-WIP.")

with tab6:
    st.markdown('<div class="section-title">📅 עדכון מלאי, סטטוס ודחיית ספקים (ETA)</div>', unsafe_allow_html=True)
    selected_pn = search_pn if search_pn != "הכל" else st.selectbox("בחר מק'ט מכלל הפריטים לעדכון", sorted(df[PN_COL].dropna().astype(str).unique()), key="update_pn_select")

    if selected_pn != "הכל":
        saved_stock, saved_eta, saved_status, saved_supplier, saved_comment, saved_by, _ = get_inventory_record(selected_pn)
        base_mrp_eta = get_base_mrp_eta(selected_pn)
        base_mrp_qty = get_base_mrp_qty(selected_pn)

        st.info(f"ℹ️ מועד ה-ETA המקורי והכמות שעלו מדוח ה-MRP (עמודות CC עד CZ) עבור מק'ט זה הם: **{base_mrp_eta}** | כמות: **{base_mrp_qty:g}**")

        with st.form("inventory_form"):
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                added_stock_input = st.number_input("תוספת למלאי זמין (קבוע)", min_value=0.0, value=float(saved_stock), step=1.0)
            with col_f2:
                try: 
                    parsed_eta = pd.to_datetime(saved_eta).date() if saved_eta else (pd.to_datetime(base_mrp_eta).date() if base_mrp_eta != "בדיקה נדרשת" else date.today())
                except: 
                    parsed_eta = date.today()
                eta_date = st.date_input("תאריך הגעה מעודכן (ETA / דחיית ספק)", value=parsed_eta)
            with col_f3:
                status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
                status_idx = status_options.index(saved_status) if saved_status in status_options else 0
                status = st.selectbox("סטטוס טיפול", status_options, index=status_idx)

            col_f4, col_f5 = st.columns(2)
            with col_f4:
                sup_idx = supplier_options.index(saved_supplier) if saved_supplier in supplier_options else 0
                supplier = st.selectbox("ספק", supplier_options, index=sup_idx)
            with col_f5:
                updated_by = st.text_input("עודכן ע'י", value=saved_by)
            comment = st.text_area("הערות (פירוט סיבת דחייה וכו')", value=saved_comment)

            if st.form_submit_button("שמור עדכון קבוע בענן"):
                save_inventory_record(selected_pn, added_stock_input, str(eta_date), status, supplier, comment, updated_by, webhook_url)
                st.success(f"העדכון למק'ט {selected_pn} נשמר בהצלחה בענן!")
                st.rerun()

with tab7:
    st.markdown('<div class="section-title">📅 מעקב ETA, דחיות, כמויות וקישורים למפיצים (Mouser / DigiKey / Findchips)</div>', unsafe_allow_html=True)
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
        
        if current_eta_raw and str(current_eta_raw).strip() not in ["", "None", "NaT", "nan"]:
            try:
                curr_eta_fmt = pd.to_datetime(current_eta_raw).strftime("%Y-%m")
            except:
                curr_eta_fmt = str(current_eta_raw)[:7]
        else:
            curr_eta_fmt = orig_eta

        if orig_eta != "בדיקה נדרשת" and curr_eta_fmt != "בדיקה נדרשת" and curr_eta_fmt != orig_eta:
            note = f"נדחה מחודש {orig_eta} לחודש {curr_eta_fmt}"
        elif curr_eta_fmt != "בדיקה נדרשת" and orig_eta == "בדיקה נדרשת":
            note = f"עודכן ל-ETA: {curr_eta_fmt}"
        else:
            note = "ללא שינוי / לפי תכנון מקורי"

        mouser_link = f"https://www.mouser.co.il/c/?q={p_num}"
        digikey_link = f"https://www.digikey.com/en/products/result?keywords={p_num}"
        findchips_link = f"https://www.findchips.com/search/{p_num}"

        eta_table_rows.append({
            "מק'ט": p_num,
            "תיאור פריט": p_desc,
            "סוג פריט": p_type,
            "ETA מקורי (MRP)": orig_eta,
            "כמות מקורית (CC-CZ)": orig_qty,
            "ETA מעודכן (בפועל)": curr_eta_fmt,
            "כמות מעודכנת": current_added_stock,
            "סטטוס ספק/דחייה": note,
            "ספק": saved_rec.get("supplier", "אופק"),
            "חיפוש במאוזר": mouser_link,
            "חיפוש בדיגיקי": digikey_link,
            "חיפוש ב-Findchips": findchips_link,
            "הערות משתמש": saved_rec.get("comment", "")
        })

    eta_df = pd.DataFrame(eta_table_rows)
    if not eta_df.empty:
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            filter_delayed_only = st.checkbox("הצג פריטים שנדחו בלבד")
        with col_s2:
            search_eta_pn = st.text_input("חיפוש חופשי לפי מק'ט או תיאור בלשונית זו", value="")

        filtered_eta_df = eta_df.copy()
        if filter_delayed_only:
            filtered_eta_df = filtered_eta_df[filtered_eta_df["סטטוס ספק/דחייה"].str.startswith("נדחה")]
        if search_eta_pn:
            filtered_eta_df = filtered_eta_df[
                filtered_eta_df["מק'ט"].str.contains(search_eta_pn, case=False, na=False) |
                filtered_eta_df["תיאור פריט"].str.contains(search_eta_pn, case=False, na=False)
            ]

        st.dataframe(
            filtered_eta_df,
            column_config={
                "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
                "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
                "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
            },
            use_container_width=True,
            height=450
        )

with tab8:
    st.markdown('<div class="section-title">↩️ חזרה לאחור וניהול היסטוריה (UNDO)</div>', unsafe_allow_html=True)
    try:
        response = supabase.table("mrp_inventory_updates").select("*").order("updated_at", desc=True).execute()
        updated_items = response.data if response.data else []
    except:
        updated_items = []

    if updated_items:
        for item in updated_items:
            i_pn = item.get("pn")
            i_stock = item.get("added_stock")
            i_eta = item.get("eta")
            i_status = item.get("status")
            i_sup = item.get("supplier")
            i_comm = item.get("comment")
            i_by = item.get("updated_by")
            i_time = item.get("updated_at")

            with st.container():
                col_u1, col_u2, col_u3 = st.columns([3, 4, 1])
                with col_u1:
                    st.markdown(f"**מק'ט:** `{i_pn}`")
                    st.text(f"ספק: {i_sup} | סטטוס: {i_status}")
                with col_u2:
                    st.text(f"תוספת: {i_stock} | ETA מעודכן: {i_eta}")
                    st.text(f"עודכן ע'י: {i_by} ({i_time})")
                with col_u3:
                    if st.button("🔄 בטל שמירה (UNDO)", key=f"undo_{i_pn}"):
                        delete_inventory_record(i_pn)
                        st.success("המידע נמחק לצמיתות מבסיס הנתונים בענן.")
                        st.rerun()
                st.divider()
    else:
        st.info("אין עדכונים קבועים במערכת כרגע.")
