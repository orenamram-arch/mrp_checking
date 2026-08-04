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
}}
.kpi-label {{ font-size: 13px; opacity: 0.75; margin-bottom: 6px; font-weight: 600; }}
.kpi-value {{ font-size: 30px; font-weight: 800; }}
.kpi-sub {{ font-size: 12px; opacity: 0.6; margin-top: 4px; }}

.kpi-green {{ border-top: 4px solid {SUCCESS}; }}
.kpi-red {{ border-top: 4px solid {DANGER}; }}
.kpi-orange {{ border-top: 4px solid {WARNING}; }}
.kpi-blue {{ border-top: 4px solid {ACCENT}; }}

.section-title {{
    font-weight: 800;
    font-size: 19px;
    margin: 18px 0 10px 0;
    border-right: 4px solid {PRIMARY};
    padding-right: 10px;
    color: var(--text-color, inherit);
}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
    <h1>🚀 MRP Executive Control Tower & Decision Hub</h1>
    <p>מערכת ניהול חוסרים מתקדמת מבוססת תוכנית עבודה מדויקת לפי סדר תצוגה מקורי</p>
</div>
""", unsafe_allow_html=True)

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
                eta_val = row.get("eta", "")
                if not eta_val or str(eta_val).strip() in ["", "None", "NaT", "nan"]:
                    eta_val = ""
                records[pn] = {
                    "added_stock": float(row.get("added_stock", 0.0) or 0.0),
                    "eta": eta_val,
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
        return (res["added_stock"], res["eta"], res["status"], res["supplier"], res["comment"], res["updated_by"], res["updated_at"])
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
    total_new_qty = existing_qty + float(wip_qty)
    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {"assembly_pn": str(assembly_pn), "wip_qty": float(total_new_qty), "updated_at": now_str}
    try:
        supabase.table("mrp_wip_assemblies").upsert(payload, on_conflict="assembly_pn").execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירת WIP: {e}")

def delete_wip_record(assembly_pn):
    try:
        supabase.table("mrp_wip_assemblies").delete().eq("assembly_pn", str(assembly_pn)).execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקת WIP: {e}")

def save_inventory_record(pn, added_stock, eta, status, supplier, comment, updated_by):
    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {
        "pn": str(pn), "added_stock": float(added_stock), "eta": str(eta),
        "status": str(status), "supplier": str(supplier), "comment": str(comment),
        "updated_by": str(updated_by), "updated_at": now_str
    }
    try:
        supabase.table("mrp_inventory_updates").upsert(payload, on_conflict="pn").execute()
        supabase.table("mrp_inventory_history").insert(payload).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירה לענן: {e}")

def delete_inventory_record(pn):
    try:
        supabase.table("mrp_inventory_updates").delete().eq("pn", str(pn)).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקה: {e}")

# ==========================================================
# DATA LOADING & EXACT ORDER PRESERVATION
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
    st.error(f"שגיאה בטעינת הקובץ: {e}")
    st.stop()

# טעינת תוכנית העבודה תוך סריקת כלל השורות האפשריות ושמירה על סדר ההצגה המקורי
if "custom_assembly_plan_df" not in st.session_state:
    header_dates = df_raw.iloc[2, 108:132].values if df_raw.shape[1] > 132 else []
    plan_rows = []
    ordered_assemblies_list = []

    # סריקה מורחבת של כל השורות כדי לוודא שכל ההרכבות (כולל התחתונות כמו REAR COVER) נכללות
    for r in range(3, df_raw.shape[0]):
        asm_pn = df_raw.iloc[r, 106] if df_raw.shape[1] > 106 else None
        if pd.notnull(asm_pn):
            clean_asm_pn = str(asm_pn).strip()
            if clean_asm_pn and clean_asm_pn != 'nan':
                if clean_asm_pn not in ordered_assemblies_list:
                    ordered_assemblies_list.append(clean_asm_pn)

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
                                    if ym_str >= "2026-09":
                                        plan_rows.append({
                                            "Assembly_PN": clean_asm_pn,
                                            "YearMonth": ym_str,
                                            "Build_Qty": q_val * system_multiplier,
                                            "Raw_Build_Qty": q_val
                                        })
                            except:
                                pass

    st.session_state["custom_assembly_plan_df"] = pd.DataFrame(plan_rows)
    st.session_state["ordered_assemblies"] = ordered_assemblies_list

assembly_plan_df = st.session_state["custom_assembly_plan_df"]
ordered_assemblies = st.session_state.get("ordered_assemblies", [])

PN_COL = df.columns[1]
DESC_COL = df.columns[4]
ITEM_TYPE_COL = df.columns[44] if len(df.columns) > 44 else df.columns[-1]
STOCK_COL = df.columns[79] if len(df.columns) > 79 else df.columns[-1]
ASSEMBLY_COLS = df.columns[10:36].tolist()
MONTH_COLS = df.columns[108:132].tolist() if len(df.columns) > 132 else []

valid_assemblies = [col for col in ASSEMBLY_COLS if col in df[PN_COL].values] or ordered_assemblies

assembly_levels = {}
for col in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(col)
        assembly_levels[col] = int(df_levels.iloc[0, col_idx])
    except:
        assembly_levels[col] = 0

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

def get_first_supply_eta(pn, inv_cache=None):
    _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        return manual_eta
    eta, _ = get_base_mrp_eta_and_qty(pn)
    return eta

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("🔍 מסננים מתקדמים")
if st.sidebar.button("🧹 איפוס כל המסננים"):
    st.rerun()

month_options = {}
for m in MONTH_COLS:
    if pd.notnull(m):
        try:
            dt = pd.to_datetime(m)
            if dt.strftime("%Y-%m") >= "2026-09":
                month_options[dt.strftime("%B %Y (%Y-%m)")] = m
        except:
            pass

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()))
selected_month_col = month_options[selected_month_label]
selected_ym = pd.to_datetime(selected_month_col).strftime("%Y-%m")
num_months_ahead = st.sidebar.slider("📅 טווח מבט קדימה בחודשים", 1, 6, 1)

assembly_mapping = {col: f"{col} - {df_desc.iloc[0, df.columns.get_loc(col)] if col in df.columns else ''}" for col in valid_assemblies}
selected_assembly = st.sidebar.selectbox("בחר הרכבה ספציפית", ["הכל"] + valid_assemblies, format_func=lambda x: assembly_mapping.get(x, x))

# ==========================================================
# CORE LOGIC FOR SHORTAGES
# ==========================================================
all_ym_list = sorted(list(set(assembly_plan_df["YearMonth"].unique())))
start_idx = next((i for i, ym in enumerate(all_ym_list) if ym >= selected_ym), 0)
selected_target_yms = all_ym_list[start_idx:start_idx + num_months_ahead] or [selected_ym]

def calculate_mrp_breakdown(target_yms=None):
    if target_yms is None:
        target_yms = selected_target_yms
    inv_cache = fetch_all_inventory_records()
    wip_cache = fetch_wip_records()

    target_month_cols_map = {pd.to_datetime(m_c).strftime("%Y-%m"): m_c for m_c in MONTH_COLS if pd.notnull(m_c)}
    temp_df = df.copy()
    shortage_records = {}

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        saved_stock, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        max_sh = 0.0
        is_short = False

        for ym in target_yms:
            col_name = target_month_cols_map.get(ym)
            if col_name and col_name in temp_df.columns:
                mrp_val = pd.to_numeric(row[col_name], errors='coerce') or 0
                eff = mrp_val + saved_stock if mrp_val < 0 else mrp_val
                if eff < 0:
                    is_short = True
                    max_sh = max(max_sh, abs(eff))
        if is_short:
            shortage_records[idx] = max_sh

    temp_df['Monthly_Balance'] = temp_df.index.map(lambda i: -shortage_records[i] if i in shortage_records else 1.0)
    mrp_shortages = temp_df[temp_df['Monthly_Balance'] < 0].copy()
    mrp_shortages['Total_MRP_Shortage'] = mrp_shortages['Monthly_Balance'].abs()

    month_plan = assembly_plan_df[assembly_plan_df["YearMonth"].isin(target_yms)]
    plan_dict = month_plan.groupby("Assembly_PN")["Raw_Build_Qty"].sum().to_dict()

    breakdown_rows = []
    for idx, row in mrp_shortages.iterrows():
        pn = str(row[PN_COL]).strip()
        desc = str(row[DESC_COL])
        stock = (pd.to_numeric(row[STOCK_COL], errors='coerce') or 0) + get_inventory_record(pn, inv_cache)[0]
        tot_sh = row['Total_MRP_Shortage']

        for asm in valid_assemblies:
            if asm in temp_df.columns:
                qty_per = pd.to_numeric(row[asm], errors='coerce') or 0
                if qty_per > 0:
                    breakdown_rows.append({
                        "PN": pn, "Description": desc, "Assembly": asm,
                        "Qty_Per_Assembly": qty_per, "Total_MRP_Shortage": tot_sh, "Stock": stock
                    })
    return pd.DataFrame(breakdown_rows)

breakdown_df = calculate_mrp_breakdown()

# ==========================================================
# TABS DEFINITION
# ==========================================================
tab1, tab2, tab10 = st.tabs(["📈 Executive Dashboard", "📊 תוכנית ייצור (Smart CTB)", "🎯 ניתוח רגישות ותוכנית"])

with tab2:
    st.markdown(f'<div class="section-title">📊 סימולציית Clear To Build (CTB) מטריציונית בדיוק לפי סדר ההרכבות</div>', unsafe_allow_html=True)

    inv_cache_ctb = fetch_all_inventory_records()
    wip_cache_ctb = fetch_wip_records()
    matrix_rows = []

    # מעבר מדויק על ההרכבות בסדר הנכון
    for asm_col in valid_assemblies:
        try:
            asm_desc = df_desc.iloc[0, df.columns.get_loc(asm_col)]
        except:
            asm_desc = ""

        row_data = {
            "קוד הרכבה": asm_col,
            "תיאור הרכבה": asm_desc,
            "רמה בעץ": assembly_levels.get(asm_col, 0)
        }
        has_build = False

        for target_m in selected_target_yms:
            sub_plan = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan["Build_Qty"].sum() if not sub_plan.empty else 0.0
            wip_qty = wip_cache_ctb.get(asm_col, 0.0)

            if raw_build > 0 or wip_qty > 0:
                has_build = True

            row_data[f"תכנית עבודה ({target_m})"] = raw_build
            row_data[f"ניתן לייצור ({target_m})"] = max(0.0, raw_build - wip_qty)
            row_data[f"WIP ({target_m})"] = wip_qty
            row_data[f"סטטוס ({target_m})"] = "✅ מוכן לייצור" if raw_build > 0 else "💤 ללא תוכנית"

        if has_build or selected_assembly == "הכל":
            matrix_rows.append(row_data)

    if matrix_rows:
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, height=500)
    else:
        st.info("לא נמצאו נתונים להצגה.")

with tab10:
    st.markdown('<div class="section-title">🎯 מטריצת תוכנית העבודה המלאה (לפי סדר התצוגה המדויק)</div>', unsafe_allow_html=True)
    if not assembly_plan_df.empty:
        pivot_plan_df = assembly_plan_df.pivot_table(
            index=["Assembly_PN"], columns="YearMonth", values="Build_Qty", fill_value=0.0
        ).reindex(valid_assemblies).reset_index()

        pivot_plan_df.insert(1, "תיאור הרכבה", pivot_plan_df["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        st.dataframe(pivot_plan_df, use_container_width=True, height=500)
    else:
        st.info("אין נתוני תוכנית עבודה זמינים.")
