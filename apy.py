"""
MRP Control Tower — מגדל בקרת חוסרים
גרסה מותאמת ומהירה: שליפת נתונים מרוכזת מ-Supabase (במקום פניות בודדות) למניעת איטיות.
גרסה 2 — עיצוב משודרג, גרפים נוספים, KPI-cards מעוצבים והתאמה אוטומטית למצב בהיר/כהה במכשיר.

הרצה:
streamlit run mrp_app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
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
# GLOBAL THEME / CSS (AUTO LIGHT/DARK MODE SUPPORT)
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

/* הגדרת RTL לתוכן הראשי ולסיידבר */
[data-testid="stAppViewContainer"] .main .block-container,
[data-testid="stSidebarContent"] {{
    font-family: 'Assistant', sans-serif;
    direction: rtl;
}}
[data-testid="stAppViewContainer"] .main .block-container * ,
[data-testid="stSidebarContent"] * {{
    font-family: 'Assistant', sans-serif;
}}

/* Hide default streamlit chrome */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

/* Header banner - צבעי גרדיאנט קבועים שיישארו בולטים ויפים בכל תמה */
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

/* KPI cards - התאמה אוטומטית למצב בהיר וכהה לפי העדפת המכשיר/דפדפן */
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

/* מניעת מסך שחור במובייל: גיבוי למקרה שהמשתמש נמצא במצב בהיר במכשיר אך הדפדפן כופה רקע */
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

/* Section titles */
.section-title {{
    font-weight: 800;
    font-size: 19px;
    margin: 18px 0 10px 0;
    border-right: 4px solid {PRIMARY};
    padding-right: 10px;
    color: var(--text-color, inherit);
}}

/* Kanban columns */
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

.stMetric {{ border: 1px solid rgba(128,128,128,0.2); border-radius: 8px; padding: 10px; }}

/* ===================== מובייל ===================== */
@media (max-width: 640px) {{
    .hero-banner {{ padding: 18px 16px; border-radius: 14px; }}
    .hero-banner h1 {{ font-size: 21px; }}
    .hero-banner p {{ font-size: 13px; }}

    .kpi-value {{ font-size: 22px; }}
    .kpi-label {{ font-size: 12px; }}
    .kpi-card {{ padding: 14px 10px; }}

    .section-title {{ font-size: 16px; }}

    /* גורם לעמודות (KPI / גרפים) להיערם אנכית זו מתחת לזו במקום להצטופף לרוחב */
    [data-testid="stHorizontalBlock"] {{
        flex-wrap: wrap !important;
    }}
    [data-testid="column"] {{
        min-width: 100% !important;
        flex: 1 1 100% !important;
    }}
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

# התאמה אוטומטית גם לגרפים של Plotly לפי התמה של המערכת/מכשיר
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

def save_inventory_record(pn, added_stock, eta, status, supplier, comment, updated_by, webhook_url=""):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
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
        msg = "🔔 עדכון מלאי/ETA למוצר!\nמק\"ט: " + str(pn) + "\nתוספת מלאי: " + str(added_stock) + "\nסטטוס: " + str(status) + "\nETA: " + str(eta)
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

def get_first_supply_eta(pn, inv_cache=None):
    _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        return manual_eta

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
                            dt = pd.to_datetime(date_val, errors='coerce')
                            if pd.notnull(dt) and dt.year >= 2024:
                                return dt.strftime("%Y-%m")
            except:
                pass

    return "בדיקה נדרשת"

# ==========================================================
# SIDEBAR FILTERS & WHAT-IF CONTROLS
# ==========================================================
st.sidebar.header("⚙️ הגדרות מערכת וחיבור")
webhook_url = st.sidebar.text_input("🔗 Teams / Slack Webhook URL (אופציונלי)", value="")
supplier_options = ["אופק", "ספק פנימי", "רכש אחר", "אחר"]

st.sidebar.header("🔍 מסננים מתקדמים")

month_options = {}
for m in MONTH_COLS:
    if pd.notnull(m):
        try:
            dt = pd.to_datetime(m)
            month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
        except:
            month_options[str(m)] = m

if not month_options:
    month_options["ברירת מחדל"] = df.columns[108] if len(df.columns) > 108 else df.columns[-1]

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()))
selected_month_col = month_options[selected_month_label]

try:
    selected_ym = pd.to_datetime(selected_month_col).strftime("%Y-%m")
except:
    selected_ym = str(selected_month_col)[:7]

level_options = ["הכל"] + sorted([str(df_levels.iloc[0, df.columns.get_loc(c)]) for c in valid_assemblies if pd.notnull(df_levels.iloc[0, df.columns.get_loc(c)])])
selected_level = st.sidebar.selectbox("סינון לפי רמת עץ (BOM Level)", level_options)

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
    format_func=lambda x: assembly_mapping.get(x, x)
)

item_types = df[ITEM_TYPE_COL].dropna().unique().tolist() if ITEM_TYPE_COL in df.columns else []
selected_item_type = st.sidebar.selectbox("בחר סוג פריט (עמודה AS)", ["הכל"] + item_types)

item_choices = ["הכל"] + sorted([f"{str(r[PN_COL]).strip()} - {str(r[DESC_COL])}" for _, r in df.iterrows() if pd.notnull(r[PN_COL])])
selected_search_item = st.sidebar.selectbox("🔎 חיפוש מהיר (בחר או הקלד מק\"ט/תיאור)", item_choices)
search_pn = selected_search_item.split(" - ")[0] if selected_search_item != "הכל" else "הכל"


# ==========================================================
# CORE LOGIC FOR SHORTAGES (WRAPPED IN FUNCTION)
# ==========================================================
def calculate_mrp_breakdown(sim_extra_stock=None):
    if sim_extra_stock is None:
        sim_extra_stock = {}

    inv_cache = fetch_all_inventory_records()

    temp_df = df.copy()
    temp_df['Monthly_Balance'] = pd.to_numeric(temp_df[selected_month_col], errors='coerce').fillna(0)

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        sim_val = sim_extra_stock.get(pn, 0.0)

        total_added_stock = saved_stock_add + sim_val

        if total_added_stock > 0:
            current_bal = temp_df.at[idx, 'Monthly_Balance']
            if current_bal < 0:
                temp_df.at[idx, 'Monthly_Balance'] = current_bal + total_added_stock

    mrp_shortages = temp_df[temp_df['Monthly_Balance'] < 0].copy()
    mrp_shortages['Total_MRP_Shortage'] = mrp_shortages['Monthly_Balance'].abs()

    month_plan = assembly_plan_df[assembly_plan_df["YearMonth"] == selected_ym]
    plan_dict = month_plan.set_index("Assembly_PN")["Build_Qty"].to_dict()

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
        matched_any = False

        for asm in filtered_assembly_cols:
            qty_per_asm = pd.to_numeric(row[asm], errors='coerce') or 0
            if qty_per_asm > 0:
                matched_any = True
                asm_build_qty = plan_dict.get(asm, 0.0)
                required_demand = qty_per_asm * asm_build_qty
                asm_desc = assembly_mapping.get(asm, asm)

                breakdown_rows.append({
                    "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                    "Status": item_status, "Assembly": asm, "Assembly_Desc": asm_desc, "Qty_Per_Assembly": qty_per_asm,
                    "Assembly_Monthly_Build": asm_build_qty, "Required_Demand": required_demand,
                    "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage
                })

        if not matched_any and selected_assembly == "הכל" and selected_level == "הכל":
            breakdown_rows.append({
                "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                "Status": item_status, "Assembly": "ללא שיוך", "Assembly_Desc": "ללא שיוך להרכבה", "Qty_Per_Assembly": 0,
                "Assembly_Monthly_Build": 0, "Required_Demand": 0, "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage
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

breakdown_df = calculate_mrp_breakdown()

# ==========================================================
# TABS
# ==========================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📈 Executive Dashboard",
    "📊 תוכנית ייצור (Smart CTB)",
    "💡 סימולציית What-If",
    "📌 לוח סטטוסים (Kanban)",
    "📅 עדכון מלאי וספקים",
    "↩️ ניהול UNDO"
])

with tab1:
    st.markdown(f'<div class="section-title">🎯 תמונת מצב ניהולית לחודש: {selected_month_label}</div>', unsafe_allow_html=True)

    dash_df = breakdown_df.copy()
    if selected_assembly != "הכל":
        dash_df = dash_df[dash_df["Assembly"] == selected_assembly]

    total_planned_assemblies = len([a for a in valid_assemblies if assembly_plan_df[(assembly_plan_df["YearMonth"] == selected_ym) & (assembly_plan_df["Assembly_PN"] == a)]["Build_Qty"].sum() > 0])
    blocked_assemblies = len(dash_df['Assembly'].unique()) if not dash_df.empty else 0
    ready_assemblies = max(0, total_planned_assemblies - blocked_assemblies)
    readiness_pct = (ready_assemblies / total_planned_assemblies * 100) if total_planned_assemblies > 0 else 100

    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    with col_k1:
        kpi_card("🟢 מוכנות קווי ייצור", f"{readiness_pct:.1f}%", f"{ready_assemblies}/{total_planned_assemblies} הרכבות מוכנות", "green")
    with col_k2:
        kpi_card("🔴 הרכבות חסומות", blocked_assemblies, "בחודש הנבחר", "red")
    with col_k3:
        kpi_card("📦 מק\"טים בגירעון", len(dash_df['PN'].unique()) if not dash_df.empty else 0, "פריטים ייחודיים", "orange")
    with col_k4:
        kpi_card("📊 כמות גירעון מצטברת", f"{dash_df['Total_MRP_Shortage'].sum():,.0f}" if not dash_df.empty else "0", "יחידות", "blue")

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

        col_g3, col_g4 = st.columns(2)

        with col_g3:
            st.markdown("##### 🔝 Top 10 מק\"טים עם החוסר הגדול ביותר")
            top10 = (dash_df.drop_duplicates(subset=["PN"])
                            .sort_values("Total_MRP_Shortage", ascending=False)
                            .head(10))
            fig_top = px.bar(top10, x="Total_MRP_Shortage", y="PN", orientation="h",
                             color="Total_MRP_Shortage", color_continuous_scale=["#F59E0B", "#EF4444"],
                             text="Total_MRP_Shortage", hover_data=["Description", "Supplier", "Status"])
            fig_top.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
            fig_top.update_layout(template=PLOTLY_TEMPLATE, height=380, yaxis={'categoryorder': 'total ascending'},
                                  margin=dict(t=10, b=10, l=10, r=10), coloraxis_showscale=False)
            st.plotly_chart(fig_top, use_container_width=True)

        with col_g4:
            st.markdown("##### 📈 מגמת חוסרים רוחבית לאורך חודשי השנה")
            inv_cache_trend = fetch_all_inventory_records()
            trend_rows = []
            for m_col in MONTH_COLS:
                if pd.notnull(m_col):
                    try:
                        m_dt = pd.to_datetime(m_col)
                        m_ym = m_dt.strftime("%Y-%m")
                        temp_b = pd.to_numeric(df[m_col], errors='coerce').fillna(0)

                        for idx_temp, row_temp in df.iterrows():
                            pn_temp = str(row_temp[PN_COL]).strip()
                            saved_stk, _, _, _, _, _, _ = get_inventory_record(pn_temp, inv_cache_trend)
                            if saved_stk > 0:
                                val_temp = temp_b.loc[idx_temp]
                                if val_temp < 0:
                                    temp_b.loc[idx_temp] = val_temp + saved_stk

                        tot_sh = temp_b[temp_b < 0].abs().sum()
                        trend_rows.append({"Month": m_ym, "Total_Shortage": tot_sh})
                    except:
                        pass
            trend_df = pd.DataFrame(trend_rows)
            if not trend_df.empty:
                fig_line = px.area(trend_df, x="Month", y="Total_Shortage", markers=True,
                                   color_discrete_sequence=[ACCENT])
                fig_line.update_traces(line=dict(width=3))
                fig_line.update_layout(template=PLOTLY_TEMPLATE, height=380, margin=dict(t=10, b=10, l=10, r=10))
                st.plotly_chart(fig_line, use_container_width=True)

        st.markdown("##### 🔻 פילוח סטטוס טיפול בפריטים חסרים")
        status_counts = dash_df.drop_duplicates(subset=["PN"])["Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        if not status_counts.empty:
            fig_funnel = px.funnel(status_counts.sort_values("Count", ascending=False), x="Count", y="Status",
                                   color_discrete_sequence=[PRIMARY])
            fig_funnel.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_funnel, use_container_width=True)

        st.markdown('<div class="section-title">📋 טבלת פירוט ניהולית עם אפשרות ייצוא</div>', unsafe_allow_html=True)
        display_df = dash_df[[
            "PN", "Description", "Item_Type", "Supplier", "Status", "Assembly", "Assembly_Desc",
            "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage"
        ]].rename(columns={
            "PN": "מק\"ט", "Description": "תיאור פריט", "Item_Type": "סוג פריט", "Supplier": "ספק",
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
        st.dataframe(styled, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='Executive_Shortages')
        processed_data = output.getvalue()

        st.download_button(
            label="📥 הורד דו\"ח מנהלים מלא ל-Excel",
            data=processed_data,
            file_name=f"MRP_Executive_Report_{selected_ym}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("🎉 אין חוסרים ב-MRP עבור ההגדרות והסינונים שנבחרו!")

with tab2:
    st.markdown(f'<div class="section-title">📊 סימולציית Clear To Build (CTB) לחודש: {selected_month_label}</div>', unsafe_allow_html=True)
    st.markdown("המערכת מציגה את מועד ה-ETA האמיתי והמדויק של הרכיבים החסרים ומדגישה ב-**BOLD** את הפריט הקריטי.")

    inv_cache_ctb = fetch_all_inventory_records()

    assemblies_to_check = [asm for asm in valid_assemblies if assembly_plan_df[(assembly_plan_df["YearMonth"] == selected_ym) & (assembly_plan_df["Assembly_PN"] == asm)]["Build_Qty"].sum() > 0]
    assemblies_to_check.sort(key=lambda x: assembly_levels.get(x, 0), reverse=True)

    production_capacity_rows = []

    for asm_col in assemblies_to_check:
        if selected_assembly != "הכל" and asm_col != selected_assembly:
            continue

        try:
            asm_desc = df_desc.iloc[0, df.columns.get_loc(asm_col)]
        except:
            asm_desc = ""

        planned_build = assembly_plan_df[(assembly_plan_df["YearMonth"] == selected_ym) & (assembly_plan_df["Assembly_PN"] == asm_col)]["Build_Qty"].sum()
        asm_shortages = breakdown_df[breakdown_df["Assembly"] == asm_col] if not breakdown_df.empty else pd.DataFrame()

        missing_items_details = []
        for _, s_row in asm_shortages.iterrows():
            c_pn = str(s_row["PN"]).strip()
            c_desc = str(s_row["Description"]).strip()
            s_qty = s_row["Total_MRP_Shortage"]

            eta_display_str = get_first_supply_eta(c_pn, inv_cache_ctb)

            if eta_display_str and eta_display_str != "בדיקה נדרשת":
                try:
                    eta_dt = pd.to_datetime(eta_display_str).date()
                except:
                    eta_dt = date(2099, 12, 31)
            else:
                eta_dt = date(2099, 12, 31)

            missing_items_details.append((c_pn, c_desc, s_qty, eta_dt, eta_display_str))

        if missing_items_details:
            missing_items_details.sort(key=lambda x: x[3], reverse=True)
            most_critical_pn = missing_items_details[0][0]
        else:
            most_critical_pn = None

        formatted_missing = []
        for c_pn, c_desc, m_qty, _, raw_eta in missing_items_details:
            eta_str = f" [ETA: {raw_eta}]" if raw_eta != "בדיקה נדרשת" else ""
            item_text = f"{c_pn} ({c_desc[:12]}) - חסר: {m_qty:g}{eta_str}"

            if c_pn == most_critical_pn:
                formatted_missing.append(f"**{item_text}**")
            else:
                formatted_missing.append(item_text)

        missing_str = " | ".join(formatted_missing) if formatted_missing else "אין חוסרים! ניתן לייצר את כל התוכנית."
        max_buildable = 0 if formatted_missing else planned_build

        production_capacity_rows.append({
            "קוד הרכבה": asm_col,
            "תיאור הרכבה": asm_desc,
            "רמה בעץ": assembly_levels.get(asm_col, 0),
            "תוכנית ייצור": planned_build,
            "ניתן לייצר בפועל (CTB)": max_buildable,
            "רכיבים חסרים בלבד (הקריטי ב-BOLD)": missing_str
        })

    if production_capacity_rows:
        cap_df = pd.DataFrame(production_capacity_rows)

        col_c1, col_c2 = st.columns([1, 1.4])
        with col_c1:
            st.markdown("##### 🚦 מוכנות הרכבות (CTB מול תוכנית)")
            cap_df["גירעון"] = cap_df["תוכנית ייצור"] - cap_df["ניתן לייצר בפועל (CTB)"]
            fig_ctb = px.bar(cap_df, x="קוד הרכבה", y=["ניתן לייצר בפועל (CTB)", "גירעון"],
                             color_discrete_sequence=[SUCCESS, DANGER], barmode="stack")
            fig_ctb.update_layout(template=PLOTLY_TEMPLATE, height=340, margin=dict(t=10, b=10, l=10, r=10),
                                  legend_title_text="")
            st.plotly_chart(fig_ctb, use_container_width=True)
        with col_c2:
            st.markdown("##### 📋 טבלת CTB מפורטת")
            st.dataframe(cap_df.drop(columns=["גירעון"]), use_container_width=True, height=340)
    else:
        st.info(f"לא נמצאו הרכבות מתוכננות לייצור לחודש {selected_month_label}.")

with tab3:
    st.markdown('<div class="section-title">💡 סימולציית What-If (מה יקרה אם...)</div>', unsafe_allow_html=True)
    st.markdown("כלי אינטראקטיבי לבחינת תרחישים. **שימו לב: תוספת מלאי בלשונית זו מחושבת באופן רגעי ואינה נשמרת בבסיס הנתונים.**")

    col_w1, col_w2 = st.columns(2)
    with col_w1:
        sim_pn = st.selectbox("בחר מק\"ט לסימולציה", sorted(df[PN_COL].dropna().astype(str).unique()), key="sim_pn")
    with col_w2:
        sim_extra_stock = st.number_input("תוספת כמות מדומיינת למלאי לצורך סימולציה", min_value=0.0, value=10.0, step=1.0)

    if st.button("🔮 הרץ סימולציית שחרור צוואר בקבוק"):
        sim_df = calculate_mrp_breakdown({sim_pn: sim_extra_stock})

        orig_blocked = set(breakdown_df['Assembly'].unique()) if not breakdown_df.empty else set()
        sim_blocked = set(sim_df['Assembly'].unique()) if not sim_df.empty else set()
        freed_assemblies = orig_blocked - sim_blocked

        st.success(f"סימולציה הופעלה בהצלחה עבור מק\"ט `{sim_pn}` עם תוספת מדומיינת של {sim_extra_stock} יחידות.")

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            kpi_card("🟢 הרכבות שהשתחררו", len(freed_assemblies), "מוכנות לייצור מלא", "green")
        with col_m2:
            kpi_card("🔴 עדיין חסום", len(sim_blocked), "הרכבות שנותרו חסומות", "red")
        with col_m3:
            before_after_delta = (breakdown_df['Total_MRP_Shortage'].sum() if not breakdown_df.empty else 0) - (sim_df['Total_MRP_Shortage'].sum() if not sim_df.empty else 0)
            kpi_card("📉 צמצום גירעון כולל", f"{before_after_delta:,.0f}", "יחידות", "blue")

        st.divider()
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.markdown("### 🟢 קווי הרכבה שישתחררו לייצור מלא:")
            if freed_assemblies:
                for asm in freed_assemblies:
                    st.success(f"✔️ **{assembly_mapping.get(asm, asm)}** - מוכנה לייצור!")
            else:
                st.info("אין הרכבות שעוברות למוכנות מלאה בעקבות תוספת זו (יתכן שנדרשים פריטים נוספים).")

        with col_res2:
            st.markdown("### 🔴 צווארי בקבוק חלופיים שעדיין תוקעים הרכבות אלו:")
            if not sim_df.empty:
                related_asm = breakdown_df[breakdown_df["PN"] == sim_pn]["Assembly"].unique()
                remaining = sim_df[sim_df["Assembly"].isin(related_asm)]
                if not remaining.empty:
                    for _, r in remaining.iterrows():
                        st.warning(f"מק\"ט: `{r['PN']}` בהרכבה {r['Assembly']} - חסר: {r['Total_MRP_Shortage']:g}")
                else:
                    st.success("אין פריטים נוספים שתוקעים את ההרכבות הרלוונטיות!")
            else:
                st.success("הסימולציה שחררה את כלל הפריטים במערכת!")

        if not breakdown_df.empty or not sim_df.empty:
            st.markdown("##### ⚖️ השוואת גירעון: לפני מול אחרי הסימולציה")
            comp_df = pd.DataFrame({
                "מצב": ["לפני סימולציה", "אחרי סימולציה"],
                "סך גירעון": [
                    breakdown_df['Total_MRP_Shortage'].sum() if not breakdown_df.empty else 0,
                    sim_df['Total_MRP_Shortage'].sum() if not sim_df.empty else 0
                ]
            })
            fig_comp = px.bar(comp_df, x="מצב", y="סך גירעון", color="מצב",
                             color_discrete_sequence=[DANGER, SUCCESS], text="סך גירעון")
            fig_comp.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
            fig_comp.update_layout(template=PLOTLY_TEMPLATE, height=320, showlegend=False,
                                   margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_comp, use_container_width=True)

with tab4:
    st.markdown('<div class="section-title">📌 לוח מעקב סטטוסים (Kanban Pipeline)</div>', unsafe_allow_html=True)
    st.markdown("מעקב ויזואלי אחר התקדמות הטיפול במק\"טים הגירעוניים מול ספקים ורכש.")

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
    st.markdown('<div class="section-title">📅 עדכון מלאי וסטטוס (שמירה קבועה בענן)</div>', unsafe_allow_html=True)
    st.markdown("**הזנת הנתונים בלשונית זו תשמור אותם באופן קבוע בבסיס הנתונים בענן ותשפיע מיידית על כל החישובים.**")

    selected_pn = search_pn if search_pn != "הכל" else st.selectbox("בחר מק\"ט מכלל הפריטים לעדכון", sorted(df[PN_COL].dropna().astype(str).unique()))

    if selected_pn != "הכל":
        saved_stock, saved_eta, saved_status, saved_supplier, saved_comment, saved_by, _ = get_inventory_record(selected_pn)
        with st.form("inventory_form"):
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                added_stock_input = st.number_input("תוספת למלאי זמין (קבוע)", min_value=0.0, value=float(saved_stock), step=1.0)
            with col_f2:
                try: parsed_eta = pd.to_datetime(saved_eta).date() if saved_eta else date.today()
                except: parsed_eta = date.today()
                eta_date = st.date_input("תאריך הגעה (ETA)", value=parsed_eta)
            with col_f3:
                status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
                status_idx = status_options.index(saved_status) if saved_status in status_options else 0
                status = st.selectbox("סטטוס טיפול", status_options, index=status_idx)

            col_f4, col_f5 = st.columns(2)
            with col_f4:
                sup_idx = supplier_options.index(saved_supplier) if saved_supplier in supplier_options else 0
                supplier = st.selectbox("ספק", supplier_options, index=sup_idx)
            with col_f5:
                updated_by = st.text_input("עודכן ע\"י", value=saved_by)
            comment = st.text_area("הערות", value=saved_comment)

            if st.form_submit_button("שמור עדכון קבוע בענן"):
                save_inventory_record(selected_pn, added_stock_input, str(eta_date), status, supplier, comment, updated_by, webhook_url)
                st.success(f"העדכון למק\"ט {selected_pn} נשמר בהצלחה בענן!")
                st.rerun()

with tab6:
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
                    st.markdown(f"**מק\"ט:** `{i_pn}`")
                    st.text(f"ספק: {i_sup} | סטטוס: {i_status}")
                with col_u2:
                    st.text(f"תוספת: {i_stock} | ETA: {i_eta}")
                    st.text(f"עודכן ע\"י: {i_by} ({i_time})")
                with col_u3:
                    if st.button("🔄 בטל שמירה (UNDO)", key=f"undo_{i_pn}"):
                        delete_inventory_record(i_pn)
                        st.success("המידע נמחק לצמיתות מבסיס הנתונים בענן.")
                        st.rerun()
                st.divider()
    else:
        st.info("אין עדכונים קבועים במערכת כרגע.")
