import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import datetime, date

# ==========================================================
# CONFIGURATION
# ==========================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"
LOCAL_DB_FILE = "eta_updates.db" 

st.set_page_config(
    page_title="MRP Control Tower",
    page_icon="📦",
    layout="wide"
)

st.title("📊 דשבורד ניתוח חוסרים מדויק לפי הרכבה - MRP")
st.markdown("חישוב ביקוש מדויק: כמות הפריט להרכבה × תוכנית הייצור החודשית של ההרכבה")

# ==========================================================
# LOCAL DATABASE SETUP (For ETA Updates)
# ==========================================================
conn = sqlite3.connect(LOCAL_DB_FILE, check_same_thread=False)

conn.execute("""
CREATE TABLE IF NOT EXISTS eta_updates
(
    pn TEXT PRIMARY KEY,
    eta TEXT,
    status TEXT,
    comment TEXT,
    updated_by TEXT,
    updated_at TEXT
)
""")
conn.commit()

def get_eta_record(pn):
    cur = conn.cursor()
    cur.execute("SELECT eta, status, comment, updated_by, updated_at FROM eta_updates WHERE pn = ?", (pn,))
    return cur.fetchone()

def save_eta_record(pn, eta, status, comment, updated_by):
    conn.execute("""
    INSERT OR REPLACE INTO eta_updates (pn, eta, status, comment, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (pn, eta, status, comment, updated_by, datetime.now().strftime("%Y-%m-%d %H:%M")))
    conn.commit()

def eta_color(eta_value):
    if eta_value in [None, "", "NaT"]:
        return "⚪"
    try:
        eta_date = pd.to_datetime(eta_value).date()
        days = (eta_date - date.today()).days
        if days < 0:
            return "🔴"
        if days <= 14:
            return "🟠"
        if days <= 30:
            return "🟡"
        return "🟢"
    except:
        return "⚪"

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
    st.error(f"שגיאה בטעינת הקובץ מ-GitHub. ודא שהקישור הוא מסוג Raw ושהמאגר ציבורי.\nפירוט השגיאה: {e}")
    st.stop()

# ==========================================================
# EXTRACT ASSEMBLY MONTHLY BUILD PLAN (Rows 3 to 23)
# ==========================================================
header_dates = df_raw.iloc[2, 108:132].values
plan_rows = []

for r in range(3, 24):
    asm_pn = df_raw.iloc[r, 106]
    asm_name = df_raw.iloc[r, 104]
    if pd.notnull(asm_pn):
        for c_idx, date_val in enumerate(header_dates):
            if pd.notnull(date_val):
                qty = df_raw.iloc[r, 108 + c_idx]
                if pd.notnull(qty) and qty != '' and qty != 'NaN':
                    try:
                        q_val = float(qty)
                        if q_val > 0:
                            plan_rows.append({
                                "Assembly_PN": str(asm_pn).strip(),
                                "Month": pd.to_datetime(date_val).strftime("%Y-%m-%d"),
                                "Build_Qty": q_val
                            })
                    except:
                        pass

assembly_plan_df = pd.DataFrame(plan_rows)

# ==========================================================
# COLUMN MAPPING
# ==========================================================
PN_COL = df.columns[1]     # מק"ט (PN_ID)
DESC_COL = df.columns[4]   # תיאור פריט
ITEM_TYPE_COL = df.columns[44] # סוג פריט (עמודה AS)
STOCK_COL = df.columns[79]     # מלאי (עמודה CB)

ASSEMBLY_COLS = df.columns[10:36].tolist()
available_months = sorted(assembly_plan_df["Month"].unique().tolist()) if not assembly_plan_df.empty else []

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("🔍 מסננים")

if available_months:
    selected_month = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", available_months)
else:
    st.sidebar.error("לא נמצאו תאריכים בתוכנית הייצור")
    st.stop()

assembly_mapping = {"הכל": "הכל"}
for col in ASSEMBLY_COLS:
    try:
        col_idx = df.columns.get_loc(col)
        desc = df_desc.iloc[0, col_idx]
        assembly_mapping[col] = f"{col} - {desc}"
    except:
        assembly_mapping[col] = col

selected_assembly = st.sidebar.selectbox(
    "בחר הרכבה (Assembly)", 
    ["הכל"] + ASSEMBLY_COLS,
    format_func=lambda x: assembly_mapping.get(x, x)
)

item_types = df[ITEM_TYPE_COL].dropna().unique().tolist()
selected_item_type = st.sidebar.selectbox("בחר סוג פריט (עמודה AS)", ["הכל"] + item_types)

# ==========================================================
# PRECISE CALCULATION ENGINE
# ==========================================================
# שליפת תוכנית הייצור להרכבות בחודש הנבחר
month_plan = assembly_plan_df[assembly_plan_df["Month"] == selected_month]
plan_dict = month_plan.set_index("Assembly_PN")["Build_Qty"].to_dict()

breakdown_rows = []

for idx, row in df.iterrows():
    pn = str(row[PN_COL]).strip()
    desc = row[DESC_COL]
    item_type = row[ITEM_TYPE_COL]
    stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
    
    for asm in ASSEMBLY_COLS:
        qty_per_asm = pd.to_numeric(row[asm], errors='coerce') or 0
        if qty_per_asm > 0:
            asm_build_qty = plan_dict.get(asm, 0.0)
            if asm_build_qty > 0:
                # דרישה מדויקת: כמות הפריט להרכבה × תוכנית הייצור החודשית להרכבה זו
                required_demand = qty_per_asm * asm_build_qty
                asm_desc = assembly_mapping.get(asm, asm)
                
                breakdown_rows.append({
                    "PN": pn,
                    "Description": desc,
                    "Item_Type": item_type,
                    "Assembly": asm,
                    "Assembly_Desc": asm_desc,
                    "Qty_Per_Assembly": qty_per_asm,
                    "Assembly_Monthly_Build": asm_build_qty,
                    "Required_Demand": required_demand,
                    "Stock": stock
                })

breakdown_df = pd.DataFrame(breakdown_rows)

if not breakdown_df.empty:
    # סינון לפי סוג פריט
    if selected_item_type != "הכל":
        breakdown_df = breakdown_df[breakdown_df["Item_Type"] == selected_item_type]

    # סינון לפי הרכבה
    if selected_assembly != "הכל":
        breakdown_df = breakdown_df[breakdown_df["Assembly"] == selected_assembly]

# ==========================================================
# DASHBOARD UI
# ==========================================================
st.subheader(f"ניתוח ביקוש מדויק לפי הרכבה לחודש: {selected_month}")

col1, col2 = st.columns(2)
col1.metric("🔴 שורות דרישה להרכבות", len(breakdown_df) if not breakdown_df.empty else 0)

total_demand = breakdown_df['Required_Demand'].sum() if not breakdown_df.empty else 0
col2.metric("📦 סך ביקוש חודשי בהרכבות", f"{total_demand:,.0f}")

st.divider()

if not breakdown_df.empty and len(breakdown_df) > 0:
    st.subheader("📋 פירוט דרישות וחוסרים פר פריט מול כל הרכבה")
    
    display_df = breakdown_df[[
        "PN", "Description", "Item_Type", "Assembly", "Assembly_Desc", 
        "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock"
    ]].rename(columns={
        "PN": "מק\"ט",
        "Description": "תיאור פריט",
        "Item_Type": "סוג פריט (AS)",
        "Assembly": "קוד הרכבה",
        "Assembly_Desc": "תיאור הרכבה",
        "Qty_Per_Assembly": "כמות נדרשת להרכבה",
        "Assembly_Monthly_Build": "ת. ייצור הרכבה לחודש",
        "Required_Demand": "ביקוש מדויק להרכבה",
        "Stock": "מלאי נוכחי"
    })
    
    st.dataframe(display_df.sort_values(by="ביקוש מדויק להרכבה", ascending=False), use_container_width=True)
else:
    st.success("🎉 לא נמצאו נתוני ייצור או דרישות עבור הסינונים שנבחרו לחודש זה!")

# ==========================================================
# ETA MANAGEMENT
# ==========================================================
st.divider()
st.subheader("📅 ניהול ומעקב ETA (נשמר מקומית)")

pn_values = sorted(df[PN_COL].dropna().astype(str).unique())
selected_pn = st.selectbox("בחר מק\"ט לעדכון סטטוס", pn_values)
record = get_eta_record(selected_pn)
default_comment = record[2] if record else ""

with st.form("eta_form"):
    eta_date = st.date_input("תאריך הגעה משוער (ETA)")
    status = st.selectbox(
        "סטטוס",
        ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
    )
    comment = st.text_area("הערות", value=default_comment)
    updated_by = st.text_input("עודכן על ידי")
    save_btn = st.form_submit_button("שמור עדכון")

if save_btn:
    save_eta_record(selected_pn, str(eta_date), status, comment, updated_by)
    st.success("העדכון נשמר בהצלחה בקובץ המקומי!")

st.subheader("🚦 טבלת סטטוסים שמורים")
eta_rows = []
for pn in pn_values:
    rec = get_eta_record(pn)
    if rec:
        eta_rows.append({
            "מק\"ט": pn,
            "ETA": rec[0],
            "סיכון": eta_color(rec[0]),
            "סטטוס": rec[1],
            "הערות": rec[2],
            "אחראי": rec[3],
            "תאריך עדכון": rec[4]
        })

eta_df = pd.DataFrame(eta_rows)
if len(eta_df) > 0:
    st.dataframe(eta_df, use_container_width=True)
else:
    st.info("עדיין לא נשמרו עדכונים במערכת.")
