import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import datetime, date

# ==========================================================
# CONFIGURATION
# ==========================================================
# הקישור הישיר (Raw) לקובץ מתוך GitHub שסיפקת
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"

# שם הקובץ המקומי שייווצר אוטומטית לשמירת העדכונים שלך
LOCAL_DB_FILE = "eta_updates.db" 

st.set_page_config(
    page_title="MRP Control Tower",
    page_icon="📦",
    layout="wide"
)

st.title("📊 דשבורד ניתוח חוסרים - MRP")
st.markdown("קריאת נתונים מ-GitHub ושמירת עדכוני סטטוס בקובץ מקומי")

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
    
    df.columns = [str(c).strip() if pd.notnull(c) else c for c in df.columns]
    return df, df_levels, df_desc

try:
    with st.spinner('טוען נתוני MRP מ-GitHub...'):
        df, df_levels, df_desc = load_data(GITHUB_URL)
except Exception as e:
    st.error(f"שגיאה בטעינת הקובץ מ-GitHub. ודא שהקישור הוא מסוג Raw ושהמאגר ציבורי.\nפירוט השגיאה: {e}")
    st.stop()

# ==========================================================
# COLUMN MAPPING
# ==========================================================
PN_COL = df.columns[1]     
DESC_COL = df.columns[4]   
ASSEMBLY_COLS = df.columns[10:36].tolist()
ITEM_TYPE_COL = df.columns[44]
STOCK_COL = df.columns[79]
MONTH_COLS = df.columns[108:132].tolist()

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("🔍 מסננים")

month_options = {str(m): m for m in MONTH_COLS if pd.notnull(m)}
selected_month_str = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()))
selected_month = month_options[selected_month_str]

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
selected_item_type = st.sidebar.selectbox("בחר סוג פריט", ["הכל"] + item_types)

# ==========================================================
# APPLY FILTERS & CALCULATE SHORTAGES
# ==========================================================
filtered_df = df.copy()

if selected_item_type != "הכל":
    filtered_df = filtered_df[filtered_df[ITEM_TYPE_COL] == selected_item_type]

if selected_assembly != "הכל":
    filtered_df[selected_assembly] = pd.to_numeric(filtered_df[selected_assembly], errors='coerce').fillna(0)
    filtered_df = filtered_df[filtered_df[selected_assembly] > 0]

filtered_df['Balance'] = pd.to_numeric(filtered_df[selected_month], errors='coerce').fillna(0)
shortage_df = filtered_df[filtered_df['Balance'] < 0].copy()
shortage_df['Shortage_Qty'] = shortage_df['Balance'].abs()

# ==========================================================
# DASHBOARD UI: SHORTAGES
# ==========================================================
st.subheader(f"ניתוח חוסרים לחודש: {selected_month_str}")

col1, col2, col3 = st.columns(3)
total_shortage_items = len(shortage_df)
total_shortage_qty = shortage_df['Shortage_Qty'].sum()

col1.metric("🔴 פריטים בחוסר", total_shortage_items)
col2.metric("📦 כמות חסרה כוללת", f"{total_shortage_qty:,.0f}")

if selected_assembly != "הכל":
    try:
        assembly_index = df.columns.get_loc(selected_assembly)
        bom_level = df_levels.iloc[0, assembly_index]
        col3.metric("🔗 רמת הרכבה (BOM Level)", str(bom_level))
    except:
        col3.metric("🔗 רמת הרכבה (BOM Level)", "לא נמצא")

st.divider()

if total_shortage_items > 0:
    st.subheader("📋 רשימת חוסרים מפורטת (מתוך ה-MRP)")
    
    display_cols = [PN_COL, DESC_COL, ITEM_TYPE_COL, STOCK_COL, 'Shortage_Qty']
    display_df = shortage_df[display_cols].sort_values(by='Shortage_Qty', ascending=False)
    display_df = display_df.rename(columns={
        PN_COL: 'מק"ט',
        DESC_COL: 'תיאור',
        ITEM_TYPE_COL: 'סוג פריט',
        STOCK_COL: 'מלאי נוכחי',
        'Shortage_Qty': 'כמות חסרה'
    })
    st.dataframe(display_df, use_container_width=True)
else:
    st.success("🎉 לא נמצאו חוסרים עבור הסינונים שנבחרו!")

# ==========================================================
# ETA MANAGEMENT (WRITES TO LOCAL FILE)
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
