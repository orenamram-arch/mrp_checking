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

st.title("📊 דשבורד ניתוח חוסרים מתקדם - MRP")
st.markdown("ניתוח חוסרים פר פריט מול הרכבות (כמות נדרשת להרכבה × תוכנית ייצור חודשית)")

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
PN_COL = df.columns[1]     # מק"ט (PN_ID)
DESC_COL = df.columns[4]   # תיאור פריט
ITEM_TYPE_COL = df.columns[44] # סוג פריט (עמודה AS)
STOCK_COL = df.columns[79]     # מלאי (עמודה CB)

# עמודות ההרכבות (K עד AJ) - אינדקסים 10 עד 35
ASSEMBLY_COLS = df.columns[10:36].tolist()

# עמודות מאזן החומרים / חודשים (DE עד EB)
MONTH_COLS = df.columns[108:132].tolist()

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("🔍 מסננים")

month_options = {str(m): m for m in MONTH_COLS if pd.notnull(m)}
selected_month_str = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()))
selected_month = month_options[selected_month_str]

# מיפוי הרכבות כולל תיאור
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
# ADVANCED BREAKDOWN ENGINE (Part per Assembly mapping)
# ==========================================================
# נבנה טבלה מפורטת המפרטת לכל פריט את הדרישה שלו בכל הרכבה
breakdown_rows = []

for idx, row in df.iterrows():
    pn = row[PN_COL]
    desc = row[DESC_COL]
    item_type = row[ITEM_TYPE_COL]
    stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
    
    for asm in ASSEMBLY_COLS:
        # כמות הפריט הנדרשת להרכבה הספציפית
        qty_per_asm = pd.to_numeric(row[asm], errors='coerce') or 0
        if qty_per_asm > 0:
            # נרצה לחלץ את תוכנית הייצור של ההרכבה בחודש הנבחר אם קיימת, 
            # לחלופין נסתמך על החישוב הקיים בשורת הפריט או נחשב לפי העניין
            asm_desc = assembly_mapping.get(asm, asm)
            breakdown_rows.append({
                "PN": pn,
                "Description": desc,
                "Item_Type": item_type,
                "Assembly": asm,
                "Assembly_Desc": asm_desc,
                "Qty_Per_Assembly": qty_per_asm,
                "Stock": stock
            })

breakdown_df = pd.DataFrame(breakdown_rows)

# סינון לפי סוג פריט
if selected_item_type != "הכל":
    breakdown_df = breakdown_df[breakdown_df["Item_Type"] == selected_item_type]

# סינון לפי הרכבה
if selected_assembly != "הכל":
    breakdown_df = breakdown_df[breakdown_df["Assembly"] == selected_assembly]

# חישוב מאזן חומרים כולל לפריט מהעמודה החודשית הנבחרת
df['Total_Balance'] = pd.to_numeric(df[selected_month], errors='coerce').fillna(0)
balance_map = df.set_index(PN_COL)['Total_Balance'].to_dict()

breakdown_df["Monthly_Balance"] = breakdown_df["PN"].map(balance_map)
# פריטים בחוסר הם בעלי מאזן שלילי
shortage_breakdown = breakdown_df[breakdown_df["Monthly_Balance"] < 0].copy()
shortage_breakdown["Shortage_Qty"] = shortage_breakdown["Monthly_Balance"].abs()

# ==========================================================
# DASHBOARD UI
# ==========================================================
st.subheader(f"ניתוח חוסרים מפורט לפי הרכבה לחודש: {selected_month_str}")

col1, col2 = st.columns(2)
col1.metric("🔴 שורות חוסר משוייכות", len(shortage_breakdown))
col2.metric("📦 סך כמות חסרה", f"{shortage_breakdown['Shortage_Qty'].sum():,.0f}")

st.divider()

if len(shortage_breakdown) > 0:
    st.subheader("📋 פירוט חוסרים לפי פריט והרכבה")
    
    display_df = shortage_breakdown[[
        "PN", "Description", "Item_Type", "Assembly", "Assembly_Desc", 
        "Qty_Per_Assembly", "Stock", "Shortage_Qty"
    ]].rename(columns={
        "PN": "מק\"ט",
        "Description": "תיאור פריט",
        "Item_Type": "סוג פריט (AS)",
        "Assembly": "קוד הרכבה",
        "Assembly_Desc": "תיאור הרכבה",
        "Qty_Per_Assembly": "כמות נדרשת להרכבה",
        "Stock": "מלאי נוכחי",
        "Shortage_Qty": "חוסר מחושב"
    })
    
    st.dataframe(display_df.sort_values(by="חוסר מחושב", ascending=False), use_container_width=True)
else:
    st.success("🎉 לא נמצאו חוסרים עבור הסינונים שנבחרו לחודש זה!")

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
