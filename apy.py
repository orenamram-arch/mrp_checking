import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import datetime, date
import io

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

st.title("📊 MRP Control Tower & Shortage Management")
st.markdown("מערכת ניהול חוסרים מתקדמת הכוללת פילוח הרכבות, מעקב ספקים (אופק), גרפי מגמות וייצוא נתונים")

# ==========================================================
# LOCAL DATABASE SETUP (For ETA & Supplier Updates)
# ==========================================================
conn = sqlite3.connect(LOCAL_DB_FILE, check_same_thread=False)

conn.execute("""
CREATE TABLE IF NOT EXISTS eta_updates
(
    pn TEXT PRIMARY KEY,
    eta TEXT,
    status TEXT,
    supplier TEXT,
    comment TEXT,
    updated_by TEXT,
    updated_at TEXT
)
""")
conn.commit()

def get_eta_record(pn):
    cur = conn.cursor()
    # תמיכה לאחור בטבלאות ישנות ללא עמודת ספק
    try:
        cur.execute("SELECT eta, status, supplier, comment, updated_by, updated_at FROM eta_updates WHERE pn = ?", (pn,))
        res = cur.fetchone()
        if res:
            return res[0], res[1], res[2], res[3], res[4], res[5]
    except:
        cur.execute("SELECT eta, status, comment, updated_by, updated_at FROM eta_updates WHERE pn = ?", (pn,))
        res = cur.fetchone()
        if res:
            return res[0], res[1], "אופק (ברירת מחדל)", res[2], res[3], res[4]
    return None

def save_eta_record(pn, eta, status, supplier, comment, updated_by):
    conn.execute("""
    INSERT OR REPLACE INTO eta_updates (pn, eta, status, supplier, comment, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (pn, eta, status, supplier, comment, updated_by, datetime.now().strftime("%Y-%m-%d %H:%M")))
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
                            dt = pd.to_datetime(date_val)
                            plan_rows.append({
                                "Assembly_PN": str(asm_pn).strip(),
                                "YearMonth": dt.strftime("%Y-%m"),
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
MONTH_COLS = df.columns[108:132].tolist()

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("🔍 מסננים ראשיים")

month_options = {}
for m in MONTH_COLS:
    if pd.notnull(m):
        try:
            dt = pd.to_datetime(m)
            month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
        except:
            month_options[str(m)] = m

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()))
selected_month_col = month_options[selected_month_label]

try:
    selected_ym = pd.to_datetime(selected_month_col).strftime("%Y-%m")
except:
    selected_ym = str(selected_month_col)[:7]

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

# חיפוש חופשי מהיר
quick_search = st.sidebar.text_input("🔎 חיפוש מהיר (מק\"ט / תיאור)", "")

# ==========================================================
# CORE LOGIC: MRP SHORTAGES + MULTI-ROW ASSEMBLY EXPANSION
# ==========================================================
df['Monthly_Balance'] = pd.to_numeric(df[selected_month_col], errors='coerce').fillna(0)
mrp_shortages = df[df['Monthly_Balance'] < 0].copy()
mrp_shortages['Total_MRP_Shortage'] = mrp_shortages['Monthly_Balance'].abs()

month_plan = assembly_plan_df[assembly_plan_df["YearMonth"] == selected_ym]
plan_dict = month_plan.set_index("Assembly_PN")["Build_Qty"].to_dict()

breakdown_rows = []

for idx, row in mrp_shortages.iterrows():
    pn = str(row[PN_COL]).strip()
    desc = str(row[DESC_COL])
    item_type = str(row[ITEM_TYPE_COL])
    stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
    total_mrp_shortage = row['Total_MRP_Shortage']
    
    matched_any = False
    for asm in ASSEMBLY_COLS:
        qty_per_asm = pd.to_numeric(row[asm], errors='coerce') or 0
        if qty_per_asm > 0:
            matched_any = True
            asm_build_qty = plan_dict.get(asm, 0.0)
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
                "Stock": stock,
                "Total_MRP_Shortage": total_mrp_shortage
            })
            
    if not matched_any:
        breakdown_rows.append({
            "PN": pn,
            "Description": desc,
            "Item_Type": item_type,
            "Assembly": "ללא שיוך",
            "Assembly_Desc": "ללא שיוך להרכבה ראשית",
            "Qty_Per_Assembly": 0,
            "Assembly_Monthly_Build": 0,
            "Required_Demand": 0,
            "Stock": stock,
            "Total_MRP_Shortage": total_mrp_shortage
        })

breakdown_df = pd.DataFrame(breakdown_rows)

# החלת סינונים
if not breakdown_df.empty:
    if selected_item_type != "הכל":
        breakdown_df = breakdown_df[breakdown_df["Item_Type"] == selected_item_type]
    if selected_assembly != "הכל":
        breakdown_df = breakdown_df[breakdown_df["Assembly"] == selected_assembly]
    if quick_search:
        q = quick_search.lower()
        breakdown_df = breakdown_df[
            breakdown_df["PN"].str.lower().str.contains(q, na=False) | 
            breakdown_df["Description"].str.lower().str.contains(q, na=False)
        ]

# ==========================================================
# DASHBOARD UI & KPIS
# ==========================================================
st.subheader(f"📈 ניתוח חוסרים לחודש: {selected_month_label}")

col1, col2, col3 = st.columns(3)
col1.metric("🔴 פריטים בחוסר ב-MRP", len(mrp_shortages))
col2.metric("📋 סך שורות פריט-הדדי להרכבה", len(breakdown_df) if not breakdown_df.empty else 0)
total_req_demand = breakdown_df['Required_Demand'].sum() if not breakdown_df.empty else 0
col3.metric("📦 סך ביקוש מחושב בהרכבות", f"{total_req_demand:,.0f}")

st.divider()

if not breakdown_df.empty and len(breakdown_df) > 0:
    # --- טבלה ראשית להצגה ---
    st.subheader("📋 פירוט חוסרים מלא ופילוח מול הרכבות")
    
    display_df = breakdown_df[[
        "PN", "Description", "Item_Type", "Assembly", "Assembly_Desc", 
        "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage"
    ]].rename(columns={
        "PN": "מק\"ט",
        "Description": "תיאור פריט",
        "Item_Type": "סוג פריט (AS)",
        "Assembly": "קוד הרכבה",
        "Assembly_Desc": "תיאור הרכבה",
        "Qty_Per_Assembly": "כמות נדרשת להרכבה",
        "Assembly_Monthly_Build": "ת. ייצור הרכבה לחודש",
        "Required_Demand": "ביקוש מדויק להרכבה",
        "Stock": "מלאי נוכחי",
        "Total_MRP_Shortage": "סך חוסר ב-MRP"
    })
    
    # עיצוב מותנה (Highlighting) וצביעת שורות לפי גובה החוסר
    def highlight_shortage(s):
        return ['background-color: #ffcccc' if v > 1000 else '' for v in s]

    st.dataframe(
        display_df.sort_values(by="סך חוסר ב-MRP", ascending=False).style.apply(highlight_shortage, subset=['סך חוסר ב-MRP']), 
        use_container_width=True
    )

    # --- כפתור ייצוא ל-Excel ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        display_df.to_excel(writer, index=False, sheet_name='Shortages_Breakdown')
    processed_data = output.getvalue()

    st.download_button(
        label="📥 הורד את הטבלה המוצגת לקובץ Excel",
        data=processed_data,
        file_name=f"MRP_Shortages_{selected_ym}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    # ==========================================================
    # TREND ANALYSIS CHART (גרף מגמות חוסרים רוחבי)
    # ==========================================================
    st.subheader("📊 גרף מגמות: סך חוסרים מצטבר לאורך החודשים הבאים")
    
    trend_data = []
    for m_label, m_col in month_options.items():
        temp_bal = pd.to_numeric(df[m_col], errors='coerce').fillna(0)
        sum_shortage = temp_bal[temp_bal < 0].abs().sum()
        # חילוץ שם חודש קצר לתצוגה בגרף
        short_name = str(m_label).split('(')[0].strip()
        trend_data.append({"חודש": short_name, "סך חוסר כספי/כמותי": sum_shortage})
    
    trend_df = pd.DataFrame(trend_data)
    fig_trend = px.bar(trend_df, x="חודש", y="סך חוסר כספי/כמותי", text_auto='.2s', color="סך חוסר כספי/כמותי", color_continuous_scale="Reds")
    fig_trend.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_trend, use_container_width=True)

else:
    st.success("🎉 לא נמצאו חוסרים ב-MRP עבור הסינונים שנבחרו לחודש זה!")

# ==========================================================
# ETA & SUPPLIER MANAGEMENT (עם מעקב קבלני משנה כמו אופק)
# ==========================================================
st.divider()
st.subheader("📅 ניהול ומעקב ETA וקבלני משנה (נשמר מקומית במסד הנתונים)")

pn_values = sorted(df[PN_COL].dropna().astype(str).unique())
selected_pn = st.selectbox("בחר מק\"ט לעדכון סטטוס וספק", pn_values)

existing_rec = get_eta_record(selected_pn)
def_eta, def_status, def_supplier, def_comment, def_by = (
    (existing_rec[0], existing_rec[1], existing_rec[2], existing_rec[3], existing_rec[4]) 
    if existing_rec else (date.today(), "פתוח", "אופק", "", "")
)

with st.form("eta_form"):
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        try:
            parsed_eta = pd.to_datetime(def_eta).date() if def_eta else date.today()
        except:
            parsed_eta = date.today()
        eta_date = st.date_input("תאריך הגעה משוער (ETA)", value=parsed_eta)
    with col_f2:
        status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
        status_idx = status_options.index(def_status) if def_status in status_options else 0
        status = st.selectbox("סטטוס", status_options, index=status_idx)
    with col_f3:
        supplier_options = ["אופק", "ספק פנימי", "רכש אחר", "אחר"]
        sup_idx = supplier_options.index(def_supplier) if def_supplier in supplier_options else 0
        supplier = st.selectbox("קבלן משנה / ספק", supplier_options, index=sup_idx)

    comment = st.text_area("הערות מעקב", value=def_comment)
    updated_by = st.text_input("עודכן על ידי", value=def_by)
    save_btn = st.form_submit_button("שמור עדכון במערכת")

if save_btn:
    save_eta_record(selected_pn, str(eta_date), status, supplier, comment, updated_by)
    st.success("העדכון נשמר בהצלחה!")

# טבלת מעקב מרכזית
st.subheader("🚦 טבלת סטטוסים וספקים שמורים במערכת")
eta_rows = []
for pn in pn_values:
    rec = get_eta_record(pn)
    if rec:
        eta_rows.append({
            "מק\"ט": pn,
            "ETA": rec[0],
            "סיכון": eta_color(rec[0]),
            "סטטוס": rec[1],
            "ספק / קב\"מ": rec[2],
            "הערות": rec[3],
            "אחראי": rec[4],
            "תאריך עדכון": rec[5]
        })

eta_df = pd.DataFrame(eta_rows)
if len(eta_df) > 0:
    st.dataframe(eta_df, use_container_width=True)
else:
    st.info("עדיין לא נשמרו עדכונים במערכת.")
