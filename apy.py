import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import datetime, date
import io
import requests
import json

# ==========================================================
# CONFIGURATION & LOGIN SYSTEM
# ==========================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"
LOCAL_DB_FILE = "eta_updates.db" 

st.set_page_config(
    page_title="MRP Control Tower & Enterprise Portal",
    page_icon="📦",
    layout="wide"
)

# מנגנון אבטחה והתחברות בסיסי
def check_password():
    def password_entered():
        if st.session_state["password"] == "ELTA2026":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("הכנס סיסמת כניסה למערכת:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("הכנס סיסמת כניסה למערכת:", type="password", on_change=password_entered, key="password")
        st.error("😕 סיסמה שגויה")
        return False
    else:
        return True

if not check_password():
    st.stop()

st.title("📊 MRP Control Tower & Enterprise Portal")
st.markdown("מערכת ניהול חוסרים מתקדמת הכוללת פורטל ספקים, התראות אוטומטיות וניתוח צווארי בקבוק בעץ המוצר")

# ==========================================================
# LOCAL DATABASE SETUP (Audit Trail & Suppliers)
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

conn.execute("""
CREATE TABLE IF NOT EXISTS eta_history
(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pn TEXT,
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
    try:
        cur.execute("SELECT eta, status, supplier, comment, updated_by, updated_at FROM eta_updates WHERE pn = ?", (pn,))
        res = cur.fetchone()
        if res:
            return res[0], res[1], res[2], res[3], res[4], res[5]
    except:
        pass
    return None

def send_teams_notification(webhook_url, message):
    """שליחת התראה אוטומטית ל-Teams / Slack / Webhook"""
    if not webhook_url:
        return False
    try:
        payload = {"text": message}
        response = requests.post(webhook_url, data=json.dumps(payload), headers={'Content-Type': 'application/json'})
        return response.status_code == 200
    except:
        return False

def save_eta_record(pn, eta, status, supplier, comment, updated_by, webhook_url=""):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn.execute("""
    INSERT OR REPLACE INTO eta_updates (pn, eta, status, supplier, comment, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (pn, eta, status, supplier, comment, updated_by, now_str))
    
    conn.execute("""
    INSERT INTO eta_history (pn, eta, status, supplier, comment, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (pn, eta, status, supplier, comment, updated_by, now_str))
    
    conn.commit()
    
    # שליחת התראה אם הוגדר Webhook
    if webhook_url:
        msg = f"🔔 עדכון MRP חדש במערכת!\nמק\"ט: {pn}\nסטטוס: {status}\nספק: {supplier}\nETA: {eta}\nעודכן ע"י: {updated_by}"
        send_teams_notification(webhook_url, msg)

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
# EXTRACT ASSEMBLY MONTHLY BUILD PLAN & BOM LEVELS
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
PN_COL = df.columns[1]     
DESC_COL = df.columns[4]   
ITEM_TYPE_COL = df.columns[44] 
STOCK_COL = df.columns[79]     

ASSEMBLY_COLS = df.columns[10:36].tolist()
MONTH_COLS = df.columns[108:132].tolist()

# ==========================================================
# SIDEBAR & PORTAL MODE CONFIGURATION
# ==========================================================
st.sidebar.header("⚙️ מצב מערכת וניהול")
portal_mode = st.sidebar.radio("בחר תצוגה:", ["מנהל מערכת מלא", "פורטל קבלן משנה (אופק בלבד)"])

webhook_url = st.sidebar.text_input("🔗 Teams / Slack Webhook URL (אופציונלי)", value="")

# הגדרת ספקים
supplier_options = ["אופק", "ספק פנימי", "רכש אחר", "אחר"]

if portal_mode == "פורטל קבלן משנה (אופק בלבד)":
    st.warning("🔒 אתה נמצא כעת בתצוגת פורטל קבלן משנה (אופק). מוצגים אך ורק פריטים שמשוייכים לאופק.")
    selected_supplier_filter = "אופק"
else:
    selected_supplier_filter = "הכל"

st.sidebar.header("🔍 מסננים מתקדמים")

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

level_options = ["הכל"] + sorted([str(df_levels.iloc[0, df.columns.get_loc(c)]) for c in ASSEMBLY_COLS if pd.notnull(df_levels.iloc[0, df.columns.get_loc(c)])])
selected_level = st.sidebar.selectbox("סינון לפי רמת עץ (BOM Level)", level_options)

assembly_mapping = {"הכל": "הכל"}
filtered_assembly_cols = []
for col in ASSEMBLY_COLS:
    try:
        col_idx = df.columns.get_loc(col)
        lvl = str(df_levels.iloc[0, col_idx])
        desc = df_desc.iloc[0, col_idx]
        if selected_level == "הכל" or lvl == selected_level:
            filtered_assembly_cols.append(col)
            assembly_mapping[col] = f"{col} - {desc} (רמה {lvl})"
    except:
        filtered_assembly_cols.append(col)
        assembly_mapping[col] = col

selected_assembly = st.sidebar.selectbox(
    "בחר הרכבה (Assembly)", 
    ["הכל"] + filtered_assembly_cols,
    format_func=lambda x: assembly_mapping.get(x, x)
)

item_types = df[ITEM_TYPE_COL].dropna().unique().tolist()
selected_item_type = st.sidebar.selectbox("בחר סוג פריט (עמודה AS)", ["הכל"] + item_types)

quick_search = st.sidebar.text_input("🔎 חיפוש מהיר (מק\"ט / תיאור)", "")

# ==========================================================
# CORE LOGIC
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
    
    # שליפת הספק הרשום במסד הנתונים עבור פריט זה
    rec_check = get_eta_record(pn)
    current_sup = rec_check[2] if rec_check else "אופק"
    
    # סינון לפי ספק אם במצב פורטל
    if selected_supplier_filter != "הכל" and current_sup != selected_supplier_filter:
        continue

    matched_any = False
    for asm in filtered_assembly_cols:
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
                "Supplier": current_sup,
                "Assembly": asm,
                "Assembly_Desc": asm_desc,
                "Qty_Per_Assembly": qty_per_asm,
                "Assembly_Monthly_Build": asm_build_qty,
                "Required_Demand": required_demand,
                "Stock": stock,
                "Total_MRP_Shortage": total_mrp_shortage
            })
            
    if not matched_any and selected_assembly == "הכל" and selected_level == "הכל":
        breakdown_rows.append({
            "PN": pn,
            "Description": desc,
            "Item_Type": item_type,
            "Supplier": current_sup,
            "Assembly": "ללא שיוך",
            "Assembly_Desc": "ללא שיוך להרכבה ראשית",
            "Qty_Per_Assembly": 0,
            "Assembly_Monthly_Build": 0,
            "Required_Demand": 0,
            "Stock": stock,
            "Total_MRP_Shortage": total_mrp_shortage
        })

breakdown_df = pd.DataFrame(breakdown_rows)

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
# TABS FOR ENTERPRISE FEATURES (כולל ניתוח צווארי בקבוק)
# ==========================================================
tab1, tab2, tab3 = st.tabs(["📦 דשבורד חוסרים ראשי", "⚠️ ניתוח צווארי בקבוק (Bottlenecks)", "📅 מעקב ETA וספקים"])

with tab1:
    st.subheader(f"📈 ניתוח חוסרים לחודש: {selected_month_label}")

    col1, col2, col3 = st.columns(3)
    col1.metric("🔴 פריטים בחוסר ב-MRP", len(mrp_shortages))
    col2.metric("📋 סך שורות פריט-הדדי להרכבה", len(breakdown_df) if not breakdown_df.empty else 0)
    total_req_demand = breakdown_df['Required_Demand'].sum() if not breakdown_df.empty else 0
    col3.metric("📦 סך ביקוש מחושב בהרכבות", f"{total_req_demand:,.0f}")

    st.divider()

    if not breakdown_df.empty and len(breakdown_df) > 0:
        st.subheader("📋 פירוט חוסרים מלא ופילוח מול הרכבות")
        
        display_df = breakdown_df[[
            "PN", "Description", "Item_Type", "Supplier", "Assembly", "Assembly_Desc", 
            "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage"
        ]].rename(columns={
            "PN": "מק\"ט",
            "Description": "תיאור פריט",
            "Item_Type": "סוג פריט (AS)",
            "Supplier": "ספק / קב\"מ",
            "Assembly": "קוד הרכבה",
            "Assembly_Desc": "תיאור הרכבה",
            "Qty_Per_Assembly": "כמות נדרשת להרכבה",
            "Assembly_Monthly_Build": "ת. ייצור הרכבה לחודש",
            "Required_Demand": "ביקוש מדויק להרכבה",
            "Stock": "מלאי נוכחי",
            "Total_MRP_Shortage": "סך חוסר ב-MRP"
        })
        
        def highlight_shortage(s):
            return ['background-color: #ffcccc' if v > 1000 else '' for v in s]

        st.dataframe(
            display_df.sort_values(by="סך חוסר ב-MRP", ascending=False).style.apply(highlight_shortage, subset=['סך חוסר ב-MRP']), 
            use_container_width=True
        )

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
    else:
        st.success("🎉 לא נמצאו חוסרים ב-MRP עבור הסינונים שנבחרו לחודש זה!")

with tab2:
    st.subheader("⚠️ ניתוח צווארי בקבוק (Bottleneck Analysis) בעץ המוצר")
    st.markdown("סקירה אוטומטית של רכיבים המופיעים במספר הרב ביותר של הרכבות וגורמים להשפעת רוחב על הייצור:")
    
    # אלגוריתם זיהוי צווארי בקבוק: ספירת הופעות של כל מק"ט בכל עמודות ההרכבות
    bottleneck_rows = []
    for idx, row in df.iterrows():
        pn = str(row[PN_COL]).strip()
        desc = str(row[DESC_COL])
        count_assemblies = 0
        total_qty_needed = 0
        
        for asm in ASSEMBLY_COLS:
            q = pd.to_numeric(row[asm], errors='coerce') or 0
            if q > 0:
                count_assemblies += 1
                total_qty_needed += q
                
        if count_assemblies > 1: # פריט שמופיע ביותר מהרכבה אחת
            bottleneck_rows.append({
                "מק\"ט": pn,
                "תיאור": desc,
                "מספר הרכבות שבהן משתתף": count_assemblies,
                "סך כמות נדרשת במצטבר": total_qty_needed
            })
            
    if bottleneck_rows:
        bn_df = pd.DataFrame(bottleneck_rows).sort_values(by="מספר הרכבות שבהן משתתף", ascending=False)
        st.dataframe(bn_df.head(20), use_container_width=True)
        st.info("💡 טיפ: רכיבים המופיעים במספר רב של הרכבות הם צווארי בקבוק פוטנציאליים. עדיפות עליונה להבטיח את זמינותם ברכש.")
    else:
        st.info("לא נמצאו פריטים משותפים למספר הרכבות.")

with tab3:
    st.subheader("📅 מעקב ETA וספקים (כולל היסטוריית שינויים ועדכונים)")

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
            sup_idx = supplier_options.index(def_supplier) if def_supplier in supplier_options else 0
            supplier = st.selectbox("קבלן משנה / ספק", supplier_options, index=sup_idx)

        comment = st.text_area("הערות מעקב", value=def_comment)
        updated_by = st.text_input("עודכן על ידי", value=def_by)
        save_btn = st.form_submit_button("שמור עדכון במערכת ושלח התראה")

    if save_btn:
        save_eta_record(selected_pn, str(eta_date), status, supplier, comment, updated_by, webhook_url)
        st.success("העדכון, ההיסטוריה וההתראה נשמרו בהצלחה!")

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

    st.subheader(f"📜 היסטוריית שינויים (Audit Trail) עבור מק\"ט: {selected_pn}")
    history_cur = conn.cursor()
    history_cur.execute("SELECT eta, status, supplier, comment, updated_by, updated_at FROM eta_history WHERE pn = ? ORDER BY id DESC", (selected_pn,))
    hist_rows = history_cur.fetchall()
    if hist_rows:
        hist_df = pd.DataFrame(hist_rows, columns=["ETA", "סטטוס", "ספק", "הערות", "עודכן על ידי", "זמן עדכון"])
        st.dataframe(hist_df, use_container_width=True)
    else:
        st.info("אין היסטוריית שינויים קודמת למק\"ט זה.")
