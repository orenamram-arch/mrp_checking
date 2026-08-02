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

st.title("📊 MRP Control Tower & Live Inventory Engine")
st.markdown("ניהול חוסרים דינמי: עדכון מלאי נכנס ו-ETA שמשפיעים ישירות על חישובי ה-MRP ומסד הנתונים")

# ==========================================================
# LOCAL DATABASE SETUP (Persistent Storage)
# ==========================================================
conn = sqlite3.connect(LOCAL_DB_FILE, check_same_thread=False)

conn.execute("""
CREATE TABLE IF NOT EXISTS inventory_updates
(
    pn TEXT PRIMARY KEY,
    added_stock REAL,
    eta TEXT,
    status TEXT,
    supplier TEXT,
    comment TEXT,
    updated_by TEXT,
    updated_at TEXT
)
""")

conn.execute("""
CREATE TABLE IF NOT EXISTS inventory_history
(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pn TEXT,
    added_stock REAL,
    eta TEXT,
    status TEXT,
    supplier TEXT,
    comment TEXT,
    updated_by TEXT,
    updated_at TEXT
)
""")
conn.commit()

def get_inventory_record(pn):
    cur = conn.cursor()
    try:
        cur.execute("SELECT added_stock, eta, status, supplier, comment, updated_by, updated_at FROM inventory_updates WHERE pn = ?", (pn,))
        res = cur.fetchone()
        if res:
            return res[0], res[1], res[2], res[3], res[4], res[5], res[6]
    except:
        pass
    return 0.0, str(date.today()), "פתוח", "אופק", "", "", ""

def save_inventory_record(pn, added_stock, eta, status, supplier, comment, updated_by, webhook_url=""):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn.execute("""
    INSERT OR REPLACE INTO inventory_updates (pn, added_stock, eta, status, supplier, comment, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (pn, added_stock, eta, status, supplier, comment, updated_by, now_str))
    
    conn.execute("""
    INSERT INTO inventory_history (pn, added_stock, eta, status, supplier, comment, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (pn, added_stock, eta, status, supplier, comment, updated_by, now_str))
    
    conn.commit()
    
    if webhook_url:
        msg = "🔔 עדכון מלאי/ETA למוצר!\nמק\"ט: " + str(pn) + "\nתוספת מלאי: " + str(added_stock) + "\nסטטוס: " + str(status) + "\nETA: " + str(eta)
        try:
            requests.post(webhook_url, data=json.dumps({"text": msg}), headers={'Content-Type': 'application/json'})
        except:
            pass

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
    st.error(f"שגיאה בטעינת הקובץ מ-GitHub. פירוט השגיאה: {e}")
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
# APPLY USER INVENTORY UPDATES TO MAIN DATAFRAME
# ==========================================================
# נטען את כל העדכונים שנשמרו במסד הנתונים ונוסיף אותם למלאי הבסיסי
for idx, row in df.iterrows():
    pn = str(row[PN_COL]).strip()
    saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn)
    if saved_stock_add > 0:
        base_stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
        df.at[idx, STOCK_COL] = base_stock + saved_stock_add

# ==========================================================
# SIDEBAR FILTERS
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
            assembly_mapping[col] = str(col) + " - " + str(desc) + " (רמה " + str(lvl) + ")"
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
# CORE LOGIC (MRP Calculation with Updated Inventory)
# ==========================================================
# חשוב מחדש את מאזן החומרים החודשי לפי המלאי החדש שנוסף
df['Monthly_Balance'] = pd.to_numeric(df[selected_month_col], errors='coerce').fillna(0)
# אם המלאי עודכן כלפי מעלה, נעדכן את ה-Balance באופן יחסי כך שהחוסר יתעדכן אוטומטית
for idx, row in df.iterrows():
    pn = str(row[PN_COL]).strip()
    added_stock, _, _, _, _, _, _ = get_inventory_record(pn)
    if added_stock > 0:
        current_bal = df.at[idx, 'Monthly_Balance']
        # אם יש חוסר שלילי, הוספת מלאי תקטין את החוסר (תקרב ל-0)
        if current_bal < 0:
            df.at[idx, 'Monthly_Balance'] = current_bal + added_stock

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
    
    _, _, _, current_sup, _, _, _ = get_inventory_record(pn)

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
# TABS
# ==========================================================
tab1, tab2, tab3 = st.tabs(["📦 דשבורד חוסרים ראשי", "⚠️ ניתוח צווארי בקבוק (Bottlenecks)", "📅 עדכון מלאי, ETA וספקים"])

with tab1:
    st.subheader("📈 ניתוח חוסרים מעודכן לחודש: " + str(selected_month_label))

    col1, col2, col3 = st.columns(3)
    col1.metric("🔴 פריטים בחוסר ב-MRP", len(mrp_shortages))
    col2.metric("📋 סך שורות פריט-הדדי להרכבה", len(breakdown_df) if not breakdown_df.empty else 0)
    total_req_demand = breakdown_df['Required_Demand'].sum() if not breakdown_df.empty else 0
    col3.metric("📦 סך ביקוש מחושב בהרכבות", f"{total_req_demand:,.0f}")

    st.divider()

    if not breakdown_df.empty and len(breakdown_df) > 0:
        st.subheader("📋 פירוט חוסרים מלא ופילוח מול הרכבות (לוקח בחשבון מלאי חדש שנכנס)")
        
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
            "Stock": "מלאי נוכחי (כולל תוספות)",
            "Total_MRP_Shortage": "סך חוסר מעודכן ב-MRP"
        })
        
        st.dataframe(display_df.sort_values(by="סך חוסר מעודכן ב-MRP", ascending=False), use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='Live_Shortages')
        processed_data = output.getvalue()

        st.download_button(
            label="📥 הורד את הטבלה המעודכנת לקובץ Excel",
            data=processed_data,
            file_name=f"Live_MRP_Shortages_{selected_ym}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("🎉 אין חוסרים ב-MRP לחודש זה (או שהם התאפסו בעקבות מלאי חדש שנכנס)!")

with tab2:
    st.subheader("⚠️ ניתוח צווארי בקבוק (Bottleneck Analysis)")
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
        if count_assemblies > 1:
            bottleneck_rows.append({
                "מק\"ט": pn,
                "תיאור": desc,
                "מספר הרכבות שבהן משתתף": count_assemblies,
                "סך כמות נדרשת במצטבר": total_qty_needed
            })
    if bottleneck_rows:
        st.dataframe(pd.DataFrame(bottleneck_rows).sort_values(by="מספר הרכבות שבהן משתתף", ascending=False).head(20), use_container_width=True)
    else:
        st.info("לא נמצאו פריטים משותפים למספר הרכבות.")

with tab3:
    st.subheader("📅 עדכון מלאי נכנס, ETA וספקים (מתעדכן מידית ב-MRP ובמסד הנתונים)")

    pn_values = sorted(df[PN_COL].dropna().astype(str).unique())
    selected_pn = st.selectbox("בחר מק\"ט לעדכון", pn_values)

    # שליפת נתונים קיימים
    saved_stock, saved_eta, saved_status, saved_supplier, saved_comment, saved_by, _ = get_inventory_record(selected_pn)

    with st.form("inventory_form"):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            added_stock_input = st.number_input("כמות מלאי חדשה שהגיעה (להוספה למלאי)", min_value=0.0, value=float(saved_stock), step=1.0)
        with col_f2:
            try:
                parsed_eta = pd.to_datetime(saved_eta).date() if saved_eta else date.today()
            except:
                parsed_eta = date.today()
            eta_date = st.date_input("תאריך הגעה משוער (ETA)", value=parsed_eta)
        with col_f3:
            status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
            status_idx = status_options.index(saved_status) if saved_status in status_options else 0
            status = st.selectbox("סטטוס", status_options, index=status_idx)

        col_f4, col_f5 = st.columns(2)
        with col_f4:
            sup_idx = supplier_options.index(saved_supplier) if saved_supplier in supplier_options else 0
            supplier = st.selectbox("ספק / קבלן משנה", supplier_options, index=sup_idx)
        with col_f5:
            updated_by = st.text_input("עודכן על ידי", value=saved_by)

        comment = st.text_area("הערות מעקב", value=saved_comment)
        save_btn = st.form_submit_button("שמור עדכון והחל על ה-MRP")

    if save_btn:
        save_inventory_record(selected_pn, added_stock_input, str(eta_date), status, supplier, comment, updated_by, webhook_url)
        st.success("הנתונים נשמרו בהצלחה במסד הנתונים והשפיעו מידית על תחשיבי ה-MRP!")
        st.rerun()

    # אפשרות גיבוי והורדת בסיס הנתונים
    st.divider()
    with open(LOCAL_DB_FILE, "rb") as db_file:
        db_bytes = db_file.read()
    st.download_button(
        label="📥 הורד גיבוי מלא של מסד הנתונים המקומי (.db)",
        data=db_bytes,
        file_name="inventory_backup.db",
        mime="application/octet-stream"
    )

    # טבלת מעקב והיסטוריה
    st.subheader("🚦 טבלת כל הפריטים שעודכנו במערכת")
    history_cur = conn.cursor()
    history_cur.execute("SELECT pn, added_stock, eta, status, supplier, comment, updated_by, updated_at FROM inventory_updates ORDER BY updated_at DESC")
    all_updated_rows = history_cur.fetchall()
    if all_updated_rows:
        up_df = pd.DataFrame(all_updated_rows, columns=["מק\"ט", "מלאי נוסף", "ETA", "סטטוס", "ספק", "הערות", "עודכן ע\"י", "זמן עדכון"])
        st.dataframe(up_df, use_container_width=True)
    else:
        st.info("עדיין לא עודכנו פריטים במערכת.")
