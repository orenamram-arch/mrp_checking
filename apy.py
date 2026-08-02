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
    page_title="MRP Control Tower & Production Plan",
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

st.title("📊 MRP Control Tower & Master Production Schedule")
st.markdown("מנוע חישוב חכם לזיהוי צווארי בקבוק וחישוב יכולת ייצור רקורסיבית (Multi-Level BOM)")

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

def delete_inventory_record(pn):
    conn.execute("DELETE FROM inventory_updates WHERE pn = ?", (pn,))
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

PN_COL = df.columns[1]     
DESC_COL = df.columns[4]   
ITEM_TYPE_COL = df.columns[44] 
STOCK_COL = df.columns[79]     
ASSEMBLY_COLS = df.columns[10:36].tolist()
MONTH_COLS = df.columns[108:132].tolist()

# הכנת מילון רכיבים מראש לביצועים מהירים בסריקה הרקורסיבית
asm_components = {}
for col in ASSEMBLY_COLS:
    asm_components[col] = df[pd.to_numeric(df[col], errors='coerce') > 0]

# ==========================================================
# APPLY USER INVENTORY UPDATES TO MAIN DATAFRAME
# ==========================================================
for idx, row in df.iterrows():
    pn = str(row[PN_COL]).strip()
    saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn)
    if saved_stock_add > 0:
        base_stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
        df.at[idx, STOCK_COL] = base_stock + saved_stock_add

# ==========================================================
# SIDEBAR FILTERS & SMART AUTOCOMPLETE SEARCH
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

item_choices = ["הכל"] + sorted([f"{str(r[PN_COL]).strip()} - {str(r[DESC_COL])}" for _, r in df.iterrows() if pd.notnull(r[PN_COL])])
selected_search_item = st.sidebar.selectbox("🔎 חיפוש מהיר (בחר או הקלד מק\"ט/תיאור)", item_choices)
search_pn = selected_search_item.split(" - ")[0] if selected_search_item != "הכל" else "הכל"

# ==========================================================
# CORE LOGIC
# ==========================================================
df['Monthly_Balance'] = pd.to_numeric(df[selected_month_col], errors='coerce').fillna(0)
for idx, row in df.iterrows():
    pn = str(row[PN_COL]).strip()
    added_stock, _, _, _, _, _, _ = get_inventory_record(pn)
    if added_stock > 0:
        current_bal = df.at[idx, 'Monthly_Balance']
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
    if search_pn != "הכל":
        breakdown_df = breakdown_df[breakdown_df["PN"] == search_pn]

# ==========================================================
# RECURSIVE INVENTORY CHECK ENGINE
# ==========================================================
memo_avail = {}
def get_actual_availability(pn, visited=None):
    """
    אלגוריתם רקורסיבי: מחשב את הזמינות האמיתית של פריט.
    אם הפריט הוא הרכבה - הזמינות היא המלאי הפיזי שלו + כמות היחידות שאפשר לבנות מהרכיבים שלו (ואלו נבדקים גם הם רקורסיבית).
    """
    if visited is None: visited = set()
    if pn in memo_avail: return memo_avail[pn]
    if pn in visited: return 0 # מניעת לולאה אינסופית בעץ מוצר מעגלי
    
    visited.add(pn)
    row_idx = df[df[PN_COL] == pn].index
    if len(row_idx) == 0:
        visited.remove(pn)
        return 0
        
    row = df.loc[row_idx[0]]
    added_st, _, _, _, _, _, _ = get_inventory_record(pn)
    phys_stock = max(0, pd.to_numeric(row[STOCK_COL], errors='coerce') or 0) + added_st
    
    # אם זה פריט קנוי (לא מופיע בראש עמודת הרכבה) -> הזמינות היא פשוט המלאי שלו
    if pn not in ASSEMBLY_COLS:
        memo_avail[pn] = phys_stock
        visited.remove(pn)
        return phys_stock
        
    # אם זה תת-הרכבה, נבדוק את הרכיבים שלה (הבנים והנכדים)
    comps = asm_components.get(pn, pd.DataFrame())
    if len(comps) == 0:
        memo_avail[pn] = phys_stock
        visited.remove(pn)
        return phys_stock
        
    min_build = float('inf')
    for _, c_row in comps.iterrows():
        c_pn = str(c_row[PN_COL]).strip()
        q_per = float(c_row[pn])
        if q_per > 0:
            # קריאה רקורסיבית למטה בעץ המוצר
            c_avail = get_actual_availability(c_pn, visited)
            possible = c_avail / q_per
            if possible < min_build:
                min_build = possible
                
    total_avail = phys_stock + (min_build if min_build != float('inf') else 0)
    memo_avail[pn] = total_avail
    visited.remove(pn)
    return total_avail

# ==========================================================
# TABS
# ==========================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📦 דשבורד חוסרים", 
    "📊 תוכנית ייצור חודשית (מבוסס עץ מוצר מלא)", 
    "⚠️ צווארי בקבוק", 
    "📅 עדכון מלאי, ETA וספקים",
    "↩️ ניהול שינויים ו-UNDO"
])

with tab1:
    st.subheader("📈 ניתוח חוסרים מעודכן לחודש: " + str(selected_month_label))

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
            "PN": "מק\"ט", "Description": "תיאור פריט", "Item_Type": "סוג פריט (AS)",
            "Supplier": "ספק / קב\"מ", "Assembly": "קוד הרכבה", "Assembly_Desc": "תיאור הרכבה",
            "Qty_Per_Assembly": "כמות נדרשת להרכבה", "Assembly_Monthly_Build": "ת. ייצור הרכבה לחודש",
            "Required_Demand": "ביקוש מדויק להרכבה", "Stock": "מלאי נוכחי", "Total_MRP_Shortage": "סך חוסר מעודכן ב-MRP"
        })
        st.dataframe(display_df.sort_values(by="סך חוסר מעודכן ב-MRP", ascending=False), use_container_width=True)
    else:
        st.success("🎉 אין חוסרים ב-MRP לחודש זה עבור הסינון שנבחר!")

with tab2:
    st.subheader("📊 תוכנית ייצור מבוססת עץ מוצר רקורסיבי (Multi-Level BOM)")
    st.markdown("חישוב זה מזהה תתי-הרכבות ובודק את זמינות הרכיבים בתוך הרכיבים כדי למצוא את המגבלה האמיתית:")
    production_capacity_rows = []

    for m_label, m_col in month_options.items():
        dt_m = pd.to_datetime(m_col)
        ym_str = dt_m.strftime("%Y-%m")
        short_month_name = str(m_label).split('(')[0].strip()

        m_plan_subset = assembly_plan_df[assembly_plan_df["YearMonth"] == ym_str]
        
        for idx_asm, asm_col in enumerate(ASSEMBLY_COLS):
            try:
                asm_desc = df_desc.iloc[0, df.columns.get_loc(asm_col)]
            except:
                asm_desc = ""

            planned_build = m_plan_subset[m_plan_subset["Assembly_PN"] == asm_col]["Build_Qty"].sum()
            if planned_build == 0:
                continue

            # חיפוש הגורם המגביל ביותר בהתבסס על הזמינות הרקורסיבית
            max_buildable = planned_build
            limiting_item = "ללא מגבלה"
            
            comps = asm_components.get(asm_col, pd.DataFrame())
            for _, comp_row in comps.iterrows():
                qty_per = float(comp_row[asm_col])
                c_pn = str(comp_row[PN_COL]).strip()
                
                if qty_per > 0:
                    c_avail = get_actual_availability(c_pn)
                    possible = c_avail / qty_per
                    if possible < max_buildable:
                        max_buildable = int(possible)
                        limiting_item = f"{c_pn} ({str(comp_row[DESC_COL]).strip()[:20]})"

            production_capacity_rows.append({
                "חודש": short_month_name,
                "קוד הרכבה": asm_col,
                "תיאור הרכבה": asm_desc,
                "תוכנית ייצור מתוכננת": planned_build,
                "יכולת ייצור בפועל": max_buildable,
                "גורם מגביל עיקרי (צוואר בקבוק שורשי)": limiting_item
            })

    if production_capacity_rows:
        cap_df = pd.DataFrame(production_capacity_rows)
        selected_cap_asm = st.selectbox("בחר הרכבה לניתוח יכולת ייצור חודשית", ["הכל"] + ASSEMBLY_COLS)
        if selected_cap_asm != "הכל":
            display_cap_df = cap_df[cap_df["קוד הרכבה"] == selected_cap_asm]
        else:
            display_cap_df = cap_df

        st.dataframe(display_cap_df, use_container_width=True)

        fig_cap = px.bar(
            display_cap_df, 
            x="חודש", 
            y=["תוכנית ייצור מתוכננת", "יכולת ייצור בפועל"], 
            barmode="group",
            title="השוואה בין תוכנית עבודה מקורית לבין יכולת ייצור אמיתית"
        )
        st.plotly_chart(fig_cap, use_container_width=True)
    else:
        st.info("לא נמצאו נתוני תוכנית ייצור להרכבות.")

with tab3:
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
                "מק\"ט": pn, "תיאור": desc,
                "מספר הרכבות שבהן משתתף": count_assemblies,
                "סך כמות נדרשת במצטבר": total_qty_needed
            })
    if bottleneck_rows:
        st.dataframe(pd.DataFrame(bottleneck_rows).sort_values(by="מספר הרכבות שבהן משתתף", ascending=False).head(20), use_container_width=True)
    else:
        st.info("לא נמצאו פריטים משותפים למספר הרכבות.")

with tab4:
    st.subheader("📅 עדכון מלאי נכנס, ETA וספקים")

    selected_pn = st.selectbox("בחר מק\"ט לעדכון", item_choices).split(" - ")[0]
    if selected_pn == "הכל":
        st.warning("אנא בחר מק\"ט ספציפי מהרשימה.")
    else:
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

        st.divider()
        with open(LOCAL_DB_FILE, "rb") as db_file:
            db_bytes = db_file.read()
        st.download_button(
            label="📥 הורד גיבוי מלא של מסד הנתונים המקומי (.db)",
            data=db_bytes,
            file_name="inventory_backup.db",
            mime="application/octet-stream"
        )

with tab5:
    st.subheader("↩️ ניהול פריטים שעודכנו ואפשרות ביצוע UNDO")
    st.markdown("כאן מוצגים כל הפריטים ששינית להם את המלאי או ה-ETA. לחיצה על כפתור ה-UNDO תמחק את העדכון ותחזיר את הפריט למצבו המקורי במערכת וב-MRP.")

    history_cur = conn.cursor()
    history_cur.execute("SELECT pn, added_stock, eta, status, supplier, comment, updated_by, updated_at FROM inventory_updates ORDER BY updated_at DESC")
    updated_items = history_cur.fetchall()

    if updated_items:
        for item in updated_items:
            i_pn, i_stock, i_eta, i_status, i_sup, i_comm, i_by, i_time = item
            
            with st.container():
                col_u1, col_u2, col_u3 = st.columns([3, 4, 1])
                with col_u1:
                    st.markdown(f"**מק\"ט:** `{i_pn}`")
                    st.text(f"ספק: {i_sup} | סטטוס: {i_status}")
                with col_u2:
                    st.text(f"תוספת מלאי: {i_stock} | ETA: {i_eta}")
                    st.text(f"עודכן ע\"י: {i_by} בתאריך: {i_time}")
                with col_u3:
                    if st.button("🔄 UNDO", key=f"undo_{i_pn}", help="החזר למצב מקורי"):
                        delete_inventory_record(i_pn)
                        st.success(f"העדכון למק\"ט {i_pn} בוטל בהצלחה! ה-MRP הוחזר לקדמותו.")
                        st.rerun()
                st.divider()
    else:
        st.info("אין כרגע פריטים שעודכנו במערכת.")
