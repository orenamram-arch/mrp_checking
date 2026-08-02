import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import datetime, date
import io
import requests
import json
import random

# ==========================================================
# CONFIGURATION & LOGIN SYSTEM
# ==========================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"
LOCAL_DB_FILE = "eta_updates.db" 

st.set_page_config(
    page_title="MRP Control Tower & Visual Dashboard",
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

st.title("📊 MRP Control Tower & Visual Analytics Dashboard")
st.markdown("דשבורד ויזואלי מתקדם לניהול חוסרים, ניתוח תוכניות ייצור וסימולציית Clear To Build ממוקדת")

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
    return 0.0, "", "פתוח", "אופק", "", "", ""

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

asm_components = {}
for col in valid_assemblies:
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

item_types = df[ITEM_TYPE_COL].dropna().unique().tolist()
selected_item_type = st.sidebar.selectbox("בחר סוג פריט (עמודה AS)", ["הכל"] + item_types)

item_choices = ["הכל"] + sorted([f"{str(r[PN_COL]).strip()} - {str(r[DESC_COL])}" for _, r in df.iterrows() if pd.notnull(r[PN_COL])])
selected_search_item = st.sidebar.selectbox("🔎 חיפוש מהיר (בחר או הקלד מק\"ט/תיאור)", item_choices)
search_pn = selected_search_item.split(" - ")[0] if selected_search_item != "הכל" else "הכל"

# ==========================================================
# CORE LOGIC FOR SHORTAGES
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
                "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                "Assembly": asm, "Assembly_Desc": asm_desc, "Qty_Per_Assembly": qty_per_asm,
                "Assembly_Monthly_Build": asm_build_qty, "Required_Demand": required_demand,
                "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage
            })
            
    if not matched_any and selected_assembly == "הכל" and selected_level == "הכל":
        breakdown_rows.append({
            "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
            "Assembly": "ללא שיוך", "Assembly_Desc": "ללא שיוך להרכבה", "Qty_Per_Assembly": 0,
            "Assembly_Monthly_Build": 0, "Required_Demand": 0, "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage
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
# TABS
# ==========================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📦 דשבורד חוסרים ויזואלי", 
    "📊 תוכנית ייצור חודשית (Smart CTB)", 
    "⚠️ צווארי בקבוק", 
    "📅 מעקב מלאי וספקים",
    "↩️ ניהול UNDO"
])

with tab1:
    st.subheader(f"📈 דשבורד חוסרים לחודש: {selected_month_label}")
    
    dash_df = breakdown_df.copy()
    if selected_assembly != "הכל":
        dash_df = dash_df[dash_df["Assembly"] == selected_assembly]
        st.info(f"🎯 מציג חוסרים ממוקדים עבור הרכבה: {assembly_mapping.get(selected_assembly, selected_assembly)}")

    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    total_shortage_items = len(dash_df['PN'].unique()) if not dash_df.empty else 0
    total_shortage_qty = dash_df['Total_MRP_Shortage'].sum() if not dash_df.empty else 0
    
    col_k1.metric("🔴 סה\"כ מק\"טים בחוסר", total_shortage_items)
    col_k2.metric("📦 סה\"כ כמות חסרה מצטברת", f"{total_shortage_qty:,.0f}")
    col_k3.metric("📅 חודש מנותח", selected_month_label.split('(')[0])
    col_k4.metric("⚙️ הרכבה נבחרת", "הכל" if selected_assembly == "הכל" else selected_assembly)

    st.divider()

    if not dash_df.empty and len(dash_df) > 0:
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("### 🥧 התפלגות חוסרים לפי סוג פריט")
            fig_pie = px.pie(dash_df, names="Item_Type", values="Total_MRP_Shortage", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_g2:
            st.markdown("### 📊 10 הפריטים הגירעוניים ביותר")
            top_shortages = dash_df.sort_values(by="Total_MRP_Shortage", ascending=False).head(10)
            fig_bar = px.bar(top_shortages, x="PN", y="Total_MRP_Shortage", text_auto='.2s', color="Total_MRP_Shortage", color_continuous_scale="Reds")
            st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader("📋 טבלת פירוט מלאה")
        display_df = dash_df[[
            "PN", "Description", "Item_Type", "Supplier", "Assembly", "Assembly_Desc", 
            "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage"
        ]].rename(columns={
            "PN": "מק\"ט", "Description": "תיאור פריט", "Item_Type": "סוג פריט (AS)",
            "Supplier": "ספק / קב\"מ", "Assembly": "קוד הרכבה", "Assembly_Desc": "תיאור הרכבה",
            "Qty_Per_Assembly": "כמות נדרשת להרכבה", "Assembly_Monthly_Build": "ת. ייצור הרכבה לחודש",
            "Required_Demand": "ביקוש מדויק להרכבה", "Stock": "מלאי נוכחי", "Total_MRP_Shortage": "סך חוסר ב-MRP"
        })
        st.dataframe(display_df.sort_values(by="סך חוסר ב-MRP", ascending=False), use_container_width=True)
    else:
        st.success("🎉 אין חוסרים ב-MRP עבור ההגדרות והסינונים שנבחרו!")

with tab2:
    st.subheader(f"📊 סימולציית Clear To Build (CTB) לחודש: {selected_month_label}")
    st.markdown("המערכת מקצה את המלאי הזמין לפי עדיפות לרמות עץ מוצר נמוכות יותר, ומציגה **אך ורק את הרכיבים שחסרים בפועל**, כאשר הפריט הקריטי מודגש ב-**BOLD**.")

    assemblies_to_check = [asm for asm in valid_assemblies if assembly_plan_df[(assembly_plan_df["YearMonth"] == selected_ym) & (assembly_plan_df["Assembly_PN"] == asm)]["Build_Qty"].sum() > 0]
    assemblies_to_check.sort(key=lambda x: assembly_levels.get(x, 0), reverse=True)

    available_inventory_pool = {}
    for idx_r, row_r in df.iterrows():
        p_num = str(row_r[PN_COL]).strip()
        base_st = pd.to_numeric(row_r[STOCK_COL], errors='coerce') or 0
        added_st, _, _, _, _, _, _ = get_inventory_record(p_num)
        available_inventory_pool[p_num] = max(0, base_st) + added_st

    production_capacity_rows = []

    for asm_col in assemblies_to_check:
        if selected_assembly != "הכל" and asm_col != selected_assembly:
            continue
            
        try:
            asm_desc = df_desc.iloc[0, df.columns.get_loc(asm_col)]
        except:
            asm_desc = ""
            
        planned_build = assembly_plan_df[(assembly_plan_df["YearMonth"] == selected_ym) & (assembly_plan_df["Assembly_PN"] == asm_col)]["Build_Qty"].sum()
        
        comps = asm_components.get(asm_col, pd.DataFrame())
        max_buildable = planned_build
        missing_items_details = []

        if len(comps) > 0:
            for _, comp_row in comps.iterrows():
                qty_per = float(comp_row[asm_col])
                c_pn = str(comp_row[PN_COL]).strip()
                c_desc = str(comp_row[DESC_COL]).strip()
                
                if qty_per > 0:
                    required_for_plan = planned_build * qty_per
                    current_pool = available_inventory_pool.get(c_pn, 0)
                    
                    _, eta_val, _, _, _, _, _ = get_inventory_record(c_pn)
                    try:
                        eta_dt = pd.to_datetime(eta_val).date() if eta_val else date(2099, 12, 31)
                    except:
                        eta_dt = date(2099, 12, 31)

                    if current_pool >= required_for_plan:
                        available_inventory_pool[c_pn] = current_pool - required_for_plan
                        possible_units = planned_build
                    else:
                        possible_units = int(current_pool / qty_per) if qty_per > 0 else 0
                        missing_qty = required_for_plan - current_pool
                        available_inventory_pool[c_pn] = 0
                        # כאן אנו מוסיפים אך ורק את הרכיבים שבאמת חסרים
                        missing_items_details.append((c_pn, c_desc, missing_qty, eta_dt, eta_val))

                    if possible_units < max_buildable:
                        max_buildable = possible_units

        # מיון החסרים כדי לזהות את הקריטי ביותר
        if missing_items_details:
            missing_items_details.sort(key=lambda x: x[3], reverse=True)
            most_critical_pn = missing_items_details[0][0]
        else:
            most_critical_pn = None

        formatted_missing = []
        for c_pn, c_desc, m_qty, _, raw_eta in missing_items_details:
            eta_str = f" [ETA: {raw_eta}]" if raw_eta else " [ללא ETA]"
            item_text = f"{c_pn} ({c_desc[:12]}) - חסר: {m_qty:g}{eta_str}"
            
            if c_pn == most_critical_pn:
                formatted_missing.append(f"**{item_text}**")
            else:
                formatted_missing.append(item_text)

        missing_str = " | ".join(formatted_missing) if formatted_missing else "אין חוסרים! ניתן לייצר את כל התוכנית."

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
        st.dataframe(cap_df, use_container_width=True)
    else:
        st.info(f"לא נמצאו הרכבות מתוכננות לייצור לחודש {selected_month_label}.")

with tab3:
    st.subheader("⚠️ ניתוח צווארי בקבוק רוחביים")
    bottleneck_rows = []
    for idx, row in df.iterrows():
        pn = str(row[PN_COL]).strip()
        desc = str(row[DESC_COL])
        count_assemblies = sum(1 for asm in valid_assemblies if (pd.to_numeric(row[asm], errors='coerce') or 0) > 0)
        
        if count_assemblies > 1:
            bottleneck_rows.append({
                "מק\"ט": pn, "תיאור": desc, "מספר הרכבות שבהן משתתף": count_assemblies
            })
    if bottleneck_rows:
        st.dataframe(pd.DataFrame(bottleneck_rows).sort_values(by="מספר הרכבות שבהן משתתף", ascending=False).head(20), use_container_width=True)

with tab4:
    st.subheader("📅 עדכון מלאי, ETA וקבלני משנה")
    selected_pn = search_pn if search_pn != "הכל" else st.selectbox("בחר מק\"ט מכלל הפריטים לעדכון", sorted(df[PN_COL].dropna().astype(str).unique()))
    
    if selected_pn != "הכל":
        saved_stock, saved_eta, saved_status, saved_supplier, saved_comment, saved_by, _ = get_inventory_record(selected_pn)
        with st.form("inventory_form"):
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                added_stock_input = st.number_input("תוספת למלאי זמין", min_value=0.0, value=float(saved_stock), step=1.0)
            with col_f2:
                try: parsed_eta = pd.to_datetime(saved_eta).date() if saved_eta else date.today()
                except: parsed_eta = date.today()
                eta_date = st.date_input("תאריך הגעה (ETA)", value=parsed_eta)
            with col_f3:
                status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
                status_idx = status_options.index(saved_status) if saved_status in status_options else 0
                status = st.selectbox("סטטוס", status_options, index=status_idx)

            col_f4, col_f5 = st.columns(2)
            with col_f4:
                sup_idx = supplier_options.index(saved_supplier) if saved_supplier in supplier_options else 0
                supplier = st.selectbox("ספק", supplier_options, index=sup_idx)
            with col_f5:
                updated_by = st.text_input("עודכן ע\"י", value=saved_by)
            comment = st.text_area("הערות", value=saved_comment)
            if st.form_submit_button("שמור עדכון"):
                save_inventory_record(selected_pn, added_stock_input, str(eta_date), status, supplier, comment, updated_by, webhook_url)
                st.success("נשמר בהצלחה!")
                st.rerun()

with tab5:
    st.subheader("↩️ חזרה לאחור וניהול היסטוריה (UNDO)")
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
                    st.text(f"תוספת: {i_stock} | ETA: {i_eta}")
                    st.text(f"עודכן ע\"י: {i_by} ({i_time})")
                with col_u3:
                    if st.button("🔄 UNDO", key=f"undo_{i_pn}"):
                        delete_inventory_record(i_pn)
                        st.success("בוטל בהצלחה!")
                        st.rerun()
                st.divider()
    else:
        st.info("אין עדכונים במערכת.")
