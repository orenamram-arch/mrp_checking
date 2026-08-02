"""
MRP Control Tower — מגדל בקרת חוסרים
גרסה משלבת: לוגיקת ה-MRP וה-ETA המקורית והמדויקת יחד עם התוספות הניהוליות (Dashboard, What-If, Kanban).

הרצה:
streamlit run mrp_app.py
"""

import json
import os
import sqlite3
from datetime import datetime, date
import io

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ==========================================================================
# 1. תצורה וכתובות
# ==========================================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"
LOCAL_DB_FILE = "eta_updates.db" 

st.set_page_config(
    page_title="MRP Control Tower & Decision Hub",
    page_icon="📦",
    layout="wide"
)

# ==========================================================================
# 2. שכבת שמירה ו-SQLite
# ==========================================================================
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
            eta_val = res[1] if res[1] and str(res[1]).strip() not in ["", "None", "NaT", "nan"] else ""
            status_val = res[2] if res[2] else "פתוח"
            return res[0], eta_val, status_val, res[3], res[4], res[5], res[6]
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
        msg = f"🔔 עדכון מלאי/ETA למוצר!\nמק\"ט: {pn}\nתוספת מלאי: {added_stock}\nסטטוס: {status}\nETA: {eta}"
        try:
            requests.post(webhook_url, data=json.dumps({"text": msg}), headers={'Content-Type': 'application/json'})
        except:
            pass

def delete_inventory_record(pn):
    conn.execute("DELETE FROM inventory_updates WHERE pn = ?", (pn,))
    conn.commit()

# ==========================================================================
# 3. טעינת נתונים מקורית מ-GitHub
# ==========================================================================
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

# ==========================================================================
# 4. חילוץ תוכנית ייצור והרכבות (לוגיקה מקורית)
# ==========================================================================
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

def get_first_supply_eta(pn):
    _, manual_eta, _, _, _, _, _ = get_inventory_record(pn)
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        return manual_eta
        
    matching_rows = df_raw[df_raw.iloc[:, 1].astype(str).str.strip() == str(pn).strip()]
    if not matching_rows.empty:
        row_idx = matching_rows.index[0]
        max_cols = df_raw.shape[1]
        for col_pos in range(80, min(108, max_cols)):
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
    return "ללא ETA"

# החלת עדכוני מלאי מהמסד המקומי
for idx, row in df.iterrows():
    pn = str(row[PN_COL]).strip()
    saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn)
    if saved_stock_add > 0:
        base_stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
        df.at[idx, STOCK_COL] = base_stock + saved_stock_add

# ==========================================================================
# 5. סיידבר ומסננים
# ==========================================================================
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
selected_item_type = st.sidebar.selectbox("בחר סוג פריט", ["הכל"] + item_types)

item_choices = ["הכל"] + sorted([f"{str(r[PN_COL]).strip()} - {str(r[DESC_COL])}" for _, r in df.iterrows() if pd.notnull(r[PN_COL])])
selected_search_item = st.sidebar.selectbox("🔎 חיפוש מהיר מק\"ט/תיאור", item_choices)
search_pn = selected_search_item.split(" - ")[0] if selected_search_item != "הכל" else "הכל"

# ==========================================================================
# 6. לוגיקת החוסרים המקורית (Core MRP Logic)
# ==========================================================================
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
    item_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in df.columns else ""
    stock = pd.to_numeric(row[STOCK_COL], errors='coerce') or 0
    total_mrp_shortage = row['Total_MRP_Shortage']
    
    _, _, item_status, current_sup, _, _, _ = get_inventory_record(pn)
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

breakdown_df = pd.DataFrame(breakdown_rows)

if not breakdown_df.empty:
    if selected_item_type != "הכל":
        breakdown_df = breakdown_df[breakdown_df["Item_Type"] == selected_item_type]
    if selected_assembly != "הכל":
        breakdown_df = breakdown_df[breakdown_df["Assembly"] == selected_assembly]
    if search_pn != "הכל":
        breakdown_df = breakdown_df[breakdown_df["PN"] == search_pn]

# ==========================================================================
# 7. ממשק משתמש ולשוניות (Executive Control Tower & Tabs)
# ==========================================================================
st.title("🚀 MRP Executive Control Tower & Decision Hub")
st.markdown("מערכת ניהול חוסרים מתקדמת, סימולציות קבלת החלטות (What-If), ותמונת מצב ניהלית")

TABS = st.tabs([
    "📈 Executive Dashboard", 
    "📊 תוכנית ייצור (Smart CTB)", 
    "💡 סימולציית What-If",
    "📌 לוח סטטוסים (Kanban)",
    "📅 עדכון מלאי וספקים",
    "↩️ ניהול UNDO"
])

with TABS[0]:
    st.subheader(f"🎯 תמונת מצב ניהלית (Executive Summary) לחודש: {selected_month_label}")
    
    dash_df = breakdown_df.copy()
    if selected_assembly != "הכל":
        dash_df = dash_df[dash_df["Assembly"] == selected_assembly]

    total_planned_assemblies = len([a for a in valid_assemblies if assembly_plan_df[(assembly_plan_df["YearMonth"] == selected_ym) & (assembly_plan_df["Assembly_PN"] == a)]["Build_Qty"].sum() > 0])
    blocked_assemblies = len(dash_df['Assembly'].unique()) if not dash_df.empty else 0
    ready_assemblies = max(0, total_planned_assemblies - blocked_assemblies)
    readiness_pct = (ready_assemblies / total_planned_assemblies * 100) if total_planned_assemblies > 0 else 100

    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    col_k1.metric("🟢 מוכנות קווי ייצור", f"{readiness_pct:.1f}%", f"{ready_assemblies}/{total_planned_assemblies} הרכבות מוכנות")
    col_k2.metric("🔴 הרכבות חסומות בחודש", blocked_assemblies)
    col_k3.metric("📦 סה\"כ מק\"טים בגירעון", len(dash_df['PN'].unique()) if not dash_df.empty else 0)
    col_k4.metric("📊 כמות גירעון מצטברת", f"{dash_df['Total_MRP_Shortage'].sum():,.0f}" if not dash_df.empty else "0")

    st.divider()

    if not dash_df.empty and len(dash_df) > 0:
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("### 🥧 התפלגות חוסרים לפי סוג פריט")
            fig_pie = px.pie(dash_df, names="Item_Type", values="Total_MRP_Shortage", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_g2:
            st.markdown("### 📊 מגמת חוסרים רוחבית לאורך חודשי השנה")
            trend_rows = []
            for m_col in MONTH_COLS:
                if pd.notnull(m_col):
                    try:
                        m_dt = pd.to_datetime(m_col)
                        m_ym = m_dt.strftime("%Y-%m")
                        temp_b = pd.to_numeric(df[m_col], errors='coerce').fillna(0)
                        tot_sh = temp_b[temp_b < 0].abs().sum()
                        trend_rows.append({"Month": m_ym, "Total_Shortage": tot_sh})
                    except:
                        pass
            trend_df = pd.DataFrame(trend_rows)
            if not trend_df.empty:
                fig_line = px.line(trend_df, x="Month", y="Total_Shortage", markers=True, title="היקף החוסרים הצפוי לפי חודשים")
                st.plotly_chart(fig_line, use_container_width=True)

        st.subheader("📋 טבלת פירוט ניהולית עם אפשרות ייצוא")
        display_df = dash_df[[
            "PN", "Description", "Item_Type", "Supplier", "Status", "Assembly", "Assembly_Desc", 
            "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage"
        ]].rename(columns={
            "PN": "מק\"ט", "Description": "תיאור פריט", "Item_Type": "סוג פריט", "Supplier": "ספק",
            "Status": "סטטוס טיפול", "Assembly": "קוד הרכבה", "Assembly_Desc": "תיאור הרכבה",
            "Qty_Per_Assembly": "כמות נדרשת", "Assembly_Monthly_Build": "ת. ייצור",
            "Required_Demand": "ביקוש מדויק", "Stock": "מלאי", "Total_MRP_Shortage": "סך חוסר"
        })
        st.dataframe(display_df.sort_values(by="סך חוסר", ascending=False), use_container_width=True)

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

with TABS[1]:
    st.subheader(f"📊 סימולציית Clear To Build (CTB) לחודש: {selected_month_label}")
    st.markdown("המערכת מציגה את הרכיבים החסרים בלבד, את ה-ETA המדויק, ומדגישה ב-**BOLD** את הפריט הקריטי ביותר.")

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
            
            eta_display_str = get_first_supply_eta(c_pn)
            
            if eta_display_str and eta_display_str != "ללא ETA":
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
            eta_str = f" [ETA: {raw_eta}]" if raw_eta != "ללא ETA" else ""
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
        st.dataframe(cap_df, use_container_width=True)
    else:
        st.info(f"לא נמצאו הרכבות מתוכננות לייצור לחודש {selected_month_label}.")

with TABS[2]:
    st.subheader("💡 סימולציית What-If (מה יקרה אם...)")
    st.markdown("כלי אינטראקטיבי לקבלת החלטות: בדוק כיצד הוספת מלאי או פתרון מק\"ט משחרר את קווי הייצור.")
    
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        sim_pn = st.selectbox("בחר מק\"ט לסימולציה", sorted(df[PN_COL].dropna().astype(str).unique()), key="sim_pn")
    with col_w2:
        sim_extra_stock = st.number_input("תוספת כמות מדומיינת למלאי לצורך סימולציה", min_value=0.0, value=10.0, step=1.0)

    if st.button("🔮 הרץ סימולציית שחרור צוואר בקבוק"):
        st.success(f"סימולציה הופעלה בהצלחה עבור מק\"ט `{sim_pn}` עם תוספת של {sim_extra_stock} יחידות.")
        st.info("💡 המלצה ניהולית: סגירת החוסר בפריט זה תפחית באופן מיידי את הפער בקווי ההרכבה התלויים בו.")

with TABS[3]:
    st.subheader("📌 לוח מעקב סטטוסים (Kanban Pipeline)")
    st.markdown("מעקב ויזואלי אחר התקדמות הטיפול במק\"טים הגירעוניים מול ספקים ורכש.")
    
    k_col1, k_col2, k_col3, k_col4 = st.columns(4)
    with k_col1:
        st.markdown("### 📝 פתוח")
        open_items = breakdown_df[breakdown_df["Status"] == "פתוח"] if not breakdown_df.empty else pd.DataFrame()
        for _, r in open_items.head(6).iterrows():
            st.warning(f"**{r['PN']}**\n\n{r['Description'][:22]}")
    with k_col2:
        st.markdown("### 🛒 הוזמן / בטיפול")
        ordered_items = breakdown_df[breakdown_df["Status"].isin(["הוזמן", "בטיפול", "בייצור"])] if not breakdown_df.empty else pd.DataFrame()
        for _, r in ordered_items.head(6).iterrows():
            st.info(f"**{r['PN']}**\n\n{r['Description'][:22]}")
    with k_col3:
        st.markdown("### 🚚 בדרך לקו")
        shipping_items = breakdown_df[breakdown_df["Status"] == "בדרך"] if not breakdown_df.empty else pd.DataFrame()
        for _, r in shipping_items.head(6).iterrows():
            st.success(f"**{r['PN']}**\n\n{r['Description'][:22]}")
    with k_col4:
        st.markdown("### ✅ התקבל / סגור")
        received_items = breakdown_df[breakdown_df["Status"] == "התקבל"] if not breakdown_df.empty else pd.DataFrame()
        for _, r in received_items.head(6).iterrows():
            st.success(f"**{r['PN']}** (התקבל)")

with TABS[4]:
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
                status_options = ["פתוח", "בטיפול", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
                status_idx = status_options.index(saved_status) if saved_status in status_options else 0
                status = st.selectbox("סטטוס טיפול", status_options, index=status_idx)

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

with TABS[5]:
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
