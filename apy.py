import streamlit as st
import pandas as pd
import plotly.express as px

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="MRP Control Tower",
    page_icon="📦",
    layout="wide"
)

st.title("📊 דשבורד ניתוח חוסרים - MRP")
st.markdown("מערכת לניתוח חוסרים לפי תוכנית עבודה, הרכבות וסוגי פריטים")

# ==========================================================
# FILE UPLOAD
# ==========================================================
uploaded_file = st.sidebar.file_uploader("העלה קובץ MRP (mrp_2.xlsx)", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("אנא העלה את קובץ ה-MRP כדי להתחיל.")
    st.stop()

# ==========================================================
# DATA LOADING
# ==========================================================
@st.cache_data
def load_data(file):
    # קריאת הנתונים המרכזיים - הכותרות בשורה 30 (אינדקס 29)
    df = pd.read_excel(file, header=29)
    
    # קריאת שורה 29 בלבד כדי לחלץ את רמות העץ (BOM levels) של ההרכבות
    # שורה 29 באקסל היא אינדקס 28
    df_levels = pd.read_excel(file, header=None, skiprows=28, nrows=1)
    
    # ניקוי שמות עמודות
    df.columns = [str(c).strip() if pd.notnull(c) else c for c in df.columns]
    
    return df, df_levels

try:
    with st.spinner('טוען נתונים...'):
        df, df_levels = load_data(uploaded_file)
except Exception as e:
    st.error(f"שגיאה בטעינת הקובץ: {e}")
    st.stop()

# ==========================================================
# COLUMN MAPPING (Based on user specs)
# ==========================================================
# עמודות A-J הן אינדקסים 0-9. 
PN_COL = df.columns[1]     # עמודה B (מק"ט) - ניתן לשנות אם המק"ט בעמודה אחרת
DESC_COL = df.columns[4]   # עמודה E (תיאור) - ניתן לשנות אם התיאור בעמודה אחרת

# הרכבות: עמודות K עד AJ -> אינדקסים 10 עד 35
ASSEMBLY_COLS = df.columns[10:36].tolist()

# סוג פריט: עמודה AS -> אינדקס 44
ITEM_TYPE_COL = df.columns[44]

# מלאי: עמודה CB -> אינדקס 79
STOCK_COL = df.columns[79]

# ETA: עמודות CC עד CZ -> אינדקסים 80 עד 103
ETA_COLS = df.columns[80:104].tolist()

# מאזן חומרים (חודשים): עמודות DE עד EB -> אינדקסים 108 עד 131
MONTH_COLS = df.columns[108:132].tolist()

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("🔍 מסננים")

# 1. סינון חודש (מתוך עמודות מאזן החומרים)
month_options = {str(m): m for m in MONTH_COLS if pd.notnull(m)}
selected_month_str = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()))
selected_month = month_options[selected_month_str]

# 2. סינון הרכבה
selected_assembly = st.sidebar.selectbox("בחר הרכבה (Assembly)", ["הכל"] + ASSEMBLY_COLS)

# 3. סינון סוג פריט
item_types = df[ITEM_TYPE_COL].dropna().unique().tolist()
selected_item_type = st.sidebar.selectbox("בחר סוג פריט", ["הכל"] + item_types)


# ==========================================================
# APPLY FILTERS & CALCULATE SHORTAGES
# ==========================================================
filtered_df = df.copy()

# סינון לפי סוג פריט
if selected_item_type != "הכל":
    filtered_df = filtered_df[filtered_df[ITEM_TYPE_COL] == selected_item_type]

# סינון לפי הרכבה
if selected_assembly != "הכל":
    # נניח שפריט שייך להרכבה אם הכמות בעמודת ההרכבה גדולה מ-0
    filtered_df[selected_assembly] = pd.to_numeric(filtered_df[selected_assembly], errors='coerce').fillna(0)
    filtered_df = filtered_df[filtered_df[selected_assembly] > 0]

# חישוב חוסרים לחודש הנבחר (חוסר = מאזן שלילי)
filtered_df['Balance'] = pd.to_numeric(filtered_df[selected_month], errors='coerce').fillna(0)
shortage_df = filtered_df[filtered_df['Balance'] < 0].copy()

# המרת החוסר לערך מוחלט להצגה נוחה
shortage_df['Shortage_Qty'] = shortage_df['Balance'].abs()

# ==========================================================
# DASHBOARD UI
# ==========================================================
st.subheader(f"ניתוח חוסרים לחודש: {selected_month_str}")

# --- KPIs ---
col1, col2, col3 = st.columns(3)
total_shortage_items = len(shortage_df)
total_shortage_qty = shortage_df['Shortage_Qty'].sum()

col1.metric("🔴 פריטים בחוסר", total_shortage_items)
col2.metric("📦 כמות חסרה כוללת", f"{total_shortage_qty:,.0f}")

# הצגת רמת עץ ההרכבה אם נבחרה הרכבה ספציפית
if selected_assembly != "הכל":
    try:
        assembly_index = df.columns.get_loc(selected_assembly)
        bom_level = df_levels.iloc[0, assembly_index]
        col3.metric("🔗 רמת הרכבה (BOM Level)", str(bom_level))
    except:
        col3.metric("🔗 רמת הרכבה (BOM Level)", "לא נמצא")

st.divider()

if total_shortage_items > 0:
    # --- Bar Chart ---
    st.subheader("🔥 20 החוסרים הגדולים ביותר")
    top_shortages = shortage_df.sort_values(by='Shortage_Qty', ascending=False).head(20)
    
    # במידה ועמודות המק"ט והתיאור הוגדרו נכון, ניצור גרף לפי מק"ט
    fig = px.bar(
        top_shortages, 
        x='Shortage_Qty', 
        y=str(PN_COL), 
        orientation='h',
        text='Shortage_Qty',
        color='Shortage_Qty',
        color_continuous_scale='Reds',
        labels={'Shortage_Qty': 'כמות חסרה', str(PN_COL): 'מק"ט'}
    )
    fig.update_layout(yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig, use_container_width=True)

    # --- Data Table ---
    st.subheader("📋 רשימת חוסרים מפורטת")
    
    # עמודות להצגה בטבלה
    display_cols = [PN_COL, DESC_COL, ITEM_TYPE_COL, STOCK_COL, 'Shortage_Qty']
    # הוספת ה-ETA הראשון שנמצא (לדוגמה) או סתם הצגת העמודות העיקריות
    display_df = shortage_df[display_cols].sort_values(by='Shortage_Qty', ascending=False)
    
    # שינוי שמות עמודות לתצוגה ברורה
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
