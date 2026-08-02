import streamlit as st
import pandas as pd
import plotly.express as px

# ==========================================================
# CONFIGURATION
# ==========================================================
# כאן אתה מגדיר את הנתיב הקבוע לקובץ במחשב שלך
# שים לב להשתמש בלוכסנים קדמיים (/) או בלוכסן כפול (\\)
FILE_PATH = "mrp_2.xlsx" # לדוגמה: "C:/Users/YourName/Documents/mrp_2.xlsx"

st.set_page_config(
    page_title="MRP Control Tower",
    page_icon="📦",
    layout="wide"
)

st.title("📊 דשבורד ניתוח חוסרים - MRP")
st.markdown("מערכת לניתוח חוסרים לפי תוכנית עבודה, הרכבות וסוגי פריטים")

# ==========================================================
# DATA LOADING
# ==========================================================
@st.cache_data
def load_data(file_path):
    # קריאת הנתונים המרכזיים - הכותרות בשורה 30 (אינדקס 29)
    df = pd.read_excel(file_path, header=29)
    
    # קריאת שורה 29 בלבד כדי לחלץ את רמות העץ (BOM levels) של ההרכבות
    df_levels = pd.read_excel(file_path, header=None, skiprows=28, nrows=1)
    
    # קריאת שורה 28 בלבד כדי לחלץ את תיאורי ההרכבות
    df_desc = pd.read_excel(file_path, header=None, skiprows=27, nrows=1)
    
    # ניקוי שמות עמודות
    df.columns = [str(c).strip() if pd.notnull(c) else c for c in df.columns]
    
    return df, df_levels, df_desc

try:
    with st.spinner('טוען נתונים מהקובץ המקומי...'):
        df, df_levels, df_desc = load_data(FILE_PATH)
except Exception as e:
    st.error(f"שגיאה בטעינת הקובץ מהנתיב: {FILE_PATH}. ודא שהקובץ סגור ושהנתיב נכון.\nפירוט השגיאה: {e}")
    st.stop()

# ==========================================================
# COLUMN MAPPING (Based on user specs)
# ==========================================================
PN_COL = df.columns[1]     
DESC_COL = df.columns[4]   

# הרכבות: עמודות K עד AJ -> אינדקסים 10 עד 35
ASSEMBLY_COLS = df.columns[10:36].tolist()

ITEM_TYPE_COL = df.columns[44]
STOCK_COL = df.columns[79]
MONTH_COLS = df.columns[108:132].tolist()

# ==========================================================
# SIDEBAR FILTERS
# ==========================================================
st.sidebar.header("🔍 מסננים")

# 1. סינון חודש
month_options = {str(m): m for m in MONTH_COLS if pd.notnull(m)}
selected_month_str = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()))
selected_month = month_options[selected_month_str]

# 2. סינון הרכבה + הוספת תיאור לתיבת הבחירה
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

# 3. סינון סוג פריט
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

# חישוב חוסרים
filtered_df['Balance'] = pd.to_numeric(filtered_df[selected_month], errors='coerce').fillna(0)
shortage_df = filtered_df[filtered_df['Balance'] < 0].copy()
shortage_df['Shortage_Qty'] = shortage_df['Balance'].abs()

# ==========================================================
# DASHBOARD UI
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
    st.subheader("🔥 20 החוסרים הגדולים ביותר")
    top_shortages = shortage_df.sort_values(by='Shortage_Qty', ascending=False).head(20)
    
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

    st.subheader("📋 רשימת חוסרים מפורטת")
    
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
