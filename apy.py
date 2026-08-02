import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import plotly.express as px
from datetime import datetime, date

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="MRP Control Tower",
    page_icon="📦",
    layout="wide"
)

# ==========================================================
# DATABASE
# ==========================================================
conn = sqlite3.connect(
    "eta_updates.db",
    check_same_thread=False
)

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

# ==========================================================
# DB FUNCTIONS
# ==========================================================
def get_eta_record(pn):
    cur = conn.cursor()
    cur.execute("""
    SELECT eta, status, comment, updated_by, updated_at
    FROM eta_updates
    WHERE pn = ?
    """, (pn,))
    return cur.fetchone()

def save_eta_record(pn, eta, status, comment, updated_by):
    conn.execute("""
    INSERT OR REPLACE INTO eta_updates (pn, eta, status, comment, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (pn, eta, status, comment, updated_by, datetime.now().strftime("%Y-%m-%d %H:%M")))
    conn.commit()

# ==========================================================
# ETA COLOR
# ==========================================================
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
# HEADER
# ==========================================================
st.title("📦 MRP Control Tower")
st.markdown("MRP / Shortage / ETA Management System")

# ==========================================================
# FILE UPLOAD
# ==========================================================
uploaded_file = st.sidebar.file_uploader(
    "Upload MRP File",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("Upload MRP Excel file")
    st.stop()

# ==========================================================
# LOAD FILE
# ==========================================================
try:
    # הקוד קורא עכשיו מהשורה ה-30 (אינדקס 29)
    df = pd.read_excel(uploaded_file, header=29)
except Exception as e:
    st.error(str(e))
    st.stop()

df.columns = [str(c).strip() for c in df.columns]

# ==========================================================
# DATE DETECTION
# ==========================================================
date_columns = []
for col in df.columns:
    try:
        pd.to_datetime(col)
        date_columns.append(col)
    except:
        pass

selected_date = None
if len(date_columns):
    selected_date = st.sidebar.selectbox(
        "MRP Date",
        date_columns
    )

# ==========================================================
# SHORTAGE ENGINE
# ==========================================================
shortage_df = pd.DataFrame()
if selected_date:
    try:
        shortage_df = df[pd.to_numeric(df[selected_date], errors="coerce") < 0].copy()
    except:
        pass

# ==========================================================
# KPI
# ==========================================================
st.subheader("Executive Dashboard")
col1, col2, col3, col4 = st.columns(4)
blocked_items = len(shortage_df)

if len(shortage_df):
    total_shortage = abs(pd.to_numeric(shortage_df[selected_date], errors="coerce").sum())
else:
    total_shortage = 0

critical_items = 0
if len(shortage_df):
    critical_items = len(shortage_df[pd.to_numeric(shortage_df[selected_date], errors="coerce") < -100])

readiness = max(0, 100 - blocked_items)

col1.metric("Blocked Items", blocked_items)
col2.metric("Shortage Qty", int(total_shortage))
col3.metric("Critical Items", critical_items)
col4.metric("Readiness %", readiness)

st.divider()

# ==========================================================
# HEATMAP
# ==========================================================
st.subheader("🔥 Top Shortages")
if len(shortage_df):
    heat_df = shortage_df.copy()
    heat_df["Shortage"] = pd.to_numeric(heat_df[selected_date], errors="coerce")
    heat_df = heat_df.sort_values("Shortage")
    
    fig = px.bar(
        heat_df.head(20),
        x="Shortage",
        y=heat_df.columns[1], # שימוש בעמודת המקט (PN_ID) אם היא השנייה, למניעת בעיות תצוגה
        orientation="h",
        color="Shortage",
        color_continuous_scale="Reds"
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.success("No shortages found")

# ==========================================================
# TABLE
# ==========================================================
if len(shortage_df):
    st.subheader("🚨 Top 20 Shortages")
    tbl = shortage_df.copy()
    tbl["Shortage"] = pd.to_numeric(tbl[selected_date], errors="coerce")
    tbl = tbl.sort_values("Shortage").head(20)
    st.dataframe(tbl, use_container_width=True)

# ==========================================================
# PN SEARCH
# ==========================================================
st.divider()
st.subheader("🔍 PN Search")
search_text = st.text_input("Search Part Number / Description")

if search_text:
    result = df[
        df.astype(str).apply(
            lambda column: column.str.contains(search_text, case=False, na=False)
        ).any(axis=1)
    ]
    st.write(f"Found {len(result)} records")
    st.dataframe(result, use_container_width=True)

# ==========================================================
# ETA MANAGEMENT
# ==========================================================
st.divider()
st.subheader("📅 ETA Management")
pn_column = st.selectbox("Select PN Column", df.columns)
pn_values = sorted(df[pn_column].dropna().astype(str).unique())
selected_pn = st.selectbox("Part Number", pn_values)
record = get_eta_record(selected_pn)
default_comment = ""

if record:
    default_comment = record[2]

with st.form("eta_form"):
    eta_date = st.date_input("ETA Date")
    status = st.selectbox(
        "Status",
        ["Open", "Ordered", "In Production", "In Transit", "Received", "Blocked"]
    )
    comment = st.text_area("Comment", value=default_comment)
    updated_by = st.text_input("Updated By")
    save_btn = st.form_submit_button("Save ETA")

if save_btn:
    save_eta_record(selected_pn, str(eta_date), status, comment, updated_by)
    st.success("ETA Updated Successfully")

# ==========================================================
# ETA OVERVIEW
# ==========================================================
st.divider()
st.subheader("🚦 ETA Tracker")
eta_rows = []

for pn in pn_values:
    rec = get_eta_record(pn)
    if rec:
        eta_rows.append({
            "PN": pn,
            "ETA": rec[0],
            "Risk": eta_color(rec[0]),
            "Status": rec[1],
            "Comment": rec[2],
            "Owner": rec[3],
            "Updated": rec[4]
        })

eta_df = pd.DataFrame(eta_rows)

if len(eta_df):
    red = len(eta_df[eta_df["Risk"] == "🔴"])
    orange = len(eta_df[eta_df["Risk"] == "🟠"])
    yellow = len(eta_df[eta_df["Risk"] == "🟡"])
    green = len(eta_df[eta_df["Risk"] == "🟢"])
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔴 Critical", red)
    c2.metric("🟠 14 Days", orange)
    c3.metric("🟡 30 Days", yellow)
    c4.metric("🟢 Healthy", green)
    
    st.dataframe(eta_df, use_container_width=True)

# ==========================================================
# BUILD CAPACITY (DUMMY FIRST VERSION)
# ==========================================================
st.divider()
st.subheader("🏭 Build Capacity")

capacity_df = pd.DataFrame({
    "Assembly": ["FE", "BE", "QGC", "GNG", "BFU"],
    "Build Qty": [
        np.random.randint(0, 50),
        np.random.randint(0, 50),
        np.random.randint(0, 50),
        np.random.randint(0, 50),
        np.random.randint(0, 50)
    ]
})

fig2 = px.bar(
    capacity_df,
    x="Assembly",
    y="Build Qty",
    color="Build Qty"
)
st.plotly_chart(fig2, use_container_width=True)

# ==========================================================
# RAW DATA
# ==========================================================
with st.expander("Raw Data"):
    st.dataframe(df, use_container_width=True)
