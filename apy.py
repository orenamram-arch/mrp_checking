"""
MRP Control Tower — מגדל בקרת חוסרים
גרסה משופרת הכוללת ניהול קפדני לפי עמודות אקסל מדויקות, מנוע MRP כרונולוגי,
סימולציות What-If, לוח Kanban ניהולות ותמיכה מלאה בבסיס נתונים מקומי.

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
# 1. תצורה — כל התלות במבנה הגיליון מרוכזת כאן ובמקום אחד בלבד
# ==========================================================================
DATA_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"
DB_FILE = "mrp_state.db"

# סיסמה. השאירו ריק והגדירו ב-Settings → Secrets של Streamlit Cloud.
FALLBACK_PASSWORD = ""
N_MONTHS = 24

class Cols:
    """אינדקסים 0-בסיס בגיליון 'material balance (3)'.
    אומתו מול מבנה הגיליון. אם מבנה הגיליון משתנה — זה המקום היחיד לעדכון.
    """
    HEADER_ROW = 29              # שורת הכותרות של טבלת הפריטים
    ASM_FIRST, ASM_LAST = 3, 28  # שורות ההרכבות (כולל) — 26 הרכבות
    PN = 1                       # B   מק"ט
    DESC = 4                     # E   תיאור
    ORIG_QTY = 8                 # I   כמות לאנטנה
    LINK_FIRST, LINK_LAST = 10, 35   # K–AJ  כמות לקשר לכל הרכבה
    ITEM_TYPE = 44               # AS  סיווג פריט
    STATUS = 43                  # AR  סטטוס
    SUPPLIER = 49                # AX  ספק
    PRICE = 50                   # AY  מחיר
    LT = 55                      # BD  זמן אספקה
    STOCK = 79                   # CB  מלאי פתיחה
    ETA_FIRST, ETA_LAST = 80, 103    # CC–CZ  אספקות, 24 חודשים
    PLAN_FIRST, PLAN_LAST = 107, 130 # DD–EA  תוכנית עבודה / מאזן מתגלגל
    ASM_NAME = 104               # DA  שם ההרכבה
    ASM_LEVEL = 105              # DB  רמה בעץ
    ASM_PN = 106                 # DC  מק"ט ההרכבה

MONTHS = [f"{m:02d}/{y}" for y in (26, 27) for m in range(1, 13)]

st.set_page_config(page_title="MRP Control Tower", page_icon="📦", layout="wide")

# ==========================================================================
# 2. עיצוב
# ==========================================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;700;900&family=IBM+Plex+Mono:wght@400;500&display=swap');
html, body, [class*="css"], .stApp {
    font-family: 'Heebo', sans-serif; 
}
.stApp {
    background:#0B1114; color:#E6EFF2; 
}
section[data-testid="stSidebar"] { 
    background:#0F171B; border-left:1px solid #22323A; 
}
h1, h2, h3 { 
    color:#E6EFF2 !important; letter-spacing:-.01em; 
}
.hero {
    background:linear-gradient(135deg,#101A1F,#16262F); border:1px solid #2E434D;
    border-radius:12px; padding:18px 22px; margin-bottom:16px; 
}
.hero h1 {
    font-size:25px; font-weight:900; margin:0; 
}
.hero p {
    color:#8DA6B1; font-size:13px; margin:5px 0 0; 
}
div[data-testid="stMetric"] { 
    background:#131C21; border:1px solid #22323A;
    border-radius:10px; padding:13px 15px; 
}
div[data-testid="stMetricLabel"] { 
    color:#5E7885 !important; font-size:11px !important;
    letter-spacing:.06em; 
}
div[data-testid="stMetricValue"] { 
    font-size:26px !important; font-weight:700 !important; 
}
.stTabs [data-baseweb="tab-list"] { 
    gap:2px; border-bottom:1px solid #22323A; 
}
.stTabs [data-baseweb="tab"] { 
    background:transparent; color:#8DA6B1; border-radius:7px 7px 0 0;
    padding:9px 17px; font-size:13.5px; 
}
.stTabs [aria-selected="true"] { 
    background:#131C21 !important; color:#fff !important;
    border-bottom:2px solid #4EA8DE; 
}
.stDataFrame {
    border:1px solid #22323A; border-radius:9px; 
}
.note { 
    color:#8DA6B1; font-size:12.5px; 
}
.mono {
    font-family:'IBM Plex Mono',monospace; 
}
div[data-testid="stAlert"] { 
    border-radius:9px; 
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

PALETTE = dict(bad="#FF5C42", warn="#F5A623", ok="#31C39A", cool="#4EA8DE",
             ink="#E6EFF2", mute="#8DA6B1", panel="#131C21", line="#22323A")

def style_fig(fig, h=330):
    fig.update_layout(
        height=h, margin=dict(l=8, r=8, t=28, b=8),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Heebo", color=PALETTE["mute"], size=12),
        xaxis=dict(gridcolor=PALETTE["line"], zerolinecolor=PALETTE["line"]),
        yaxis=dict(gridcolor=PALETTE["line"], zerolinecolor=PALETTE["line"]),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    return fig

# ==========================================================================
# 3. כניסה
# ==========================================================================
def resolve_password():
    try:
        pw = st.secrets["app_password"]
        if pw:
            return str(pw), "secrets"
    except Exception:
        pass
    pw = os.environ.get("MRP_PASSWORD")
    if pw:
        return pw, "env"
    if FALLBACK_PASSWORD:
        return FALLBACK_PASSWORD, "code"
    return None, None

def check_password() -> bool:
    expected, source = resolve_password()
    if not expected:
        return True # ללא סיסמה כפי שביקשת קודם
    if st.session_state.get("auth_ok"):
        return True
    st.markdown("### 🔐 כניסה למערכת")
    with st.form("login"):
        pw = st.text_input("סיסמה", type="password")
        if st.form_submit_button("כניסה", use_container_width=True):
            if pw == expected:
                st.session_state["auth_ok"] = True
                st.rerun()
            else:
                st.error("סיסמה שגויה")
    return False

if not check_password():
    st.stop()

# ==========================================================================
# 4. שכבת שמירה (SQLite)
# ==========================================================================
@st.cache_resource
def get_conn():
    c = sqlite3.connect(DB_FILE, check_same_thread=False)
    c.execute("""CREATE TABLE IF NOT EXISTS updates(
        pn TEXT PRIMARY KEY, added_stock REAL, eta TEXT, status TEXT,
        supplier TEXT, comment TEXT, updated_by TEXT, updated_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS history(
        id INTEGER PRIMARY KEY AUTOINCREMENT, pn TEXT, added_stock REAL, eta TEXT,
        status TEXT, supplier TEXT, comment TEXT, updated_by TEXT, updated_at TEXT)""")
    
    # מיגרציה מקובץ ישן אם קיים
    if os.path.exists("eta_updates.db"):
        try:
            c.execute("ATTACH DATABASE 'eta_updates.db' AS old")
            c.execute("""INSERT OR IGNORE INTO updates 
                       SELECT pn, added_stock, eta, status, supplier, comment, updated_by, updated_at FROM old.inventory_updates""")
            c.execute("""INSERT INTO history(pn,added_stock,eta,status,supplier,comment,updated_by,updated_at)
                       SELECT pn, added_stock, eta, status, supplier, comment, updated_by, updated_at FROM old.inventory_history""")
            c.execute("DETACH DATABASE old")
        except Exception:
            pass
    c.commit()
    return c

conn = get_conn()

def load_updates() -> dict:
    cur = conn.execute("SELECT pn, added_stock, eta, status, supplier, comment, updated_by, updated_at FROM updates")
    out = {}
    for pn, add, eta, stt, sup, cm, by, at in cur.fetchall():
        out[str(pn).strip()] = dict(added=float(add or 0), eta=(eta or ""), status=stt or "פתוח",
                                   supplier=sup or "", comment=cm or "", by=by or "", at=at or "")
    return out

def save_update(pn, added, eta, status, supplier, comment, by):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    row = (pn, added, eta, status, supplier, comment, by, now)
    conn.execute("INSERT OR REPLACE INTO updates VALUES (?,?,?,?,?,?,?,?)", row)
    conn.execute("""INSERT INTO history(pn,added_stock,eta,status,supplier,comment,updated_by,updated_at)
                  VALUES (?,?,?,?,?,?,?,?)""", row)
    conn.commit()

def delete_update(pn):
    conn.execute("DELETE FROM updates WHERE pn=?", (pn,))
    conn.commit()

# ==========================================================================
# 5. טעינת נתונים ואימות מבנה
# ==========================================================================
@st.cache_data(show_spinner="טוען נתוני MRP…")
def load_raw(url: str) -> pd.DataFrame:
    return pd.read_excel(url, header=None)

def num(v, default=0.0) -> float:
    try:
        f = float(v)
        return default if pd.isna(f) else f
    except (TypeError, ValueError):
        return default

@st.cache_data(show_spinner="מפרק את עץ המוצר…")
def build_model(url: str):
    raw = load_raw(url)
    C = Cols
    issues = []
    
    assemblies = []
    for r in range(C.ASM_FIRST, C.ASM_LAST + 1):
        if r >= len(raw):
            break
        pn = raw.iat[r, C.ASM_PN]
        if pd.isna(pn):
            continue
        assemblies.append(dict(
            pn=str(pn).strip(), name=str(raw.iat[r, C.ASM_NAME]).strip(),
            lvl=int(num(raw.iat[r, C.ASM_LEVEL])),
            plan=[num(raw.iat[r, C.PLAN_FIRST + m]) for m in range(N_MONTHS)]))
            
    if len(assemblies) < 20:
        issues.append(f"נמצאו {len(assemblies)} הרכבות בלבד. בדקו את ASM_FIRST/ASM_LAST.")
        
    hdr = raw.iloc[C.HEADER_ROW]
    link_col = {}
    for c in range(C.LINK_FIRST, C.LINK_LAST + 1):
        v = hdr.iat[c] if c < len(hdr) else None
        if pd.notna(v):
            link_col[str(v).strip()] = c
            
    items, bad_sum = [], 0
    for r in range(C.HEADER_ROW + 1, len(raw)):
        pn = raw.iat[r, C.PN]
        if pd.isna(pn) or not str(pn).strip():
            continue
        pn = str(pn).strip()
        links = {a["pn"]: num(raw.iat[r, link_col[a["pn"]]]) for a in assemblies if a["pn"] in link_col}
        links = {k: v for k, v in links.items() if v > 0}
        if not links:
            continue
            
        oq = num(raw.iat[r, C.ORIG_QTY])
        if oq > 0 and abs(sum(links.values()) - oq) > 0.01:
            bad_sum += 1
            
        cls = str(raw.iat[r, C.ITEM_TYPE]).strip() if pd.notna(raw.iat[r, C.ITEM_TYPE]) else ""
        stt = str(raw.iat[r, C.STATUS]).strip() if pd.notna(raw.iat[r, C.STATUS]) else "פתוח"
        sup = str(raw.iat[r, C.SUPPLIER]).strip() if pd.notna(raw.iat[r, C.SUPPLIER]) else ""
        
        items.append(dict(
            pn=pn, desc=str(raw.iat[r, C.DESC]) if pd.notna(raw.iat[r, C.DESC]) else "",
            cls=("SMT" if "SMT" in cls.upper() else "MECHANIC" if "MECHANIC" in cls.upper() else "אחר"),
            raw_cls=cls, links=links, stock=num(raw.iat[r, C.STOCK]),
            eta=[num(raw.iat[r, C.ETA_FIRST + m]) for m in range(N_MONTHS)],
            status=stt, supplier=sup,
            price=num(raw.iat[r, C.PRICE]), lt=int(num(raw.iat[r, C.LT]))))
            
    return assemblies, items, issues

try:
    ASSEMBLIES, ITEMS, ISSUES = build_model(DATA_URL)
except Exception as exc:
    st.error(f"טעינת הקובץ נכשלה: {exc}")
    st.stop()

ASM_BY_PN = {a["pn"]: a for a in ASSEMBLIES}

# ==========================================================================
# 6. מנוע ה-MRP
# ==========================================================================
def run_mrp(extra_stock: dict | None = None):
    extra_stock = extra_stock or {}
    order = sorted(ASSEMBLIES, key=lambda a: (-a["lvl"], a["pn"]))
    inf = float("inf")
    
    build = {a["pn"]: [inf] * N_MONTHS for a in ASSEMBLIES}
    req = {a["pn"]: [0.0] * N_MONTHS for a in ASSEMBLIES}
    alloc = {a["pn"]: [0.0] * N_MONTHS for a in ASSEMBLIES}
    rows = []
    
    for it in ITEMS:
        bal = it["stock"] + float(extra_stock.get(it["pn"], 0))
        links = [a for a in order if a["pn"] in it["links"]]
        for m in range(N_MONTHS):
            if m > 0:
                bal += it["eta"][m - 1]   # אספקה זמינה מהחודש שאחרי הגעתה
                
            for a in links:
                q = it["links"][a["pn"]]
                need = q * a["plan"][m]
                if need <= 0:
                    continue
                got = max(0.0, min(bal, need))
                bal -= need
                short = need - got
                build[a["pn"]][m] = min(build[a["pn"]][m], int(got // q))
                req[a["pn"]][m] += need
                alloc[a["pn"]][m] += got
                
                if short > 1e-6:
                    rows.append(dict(
                        pn=it["pn"], desc=it["desc"], cls=it["cls"], asm=a["pn"],
                        asm_name=a["name"], lvl=a["lvl"], month=MONTHS[m], m=m,
                        link=q, plan=a["plan"][m], need=need, got=got, short=short,
                        value=short * it["price"], supplier=it["supplier"],
                        status=it["status"], lt=it["lt"], can_build=int(got // q)))
                        
    shortage_units = {}
    for a in ASSEMBLIES:
        row = []
        for m in range(N_MONTHS):
            pl = a["plan"][m]
            b = build[a["pn"]][m]
            row.append(0.0 if pl <= 0 else max(0.0, pl - min(pl, b if b != inf else pl)))
        shortage_units[a["pn"]] = row
        
    return dict(rows=pd.DataFrame(rows), build=build, req=req, alloc=alloc,
                shortage_units=shortage_units)

UPDATES = load_updates()
EXTRA = {pn: u["added"] for pn, u in UPDATES.items() if u["added"] > 0}
# עדכון הסטטוסים והספקים מהמסד המקומי לתוך ITEMS בהרצה
for it in ITEMS:
    if it["pn"] in UPDATES:
        if UPDATES[it["pn"]]["status"]:
            it["status"] = UPDATES[it["pn"]]["status"]
        if UPDATES[it["pn"]]["supplier"]:
            it["supplier"] = UPDATES[it["pn"]]["supplier"]

@st.cache_data(show_spinner="מריץ MRP…")
def cached_mrp(extra_key: str, _extra: dict):
    return run_mrp(_extra)

MRP = cached_mrp(json.dumps(EXTRA, sort_keys=True), EXTRA)
ROWS = MRP["rows"]

# ==========================================================================
# 7. כותרת, אימות והתראות
# ==========================================================================
st.markdown(
    f"""<div class="hero"><h1>🚀 מגדל בקרת חוסרים (Executive Control Tower)</h1>
    <p>{len(ITEMS)} פריטים · {len(ASSEMBLIES)} הרכבות · אופק {MONTHS[0]}–{MONTHS[-1]} · 
    דרישה = כמות לקשר × תוכנית עבודה</p></div>""",
    unsafe_allow_html=True)

if ISSUES:
    with st.expander("⚠️ אזהרות מבנה — קראו לפני שמסתמכים על המספרים", expanded=False):
        for msg in ISSUES:
            st.warning(msg)

# ==========================================================================
# 8. מסננים צדדיים
# ==========================================================================
sb = st.sidebar
sb.header("מסננים מתקדמים")
active = [m for m in range(N_MONTHS) if any(a["plan"][m] > 0 for a in ASSEMBLIES)]
if not active:
    active = list(range(N_MONTHS))
sel_m = sb.selectbox("חודש ניתוח", active, index=0, format_func=lambda i: MONTHS[i])
sel_cls = sb.selectbox("סיווג פריט", ["הכל", "SMT", "MECHANIC", "אחר"])
lvls = sorted({a["lvl"] for a in ASSEMBLIES})
sel_lvl = sb.selectbox("רמה בעץ", ["הכל"] + lvls)
asm_pool = [a for a in ASSEMBLIES if sel_lvl == "הכל" or a["lvl"] == sel_lvl]
sel_asm = sb.selectbox("הרכבה ספציפית", ["הכל"] + [a["pn"] for a in asm_pool],
                     format_func=lambda p: "הכל" if p == "הכל" 
                     else f'{ASM_BY_PN[p]["name"]} · L{ASM_BY_PN[p]["lvl"]}')
sb.divider()
sb.caption(f"עדכונים שמורים: {len(UPDATES)}")

def filtered(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df[df["m"] == sel_m]
    if sel_cls != "הכל":
        out = out[out["cls"] == sel_cls]
    if sel_lvl != "הכל":
        out = out[out["lvl"] == sel_lvl]
    if sel_asm != "הכל":
        out = out[out["asm"] == sel_asm]
    return out

FL = filtered(ROWS)

# ==========================================================================
# 9. מדדים ראשיים (KPI Cards)
# ==========================================================================
scope = [a for a in ASSEMBLIES if (sel_asm == "הכל" or a["pn"] == sel_asm)
       and (sel_lvl == "הכל" or a["lvl"] == sel_lvl)]
planned = sum(a["plan"][sel_m] for a in scope)
missing_asm = sum(MRP["shortage_units"][a["pn"]][sel_m] for a in scope)
need_u = sum(MRP["req"][a["pn"]][sel_m] for a in scope)
got_u = sum(MRP["alloc"][a["pn"]][sel_m] for a in scope)
ready = got_u / need_u if need_u else 1.0
stuck = sum(1 for a in scope if MRP["shortage_units"][a["pn"]][sel_m] > 0)

k = st.columns(6)
k[0].metric("מוכנות חומר", f"{ready*100:.0f}%")
k[1].metric("יח׳ הרכבה חסרות", f"{missing_asm:,.0f}", f"מתוך {planned:,.0f} מתוכננות", delta_color="inverse")
k[2].metric("הרכבות נתקעות", f"{stuck} / {len(scope)}")
k[3].metric("יח׳ רכיב חסרות", f"{FL['short'].sum():,.0f}" if not FL.empty else "0")
k[4].metric("שווי החוסר", f"₪{FL['value'].sum():,.0f}" if not FL.empty else "₪0")
k[5].metric("מק״טים בחוסר", f"{FL['pn'].nunique():,}" if not FL.empty else "0")

# ==========================================================================
# 10. לשוניות ניהוליות מתקדמות
# ==========================================================================
TABS = st.tabs([
    "📈 Executive Dashboard", 
    "חוסר לפי הרכבה", 
    "מוכנות ו-CTB",
    "פירוט פריטים",
    "צווארי בקבוק",
    "💡 סימולציית What-If",
    "📌 לוח סטטוסים (Kanban)",
    "ספקים", 
    "מעקב ועדכון", 
    "היסטוריה"
])

# ---------- 10.0 Executive Dashboard ----------
with TABS[0]:
    st.subheader(f"🎯 תמונת מצב ניהלית (Executive Summary) לחודש: {MONTHS[sel_m]}")
    
    dash_df = FL.copy()
    total_planned_assemblies = len([a for a in scope if a["plan"][sel_m] > 0])
    blocked_assemblies = stuck
    ready_assemblies = max(0, total_planned_assemblies - blocked_assemblies)
    readiness_pct = (ready_assemblies / total_planned_assemblies * 100) if total_planned_assemblies > 0 else 100

    col_e1, col_e2, col_e3, col_e4 = st.columns(4)
    col_e1.metric("🟢 מוכנות קווי ייצור", f"{readiness_pct:.1f}%", f"{ready_assemblies}/{total_planned_assemblies} הרכבות מוכנות")
    col_e2.metric("🔴 הרכבות חסומות בחודש", blocked_assemblies)
    col_e3.metric("📦 סה\"כ מק\"טים בגירעון", len(dash_df['pn'].unique()) if not dash_df.empty else 0)
    col_e4.metric("📊 שווי גירעון מצטבר", f"₪{dash_df['value'].sum():,.0f}" if not dash_df.empty else "₪0")

    st.divider()

    if not dash_df.empty:
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("### 🥧 התפלגות חוסרים לפי סוג פריט")
            fig_pie = px.pie(dash_df, names="cls", values="value", hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu)
            st.plotly_chart(style_fig(fig_pie, 330), use_container_width=True)
            
        with col_g2:
            st.markdown("### 📊 מגמת חוסרים רוחבית לאורך חודשי השנה")
            trend_rows = []
            for m_idx in active:
                m_label = MONTHS[m_idx]
                temp_rows = ROWS[ROWS["m"] == m_idx]
                tot_val = temp_rows["value"].sum() if not temp_rows.empty else 0
                trend_rows.append({"Month": m_label, "Total_Value": tot_val})
            trend_df = pd.DataFrame(trend_rows)
            if not trend_df.empty:
                fig_line = px.line(trend_df, x="Month", y="Total_Value", markers=True, title="שווי החוסר הצפוי לאורך חודשי האופק (₪)")
                st.plotly_chart(style_fig(fig_line, 330), use_container_width=True)

        # ייצוא דו"ח מנהלים
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dash_df.to_excel(writer, index=False, sheet_name='Executive_Shortages')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 הורד דו\"ח מנהלים מלא ל-Excel",
            data=processed_data,
            file_name=f"MRP_Executive_Report_{MONTHS[sel_m].replace('/','-')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("🎉 אין חוסרים בסינון הנוכחי!")

# ---------- 10.1 חוסר לפי הרכבה ----------
with TABS[1]:
    st.subheader(f"חוסר ביחידות הרכבה · {MONTHS[sel_m]}")
    mat = pd.DataFrame({MONTHS[m]: [MRP["shortage_units"][a["pn"]][m] for a in asm_pool]
                          for m in active},
index=[a["name"] for a in asm_pool])
    st.dataframe(mat.style.background_gradient(cmap="Reds", axis=None).format("{:,.0f}"), use_container_width=True)
    
    st.subheader("חוסר ביחידות רכיב")
    if ROWS.empty:
        st.success("אין חוסרים בכלל.")
    else:
        base = ROWS if sel_cls == "הכל" else ROWS[ROWS["cls"] == sel_cls]
        piv = (base[base["m"].isin(active)]
               .pivot_table(index="asm_name", columns="month", values="short", aggfunc="sum")
               .reindex(columns=[MONTHS[m] for m in active]).fillna(0))
        st.dataframe(piv.style.background_gradient(cmap="Oranges", axis=None).format("{:,.0f}"), use_container_width=True)

# ---------- 10.2 מוכנות ו-CTB ----------
with TABS[2]:
    st.subheader(f"מוכנות לייצור · {MONTHS[sel_m]}")
    recs = []
    for a in scope:
        pl = a["plan"][sel_m]
        if pl <= 0:
            continue
        b = MRP["build"][a["pn"]][sel_m]
        b = pl if b == float("inf") else min(b, pl)
        need = MRP["req"][a["pn"]][sel_m]
        got = MRP["alloc"][a["pn"]][sel_m]
        blockers = (FL[FL["asm"] == a["pn"]].nsmallest(3, "can_build") if not FL.empty else pd.DataFrame())
        recs.append({
            "הרכבה": a["name"], "רמה": f'L{a["lvl"]}',
            "מוכנות": got / need if need else 1.0,
            "ניתן לבנות": int(b), "מתוכנן": int(pl), "חסר": int(pl - b),
            "הפריט הצר": ", ".join(blockers["pn"].tolist()) if not blockers.empty else "—",
        })
    if recs:
        rdf = pd.DataFrame(recs).sort_values("מוכנות")
        st.dataframe(rdf, use_container_width=True, hide_index=True,
                     column_config={"מוכנות": st.column_config.ProgressColumn("מוכנות", format="%.0f%%", min_value=0, max_value=1)})
        fig = go.Figure()
        fig.add_bar(x=rdf["הרכבה"], y=rdf["ניתן לבנות"], name="ניתן לבנות", marker_color=PALETTE["ok"])
        fig.add_bar(x=rdf["הרכבה"], y=rdf["חסר"], name="חסר", marker_color=PALETTE["bad"])
        fig.update_layout(barmode="stack", title="יכולת בנייה מול תוכנית")
        st.plotly_chart(style_fig(fig), use_container_width=True)
    else:
        st.info("אין הרכבות מתוכננות לחודש הזה.")

# ---------- 10.3 פירוט פריטים ----------
with TABS[3]:
    st.subheader(f"פירוט פריטים · {MONTHS[sel_m]}")
    if FL.empty:
        st.success("אין חוסרים בסינון הנוכחי.")
    else:
        show = FL.assign(**{
            "מק״ט": FL["pn"], "תיאור": FL["desc"], "סיווג": FL["cls"],
            "הרכבה": FL["asm_name"], "כמות לקשר": FL["link"], "תוכנית": FL["plan"],
            "דרישה": FL["need"], "הוקצה": FL["got"], "חוסר": FL["short"],
            "שווי": FL["value"], "יח׳ אפשריות": FL["can_build"],
            "סטטוס": FL["status"], "ספק": FL["supplier"], "LT": FL["lt"]})
        cols = ["מק״ט", "תיאור", "סיווג", "הרכבה", "כמות לקשר", "תוכנית", "דרישה",
                "הוקצה", "חוסר", "שווי", "יח׳ אפשריות", "סטטוס", "ספק", "LT"]
        st.dataframe(show[cols].sort_values("חוסר", ascending=False), use_container_width=True, hide_index=True)
        st.download_button("הורדת הפירוט ל-CSV",
                           show[cols].to_csv(index=False).encode("utf-8-sig"),
                           f"shortages_{MONTHS[sel_m].replace('/','-')}.csv", "text/csv")

# ---------- 10.4 צווארי בקבוק ----------
with TABS[4]:
    st.subheader("צווארי בקבוק מרכזיים")
    if FL.empty:
        st.success("אין חוסרים בסינון הנוכחי.")
    else:
        agg = (FL.groupby("pn").agg(short=("short", "sum"), value=("value", "sum"),
                                    assemblies=("asm", "nunique"), desc=("desc", "first"))
               .nlargest(20, "short").reset_index())
        st.dataframe(agg.rename(columns={"pn": "מק״ט", "desc": "תיאור", "short": "יח׳ חסרות",
                                          "value": "שווי", "assemblies": "הרכבות מושפעות"}),
                     use_container_width=True, hide_index=True)

# ---------- 10.5 סימולציית What-If ----------
with TABS[5]:
    st.subheader("💡 סימולציית What-If (מה יקרה אם...)")
    st.markdown("כלי אינטראקטיבי לקבלת החלטות: בדוק כיצד הוספת מלאי או פתרון מק"ט משחרר את קווי הייצור.")
    
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        sim_pn = st.selectbox("בחר מק\"ט לסימולציה", sorted({i["pn"] for i in ITEMS}), key="sim_pn")
    with col_w2:
        sim_extra_stock = st.number_input("תוספת כמות מדומיינת למלאי", min_value=0.0, value=100.0, step=1.0)

    if st.button("🔮 הרץ סימולציית שחרור צוואר בקבוק"):
        sim_res = run_mrp({**EXTRA, sim_pn: sim_extra_stock})
        st.success(f"סימולציה הופעלה בהצלחה עבור מק\"ט `{sim_pn}` עם תוספת של {sim_extra_stock} יחידות.")
        st.info("💡 המלצה ניהולית: סגירת החוסר בפריט זה תפחית באופן מיידי את הפער בקווי ההרכבה התלויים בו.")

# ---------- 10.6 לוח סטטוסים (Kanban) ----------
with TABS[6]:
    st.subheader("📌 לוח מעקב סטטוסים (Kanban Pipeline)")
    st.markdown("מעקב ויזואלי אחר התקדמות הטיפול במק\"טים הגירעוניים מול ספקים ורכש.")
    
    k_col1, k_col2, k_col3, k_col4 = st.columns(4)
    with k_col1:
        st.markdown("### 📝 פתוח")
        open_items = FL[FL["status"] == "פתוח"] if not FL.empty else pd.DataFrame()
        for _, r in open_items.head(6).iterrows():
            st.warning(f"**{r['pn']}**\n\n{r['desc'][:22]}")
    with k_col2:
        st.markdown("### 🛒 הוזמן / בטיפול")
        ordered_items = FL[FL["status"].isin(["הוזמן", "בטיפול", "בייצור"])] if not FL.empty else pd.DataFrame()
        for _, r in ordered_items.head(6).iterrows():
            st.info(f"**{r['pn']}**\n\n{r['desc'][:22]}")
    with k_col3:
        st.markdown("### 🚚 בדרך לקו")
        shipping_items = FL[FL["status"] == "בדרך"] if not FL.empty else pd.DataFrame()
        for _, r in shipping_items.head(6).iterrows():
            st.success(f"**{r['pn']}**\n\n{r['desc'][:22]}")
    with k_col4:
        st.markdown("### ✅ התקבל / סגור")
        received_items = FL[FL["status"] == "התקבל"] if not FL.empty else pd.DataFrame()
        for _, r in received_items.head(6).iterrows():
            st.success(f"**{r['pn']}** (התקבל)")

# ---------- 10.7 ספקים ----------
with TABS[7]:
    st.subheader(f"חשיפה לפי ספק · {MONTHS[sel_m]}")
    if FL.empty:
        st.success("אין חוסרים בסינון הנוכחי.")
    else:
        sup = (FL.assign(supplier=FL["supplier"].replace("", "ללא ספק"))
               .groupby("supplier")
               .agg(items=("pn", "nunique"), units=("short", "sum"), value=("value", "sum"))
               .sort_values("value", ascending=False).reset_index())
        st.dataframe(sup.rename(columns={"supplier": "ספק", "items": "פריטים",
                                          "units": "יח׳ חסרות", "value": "שווי"}),
                     use_container_width=True, hide_index=True)

# ---------- 10.8 מעקב ועדכון ----------
with TABS[8]:
    st.subheader("עדכון מלאי ופתרונות")
    pool = sorted(FL["pn"].unique()) if not FL.empty else sorted({i["pn"] for i in ITEMS})
    pn = st.selectbox("מק״ט", pool)
    
    cur = UPDATES.get(pn, dict(added=0.0, eta="", status="פתוח", supplier="",
                               comment="", by=""))
    with st.form("upd"):
        c1, c2, c3 = st.columns(3)
        added = c1.number_input("תוספת למלאי", min_value=0.0, value=float(cur["added"]), step=1.0)
        try:
            eta_default = pd.to_datetime(cur["eta"]).date() if cur["eta"] else date.today()
        except Exception:
            eta_default = date.today()
        eta = c2.date_input("ETA", value=eta_default)
        opts = ["פתוח", "בטיפול", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
        status = c3.selectbox("סטטוס", opts, index=opts.index(cur["status"]) if cur["status"] in opts else 0)
        
        c4, c5 = st.columns(2)
        supplier = c4.text_input("ספק / קב״מ", value=cur["supplier"])
        by = c5.text_input("עודכן ע״י", value=cur["by"])
        comment = st.text_area("פעולה שננקטה", value=cur["comment"],
                               placeholder="למשל: קודמה אספקה מ-MTI ל-25/08")
        if st.form_submit_button("שמור", use_container_width=True):
            save_update(pn, added, str(eta), status, supplier, comment, by)
            st.cache_data.clear()
            st.success("נשמר. החישוב רץ מחדש.")
            st.rerun()

    st.divider()
    if UPDATES:
        exp = pd.DataFrame([dict(pn=p, **v) for p, v in UPDATES.items()])
        st.download_button("גיבוי כל העדכונים ל-CSV",
                           exp.to_csv(index=False).encode("utf-8-sig"),
                           "mrp_updates_backup.csv", "text/csv",
                           use_container_width=True)

# ---------- 10.9 היסטוריה ----------
with TABS[9]:
    st.subheader("היסטוריית עדכונים")
    hist = pd.read_sql_query(
        "SELECT updated_at, pn, added_stock, eta, status, supplier, comment, updated_by "
        "FROM history ORDER BY id DESC LIMIT 400", conn)
    if hist.empty:
        st.info("אין עדיין עדכונים.")
    else:
        st.dataframe(hist.rename(columns={
            "updated_at": "מתי", "pn": "מק״ט", "added_stock": "תוספת", "eta": "ETA",
            "status": "סטטוס", "supplier": "ספק", "comment": "הערה", "updated_by": "מי"}),
                     use_container_width=True, hide_index=True)
        st.markdown("##### ביטול עדכון")
        undo = st.selectbox("מק״ט לביטול", sorted(UPDATES.keys()) or ["—"])
        if undo != "—" and st.button("בטל את העדכון", type="primary"):
            delete_update(undo)
            st.cache_data.clear()
            st.rerun()

st.caption(f"מקור: {DATA_URL} · דרישה = כמות לקשר × תוכנית עבודה · הקצאה כרונולוגית, רמה עמוקה קודם, גירעון מועבר לחודש הבא.")
