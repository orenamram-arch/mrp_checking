# MRP Control Tower — מגדל בקרת חוסרים
# גרסה מתוקנת: תיקון באג תאריכים בתוכנית ההרכבה, תיקון באג NaN בהמרות
# מספריות, תיקון שליפת ETA בסיסי מה-MRP, ותיקון הצטברות מלאי לפי ETA.
# ראו הערות "תיקון קריטי" לאורך הקובץ.

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import io
import requests
import json
from supabase import create_client, Client

# ==========================================================
# HELPERS
# ==========================================================
def safe_num(value, default=0.0):
    # תיקון קריטי: pd.to_numeric(x, errors='coerce') מחזיר NaN לתאים לא-תקינים,
    # אבל NaN הוא "truthy" בפייתון, ולכן התבנית הישנה 'X or 0' לא באמת הופכת
    # NaN ל-0. safe_num מטפל בזה נכון עם pd.isna().
    n = pd.to_numeric(value, errors='coerce')
    if pd.isna(n):
        return default
    return float(n)

# ==========================================================
# CONFIGURATION & SYSTEM FACTORS
# ==========================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"

# תיקון: זהו ערך ברירת מחדל/גיבוי בלבד. הפקטורים האמיתיים נגזרים
# אוטומטית מתוך עמודת "QTY PER ASSY" בטבלת עץ ההרכבות שבקובץ עצמו
# (ראו ASSEMBLY_SYSTEM_FACTORS המחושב בהמשך, אחרי טעינת הנתונים) - כי
# התברר שהרשימה הידנית הזו הייתה חסרה לפחות פריט אחד (6930N127-001,
# שפקטור האמת שלו הוא 2 ולא 1 כברירת המחדל). אם טעינת הקובץ נכשלת
# מכל סיבה, המערכת תיפול חזרה לרשימה הידנית הזו.
ASSEMBLY_SYSTEM_FACTORS_FALLBACK = {
    "1096G860-002": 4,
    "1093U447-001": 4,
    "1093M635-003": 16,
    "1096B650-003": 16,
    "1096G880-003": 4
}
ASSEMBLY_SYSTEM_FACTORS = dict(ASSEMBLY_SYSTEM_FACTORS_FALLBACK)

st.set_page_config(
    page_title="MRP Executive Control Tower",
    page_icon="🚀",
    layout="wide"
)

# ==========================================================
# GLOBAL THEME / CSS
# ==========================================================
PRIMARY = "#4F46E5"
PRIMARY_DARK = "#3730A3"
ACCENT = "#06B6D4"
DANGER = "#EF4444"
WARNING = "#F59E0B"
SUCCESS = "#10B981"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Assistant:wght@400;600;700;800&display=swap');

[data-testid="stAppViewContainer"] .main .block-container,
[data-testid="stSidebarContent"] {{
    font-family: 'Assistant', sans-serif;
    direction: rtl;
}}
[data-testid="stAppViewContainer"] .main .block-container * ,
[data-testid="stSidebarContent"] * {{
    font-family: 'Assistant', sans-serif;
}}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

.hero-banner {{
    background: linear-gradient(120deg, {PRIMARY} 0%, {PRIMARY_DARK} 45%, {ACCENT} 100%);
    padding: 28px 32px;
    border-radius: 18px;
    margin-bottom: 22px;
    box-shadow: 0 10px 30px rgba(79,70,229,0.35);
}}
.hero-banner h1 {{
    color: white;
    font-weight: 800;
    font-size: 30px;
    margin: 0;
}}
.hero-banner p {{
    color: rgba(255,255,255,0.9);
    font-size: 15px;
    margin-top: 6px;
}}

.kpi-card {{
    background-color: var(--secondary-background-color);
    color: var(--text-color);
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 14px;
    padding: 18px 16px;
    text-align: center;
    box-shadow: 0 4px 14px rgba(0,0,0,0.1);
    transition: transform 0.15s ease;
}}
.kpi-card:hover {{ transform: translateY(-3px); }}
.kpi-label {{
    font-size: 13px;
    opacity: 0.75;
    margin-bottom: 6px;
    font-weight: 600;
}}
.kpi-value {{
    font-size: 30px;
    font-weight: 800;
}}
.kpi-sub {{
    font-size: 12px;
    opacity: 0.6;
    margin-top: 4px;
}}

.kpi-green {{ border-top: 4px solid {SUCCESS}; }}
.kpi-red {{ border-top: 4px solid {DANGER}; }}
.kpi-orange {{ border-top: 4px solid {WARNING}; }}
.kpi-blue {{ border-top: 4px solid {ACCENT}; }}

@media (prefers-color-scheme: light) {{
    .kpi-card, .kanban-card {{
        background-color: #ffffff !important;
        color: #111827 !important;
        border-color: #e5e7eb !important;
    }}
}}

@media (prefers-color-scheme: dark) {{
    .kpi-card, .kanban-card {{
        background-color: #1f2937 !important;
        color: #f9fafb !important;
        border-color: #374151 !important;
    }}
}}

.section-title {{
    font-weight: 800;
    font-size: 19px;
    margin: 18px 0 10px 0;
    border-right: 4px solid {PRIMARY};
    padding-right: 10px;
    color: var(--text-color, inherit);
}}

.kanban-col-header {{
    font-weight: 800;
    font-size: 15px;
    padding: 8px 12px;
    border-radius: 10px;
    margin-bottom: 8px;
    text-align: center;
}}
.kanban-card {{
    border: 1px solid rgba(128,128,128,0.2);
    border-radius: 10px;
    padding: 10px 12px;
    margin-bottom: 8px;
    border-right: 3px solid {PRIMARY};
    font-size: 13px;
}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
    <h1>🚀 MRP Executive Control Tower & Decision Hub</h1>
    <p>מערכת ניהול חוסרים מתקדמת, סימולציות קבלת החלטות (What-If), ותמונת מצב ניהולית מסונכרנת לענן במהירות שיא</p>
</div>
""", unsafe_allow_html=True)

def kpi_card(label, value, sub="", color="blue"):
    st.markdown(f"""
    <div class="kpi-card kpi-{color}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

try:
    _theme_base = st.get_option("theme.base")
except Exception:
    _theme_base = None

PLOTLY_TEMPLATE = "plotly_white" if _theme_base == "light" else "plotly_dark"
COLOR_SEQ = [PRIMARY, ACCENT, WARNING, DANGER, SUCCESS, "#A78BFA", "#F472B6", "#34D399"]

# ==========================================================
# SUPABASE SETUP & FAST CACHED STORAGE
# ==========================================================
SUPABASE_URL = "https://vobzhjutimeowgsjhgyt.supabase.co"
SUPABASE_KEY = "sb_publishable_OC3UKQ-UdO3ba4yHgvt9RQ_-AZdenBv"

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

@st.cache_data(ttl=60)
def fetch_all_inventory_records():
    try:
        response = supabase.table("mrp_inventory_updates").select("*").execute()
        records = {}
        if response.data:
            for row in response.data:
                pn = str(row.get("pn")).strip()
                eta_val = row.get("eta", "")
                if not eta_val or str(eta_val).strip() in ["", "None", "NaT", "nan"]:
                    eta_val = ""
                status_val = row.get("status", "פתוח") or "פתוח"
                records[pn] = {
                    "added_stock": float(row.get("added_stock", 0.0) or 0.0),
                    "eta": eta_val,
                    "status": status_val,
                    "supplier": row.get("supplier", "אופק"),
                    "comment": row.get("comment", ""),
                    "updated_by": row.get("updated_by", ""),
                    "updated_at": row.get("updated_at", "")
                }
        return records
    except Exception:
        return {}

def get_inventory_record(pn, cache=None):
    all_recs = cache if cache is not None else fetch_all_inventory_records()
    res = all_recs.get(str(pn).strip())
    if res:
        return (
            res["added_stock"],
            res["eta"],
            res["status"],
            res["supplier"],
            res["comment"],
            res["updated_by"],
            res["updated_at"]
        )
    return 0.0, "", "פתוח", "אופק", "", "", ""

@st.cache_data(ttl=60)
def fetch_wip_records():
    try:
        response = supabase.table("mrp_wip_assemblies").select("*").execute()
        if response.data:
            return {str(row.get("assembly_pn")).strip(): float(row.get("wip_qty", 0.0)) for row in response.data}
    except:
        pass
    return {}

def save_wip_record(assembly_pn, wip_qty):
    current_wip_dict = fetch_wip_records()
    existing_qty = current_wip_dict.get(str(assembly_pn), 0.0)
    total_new_qty = existing_qty + float(wip_qty)

    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {
        "assembly_pn": str(assembly_pn),
        "wip_qty": float(total_new_qty),
        "updated_at": now_str
    }
    try:
        supabase.table("mrp_wip_assemblies").upsert(payload, on_conflict="assembly_pn").execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירת WIP ל-Supabase: {e}")

def delete_wip_record(assembly_pn):
    try:
        supabase.table("mrp_wip_assemblies").delete().eq("assembly_pn", str(assembly_pn)).execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקת WIP מ-Supabase: {e}")

def save_inventory_record(pn, added_stock, eta, status, supplier, comment, updated_by, webhook_url=""):
    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {
        "pn": str(pn),
        "added_stock": float(added_stock),
        "eta": str(eta),
        "status": str(status),
        "supplier": str(supplier),
        "comment": str(comment),
        "updated_by": str(updated_by),
        "updated_at": now_str
    }
    try:
        supabase.table("mrp_inventory_updates").upsert(payload, on_conflict="pn").execute()
        supabase.table("mrp_inventory_history").insert(payload).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירה ל-Supabase: {e}")

    if webhook_url:
        msg = "🔔 עדכון מלאי/ETA למוצר!\nמק'ט: " + str(pn) + "\nתוספת מלאי: " + str(added_stock) + "\nסטטוס: " + str(status) + "\nETA: " + str(eta)
        try:
            requests.post(webhook_url, data=json.dumps({"text": msg}), headers={'Content-Type': 'application/json'})
        except:
            pass

def delete_inventory_record(pn):
    try:
        supabase.table("mrp_inventory_updates").delete().eq("pn", str(pn)).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקה מ-Supabase: {e}")

@st.cache_data(ttl=60)
def fetch_cloud_assembly_plan():
    try:
        response = supabase.table("mrp_assembly_plans").select("*").execute()
        if response.data:
            return pd.DataFrame(response.data)
    except:
        pass
    return pd.DataFrame()

def save_cloud_assembly_plan(plan_df):
    try:
        records = plan_df.to_dict(orient="records")
        supabase.table("mrp_assembly_plans").delete().neq("Assembly_PN", "DUMMY").execute()
        if records:
            supabase.table("mrp_assembly_plans").upsert(records).execute()
        fetch_cloud_assembly_plan.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירת תוכנית הייצור לענן: {e}")

# ==========================================================
# DATA LOADING FROM GITHUB & SESSION STATE
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

PN_COL = df.columns[1]
DESC_COL = df.columns[4]
ITEM_TYPE_COL = df.columns[44] if len(df.columns) > 44 else df.columns[-1]
STOCK_COL = df.columns[79] if len(df.columns) > 79 else df.columns[-1]
ASSEMBLY_COLS = df.columns[10:36].tolist()
MONTH_COLS = df.columns[108:132].tolist() if len(df.columns) > 132 else []

# ==========================================================
# שיפור: בדיקת תקינות אוטומטית למבנה כותרות התאריכים
# ==========================================================
# מגלה אוטומטית אם קובץ ה-Excise (מה-GitHub) חזר על אותו באג שתיקנו -
# חודשים כפולים בכותרת - כך שאם הקובץ ישתנה בעתיד, המשתמש יקבל התרעה
# ברורה בממשק במקום חישוב שקט ושגוי.
def _validate_month_headers(month_cols):
    warnings_list = []
    ym_list = []
    for c in month_cols:
        if pd.notnull(c):
            try:
                ym_list.append(pd.to_datetime(c).strftime("%Y-%m"))
            except Exception:
                pass
    if len(ym_list) != len(set(ym_list)):
        warnings_list.append(
            "⚠️ נמצאו חודשים כפולים בכותרת עמודות ה-MRP הראשיות (108-131). "
            "מבנה קובץ ה-Excel כנראה השתנה מאז התיקון האחרון - יש לבדוק ידנית "
            "לפני שסומכים על תוצאות המערכת."
        )
    if len(ym_list) < 24:
        warnings_list.append(
            f"⚠️ נמצאו רק {len(ym_list)} חודשים תקינים בכותרת (במקום 24 צפויים). "
            "ייתכן שהוספו/הוסרו עמודות בקובץ."
        )
    return warnings_list

_month_header_warnings = _validate_month_headers(MONTH_COLS)
for _w in _month_header_warnings:
    st.warning(_w)

valid_assemblies = []
for col in ASSEMBLY_COLS:
    valid_assemblies.append(col)

assembly_levels = {}
for col in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(col)
        lvl_val = int(df_levels.iloc[0, col_idx])
        assembly_levels[col] = lvl_val
    except:
        assembly_levels[col] = 0

valid_assemblies = sorted(valid_assemblies, key=lambda x: (assembly_levels.get(x, 0), str(x)))

# ==========================================================
# תיקון + שיפור: פענוח עץ ההרכבות האמיתי מהקובץ (BOM היררכי)
# ==========================================================
# בטבלה בשורות 3-28 (עמודות DESC/LEVEL/PN/QTY PER ASSY, עמודות 104-107)
# יש עץ הרכבות מלא: כל שורה = הרכבה אחת, עם רמת עומק (LEVEL) וכמות
# ליחידת ההורה הישיר שלה (QTY PER ASSY). זהו "BOM מוזח" קלאסי - ההורה
# של כל שורה הוא השורה הקרובה שלפניה עם LEVEL נמוך ב-1.
#
# תיקון: הקוד המקורי לא השתמש בטבלה הזו כלל להיררכיה, ובמקום זה
# הסתמך על מילון קשיח (ASSEMBLY_SYSTEM_FACTORS) עם 5 פריטים בלבד -
# שהתברר שחסר בו לפחות פריט אחד עם פקטור שונה מ-1 (6930N127-001,
# פקטור אמיתי=2). כאן אנחנו גוזרים את הפקטורים ישירות מהקובץ, לכל
# 26 ההרכבות, ובנוסף בונים עץ הורה-ילד מלא לבדיקת זמינות היררכית ב-WIP.
ASSEMBLY_BOM_TREE = {}     # pn -> {"desc","level","qty_per_parent","parent"}
ASSEMBLY_CHILDREN = {}     # parent_pn -> [child_pn, ...]

try:
    _level_stack = []  # [(level, pn), ...] מהשורש ועד הענף הנוכחי
    for _r in range(3, 29):
        _desc = df_raw.iloc[_r, 104]
        _level = df_raw.iloc[_r, 105]
        _pn = df_raw.iloc[_r, 106]
        _qty = df_raw.iloc[_r, 107]
        if pd.isna(_pn):
            continue
        _pn = str(_pn).strip()
        _level = int(_level) if pd.notnull(_level) else 0
        _qty = safe_num(_qty, default=1.0)

        while _level_stack and _level_stack[-1][0] >= _level:
            _level_stack.pop()
        _parent_pn = _level_stack[-1][1] if _level_stack else None

        ASSEMBLY_BOM_TREE[_pn] = {
            "desc": str(_desc), "level": _level, "qty_per_parent": _qty, "parent": _parent_pn
        }
        if _parent_pn:
            ASSEMBLY_CHILDREN.setdefault(_parent_pn, []).append(_pn)

        _level_stack.append((_level, _pn))
except Exception:
    ASSEMBLY_BOM_TREE = {}
    ASSEMBLY_CHILDREN = {}

if ASSEMBLY_BOM_TREE:
    ASSEMBLY_SYSTEM_FACTORS = {
        pn: info["qty_per_parent"] for pn, info in ASSEMBLY_BOM_TREE.items() if info["qty_per_parent"] != 1
    }
    if not ASSEMBLY_SYSTEM_FACTORS:
        ASSEMBLY_SYSTEM_FACTORS = dict(ASSEMBLY_SYSTEM_FACTORS_FALLBACK)
else:
    ASSEMBLY_SYSTEM_FACTORS = dict(ASSEMBLY_SYSTEM_FACTORS_FALLBACK)

if "custom_assembly_plan_df" not in st.session_state:
    cloud_plan = fetch_cloud_assembly_plan()
    if not cloud_plan.empty:
        st.session_state["custom_assembly_plan_df"] = cloud_plan
    else:
        # ==========================================================
        # תיקון קריטי: באג תאריכים בכותרת טבלת תוכנית ההרכבה
        # ==========================================================
        # בקובץ המקור, שורת הכותרת של טבלת תוכנית ההרכבה (df_raw שורה 2,
        # עמודות 108 ואילך) "שבורה": השנה לא מתקדמת כמו שצריך, ובמקום זה
        # יש שני בלוקים של 12 חודשים (ינואר-דצמבר) שמסומנים שניהם "2026"
        # (רק היום בחודש שונה - 26 מול 27). בפועל, הבלוק השני הוא נתוני
        # השנה הבאה (2027), לא חזרה על 2026. הקוד המקורי קרא את התאריכים
        # משורה 2 ולכן שיוך כל ה-Build_Qty של השנה השנייה יוחס בטעות
        # לחודשים המקבילים בשנה הראשונה (למשל ייצור מתוכנן ל-2027-01 נספר
        # כאילו הוא ב-2026-01) - מה שמערבב תוכניות ייצור בין שנים ומעוות
        # לגמרי את חישובי החוסרים לטווח הארוך.
        #
        # שורת הכותרת של הטבלה הראשית (df, header=29) לעומת זאת תקינה
        # לחלוטין באותו טווח עמודות בדיוק (108-131: ינואר 2026 עד דצמבר
        # 2027 ברצף, ללא כפילות). מכיוון שהעמודות מיושרות 1:1 בין שתי
        # הטבלאות, אנחנו משתמשים בתאריכי הכותרת התקינים של הטבלה הראשית
        # (MONTH_COLS) גם עבור טבלת תוכנית ההרכבה, במקום בתאריכים השבורים
        # משורה 2.
        if len(MONTH_COLS) >= 24:
            header_dates = list(MONTH_COLS[:24])
        else:
            header_dates = df_raw.iloc[2, 108:132].values if df_raw.shape[1] > 132 else []
        plan_rows = []

        for r in range(3, df_raw.shape[0]):
            asm_pn = df_raw.iloc[r, 106] if df_raw.shape[1] > 106 else None
            if pd.notnull(asm_pn):
                clean_asm_pn = str(asm_pn).strip()
                system_multiplier = ASSEMBLY_SYSTEM_FACTORS.get(clean_asm_pn, 1)

                for c_idx, date_val in enumerate(header_dates):
                    if pd.notnull(date_val):
                        qty = df_raw.iloc[r, 108 + c_idx]
                        if pd.notnull(qty) and qty != '' and qty != 'NaN':
                            try:
                                q_val = float(qty)
                                if q_val > 0:
                                    dt = pd.to_datetime(date_val)
                                    ym_str = dt.strftime("%Y-%m")
                                    if dt.month >= 9 or ym_str >= "2026-09":
                                        displayed_build_qty = q_val * system_multiplier
                                        plan_rows.append({
                                            "Assembly_PN": clean_asm_pn,
                                            "YearMonth": ym_str,
                                            "Build_Qty": displayed_build_qty,
                                            "Raw_Build_Qty": q_val
                                        })
                            except:
                                pass
        st.session_state["custom_assembly_plan_df"] = pd.DataFrame(plan_rows)

assembly_plan_df = st.session_state["custom_assembly_plan_df"]

# ==========================================================
# תיקון קריטי: מיפוי תאריכים אמיתי לעמודות לוח האספקה הפריטני
# ==========================================================
# הקוד המקורי בנה raw_eta_dates משורה 2 בלבד (df_raw.iloc[2, :]).
# הבעיה: שורה 2 ריקה (NaN) בדיוק בטווח העמודות 80-103, שבו נמצא לוח
# האספקה הצפויה בפועל לכל פריט (כמויות הזמנות פתוחות לפי תאריך יעד).
# כתוצאה מכך הפונקציה דילגה על הבלוק האמיתי הזה לגמרי, וקפצה ישר
# לעמודות 108 ואילך (עמודות היתרה החודשית הנטו של ה-MRP), שהן דבר שונה
# לחלוטין מ"תאריך ההגעה של אספקה מתוכננת" - ולכן ה-ETA שהוצג בכרטיסי
# הפריטים לא שיקף אספקה אמיתית אלא את החודש הראשון שבו היתרה הנטו
# חיובית (מה שכבר תלוי בעצמו בכל שרשרת החישוב).
#
# בנוסף, שורת הכותרת עצמה (שורה 29) בטווח 80-103 סובלת מאותו באג שנה
# שתיקנו למעלה בתוכנית ההרכבה: 12 העמודות הראשונות מתויגות "2026" (יום
# 26) ו-12 הבאות מתויגות שוב "2026" (יום 27) במקום "2027". אנחנו מתקנים
# את זה כאן באותו אופן - הבלוק השני מקודם בשנה אחת.
def _build_supply_date_map():
    date_map = {}
    # בלוק אספקה פריטני: עמודות 80-103 (24 עמודות, מקור: שורת כותרת 29)
    supply_start = 80
    for i in range(24):
        col_pos = supply_start + i
        if col_pos >= df_raw.shape[1]:
            break
        year_offset = 0 if i < 12 else 1
        month = (i % 12) + 1
        base_year = 2026  # שנת הבסיס של תוכנית זו, כפי שמופיעה בקובץ
        try:
            date_map[col_pos] = pd.Timestamp(year=base_year + year_offset, month=month, day=1)
        except Exception:
            pass
    # בלוק יתרת MRP חודשית: עמודות 108-131 - התאריכים כאן כבר תקינים
    # ורציפים בשורה הראשית (df, header=29), אז פשוט קוראים אותם משם.
    for i, col in enumerate(MONTH_COLS[:24]):
        col_pos = 108 + i
        if pd.notnull(col):
            try:
                date_map[col_pos] = pd.to_datetime(col)
            except Exception:
                pass
    return date_map

SUPPLY_DATE_MAP = _build_supply_date_map()

def get_base_mrp_eta_and_qty(pn):
    # מחזיר (YearMonth, כמות) של האספקה/היתרה הראשונה החיובית עבור מק"ט,
    # לפי לוח האספקה הפריטני (עמודות 80-103) קודם, ואם לא נמצא - לפי
    # היתרה החודשית הנטו של ה-MRP (עמודות 108-131) כגיבוי.
    matching_rows = df_raw[df_raw.iloc[:, 1].astype(str).str.strip() == str(pn).strip()]
    if matching_rows.empty:
        return "בדיקה נדרשת", 0.0

    row_idx = matching_rows.index[0]
    for col_pos in sorted(SUPPLY_DATE_MAP.keys()):
        try:
            val = df_raw.iloc[row_idx, col_pos]
            q = safe_num(val)
            if q > 0:
                dt = SUPPLY_DATE_MAP[col_pos]
                return dt.strftime("%Y-%m"), q
        except Exception:
            pass
    return "בדיקה נדרשת", 0.0

def get_base_mrp_eta(pn):
    eta, _ = get_base_mrp_eta_and_qty(pn)
    return eta

def get_base_mrp_qty(pn):
    _, qty = get_base_mrp_eta_and_qty(pn)
    return qty

# ==========================================================
# תיקון קריטי: זמינות עתידית לפי ETA - לא רק מלאי נוכחי בקופה
# ==========================================================
# זה בדיוק המנגנון הבסיסי של MRP שציינת: פריט לא חייב להיות כבר
# פיזית במלאי כדי שהתוכנית תיחשב אפשרית - מספיק שה-ETA שלו חל **לפני**
# חודש הבנייה המתוכנן (חודש קודם, לא אותו חודש עצמו - כי אין ודאות
# שהפריט יגיע *לפני* שהבנייה בפועל מתחילה בתוך חודש היעד עצמו). לכן,
# זמינות של רכיב לחודש יעד נתון היא: מלאי נוכחי (STOCK) + כל האספקה
# הצפויה מלוח ה-PO הפריטני (עמודות 80-103) שה-ETA שלה קודם לחודש
# היעד (לא כולל אותו חודש) + תוספת מלאי ידנית שנרשמה עם ETA שכבר חל
# בחודש קודם לחודש היעד.
#
# תיקון (בעקבות משוב נוסף): בגרסה הקודמת השתמשתי ב-"<=" (כולל את אותו
# חודש) - זו הייתה טעות. הכלל הנכון, כפי שהוגדר: "במלאי, או שה-ETA
# הוא חודש לפני התוכנית" - כלומר "<" (קודם, לא כולל).
def get_cumulative_incoming_supply(pn, target_ym):
    matching_rows = df_raw[df_raw.iloc[:, 1].astype(str).str.strip() == str(pn).strip()]
    if matching_rows.empty:
        return 0.0
    row_idx = matching_rows.index[0]
    total = 0.0
    for col_pos, dt in SUPPLY_DATE_MAP.items():
        if col_pos > 103:
            continue  # רק בלוק לוח האספקה הפריטני (80-103), לא עמודות היתרה החודשית
        try:
            ym = dt.strftime("%Y-%m")
            if ym < target_ym:
                q = safe_num(df_raw.iloc[row_idx, col_pos])
                if q > 0:
                    total += q
        except Exception:
            pass
    return total

def get_component_available_by_month(pn, target_ym, inv_cache=None):
    if inv_cache is None:
        inv_cache = fetch_all_inventory_records()
    match = df[df[PN_COL].astype(str).str.strip() == pn]
    base_stock = safe_num(match.iloc[0][STOCK_COL]) if not match.empty else 0.0

    saved_add, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    manual_eta_ym = None
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        try:
            manual_eta_ym = pd.to_datetime(manual_eta).strftime("%Y-%m")
        except Exception:
            manual_eta_ym = None
    manual_stock_effective = saved_add if (manual_eta_ym is None or manual_eta_ym < target_ym) else 0.0

    incoming_supply = get_cumulative_incoming_supply(pn, target_ym)
    return base_stock + manual_stock_effective + incoming_supply

def get_first_supply_eta(pn, inv_cache=None):
    _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        return manual_eta
    return get_base_mrp_eta(pn)

# ==========================================================
# SIDEBAR FILTERS & WHAT-IF CONTROLS
# ==========================================================
st.sidebar.header("⚙️ הגדרות מערכת וחיבור")
webhook_url = st.sidebar.text_input("🔗 Teams / Slack Webhook URL (אופציונלי)", value="")
supplier_options = ["אופק", "ספק פנימי", "רכש אחר", "אחר"]

st.sidebar.header("🔍 מסננים מתקדמים")

if st.sidebar.button("🧹 איפוס כל המסננים (Clear All)"):
    keys_to_clear = ["selected_month_label", "num_months_ahead", "selected_level", "selected_assembly", "selected_item_type", "selected_search_item"]
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

month_options = {}
for m in MONTH_COLS:
    if pd.notnull(m):
        try:
            dt = pd.to_datetime(m)
            m_ym = dt.strftime("%Y-%m")
            if m_ym >= "2026-09":
                month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
        except:
            pass

if not month_options:
    for m in MONTH_COLS:
        if pd.notnull(m):
            try:
                dt = pd.to_datetime(m)
                if dt.month >= 9 or dt.strftime("%Y-%m") >= "2026-09":
                    month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
            except:
                pass

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()), key="selected_month_label")
selected_month_col = month_options[selected_month_label]

try:
    selected_ym = pd.to_datetime(selected_month_col).strftime("%Y-%m")
except:
    selected_ym = str(selected_month_col)[:7]

num_months_ahead = st.sidebar.slider("📅 טווח מבט קדימה במספר חודשים", min_value=1, max_value=6, value=1, key="num_months_ahead")

level_options = ["הכל"] + sorted(list(set(str(assembly_levels[c]) for c in valid_assemblies)), key=lambda x: int(x) if x.isdigit() else 0)
selected_level = st.sidebar.selectbox("סינון לפי רמת עץ (BOM Level)", level_options, key="selected_level")

assembly_mapping = {"הכל": "הכל"}
filtered_assembly_cols = []
for col in valid_assemblies:
    try:
        col_idx = df.columns.get_loc(col)
        lvl = str(assembly_levels.get(col, 0))
        desc = df_desc.iloc[0, col_idx]
        if selected_level == "הכל" or lvl == selected_level:
            filtered_assembly_cols.append(col)
            assembly_mapping[col] = f"[רמה {lvl}] {str(col)} - {str(desc)}"
    except:
        filtered_assembly_cols.append(col)
        assembly_mapping[col] = col

selected_assembly = st.sidebar.selectbox(
    "בחר הרכבה ספציפית לדשבורד",
    ["הכל"] + filtered_assembly_cols,
    format_func=lambda x: assembly_mapping.get(x, x),
    key="selected_assembly"
)

item_types = df[ITEM_TYPE_COL].dropna().unique().tolist() if ITEM_TYPE_COL in df.columns else []
selected_item_type = st.sidebar.selectbox("בחר סוג פריט (עמודה AS)", ["הכל"] + item_types, key="selected_item_type")

item_choices = ["הכל"] + sorted([f"{str(r[PN_COL]).strip()} - {str(r[DESC_COL])}" for _, r in df.iterrows() if pd.notnull(r[PN_COL])])
selected_search_item = st.sidebar.selectbox("🔎 חיפוש מהיר (בחר או הקלד מק'ט/תיאור)", item_choices, key="selected_search_item")
search_pn = selected_search_item.split(" - ")[0] if selected_search_item != "הכל" else "הכל"

# ==========================================================
# FILE UPLOADS & TEMPLATES
# ==========================================================
st.sidebar.divider()
st.sidebar.markdown("##### 📥 עדכון ETA וכמות אספקה מקובץ ספק")

eta_template_df = pd.DataFrame(columns=["PN", "ETA", "Qty"])
eta_template_output = io.BytesIO()
with pd.ExcelWriter(eta_template_output, engine='openpyxl') as writer:
    eta_template_df.to_excel(writer, index=False, sheet_name='ETA_Template')
st.sidebar.download_button(
    label="📄 הורד תבנית Excel לעדכון ETA",
    data=eta_template_output.getvalue(),
    file_name="ETA_Update_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded_eta_file = st.sidebar.file_uploader("העלה קובץ ETA (עמודות: PN, ETA, Qty)", type=["xlsx", "xls"], key="eta_uploader")
if uploaded_eta_file is not None:
    try:
        eta_df_sup = pd.read_excel(uploaded_eta_file)
        if st.sidebar.button("⚡ עדכן ETA וכמות אספקה"):
            eta_count = 0
            for _, s_row in eta_df_sup.iterrows():
                p_code = str(s_row.iloc[0]).strip()
                new_eta = str(s_row.iloc[1]).strip() if len(s_row) > 1 and pd.notnull(s_row.iloc[1]) else ""
                new_supply_qty = float(s_row.iloc[2]) if len(s_row) > 2 and pd.notnull(s_row.iloc[2]) else 0.0

                if p_code and p_code != 'nan' and new_eta and new_eta not in ["nan", "NaT", "None"]:
                    curr_stock, _, curr_status, curr_sup, curr_comm, _, _ = get_inventory_record(p_code)
                    updated_total_stock = curr_stock + new_supply_qty if new_supply_qty > 0 else curr_stock

                    save_inventory_record(
                        pn=p_code,
                        added_stock=updated_total_stock,
                        eta=new_eta,
                        status=curr_status if curr_status != "פתוח" else "הוזמן",
                        supplier=curr_sup,
                        comment=f"{curr_comm} | אספקה בסך {new_supply_qty} בתאריך ETA {new_eta} מקובץ ספק",
                        updated_by="ETA & Qty File Upload",
                        webhook_url=webhook_url
                    )
                    eta_count += 1
            st.sidebar.success(f"עודכנו בהצלחה ETA וכמויות אספקה עבור {eta_count} שורות!")
    except Exception as e:
        st.sidebar.error(f"שגיאה בקריאת קובץ ה-ETA: {e}")

# ==========================================================
# OPTIMIZED SHORTAGE CALCULATION (CACHED & FIXED)
# ==========================================================
all_ym_list = sorted(list(set(assembly_plan_df["YearMonth"].unique())))
start_idx = 0
for idx, ym in enumerate(all_ym_list):
    if ym >= selected_ym:
        start_idx = idx
        break
selected_target_yms = all_ym_list[start_idx:start_idx + num_months_ahead]
if not selected_target_yms:
    selected_target_yms = [selected_ym]

@st.cache_data(ttl=60)
def calculate_mrp_breakdown_cached(target_yms_tuple, sim_extra_stock_items_tuple, active_plan_df):
    sim_extra_stock_dict = dict(sim_extra_stock_items_tuple)
    inv_cache = fetch_all_inventory_records()
    wip_cache = fetch_wip_records()

    target_month_cols_map = {}
    for m_c in MONTH_COLS:
        if pd.notnull(m_c):
            try:
                m_dt_ym = pd.to_datetime(m_c).strftime("%Y-%m")
                if m_dt_ym in target_yms_tuple:
                    target_month_cols_map[m_dt_ym] = m_c
            except:
                pass

    temp_df = df.copy()
    shortage_records = {}

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        saved_stock_add, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        sim_val = sim_extra_stock_dict.get(pn, 0.0)

        # ==========================================================
        # תיקון קריטי: מלאי מצטבר שמכבד את חודש ה-ETA שלו
        # ==========================================================
        # הקוד המקורי החיל את כל תוספת המלאי (total_added_stock) על כל
        # חודש שנבדק, כולל חודשים שקודמים לתאריך ההגעה (ETA) שהמשתמש
        # עצמו רשם לאותה תוספת - כאילו המלאי כבר זמין גם *לפני* שהוא
        # בכלל הגיע. בהמשך הפונקציה היה גם בלוק "תיקון" שניסה לנטרל את
        # זה עבור ETA עתידי מדי, אבל הוא בפועל לא עשה כלום (no-op: הוא
        # החיל abs() על ערך שכבר שלילי, בלי לשנות אותו).
        #
        # התיקון האמיתי: תוספת מלאי שנשמרה עם ETA מוגדר משפיעה רק
        # החל מהחודש שאחרי חודש ה-ETA שלה (זה בדיוק ה"מצטבר קדימה"
        # שביקשת) - היא לא "נעלמת" בחודשים שאחרי, אבל גם לא מוחלת לא
        # בחודש ה-ETA עצמו ולא לפניו, כי אין ודאות שהיא מגיעה *לפני*
        # שהבנייה בפועל מתחילה בתוך אותו חודש (בדיוק כפי שהובהר: "או
        # שהמלאי במלאי, או שה-ETA שלו הוא חודש לפני התוכנית"). תוספת
        # בלי ETA (או תוספת מסימולציית What-If) ממשיכה להיחשב זמינה
        # מיידית, כברירת מחדל.
        manual_eta_ym = None
        if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
            try:
                manual_eta_ym = pd.to_datetime(manual_eta).strftime("%Y-%m")
            except Exception:
                manual_eta_ym = None

        max_shortage_val = 0.0
        is_short_or = False

        for ym in target_yms_tuple:
            col_name = target_month_cols_map.get(ym)
            if col_name and col_name in temp_df.columns:
                mrp_val = safe_num(row[col_name])

                stock_arrived_by_this_month = (manual_eta_ym is None) or (manual_eta_ym < ym)
                effective_addition = (saved_stock_add if stock_arrived_by_this_month else 0.0) + sim_val

                effective_mrp_val = mrp_val + effective_addition if mrp_val < 0 else mrp_val

                if effective_mrp_val < 0:
                    is_short_or = True
                    sh_qty = abs(effective_mrp_val)
                    if sh_qty > max_shortage_val:
                        max_shortage_val = sh_qty

        if is_short_or:
            shortage_records[idx] = max_shortage_val

    temp_df['Monthly_Balance'] = temp_df.index.map(lambda i: -shortage_records[i] if i in shortage_records else 1.0)

    mrp_shortages = temp_df[temp_df['Monthly_Balance'] < 0].copy()
    mrp_shortages['Total_MRP_Shortage'] = mrp_shortages['Monthly_Balance'].abs()

    month_plan = active_plan_df[active_plan_df["YearMonth"].isin(target_yms_tuple)]
    # תיקון קריטי: חישוב דרישות MRP מבוסס אך ורק על Raw_Build_Qty הגולמי משום שהעץ בקובץ כבר כולל את הפקטורים
    plan_dict = month_plan.groupby("Assembly_PN")["Raw_Build_Qty"].sum().to_dict()

    for asm_wip, wip_qty in wip_cache.items():
        if wip_qty > 0 and asm_wip in plan_dict:
            sys_factor = ASSEMBLY_SYSTEM_FACTORS.get(asm_wip, 1)
            raw_wip_qty = wip_qty / sys_factor
            plan_dict[asm_wip] = max(0.0, plan_dict[asm_wip] - raw_wip_qty)

    breakdown_rows = []
    # תיקון קריטי: עמודת "מלאי" בטבלת החוסרים תשקף עכשיו את אותה
    # זמינות מתחשבת-ETA כמו כל שאר המערכת (מלאי + אספקה שה-ETA שלה חל
    # עד סוף טווח החודשים הנבדק) - כדי שלא יהיה פער בין מה שמוצג כאן
    # לבין מה שקובע בפועל אם הרכבה "ניתנת לייצור" בטאבים אחרים.
    reference_ym = max(target_yms_tuple) if target_yms_tuple else None

    for idx, row in mrp_shortages.iterrows():
        pn = str(row[PN_COL]).strip()
        desc = str(row[DESC_COL])
        item_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in temp_df.columns else ""

        if reference_ym:
            stock = get_component_available_by_month(pn, reference_ym, inv_cache) + sim_extra_stock_dict.get(pn, 0.0)
        else:
            base_stock = safe_num(row[STOCK_COL])
            saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
            stock = base_stock + saved_stock_add + sim_extra_stock_dict.get(pn, 0.0)

        total_mrp_shortage = row['Total_MRP_Shortage']
        _, _, item_status, current_sup, _, _, _ = get_inventory_record(pn, inv_cache)

        mouser_link = f"https://www.mouser.co.il/c/?q={pn}"
        digikey_link = f"https://www.digikey.com/en/products/result?keywords={pn}"
        findchips_link = f"https://www.findchips.com/search/{pn}"

        added_for_this_pn = False
        for asm in filtered_assembly_cols:
            qty_per_asm = safe_num(row[asm])
            if qty_per_asm > 0:
                added_for_this_pn = True
                asm_raw_build = plan_dict.get(asm, 0.0)
                required_demand = qty_per_asm * asm_raw_build
                asm_desc = assembly_mapping.get(asm, asm)

                breakdown_rows.append({
                    "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                    "Status": item_status, "Assembly": asm, "Assembly_Desc": asm_desc, "Qty_Per_Assembly": qty_per_asm,
                    "Assembly_Monthly_Build": asm_raw_build * ASSEMBLY_SYSTEM_FACTORS.get(asm, 1),
                    "Required_Demand": required_demand,
                    "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage,
                    "חיפוש במאוזר": mouser_link, "חיפוש בדיגיקי": digikey_link, "חיפוש ב-Findchips": findchips_link
                })

        if not added_for_this_pn:
            breakdown_rows.append({
                "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                "Status": item_status, "Assembly": "ללא שיוך", "Assembly_Desc": "ללא שיוך להרכבה", "Qty_Per_Assembly": 0,
                "Assembly_Monthly_Build": 0, "Required_Demand": 0, "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage,
                "חיפוש במאוזר": mouser_link, "חיפוש בדיגיקי": digikey_link, "חיפוש ב-Findchips": findchips_link
            })

    res_df = pd.DataFrame(breakdown_rows)
    return res_df

def calculate_mrp_breakdown(sim_extra_stock=None, target_yms=None, plan_df_override=None):
    if sim_extra_stock is None:
        sim_extra_stock = {}
    if target_yms is None:
        target_yms = selected_target_yms
    active_plan = plan_df_override if plan_df_override is not None else assembly_plan_df
    
    res = calculate_mrp_breakdown_cached(tuple(target_yms), tuple(sorted(sim_extra_stock.items())), active_plan)
    res_df = res.copy()

    if not res_df.empty:
        if selected_item_type != "הכל":
            res_df = res_df[res_df["Item_Type"] == selected_item_type]
        if selected_assembly != "הכל":
            res_df = res_df[res_df["Assembly"] == selected_assembly]
        if search_pn != "הכל":
            res_df = res_df[res_df["PN"] == search_pn]

    return res_df

breakdown_df = calculate_mrp_breakdown(target_yms=selected_target_yms)

# ==========================================================
# תיקון קריטי: בדיקת זמינות היררכית אמיתית לפני הוספה ל-WIP
# ==========================================================
# הכפתור בטאב ה-WIP הבטיח "בדיקת זמינות היררכית מלאה", אבל בפועל לא
# הייתה שום בדיקה - היה אפשר להוסיף ל-WIP כל כמות של כל הרכבה (כולל
# ההרכבה הסופית, רמה 0) גם אם אין שום סיכוי לבנות אותה, כי תתי-ההרכבות
# שלה חסרות ברכיבי גלם. הפונקציה הזו הולכת רקורסיבית על עץ ה-BOM
# (ASSEMBLY_BOM_TREE / ASSEMBLY_CHILDREN שנבנה מהקובץ) ובודקת בכל
# רמה: (א) האם יש מספיק רכיבי גלם לכמות המבוקשת של ההרכבה הזו עד
# (כולל) חודש היעד - כלומר מלאי נוכחי + כל אספקה שה-ETA שלה כבר חל,
# בדיוק כמו במנגנון ה-MRP הרגיל (לא רק "יש כרגע בקופה") - ו-(ב) לכל
# תת-הרכבה, האם ה-WIP הקיים שלה מכסה את הכמות הדרושה, ואם לא, ממשיכה
# לבדוק את רכיבי הגלם שלה (רקורסיבית, עד לעלים).
#
# תיקון נוסף (בעקבות משוב): הגרסה הקודמת השוותה מול מלאי נוכחי בלבד
# ("STOCK") והתעלמה לגמרי מ-ETA - זו הייתה טעות, כי בדיוק ככה עובד
# ה-MRP: רכיב לא חייב להיות כבר במלאי, מספיק שה-ETA שלו חל עד (כולל)
# חודש הבנייה המתוכנן.
def check_hierarchical_ctb(asm_pn, requested_qty, target_ym, inv_cache=None, wip_cache=None, _visited=None):
    if inv_cache is None:
        inv_cache = fetch_all_inventory_records()
    if wip_cache is None:
        wip_cache = fetch_wip_records()
    if _visited is None:
        _visited = set()
    if asm_pn in _visited or requested_qty <= 0:
        return []
    _visited.add(asm_pn)

    blockers = []

    # (א) רכיבי גלם ישירים של ההרכבה הזו (עמודת ההרכבה ב-df הראשי)
    if asm_pn in df.columns:
        for _, row in df.iterrows():
            qty_per = safe_num(row[asm_pn])
            if qty_per <= 0:
                continue
            comp_pn = str(row[PN_COL]).strip()
            base_stock_check = safe_num(row[STOCK_COL])
            if base_stock_check >= 9000000:
                # ערך "מלאי אינסופי" - זהו כנראה מק"ט של הרכבה אחרת
                # שמטופלת כבר דרך העץ הרקורסיבי, לא רכיב גלם אמיתי.
                continue
            required = qty_per * requested_qty
            available = get_component_available_by_month(comp_pn, target_ym, inv_cache)
            if available < required:
                blockers.append({
                    "assembly": asm_pn,
                    "assembly_desc": assembly_mapping.get(asm_pn, asm_pn),
                    "component": comp_pn,
                    "component_desc": str(row[DESC_COL]),
                    "required": required, "available": available,
                    "shortage": required - available
                })

    # (ב) תתי-הרכבות (רקורסיה) - מנוכה מהן ה-WIP הקיים
    for child_pn in ASSEMBLY_CHILDREN.get(asm_pn, []):
        child_info = ASSEMBLY_BOM_TREE.get(child_pn, {})
        qty_per_parent = child_info.get("qty_per_parent", 1.0)
        child_needed = requested_qty * qty_per_parent
        child_wip = wip_cache.get(child_pn, 0.0)
        net_needed = max(0.0, child_needed - child_wip)
        if net_needed > 0:
            blockers.extend(check_hierarchical_ctb(child_pn, net_needed, target_ym, inv_cache, wip_cache, _visited))

    return blockers

# ==========================================================
# TABS DEFINITION (10 הטאבים המקוריים + טאב חדש לעריכת ETA מרוכזת)
# ==========================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11 = st.tabs([
    "📈 Executive Dashboard",
    "📊 תוכנית ייצור (Smart CTB)",
    "💡 סימולציית What-If",
    "📌 לוח סטטוסים (Kanban)",
    "🏭 ניהול WIP (בייצור)",
    "📅 עדכון מלאי וספקים",
    "📅 מעקב ETA ודחיות",
    "↩️ ניהול UNDO",
    "📦 ניהול מלאי מעודכן",
    "🎯 ניתוח רגישות ותוכנית",
    "✏️ עריכת ETA מרוכזת"
])

with tab1:
    israel_time = datetime.utcnow() + timedelta(hours=3)
    current_time_str = israel_time.strftime("%d/%m/%Y | %H:%M:%S")
    st.markdown(f"""
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; opacity: 0.85; font-weight: 600;">
        <div>🎯 תמונת מצב ניהולית לטווח חודשים: {', '.join(selected_target_yms)}</div>
        <div>🕒 שעון ישראל (עדכני): {current_time_str}</div>
    </div>
    """, unsafe_allow_html=True)

    dash_df = breakdown_df.copy()
    if selected_assembly != "הכל":
        dash_df = dash_df[dash_df["Assembly"] == selected_assembly]

    wip_cache_dash = fetch_wip_records()
    inv_cache_dash = fetch_all_inventory_records()

    total_planned_qty = 0.0
    total_executable_qty = 0.0
    total_planned_assemblies_count = 0
    blocked_assemblies = len(dash_df['Assembly'].unique()) if not dash_df.empty else 0

    assemblies_to_evaluate = [a for a in valid_assemblies if selected_assembly == "הכל" or a == selected_assembly]

    for asm_col in assemblies_to_evaluate:
        for target_m in selected_target_yms:
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan_df["Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache_dash.get(asm_col, 0.0)

            if raw_build > 0 or current_wip_qty > 0:
                total_planned_assemblies_count += 1
                total_planned_qty += raw_build

                month_breakdown = calculate_mrp_breakdown(target_yms=[target_m])
                asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

                max_possible_build = raw_build
                if not asm_shortages.empty and raw_build > 0:
                    for _, s_row in asm_shortages.iterrows():
                        req_per = s_row["Qty_Per_Assembly"]
                        if req_per > 0:
                            comp_pn = str(s_row["PN"]).strip()
                            # תיקון קריטי: זמינות מתחשבת ב-ETA (מלאי + כל אספקה
                            # שה-ETA שלה חל עד חודש היעד), לא רק מלאי סטטי -
                            # בדיוק אותה לוגיקה שתיקנו לבדיקת ה-WIP ההיררכית,
                            # עכשיו גם כאן ובכל מקום אחר שמחשב "ניתן לייצור".
                            total_comp_stock = get_component_available_by_month(comp_pn, target_m, inv_cache_dash)
                            possible_from_this = total_comp_stock / req_per
                            max_possible_build = min(max_possible_build, possible_from_this)
                    gross_executable = max(0.0, min(raw_build, max_possible_build))
                else:
                    gross_executable = raw_build

                net_executable_qty = max(0.0, gross_executable - current_wip_qty)
                total_executable_qty += net_executable_qty

    readiness_pct = (total_executable_qty / total_planned_qty * 100) if total_planned_qty > 0 else 100
    unique_shortage_count = len(dash_df['PN'].unique()) if not dash_df.empty else 0
    active_wip_list = [(w, q) for w, q in wip_cache_dash.items() if q > 0]
    total_wip_active_count = len(active_wip_list)

    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    with col_k1:
        kpi_card("🟢 מוכנות ייצור משוקללת", f"{readiness_pct:.1f}%", f"{total_executable_qty:,.0f} / {total_planned_qty:,.0f} יחידות ניתן לייצור", "green")
    with col_k2:
        kpi_card("🔴 הרכבות חסומות", blocked_assemblies, "בטווח הנבחר", "red")
    with col_k3:
        kpi_card("🏭 פעילים ב-WIP", total_wip_active_count, "הודעות ייצור פעילות", "blue")
    with col_k4:
        kpi_card("📦 מק'טים בגירעון", unique_shortage_count, "פריטים ייחודיים", "orange")
    with col_k5:
        kpi_card("📊 גירעון מצטברת", f"{dash_df['Total_MRP_Shortage'].sum():,.0f}" if not dash_df.empty else "0", "יחידות", "blue")

    with st.expander("🔍 הצג פירוט כרטיסי הרכבות פעילים ב-WIP (לחץ לפתיחה)", expanded=False):
        if active_wip_list:
            wip_detail_rows = [{"קוד הרכבה": asm_pn, "תיאור הרכבה": df_desc.iloc[0, df.columns.get_loc(asm_pn)] if asm_pn in df.columns else "", "כמות ב-WIP": asm_qty, "רמה בעץ": assembly_levels.get(asm_pn, 0)} for asm_pn, asm_qty in active_wip_list]
            st.dataframe(pd.DataFrame(wip_detail_rows), use_container_width=True)
        else:
            st.info("אין כרגע הרכבות פעילות ב-WIP.")

    st.divider()

    if not dash_df.empty and len(dash_df) > 0:
        col_g0, col_g1, col_g2 = st.columns([1, 1.2, 1.2])

        with col_g0:
            st.markdown("##### 🎯 מד מוכנות ייצור")
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=readiness_pct,
                number={'suffix': "%", 'font': {'size': 34}},
                gauge={
                    'axis': {'range': [0, 100]},
                    'bar': {'color': PRIMARY},
                    'steps': [
                        {'range': [0, 50], 'color': '#3B1F1F'},
                        {'range': [50, 80], 'color': '#3B2F1F'},
                        {'range': [80, 100], 'color': '#1F3B2A'},
                    ],
                }
            ))
            fig_gauge.update_layout(template=PLOTLY_TEMPLATE, height=260, margin=dict(t=10, b=10, l=20, r=20))
            st.plotly_chart(fig_gauge, use_container_width=True)

        with col_g1:
            st.markdown("##### 🥧 התפלגות חוסרים לפי סוג פריט")
            fig_pie = px.pie(dash_df, names="Item_Type", values="Total_MRP_Shortage", hole=0.5, color_discrete_sequence=COLOR_SEQ)
            fig_pie.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_g2:
            st.markdown("##### 🏭 התפלגות חוסרים לפי ספק")
            fig_sup = px.pie(dash_df, names="Supplier", values="Total_MRP_Shortage", hole=0.5, color_discrete_sequence=COLOR_SEQ)
            fig_sup.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_sup, use_container_width=True)

        st.markdown('<div class="section-title">📋 טבלת פירוט ניהולית עם אפשרות ייצוא וקישורי חיפוש מלאי</div>', unsafe_allow_html=True)
        display_df = dash_df[[
            "PN", "Description", "Item_Type", "Supplier", "Status", "Assembly", "Assembly_Desc",
            "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage",
            "חיפוש במאוזר", "חיפוש בדיגיקי", "חיפוש ב-Findchips"
        ]].rename(columns={
            "PN": "מק'ט", "Description": "תיאור פריט", "Item_Type": "סוג פריט", "Supplier": "ספק",
            "Status": "סטטוס טיפול", "Assembly": "קוד הרכבה", "Assembly_Desc": "תיאור הרכבה",
            "Qty_Per_Assembly": "כמות נדרשת", "Assembly_Monthly_Build": "ת. ייצור",
            "Required_Demand": "ביקוש מדויק", "Stock": "מלאי", "Total_MRP_Shortage": "סך חוסר"
        })

        def _shortage_color(val, vmax):
            if vmax <= 0:
                return ""
            ratio = min(1.0, float(val) / vmax)
            return f"background-color: rgba(239,{int(180 - ratio * 140)},{int(120 - ratio * 100)},0.55); color: white;"

        sorted_display_df = display_df.sort_values(by="סך חוסר", ascending=False)
        max_shortage = sorted_display_df["סך חוסר"].max() if not sorted_display_df.empty else 0

        styled = sorted_display_df.style.map(lambda v: _shortage_color(v, max_shortage), subset=["סך חוסר"])
        st.dataframe(styled, column_config={
            "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
            "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
            "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
        }, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='Executive_Shortages')
        st.download_button(label="📥 הורד דו'ח מנהלים מלא ל-Excel", data=output.getvalue(), file_name=f"MRP_Executive_Report_{selected_ym}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.success("🎉 אין חוסרים ב-MRP עבור ההגדרות והסינונים שנבחרו!")

with tab2:
    st.markdown(f'<div class="section-title">📊 סימולציית Clear To Build (CTB) מטריציונית עם השוואת כמויות וגרף הרכבות מפורט</div>', unsafe_allow_html=True)
    inv_cache_ctb = fetch_all_inventory_records()
    wip_cache_ctb = fetch_wip_records()
    matrix_rows, chart_assembly_data = [], []
    assemblies_to_check = [asm for asm in valid_assemblies if selected_assembly == "הכל" or asm == selected_assembly]

    for asm_col in assemblies_to_check:
        asm_desc = df_desc.iloc[0, df.columns.get_loc(asm_col)] if asm_col in df.columns else ""
        row_data = {"קוד הרכבה": asm_col, "תיאור הרכבה": asm_desc, "רמה בעץ": assembly_levels.get(asm_col, 0)}
        has_any_build = False

        for target_m in selected_target_yms:
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan_df["Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache_ctb.get(asm_col, 0.0)

            if raw_build > 0 or current_wip_qty > 0:
                has_any_build = True

            month_breakdown = calculate_mrp_breakdown(target_yms=[target_m])
            asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

            max_possible_build = raw_build
            if not asm_shortages.empty and raw_build > 0:
                for _, s_row in asm_shortages.iterrows():
                    req_per = s_row["Qty_Per_Assembly"]
                    if req_per > 0:
                        comp_pn = str(s_row["PN"]).strip()
                        # תיקון קריטי: אותה זמינות מתחשבת-ETA כמו בכל שאר המערכת
                        total_comp_stock = get_component_available_by_month(comp_pn, target_m, inv_cache_ctb)
                        max_possible_build = min(max_possible_build, total_comp_stock / req_per)
                gross_executable = max(0.0, min(raw_build, max_possible_build))
            else:
                gross_executable = raw_build

            net_executable_qty = max(0.0, gross_executable - current_wip_qty)
            row_data[f"תכנית ייצור ({target_m})"] = raw_build
            row_data[f"ניתן לייצור ({target_m})"] = net_executable_qty
            row_data[f"WIP ({target_m})"] = current_wip_qty

            if raw_build > 0 or current_wip_qty > 0:
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "תכנית ייצור", "כמות": raw_build})
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "ניתן לייצור בפועל", "כמות": net_executable_qty})
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "WIP", "כמות": current_wip_qty})

        for target_m in selected_target_yms:
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == target_m) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            raw_build = sub_plan_df["Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache_ctb.get(asm_col, 0.0)
            net_build = max(0.0, raw_build - current_wip_qty)

            month_breakdown = calculate_mrp_breakdown(target_yms=[target_m])
            asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

            missing_items_details = []
            for _, s_row in asm_shortages.iterrows():
                c_pn, c_desc, s_qty = str(s_row["PN"]).strip(), str(s_row["Description"]).strip(), s_row["Total_MRP_Shortage"]
                raw_eta = get_first_supply_eta(c_pn, inv_cache_ctb)
                missing_items_details.append((c_pn, c_desc, s_qty, raw_eta))

            if missing_items_details:
                formatted_missing = [f"{c_pn} ({c_desc[:10]}) - חסר: {m_qty:g} [ETA: {raw_eta}]" for c_pn, c_desc, m_qty, raw_eta in missing_items_details]
                row_data[f"סטטוס וחוסרים ({target_m})"] = "❌ חסר: " + " | ".join(formatted_missing)
            else:
                row_data[f"סטטוס וחוסרים ({target_m})"] = "✅ מוכן לייצור מלא" if net_build > 0 else "💤 ללא תוכנית ייצור"

        if has_any_build:
            matrix_rows.append(row_data)

    if matrix_rows:
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, height=420)
        if chart_assembly_data:
            fig_bar_asm = px.bar(pd.DataFrame(chart_assembly_data), x="הרכבה ותיאור", y="כמות", color="מדד", barmode="group", color_discrete_sequence=[PRIMARY, SUCCESS, ACCENT])
            fig_bar_asm.update_layout(template=PLOTLY_TEMPLATE, height=400, margin=dict(t=20, b=40, l=20, r=20), xaxis_tickangle=-25)
            st.plotly_chart(fig_bar_asm, use_container_width=True)

with tab3:
    st.markdown('<div class="section-title">💡 סימולציית What-If (מה יקרה אם...)</div>', unsafe_allow_html=True)
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        sim_pn = st.selectbox("בחר מק'ט לסימולציה", sorted(df[PN_COL].dropna().astype(str).unique()), key="sim_pn")
    with col_w2:
        sim_extra_stock = st.number_input("תוספת כמות מדומיינת למלאי", min_value=0.0, value=10.0, step=1.0)

    if st.button("🔮 הרץ סימולציית שחרור צוואר בקבוק"):
        sim_df = calculate_mrp_breakdown({sim_pn: sim_extra_stock}, target_yms=selected_target_yms)
        orig_blocked = set(breakdown_df['Assembly'].unique()) if not breakdown_df.empty else set()
        sim_blocked = set(sim_df['Assembly'].unique()) if not sim_df.empty else set()
        st.success(f"סימולציה הופעלה בהצלחה עבור מק'ט `{sim_pn}`.")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            kpi_card("🟢 הרכבות שהשתחררו", len(orig_blocked - sim_blocked), "", "green")
        with col_m2:
            kpi_card("🔴 עדיין חסום", len(sim_blocked), "", "red")
        with col_m3:
            before_after_delta = (breakdown_df['Total_MRP_Shortage'].sum() if not breakdown_df.empty else 0) - (sim_df['Total_MRP_Shortage'].sum() if not sim_df.empty else 0)
            kpi_card("📉 צמצום גירעון", f"{before_after_delta:,.0f}", "יחידות", "blue")

with tab4:
    st.markdown('<div class="section-title">📌 לוח מעקב סטטוסים (Kanban Pipeline)</div>', unsafe_allow_html=True)
    statuses = [("פתוח", "📝 פתוח לטיפול", "#3B1F1F", DANGER), ("הוזמן", "🛒 הוזמן / בטיפול רכש", "#3B2F1F", WARNING), ("בדרך", "🚚 בדרך לקו", "#1F2A3B", ACCENT), ("התקבל", "✅ התקבל / סגור", "#1F3B2A", SUCCESS)]
    dedup_all = breakdown_df.drop_duplicates(subset=["PN"]) if not breakdown_df.empty else pd.DataFrame()
    kcols = st.columns(4)
    for (status_key, title, bg, accent_color), kcol in zip(statuses, kcols):
        with kcol:
            items = dedup_all[dedup_all["Status"] == status_key] if not dedup_all.empty else pd.DataFrame()
            st.markdown(f'<div class="kanban-col-header" style="background:{bg}; color:{accent_color};">{title} ({len(items)})</div>', unsafe_allow_html=True)
            for _, r in items.head(6).iterrows():
                st.markdown(f'<div class="kanban-card" style="border-color:{accent_color};"><b>{r["PN"]}</b><br><span style="opacity:0.75;">{str(r["Description"])[:24]}</span></div>', unsafe_allow_html=True)

with tab5:
    st.markdown(f'<div class="section-title">🏭 ניהול WIP חכם (כולל סגירת מחזור ייצור ואימות היררכיה)</div>', unsafe_allow_html=True)
    wip_current = fetch_wip_records()
    if wip_current:
        with st.form("close_wip_form"):
            wip_to_close = st.selectbox("בחר הרכבה שסיימה ייצור לחודש זה", list(wip_current.keys()), format_func=lambda x: f"{x} (כמות ב-WIP: {wip_current[x]})")
            is_finished = st.checkbox("האם ההרכבה הסתיימה לחלוטין והושלמה בהצלחה?")
            if st.form_submit_button("סגור WIP והוסף למלאי הזמין"):
                if is_finished:
                    closing_qty = wip_current[wip_to_close]
                    curr_stk, curr_eta, curr_stat, curr_sup, curr_comm, _, _ = get_inventory_record(wip_to_close)
                    save_inventory_record(wip_to_close, curr_stk + closing_qty, curr_eta, "התקבל", curr_sup, f"{curr_comm} | הושלם מייצור WIP בסך {closing_qty}", "WIP Close", webhook_url)
                    delete_wip_record(wip_to_close)
                    st.rerun()

    st.divider()
    st.markdown("##### ➕ צירוף הרכבה חדשה ל-WIP")
    wip_asm_choice = st.selectbox("בחר הרכבה חדשה לצירוף ל-WIP", filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x), key="wip_asm_choice")
    wip_qty_input = st.number_input("כמות יחידות הרכבה להוספה לייצור (WIP)", min_value=0.0, value=1.0, step=1.0, key="wip_qty_input")
    wip_target_month_label = st.selectbox(
        "לאיזה חודש בונים? (הבדיקה תתחשב באספקה/ETA שכבר אמורים להגיע עד חודש זה, לא רק במלאי הנוכחי בקופה)",
        list(month_options.keys()),
        index=list(month_options.keys()).index(selected_month_label) if selected_month_label in month_options else 0,
        key="wip_target_month_label"
    )
    wip_target_ym = pd.to_datetime(month_options[wip_target_month_label]).strftime("%Y-%m")

    # תיקון קריטי: כאן בפועל רצה עכשיו בדיקת הזמינות ההיררכית שהכפתור
    # תמיד הבטיח ולא ביצע. הבדיקה מתחשבת ב-ETA (מלאי נוכחי + כל אספקה
    # שה-ETA שלה חל עד חודש הבנייה שנבחר, בדיוק כמו מנגנון ה-MRP הרגיל)
    # ולא רק במה שכבר פיזית בקופה. הבדיקה רצה בכל שינוי (מחוץ לטופס) כדי
    # שהמשתמש יראה מיד את התוצאה, והאישור הסופי נשאר בתוך טופס לשמירה אטומית.
    hierarchy_blockers = check_hierarchical_ctb(wip_asm_choice, wip_qty_input, wip_target_ym) if wip_qty_input > 0 else []

    if hierarchy_blockers:
        st.error(f"⛔ נמצאו {len(hierarchy_blockers)} חוסרים בעץ ההרכבה עד חודש {wip_target_ym} (בהרכבה עצמה ו/או בתתי-ההרכבות שלה) - כולל התחשבות ב-ETA צפוי:")
        blockers_df = pd.DataFrame(hierarchy_blockers).rename(columns={
            "assembly": "קוד הרכבה חוסמת", "assembly_desc": "תיאור הרכבה חוסמת",
            "component": "מק\"ט חסר", "component_desc": "תיאור פריט",
            "required": "נדרש", "available": "זמין", "shortage": "חוסר"
        })
        st.dataframe(blockers_df, use_container_width=True, height=min(300, 45 + 35 * len(blockers_df)))
        with st.form("wip_form"):
            override_confirm = st.checkbox("⚠️ ידוע לי שיש חוסרים בעץ ההרכבה, ואני מאשר בכל זאת להוסיף ל-WIP (למשל אם מדובר בהזמנת ייצור מתוכננת מראש)")
            if st.form_submit_button("שמור ל-WIP בכל זאת"):
                if override_confirm:
                    save_wip_record(wip_asm_choice, wip_qty_input)
                    st.success("ההרכבה נוספה ל-WIP (עם חוסרים ידועים).")
                    st.rerun()
                else:
                    st.warning("יש לסמן את תיבת האישור כדי לשמור למרות החוסרים.")
    else:
        st.success("✅ נבדק עץ ההרכבה המלא - כל רכיבי הגלם וכל תתי-ההרכבות זמינים לכמות המבוקשת.")
        with st.form("wip_form"):
            if st.form_submit_button("בדיקת זמינות היררכית מלאה ושמור WIP"):
                save_wip_record(wip_asm_choice, wip_qty_input)
                st.success("ההרכבה נוספה בהצלחה ל-WIP!")
                st.rerun()

with tab6:
    st.markdown('<div class="section-title">📅 עדכון מלאי, סטטוס ודחיית ספקים (ETA)</div>', unsafe_allow_html=True)
    selected_pn = search_pn if search_pn != "הכל" else st.selectbox("בחר מק'ט מכלל הפריטים לעדכון", sorted(df[PN_COL].dropna().astype(str).unique()), key="update_pn_select")
    if selected_pn != "הכל":
        saved_stock, saved_eta, saved_status, saved_supplier, saved_comment, saved_by, _ = get_inventory_record(selected_pn)
        with st.form("inventory_form"):
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                added_stock_input = st.number_input("תוספת למלאי זמין", min_value=0.0, value=float(saved_stock), step=1.0)
            with col_f2:
                try: parsed_eta = pd.to_datetime(saved_eta).date() if saved_eta else date.today()
                except: parsed_eta = date.today()
                eta_date = st.date_input("תאריך הגעה מעודכן (ETA)", value=parsed_eta)
            with col_f3:
                status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
                status = st.selectbox("סטטוס טיפול", status_options, index=status_options.index(saved_status) if saved_status in status_options else 0)
            supplier = st.selectbox("ספק", supplier_options, index=supplier_options.index(saved_supplier) if saved_supplier in supplier_options else 0)
            comment = st.text_area("הערות", value=saved_comment)
            if st.form_submit_button("שמור עדכון קבוע בענן"):
                save_inventory_record(selected_pn, added_stock_input, str(eta_date), status, supplier, comment, "User", webhook_url)
                st.success("העדכון נשמר!")
                st.rerun()

with tab7:
    st.markdown('<div class="section-title">📅 מעקב ETA, דחיות, כמויות וקישורים למפיצים</div>', unsafe_allow_html=True)
    inv_cache_all = fetch_all_inventory_records()
    eta_table_rows = []

    for _, row in df.iterrows():
        p_num = str(row[PN_COL]).strip()
        if not p_num or p_num == 'nan':
            continue
        p_desc = str(row[DESC_COL])
        p_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in df.columns else ""

        orig_eta = get_base_mrp_eta(p_num)
        orig_qty = get_base_mrp_qty(p_num)
        
        saved_rec = inv_cache_all.get(p_num, {})
        current_eta_raw = saved_rec.get("eta", "")
        current_added_stock = saved_rec.get("added_stock", 0.0)
        curr_eta_fmt = pd.to_datetime(current_eta_raw).strftime("%Y-%m") if current_eta_raw else orig_eta

        eta_table_rows.append({
            "מק'ט": p_num,
            "תיאור פריט": p_desc,
            "סוג פריט": p_type,
            "ETA מקורי (MRP)": orig_eta,
            "כמות מקורית": orig_qty,
            "ETA מעודכן": curr_eta_fmt,
            "כמות מעודכנת": current_added_stock,
            "ספק": saved_rec.get("supplier", "אופק"),
            "חיפוש במאוזר": f"https://www.mouser.co.il/c/?q={p_num}",
            "חיפוש בדיגיקי": f"https://www.digikey.com/en/products/result?keywords={p_num}",
            "חיפוש ב-Findchips": f"https://www.findchips.com/search/{p_num}"
        })

    eta_df = pd.DataFrame(eta_table_rows)
    if not eta_df.empty:
        # שיפור שקיפות: מציג בבירור כמה פריטים עדיין ללא ETA בסיסי מה-MRP
        # (מק"טים שהפונקציה לא הצליחה לשייך להם אספקה/יתרה עתידית -
        # ראו "בדיקה נדרשת"), כדי שהצוות ידע לתעדף בדיקה ידנית שלהם.
        needs_review_count = int((eta_df["ETA מקורי (MRP)"] == "בדיקה נדרשת").sum())
        st.metric("🔎 מק\"טים ללא ETA בסיסי במערכת (דורשים בדיקה ידנית)", needs_review_count)

        def _highlight_needs_review(val):
            if val == "בדיקה נדרשת":
                return f"background-color: {WARNING}; color: white; font-weight: 700;"
            return ""

        eta_df_styled = eta_df.style.map(_highlight_needs_review, subset=["ETA מקורי (MRP)"])
        st.dataframe(eta_df_styled, column_config={
            "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
            "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
            "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
        }, use_container_width=True, height=450)

with tab8:
    st.markdown('<div class="section-title">↩️ חזרה לאחור וניהול היסטוריה (UNDO)</div>', unsafe_allow_html=True)
    try: updated_items = supabase.table("mrp_inventory_updates").select("*").order("updated_at", desc=True).execute().data or []
    except: updated_items = []
    for item in updated_items:
        col_u1, col_u2, col_u3 = st.columns([3, 4, 1])
        with col_u1: st.markdown(f"**מק'ט:** `{item.get('pn')}`")
        with col_u2: st.text(f"תוספת: {item.get('added_stock')} | ETA: {item.get('eta')}")
        with col_u3:
            if st.button("🔄 UNDO", key=f"undo_{item.get('pn')}"):
                delete_inventory_record(item.get('pn'))
                st.rerun()

with tab9:
    st.markdown('<div class="section-title">📦 ניהול מלאי מעודכן (עריכה וגריעת כמויות)</div>', unsafe_allow_html=True)
    active_stock_items = {k: v for k, v in fetch_all_inventory_records().items() if float(v.get("added_stock", 0.0)) > 0}
    if active_stock_items:
        st.dataframe(pd.DataFrame([{"מק'ט": k, "כמות": v["added_stock"], "ETA": v["eta"]} for k, v in active_stock_items.items()]), use_container_width=True)
        selected_mgmt_pn = st.selectbox("בחר מק'ט לעריכה או גריעה", list(active_stock_items.keys()), key="mgmt_pn_select")
        if selected_mgmt_pn:
            with st.form("edit_mgmt_form"):
                new_qty = st.number_input("עדכן כמות", min_value=0.0, value=float(active_stock_items[selected_mgmt_pn]["added_stock"]), step=1.0)
                if st.form_submit_button("🗑️ אפס או עדכן"):
                    delete_inventory_record(selected_mgmt_pn)
                    if new_qty > 0:
                        save_inventory_record(selected_mgmt_pn, new_qty, "", "פתוח", "אופק", "", "Tab 9", webhook_url)
                    st.rerun()

with tab10:
    st.markdown('<div class="section-title">🎯 ניתוח רגישות וניהול תוכנית הייצור (עריכה פרטנית לפי הרכבה וחודש)</div>', unsafe_allow_html=True)
    if not assembly_plan_df.empty:
        orig_pivot_plan = assembly_plan_df.pivot_table(index=["Assembly_PN"], columns="YearMonth", values="Build_Qty", fill_value=0.0).reset_index()
        orig_pivot_plan.insert(1, "רמה", orig_pivot_plan["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        orig_pivot_plan.insert(2, "תיאור הרכבה", orig_pivot_plan["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        orig_pivot_plan = orig_pivot_plan.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(orig_pivot_plan, use_container_width=True, height=280)

    st.divider()
    st.markdown("##### ⚙️ הגדרת שינוי רגישות: גורף או חודש ספציפי")
    
    col_mode_choice = st.columns(2)
    with col_mode_choice[0]:
        sens_scope = st.radio("היקף השינוי", ["שינוי גורף לכל החודשים", "שינוי לחודש ספציפי בלבד"], horizontal=True, key="sens_scope")
    with col_mode_choice[1]:
        if sens_scope == "שינוי לחודש ספציפי בלבד":
            available_yms = sorted(assembly_plan_df["YearMonth"].unique())
            target_sens_month = st.selectbox("בחר חודש ספציפי לעדכון", available_yms, key="target_sens_month")

    col_sens1, col_sens2, col_sens3 = st.columns([1.2, 1, 1])
    with col_sens1:
        sens_assembly_target = st.selectbox("בחר הרכבה לניתוח רגישות", ["הכל (כלל ההרכבות)"] + filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x), key="sens_assembly_target")
    with col_sens2:
        sens_mode = st.radio("סוג שינוי", ["אחוזים (%)", "מספרי (יחידות)"], horizontal=True, key="sens_mode")
    with col_sens3:
        if sens_mode == "אחוזים (%)":
            sensitivity_val = st.slider("שינוי אחוז תוכנית הייצור (%)", -50, 100, 0, 5, key="sens_slider")
        else:
            sensitivity_val = st.number_input("תוספת/הפחתה מספרית (יחידות)", -500, 500, 0, 1, key="sens_num")

    if st.button("🚀 הרץ ניתוח רגישות לתוכנית", key="run_sensitivity"):
        simulated_plan_df = assembly_plan_df.copy()
        sys_factor_map = ASSEMBLY_SYSTEM_FACTORS
        
        if sens_scope == "שינוי גורף לכל החודשים":
            if sens_mode == "אחוזים (%)" and sensitivity_val != 0:
                multiplier = 1.0 + (sensitivity_val / 100.0)
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    simulated_plan_df["Raw_Build_Qty"] *= multiplier
                    simulated_plan_df["Build_Qty"] *= multiplier
                else:
                    mask = simulated_plan_df["Assembly_PN"] == sens_assembly_target
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
            elif sens_mode == "מספרי (יחידות)" and sensitivity_val != 0:
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    for idx_row in simulated_plan_df.index:
                        asm_code = str(simulated_plan_df.loc[idx_row, "Assembly_PN"]).strip()
                        sys_f = sys_factor_map.get(asm_code, 1)
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f
                else:
                    mask = simulated_plan_df["Assembly_PN"] == sens_assembly_target
                    sys_f = sys_factor_map.get(sens_assembly_target, 1)
                    for idx_row in simulated_plan_df[mask].index:
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f
        else:
            if sens_mode == "אחוזים (%)" and sensitivity_val != 0:
                multiplier = 1.0 + (sensitivity_val / 100.0)
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    mask = simulated_plan_df["YearMonth"] == target_sens_month
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
                else:
                    mask = (simulated_plan_df["Assembly_PN"] == sens_assembly_target) & (simulated_plan_df["YearMonth"] == target_sens_month)
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
            elif sens_mode == "מספרי (יחידות)" and sensitivity_val != 0:
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    mask = simulated_plan_df["YearMonth"] == target_sens_month
                    for idx_row in simulated_plan_df[mask].index:
                        asm_code = str(simulated_plan_df.loc[idx_row, "Assembly_PN"]).strip()
                        sys_f = sys_factor_map.get(asm_code, 1)
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f
                else:
                    mask = (simulated_plan_df["Assembly_PN"] == sens_assembly_target) & (simulated_plan_df["YearMonth"] == target_sens_month)
                    sys_f = sys_factor_map.get(sens_assembly_target, 1)
                    for idx_row in simulated_plan_df[mask].index:
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f

        st.session_state["temp_simulated_plan"] = simulated_plan_df
        st.success("ניתוח הרגישות בוצע בהצלחה! צפה בתוצאות המעודכנות למטה.")

    if "temp_simulated_plan" in st.session_state:
        st.divider()
        st.markdown("##### 📋 תצוגה מקדימה של התוכנית הסימולטיבית (לאחר ניתוח רגישות):")
        preview_pivot = st.session_state["temp_simulated_plan"].pivot_table(index=["Assembly_PN"], columns="YearMonth", values="Build_Qty", fill_value=0.0).reset_index()
        preview_pivot.insert(1, "רמה", preview_pivot["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        preview_pivot.insert(2, "תיאור הרכבה", preview_pivot["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        preview_pivot = preview_pivot.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(preview_pivot, use_container_width=True, height=240)

        with st.form("update_plan_form"):
            update_confirmation = st.checkbox("❓ האם אתה מאשר לשמור את השינויים ולהחיל את תוכנית הייצור החדשה על כלל המערכת?")
            if st.form_submit_button("💾 שמור שינויים ועדכן את תוכנית העבודה"):
                if update_confirmation:
                    st.session_state["previous_approved_plan"] = assembly_plan_df.copy()
                    st.session_state["custom_assembly_plan_df"] = st.session_state["temp_simulated_plan"]
                    save_cloud_assembly_plan(st.session_state["temp_simulated_plan"])
                    del st.session_state["temp_simulated_plan"]
                    st.success("תוכנית הייצור עודכנה ונשמרה בהצלחה בענן ובמערכת!")
                    st.rerun()
                else:
                    st.warning("יש לסמן את תיבת האישור כדי לשמור את השינויים.")

    if "previous_approved_plan" in st.session_state:
        st.divider()
        st.markdown("##### 📊 השוואה בין תוכנית הייצור הקודמת לחדשה")
        orig_plan_pivot = st.session_state["previous_approved_plan"].pivot_table(index="Assembly_PN", columns="YearMonth", values="Build_Qty", fill_value=0.0)
        new_plan_pivot = assembly_plan_df.pivot_table(index="Assembly_PN", columns="YearMonth", values="Build_Qty", fill_value=0.0)
        comparison_diff = new_plan_pivot.sub(orig_plan_pivot, fill_value=0.0).reset_index()
        comparison_diff.insert(1, "רמה", comparison_diff["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        comparison_diff.insert(2, "תיאור הרכבה", comparison_diff["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        comparison_diff = comparison_diff.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(comparison_diff, use_container_width=True)

with tab11:
    # ==========================================================
    # טאב חדש: עריכת ETA מרוכזת לכל הפריטים
    # ==========================================================
    # טבלה אחת עם כל הפריטים, שאפשר לערוך בה ישירות ETA / תוספת מלאי /
    # סטטוס / ספק / הערות, ולשמור הכל ל-DB (Supabase) בלחיצה אחת.
    # חשוב: זו בדיוק אותה טבלה (mrp_inventory_updates) שממנה כל שאר
    # המערכת קוראת - חישובי ה-MRP, ה-CTB, ובדיקת הזמינות ההיררכית ב-WIP
    # (get_inventory_record) - כך שכל שינוי שנשמר כאן מוזן אוטומטית לכל
    # החישובים בכל שאר הטאבים, בלי צורך בשום שינוי נוסף.
    st.markdown('<div class="section-title">✏️ עריכת ETA מרוכזת לכל הפריטים</div>', unsafe_allow_html=True)
    st.caption("עדכון כאן נשמר ישירות ב-DB, ומשפיע מיידית על כל חישובי ה-MRP, ה-CTB ובדיקת הזמינות ההיררכית ב-WIP בכל שאר הטאבים.")

    inv_cache_bulk = fetch_all_inventory_records()

    col_bf1, col_bf2, col_bf3 = st.columns([1.2, 1.5, 1])
    with col_bf1:
        bulk_item_type = st.selectbox("סינון לפי סוג פריט", ["הכל"] + item_types, key="bulk_item_type")
    with col_bf2:
        bulk_search = st.text_input("חיפוש לפי מק\"ט או תיאור", key="bulk_search")
    with col_bf3:
        bulk_only_shortage = st.checkbox("הצג רק פריטים שכרגע בחוסר", key="bulk_only_shortage")

    shortage_pns = set(breakdown_df["PN"].unique()) if not breakdown_df.empty else set()
    bulk_rows = []

    for _, row in df.iterrows():
        p_num = str(row[PN_COL]).strip()
        if not p_num or p_num == 'nan':
            continue
        if bulk_only_shortage and p_num not in shortage_pns:
            continue
        p_desc = str(row[DESC_COL])
        p_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in df.columns else ""
        if bulk_item_type != "הכל" and p_type != bulk_item_type:
            continue
        if bulk_search and bulk_search.strip():
            needle = bulk_search.strip().lower()
            if needle not in p_num.lower() and needle not in p_desc.lower():
                continue

        orig_eta = get_base_mrp_eta(p_num)
        orig_qty = get_base_mrp_qty(p_num)
        saved_rec = inv_cache_bulk.get(p_num, {})
        current_eta_raw = saved_rec.get("eta", "")
        try:
            current_eta_date = pd.to_datetime(current_eta_raw).date() if current_eta_raw else None
        except Exception:
            current_eta_date = None

        bulk_rows.append({
            "מק\"ט": p_num,
            "תיאור פריט": p_desc,
            "סוג פריט": p_type,
            "ETA מקורי (MRP)": orig_eta,
            "כמות מקורית (MRP)": orig_qty,
            "ETA מעודכן": current_eta_date,
            "תוספת מלאי": float(saved_rec.get("added_stock", 0.0) or 0.0),
            "סטטוס": saved_rec.get("status", "פתוח") or "פתוח",
            "ספק": saved_rec.get("supplier", "אופק") or "אופק",
            "הערות": saved_rec.get("comment", "") or "",
        })

    if not bulk_rows:
        st.info("לא נמצאו פריטים התואמים לסינון שנבחר.")
    else:
        bulk_df = pd.DataFrame(bulk_rows)
        st.caption(f"מציג {len(bulk_df)} פריטים. ניתן לערוך ETA מעודכן / תוספת מלאי / סטטוס / ספק / הערות ישירות בטבלה, ואז ללחוץ על 'שמור' למטה.")

        edited_df = st.data_editor(
            bulk_df,
            key="bulk_eta_editor",
            use_container_width=True,
            height=520,
            hide_index=True,
            disabled=["מק\"ט", "תיאור פריט", "סוג פריט", "ETA מקורי (MRP)", "כמות מקורית (MRP)"],
            column_config={
                "ETA מעודכן": st.column_config.DateColumn("ETA מעודכן", format="YYYY-MM-DD"),
                "תוספת מלאי": st.column_config.NumberColumn("תוספת מלאי", min_value=0.0, step=1.0),
                "סטטוס": st.column_config.SelectboxColumn("סטטוס", options=["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]),
                "ספק": st.column_config.SelectboxColumn("ספק", options=supplier_options),
            }
        )

        if st.button("💾 שמור את כל השינויים ל-DB", key="bulk_save_btn"):
            changed_count = 0
            for i in range(len(bulk_df)):
                orig_row = bulk_df.iloc[i]
                new_row = edited_df.iloc[i]

                new_stock_val = float(new_row["תוספת מלאי"]) if pd.notnull(new_row["תוספת מלאי"]) else 0.0
                changed = (
                    str(orig_row["ETA מעודכן"]) != str(new_row["ETA מעודכן"]) or
                    float(orig_row["תוספת מלאי"]) != new_stock_val or
                    orig_row["סטטוס"] != new_row["סטטוס"] or
                    orig_row["ספק"] != new_row["ספק"] or
                    orig_row["הערות"] != new_row["הערות"]
                )
                if changed:
                    save_inventory_record(
                        pn=orig_row["מק\"ט"],
                        added_stock=new_stock_val,
                        eta=str(new_row["ETA מעודכן"]) if new_row["ETA מעודכן"] else "",
                        status=new_row["סטטוס"],
                        supplier=new_row["ספק"],
                        comment=new_row["הערות"],
                        updated_by="Bulk ETA Editor",
                        webhook_url=webhook_url
                    )
                    changed_count += 1

            if changed_count > 0:
                st.success(f"נשמרו {changed_count} שינויים ל-DB. כל החישובים בכל הטאבים יתעדכנו בהתאם.")
                st.rerun()
            else:
                st.info("לא זוהו שינויים לשמירה.")
