st.markdown(f"""
<style>
/* -----------------------------------------------------------
   GLOBAL COLORS & THEME TOKENS
----------------------------------------------------------- */
:root {{
    --clr-primary: #4F46E5;
    --clr-primary-dark: #3730A3;
    --clr-accent: #06B6D4;
    --clr-danger: #EF4444;
    --clr-warning: #F59E0B;
    --clr-success: #10B981;
    --clr-bg-card-light: #ffffff;
    --clr-bg-card-dark: #1f2937;
    --clr-text-light: #111827;
    --clr-text-dark: #f9fafb;
}}

@import url('https://fonts.googleapis.com/css2?family=Assistant:wght@400;600;700;800&display=swap');

[data-testid="stAppViewContainer"] .main .block-container,
[data-testid="stSidebarContent"] {{
    font-family: 'Assistant', sans-serif;
    direction: rtl;
}}
[data-testid="stAppViewContainer"] .main .block-container *,
[data-testid="stSidebarContent"] * {{
    font-family: 'Assistant', sans-serif;
}}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

/* -----------------------------------------------------------
   TOP NAVIGATION BAR
----------------------------------------------------------- */
.top-nav {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    background: var(--clr-primary-dark);
    padding: 12px 20px;
    border-radius: 12px;
    margin-bottom: 20px;
    color: white;
}}
.top-nav a {{
    color: white;
    margin-left: 20px;
    font-weight: 600;
    text-decoration: none;
}}
.top-nav a:hover {{
    opacity: 0.8;
}}

/* -----------------------------------------------------------
   HERO BANNER
----------------------------------------------------------- */
.hero-banner {{
    background: linear-gradient(120deg, var(--clr-primary) 0%, var(--clr-primary-dark) 45%, var(--clr-accent) 100%);
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

/* -----------------------------------------------------------
   KPI CARDS — MODERN ENTERPRISE STYLE
----------------------------------------------------------- */
.kpi-card {{
    display: flex;
    align-items: center;
    gap: 14px;
    background: linear-gradient(135deg, var(--clr-primary) 0%, var(--clr-primary-dark) 100%);
    color: white;
    padding: 20px;
    border-radius: 16px;
    box-shadow: 0 8px 20px rgba(0,0,0,0.15);
    transition: transform 0.15s ease;
}}
.kpi-card:hover {{
    transform: translateY(-4px);
}}
.kpi-icon {{
    font-size: 40px;
    opacity: 0.85;
}}
.kpi-label {{
    font-size: 13px;
    opacity: 0.85;
    margin-bottom: 4px;
    font-weight: 600;
}}
.kpi-value {{
    font-size: 34px;
    font-weight: 800;
}}
.kpi-sub {{
    font-size: 12px;
    opacity: 0.7;
    margin-top: 4px;
}}

/* -----------------------------------------------------------
   SECTION TITLES
----------------------------------------------------------- */
.section-title {{
    font-weight: 800;
    font-size: 22px;
    margin: 28px 0 14px 0;
    border-right: 6px solid var(--clr-primary);
    padding-right: 14px;
    color: var(--clr-text-light);
}}

/* -----------------------------------------------------------
   KANBAN CARDS
----------------------------------------------------------- */
.kanban-card {{
    background: var(--clr-bg-card-light);
    border-left: 6px solid var(--clr-danger);
    padding: 14px;
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    margin-bottom: 10px;
    font-size: 13px;
}}

/* -----------------------------------------------------------
   BUTTONS — MATERIAL STYLE
----------------------------------------------------------- */
.stButton>button {{
    background-color: var(--clr-primary);
    color: white;
    padding: 10px 18px;
    border-radius: 10px;
    font-weight: 600;
    border: none;
    transition: 0.2s;
}}
.stButton>button:hover {{
    background-color: var(--clr-primary-dark);
    transform: translateY(-2px);
}}

/* -----------------------------------------------------------
   TABLE HIGHLIGHTING FOR SHORTAGES
----------------------------------------------------------- */
.shortage-high {{
    background-color: rgba(239, 68, 68, 0.15) !important;
}}
.shortage-mid {{
    background-color: rgba(245, 158, 11, 0.15) !important;
}}

/* -----------------------------------------------------------
   LOADER ANIMATION
----------------------------------------------------------- */
.loader {{
  border: 6px solid #f3f3f3;
  border-top: 6px solid var(--clr-accent);
  border-radius: 50%;
  width: 40px;
  height: 40px;
  animation: spin 1s linear infinite;
  margin: auto;
}}
@keyframes spin {{
  0% {{ transform: rotate(0deg); }}
  100% {{ transform: rotate(360deg); }}
}}

</style>
""", unsafe_allow_html=True)
