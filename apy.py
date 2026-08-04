# ==========================================================
# עדכון רשימת ההרכבות ותצוגתן תוך שמירה מלאה על כלל יכולות המערכת
# ==========================================================

# החלף את מקטע טעינת התוכנית ורשימת ההרכבות בקוד שלך בקטע הבא:

if "custom_assembly_plan_df" not in st.session_state:
    header_dates = df_raw.iloc[2, 108:132].values if df_raw.shape[1] > 132 else []
    plan_rows = []
    ordered_assemblies_list = []

    # סריקה רוחבית של כלל השורות בקובץ הגולמי כדי לאפס את מגבלת ה-24 שורות הקודמת
    # ולכלול את כל ההרכבות (כולל ההרכבות התחתונות כגון REAR COVER ו-FINGER GUARD) בסדר המקורי שלהן
    for r in range(3, df_raw.shape[0]):
        asm_pn = df_raw.iloc[r, 106] if df_raw.shape[1] > 106 else None
        if pd.notnull(asm_pn):
            clean_asm_pn = str(asm_pn).strip()
            if clean_asm_pn and clean_asm_pn != 'nan':
                if clean_asm_pn not in ordered_assemblies_list:
                    ordered_assemblies_list.append(clean_asm_pn)

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
                                    if ym_str >= "2026-09":
                                        plan_rows.append({
                                            "Assembly_PN": clean_asm_pn,
                                            "YearMonth": ym_str,
                                            "Build_Qty": q_val * system_multiplier,
                                            "Raw_Build_Qty": q_val
                                        })
                            except:
                                pass

    st.session_state["custom_assembly_plan_df"] = pd.DataFrame(plan_rows)
    st.session_state["ordered_assemblies"] = ordered_assemblies_list

assembly_plan_df = st.session_state["custom_assembly_plan_df"]
ordered_assemblies = st.session_state.get("ordered_assemblies", [])

PN_COL = df.columns[1]
DESC_COL = df.columns[4]
ITEM_TYPE_COL = df.columns[44] if len(df.columns) > 44 else df.columns[-1]
STOCK_COL = df.columns[79] if len(df.columns) > 79 else df.columns[-1]
ASSEMBLY_COLS = df.columns[10:36].tolist()
MONTH_COLS = df.columns[108:132].tolist() if len(df.columns) > 132 else []

# הבטחת שמירה על כל ההרכבות והסדר המדויק שהוצג בתמונה ללא פגיעה בלוגיקה קיימת
valid_assemblies = [col for col in ASSEMBLY_COLS if col in df[PN_COL].values]
if not valid_assemblies:
    valid_assemblies = ordered_assemblies
else:
    for asm in ordered_assemblies:
        if asm not in valid_assemblies:
            valid_assemblies.append(asm)

filtered_assembly_cols = valid_assemblies
