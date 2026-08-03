with tab5:
    st.markdown('<div class="section-title">🏭 ניהול WIP חכם (בדיקת עץ היררכית מדויקת Strict BOM Explosion)</div>', unsafe_allow_html=True)
    st.markdown("בדיקה הנדסית מדויקת: סריקה אך ורק של השורות התלויות ישירות בהרכבה הנבחרת בעצי המוצר.")

    wip_current = fetch_wip_records()

    with st.form("wip_form"):
        wip_asm_choice = st.selectbox("בחר הרכבה לצירוף ל-WIP", filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x))
        current_wip_val = wip_current.get(wip_asm_choice, 0.0)
        wip_qty_input = st.number_input("כמות יחידות הרכבה להורדה לייצור", min_value=0.0, value=float(current_wip_val if current_wip_val > 0 else 1.0), step=1.0)

        submitted_wip = st.form_submit_button("בדיקת זמינות קפדנית ושמור WIP")

        if submitted_wip:
            inv_cache_wip = fetch_all_inventory_records()
            
            def get_strict_bom_issues(target_asm):
                strict_issues = []
                
                # בדיקה האם ההרכבה קיימת כעמודה בעץ המוצר
                if target_asm in df.columns:
                    # סינון אך ורק של השורות שבהן יש דרישה/כמות בעמודה של ההרכבה הנבחרת (בנים ישירים)
                    child_rows = df[df[target_asm].notnull() & (pd.to_numeric(df[target_asm], errors='coerce') > 0)]
                    
                    for _, c_row in child_rows.iterrows():
                        c_pn = str(c_row[PN_COL]).strip()
                        if c_pn == target_asm:
                            continue
                        
                        # חישוב יתרה מעודכנת לפריט הבן
                        monthly_bal = pd.to_numeric(c_row.get(selected_month_col, 0), errors='coerce') or 0
                        added_stk, man_eta, _, _, _, _, _ = get_inventory_record(c_pn, inv_cache_wip)
                        eff_bal = monthly_bal + added_stk

                        is_late = False
                        if man_eta and str(man_eta).strip() not in ["", "None", "NaT", "nan"]:
                            try:
                                if pd.to_datetime(man_eta).strftime("%Y-%m") > selected_ym:
                                    is_late = True
                            except:
                                pass

                        if eff_bal < 0 or is_late:
                            q_missing = abs(eff_bal) if eff_bal < 0 else 1.0
                            strict_issues.append({
                                "PN": c_pn,
                                "Description": str(c_row[DESC_COL]),
                                "Reason": f"גירעון של {q_missing:g} יח'" if eff_bal < 0 else f"עיכוב ספק (ETA: {man_eta})"
                            })
                else:
                    # גיבוי למקרה שההרכבה מופיעה רק בדוח החוסרים הכללי
                    gen_shortages = breakdown_df[breakdown_df["Assembly"] == target_asm]
                    for _, g_row in gen_shortages.iterrows():
                        strict_issues.append({
                            "PN": str(g_row["PN"]).strip(),
                            "Description": str(g_row["Description"]),
                            "Reason": f"חוסר של {g_row['Total_MRP_Shortage']:g} יח'"
                        })

                return strict_issues

            strict_issues_list = get_strict_bom_issues(wip_asm_choice)

            if strict_issues_list:
                st.error(f"❌ שגיאה: לא ניתן להכניס את ההרכבה `{wip_asm_choice}` ל-WIP מכיוון שרכיבי הבן **הישירים והאמיתיים** שלה חסרים או באיחור!")
                st.markdown("**רכיבי הבן החסרים בעץ של הרכבה זו בלבד:**")
                for item in strict_issues_list:
                    st.markdown(f"- מק'ט: `{item['PN']}` | תיאור: {item['Description']} | סיבה: {item['Reason']}")
                st.warning("💡 עדכן את המלאי או ה-ETA של רכיבים ספציפיים אלו בלשונית 'עדכון מלאי וספקים'.")
            else:
                save_wip_record(wip_asm_choice, wip_qty_input)
                st.success(f"✅ בדיקת העץ הקפדנית עברה בהצלחה! כל הבנים הישירים של `{wip_asm_choice}` זמינים. ה-WIP נשמר!")
                st.rerun()

    st.markdown("##### 📋 רשימת ההרכבות הפעילות ב-WIP כרגע:")
    if wip_current:
        wip_display_rows = []
        for asm_k, qty_v in wip_current.items():
            if qty_v > 0:
                wip_display_rows.append({
                    "קוד הרכבה": asm_k,
                    "תיאור הרכבה": assembly_mapping.get(asm_k, ""),
                    "כמות ב-WIP": qty_v
                })
        if wip_display_rows:
            st.dataframe(pd.DataFrame(wip_display_rows), use_container_width=True)
            if st.button("🗑️ איפוס כל ה-WIP"):
                for asm_k in list(wip_current.keys()):
                    delete_wip_record(asm_k)
                st.success("כל נתוני ה-WIP אופסו.")
                st.rerun()
        else:
            st.info("אין כרגע הרכבות פעילות ב-WIP.")
    else:
        st.info("אין כרגע הרכבות פעילות ב-WIP.")
