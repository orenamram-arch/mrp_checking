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
            # שינוי מדויק לחודש ספציפי בלבד
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
                    del st.session_state["temp_simulated_plan"]
                    st.success("תוכנית הייצור עודכנה ונשמרה בהצלחה במערכת!")
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
