# ==========================================================
# ETA & SUPPLIER MANAGEMENT (עם שמירה מקומית וגיבוי קבצים)
# ==========================================================
with tab3:
    st.subheader("📅 מעקב ETA וספקים (כולל היסטוריית שינויים ועדכונים)")

    pn_values = sorted(df[PN_COL].dropna().astype(str).unique())
    selected_pn = st.selectbox("בחר מק\"ט לעדכון סטטוס וספק", pn_values)

    existing_rec = get_eta_record(selected_pn)
    def_eta, def_status, def_supplier, def_comment, def_by = (
        (existing_rec[0], existing_rec[1], existing_rec[2], existing_rec[3], existing_rec[4]) 
        if existing_rec else (date.today(), "פתוח", "אופק", "", "")
    )

    with st.form("eta_form"):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            try:
                parsed_eta = pd.to_datetime(def_eta).date() if def_eta else date.today()
            except:
                parsed_eta = date.today()
            eta_date = st.date_input("תאריך הגעה משוער (ETA)", value=parsed_eta)
        with col_f2:
            status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
            status_idx = status_options.index(def_status) if def_status in status_options else 0
            status = st.selectbox("סטטוס", status_options, index=status_idx)
        with col_f3:
            sup_idx = supplier_options.index(def_supplier) if def_supplier in supplier_options else 0
            supplier = st.selectbox("קבלן משנה / ספק", supplier_options, index=sup_idx)

        comment = st.text_area("הערות מעקב", value=def_comment)
        updated_by = st.text_input("עודכן על ידי", value=def_by)
        save_btn = st.form_submit_button("שמור עדכון במערכת ושלח התראה")

    if save_btn:
        save_eta_record(selected_pn, str(eta_date), status, supplier, comment, updated_by, webhook_url)
        st.success("העדכון, ההיסטוריה וההתראה נשמרו בהצלחה במסד הנתונים המקומי!")

    # כפתור גיבוי והורדת קובץ מסד הנתונים למחשב שלך
    st.divider()
    st.markdown("### 💾 גיבוי נתונים")
    with open(LOCAL_DB_FILE, "rb") as db_file:
        db_bytes = db_file.read()
    st.download_button(
        label="📥 הורד גיבוי של מסד הנתונים המקומי (.db)",
        data=db_bytes,
        file_name="eta_updates_backup.db",
        mime="application/octet-stream",
        help="שמור עותק של כל העדכונים והסטטוסים שהוזנו למערכת במחשב שלך."
    )

    st.subheader("🚦 טבלת סטטוסים וספקים שמורים במערכת")
    eta_rows = []
    for pn in pn_values:
        rec = get_eta_record(pn)
        if rec:
            eta_rows.append({
                "מק\"ט": pn,
                "ETA": rec[0],
                "סיכון": eta_color(rec[0]),
                "סטטוס": rec[1],
                "ספק / קב\"מ": rec[2],
                "הערות": rec[3],
                "אחראי": rec[4],
                "תאריך עדכון": rec[5]
            })

    eta_df = pd.DataFrame(eta_rows)
    if len(eta_df) > 0:
        st.dataframe(eta_df, use_container_width=True)
    else:
        st.info("עדיין לא נשמרו עדכונים במערכת.")

    st.subheader(f"📜 היסטוריית שינויים (Audit Trail) עבור מק\"ט: {selected_pn}")
    history_cur = conn.cursor()
    history_cur.execute("SELECT eta, status, supplier, comment, updated_by, updated_at FROM eta_history WHERE pn = ? ORDER BY id DESC", (selected_pn,))
    hist_rows = history_cur.fetchall()
    if hist_rows:
        hist_df = pd.DataFrame(hist_rows, columns=["ETA", "סטטוס", "ספק", "הערות", "עודכן על ידי", "זמן עדכון"])
        st.dataframe(hist_df, use_container_width=True)
    else:
        st.info("אין היסטוריית שינויים קודמת למק\"ט זה.")
