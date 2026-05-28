        buf = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        f_cfg = {"Low": ["Starlette", "Stella", "Lares"], "Meo": ["Lageos 1", "Lageos 2", "Lares 2"], "High": ["Etalon 1", "Etalon 2", "Galileo-Bridge"]}
        f_hd, f_dt, f_vrd = Font(name="Calibri", size=10, bold=True), Font(name="Calibri", size=10), PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        f_grigio = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        brd, al_c = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')), Alignment(horizontal="center", vertical="center")
        
        for f_nm, s_lst in f_cfg.items():
            ws = wb.create_sheet(title=f_nm)
            ws.sheet_view.showGridLines = True
            ws.append([])
            c_idx = 2
            for s in s_lst:
                ws.merge_cells(start_row=2, start_column=c_idx, end_row=2, end_column=c_idx+9)
                
                display_name = "galileo-xxx" if s == "Galileo-Bridge" else s.lower().replace(" ", "")
                cell_s = ws.cell(row=2, column=c_idx, value=display_name)
                cell_s.font = f_hd; cell_s.alignment = al_c; cell_s.fill = f_grigio
                
                ws.merge_cells(start_row=3, start_column=c_idx, end_row=3, end_column=c_idx+9)
                orb_k = "Bassa (LEO)" if f_nm=="Low" else ("Media (MEO)" if f_nm=="Meo" else "Alta (HEO/GEO)")
                sic_label = "71xx" if s == "Galileo-Bridge" else sat_info[orb_k].get(s, "0000")
                ws.cell(row=3, column=c_idx, value=sic_label).font = f_dt
                ws.cell(row=3, column=c_idx).alignment = al_c; ws.cell(row=3, column=c_idx).fill = f_grigio
                
                cell_d_title = ws.cell(row=4, column=c_idx, value="Data")
                cell_d_title.font = f_hd; cell_d_title.alignment = al_c; cell_d_title.fill = f_grigio
                
                ws.merge_cells(start_row=4, start_column=c_idx+1, end_row=4, end_column=c_idx+4)
                cell_ms = ws.cell(row=4, column=c_idx+1, value="MSLR")
                cell_ms.font = f_hd; cell_ms.alignment = al_c; cell_ms.fill = f_grigio
                
                ws.merge_cells(start_row=4, start_column=c_idx+5, end_row=4, end_column=c_idx+9)
                cell_ml = ws.cell(row=4, column=c_idx+5, value="MLRO")
                cell_ml.font = f_hd; cell_ml.alignment = al_c; cell_ml.fill = f_grigio
                
                for r_h in range(2, 5):
                    for c_h in range(c_idx, c_idx+10): ws.cell(row=r_h, column=c_h).border = brd
                
                ws.cell(row=5, column=c_idx, value="").font = f_hd
                ws.cell(row=5, column=c_idx).fill = f_grigio; ws.cell(row=5, column=c_idx).border = brd
                
                for p_idx, p_tx in enumerate(["Time start", "Rms", "NP", ""]):
                    cell_p1 = ws.cell(row=5, column=c_idx + 1 + p_idx, value=p_tx)
                    cell_p1.font = f_hd; cell_p1.alignment = al_c; cell_p1.fill = f_grigio; cell_p1.border = brd
                    
                    cell_p2 = ws.cell(row=5, column=c_idx + 5 + p_idx, value=p_tx)
                    cell_p2.font = f_hd; cell_p2.alignment = al_c; cell_p2.fill = f_grigio; cell_p2.border = brd
                
                cell_n_title = ws.cell(row=5, column=c_idx + 9, value="Note")
                cell_n_title.font = f_hd; cell_n_title.alignment = al_c; cell_n_title.fill = f_grigio; cell_n_title.border = brd
                
                c_idx += 10
            
            r_dest = 6
            if not df_v.empty:
                for row in df_v.to_dict(orient="records"):
                    is_galileo = row["Satellite"].startswith("Galileo-")
                    if f_nm == "High" and not is_galileo and row["Satellite"] not in s_lst: continue
                    if f_nm == "High" and is_galileo and "Galileo-Bridge" not in s_lst: continue
                    if f_nm != "High" and row["Satellite"] not in s_lst: continue
                    
                    s_pos = s_lst.index("Galileo-Bridge") if is_galileo and f_nm == "High" else s_lst.index(row["Satellite"])
                    b_col = 2 + (s_pos * 10)
                    
                    cell_d = ws.cell(row=r_dest, column=b_col, value=row["Data"])
                    cell_d.font = f_dt; cell_d.fill = f_vrd; cell_d.alignment = al_c
                    
                    if row["T_MSLR"]:
                        dt_ms_obj = datetime.strptime(row["T_MSLR"], "%Y-%m-%d %H:%M:%S")
                        t_ms = f"{dt_ms_obj.hour}.{dt_ms_obj.strftime('%M')}"
                        vals = [t_ms, row["R_MS"], row["N_MS"], row["fr2_ms_old"]]
                        for o_idx, val in enumerate(vals):
                            cell = ws.cell(row=r_dest, column=b_col + 1 + o_idx, value=val)
                            cell.font = f_dt; cell.fill = f_vrd; cell.alignment = al_c
                    if row["T_MLRO"]:
                        dt_mo_obj = datetime.strptime(row["T_MLRO"], "%Y-%m-%d %H:%M:%S")
                        t_ml = f"{dt_mo_obj.hour}.{dt_mo_obj.strftime('%M')}"
                        vals = [t_ml, row["R_MO"], row["N_MO"], row["fr2"]]
                        for o_idx, val in enumerate(vals):
                            cell = ws.cell(row=r_dest, column=b_col + 5 + o_idx, value=val)
                            cell.font = f_dt; cell.fill = f_vrd; cell.alignment = al_c
                        
                    cell_nt = ws.cell(row=r_dest, column=b_col + 9, value=row["Note_Accoppiate"])
                    cell_nt.font = f_dt; cell_nt.fill = f_vrd; cell_nt.alignment = al_c
                    
                    for c_b in range(2, 32):
                        ws.cell(row=r_dest, column=c_b).border = brd
                    r_dest += 1
            
            for r_empty in range(r_dest, 31):
                for c_b in range(2, 32):
                    ws.cell(row=r_empty, column=c_b).border = brd
            
            for col in range(1, 32):
                col_letter = get_column_letter(col)
                max_len = max(len(str(ws.cell(row_idx, column=col).value or '')) for row_idx in range(1, 31))
                ws.column_dimensions[col_letter].width = max(max_len + 2, 11)
                
        wb.save(buf)
        st.write("")
        st.download_button(label="📥 Scarica Registro Strutturato (.xlsx)", data=buf.getvalue(), file_name=f"satelliti_collocazione_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M')}.xlsx", width="stretch")
        st.write("---")
        
        df_mslr_raw = df[df["sistema"] == "MSLR"]
        st.markdown("### 🔴 Dati MSLR")
        if not df_mslr_raw.empty: st.dataframe(df_mslr_raw[["id", "orbita", "satellite", "sic_code", "data_ora", "rms", "normal_point", "note"]], width="stretch", hide_index=True)
        
        df_mlro_raw = df[df["sistema"] == "MLRO"]
        st.markdown("### 🔵 Dati MLRO")
        if not df_mlro_raw.empty: st.dataframe(df_mlro_raw[["id", "orbita", "satellite", "sic_code", "data_ora", "rms", "normal_point", "note"]], width="stretch", hide_index=True)
    else: st.info("Nessun dato ancora registrato.")
