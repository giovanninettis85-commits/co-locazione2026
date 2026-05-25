import streamlit as st
import pandas as pd
import sqlite3, io, openpyxl, os
from datetime import datetime, time, timedelta, timezone
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Laser Ranging Tracking", layout="centered")

col1, _, col2 = st.columns(3)
with col1:
    if os.path.exists("logo_asi.png"): st.image("logo_asi.png", width=120)
with col2:
    if os.path.exists("logo_egeos.png"): st.image("logo_egeos.png", width=130)

def q(sql, p=()):
    with sqlite3.connect("laser_data_v2.db") as c:
        cursor = c.cursor()
        cursor.execute(sql, p)
        c.commit()
        return cursor.fetchall()

q("""CREATE TABLE IF NOT EXISTS acquisitions (id INTEGER PRIMARY KEY AUTOINCREMENT,
     sistema TEXT, orbita TEXT, satellite TEXT, sic_code TEXT, data_ora TEXT, rms REAL, normal_point INTEGER)""")

sat_info = {
    "Bassa (LEO)": {"Starlette": "1134", "Stella": "0643", "Lares": "5987"},
    "Media (MEO)": {"Lageos 1": "1155", "Lageos 2": "5986", "Lares 2": "5988"},
    "Alta (HEO/GEO)": {
        "Etalon 1": "0525", 
        "Etalon 2": "4146",
        "Galileo-101": "7101", 
        "Galileo-102": "7102", 
        "Galileo-201": "7201", 
        "Galileo-202": "7202", 
        "Galileo-209": "7209", 
        "Galileo-211": "7211"
    }
}

st.title("🛰️ Laser Ranging Data Sync")
t_mslr, t_mlro, t_dati = st.tabs(["🔴 MSLR", "🔵 MLRO", "📊 Registro Dati"])

def f_form(sys):
    st.subheader(f"Nuovo inserimento {sys}")
    orb = st.selectbox("Seleziona Orbita", list(sat_info.keys()), key=f"o_{sys}")
    sat_name = st.selectbox("Seleziona Satellite", list(sat_info[orb].keys()), key=f"s_{sys}")
    sic = sat_info[orb][sat_name]
    dt = datetime.now(timezone.utc)
    c1, c2 = st.columns(2)
    with c1: d = st.date_input("Data (UTC)", value=dt.date(), key=f"d_{sys}")
    
    # MODIFICA: Inserito step=60 per sbloccare la selezione dell'orario minuto per minuto
    with c2: t = st.time_input("Ora (UTC)", value=dt.time().replace(second=0, microsecond=0), step=60, key=f"t_{sys}")
    
    rms = st.number_input("RMS (mm)", min_value=0.0, value=1.0, step=0.1, format="%.1f", key=f"r_{sys}")
    np = st.number_input("Normal Point", min_value=1, value=15, step=1, key=f"n_{sys}")
    
    if f"saved_{sys}" not in st.session_state:
        st.session_state[f"saved_{sys}"] = False
        
    if st.button(f"💾 Salva in {sys}", key=f"b_{sys}", type="primary", width="stretch"):
        dt_c = datetime.combine(d, t).strftime("%Y-%m-%d %H:%M:%S")
        q("INSERT INTO acquisitions VALUES (NULL,?,?,?,?,?,?,?)", (sys, orb, sat_name, sic, dt_c, round(rms, 1), np))
        st.session_state[f"saved_{sys}"] = True
        st.rerun()
        
    if st.session_state[f"saved_{sys}"]:
        st.success("Salvato")
        st.session_state[f"saved_{sys}"] = False

with t_mslr: f_form("MSLR")
with t_mlro: f_form("MLRO")
with t_dati:
    st.subheader("📋 Gestione Registro ed Obiettivi")
    raw = q("SELECT id, sistema, orbita, satellite, sic_code, data_ora, rms, normal_point FROM acquisitions ORDER BY id DESC")
    df = pd.DataFrame(raw, columns=["id", "sistema", "orbita", "satellite", "sic_code", "data_ora", "rms", "normal_point"]) if raw else pd.DataFrame()
    valid, c_leo, c_meo, c_heo = [], 0, 0, 0
    
    if not df.empty:
        mslr, mlro = df[df["sistema"] == "MSLR"], df[df["sistema"] == "MLRO"]
        u_mlro = set()
        for _, r_ms in mslr.iterrows():
            dt_ms = datetime.strptime(r_ms["data_ora"], "%Y-%m-%d %H:%M:%S")
            for _, r_mo in mlro.iterrows():
                if r_mo["id"] in u_mlro or r_ms["satellite"] != r_mo["satellite"]: continue
                dt_mo = datetime.strptime(r_mo["data_ora"], "%Y-%m-%d %H:%M:%S")
                if abs(dt_ms - dt_mo) <= timedelta(minutes=5):
                    u_mlro.add(r_mo["id"])
                    
                    fr2_name = f"9991_{r_mo['satellite'].lower().replace(' ','')}_crd_{dt_mo.strftime('%Y%m%d_%H%M')}_00.fr2"
                    fr2_ms_old = f"{dt_ms.strftime('%Y%m%d_%H%M')}_{r_ms['satellite'].lower().replace(' ','')}_{r_ms['sic_code']}.fr2"
                    
                    valid.append({
                        "Data": dt_ms.strftime("%Y-%m-%d"),
                        "fr2": fr2_name, 
                        "fr2_ms_old": fr2_ms_old,
                        "Satellite": r_ms["satellite"], 
                        "SIC": r_ms["sic_code"], 
                        "Orbita": r_ms["orbita"], 
                        "T_MSLR": r_ms["data_ora"], 
                        "R_MS": round(r_ms["rms"], 1), 
                        "N_MS": r_ms["normal_point"], 
                        "T_MLRO": r_mo["data_ora"], 
                        "R_MO": round(r_mo["rms"], 1), 
                        "N_MO": r_mo["normal_point"]
                    })
                    break
        df_v = pd.DataFrame(valid) if valid else pd.DataFrame()
        if not df_v.empty:
            c_leo, c_meo, c_heo = len(df_v[df_v["Orbita"] == "Bassa (LEO)"]), len(df_v[df_v["Orbita"] == "Media (MEO)"]), len(df_v[df_v["Orbita"] == "Alta (HEO/GEO)"])
    else: df_v = pd.DataFrame()
    
    cx1, cx2, cx3 = st.columns(3)
    with cx1: st.markdown(f"🟢 **Low**: {c_leo}/20"); st.progress(min(c_leo/20, 1.0))
    with cx2: st.markdown(f"🟡 **Meo**: {c_meo}/20"); st.progress(min(c_meo/20, 1.0))
    with cx3: st.markdown(f"🔴 **High**: {c_heo}/20"); st.progress(min(c_heo/20, 1.0))
    st.write("---")
    
    if not df.empty:
        st.markdown("### 🌟 Tabella Sincronizzata")
        if not df_v.empty:
            st.dataframe(df_v[["Data", "fr2", "Satellite", "SIC", "Orbita", "T_MSLR", "R_MS", "N_MS", "T_MLRO", "R_MO", "N_MO"]], width="stretch", hide_index=True)
        else: st.warning("⚠️ Nessun passaggio accoppiato entro i 5 minuti.")
        with st.expander("🗑️ Elimina record"):
            opt = {r["id"]: f"ID {r['id']} - {r['satellite']} ({r['sistema']}) del {r['data_ora']}" for _, r in df.iterrows()}
            sel = st.selectbox("Seleziona riga:", list(opt.keys()), format_func=lambda x: opt[x])
            if st.button("🚨 Nel Registro Elimina Riga", width="stretch"):
                q("DELETE FROM acquisitions WHERE id = ?", (sel,)); st.rerun()
                
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
                ws.merge_cells(start_row=2, start_column=c_idx, end_row=2, end_column=c_idx+8)
                
                display_name = "galileo-xxx" if s == "Galileo-Bridge" else s.lower().replace(" ", "")
                cell_s = ws.cell(row=2, column=c_idx, value=display_name)
                cell_s.font = f_hd; cell_s.alignment = al_c; cell_s.fill = f_grigio
                
                ws.merge_cells(start_row=3, start_column=c_idx, end_row=3, end_column=c_idx+8)
                orb_k = "Bassa (LEO)" if f_nm=="Low" else ("Media (MEO)" if f_nm=="Meo" else "Alta (HEO/GEO)")
                sic_label = "71xx" if s == "Galileo-Bridge" else sat_info[orb_k].get(s, "0000")
                ws.cell(row=3, column=c_idx, value=sic_label).font = f_dt
                ws.cell(row=3, column=c_idx).alignment = al_c; ws.cell(row=3, column=c_idx).fill = f_grigio
                
                cell_d_title = ws.cell(row=4, column=c_idx, value="Data")
                cell_d_title.font = f_hd; cell_d_title.alignment = al_c; cell_d_title.fill = f_grigio
                
                ws.merge_cells(start_row=4, start_column=c_idx+1, end_row=4, end_column=c_idx+4)
                cell_ms = ws.cell(row=4, column=c_idx+1, value="MSLR")
                cell_ms.font = f_hd; cell_ms.alignment = al_c; cell_ms.fill = f_grigio
                
                ws.merge_cells(start_row=4, start_column=c_idx+5, end_row=4, end_column=c_idx+8)
                cell_ml = ws.cell(row=4, column=c_idx+5, value="MLRO")
                cell_ml.font = f_hd; cell_ml.alignment = al_c; cell_ml.fill = f_grigio
                
                for r_h in range(2, 5):
                    for c_h in range(c_idx, c_idx+9): ws.cell(row=r_h, column=c_h).border = brd
                
                ws.cell(row=5, column=c_idx, value="").font = f_hd
                ws.cell(row=5, column=c_idx).fill = f_grigio; ws.cell(row=5, column=c_idx).border = brd
                
                for p_idx, p_tx in enumerate(["Time start", "Rms", "NP", ""]):
                    cell_p1 = ws.cell(row=5, column=c_idx + 1 + p_idx, value=p_tx)
                    cell_p1.font = f_hd; cell_p1.alignment = al_c; cell_p1.fill = f_grigio; cell_p1.border = brd
                    
                    cell_p2 = ws.cell(row=5, column=c_idx + 5 + p_idx, value=p_tx)
                    cell_p2.font = f_hd; cell_p2.alignment = al_c; cell_p2.fill = f_grigio; cell_p2.border = brd
                c_idx += 9
            
            r_dest = 6
            if not df_v.empty:
                for row in df_v.to_dict(orient="records"):
                    is_galileo = row["Satellite"].startswith("Galileo-")
                    if f_nm == "High" and not is_galileo and row["Satellite"] not in s_lst: continue
                    if f_nm == "High" and is_galileo and "Galileo-Bridge" not in s_lst: continue
                    if f_nm != "High" and row["Satellite"] not in s_lst: continue
                    
                    s_pos = s_lst.index("Galileo-Bridge") if is_galileo and f_nm == "High" else s_lst.index(row["Satellite"])
                    b_col = 2 + (s_pos * 9)
                    dt_ms_obj = datetime.strptime(row["T_MSLR"], "%Y-%m-%d %H:%M:%S")
                    dt_mo_obj = datetime.strptime(row["T_MLRO"], "%Y-%m-%d %H:%M:%S")
                    
                    t_ms = f"{dt_ms_obj.hour}.{dt_ms_obj.strftime('%M')}"
                    t_ml = f"{dt_mo_obj.hour}.{dt_mo_obj.strftime('%M')}"
                    
                    cell_d = ws.cell(row=r_dest, column=b_col, value=row["Data"])
                    cell_d.font = f_dt; cell_d.fill = f_vrd; cell_d.alignment = al_c
                    
                    for o_idx, val in enumerate([t_ms, row["R_MS"], row["N_MS"], row["fr2_ms_old"]]):
                        cell = ws.cell(row=r_dest, column=b_col + 1 + o_idx, value=val)
                        cell.font = f_dt; cell.fill = f_vrd; cell.alignment = al_c
                    for o_idx, val in enumerate([t_ml, row["R_MO"], row["N_MO"], row["fr2"]]):
                        cell = ws.cell(row=r_dest, column=b_col + 5 + o_idx, value=val)
                        cell.font = f_dt; cell.fill = f_vrd; cell.alignment = al_c
                    for c_b in range(2, 29):
                        ws.cell(row=r_dest, column=c_b).border = brd
                    r_dest += 1
            
            for r_empty in range(r_dest, 31):
                for c_b in range(2, 29):
                    ws.cell(row=r_empty, column=c_b).border = brd
                    
            for col in range(1, 29):
                col_letter = get_column_letter(col)
                max_len = max(len(str(ws.cell(row=row, column=col).value or '')) for row in range(1, 31))
                ws.column_dimensions[col_letter].width = max(max_len + 2, 11)
        wb.save(buf)
        st.write("")
        st.download_button(label="📥 Scarica Registro Strutturato (.xlsx)", data=buf.getvalue(), file_name=f"satelliti_collocazione_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M')}.xlsx", width="stretch")
        st.write("---")
        
        df_mslr_raw = df[df["sistema"] == "MSLR"]
        st.markdown("### 🔴 Dati MSLR")
        if not df_mslr_raw.empty: st.dataframe(df_mslr_raw, width="stretch", hide_index=True)
        
        df_mlro_raw = df[df["sistema"] == "MLRO"]
        st.markdown("### 🔵 Dati MLRO")
        if not df_mlro_raw.empty: st.dataframe(df_mlro_raw, width="stretch", hide_index=True)
    else: st.info("Nessun dato ancora registrato.")
