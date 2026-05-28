import streamlit as st
import pandas as pd
import sqlite3, io, openpyxl, os, base64
from datetime import datetime, time, timedelta, timezone
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Laser Ranging Tracking", layout="centered")

# SOLUZIONE CONTENITORI CSS PER ENTI ISTITUZIONALI
col1, _, col2 = st.columns(3)
with col1:
    st.markdown('<div style="background-color:#002F6C; color:white; padding:12px; border-radius:6px; text-align:center; font-family:Arial, sans-serif; font-weight:bold; font-size:18px; border-left: 5px solid #00A699;">🛰️ ASI<br><span style="font-size:10px; font-weight:normal; opacity:0.85;">Agenzia Spaziale Italiana</span></div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div style="background-color:#2F2F2F; color:#92D050; padding:12px; border-radius:6px; text-align:center; font-family:Arial, sans-serif; font-weight:bold; font-size:18px; border-right: 5px solid #92D050;">📡 e-GEOS<br><span style="font-size:10px; color:white; font-weight:normal; opacity:0.85;">AN ASI / TELESPAZIO COMPANY</span></div>', unsafe_allow_html=True)

def q(sql, p=()):
    with sqlite3.connect("laser_data_v2.db") as c:
        cursor = c.cursor()
        cursor.execute(sql, p)
        c.commit()
        return cursor.fetchall()

q("""CREATE TABLE IF NOT EXISTS acquisitions (id INTEGER PRIMARY KEY AUTOINCREMENT,
     sistema TEXT, orbita TEXT, satellite TEXT, sic_code TEXT, data_ora TEXT, rms REAL, normal_point INTEGER)""")

try:
    q("ALTER TABLE acquisitions ADD COLUMN note TEXT")
except:
    pass

# --- DISLOCAZIONE PULSANTE DI EMERGENZA NELLA SIDEBAR PER ARCHIVIO FOGLI ---
with st.sidebar:
    st.markdown("### ⚙️ Strumenti Archivio")
    if st.button("🚀 IMPORTA TUTTI I 3 FOGLI ADESSO", type="primary", use_container_width=True):
        tutto = [
            ('MSLR', 'Bassa (LEO)', 'Stella', '0643', '2026-05-12 09:20:00', 2.4, 16, ''),
            ('MLRO', 'Bassa (LEO)', 'Stella', '0643', '2026-05-12 09:20:00', 2.3, 10, 'Session1 process manuale Digos'),
            ('MSLR', 'Bassa (LEO)', 'Stella', '0643', '2026-05-14 10:10:00', 4.1, 4, ''),
            ('MLRO', 'Bassa (LEO)', 'Stella', '0643', '2026-05-14 10:04:00', 3.4, 2, 'Session2 process manuale Digos'),
            ('MSLR', 'Bassa (LEO)', 'Stella', '0643', '2026-05-19 20:51:00', 8.5, 13, ''),
            ('MLRO', 'Bassa (LEO)', 'Stella', '0643', '2026-05-19 20:51:00', 3.2, 5, 'Session1'),
            ('MSLR', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-12 07:48:00', 3.1, 10, ''),
            ('MLRO', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-12 07:48:00', 2.8, 15, 'Session1 process manuale Digos'),
            ('MSLR', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-12 09:37:00', 5.9, 12, ''),
            ('MLRO', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-12 09:37:00', 3.2, 7, 'Session1 process manuale Digos'),
            ('MSLR', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-13 06:37:00', 1.3, 5, ''),
            ('MLRO', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-13 06:37:00', 2.5, 12, 'Session1 process manuale Digos'),
            ('MSLR', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-14 08:26:00', 6.2, 10, ''),
            ('MLRO', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-14 08:27:00', 2.9, 6, 'Session2 process manuale Digos'),
            ('MSLR', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-15 06:45:00', 6.5, 16, ''),
            ('MLRO', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-15 06:45:00', 2.7, 7, 'Definizione 1'),
            ('MSLR', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-21 07:07:00', 6.3, 15, ''),
            ('MLRO', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-21 07:03:00', 3.3, 9, 'session1'),
            ('MSLR', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-22 07:23:00', 5.1, 4, ''),
            ('MLRO', 'Bassa (LEO)', 'Starlette', '1134', '2026-05-22 07:23:00', 2.5, 5, 'session1'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-11 19:39:36', 3.1, 13, ''),
            ('MLRO', 'Bassa (LEO)', 'Lares', '5987', '2026-05-11 19:39:00', 1.8, 10, 'Session3 process manuale Digos')
        ]
        tutto += [
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-14 08:21:15', 5.1, 7, 'sessione2'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-14 16:21:21', 8.2, 5, ''),
            ('MLRO', 'Bassa (LEO)', 'Lares', '5987', '2026-05-14 16:21:21', 2.3, 10, 'Session3 process manuale Digos'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-20 15:56:24', 4.1, 24, 'session5'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-22 07:40:11', 6.2, 11, 'session1'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-22 15:40:39', 3.8, 21, ''),
            ('MLRO', 'Bassa (LEO)', 'Lares', '5987', '2026-05-22 15:40:00', 3.0, 21, 'sessione4'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-25 14:31:00', 14.3, 21, ''),
            ('MLRO', 'Bassa (LEO)', 'Lares', '5987', '2026-05-25 14:32:00', 2.9, 6, 'sessione3'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-28 05:04:00', 6.1, 16, ''),
            ('MLRO', 'Bassa (LEO)', 'Lares', '5987', '2026-05-28 05:04:00', 2.8, 4, 'sessione2'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-28 13:16:16', 13.1, 16, 'sessione 4'),
            ('MSLR', 'Bassa (LEO)', 'Lares', '5987', '2026-05-28 15:15:15', 12.7, 19, ''),
            ('MLRO', 'Bassa (LEO)', 'Lares', '5987', '2026-05-28 15:08:00', 3.5, 21, 'sessione5'),
            ('MSLR', 'Alta (HEO/GEO)', 'Etalon 1', '0525', '2026-05-11 19:44:00', 37.38, 5, 'da verificare riga alto rms'),
            ('MLRO', 'Alta (HEO/GEO)', 'Etalon 1', '0525', '2026-05-11 19:43:00', 11.19, 4, 'Session1 process manuale Digos'),
            ('MSLR', 'Alta (HEO/GEO)', 'Galileo-233', '7101', '2026-05-22 21:48:00', 0.0, 1, 'Mancano dati meteo'),
            ('MLRO', 'Alta (HEO/GEO)', 'Galileo-233', '7101', '2026-05-22 21:35:00', 3.47, 4, 'Mancano dati meteo'),
            ('MSLR', 'Alta (HEO/GEO)', 'Galileo-209', '7209', '2026-05-23 00:45:00', 7.87, 1, 'sessione1'),
            ('MLRO', 'Alta (HEO/GEO)', 'Galileo-209', '7209', '2026-05-23 00:43:00', 4.33, 4, 'sessione1'),
            ('MSLR', 'Alta (HEO/GEO)', 'Galileo-101', '7101', '2026-05-27 23:12:00', 25.21, 2, 'Mlro non acquisito | sessione2'),
            ('MSLR', 'Alta (HEO/GEO)', 'Galileo-101', '7101', '2026-05-28 20:36:00', 11.21, 4, 'MLRO NON ACQUISITO | SESSIONE 8'),
            ('MSLR', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-11 10:11:00', 3.86, 10, ''),
            ('MLRO', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-11 10:05:00', 4.61, 10, 'Mlro output generato'),
            ('MSLR', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-11 11:17:00', 2.37, 8, ''),
            ('MLRO', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-11 11:11:00', 4.41, 10, 'Mlro output generato'),
            ('MSLR', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-11 13:58:00', 5.12, 11, ''),
            ('MLRO', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-11 13:42:00', 4.37, 5, 'Mlro output generato'),
            ('MSLR', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-15 13:16:00', 6.01, 11, ''),
            ('MLRO', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-15 13:16:00', 4.35, 10, 'Mlro output generato'),
            ('MSLR', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-15 15:10:00', 4.10, 4, ''),
            ('MLRO', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-15 15:02:00', 4.31, 10, 'Mlro output generato'),
            ('MSLR', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-20 12:38:00', 6.01, 11, 'Mlro no data'),
            ('MSLR', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-25 15:16:00', 3.00, 10, ''),
            ('MLRO', 'Media (MEO)', 'Lageos 1', '1155', '2026-05-25 15:17:00', 4.02, 17, 'sessione 1'),
            ('MLRO', 'Media (MEO)', 'Lageos 2', '5986', '2026-05-14 05:31:00', 4.04, 15, 'Mlro resulta male'),
            ('MSLR', 'Media (MEO)', 'Lageos 2', '5986', '2026-05-21 07:20:00', 5.01, 10, 'sessione1'),
            ('MLRO', 'Media (MEO)', 'Lageos 2', '5986', '2026-05-21 07:11:00', 3.20, 10, 'sessione1'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-11 14:24:00', 4.06, 18, ''),
            ('MLRO', 'Media (MEO)', 'Lares 2', '5988', '2026-05-11 13:57:00', 3.93, 22, 'Session1 process manuale Digos'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-11 18:15:00', 3.41, 10, ''),
            ('MLRO', 'Media (MEO)', 'Lares 2', '5988', '2026-05-11 18:11:00', 3.89, 11, 'Session3 process manuale Digos'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-20 03:41:00', 4.02, 23, ''),
            ('MLRO', 'Media (MEO)', 'Lares 2', '5988', '2026-05-20 03:38:00', 3.52, 20, ''),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-21 06:01:00', 3.98, 9, ''),
            ('MLRO', 'Media (MEO)', 'Lares 2', '5988', '2026-05-21 06:03:00', 4.05, 19, 'errore npt'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-21 14:31:00', 3.85, 5, ''),
            ('MLRO', 'Media (MEO)', 'Lares 2', '5988', '2026-05-21 14:31:00', 3.85, 5, ''),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-22 15:39:00', 4.02, 16, ''),
            ('MLRO', 'Media (MEO)', 'Lares 2', '5988', '2026-05-22 15:40:00', 4.02, 16, 'sessione 2'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-23 03:20:00', 3.09, 7, ''),
            ('MLRO', 'Media (MEO)', 'Lares 2', '5988', '2026-05-23 03:02:00', 3.80, 2, 'sessione 2'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-25 08:23:00', 3.12, 10, 'Mlro no data | errore npt'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-25 12:41:00', 3.80, 7, 'Mlro no data | errore npt'),
            ('MSLR', 'Media (MEO)', 'Lares 2', '5988', '2026-05-28 11:36:00', 4.12, 4, 'MLRO no data | meteo Sclat')
        ]
        with sqlite3.connect("laser_data_v2.db") as c:
            c.cursor().executemany("INSERT INTO acquisitions VALUES (NULL, ?, ?, ?, ?, ?, ?, ?, ?)", tutto)
            c.commit()
        st.sidebar.success("✅ Fogli Caricati!")
        st.rerun()

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

st.title("🛰️ Laser Ranging Tracking Data Sync")
t_mslr, t_mlro, t_dati = st.tabs(["🔴 MSLR", "🔵 MLRO", "📊 Registro Dati"])

def f_form(sys):
    st.subheader(f"Nuovo inserimento {sys}")
    orb = st.selectbox("Seleziona Orbita", list(sat_info.keys()), key=f"o_{sys}")
    sat_name = st.selectbox("Seleziona Satellite", list(sat_info[orb].keys()), key=f"s_{sys}")
    sic = sat_info[orb][sat_name]
    dt = datetime.now(timezone.utc)
    c1, c2 = st.columns(2)
    with c1: d = st.date_input("Data (UTC)", value=dt.date(), key=f"d_{sys}")
    with c2: t = st.time_input("Ora (UTC)", value=dt.time().replace(second=0, microsecond=0), step=60, key=f"t_{sys}")
    rms = st.number_input("RMS (mm)", min_value=0.0, value=1.0, step=0.1, format="%.1f", key=f"r_{sys}")
    np = st.number_input("Normal Point", min_value=1, value=15, step=1, key=f"n_{sys}")
    
    if f"input_note_{sys}" not in st.session_state:
        st.session_state[f"input_note_{sys}"] = ""
        
    nota = st.text_input("Note", value=st.session_state[f"input_note_{sys}"], key=f"nt_{sys}", placeholder="Inserisci eventuali annotazioni qui...")
    
    if f"saved_{sys}" not in st.session_state:
        st.session_state[f"saved_{sys}"] = False
        
    if st.button(f"💾 Salva in {sys}", key=f"b_{sys}", type="primary", width="stretch"):
        dt_c = datetime.combine(d, t).strftime("%Y-%m-%d %H:%M:%S")
        q("INSERT INTO acquisitions VALUES (NULL,?,?,?,?,?,?,?,?)", (sys, orb, sat_name, sic, dt_c, round(rms, 1), np, nota))
        st.session_state[f"saved_{sys}"] = True
        st.session_state[f"input_note_{sys}"] = ""
        st.rerun()
        
    if st.session_state[f"saved_{sys}"]:
        st.success("Salvato")
        st.session_state[f"saved_{sys}"] = False

with t_mslr: f_form("MSLR")
with t_mlro: f_form("MLRO")
with t_dati:
    st.subheader("📋 Gestione Registro ed Obiettivi")
    raw = q("SELECT id, sistema, orbita, satellite, sic_code, data_ora, rms, normal_point, note FROM acquisitions ORDER BY id DESC")
    df = pd.DataFrame(raw, columns=["id", "sistema", "orbita", "satellite", "sic_code", "data_ora", "rms", "normal_point", "note"]) if raw else pd.DataFrame()
    valid, c_leo, c_meo, c_heo = [], 0, 0, 0
    
    if not df.empty:
        mslr, mlro = df[df["sistema"] == "MSLR"], df[df["sistema"] == "MLRO"]
        u_mlro = set()
        for _, r_ms in mslr.iterrows():
            dt_ms = datetime.strptime(r_ms["data_ora"], "%Y-%m-%d %H:%M:%S")
            for _, r_mo in mlro.iterrows():
                if r_mo["id"] in u_mlro or r_ms["satellite"] != r_mo["satellite"]: continue
                dt_mo = datetime.strptime(r_mo["data_ora"], "%Y-%m-%d %H:%M:%S")
                
                if abs(dt_ms - dt_mo) <= timedelta(minutes=20):
                    u_mlro.add(r_mo["id"])
                    
                    fr2_name = f"7941_{r_mo['satellite'].lower().replace(' ','')}_crd_{dt_mo.strftime('%Y%m%d_%H%M')}_00.fr2"
                    fr2_ms_old = f"9991_{r_ms['satellite'].lower().replace(' ','')}_crd_{dt_ms.strftime('%Y%m%d_%H%M')}_00.fr2"
                    
                    n_ms = r_ms["note"] if r_ms["note"] else ""
                    n_mo = r_mo["note"] if r_mo["note"] else ""
                    if n_ms and n_mo:
                        nota_unita = f"{n_ms} | {n_mo}"
                    elif n_ms:
                        nota_unita = n_ms
                    else:
                        nota_unita = n_mo
                    
                    valid.append({
                        "Data": dt_ms.strftime("%Y-%m-%d"),
                        "fr2_mlro": fr2_name, 
                        "fr2_mslr": fr2_ms_old,
                        "Satellite": r_ms["satellite"], 
                        "SIC": r_ms["sic_code"], 
                        "Orbita": r_ms["orbita"], 
                        "T_MSLR": r_ms["data_ora"], 
                        "R_MS": round(r_ms["rms"], 1), 
                        "N_MS": r_ms["normal_point"], 
                        "T_MLRO": r_mo["data_ora"], 
                        "R_MO": round(r_mo["rms"], 1), 
                        "N_MO": r_mo["normal_point"],
                        "Note_Accoppiate": nota_unita
                    })
                    break
                    
        df_v = pd.DataFrame(valid) if valid else pd.DataFrame()
        c_leo = len(df_v[df_v["Orbita"] == "Bassa (LEO)"]) if not df_v.empty else 0
        c_meo = len(df_v[df_v["Orbita"] == "Media (MEO)"]) if not df_v.empty else 0
        c_heo = len(df_v[df_v["Orbita"] == "Alta (HEO/GEO)"]) if not df_v.empty else 0
    else:
        df_v = pd.DataFrame()
    
    cx1, cx2, cx3 = st.columns(3)
    with cx1: st.markdown(f"🟢 **Low**: {c_leo}/20"); st.progress(min(c_leo/20, 1.0))
    with cx2: st.markdown(f"🟡 **Meo**: {c_meo}/20"); st.progress(min(c_meo/20, 1.0))
    with cx3: st.markdown(f"🔴 **High**: {c_heo}/20"); st.progress(min(c_heo/20, 1.0))
    st.write("---")
    
    if not df.empty:
        st.markdown("### 🌟 Tabella Sincronizzata (Co-locazione)")
        if not df_v.empty:
            st.dataframe(df_v[["Data", "fr2_mslr", "fr2_mlro", "Satellite", "SIC", "Orbita", "T_MSLR", "R_MS", "N_MS", "T_MLRO", "R_MO", "N_MO", "Note_Accoppiate"]], width="stretch", hide_index=True)
        else:
            st.warning("⚠️ Nessun passaggio accoppiato entro i 20 minuti.")
            
        with st.expander("🗑️ Nel Registro Elimina Riga"):
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
                ws.merge_cells(start_row=2, start_column=c_idx, end_row=2, end_column=c_idx+9)
                
                display_name = "galileo-xxx" if s == "Galileo-Bridge" else s.lower().replace(" ", "")
                cell_s = ws.cell(row=2, column=c_idx, value=display_name)
                cell_s.font = f_hd;
