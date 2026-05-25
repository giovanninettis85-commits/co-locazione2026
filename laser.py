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
    "Alta (HEO/GEO)": {"Etalon 1": "0525", "Etalon 2": "4146", "Galileo-101": "7101"}
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
    with c2: t = st.time_input("Ora (UTC)", value=dt.time().replace(second=0, microsecond=0), key=f"t_{sys}")
    rms = st.number_input("RMS (mm)", min_value=0.0, value=1.0, step=0.1, format="%.1f", key=f"r_{sys}")
    np = st.number_input("Normal Point", min_value=1, value=15, step=1, key=f"n_{sys}")
    
    if st.button(f"💾 Salva in {sys}", key=f"b_{sys}", type="primary", width="stretch"):
        dt_c = datetime.combine(d, t).strftime("%Y-%m-%d %H:%M:%S")
        q("INSERT INTO acquisitions VALUES (NULL,?,?,?,?,?,?,?)", (sys, orb, sat_name, sic, dt_c, round(rms, 1), np))
        st.success(f"✅ Satellite {sat_name} salvato con successo!")
        st.rerun()
        
    st.write("---")
    st.markdown(f"### 📋 Ultimi inserimenti {sys}")
    raw_sys = q("SELECT data_ora, satellite, sic_code, rms, normal_point FROM acquisitions WHERE sistema = ? ORDER BY id DESC LIMIT 5", (sys,))
    if raw_sys:
        df_sys = pd.DataFrame(raw_sys, columns=["Data/Ora UTC", "Satellite", "Codice SIC", "RMS (mm)", "Normal Point"])
        st.dataframe(df_sys, use_container_width=True, hide_index=True)
    else:
        st.info(f"Nessun satellite salvato oggi per il sistema {sys}.")

with t_mslr: f_form("MSLR")
with t_mlro: f_form("MLRO")
