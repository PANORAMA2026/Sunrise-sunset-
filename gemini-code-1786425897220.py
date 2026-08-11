import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from astral import LocationInfo
from astral.sun import sun
import datetime
import re

# Configurazione della pagina Streamlit
st.set_page_config(page_title="Calcolatore Alba/Tramonto Rotta Navale", layout="wide")

st.title("🚢 Calcolatore Alba & Tramonto per Rotte Navali")
st.write("Carica il file della rotta per visualizzare la mappa interattiva e le effemeridi per ogni punto di passaggio.")

# Sidebar per caricamento file
st.sidebar.header("Carica Dati Rotta")
uploaded_file = st.sidebar.file_uploader("Carica un file Excel (.xlsx, .xls) o CSV (.csv)", type=["xlsx", "xls", "csv"])

# Dati di esempio predefiniti
def get_sample_data():
    return pd.DataFrame({
        'Waypoint': ['Genova', 'Stretto Bonifacio', 'Palermo'],
        'Latitudine': [44.4056, 41.3142, 38.1157],
        'Longitudine': [8.9463, 9.2084, 13.3615],
        'Data_Ora': ['2026-05-10 08:00', '2026-05-10 22:30', '2026-05-11 14:00']
    })

# Funzione per leggere CSV gestendo vari encoding e separatori
def load_csv_safely(file):
    encodings_to_try = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
    for enc in encodings_to_try:
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=enc, sep=None, engine='python')
        except Exception:
            continue
    file.seek(0)
    return pd.read_csv(file, encoding='latin1')

# Funzione per convertire coordinate nautiche o stringhe decimali
def parse_coordinate(val):
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    
    val_str = str(val).strip().upper()
    # Riconosce formati tipo "44° 24.3' N" oppure "44 24.3 N"
    match = re.match(r'([+-]?\d+[\.,]?\d*)\s*°?\s*(\d*[\.,]?\d*)?\s*\'?\s*([NSEW])?', val_str)
    if match:
        deg = float(match.group(1).replace(',', '.'))
        minutes = float(match.group(2).replace(',', '.')) if match.group(2) else 0.0
        direction = match.group(3)
        
        dec = deg + (minutes / 60.0) if deg >= 0 else deg - (minutes / 60.0)
        if direction in ['S', 'W']:
            dec = -abs(dec)
        elif direction in ['N', 'E']:
            dec = abs(dec)
        return dec
    
    try:
        return float(val_str.replace(',', '.'))
    except ValueError:
        return None

# Funzione per normalizzare le colonne del dataframe
def normalize_dataframe(df):
    cols = {str(c).strip().lower(): c for c in df.columns}
    mapping = {}
    
    # Ricerca Latitudine
    for alias in ['lat', 'latitude', 'latitudine', 'lat (deg)']:
        if alias in cols:
            mapping[cols[alias]] = 'Latitudine'
            break
            
    # Ricerca Longitudine
    for alias in ['lon', 'long', 'longitude', 'longitudine', 'lon (deg)']:
        if alias in cols:
            mapping[cols[alias]] = 'Longitudine'
            break

    # Ricerca Waypoint
    for alias in ['waypoint', 'wp', 'name', 'nome', 'point', 'punto', 'station']:
        if alias in cols:
            mapping[cols[alias]] = 'Waypoint'
            break

    # Ricerca Data/Ora
    for alias in ['data_ora', 'datetime', 'date_time', 'eta', 'time', 'date', 'data', 'ora']:
        if alias in cols:
            mapping[cols[alias]] = 'Data_Ora'
            break

    df = df.rename(columns=mapping)

    # Se Date e Time sono separate, prova a unirle
    if 'Data_Ora' not in df.columns:
        date_col = next((c for c in df.columns if c.lower() in ['date', 'data']), None)
        time_col = next((c for c in df.columns if c.lower() in ['time', 'ora']), None)
        if date_col and time_col:
            df['Data_Ora'] = df[date_col].astype(str) + ' ' + df[time_col].astype(str)

    if 'Waypoint' not in df.columns:
        df['Waypoint'] = [f"WP {i+1}" for i in range(len(df))]

    return df

# Caricamento ed elaborazione dati con fallback robusto
if uploaded_file is None:
    st.info("👋 Nessun file caricato. Sto mostrando una rotta di esempio.")
    df = get_sample_data()
else:
    df = None
    
    # 1. Tenta la lettura come CSV
    try:
        uploaded_file.seek(0)
        df = load_csv_safely(uploaded_file)
    except Exception:
        pass

    # 2. Tenta la lettura come file Excel
    if df is None:
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        except Exception:
            try:
                uploaded_file.seek(0)
                df = pd.read_excel(uploaded_file, engine='xlrd')
            except Exception:
                pass

    # 3. Esito del caricamento
    if df is not None:
        df = normalize_dataframe(df)
        st.success(f"File '{uploaded_file.name}' caricato con successo!")
    else:
        st.error("Errore nella lettura del file: impossibile determinare il formato o file danneggiato.")
        st.stop()

# Conversione coordinate e date
if 'Latitudine' in df.columns and 'Longitudine' in df.columns:
    df['Latitudine'] = df['Latitudine'].apply(parse_coordinate)
    df['Longitudine'] = df['Longitudine'].apply(parse_coordinate)

if 'Data_Ora' in df.columns:
    df['Data_Ora'] = pd.to_datetime(df['Data_Ora'], errors='coerce')
else:
    st.error("❌ Impossibile identificare le colonne contenenti data e ora.")
    st.stop()

# Calcolo Alba e Tramonto con Astral
sunrises, sunsets, condizioni = [], [], []

for idx, row in df.iterrows():
    try:
        lat = row['Latitudine']
        lon = row['Longitudine']
        dt = row['Data_Ora']
        
        if pd.isna(lat) or pd.isna(lon) or pd.isna(dt):
            raise ValueError("Dati incompleti")

        city = LocationInfo("Point", "Region", "UTC", lat, lon)
        s = sun(city.observer, date=dt.date())
        
        sr_time = s['sunrise'].strftime('%H:%M')
        ss_time = s['sunset'].strftime('%H:%M')
        
        sunrises.append(sr_time)
        sunsets.append(ss_time)
        
        if s['sunrise'].time() <= dt.time() <= s['sunset'].time():
            condizioni.append("☀️ Giorno")
        else:
            condizioni.append("🌙 Notte")
    except Exception:
        sunrises.append("N/D")
        sunsets.append("N/D")
        condizioni.append("Errore dati")

df['Alba (UTC)'] = sunrises
df['Tramonto (UTC)'] = sunsets
df['Luce'] = condizioni

# Layout del dashboard Streamlit
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📍 Tabella Dati ed Effemeridi")
    st.dataframe(df, use_container_width=True)
    
    csv_data = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Scarica Tabella Elaborata (CSV)",
        data=csv_data,
        file_name='rotta_effemeridi_calcolate.csv',
        mime='text/csv',
    )

with col2:
    st.subheader("🗺️ Mappa Interattiva della Rotta")
    
    valid_points = df.dropna(subset=['Latitudine', 'Longitudine'])
    if not valid_points.empty:
        try:
            center_lat = valid_points['Latitudine'].mean()
            center_lon = valid_points['Longitudine'].mean()
            m = folium.Map(location=[center_lat, center_lon], zoom_start=5)
            
            points = list(zip(valid_points['Latitudine'], valid_points['Longitudine']))
            folium.PolyLine(points, color="blue", weight=3, opacity=0.8).add_to(m)
            
            for idx, row in valid_points.iterrows():
                nome_wp = row.get('Waypoint', f'Punto {idx+1}')
                popup_text = f"""
                <b>{nome_wp}</b><br>
                Ora passaggio: {row['Data_Ora']}<br>
                Alba (UTC): {row['Alba (UTC)']}<br>
                Tramonto (UTC): {row['Tramonto (UTC)']}<br>
                Stato: {row['Luce']}
                """
                color = 'orange' if 'Giorno' in str(row['Luce']) else 'darkblue'
                folium.Marker(
                    location=[row['Latitudine'], row['Longitudine']],
                    popup=popup_text,
                    tooltip=f"{nome_wp} ({row['Luce']})",
                    icon=folium.Icon(color=color, icon='ship', prefix='fa')
                ).add_to(m)
                
            st_folium(m, width=600, height=500)
        except Exception as e:
            st.warning(f"Non è stato possibile generare la mappa: {e}")
    else:
        st.warning("Nessuna coordinata valida trovata per generare la mappa.")
