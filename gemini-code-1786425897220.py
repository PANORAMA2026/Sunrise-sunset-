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

# Tentativo di individuazione automatica delle colonne
def auto_detect_columns(df):
    cols = list(df.columns)
    mapping = {}

    for c in cols:
        c_clean = str(c).strip().lower()
        if any(alias in c_clean for alias in ['lat', 'latitude', 'latitudine']) and 'Latitudine' not in mapping.values():
            mapping[c] = 'Latitudine'
        elif any(alias in c_clean for alias in ['lon', 'long', 'longitude', 'longitudine']) and 'Longitudine' not in mapping.values():
            mapping[c] = 'Longitudine'
        elif any(alias in c_clean for alias in ['waypoint', 'wp', 'name', 'nome', 'point', 'punto', 'station', 'wpt']) and 'Waypoint' not in mapping.values():
            mapping[c] = 'Waypoint'
        elif any(alias in c_clean for alias in ['data_ora', 'datetime', 'date_time', 'date/time', 'eta', 'etd', 'utc', 'time', 'date', 'timestamp']) and 'Data_Ora' not in mapping.values():
            mapping[c] = 'Data_Ora'

    return mapping

# Caricamento ed elaborazione dati
if uploaded_file is None:
    st.info("👋 Nessun file caricato. Sto mostrando una rotta di esempio.")
    df = get_sample_data()
else:
    df = None
    try:
        uploaded_file.seek(0)
        df = load_csv_safely(uploaded_file)
    except Exception:
        pass

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

    if df is not None:
        st.success(f"File '{uploaded_file.name}' caricato con successo!")
    else:
        st.error("Errore nella lettura del file: impossibile determinare il formato o file danneggiato.")
        st.stop()

# Mappatura e Selezione Colonne
st.sidebar.subheader("⚙️ Mappatura Colonne")

all_columns = list(df.columns)
detected = auto_detect_columns(df)

# Latitudine
default_lat_idx = all_columns.index(next((k for k, v in detected.items() if v == 'Latitudine'), all_columns[0]))
col_lat = st.sidebar.selectbox("Colonna Latitudine:", all_columns, index=default_lat_idx)

# Longitudine
default_lon_idx = all_columns.index(next((k for k, v in detected.items() if v == 'Longitudine'), all_columns[min(1, len(all_columns)-1)]))
col_lon = st.sidebar.selectbox("Colonna Longitudine:", all_columns, index=default_lon_idx)

# Waypoint
default_wp_idx = all_columns.index(next((k for k, v in detected.items() if v == 'Waypoint'), all_columns[0])) if any(v == 'Waypoint' for v in detected.values()) else None
col_wp = st.sidebar.selectbox("Colonna Nome Waypoint (Opzionale):", ["Genera automatico (WP 1, WP 2...)"] + all_columns, index=0 if default_wp_idx is None else default_wp_idx + 1)

# Data / Ora
dt_options = ["Nessuna nel file (Usa data/ora manuale)"] + all_columns
default_dt_idx = (all_columns.index(next(k for k, v in detected.items() if v == 'Data_Ora')) + 1) if any(v == 'Data_Ora' for v in detected.values()) else 0
col_dt = st.sidebar.selectbox("Colonna Data e Ora:", dt_options, index=default_dt_idx)

# Assegnazione colonne nel DataFrame
df['Latitudine'] = df[col_lat].apply(parse_coordinate)
df['Longitudine'] = df[col_lon].apply(parse_coordinate)

if col_wp != "Genera automatico (WP 1, WP 2...)":
    df['Waypoint'] = df[col_wp]
else:
    df['Waypoint'] = [f"WP {i+1}" for i in range(len(df))]

# Gestione Data / Ora manuale o da colonna
if col_dt != "Nessuna nel file (Usa data/ora manuale)":
    df['Data_Ora'] = pd.to_datetime(df[col_dt], errors='coerce')
else:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Imposta Data/Ora Manuale")
    start_date = st.sidebar.date_input("Data di partenza rotta:", datetime.date.today())
    start_time = st.sidebar.time_input("Ora di partenza (UTC):", datetime.time(8, 0))
    start_dt = datetime.datetime.combine(start_date, start_time)
    
    # Assegna orari progressivi (es. 1 ora di differenza tra waypoint per default)
    hours_step = st.sidebar.number_input("Ore stimati tra ciascun waypoint:", min_value=0.1, value=1.0, step=0.5)
    df['Data_Ora'] = [start_dt + datetime.timedelta(hours=i*hours_step) for i in range(len(df))]

# Verifica presenza date valide
if df['Data_Ora'].isna().all():
    st.error("❌ La colonna selezionata per Data/Ora non contiene valori validi. Seleziona un'altra colonna dal menu a sinistra oppure usa la data/ora manuale.")
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
    
    # Selezione colonne pulite da mostrare
    display_cols = ['Waypoint', 'Latitudine', 'Longitudine', 'Data_Ora', 'Alba (UTC)', 'Tramonto (UTC)', 'Luce']
    st.dataframe(df[display_cols], use_container_width=True)
    
    csv_data = df[display_cols].to_csv(index=False).encode('utf-8')
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
