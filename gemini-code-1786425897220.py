import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from astral import LocationInfo
from astral.sun import sun
import datetime

st.set_page_config(page_title="Calcolatore Alba/Tramonto Rotta", layout="wide")

st.title("🚢 Calcolatore Alba & Tramonto per Rotte Navali")
st.write("Carica il file della rotta per visualizzare la mappa interattiva e le effemeridi per ogni punto di passaggio.")

# Sidebar per caricamento file
st.sidebar.header("Carica Dati")
uploaded_file = st.sidebar.file_uploader("Carica un file Excel o CSV", type=["xlsx", "csv"])

# Template scaricabile
def get_sample_data():
    return pd.DataFrame({
        'Waypoint': ['Genova', 'Stretto Bonifacio', 'Palermo'],
        'Latitudine': [44.4056, 41.3142, 38.1157],
        'Longitudine': [8.9463, 9.2084, 13.3615],
        'Data_Ora': ['2026-05-10 08:00', '2026-05-10 22:30', '2026-05-11 14:00']
    })

if uploaded_file is None:
    st.info("👋 Nessun file caricato. Sto usando una rotta di esempio.")
    df = get_sample_data()
else:
    if uploaded_file.name.endswith('.csv'):
    try:
        df = pd.read_csv(uploaded_file, encoding='utf-8')
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding='latin1')
else:
    df = pd.read_excel(uploaded_file)

# Conversione date
df['Data_Ora'] = pd.to_datetime(df['Data_Ora'])

# Calcolo Effemeridi
sunrises, sunsets, condizioni = [], [], []

for idx, row in df.iterrows():
    lat, lon, dt = row['Latitudine'], row['Longitudine'], row['Data_Ora']
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

df['Alba (UTC)'] = sunrises
df['Tramonto (UTC)'] = sunsets
df['Luce'] = condizioni

# Visualizzazione Layout a due colonne
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📍 Dati Rotta ed Effemeridi")
    st.dataframe(df, use_container_width=True)
    
    # Download del risultato
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Scarica Dati Elaborati (CSV)",
        data=csv,
        file_name='rotta_effemeridi.csv',
        mime='text/csv',
    )

with col2:
    st.subheader("🗺️ Mappe Interattiva della Rotta")
    
    # Centro mappa
    center_lat = df['Latitudine'].mean()
    center_lon = df['Longitudine'].mean()
    m = folium.Map(location=[center_lat, center_lon], zoom_start=6)
    
    # Traccia linea rotta
    points = list(zip(df['Latitudine'], df['Longitudine']))
    folium.PolyLine(points, color="blue", weight=3, opacity=0.8).add_to(m)
    
    # Marker per ogni waypoint
    for idx, row in df.iterrows():
        popup_text = f"""
        <b>{row.get('Waypoint', f'Punto {idx+1}')}</b><br>
        Ora passaggio: {row['Data_Ora']}<br>
        Alba: {row['Alba (UTC)']}<br>
        Tramonto: {row['Tramonto (UTC)']}<br>
        Stato: {row['Luce']}
        """
        color = 'orange' if 'Giorno' in row['Luce'] else 'darkblue'
        folium.Marker(
            location=[row['Latitudine'], row['Longitudine']],
            popup=popup_text,
            tooltip=f"{row.get('Waypoint', idx+1)} ({row['Luce']})",
            icon=folium.Icon(color=color, icon='ship', prefix='fa')
        ).add_to(m)
        
    st_folium(m, width=600, height=500)
