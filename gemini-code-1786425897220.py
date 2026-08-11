import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from astral import LocationInfo
from astral.sun import sun
import datetime

# Configurazione della pagina Streamlit
st.set_page_config(page_title="Calcolatore Alba/Tramonto Rotta Navale", layout="wide")

st.title("🚢 Calcolatore Alba & Tramonto per Rotte Navali")
st.write("Carica il file della rotta per visualizzare la mappa interattiva e le effemeridi per ogni punto di passaggio.")

# Sidebar per caricamento file
st.sidebar.header("Carica Dati Rotta")
uploaded_file = st.sidebar.file_uploader("Carica un file Excel (.xlsx) o CSV (.csv)", type=["xlsx", "csv"])

# Dati di esempio predefiniti se non viene caricato alcun file
def get_sample_data():
    return pd.DataFrame({
        'Waypoint': ['Genova', 'Stretto Bonifacio', 'Palermo'],
        'Latitudine': [44.4056, 41.3142, 38.1157],
        'Longitudine': [8.9463, 9.2084, 13.3615],
        'Data_Ora': ['2026-05-10 08:00', '2026-05-10 22:30', '2026-05-11 14:00']
    })

# Funzione per leggere CSV gestendo gli errori di codifica/encoding
def load_csv_safely(file):
    encodings_to_try = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
    for enc in encodings_to_try:
        try:
            file.seek(0)
            # engine='python' e sep=None rilevano in automatico virgola o punto e virgola
            return pd.read_csv(file, encoding=enc, sep=None, engine='python')
        except Exception:
            continue
    # Se i tentativi precedenti falliscono, prova la lettura standard con fallback
    file.seek(0)
    return pd.read_csv(file, encoding='latin1')

# Caricamento ed elaborazione dati
if uploaded_file is None:
    st.info("👋 Nessun file caricato. Sto mostrando una rotta di esempio.")
    df = get_sample_data()
else:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = load_csv_safely(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.success(f"File '{uploaded_file.name}' caricato con successo!")
    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")
        st.stop()

# Pulizia e conversione date
try:
    df['Data_Ora'] = pd.to_datetime(df['Data_Ora'])
except Exception:
    st.error("Assicurati che la colonna 'Data_Ora' contenga date valide (es. YYYY-MM-DD HH:MM).")

# Calcolo Alba e Tramonto con Astral
sunrises, sunsets, condizioni = [], [], []

for idx, row in df.iterrows():
    try:
        lat = float(row['Latitudine'])
        lon = float(row['Longitudine'])
        dt = row['Data_Ora']
        
        city = LocationInfo("Point", "Region", "UTC", lat, lon)
        s = sun(city.observer, date=dt.date())
        
        sr_time = s['sunrise'].strftime('%H:%M')
        ss_time = s['sunset'].strftime('%H:%M')
        
        sunrises.append(sr_time)
        sunsets.append(ss_time)
        
        # Verifico se la nave si trova nell'intervallo di luce solare
        if s['sunrise'].time() <= dt.time() <= s['sunset'].time():
            condizioni.append("☀️ Giorno")
        else:
            condizioni.append("🌙 Notte")
    except Exception as e:
        sunrises.append("N/D")
        sunsets.append("N/D")
        condizioni.append("Errore dati")

# Aggiunta colonne elaborate al DataFrame
df['Alba (UTC)'] = sunrises
df['Tramonto (UTC)'] = sunsets
df['Luce'] = condizioni

# Visualizzazione del Layout a due colonne
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📍 Tabella Dati ed Effemeridi")
    st.dataframe(df, use_container_width=True)
    
    # Pulsante di Download
    csv_data = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Scarica Tabella Elaborata (CSV)",
        data=csv_data,
        file_name='rotta_effemeridi_calcolate.csv',
        mime='text/csv',
    )

with col2:
    st.subheader("🗺️ Mappa Interattiva della Rotta")
    
    try:
        # Calcolo centro della mappa
        center_lat = df['Latitudine'].astype(float).mean()
        center_lon = df['Longitudine'].astype(float).mean()
        m = folium.Map(location=[center_lat, center_lon], zoom_start=5)
        
        # Traccia linea rotta navale
        points = list(zip(df['Latitudine'].astype(float), df['Longitudine'].astype(float)))
        folium.PolyLine(points, color="blue", weight=3, opacity=0.8).add_to(m)
        
        # Marker per ciascun Waypoint
        for idx, row in df.iterrows():
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
                location=[float(row['Latitudine']), float(row['Longitudine'])],
                popup=popup_text,
                tooltip=f"{nome_wp} ({row['Luce']})",
                icon=folium.Icon(color=color, icon='ship', prefix='fa')
            ).add_to(m)
            
        st_folium(m, width=600, height=500)
    except Exception as e:
        st.warning(f"Non è stato possibile generare la mappa: {e}")
