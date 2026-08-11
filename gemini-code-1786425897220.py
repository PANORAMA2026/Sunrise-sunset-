import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from astral import LocationInfo
from astral.sun import sun
import re

# Configurazione della pagina Streamlit
st.set_page_config(page_title="Calcolatore Alba/Tramonto Rotta Navale", layout="wide")

st.title("🚢 Calcolatore Alba & Tramonto per Rotte Navali")
st.write("Carica il file della rotta per visualizzare le effemeridi (UTC e Locale) e la mappa interattiva.")

# Sidebar per caricamento file
st.sidebar.header("Carica Dati Rotta")
uploaded_file = st.sidebar.file_uploader("Carica file CSV (.csv)", type=["csv"])

# Funzione per convertire coordinate nautiche (es. 33° 45.007' N) in float decimale
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

# Funzione per la lettura del file CSV e selezione delle sole colonne richieste
def load_and_process_csv(file):
    encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']
    df = None
    for enc in encodings:
        try:
            file.seek(0)
            df = pd.read_csv(file, encoding=enc, sep=None, engine='python')
            break
        except Exception:
            continue
            
    if df is None:
        raise ValueError("Impossibile leggere il file CSV con gli encoding supportati.")

    # Pulisce i nomi delle colonne rimuovendo caratteri speciali come BOM (ï»¿)
    df.columns = df.columns.str.replace('ï»¿', '').str.strip()

    # Mappatura flessibile per individuare esattamente le 6 colonne richieste
    target_columns = {
        'Waypoint': next((c for c in df.columns if 'waypoint' in c.lower()), None),
        'Name': next((c for c in df.columns if c.lower() == 'name'), None),
        'Latitude': next((c for c in df.columns if 'latitude' in c.lower()), None),
        'Longitude': next((c for c in df.columns if 'longitude' in c.lower()), None),
        'Arrival Time (UTC)': next((c for c in df.columns if 'arrival time (utc)' in c.lower()), None),
        'Arrival Time (Local)': next((c for c in df.columns if 'arrival time (local)' in c.lower()), None),
    }

    # Verifico che le colonne fondamentali (Lat, Lon, UTC) esistano
    missing = [k for k, v in target_columns.items() if v is None and k in ['Latitude', 'Longitude', 'Arrival Time (UTC)']]
    if missing:
        raise KeyError(f"Mancano le seguenti colonne obbligatorie nel CSV: {', '.join(missing)}")

    # Filtro mantenendo solo le colonne rilevate
    selected_cols = {v: k for k, v in target_columns.items() if v is not None}
    df = df[list(selected_cols.keys())].rename(columns=selected_cols)

    return df

# Elaborazione Dati
if uploaded_file is None:
    st.info("👋 Carica il tuo file CSV dalla barra laterale per iniziare.")
else:
    try:
        df = load_and_process_csv(uploaded_file)
        st.success(f"File '{uploaded_file.name}' caricato ed elaborato correttamente!")
        
        # Conversione coordinate
        df['Lat_Decimal'] = df['Latitude'].apply(parse_coordinate)
        df['Lon_Decimal'] = df['Longitude'].apply(parse_coordinate)

        # Conversione date/ore
        df['Arrival Time (UTC)'] = pd.to_datetime(df['Arrival Time (UTC)'], errors='coerce')
        if 'Arrival Time (Local)' in df.columns:
            df['Arrival Time (Local)'] = pd.to_datetime(df['Arrival Time (Local)'], errors='coerce')

        # Calcolo Alba, Tramonto e Condizione Luce (UTC)
        sunrises, sunsets, condizioni = [], [], []

        for idx, row in df.iterrows():
            lat = row['Lat_Decimal']
            lon = row['Lon_Decimal']
            dt = row['Arrival Time (UTC)']
            
            if pd.isna(lat) or pd.isna(lon) or pd.isna(dt):
                sunrises.append("N/D")
                sunsets.append("N/D")
                condizioni.append("Dati Incompleti")
                continue

            try:
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
                condizioni.append("Errore Calcolo")

        df['Alba (UTC)'] = sunrises
        df['Tramonto (UTC)'] = sunsets
        df['Luce'] = condizioni

        # Organizzazione colonne per la visualizzazione finale
        output_cols = []
        for col in ['Waypoint', 'Name', 'Latitude', 'Longitude', 'Arrival Time (UTC)', 'Arrival Time (Local)', 'Alba (UTC)', 'Tramonto (UTC)', 'Luce']:
            if col in df.columns:
                output_cols.append(col)

        col1, col2 = st.columns([1.2, 0.8])

        with col1:
            st.subheader("📍 Tabella Effemeridi Rotta")
            st.dataframe(df[output_cols], use_container_width=True)
            
            # Download del CSV elaborato
            csv_data = df[output_cols].to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Scarica Tabella Elaborata (CSV)",
                data=csv_data,
                file_name='rotta_effemeridi.csv',
                mime='text/csv',
            )

        with col2:
            st.subheader("🗺️ Mappa della Rotta")
            valid_points = df.dropna(subset=['Lat_Decimal', 'Lon_Decimal'])
            if not valid_points.empty:
                center_lat = valid_points['Lat_Decimal'].mean()
                center_lon = valid_points['Lon_Decimal'].mean()
                m = folium.Map(location=[center_lat, center_lon], zoom_start=5)
                
                points = list(zip(valid_points['Lat_Decimal'], valid_points['Lon_Decimal']))
                folium.PolyLine(points, color="blue", weight=3, opacity=0.8).add_to(m)
                
                for idx, row in valid_points.iterrows():
                    wp_name = row.get('Name') if pd.notna(row.get('Name')) else f"WP {row.get('Waypoint', idx+1)}"
                    popup_text = f"""
                    <b>{wp_name}</b><br>
                    UTC: {row['Arrival Time (UTC)']}<br>
                    Alba (UTC): {row['Alba (UTC)']}<br>
                    Tramonto (UTC): {row['Tramonto (UTC)']}<br>
                    Stato: {row['Luce']}
                    """
                    color = 'orange' if 'Giorno' in str(row['Luce']) else 'darkblue'
                    folium.Marker(
                        location=[row['Lat_Decimal'], row['Lon_Decimal']],
                        popup=popup_text,
                        tooltip=f"{wp_name} ({row['Luce']})",
                        icon=folium.Icon(color=color, icon='ship', prefix='fa')
                    ).add_to(m)
                    
                st_folium(m, width=550, height=500)
            else:
                st.warning("Nessuna coordinata valida trovata nel file.")

    except Exception as e:
        st.error(f"Errore durante l'elaborazione del file: {e}")
