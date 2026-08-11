import streamlit as st
import pandas as pd
import datetime
import re
import folium
from streamlit_folium import st_folium
from astral import LocationInfo
from astral.sun import sun

# Configurazione Pagina
st.set_page_config(page_title="Riepilogo Giornaliero Alba/Tramonto", layout="wide")
st.title("🌅 Riepilogo Giornaliero Effemeridi Rotta Navale")
st.write("Visualizza per ogni giorno di viaggio (Partenza, Sea Day, Arrivo) un singolo orario di alba e tramonto in Ora Locale.")

# Sidebar per caricamento file
st.sidebar.header("Carica Dati Rotta")
uploaded_file = st.sidebar.file_uploader("Carica file CSV (.csv)", type=["csv"])

# Conversione coordinate nautiche (es. 33° 45.007' N) in decimali
def parse_coordinate(val):
    if pd.isna(val) or isinstance(val, (int, float)):
        return float(val) if pd.notna(val) else None
    
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

# Funzione Calcolo Alba/Tramonto locale per un dato Waypoint
def get_sun_time_local(wp_row, target_event='sunrise'):
    lat = wp_row['Lat_Decimal']
    lon = wp_row['Lon_Decimal']
    dt_utc = wp_row['Arrival Time (UTC)']
    dt_local = wp_row['Arrival Time (Local)']
    
    if pd.isna(lat) or pd.isna(lon) or pd.isna(dt_utc):
        return "N/D"
    
    try:
        city = LocationInfo("Point", "Region", "UTC", lat, lon)
        s = sun(city.observer, date=dt_utc.date())
        
        utc_time = s['sunrise'] if target_event == 'sunrise' else s['sunset']
        
        # Applica l'offset locale
        if pd.notna(dt_local):
            offset = dt_local - dt_utc
            local_time = utc_time + offset
        else:
            local_time = utc_time
            
        return local_time.strftime('%H:%M')
    except Exception:
        return "N/D"

# Elaborazione File CSV
if uploaded_file is None:
    st.info("👋 Carica il file CSV della rotta dalla barra laterale per generare il riepilogo.")
else:
    try:
        encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']
        df = None
        for enc in encodings:
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding=enc, sep=None, engine='python')
                break
            except Exception:
                continue

        # Pulizia nomi colonne
        df.columns = df.columns.str.replace('ï»¿', '').str.strip()

        # Normalizzazione Colonne
        col_utc = next((c for c in df.columns if 'arrival time (utc)' in c.lower()), None)
        col_local = next((c for c in df.columns if 'arrival time (local)' in c.lower()), None)
        col_lat = next((c for c in df.columns if 'latitude' in c.lower()), None)
        col_lon = next((c for c in df.columns if 'longitude' in c.lower()), None)
        col_name = next((c for c in df.columns if c.lower() == 'name'), None)
        col_wp = next((c for c in df.columns if 'waypoint' in c.lower()), None)

        if not (col_utc and col_lat and col_lon):
            st.error("Il file deve contenere almeno le colonne di Latitudine, Longitudine e Arrival Time (UTC).")
            st.stop()

        df['Lat_Decimal'] = df[col_lat].apply(parse_coordinate)
        df['Lon_Decimal'] = df[col_lon].apply(parse_coordinate)
        df['Arrival Time (UTC)'] = pd.to_datetime(df[col_utc], errors='coerce')
        df['Arrival Time (Local)'] = pd.to_datetime(df[col_local], errors='coerce') if col_local else df['Arrival Time (UTC)']
        
        # Estrazione Data Locale
        df['Date_Local'] = df['Arrival Time (Local)'].dt.date
        unique_dates = sorted(df['Date_Local'].dropna().unique())

        daily_summary = []

        for i, d in enumerate(unique_dates):
            day_df = df[df['Date_Local'] == d].copy()
            if day_df.empty:
                continue

            # Determinazione Tipo Giorno
            if i == 0:
                tipo_giorno = "🛫 Partenza"
            elif i == len(unique_dates) - 1:
                tipo_giorno = "🛬 Arrivo"
            else:
                tipo_giorno = "🌊 Sea Day"

            # WP più vicino alle 07:00 (Alba)
            target_7am = pd.Timestamp.combine(d, datetime.time(7, 0))
            day_df['diff_7am'] = (day_df['Arrival Time (Local)'] - target_7am).abs()
            wp_7am = day_df.loc[day_df['diff_7am'].idxmin()]

            # WP più vicino alle 18:00 (Tramonto)
            target_6pm = pd.Timestamp.combine(d, datetime.time(18, 0))
            day_df['diff_6pm'] = (day_df['Arrival Time (Local)'] - target_6pm).abs()
            wp_6pm = day_df.loc[day_df['diff_6pm'].idxmin()]

            # Calcolo orari
            alba_local = get_sun_time_local(wp_7am, 'sunrise')
            tramonto_local = get_sun_time_local(wp_6pm, 'sunset')

            wp_7am_label = wp_7am.get(col_name) if col_name and pd.notna(wp_7am.get(col_name)) else f"WP {wp_7am.get(col_wp, '')}"
            wp_6pm_label = wp_6pm.get(col_name) if col_name and pd.notna(wp_6pm.get(col_name)) else f"WP {wp_6pm.get(col_wp, '')}"

            daily_summary.append({
                'Data': d.strftime('%d/%m/%Y'),
                'Tipo Giorno': tipo_giorno,
                'Alba (Local)': alba_local,
                'Tramonto (Local)': tramonto_local,
                'WP Rif. Alba (~07:00)': f"{wp_7am_label} ({wp_7am['Arrival Time (Local)'].strftime('%H:%M')})",
                'WP Rif. Tramonto (~18:00)': f"{wp_6pm_label} ({wp_6pm['Arrival Time (Local)'].strftime('%H:%M')})"
            })

        summary_df = pd.DataFrame(daily_summary)

        # Output Interfaccia
        col1, col2 = st.columns([1.3, 0.7])

        with col1:
            st.subheader("📋 Tabella Giornaliera Effemeridi")
            st.dataframe(summary_df, use_container_width=True)

            csv_out = summary_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Scarica Report Giornaliero (CSV)",
                data=csv_out,
                file_name="riepilogo_giornaliero_effemeridi.csv",
                mime="text/csv"
            )

        with col2:
            st.subheader("🗺️ Mappa della Rotta")
            valid_points = df.dropna(subset=['Lat_Decimal', 'Lon_Decimal'])
            if not valid_points.empty:
                m = folium.Map(location=[valid_points['Lat_Decimal'].mean(), valid_points['Lon_Decimal'].mean()], zoom_start=4)
                points = list(zip(valid_points['Lat_Decimal'], valid_points['Lon_Decimal']))
                folium.PolyLine(points, color="blue", weight=3, opacity=0.7).add_to(m)

                for idx, row in valid_points.iterrows():
                    wp_label = row.get(col_name) if col_name and pd.notna(row.get(col_name)) else f"WP {row.get(col_wp, idx+1)}"
                    folium.CircleMarker(
                        location=[row['Lat_Decimal'], row['Lon_Decimal']],
                        radius=4,
                        color="red",
                        fill=True,
                        tooltip=f"{wp_label} ({row['Arrival Time (Local)'].strftime('%d/%m %H:%M') if pd.notna(row['Arrival Time (Local)']) else ''})"
                    ).add_to(m)

                st_folium(m, width=500, height=450)

    except Exception as e:
        st.error(f"Errore durante l'elaborazione del file: {e}")
