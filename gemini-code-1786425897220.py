import streamlit as st
import pandas as pd
import datetime
import re
import io
import folium
from streamlit_folium import st_folium
from astral import LocationInfo
from astral.sun import sun
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Configurazione della pagina Streamlit
st.set_page_config(page_title="Gestore Rotte & Effemeridi", layout="wide")

st.title("🚢 Gestore Rotte Navali & Calcolatore Effemeridi")
st.write("Carica le tue rotte, definisci le date di partenza (anche ripetute) ed esporta il report Excel formattato.")

# ----------------------------------------------------
# FUNZIONI UTILI
# ----------------------------------------------------

def parse_coordinate(val):
    """Converte coordinate in vari formati (es. 33° 45.007' N) in decimali."""
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

def format_timezone(offset_seconds):
    """Formatta l'offset del fuso orario nello stile di Book1.xlsx (es. 7 W, 8 W, 2 E)."""
    if pd.isna(offset_seconds):
        return "N/D"
    hours = round(offset_seconds / 3600)
    if hours < 0:
        return f"{abs(hours)} W"
    elif hours > 0:
        return f"{hours} E"
    else:
        return "UTC"

def get_sun_time_local(wp_row, target_event='sunrise'):
    """Calcola l'orario di alba/tramonto locale per un specifico waypoint."""
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
        
        if pd.notna(dt_local):
            offset = dt_local - dt_utc
            local_time = utc_time + offset
        else:
            local_time = utc_time
            
        return local_time.strftime('%H:%M:%S')
    except Exception:
        return "N/D"

def generate_excel_output(df_report):
    """Genera un file Excel (.xlsx) formattato esattamente come Book1.xlsx."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EFFEMERIDI"
    
    # Intestazioni di colonna
    headers = ["DATE", "", "PORT OF CALL", "TIME ZONE", "SUNRISE", "SUNSET"]
    ws.append(headers)
    
    # Stile Intestazione
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # Righe Dati
    for _, row in df_report.iterrows():
        ws.append([
            row['DATE'],
            "",
            row['PORT OF CALL'],
            row['TIME ZONE'],
            row['SUNRISE'],
            row['SUNSET']
        ])
        
    # Bordi e Allineamento Dati
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    for row in ws.iter_rows(min_row=2, max_row=len(df_report)+1, min_col=1, max_col=6):
        for cell in row:
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
    # Larghezza Colonne
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 4
    ws.column_dimensions['C'].width = 32
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ----------------------------------------------------
# SIDEBAR: CARICAMENTO & CONFIGURAZIONE DATE
# ----------------------------------------------------

st.sidebar.header("📁 1. Carica File Rotte")
uploaded_files = st.sidebar.file_uploader("Carica uno o più file CSV (.csv)", type=["csv"], accept_multiple_files=True)

route_configs = {}

if uploaded_files:
    st.sidebar.header("📅 2. Date di Partenza Rotta")
    st.sidebar.write("Inserisci le date di partenza (separate da virgola) per ripercorrere la stessa rotta in giorni diversi:")

    for file in uploaded_files:
        st.sidebar.markdown(f"**Rotta:** `{file.name}`")
        dates_str = st.sidebar.text_input(
            f"Date di Partenza (AAAA-MM-GG):",
            value=datetime.date.today().strftime("%Y-%m-%d"),
            key=f"dates_{file.name}"
        )
        
        # Parsing date inserite dall'utente
        parsed_dates = []
        for d_str in dates_str.split(','):
            d_str = d_str.strip()
            if d_str:
                try:
                    dt = datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
                    parsed_dates.append(dt)
                except ValueError:
                    st.sidebar.warning(f"Formato data non valido '{d_str}'. Usa AAAA-MM-GG.")
        
        route_configs[file.name] = {
            'file': file,
            'start_dates': parsed_dates
        }

# ----------------------------------------------------
# MAIN LOGIC & REPORT GENERATION
# ----------------------------------------------------

if not uploaded_files:
    st.info("👋 Carica uno o più file CSV di rotta dalla barra laterale per iniziare.")
else:
    all_summary_rows = []
    all_route_dfs = []

    for file_name, config in route_configs.items():
        file = config['file']
        start_dates = config['start_dates']

        if not start_dates:
            continue

        # Lettura file CSV
        encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']
        df_base = None
        for enc in encodings:
            try:
                file.seek(0)
                df_base = pd.read_csv(file, encoding=enc, sep=None, engine='python')
                break
            except Exception:
                continue

        if df_base is None:
            st.error(f"Impossibile leggere il file {file_name}.")
            continue

        # Pulizia colonne
        df_base.columns = df_base.columns.str.replace('ï»¿', '').str.strip()

        col_utc = next((c for c in df_base.columns if 'arrival time (utc)' in c.lower()), None)
        col_local = next((c for c in df_base.columns if 'arrival time (local)' in c.lower()), None)
        col_lat = next((c for c in df_base.columns if 'latitude' in c.lower()), None)
        col_lon = next((c for c in df_base.columns if 'longitude' in c.lower()), None)
        col_name = next((c for c in df_base.columns if c.lower() == 'name'), None)
        col_wp = next((c for c in df_base.columns if 'waypoint' in c.lower()), None)

        if not (col_utc and col_lat and col_lon):
            st.error(f"Il file {file_name} non contiene tutte le colonne obbligatorie (Latitudine, Longitudine, Arrival Time UTC).")
            continue

        df_base['Lat_Decimal'] = df_base[col_lat].apply(parse_coordinate)
        df_base['Lon_Decimal'] = df_base[col_lon].apply(parse_coordinate)
        df_base['Arrival Time (UTC)'] = pd.to_datetime(df_base[col_utc], errors='coerce')
        df_base['Arrival Time (Local)'] = pd.to_datetime(df_base[col_local], errors='coerce') if col_local else df_base['Arrival Time (UTC)']

        # Data di partenza originaria del file CSV
        base_start_datetime = df_base['Arrival Time (Local)'].dropna().iloc[0]
        base_start_date = base_start_datetime.date()

        # Genero l'esecuzione della rotta per OGNI data di partenza richiesta
        for start_date in start_dates:
            days_shift = (start_date - base_start_date).days
            
            df_instance = df_base.copy()
            df_instance['Arrival Time (UTC)'] = df_instance['Arrival Time (UTC)'] + pd.Timedelta(days=days_shift)
            df_instance['Arrival Time (Local)'] = df_instance['Arrival Time (Local)'] + pd.Timedelta(days=days_shift)
            df_instance['Date_Local'] = df_instance['Arrival Time (Local)'].dt.date
            
            unique_dates = sorted(df_instance['Date_Local'].dropna().unique())

            for i, d in enumerate(unique_dates):
                day_df = df_instance[df_instance['Date_Local'] == d].copy()
                if day_df.empty:
                    continue

                # Identificazione Nome Porto / Fun Day at Sea
                if i == 0:
                    port_of_call = day_df.iloc[0].get(col_name) if col_name and pd.notna(day_df.iloc[0].get(col_name)) else "Porto di Partenza"
                elif i == len(unique_dates) - 1:
                    port_of_call = day_df.iloc[-1].get(col_name) if col_name and pd.notna(day_df.iloc[-1].get(col_name)) else "Porto di Arrivo"
                else:
                    port_of_call = "Fun Day at Sea"

                # WP più vicino alle 07:00 (Alba)
                target_7am = pd.Timestamp.combine(d, datetime.time(7, 0))
                day_df['diff_7am'] = (day_df['Arrival Time (Local)'] - target_7am).abs()
                wp_7am = day_df.loc[day_df['diff_7am'].idxmin()]

                # WP più vicino alle 18:00 (Tramonto)
                target_6pm = pd.Timestamp.combine(d, datetime.time(18, 0))
                day_df['diff_6pm'] = (day_df['Arrival Time (Local)'] - target_6pm).abs()
                wp_6pm = day_df.loc[day_df['diff_6pm'].idxmin()]

                # Calcolo Fuso Orario (Time Zone)
                offset_sec = (wp_7am['Arrival Time (Local)'] - wp_7am['Arrival Time (UTC)']).total_seconds()
                time_zone_str = format_timezone(offset_sec)

                # Calcolo Effemeridi
                sunrise_str = get_sun_time_local(wp_7am, 'sunrise')
                sunset_str = get_sun_time_local(wp_6pm, 'sunset')

                all_summary_rows.append({
                    'DATE': d.strftime('%Y-%m-%d'),
                    'PORT OF CALL': port_of_call,
                    'TIME ZONE': time_zone_str,
                    'SUNRISE': sunrise_str,
                    'SUNSET': sunset_str,
                    'ROTTA': file_name
                })

            all_route_dfs.append(df_instance)

    if all_summary_rows:
        df_report_full = pd.DataFrame(all_summary_rows)

        # ----------------------------------------------------
        # OUTPUT VISUALE & DOWNLOAD
        # ----------------------------------------------------
        st.subheader("📊 Report Giornaliero Effemeridi (Tutte le Rotte & Date)")
        st.dataframe(df_report_full[['DATE', 'PORT OF CALL', 'TIME ZONE', 'SUNRISE', 'SUNSET', 'ROTTA']], use_container_width=True)

        col_dl1, col_dl2 = st.columns(2)

        with col_dl1:
            # Download File Excel Formattato (come Book1.xlsx)
            excel_bytes = generate_excel_output(df_report_full)
            st.download_button(
                label="🟢 Scarica Report in Excel (.xlsx)",
                data=excel_bytes,
                file_name="Effemeridi_Rotte_Navali.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_dl2:
            # Download CSV
            csv_data = df_report_full[['DATE', 'PORT OF CALL', 'TIME ZONE', 'SUNRISE', 'SUNSET']].to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Scarica Report in CSV",
                data=csv_data,
                file_name="Effemeridi_Rotte_Navali.csv",
                mime="text/csv"
            )

        # Mappa della Rotta
        st.subheader("🗺️ Mappa Interattiva delle Rotte")
        combined_points = pd.concat([df.dropna(subset=['Lat_Decimal', 'Lon_Decimal']) for df in all_route_dfs])
        if not combined_points.empty:
            m = folium.Map(location=[combined_points['Lat_Decimal'].mean(), combined_points['Lon_Decimal'].mean()], zoom_start=4)
            for df_inst in all_route_dfs:
                pts = list(zip(df_inst['Lat_Decimal'], df_inst['Lon_Decimal']))
                folium.PolyLine(pts, color="blue", weight=3, opacity=0.7).add_to(m)
            st_folium(m, width=900, height=450)
