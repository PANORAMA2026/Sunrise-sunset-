import streamlit as st
import pandas as pd
import datetime
import re
import io
import urllib.request
import json
import folium
from streamlit_folium import st_folium
from astral import LocationInfo
from astral.sun import sun
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Configurazione della pagina Streamlit
st.set_page_config(page_title="Gestore Rotte & Effemeridi", layout="wide")

st.title("🚢 Calcolatore Effemeridi e Gestore Rotte Navali")
st.write("Visualizza le effemeridi per la rotta con identificazione automatica dei porti e giorni di navigazione (Seaday).")

# ----------------------------------------------------
# FUNZIONI UTILI, REVERSE GEOCODING & PARSING
# ----------------------------------------------------

def parse_coordinate(val):
    """Converte coordinate nautiche (es. 33° 45.007' N) in formato decimale."""
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

def get_port_name_from_coords(lat, lon):
    """Effettua il Reverse Geocoding usando OpenStreetMap Nominatim per trovare il nome del porto/città."""
    if lat is None or lon is None:
        return "N/D"
    try:
        url = f"https://nominatim.openstreetmap.org/reverse?lat={lat}&lon={lon}&format=json"
        req = urllib.request.Request(url, headers={'User-Agent': 'NavalRoutePlanner/1.0'})
        with urllib.request.urlopen(req, timeout=3) as response:
            data = json.loads(response.read().decode())
            address = data.get('address', {})
            # Cerca nell'ordine: nome del porto, città, comune o paese
            port = address.get('port') or address.get('city') or address.get('town') or address.get('village') or address.get('county')
            return port.upper() if port else "PORTO"
    except Exception:
        return "PORTO"

def parse_user_date(d_str):
    """Accetta sia AAAA-MM-GG che GG/MM/AAAA."""
    d_str = d_str.strip()
    if not d_str:
        return None
    
    if '-' in d_str:
        try:
            return pd.to_datetime(d_str, format='%Y-%m-%d').date()
        except Exception:
            pass
            
    if '/' in d_str:
        try:
            return pd.to_datetime(d_str, dayfirst=True).date()
        except Exception:
            pass

    try:
        return pd.to_datetime(d_str, dayfirst=True).date()
    except Exception:
        return None

def format_timezone(offset_seconds):
    """Formatta l'offset del fuso orario nello stile 7 W / 8 W / 2 E."""
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
    """Calcola l'orario di alba/tramonto locale per un dato waypoint."""
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
    """Genera file Excel formattato."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EFFEMERIDI"
    
    headers = ["DATE", "", "PORT OF CALL", "TIME ZONE", "SUNRISE", "SUNSET"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for _, row in df_report.iterrows():
        ws.append([
            row['DATE'],
            "",
            row['PORT OF CALL'],
            row['TIME ZONE'],
            row['SUNRISE'],
            row['SUNSET']
        ])
        
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
# SIDEBAR: CONFIGURAZIONE
# ----------------------------------------------------

st.sidebar.header("📁 1. Carica File Rotte")
uploaded_files = st.sidebar.file_uploader("Carica uno o più file CSV (.csv)", type=["csv"], accept_multiple_files=True)

route_configs = {}

if uploaded_files:
    st.sidebar.header("📅 2. Impostazioni Rotta e Porti")

    for file in uploaded_files:
        st.sidebar.markdown(f"---")
        st.sidebar.markdown(f"**File:** `{file.name}`")
        
        # Lettura preliminare per individuare i punti estremi della rotta
        encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']
        df_preview = None
        for enc in encodings:
            try:
                file.seek(0)
                df_preview = pd.read_csv(file, encoding=enc, sep=None, engine='python')
                break
            except Exception:
                continue

        auto_dep = "PORTO DI PARTENZA"
        auto_arr = "PORTO DI ARRIVO"

        if df_preview is not None:
            df_preview.columns = df_preview.columns.str.replace('ï»¿', '').str.strip()
            col_lat = next((c for c in df_preview.columns if 'latitude' in c.lower()), None)
            col_lon = next((c for c in df_preview.columns if 'longitude' in c.lower()), None)
            
            if col_lat and col_lon:
                first_lat = parse_coordinate(df_preview.iloc[0][col_lat])
                first_lon = parse_coordinate(df_preview.iloc[0][col_lon])
                last_lat = parse_coordinate(df_preview.iloc[-1][col_lat])
                last_lon = parse_coordinate(df_preview.iloc[-1][col_lon])

                # Calcolo automatico tramite Reverse Geocoding
                auto_dep = get_port_name_from_coords(first_lat, first_lon)
                auto_arr = get_port_name_from_coords(last_lat, last_lon)

        # Input utente con valore precompilato dall'algoritmo
        dep_port = st.sidebar.text_input("Porto di Partenza:", value=auto_dep, key=f"dep_{file.name}")
        arr_port = st.sidebar.text_input("Porto di Arrivo:", value=auto_arr, key=f"arr_{file.name}")
        
        dates_str = st.sidebar.text_input(
            "Nuova Data di Partenza (AAAA-MM-GG o GG/MM/AAAA):",
            value="",
            key=f"dates_{file.name}",
            placeholder="Opzionale (es. 2026-08-01)"
        )
        
        parsed_dates = []
        if dates_str.strip():
            for d_str in dates_str.split(','):
                parsed = parse_user_date(d_str)
                if parsed:
                    parsed_dates.append(parsed)
                else:
                    st.sidebar.warning(f"Formato non valido: '{d_str}'. Usa AAAA-MM-GG o GG/MM/AAAA.")
        
        route_configs[file.name] = {
            'file': file,
            'start_dates': parsed_dates,
            'dep_port': dep_port,
            'arr_port': arr_port
        }

# ----------------------------------------------------
# PROCESSING & TABELLA DI OUTPUT
# ----------------------------------------------------

if not uploaded_files:
    st.info("👋 Carica uno o più file CSV di rotta dalla barra laterale per iniziare.")
else:
    all_summary_rows = []
    all_route_dfs = []

    for file_name, config in route_configs.items():
        file = config['file']
        user_start_dates = config['start_dates']
        dep_port_name = config['dep_port']
        arr_port_name = config['arr_port']

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

        df_base.columns = df_base.columns.str.replace('ï»¿', '').str.strip()

        col_utc = next((c for c in df_base.columns if 'arrival time (utc)' in c.lower()), None)
        col_local = next((c for c in df_base.columns if 'arrival time (local)' in c.lower()), None)
        col_lat = next((c for c in df_base.columns if 'latitude' in c.lower()), None)
        col_lon = next((c for c in df_base.columns if 'longitude' in c.lower()), None)

        if not (col_utc and col_lat and col_lon):
            st.error(f"Il file {file_name} non contiene tutte le colonne obbligatorie.")
            continue

        df_base['Lat_Decimal'] = df_base[col_lat].apply(parse_coordinate)
        df_base['Lon_Decimal'] = df_base[col_lon].apply(parse_coordinate)
        
        df_base['Arrival Time (UTC)'] = pd.to_datetime(df_base[col_utc], dayfirst=True, errors='coerce')
        df_base['Arrival Time (Local)'] = pd.to_datetime(df_base[col_local], dayfirst=True, errors='coerce') if col_local else df_base['Arrival Time (UTC)']

        original_base_datetime = df_base['Arrival Time (Local)'].dropna().iloc[0]
        original_base_date = original_base_datetime.date()

        target_start_dates = user_start_dates if user_start_dates else [original_base_date]

        for start_date in target_start_dates:
            days_shift = (start_date - original_base_date).days
            
            df_instance = df_base.copy()
            df_instance['Arrival Time (UTC)'] = df_instance['Arrival Time (UTC)'] + pd.Timedelta(days=days_shift)
            df_instance['Arrival Time (Local)'] = df_instance['Arrival Time (Local)'] + pd.Timedelta(days=days_shift)
            df_instance['Date_Local'] = df_instance['Arrival Time (Local)'].dt.date
            
            unique_dates = sorted(df_instance['Date_Local'].dropna().unique())

            for i, d in enumerate(unique_dates):
                day_df = df_instance[df_instance['Date_Local'] == d].copy()
                if day_df.empty:
                    continue

                # LOGICA ASSEGNAZIONE NOME PORTO / SEADAY
                if i == 0:
                    port_of_call = dep_port_name
                elif i == len(unique_dates) - 1:
                    port_of_call = arr_port_name
                else:
                    port_of_call = "Seaday"

                # Waypoint per Alba (07:00) e Tramonto (18:00)
                target_7am = pd.Timestamp.combine(d, datetime.time(7, 0))
                day_df['diff_7am'] = (day_df['Arrival Time (Local)'] - target_7am).abs()
                wp_7am = day_df.loc[day_df['diff_7am'].idxmin()]

                target_6pm = pd.Timestamp.combine(d, datetime.time(18, 0))
                day_df['diff_6pm'] = (day_df['Arrival Time (Local)'] - target_6pm).abs()
                wp_6pm = day_df.loc[day_df['diff_6pm'].idxmin()]

                offset_sec = (wp_7am['Arrival Time (Local)'] - wp_7am['Arrival Time (UTC)']).total_seconds()
                time_zone_str = format_timezone(offset_sec)

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

        st.subheader("📊 Report Effemeridi Calcolato")
        st.dataframe(df_report_full[['DATE', 'PORT OF CALL', 'TIME ZONE', 'SUNRISE', 'SUNSET', 'ROTTA']], use_container_width=True)

        col_dl1, col_dl2 = st.columns(2)

        with col_dl1:
            excel_bytes = generate_excel_output(df_report_full)
            st.download_button(
                label="🟢 Scarica Report in Excel (.xlsx)",
                data=excel_bytes,
                file_name="Effemeridi_Rotte_Navali.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_dl2:
            csv_data = df_report_full[['DATE', 'PORT OF CALL', 'TIME ZONE', 'SUNRISE', 'SUNSET']].to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Scarica Report in CSV",
                data=csv_data,
                file_name="Effemeridi_Rotte_Navali.csv",
                mime="text/csv"
            )

        st.subheader("🗺️ Mappa Interattiva delle Rotte")
        combined_points = pd.concat([df.dropna(subset=['Lat_Decimal', 'Lon_Decimal']) for df in all_route_dfs])
        if not combined_points.empty:
            m = folium.Map(location=[combined_points['Lat_Decimal'].mean(), combined_points['Lon_Decimal'].mean()], zoom_start=4)
            for df_inst in all_route_dfs:
                pts = list(zip(df_inst['Lat_Decimal'], df_inst['Lon_Decimal']))
                folium.PolyLine(pts, color="blue", weight=3, opacity=0.7).add_to(m)
            st_folium(m, width=900, height=450)
