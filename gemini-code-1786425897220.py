import streamlit as st
import pandas as pd
import datetime
import re
import io
import pydeck as pdk
from astral import LocationInfo
from astral.sun import sun
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="Gestore Effemeridi & Calendario Crociere", layout="wide")

st.title("🚢 Generatore Effemeridi da Calendario Crociere (Date, Location, ETA, ETD)")
st.write("Carica il file Excel del **Calendario Crociere** e il file CSV della **Rotta Navale** per calcolare automaticamente le effemeridi locali.")

# ----------------------------------------------------
# FUNZIONI UTILI & PARSING
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

def format_timezone(offset_seconds):
    """Formatta l'offset del fuso orario nello stile 7 W / 8 W / 2 E / UTC."""
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

def parse_cruise_calendar(excel_file):
    """Legge il file Excel del calendario crociere riconoscendo Date, Location, ETA, ETD."""
    df_cal = pd.read_excel(excel_file)
    col_map = {str(c).strip().lower(): c for c in df_cal.columns}
    
    date_col = col_map.get('date')
    loc_col = col_map.get('location')
    eta_col = col_map.get('eta')
    etd_col = col_map.get('etd')
    
    return df_cal, date_col, loc_col, eta_col, etd_col

def generate_excel_output(df_report):
    """Genera file Excel formattato secondo il layout standard."""
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
# BARRA LATERALE: CARICAMENTO FILE
# ----------------------------------------------------

st.sidebar.header("📁 1. Calendario Crociere (Excel)")
calendar_file = st.sidebar.file_uploader("Carica Calendario Excel (.xlsx / .xls)", type=["xlsx", "xls"])

st.sidebar.header("🗺️ 2. Rotta Navale (CSV)")
route_files = st.sidebar.file_uploader("Carica file Rotta CSV (.csv)", type=["csv"], accept_multiple_files=True)

# ----------------------------------------------------
# ELABORAZIONE
# ----------------------------------------------------

if not calendar_file or not route_files:
    st.info("👈 Carica il file **Calendario Excel** e almeno un file **CSV di Rotta** dalla barra laterale per proseguire.")
else:
    # 1. LETTURA ED ESTRAZIONE COLONNE CALENDARIO
    try:
        df_cal_raw, auto_date, auto_loc, auto_eta, auto_etd = parse_cruise_calendar(calendar_file)
    except Exception as e:
        st.error(f"Errore nella lettura del file Excel: {e}")
        st.stop()

    cols = list(df_cal_raw.columns)
    
    st.subheader("⚙️ Verifica Mappatura Colonne Calendario")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        sel_date = st.selectbox("Colonna DATE:", cols, index=cols.index(auto_date) if auto_date in cols else 0)
    with c2:
        sel_loc = st.selectbox("Colonna LOCATION:", cols, index=cols.index(auto_loc) if auto_loc in cols else min(1, len(cols)-1))
    with c3:
        sel_eta = st.selectbox("Colonna ETA (Opzionale):", ["Nessuna"] + cols, index=(cols.index(auto_eta)+1) if auto_eta in cols else 0)
    with c4:
        sel_etd = st.selectbox("Colonna ETD (Opzionale):", ["Nessuna"] + cols, index=(cols.index(auto_etd)+1) if auto_etd in cols else 0)

    # Pulizia e strutturazione dati del calendario
    df_cal = df_cal_raw.copy()
    df_cal['Date_Parsed'] = pd.to_datetime(df_cal[sel_date], errors='coerce').dt.date
    df_cal = df_cal.dropna(subset=['Date_Parsed']).copy()
    df_cal['Location_Clean'] = df_cal[sel_loc].astype(str).str.strip()

    # Eliminazione duplicati sullo stesso giorno dando priorità allo scalo reale rispetto a "at Sea"
    def deduplicate_day(group):
        if len(group) == 1:
            return group.iloc[0]
        non_sea = group[~group['Location_Clean'].str.contains('Sea', case=False, na=False)]
        if not non_sea.empty:
            return non_sea.iloc[0]
        return group.iloc[0]

    df_cal = df_cal.groupby('Date_Parsed', as_index=False, group_keys=False).apply(deduplicate_day).reset_index(drop=True)

    # 2. LETTURA ED UNIONE ROTTE CSV
    all_route_waypoints = []
    for r_file in route_files:
        encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']
        df_r = None
        for enc in encodings:
            try:
                r_file.seek(0)
                df_r = pd.read_csv(r_file, encoding=enc, sep=None, engine='python')
                break
            except Exception:
                continue

        if df_r is not None:
            df_r.columns = df_r.columns.str.replace('ï»¿', '').str.strip()
            c_utc = next((c for c in df_r.columns if 'arrival time (utc)' in c.lower()), None)
            c_local = next((c for c in df_r.columns if 'arrival time (local)' in c.lower()), None)
            c_lat = next((c for c in df_r.columns if 'latitude' in c.lower()), None)
            c_lon = next((c for c in df_r.columns if 'longitude' in c.lower()), None)

            if c_utc and c_lat and c_lon:
                df_r['Lat_Decimal'] = df_r[c_lat].apply(parse_coordinate)
                df_r['Lon_Decimal'] = df_r[c_lon].apply(parse_coordinate)
                df_r['Arrival Time (UTC)'] = pd.to_datetime(df_r[c_utc], dayfirst=True, errors='coerce')
                df_r['Arrival Time (Local)'] = pd.to_datetime(df_r[c_local], dayfirst=True, errors='coerce') if c_local else df_r['Arrival Time (UTC)']
                all_route_waypoints.append(df_r)

    if not all_route_waypoints:
        st.error("Nessun file di rotta valido caricato o colonne obbligatorie mancanti (Latitude, Longitude, Arrival Time UTC).")
        st.stop()

    df_full_route = pd.concat(all_route_waypoints, ignore_index=True)
    route_dates = sorted(df_full_route['Arrival Time (Local)'].dt.date.dropna().unique())
    days_in_route = len(route_dates)
    cal_start_date = df_cal['Date_Parsed'].iloc[0]

    # 3. ELABORAZIONE CALENDARIO - CALCOLO EFFEMERIDI
    summary_rows = []
    processed_dfs = []

    for _, cal_row in df_cal.iterrows():
        c_date = cal_row['Date_Parsed']
        port_name = cal_row['Location_Clean']
        
        # Mappatura del giorno del calendario sulla sequenza waypoint della rotta
        day_offset = (c_date - cal_start_date).days % days_in_route
        target_route_date = route_dates[day_offset]

        day_route = df_full_route[df_full_route['Arrival Time (Local)'].dt.date == target_route_date].copy()
        if day_route.empty:
            continue

        time_shift = pd.Timestamp(c_date) - pd.Timestamp(target_route_date)
        day_route['Arrival Time (UTC)'] = day_route['Arrival Time (UTC)'] + time_shift
        day_route['Arrival Time (Local)'] = day_route['Arrival Time (Local)'] + time_shift

        # Punto rotta ore 07:00 (Alba) e ore 18:00 (Tramonto)
        target_7am = pd.Timestamp.combine(c_date, datetime.time(7, 0))
        day_route['diff_7am'] = (day_route['Arrival Time (Local)'] - target_7am).abs()
        wp_7am = day_route.loc[day_route['diff_7am'].idxmin()]

        target_6pm = pd.Timestamp.combine(c_date, datetime.time(18, 0))
        day_route['diff_6pm'] = (day_route['Arrival Time (Local)'] - target_6pm).abs()
        wp_6pm = day_route.loc[day_route['diff_6pm'].idxmin()]

        offset_sec = (wp_7am['Arrival Time (Local)'] - wp_7am['Arrival Time (UTC)']).total_seconds()
        tz_str = format_timezone(offset_sec)

        sunrise_str = get_sun_time_local(wp_7am, 'sunrise')
        sunset_str = get_sun_time_local(wp_6pm, 'sunset')

        summary_rows.append({
            'DATE': c_date.strftime('%Y-%m-%d'),
            'PORT OF CALL': port_name,
            'TIME ZONE': tz_str,
            'SUNRISE': sunrise_str,
            'SUNSET': sunset_str
        })
        processed_dfs.append(day_route)

    # 4. TABELLA DI OUTPUT E DOWNLOAD
    if summary_rows:
        df_result = pd.DataFrame(summary_rows)

        st.subheader("📊 Report Effemeridi Calcolato")
        st.dataframe(df_result, use_container_width=True)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            excel_bytes = generate_excel_output(df_result)
            st.download_button(
                label="🟢 Scarica Report Excel (.xlsx)",
                data=excel_bytes,
                file_name="Effemeridi_Calendario_Crociere.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_dl2:
            csv_data = df_result.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Scarica Report CSV",
                data=csv_data,
                file_name="Effemeridi_Calendario_Crociere.csv",
                mime="text/csv"
            )

        # 5. VISUALIZZAZIONE MAPPA INTERATTIVA (PyDeck)
        st.subheader("🗺️ Mappa della Rotta Proiettata")
        combined_pts = pd.concat([df.dropna(subset=['Lat_Decimal', 'Lon_Decimal']) for df in processed_dfs])
        if not combined_pts.empty:
            path_coords = combined_pts[['Lon_Decimal', 'Lat_Decimal']].values.tolist()
            view_state = pdk.ViewState(
                latitude=combined_pts['Lat_Decimal'].mean(),
                longitude=combined_pts['Lon_Decimal'].mean(),
                zoom=4
            )
            line_layer = pdk.Layer(
                "PathLayer",
                [{"path": path_coords}],
                get_path="path",
                get_color=[31, 78, 120, 255],
                width_scale=20,
                width_min_pixels=3
            )
            st.pydeck_chart(pdk.Deck(layers=[line_layer], initial_view_state=view_state))
