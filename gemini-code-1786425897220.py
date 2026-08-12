import io
import math
import re
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Naval Ephemeris Calculator", layout="wide"
)

st.title("⚓ Naval Ephemeris & Route Analyzer")
st.write(
    "Calcolo di Alba, Tramonto e Time Zone basato esclusivamente sui file rotta caricate."
)

# -----------------------------------------------------------------------------
# UTILITY FUNCTIONS
# -----------------------------------------------------------------------------

def parse_coordinate(coord):
    """Converte coordinate nautiche o decimali in float."""
    if pd.isna(coord):
        return 0.0
    if isinstance(coord, (int, float)):
        return float(coord)
    
    s = str(coord).strip()
    try:
        return float(s)
    except ValueError:
        pass
    
    match = re.search(r"(\d+)°\s*(\d+(?:\.\d+)?)\s*['′]?\s*([NSEWnsew]?)", s)
    if match:
        degrees = float(match.group(1))
        minutes = float(match.group(2))
        direction = match.group(3).upper()
        
        decimal_degrees = degrees + (minutes / 60.0)
        if direction in ['S', 'W']:
            decimal_degrees = -decimal_degrees
        return decimal_degrees
        
    match_simple = re.search(r"([+-]?\d+(?:\.\d+)?)\s*([NSEWnsew]?)", s)
    if match_simple:
        val = float(match_simple.group(1))
        direction = match_simple.group(2).upper()
        if direction in ['S', 'W']:
            val = -abs(val)
        return val
        
    return 0.0


def parse_tz_offset(tz_val):
    """Estrae l'offset orario numerico dal CSV."""
    if pd.isna(tz_val):
        return 0.0
    s = str(tz_val).strip()
    if not s or s.lower() == 'nan':
        return 0.0

    match_hhmmss = re.search(r'([+-]?)\s*(\d{1,2}):(\d{2})(?::(\d{2}))?', s)
    if match_hhmmss:
        sign = -1.0 if match_hhmmss.group(1) == '-' else 1.0
        hours = float(match_hhmmss.group(2))
        minutes = float(match_hhmmss.group(3))
        return sign * (hours + minutes / 60.0)

    match_we = re.search(r'(\d+(?:\.\d+)?)\s*([WEwe])', s)
    if match_we:
        val = float(match_we.group(1))
        direction = match_we.group(2).upper()
        return -val if direction == 'W' else val

    try:
        return float(s)
    except ValueError:
        return 0.0


def format_tz_string(tz_val, offset_hours):
    """Formatta la Time Zone (es. '7 W' o '2 E')."""
    s = str(tz_val).strip()
    if not s or s.lower() == 'nan':
        return "0"
    
    match_we = re.search(r'(\d+(?:\.\d+)?)\s*([WEwe])', s)
    if match_we:
        return f"{match_we.group(1)} {match_we.group(2).upper()}"
        
    if offset_hours == 0:
        return "0"
    
    abs_h = abs(offset_hours)
    h_str = f"{int(abs_h)}" if abs_h.is_integer() else f"{abs_h}"
    direction = "W" if offset_hours < 0 else "E"
    return f"{h_str} {direction}"


def calculate_sun_events_native(lat, lon, date_obj, tz_offset_hours):
    """Calcola alba e tramonto in ora locale."""
    try:
        day_of_year = date_obj.timetuple().tm_yday
        declination = 0.409 * math.sin((2 * math.pi / 365) * (day_of_year - 81))
        lat_rad = math.radians(lat)

        cos_ha = (
            math.sin(math.radians(-0.833))
            - math.sin(lat_rad) * math.sin(declination)
        ) / (math.cos(lat_rad) * math.cos(declination))

        if cos_ha >= 1:
            return "Mai sorge", "Mai sorge"
        if cos_ha <= -1:
            return "Mai tramonta", "Mai tramonta"

        ha = math.degrees(math.acos(cos_ha))
        b = (2 * math.pi / 365) * (day_of_year - 81)
        eot = 9.87 * math.sin(2 * b) - 7.53 * math.cos(b) - 1.5 * math.sin(b)

        solar_noon_utc = (720 - (4 * lon) - eot) / 60.0

        sunrise_utc_hours = solar_noon_utc - (ha / 15.0)
        sunset_utc_hours = solar_noon_utc + (ha / 15.0)

        sunrise_local_hours = (sunrise_utc_hours + tz_offset_hours) % 24
        sunset_local_hours = (sunset_utc_hours + tz_offset_hours) % 24

        def hours_to_time_str(h_decimal):
            h = int(h_decimal)
            m = int((h_decimal - h) * 60)
            s = int((((h_decimal - h) * 60) - m) * 60)
            return f"{h:02d}:{m:02d}:{s:02d}"

        return hours_to_time_str(sunrise_local_hours), hours_to_time_str(sunset_local_hours)
    except Exception:
        return "N/D", "N/D"


# -----------------------------------------------------------------------------
# SIDEBAR - CARICAMENTO FILE
# -----------------------------------------------------------------------------
st.sidebar.header("📁 Caricamento Dati")

cal_file = st.sidebar.file_uploader(
    "1. Carica Calendario Mensile (Excel)", type=["xlsx", "xls"]
)
route_files = st.sidebar.file_uploader(
    "2. Carica File Rotta (CSV)", type=["csv"], accept_multiple_files=True
)

if cal_file and route_files:
    df_cal = pd.read_excel(cal_file)

    cal_date_col = next(
        (c for c in df_cal.columns if 'date' in c.lower() or 'data' in c.lower()),
        df_cal.columns[0],
    )
    cal_port_col = next(
        (c for c in df_cal.columns if 'location' in c.lower() or 'port' in c.lower()),
        df_cal.columns[1],
    )
    cal_cruise_col = next(
        (c for c in df_cal.columns if 'cruise' in c.lower() or 'crociera' in c.lower()),
        None,
    )

    df_cal['Date_Parsed'] = pd.to_datetime(df_cal[cal_date_col]).dt.date
    df_cal['Location_Clean'] = df_cal[cal_port_col].astype(str).str.strip()

    routes_data = []
    for rf in route_files:
        try:
            df_r = pd.read_csv(rf)
            time_col = next(
                c for c in df_r.columns if 'arrival time' in c.lower() or 'time' in c
            )
            df_r['dt_local'] = pd.to_datetime(df_r[time_col], dayfirst=True)
            df_r['date_only'] = df_r['dt_local'].dt.date

            # Ricerca nome waypoint/porto nel CSV se esiste
            name_col = next((c for c in df_r.columns if 'name' in c.lower() or 'waypoint' in c.lower()), None)
            waypoint_names = df_r[name_col].astype(str).str.lower().tolist() if name_col else []

            unique_dates = sorted(df_r['date_only'].dropna().unique())
            if unique_dates:
                start_d = unique_dates[0]
                df_r['rel_day'] = df_r['date_only'].apply(
                    lambda d: (d - start_d).days if pd.notnull(d) else None
                )
                routes_data.append({
                    'filename': rf.name,
                    'df': df_r,
                    'num_days': len(unique_dates),
                    'max_rel_day': max(df_r['rel_day'].dropna()),
                    'waypoint_names': waypoint_names
                })
        except Exception as e:
            st.sidebar.error(f"Errore nella lettura di {rf.name}: {e}")

    if not routes_data:
        st.warning("Nessun file rotta valido elaborato.")
        st.stop()

    # -----------------------------------------------------------------------------
    # ABBINAMENTO ESPLICITO
    # -----------------------------------------------------------------------------
    st.subheader("🎯 Abbinamento File Rotta")
    st.info("Associa il file di rotta caricato solo alla specifica crociera a cui si riferisce.")

    cruise_options = (
        df_cal[cal_cruise_col].dropna().unique().tolist()
        if cal_cruise_col
        else df_cal['Date_Parsed'].astype(str).unique().tolist()
    )

    mapped_routes_list = []

    for idx, r_info in enumerate(routes_data):
        col1, col2 = st.columns([2, 3])
        with col1:
            st.write(
                f"📄 **File Rotta:** `{r_info['filename']}` ({r_info['num_days']} giorni)"
            )
        with col2:
            selected_cruises = st.multiselect(
                f"Associa `{r_info['filename']}` alla Crociera:",
                options=cruise_options,
                default=[],
                key=f"multiselect_{idx}",
                placeholder="Seleziona il codice crociera dal calendario...",
            )

            for cruise_id in selected_cruises:
                if cal_cruise_col:
                    start_date = df_cal[df_cal[cal_cruise_col] == cruise_id]['Date_Parsed'].min()
                else:
                    start_date = pd.to_datetime(cruise_id).date()

                mapped_routes_list.append({
                    'filename': r_info['filename'],
                    'df': r_info['df'],
                    'cruise_id': cruise_id,
                    'start_cal_date': start_date,
                    'max_rel_day': r_info['max_rel_day'],
                    'waypoint_names': r_info['waypoint_names']
                })

    filter_by_port = st.checkbox(
        "Filtra anche per corrispondenza Nome Porto / Mare (Escludi porti intermedi non presenti nel CSV)",
        value=True,
    )

    # -----------------------------------------------------------------------------
    # ELABORAZIONE EFFEMERIDI
    # -----------------------------------------------------------------------------
    results = []

    for _, cal_row in df_cal.iterrows():
        c_date = cal_row['Date_Parsed']
        port_name = cal_row['Location_Clean']
        c_cruise = cal_row[cal_cruise_col] if cal_cruise_col else str(c_date)

        matching_waypoint = None
        matching_tz_val = '0'
        matched_filename = None

        for item in mapped_routes_list:
            r_df = item['df']
            r_start_date = item['start_cal_date']
            r_cruise_id = item['cruise_id']
            max_rel_day = item['max_rel_day']

            if c_cruise != r_cruise_id:
                continue

            rel_day = (c_date - r_start_date).days

            if 0 <= rel_day <= max_rel_day:
                # Controlla se il porto nel calendario coincide con i waypoint o con la rotta
                if filter_by_port:
                    port_lower = port_name.lower()
                    # Se il calendario dice p.es. "Cabo San Lucas" e noi stiamo caricando la rotta Long Beach-Puerto Vallarta,
                    # verifichiamo se "cabo" compare nei waypoint del CSV. Se non compare, ignoriamo questa riga!
                    csv_names_str = " ".join(item['waypoint_names'])
                    
                    # Estraiamo parole chiave significative dal nome porto (es. "Cabo", "Mazatlan", "Vallarta")
                    keywords = [w for w in re.findall(r'\w+', port_lower) if len(w) > 3 and w not in ['port', 'puebla', 'dock', 'sea', 'funday']]
                    
                    if keywords:
                        match_found = any(kw in csv_names_str for kw in keywords) or ("fun day" in port_lower and "sea" in csv_names_str)
                        if not match_found and "fun day" not in port_lower:
                            # Porto non presente nel CSV -> Salta
                            continue

                day_points = r_df[r_df['rel_day'] == rel_day]
                if not day_points.empty:
                    mid_idx = len(day_points) // 2
                    matching_waypoint = day_points.iloc[mid_idx]
                    matched_filename = item['filename']

                    tz_col = next(
                        (
                            c
                            for c in day_points.columns
                            if any(
                                k in c.lower()
                                for k in [
                                    'utc offset',
                                    'time zone',
                                    'timezone',
                                    'offset',
                                    'tz',
                                ]
                            )
                        ),
                        None,
                    )
                    if tz_col:
                        matching_tz_val = day_points.iloc[0][tz_col]
                    break

        if matching_waypoint is not None:
            lat = parse_coordinate(matching_waypoint['Latitude'])
            lon = parse_coordinate(matching_waypoint['Longitude'])

            tz_offset = parse_tz_offset(matching_tz_val)
            tz_str_formatted = format_tz_string(matching_tz_val, tz_offset)

            sunrise, sunset = calculate_sun_events_native(lat, lon, c_date, tz_offset)

            results.append({
                "CRUISE CODE": c_cruise,
                "DATE": c_date.strftime("%Y-%m-%d"),
                "PORT OF CALL": port_name,
                "TIME ZONE": tz_str_formatted,
                "SUNRISE": sunrise,
                "SUNSET": sunset,
                "FILE ROTTA": matched_filename,
                "STATUS": "OK",
            })

    # -----------------------------------------------------------------------------
    # OUTPUT TABELLA ED EXPORT EXCEL
    # -----------------------------------------------------------------------------
    if results:
        df_out = pd.DataFrame(results)

        st.subheader("📊 Risultati Calcolo Effemeridi Navali")
        st.dataframe(df_out, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Ephemeris")
        excel_data = output.getvalue()

        st.download_button(
            label="📥 Scarica Tabella Excel",
            data=excel_data,
            file_name=f'Naval_Ephemeris_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Nessuna rotta corrispondente trovata per la selezione. Assicurati di aver abbinato il file CSV alla crociera corretta.")

else:
    st.info("Carica il calendario e i file CSV dalla barra laterale per iniziare.")
