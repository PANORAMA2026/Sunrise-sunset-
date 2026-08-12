import io
import math
import re
from datetime import datetime, time, timedelta
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Naval Ephemeris Calculator", layout="wide"
)

st.title("⚓ Naval Ephemeris & Route Analyzer")
st.write(
    "Calcolo di Alba, Tramonto e Time Zone per rotte singole o ripetute nel"
    " calendario."
)

# -----------------------------------------------------------------------------
# UTILITY FUNCTIONS
# -----------------------------------------------------------------------------

def parse_coordinate(coord):
    """
    Converte coordinate nautiche (es. "33° 45.007' N" o "118° 11.209' W")
    o numeriche in gradi decimali trasformabili in float.
    """
    if pd.isna(coord):
        return 0.0
    if isinstance(coord, (int, float)):
        return float(coord)
    
    s = str(coord).strip()
    try:
        return float(s)
    except ValueError:
        pass
    
    # Formato Gradi° Minuti' Emisfero (es. 33° 45.007' N)
    match = re.search(r"(\d+)°\s*(\d+(?:\.\d+)?)\s*['′]?\s*([NSEWnsew]?)", s)
    if match:
        degrees = float(match.group(1))
        minutes = float(match.group(2))
        direction = match.group(3).upper()
        
        decimal_degrees = degrees + (minutes / 60.0)
        if direction in ['S', 'W']:
            decimal_degrees = -decimal_degrees
        return decimal_degrees
        
    # Formato con direzione finale senza gradi (es. 33.7501 N)
    match_simple = re.search(r"([+-]?\d+(?:\.\d+)?)\s*([NSEWnsew]?)", s)
    if match_simple:
        val = float(match_simple.group(1))
        direction = match_simple.group(2).upper()
        if direction in ['S', 'W']:
            val = -abs(val)
        return val
        
    return 0.0


def parse_tz_offset(tz_str):
    """Estrae l'offset orario numerico in ore da stringhe come '7 W', '8 W', '2 E'."""
    if not isinstance(tz_str, str):
        return 0.0
    match = re.search(r'(\d+(?:\.\d+)?)\s*([WEwe]?)', tz_str.strip())
    if not match:
        return 0.0
    val = float(match.group(1))
    direction = match.group(2).upper()
    return -val if direction == 'W' else val


def format_tz_string(tz_str):
    """Standardizza il formato della Time Zone per l'output (es. '7 W')."""
    if not isinstance(tz_str, str) or not tz_str.strip():
        return '0'
    match = re.search(r'(\d+(?:\.\d+)?)\s*([WEwe]?)', tz_str.strip())
    if not match:
        return tz_str.strip()
    val = match.group(1)
    direction = match.group(2).upper()
    return f'{val} {direction}'.strip()


def calculate_sun_events_native(lat, lon, date_obj, tz_offset_hours):
    """Calcola alba e tramonto in ora locale usando formule astronomiche standard (NOAA)."""
    try:
        day_of_year = date_obj.timetuple().tm_yday

        # Declinazione solare (radianti)
        declination = 0.409 * math.sin((2 * math.pi / 365) * (day_of_year - 81))

        lat_rad = math.radians(lat)

        # Angolo orario dell'alba/tramonto (-0.833° tiene conto di rifrazione e diametro)
        cos_ha = (
            math.sin(math.radians(-0.833))
            - math.sin(lat_rad) * math.sin(declination)
        ) / (math.cos(lat_rad) * math.cos(declination))

        if cos_ha >= 1:
            return "Mai sorge", "Mai sorge"
        if cos_ha <= -1:
            return "Mai tramonta", "Mai tramonta"

        ha = math.degrees(math.acos(cos_ha))

        # Equazione del tempo (minuti)
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
                })
        except Exception as e:
            st.sidebar.error(f"Errore nella lettura di {rf.name}: {e}")

    if not routes_data:
        st.warning("Nessun file rotta valido elaborato.")
        st.stop()

    # -----------------------------------------------------------------------------
    # CONFIGURAZIONE ABBINAMENTO MULTIPLO
    # -----------------------------------------------------------------------------
    st.subheader("🎯 Abbinamento Rotte a una o più Crociere/Ricorrenze")

    cruise_options = (
        df_cal[cal_cruise_col].dropna().unique().tolist()
        if cal_cruise_col
        else df_cal['Date_Parsed'].unique().tolist()
    )

    mapped_routes_list = []

    for idx, r_info in enumerate(routes_data):
        col1, col2 = st.columns([2, 3])
        with col1:
            st.write(
                f"**Rotta `{r_info['filename']}`** ({r_info['num_days']} giorni)"
            )
        with col2:
            selected_cruises = st.multiselect(
                f"Seleziona tutte le Crociere/Ricorrenze per `{r_info['filename']}`:",
                options=cruise_options,
                default=[cruise_options[0]] if idx == 0 else [],
                key=f"multiselect_{idx}",
            )

            for cruise_id in selected_cruises:
                if cal_cruise_col:
                    start_date = df_cal[df_cal[cal_cruise_col] == cruise_id]['Date_Parsed'].min()
                else:
                    start_date = cruise_id

                mapped_routes_list.append({
                    'df': r_info['df'],
                    'cruise_id': cruise_id,
                    'start_cal_date': start_date,
                })

    only_mapped = st.checkbox(
        "Mostra solo i giorni con rotta caricata ed elaborata", value=False
    )

    # -----------------------------------------------------------------------------
    # ELABORAZIONE EFFEMERIDI
    # -----------------------------------------------------------------------------
    results = []

    for _, cal_row in df_cal.iterrows():
        c_date = cal_row['Date_Parsed']
        port_name = cal_row['Location_Clean']
        c_cruise = cal_row[cal_cruise_col] if cal_cruise_col else None

        matching_waypoint = None
        matching_tz = '0'

        for item in mapped_routes_list:
            r_df = item['df']
            r_start_date = item['start_cal_date']
            r_cruise_id = item['cruise_id']

            if cal_cruise_col and c_cruise != r_cruise_id:
                continue

            rel_day = (c_date - r_start_date).days

            if rel_day >= 0:
                day_points = r_df[r_df['rel_day'] == rel_day]
                if not day_points.empty:
                    mid_idx = len(day_points) // 2
                    matching_waypoint = day_points.iloc[mid_idx]

                    if 'Time Zone' in day_points.columns:
                        matching_tz = str(day_points.iloc[0]['Time Zone'])
                    break

        if matching_waypoint is not None:
            lat = parse_coordinate(matching_waypoint['Latitude'])
            lon = parse_coordinate(matching_waypoint['Longitude'])
            tz_offset = parse_tz_offset(matching_tz)
            tz_str_formatted = format_tz_string(matching_tz)

            sunrise, sunset = calculate_sun_events_native(lat, lon, c_date, tz_offset)
            status = "OK"
        else:
            if only_mapped:
                continue
            sunrise, sunset = "N/D", "N/D"
            tz_str_formatted = "N/D"
            status = "N/D (Rotta non caricata)"

        results.append({
            "DATE": c_date.strftime("%Y-%m-%d"),
            "PORT OF CALL": port_name,
            "TIME ZONE": tz_str_formatted,
            "SUNRISE": sunrise,
            "SUNSET": sunset,
            "STATUS": status,
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
        st.info("Nessun dato da mostrare con i filtri selezionati.")

else:
    st.info("Carica il calendario e i file CSV per iniziare.")
