import io
import math
import re
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Naval Ephemeris Calculator", layout="wide"
)

st.title("⚓ Naval Ephemeris & Route Analyzer")
st.write(
    "I file CSV di rotta guidano il calcolo delle effemeridi; il calendario Excel ne pianifica le date di esecuzione."
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
    "2. Carica File Rotta Master (CSV)", type=["csv"], accept_multiple_files=True
)

if cal_file and route_files:
    df_cal = pd.read_excel(cal_file)

    cal_date_col = next(
        (c for c in df_cal.columns if 'date' in c.lower() or 'data' in c.lower()),
        df_cal.columns[0],
    )
    cal_cruise_col = next(
        (c for c in df_cal.columns if 'cruise' in c.lower() or 'crociera' in c.lower()),
        None,
    )

    df_cal['Date_Parsed'] = pd.to_datetime(df_cal[cal_date_col]).dt.date

    # Estrazione delle date di partenza delle crociere dal calendario
    cruise_start_dates = {}
    if cal_cruise_col:
        grouped = df_cal.groupby(cal_cruise_col)['Date_Parsed'].min()
        for cruise_id, start_d in grouped.items():
            cruise_start_dates[str(cruise_id)] = start_d
    else:
        for idx, d in enumerate(sorted(df_cal['Date_Parsed'].unique())):
            cruise_start_dates[f"Crociera_{idx+1}"] = d

    # ELABORAZIONE DEI FILE CSV MASTER
    parsed_routes = []
    for rf in route_files:
        try:
            df_r = pd.read_csv(rf)

            time_col = next(
                c for c in df_r.columns if 'arrival time' in c.lower() or 'time' in c
            )
            lat_col = next(c for c in df_r.columns if 'lat' in c.lower())
            lon_col = next(c for c in df_r.columns if 'lon' in c.lower())
            name_col = next(
                (c for c in df_r.columns if 'name' in c.lower() or 'waypoint' in c.lower()),
                None
            )
            tz_col = next(
                (c for c in df_r.columns if any(k in c.lower() for k in ['utc offset', 'time zone', 'timezone', 'offset', 'tz'])),
                None
            )

            df_r['dt_parsed'] = pd.to_datetime(df_r[time_col], dayfirst=True)
            df_r['date_only'] = df_r['dt_parsed'].dt.date

            min_date = df_r['date_only'].min()
            # Calcolo dei giorni relativi dall'inizio della rotta master (0, 1, 2...)
            df_r['rel_day'] = df_r['date_only'].apply(lambda d: (d - min_date).days if pd.notnull(d) else 0)

            parsed_routes.append({
                'filename': rf.name,
                'df': df_r,
                'lat_col': lat_col,
                'lon_col': lon_col,
                'name_col': name_col,
                'tz_col': tz_col,
                'total_days': df_r['rel_day'].max() + 1
            })
        except Exception as e:
            st.sidebar.error(f"Errore nella lettura del file {rf.name}: {e}")

    if not parsed_routes:
        st.warning("Nessun file rotta valido elaborato.")
        st.stop()

    # -----------------------------------------------------------------------------
    # SELEZIONE ED ASSEGNAZIONE DELLE CROCIERE PER CIASCUNA ROTTA CSV
    # -----------------------------------------------------------------------------
    st.subheader("🎯 Assegnazione Rotte alle Crociere del Calendario")
    st.info(
        "Seleziona per quali crociere del calendario deve essere eseguito ogni file di rotta CSV."
    )

    results = []

    for r_idx, r_info in enumerate(parsed_routes):
        st.write(f"### 🚢 Rotta CSV: `{r_info['filename']}` ({r_info['total_days']} giorni di navigazione)")

        selected_cruises = st.multiselect(
            f"Esegui la rotta `{r_info['filename']}` per le seguenti crociere:",
            options=list(cruise_start_dates.keys()),
            default=[],
            key=f"cruise_select_{r_idx}",
            placeholder="Scegli le crociere in cui la nave compie questa rotta..."
        )

        # GENERAZIONE EFFEMERIDI DIRETTAMENTE DAI PUNTI DELLA ROTTA CSV
        df_r = r_info['df']

        for c_code in selected_cruises:
            start_date = cruise_start_dates[c_code]

            # Raggruppiamo i waypoint della rotta giorno per giorno
            for rel_day, day_group in df_r.groupby('rel_day'):
                actual_date = start_date + timedelta(days=int(rel_day))

                # Prendiamo il punto mediano del giorno di navigazione
                mid_point = day_group.iloc[len(day_group) // 2]

                lat = parse_coordinate(mid_point[r_info['lat_col']])
                lon = parse_coordinate(mid_point[r_info['lon_col']])

                waypoint_name = (
                    str(mid_point[r_info['name_col']])
                    if r_info['name_col'] and pd.notna(mid_point[r_info['name_col']])
                    else f"Waypoint Giorno {rel_day}"
                )

                tz_raw = mid_point[r_info['tz_col']] if r_info['tz_col'] else '0'
                tz_offset = parse_tz_offset(tz_raw)
                tz_str_formatted = format_tz_string(tz_raw, tz_offset)

                sunrise, sunset = calculate_sun_events_native(
                    lat, lon, actual_date, tz_offset
                )

                results.append({
                    "CODICE CROCIERA": c_code,
                    "DATA": actual_date.strftime("%Y-%m-%d"),
                    "GIORNO ROTTA": f"Giorno {rel_day + 1}",
                    "WAYPOINT / NOME": waypoint_name,
                    "LATITUDINE": lat,
                    "LONGITUDINE": lon,
                    "TIME ZONE": tz_str_formatted,
                    "ALBA (Local Time)": sunrise,
                    "TRAMONTO (Local Time)": sunset,
                    "FILE ROTTA ORIGINE": r_info['filename'],
                })

    # -----------------------------------------------------------------------------
    # TABELLA FINALE ED EXPORT EXCEL
    # -----------------------------------------------------------------------------
    if results:
        df_out = pd.DataFrame(results)
        df_out = df_out.sort_values(by=["DATA", "CODICE CROCIERA"])

        st.markdown("---")
        st.subheader("📊 Risultati Effemeridi Generati dai File CSV Master")
        st.dataframe(df_out, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Ephemeris")
        excel_data = output.getvalue()

        st.download_button(
            label="📥 Scarica Report Excel Completato",
            data=excel_data,
            file_name=f'Naval_Ephemeris_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
