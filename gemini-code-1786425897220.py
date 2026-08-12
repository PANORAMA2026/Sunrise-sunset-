import io
import math
import re
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
from geopy.geocoders import Nominatim

st.set_page_config(
    page_title="Naval Ephemeris & Automatic Route Mapper", layout="wide"
)

st.title("⚓ Naval Ephemeris & Automatic Route Mapper")
st.write(
    "Il sistema riconosce automaticamente i porti di partenza e arrivo dalle coordinate del CSV e li abbina al calendario senza intervento manuale."
)

# -----------------------------------------------------------------------------
# UTILITY FUNCTIONS & GEOCODING
# -----------------------------------------------------------------------------

@st.cache_data(ttl=86400)
def reverse_geocode_port(lat, lon):
    """Converte Lat/Lon nel nome della città/porto tramite sistema di mappe."""
    try:
        geolocator = Nominatim(user_agent="naval_ephemeris_app")
        location = geolocator.reverse((lat, lon), language="en", timeout=5)
        if location and location.raw.get("address"):
            addr = location.raw["address"]
            # Cerca la città, il porto, il comune o la regione
            port_name = (
                addr.get("city")
                or addr.get("town")
                or addr.get("port")
                or addr.get("county")
                or addr.get("state")
                or ""
            )
            return port_name.lower().strip()
    except Exception:
        pass
    return ""


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
        if direction in ["S", "W"]:
            decimal_degrees = -decimal_degrees
        return decimal_degrees

    match_simple = re.search(r"([+-]?\d+(?:\.\d+)?)\s*([NSEWnsew]?)", s)
    if match_simple:
        val = float(match_simple.group(1))
        direction = match_simple.group(2).upper()
        if direction in ["S", "W"]:
            val = -abs(val)
        return val

    return 0.0


def parse_tz_offset(tz_val):
    """Estrae l'offset orario numerico dal CSV."""
    if pd.isna(tz_val):
        return 0.0
    s = str(tz_val).strip()
    if not s or s.lower() == "nan":
        return 0.0

    match_hhmmss = re.search(r"([+-]?)\s*(\d{1,2}):(\d{2})(?::(\d{2}))?", s)
    if match_hhmmss:
        sign = -1.0 if match_hhmmss.group(1) == "-" else 1.0
        hours = float(match_hhmmss.group(2))
        minutes = float(match_hhmmss.group(3))
        return sign * (hours + minutes / 60.0)

    match_we = re.search(r"(\d+(?:\.\d+)?)\s*([WEwe])", s)
    if match_we:
        val = float(match_we.group(1))
        direction = match_we.group(2).upper()
        return -val if direction == "W" else val

    try:
        return float(s)
    except ValueError:
        return 0.0


def format_tz_string(tz_val, offset_hours):
    """Formatta la Time Zone (es. '7 W' o '2 E')."""
    s = str(tz_val).strip()
    if not s or s.lower() == "nan":
        return "0"

    match_we = re.search(r"(\d+(?:\.\d+)?)\s*([WEwe])", s)
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
        declination = 0.409 * math.sin(
            (2 * math.pi / 365) * (day_of_year - 81)
        )
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

        return hours_to_time_str(sunrise_local_hours), hours_to_time_str(
            sunset_local_hours
        )
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
        (
            c
            for c in df_cal.columns
            if "date" in c.lower() or "data" in c.lower()
        ),
        df_cal.columns[0],
    )
    cal_port_col = next(
        (
            c
            for c in df_cal.columns
            if "location" in c.lower() or "port" in c.lower()
        ),
        df_cal.columns[1],
    )
    cal_cruise_col = next(
        (
            c
            for c in df_cal.columns
            if "cruise" in c.lower() or "crociera" in c.lower()
        ),
        None,
    )

    df_cal["Date_Parsed"] = pd.to_datetime(df_cal[cal_date_col]).dt.date
    df_cal["Location_Clean"] = df_cal[cal_port_col].astype(str).str.strip()

    st.subheader("🔍 Analisi Automatica e Mappatura delle Rotte")

    results = []

    for rf in route_files:
        try:
            df_r = pd.read_csv(rf)

            time_col = next(
                c
                for c in df_r.columns
                if "arrival time" in c.lower() or "time" in c
            )
            lat_col = next(c for c in df_r.columns if "lat" in c.lower())
            lon_col = next(c for c in df_r.columns if "lon" in c.lower())
            name_col = next(
                (
                    c
                    for c in df_r.columns
                    if "name" in c.lower() or "waypoint" in c.lower()
                ),
                None,
            )
            tz_col = next(
                (
                    c
                    for c in df_r.columns
                    if any(
                        k in c.lower()
                        for k in [
                            "utc offset",
                            "time zone",
                            "timezone",
                            "offset",
                            "tz",
                        ]
                    )
                ),
                None,
            )

            df_r["dt_parsed"] = pd.to_datetime(df_r[time_col], dayfirst=True)
            df_r["date_only"] = df_r["dt_parsed"].dt.date
            min_date = df_r["date_only"].min()
            df_r["rel_day"] = df_r["date_only"].apply(
                lambda d: (d - min_date).days if pd.notnull(d) else 0
            )

            # ESTREMI DELLA ROTTA (Start e End Waypoints)
            start_wp = df_r.iloc[0]
            end_wp = df_r.iloc[-1]

            start_lat, start_lon = parse_coordinate(
                start_wp[lat_col]
            ), parse_coordinate(start_wp[lon_col])
            end_lat, end_lon = parse_coordinate(
                end_wp[lat_col]
            ), parse_coordinate(end_wp[lon_col])

            # MAPPATURA AUTOMATICA VIA MAPPE / GEOPY
            start_geo_name = reverse_geocode_port(start_lat, start_lon)
            end_geo_name = reverse_geocode_port(end_lat, end_lon)

            # Estrazione parole dal nome waypoint se disponibili
            start_csv_name = (
                str(start_wp[name_col]).lower() if name_col else ""
            )
            end_csv_name = str(end_wp[name_col]).lower() if name_col else ""

            st.write(
                f"📄 **Rotta CSV:** `{rf.name}`  \n"
                f"📍 *Origine Riconosciuta:* {start_geo_name.upper() or start_csv_name.upper()} | "
                f"📍 *Destinazione Riconosciuta:* {end_geo_name.upper() or end_csv_name.upper()}"
            )

            # RICERCA AUTOMATICA NEL CALENDARIO EXCEL
            # Identifichiamo le crociere il cui porto iniziale e finale coincidono con la rotta
            matching_cruises = []

            if cal_cruise_col:
                for cruise_id, group in df_cal.groupby(cal_cruise_col):
                    ports_in_cruise = (
                        group["Location_Clean"].str.lower().tolist()
                    )
                    cruise_start_port = (
                        ports_in_cruise[0] if ports_in_cruise else ""
                    )
                    cruise_end_port = (
                        ports_in_cruise[-1] if ports_in_cruise else ""
                    )

                    # Verifichiamo se i porti coincidono con le coordinate verificate
                    start_match = (
                        start_geo_name in cruise_start_port
                        or any(
                            w in cruise_start_port
                            for w in start_csv_name.split()
                            if len(w) > 3
                        )
                    )
                    end_match = (
                        end_geo_name in cruise_end_port
                        or any(
                            w in cruise_end_port
                            for w in end_csv_name.split()
                            if len(w) > 3
                        )
                    )

                    if start_match or end_match:
                        start_date = group["Date_Parsed"].min()
                        matching_cruises.append((str(cruise_id), start_date))

            if not matching_cruises:
                st.warning(
                    f"⚠️ Nessun abbinamento automatico trovato per `{rf.name}` nel calendario. "
                    f"Controlla i nomi dei porti nel calendario Excel."
                )

            # APPLICAZIONE AUTOMATICA PER OGNI CROCIERA TROVATA
            for c_code, start_date in matching_cruises:
                for rel_day, day_group in df_r.groupby("rel_day"):
                    actual_date = start_date + timedelta(days=int(rel_day))
                    mid_point = day_group.iloc[len(day_group) // 2]

                    lat = parse_coordinate(mid_point[lat_col])
                    lon = parse_coordinate(mid_point[lon_col])

                    wp_name = (
                        str(mid_point[name_col])
                        if name_col and pd.notna(mid_point[name_col])
                        else f"Waypoint Giorno {rel_day}"
                    )

                    tz_raw = mid_point[tz_col] if tz_col else "0"
                    tz_offset = parse_tz_offset(tz_raw)
                    tz_str_formatted = format_tz_string(tz_raw, tz_offset)

                    sunrise, sunset = calculate_sun_events_native(
                        lat, lon, actual_date, tz_offset
                    )

                    results.append({
                        "CODICE CROCIERA": c_code,
                        "DATA ESECUZIONE": actual_date.strftime("%Y-%m-%d"),
                        "GIORNO ROTTA": f"Giorno {rel_day + 1}",
                        "WAYPOINT / NOME": wp_name,
                        "LATITUDINE": lat,
                        "LONGITUDINE": lon,
                        "TIME ZONE": tz_str_formatted,
                        "ALBA (Local Time)": sunrise,
                        "TRAMONTO (Local Time)": sunset,
                        "FILE ROTTA ORIGINE": rf.name,
                    })

        except Exception as e:
            st.error(f"Errore durante la lettura del file {rf.name}: {e}")

    # -----------------------------------------------------------------------------
    # TABELLA FINALE ED EXPORT EXCEL
    # -----------------------------------------------------------------------------
    if results:
        df_out = pd.DataFrame(results)
        df_out = df_out.sort_values(by=["DATA ESECUZIONE", "CODICE CROCIERA"])

        st.markdown("---")
        st.subheader("📊 Risultati Calcolo Effemeridi Automatica")
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
else:
    st.info("Carica il calendario Excel e i file CSV di rotta per avviare il riconoscimento automatico.")
