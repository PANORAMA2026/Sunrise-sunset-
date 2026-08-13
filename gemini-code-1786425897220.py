import io
import math
import re
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Naval Ephemeris & Automatic Route Mapper", layout="wide"
)

st.title("⚓ Naval Ephemeris & Automatic Route Mapper")
st.write(
    "Accoppiamento rigido di sequenza: la rotta viene applicata SOLO se l'origine e la destinazione rispettano la sequenza esatta e i giorni sul calendario."
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
    if not s or s.lower() in ["nan", "none"]:
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
    if not s or s.lower() in ["nan", "none"]:
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

        return hours_to_time_str(sunrise_local_hours), hours_to_time_str(
            sunset_local_hours
        )
    except Exception:
        return "N/D", "N/D"


def clean_text_val(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "null"]:
        return ""
    return s


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

    # Identificazione colonne Calendario
    cal_date_col = next(
        (c for c in df_cal.columns if "date" in c.lower() or "data" in c.lower()),
        df_cal.columns[0],
    )
    cal_port_col = next(
        (c for c in df_cal.columns if "location" in c.lower() or "port" in c.lower()),
        df_cal.columns[1],
    )
    cal_code_col = next(
        (c for c in df_cal.columns if "port code" in c.lower() or "code" in c.lower()),
        None,
    )
    cal_cruise_col = next(
        (c for c in df_cal.columns if "cruise" in c.lower() or "crociera" in c.lower()),
        None,
    )

    df_cal["Date_Parsed"] = pd.to_datetime(df_cal[cal_date_col]).dt.date
    df_cal["Location_Clean"] = df_cal[cal_port_col].astype(str).str.strip().str.upper()
    df_cal["Port_Code_Clean"] = (
        df_cal[cal_code_col].astype(str).str.strip().str.upper()
        if cal_code_col
        else ""
    )

    # Ordina il calendario per data
    df_cal = df_cal.sort_values(by="Date_Parsed").reset_index(drop=True)

    st.subheader("🔍 Analisi Sequenziale della Rotta")

    results = []

    for rf in route_files:
        try:
            df_r = pd.read_csv(rf)

            time_col = next(
                c for c in df_r.columns if "arrival time" in c.lower() or "time" in c
            )
            lat_col = next(c for c in df_r.columns if "lat" in c.lower())
            lon_col = next(c for c in df_r.columns if "lon" in c.lower())

            name_cols = [
                c for c in df_r.columns
                if any(k in c.lower() for k in ["name", "port", "location"])
                and not any(k in c.lower() for k in ["no", "num", "id"])
            ]
            name_col = name_cols[0] if name_cols else next((c for c in df_r.columns if "waypoint" in c.lower()), None)

            tz_col = next(
                (
                    c for c in df_r.columns
                    if any(k in c.lower() for k in ["utc offset", "time zone", "timezone", "offset", "tz"])
                ),
                None,
            )

            df_r["dt_parsed"] = pd.to_datetime(df_r[time_col], dayfirst=True)
            df_r["date_only"] = df_r["dt_parsed"].dt.date
            min_date = df_r["date_only"].min()
            df_r["rel_day"] = df_r["date_only"].apply(
                lambda d: (d - min_date).days if pd.notnull(d) else 0
            )

            # Durata in giorni della rotta CSV (es. da giorno 0 a giorno 3 = 3 giorni di differenza)
            route_duration_days = int(df_r["rel_day"].max() - df_r["rel_day"].min())

            # ESTRAZIONE CODICI/NOMI ORIGINE E DESTINAZIONE
            filename_port_codes = re.findall(r"\b[A-Z]{3}\b", rf.name.upper())

            start_code = filename_port_codes[0] if len(filename_port_codes) >= 1 else ""
            end_code = filename_port_codes[1] if len(filename_port_codes) >= 2 else ""

            start_wp = df_r.iloc[0]
            end_wp = df_r.iloc[-1]

            start_name_csv = clean_text_val(start_wp[name_col]).upper() if name_col else ""
            end_name_csv = clean_text_val(end_wp[name_col]).upper() if name_col else ""

            st.write(
                f"📄 **Rotta CSV:** `{rf.name}` | "
                f"**Origine:** `{start_code or start_name_csv}` $\\rightarrow$ **Destinazione:** `{end_code or end_name_csv}` | "
                f"**Durata Rotta:** {route_duration_days + 1} giorni"
            )

            # SEARCH SEQUENZIALE SUL CALENDARIO
            # Cerca nel calendario una data D (Origine) e una data D + route_duration_days (Destinazione)
            matched_sequence_dates = []

            for i in range(len(df_cal)):
                row_start = df_cal.iloc[i]
                date_start = row_start["Date_Parsed"]

                # Verifica se questa riga del calendario corrisponde al porto di Partenza
                start_match = (
                    (start_code and start_code == row_start["Port_Code_Clean"])
                    or (start_code and start_code in row_start["Location_Clean"])
                    or (start_name_csv and start_name_csv in row_start["Location_Clean"])
                )

                if start_match:
                    target_end_date = date_start + timedelta(days=route_duration_days)

                    # Cerca se nel calendario esiste una riga con la data esatta di arrivo e il porto di Destinazione
                    matching_end_rows = df_cal[df_cal["Date_Parsed"] == target_end_date]

                    for _, row_end in matching_end_rows.iterrows():
                        end_match = (
                            (end_code and end_code == row_end["Port_Code_Clean"])
                            or (end_code and end_code in row_end["Location_Clean"])
                            or (end_name_csv and end_name_csv in row_end["Location_Clean"])
                        )

                        if end_match:
                            cruise_id = (
                                str(row_start[cal_cruise_col])
                                if cal_cruise_col and pd.notna(row_start[cal_cruise_col])
                                else f"PTP_{date_start}"
                            )
                            matched_sequence_dates.append((cruise_id, date_start))

            if matched_sequence_dates:
                st.success(f"✅ Sequenza trovata nel calendario per le date: {[d.strftime('%Y-%m-%d') for _, d in matched_sequence_dates]}")
            else:
                st.warning("⚠️ Nessuna sequenza di date esatta trovata nel calendario per questo itinerario.")

            # GENERAZIONE EFFEMERIDI SOLO PER LE SEQUENZE CONVALIDATE
            for c_code, start_date in matched_sequence_dates:
                for rel_day, day_group in df_r.groupby("rel_day"):
                    actual_date = start_date + timedelta(days=int(rel_day))
                    mid_point = day_group.iloc[len(day_group) // 2]

                    lat = parse_coordinate(mid_point[lat_col])
                    lon = parse_coordinate(mid_point[lon_col])

                    wp_name = clean_text_val(mid_point[name_col]) if name_col else ""
                    if not wp_name:
                        wp_name = f"Waypoint Giorno {rel_day + 1}"

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
            st.error(f"Errore durante l'elaborazione di {rf.name}: {e}")

    # -----------------------------------------------------------------------------
    # TABELLA FINALE ED EXPORT EXCEL
    # -----------------------------------------------------------------------------
    if results:
        df_out = pd.DataFrame(results)
        df_out = df_out.drop_duplicates(subset=["CODICE CROCIERA", "DATA ESECUZIONE", "WAYPOINT / NOME"])
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
