import re


def _normalise_body(body: str):
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = body.replace('\xa0', ' ')
    lines = [line.strip() for line in body.split('\n')]
    return body, lines


# ---------------------------------------------------------------------------
# PNR
# ---------------------------------------------------------------------------

def extract_paid_pnr(body: str, subject: str = "") -> str:
    # Try embedded marker first
    m = re.search(r'BOOKING_REFERENCE_START([A-Z0-9]{6})BOOKING_REFERENCE_END', body)
    if m:
        return m.group(1)
    # Fall back to subject line
    m = re.search(r'\b([A-Z0-9]{6})\b', subject)
    if m:
        return m.group(1)
    return ""


# ---------------------------------------------------------------------------
# Aeroplan
# ---------------------------------------------------------------------------

def extract_paid_aeroplan(body: str) -> str:
    m = re.search(r'Aeroplan\s*#:\s*(\d+)', body)
    if m:
        return m.group(1)
    return ""


# ---------------------------------------------------------------------------
# Cost
# ---------------------------------------------------------------------------

def extract_paid_ticket_cost(body: str) -> str:
    _, lines = _normalise_body(body)
    in_total = False
    for line in lines:
        if "Grand total" in line:
            in_total = True
            continue
        if in_total and line:
            m = re.search(r'CAD\s*\$\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})', line)
            if m:
                return m.group(0)
    return ""


# ---------------------------------------------------------------------------
# Airport helpers
# ---------------------------------------------------------------------------

def is_paid_airport_code_line(line: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-zÀ-ÿ' .\-]+\s+[A-Z]{3}", line))


def simplify_paid_code_airport(line: str) -> str:
    m = re.match(r'^(.*)\s+[A-Z]{3}$', line)
    if m:
        return m.group(1).strip()
    return line.strip()


def is_paid_flight_number_line(line: str) -> bool:
    return bool(re.match(r'^AC\s*\d{1,4}', line))


def _is_montreal_airport(text: str) -> bool:
    t = text.lower()
    return any(x in t for x in ("montreal", "montréal", "yul"))


def _is_time_line(line: str) -> bool:
    return bool(re.fullmatch(r'\d{2}:\d{2}', line))


# ---------------------------------------------------------------------------
# Segment extraction — Format A (airport code lines like "Montreal Trudeau YUL")
# ---------------------------------------------------------------------------

def extract_paid_segments_code_line_format(body: str) -> list:
    """
    Collect segments from the Flights…Passengers section.
    Handles lines like "Montreal Trudeau YUL".
    """
    _, lines = _normalise_body(body)

    in_flights = False
    segments   = []

    dep_airport = arr_airport = dep_date = dep_time = arr_time = ""
    times_found = 0

    for line in lines:
        if line == "Flights":
            in_flights = True
            continue
        if line == "Passengers":
            break
        if not in_flights:
            continue

        if "Departure •" in line:
            dep_date    = line
            dep_airport = arr_airport = dep_time = arr_time = ""
            times_found = 0
            continue
        if "Return •" in line:
            dep_date    = line
            dep_airport = arr_airport = dep_time = arr_time = ""
            times_found = 0
            continue

        if is_paid_airport_code_line(line):
            name = simplify_paid_code_airport(line)
            if not dep_airport:
                dep_airport = name
            elif not arr_airport:
                arr_airport = name
            continue

        if _is_time_line(line):
            times_found += 1
            if times_found == 1:
                dep_time = line
            elif times_found == 2:
                arr_time = line
            continue

        if is_paid_flight_number_line(line):
            if dep_airport and arr_airport and dep_time and arr_time:
                segments.append({
                    "dep":      dep_airport,
                    "arr":      arr_airport,
                    "date":     dep_date,
                    "dep_time": dep_time,
                    "arr_time": arr_time,
                })
            dep_airport = arr_airport = dep_time = arr_time = ""
            times_found = 0

    return segments


# ---------------------------------------------------------------------------
# Segment extraction — Format B (embedded marker tokens)
# ---------------------------------------------------------------------------

def extract_paid_segments_marker_format(body: str) -> list:
    """
    Collect segments using embedded tokens like DEPARTURE_LOCATIONCODE_START.
    """
    _, lines = _normalise_body(body)

    in_flights = False
    segments   = []

    dep_airport = arr_airport = dep_date = dep_time = arr_time = ""
    times_found = 0

    for line in lines:
        if line == "Flights":
            in_flights = True
            continue
        if line == "Passengers":
            break
        if not in_flights:
            continue

        if "Departure •" in line or "Return •" in line:
            dep_date    = line
            dep_airport = arr_airport = dep_time = arr_time = ""
            times_found = 0
            continue

        if "DEPARTURE_LOCATIONCODE_START" in line:
            dep_airport = line[:line.index("DEPARTURE_LOCATIONCODE_START")].strip()
            continue
        if "ARRIVAL_LOCATIONCODE_START" in line:
            arr_airport = line[:line.index("ARRIVAL_LOCATIONCODE_START")].strip()
            continue

        if _is_time_line(line):
            times_found += 1
            if times_found == 1:
                dep_time = line
            elif times_found == 2:
                arr_time = line
            continue

        if "FLIGHT_NUMBER_STARTAC" in line:
            if dep_airport and arr_airport and dep_time and arr_time:
                segments.append({
                    "dep":      dep_airport,
                    "arr":      arr_airport,
                    "date":     dep_date,
                    "dep_time": dep_time,
                    "arr_time": arr_time,
                })
            dep_airport = arr_airport = dep_time = arr_time = ""
            times_found = 0

    return segments


# ---------------------------------------------------------------------------
# Combined segment extractor (try Format A, fall back to Format B)
# ---------------------------------------------------------------------------

def extract_paid_segments(body: str) -> list:
    segments = extract_paid_segments_code_line_format(body)
    if not segments:
        segments = extract_paid_segments_marker_format(body)
    return segments


# ---------------------------------------------------------------------------
# Trip segment text
# ---------------------------------------------------------------------------

def _format_segment_text(seg: dict) -> str:
    date_raw = seg.get("date", "")
    # Try to pull a clean date from "Departure • Thursday, April 15, 2025" style
    m = re.search(r'(\w+ \d+,?\s*\d{4})', date_raw)
    date_str = m.group(1) if m else date_raw
    return (f"{date_str} at {seg['dep_time']} ({seg['dep']}) "
            f"- {seg['arr_time']} ({seg['arr']})")


def extract_trip_segment_groups_paid(body: str) -> tuple:
    segments = extract_paid_segments(body)
    if not segments:
        return "", ""

    outbound_segments = []
    return_segments   = []
    reached_montreal  = False
    starts_in_montreal = _is_montreal_airport(segments[0].get("dep", ""))

    for seg in segments:
        text = _format_segment_text(seg)
        if starts_in_montreal:
            return_segments.append(text)
        elif not reached_montreal:
            outbound_segments.append(text)
            if _is_montreal_airport(seg.get("arr", "")):
                reached_montreal = True
        else:
            return_segments.append(text)

    return "\n".join(outbound_segments), "\n".join(return_segments)


# ---------------------------------------------------------------------------
# Montreal times
# ---------------------------------------------------------------------------

def extract_montreal_times_paid(body: str) -> tuple:
    segments         = extract_paid_segments(body)
    montreal_arrival = ""
    montreal_departure = ""

    for seg in segments:
        if not montreal_arrival and _is_montreal_airport(seg.get("arr", "")):
            montreal_arrival = seg["arr_time"]
        if not montreal_departure and _is_montreal_airport(seg.get("dep", "")):
            montreal_departure = seg["dep_time"]

    return montreal_arrival, montreal_departure


# ---------------------------------------------------------------------------
# First departure airport
# ---------------------------------------------------------------------------

def extract_first_departure_airport_paid(body: str) -> str:
    segments = extract_paid_segments(body)
    if segments:
        return segments[0].get("dep", "")
    return ""


# ---------------------------------------------------------------------------
# Passenger names
# ---------------------------------------------------------------------------

def extract_paid_passenger_names(body: str) -> list:
    names = re.findall(
        r'PASSENGER_NAME_START\s*(.*?)\s*PASSENGER_NAME_END',
        body
    )
    # Deduplicate while preserving order
    seen   = set()
    unique = []
    for name in names:
        if name not in seen:
            seen.add(name)
            unique.append(name)
    return unique


# ---------------------------------------------------------------------------
# Top-level parser
# ---------------------------------------------------------------------------

def parse_paid_passengers(body: str, pnr: str, entry_id: str,
                           subject: str = "") -> list:
    if not pnr:
        pnr = extract_paid_pnr(body, subject)

    first_dep          = extract_first_departure_airport_paid(body)
    outbound, ret      = extract_trip_segment_groups_paid(body)
    mtl_arr, mtl_dep   = extract_montreal_times_paid(body)
    cost               = extract_paid_ticket_cost(body)
    names              = extract_paid_passenger_names(body)

    if not names:
        names = [""]   # one row with blank name if no names found

    rows = []
    for name in names:
        rows.append({
            "EntryID":               entry_id,
            "PNR":                   pnr,
            "PassengerName":         name,
            "AeroplanNumber":        extract_paid_aeroplan(body),
            "FirstDepartureAirport": first_dep,
            "OutboundSegments":      outbound,
            "ReturnSegments":        ret,
            "MontrealArrivalTime":   mtl_arr,
            "MontrealDepartureTime": mtl_dep,
            "FlightPassProduct":     "",
            "CreditsPerPassenger":   "",
            "Cost":                  cost,
            "Type":                  "Paid Ticket",
        })

    return rows


def parse_paid_email(body: str, entry_id: str, subject: str = "") -> list:
    pnr = extract_paid_pnr(body, subject)
    return parse_paid_passengers(body, pnr, entry_id, subject)
