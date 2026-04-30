import re
import config


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalise_body(body: str):
    """Return (cleaned_body, lines_list) with consistent line endings."""
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = body.replace('\xa0', ' ')
    lines = [line.strip() for line in body.split('\n')]
    return body, lines


def get_email_type(subject: str) -> str:
    subject_lower = subject.lower()
    if config.SUBJECT_PAID.lower() in subject_lower:
        return "paidTickets"
    if config.SUBJECT_FLIGHT_PASS.lower() in subject_lower:
        return "flightPass"
    return ""


def already_processed(ws, entry_id: str) -> bool:
    from excel_writer import get_all_entry_ids
    return entry_id in get_all_entry_ids(ws)


def get_next_nonblank_after_label(lines: list, label: str) -> str:
    """Return the first non-blank line that appears after a line containing label."""
    found_label = False
    for line in lines:
        if found_label:
            if line:
                return line
        elif label in line:
            found_label = True
    return ""


# ---------------------------------------------------------------------------
# PNR
# ---------------------------------------------------------------------------

def extract_flight_pass_pnr(body: str) -> str:
    _, lines = _normalise_body(body)
    candidate = get_next_nonblank_after_label(lines, "Booking Reference:")
    candidate = candidate.upper().strip()
    if re.fullmatch(r'[A-Z0-9]{6}', candidate):
        return candidate
    return ""


# ---------------------------------------------------------------------------
# Aeroplan
# ---------------------------------------------------------------------------

def extract_aeroplan_from_block(block: str) -> str:
    _, block_lines = _normalise_body(block)
    candidate = get_next_nonblank_after_label(block_lines, "Air Canada Aeroplan:")
    candidate = candidate.strip()
    if candidate.isdigit():
        return candidate
    return ""


# ---------------------------------------------------------------------------
# Flight number / segment helpers
# ---------------------------------------------------------------------------

def is_flight_number_line(line: str) -> bool:
    return bool(re.fullmatch(r'AC\d{1,4}', line))


def extract_six_field_block(lines: list, start_index: int) -> list:
    """
    From start_index+1, collect the next 6 non-blank lines.
    Returns [dep_airport, dep_date, dep_time, arr_airport, arr_date, arr_time]
    or [] if not enough lines.
    """
    collected = []
    i = start_index + 1
    while i < len(lines) and len(collected) < 6:
        if lines[i]:
            collected.append(lines[i])
        i += 1
    return collected if len(collected) == 6 else []


def simplify_airport_name(airport_text: str) -> str:
    if '(' in airport_text:
        return airport_text[:airport_text.index('(')].strip()
    return airport_text.strip()


def is_montreal_airport(airport_text: str) -> bool:
    text_lower = airport_text.lower()
    return any(x in text_lower for x in ("montreal", "montréal", "yul"))


def format_flight_date(raw_date: str) -> str:
    """
    Input like "Thursday 15-Apr-2025 09:45".
    VBA splits on space → index 1 → "15-Apr-2025", splits on "-" → ["15","Apr","2025"].
    Returns "Apr 15".
    """
    try:
        parts = raw_date.split()
        if len(parts) >= 2:
            day_part = parts[1]          # "15-Apr-2025"
        else:
            day_part = parts[0]
        segments = day_part.split('-')   # ["15","Apr","2025"]
        return f"{segments[1]} {segments[0]}"
    except (IndexError, AttributeError):
        return raw_date


def _build_segment_text(dep_airport: str, dep_date: str, dep_time: str,
                         arr_airport: str, arr_date: str, arr_time: str) -> str:
    dep_simple = simplify_airport_name(dep_airport)
    arr_simple = simplify_airport_name(arr_airport)
    date_str   = format_flight_date(dep_date)
    return f"{date_str} at {dep_time} ({dep_simple}) - {arr_time} ({arr_simple})"


# ---------------------------------------------------------------------------
# First departure airport
# ---------------------------------------------------------------------------

def extract_first_departure_airport(body: str) -> str:
    _, lines = _normalise_body(body)
    for i, line in enumerate(lines):
        if is_flight_number_line(line):
            for j in range(i + 1, len(lines)):
                if lines[j]:
                    return simplify_airport_name(lines[j])
    return ""


# ---------------------------------------------------------------------------
# Segment groups
# ---------------------------------------------------------------------------

def extract_trip_segment_groups(body: str) -> tuple:
    """Returns (outbound_text, return_text) with segments joined by newlines."""
    _, lines = _normalise_body(body)
    outbound_segments = []
    return_segments   = []
    reached_montreal  = False

    for i, line in enumerate(lines):
        if is_flight_number_line(line):
            block = extract_six_field_block(lines, i)
            if not block:
                continue
            dep_airport, dep_date, dep_time, arr_airport, arr_date, arr_time = block
            segment_text = _build_segment_text(dep_airport, dep_date, dep_time,
                                               arr_airport, arr_date, arr_time)
            if not reached_montreal:
                outbound_segments.append(segment_text)
                if is_montreal_airport(arr_airport):
                    reached_montreal = True
            else:
                return_segments.append(segment_text)

    return "\n".join(outbound_segments), "\n".join(return_segments)


# ---------------------------------------------------------------------------
# Montreal times
# ---------------------------------------------------------------------------

def extract_montreal_times(body: str) -> tuple:
    """Returns (montreal_arrival_time, montreal_departure_time)."""
    _, lines = _normalise_body(body)
    montreal_arrival   = ""
    montreal_departure = ""

    for i, line in enumerate(lines):
        if is_flight_number_line(line):
            block = extract_six_field_block(lines, i)
            if not block:
                continue
            dep_airport, _, dep_time, arr_airport, _, arr_time = block
            if not montreal_arrival and is_montreal_airport(arr_airport):
                montreal_arrival = arr_time
            if not montreal_departure and is_montreal_airport(dep_airport):
                montreal_departure = dep_time

    return montreal_arrival, montreal_departure


# ---------------------------------------------------------------------------
# Flight pass product info
# ---------------------------------------------------------------------------

def extract_flight_pass_info(body: str) -> tuple:
    """Returns (product_name, credits_per_pax)."""
    _, lines = _normalise_body(body)

    # Count passengers: lines matching  "1:Name:Ticket Number:"
    pax_count = sum(
        1 for line in lines
        if re.search(r'^\d+\s*:.*:\s*Ticket Number:', line)
    )
    if pax_count == 0:
        pax_count = 1

    product_name   = ""
    credits_per_pax = ""

    for i, line in enumerate(lines):
        if "Flight Credit Summary" in line:
            non_blank = []
            for j in range(i + 1, len(lines)):
                if lines[j]:
                    non_blank.append(lines[j])
                if len(non_blank) == 2:
                    break
            if len(non_blank) >= 1:
                product_name = non_blank[0]
            if len(non_blank) >= 2:
                m = re.search(r'\d+', non_blank[1])
                if m:
                    total_credits = int(m.group())
                    credits_per_pax = str(total_credits // pax_count)
            break

    return product_name, credits_per_pax


# ---------------------------------------------------------------------------
# Passenger parsing
# ---------------------------------------------------------------------------

def is_passenger_header_line(line: str) -> bool:
    return bool(re.search(r'^\d+\s*:\s*.+\s*:\s*Ticket Number:', line))


def extract_passenger_name(header_line: str) -> str:
    m = re.search(r'^\d+\s*:\s*(.*?)\s*:\s*Ticket Number:', header_line)
    if m:
        return m.group(1).strip()
    return ""


def parse_passengers_from_body(body: str, pnr: str, entry_id: str,
                                first_dep: str, outbound: str, return_seg: str,
                                mtl_arrival: str, mtl_departure: str,
                                fp_product: str, credits_per_pax: str) -> list:
    """
    Scan for passenger header lines, accumulate blocks, return list of row dicts.
    """
    _, lines = _normalise_body(body)

    rows          = []
    current_block = []
    current_header = None

    def flush_block(header_line, block_lines):
        if header_line is None:
            return
        name     = extract_passenger_name(header_line)
        aeroplan = extract_aeroplan_from_block("\n".join(block_lines))
        rows.append({
            "EntryID":               entry_id,
            "PNR":                   pnr,
            "PassengerName":         name,
            "AeroplanNumber":        aeroplan,
            "FirstDepartureAirport": first_dep,
            "OutboundSegments":      outbound,
            "ReturnSegments":        return_seg,
            "MontrealArrivalTime":   mtl_arrival,
            "MontrealDepartureTime": mtl_departure,
            "FlightPassProduct":     fp_product,
            "CreditsPerPassenger":   credits_per_pax,
            "Cost":                  "",
            "Type":                  "Flight Pass",
        })

    for line in lines:
        if is_passenger_header_line(line):
            flush_block(current_header, current_block)
            current_header = line
            current_block  = []
        else:
            current_block.append(line)

    flush_block(current_header, current_block)

    return rows


# ---------------------------------------------------------------------------
# Top-level entry point
# ---------------------------------------------------------------------------

def parse_flight_pass_email(body: str, entry_id: str) -> list:
    """Parse a flight-pass email body and return list of passenger row dicts."""
    pnr             = extract_flight_pass_pnr(body)
    first_dep       = extract_first_departure_airport(body)
    outbound, ret   = extract_trip_segment_groups(body)
    mtl_arr, mtl_dep = extract_montreal_times(body)
    fp_product, credits = extract_flight_pass_info(body)

    rows = parse_passengers_from_body(
        body, pnr, entry_id,
        first_dep, outbound, ret,
        mtl_arr, mtl_dep,
        fp_product, credits
    )
    return rows
