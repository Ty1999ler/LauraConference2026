import re
import config

# Canadian province/territory codes used in Air Canada airport names
_PROVINCE = r'(?:ON|PQ|QC|BC|AB|MB|SK|NS|NB|PE|NL|NF|YT|NT|NU)'


# ── Normalisation ─────────────────────────────────────────────────────────

def _normalise(body: str) -> list:
    body = body.replace('\r\n', '\n').replace('\r', '\n').replace('\xa0', ' ')
    result = []
    for line in body.split('\n'):
        for part in line.split('\t'):
            result.append(part.strip())
    return result


# ── Line type detectors ───────────────────────────────────────────────────

def _is_flight_number(line: str) -> bool:
    # Matches "AC410", "AC8356" etc — flight number only on the line
    return bool(re.match(r'^AC\d{1,4}\s*$', line))


def _is_airport_line(line: str) -> bool:
    # Matches "Toronto Pearson (ON)", "Montreal Trudeau (PQ)" etc
    return bool(re.search(rf'\({_PROVINCE}\)\s*$', line))


def _is_date_line(line: str) -> bool:
    # Matches "Mon 06-Jul 2026"
    return bool(re.match(r'^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\d{2}-\w{3}[\s-]\d{4}$', line))


def _is_time_line(line: str) -> bool:
    # Matches "12:00", "09:05"
    return bool(re.match(r'^\d{1,2}:\d{2}$', line))


def _is_montreal(text: str) -> bool:
    t = text.lower()
    return any(x in t for x in ('montreal', 'montréal', 'trudeau'))


# ── Formatting helpers ────────────────────────────────────────────────────

def _simplify_airport(text: str) -> str:
    """Remove province code: 'Toronto Pearson (ON)' -> 'Toronto Pearson'"""
    return re.sub(r'\s*\([A-Z]+\)\s*$', '', text).strip()


def _fmt_date(date_str: str) -> str:
    """'Mon 06-Jul 2026' -> 'Jul 06'"""
    try:
        parts = date_str.split()        # ["Mon", "06-Jul", "2026"]
        day, mon = parts[1].split('-')[:2]
        return f"{mon} {day}"
    except Exception:
        return date_str


def _fmt_segment(seg: dict) -> str:
    dep  = _simplify_airport(seg['dep_airport'])
    arr  = _simplify_airport(seg['arr_airport'])
    date = _fmt_date(seg['dep_date'])
    return f"{date} at {seg['dep_time']} ({dep}) -> {seg['arr_time']} ({arr})"


# ── Booking Reference ─────────────────────────────────────────────────────

def _extract_pnr(lines: list) -> str:
    for i, line in enumerate(lines):
        if 'Booking Reference:' in line:
            # Check same line first (e.g. "Booking Reference: B2ZT9L")
            rest = line.split('Booking Reference:')[-1].strip()
            if re.fullmatch(r'[A-Z0-9]{6}', rest):
                return rest
            # Then check next non-blank lines
            for j in range(i + 1, min(i + 6, len(lines))):
                if lines[j]:
                    candidate = lines[j].strip()
                    if re.fullmatch(r'[A-Z0-9]{6}', candidate):
                        return candidate
    return ''


# ── Flight Segment Extraction ─────────────────────────────────────────────

_SKIP_LINES = frozenset(['Flight', 'From', 'To', 'Stops', 'Fare Type', 'Meal',
                          'Flex,', 'M', '0', '1', '2'])


def _extract_segments(lines: list) -> list:
    """
    Extract all flight segments from the Flight Itinerary section.
    Returns list of dicts with dep_airport, dep_date, dep_time,
    arr_airport, arr_date, arr_time.
    """
    # Find section boundaries
    start = end = None
    for i, line in enumerate(lines):
        if line == 'Flight Itinerary':
            start = i
        if start is not None and line == 'Passenger Information':
            end = i
            break

    if start is None:
        return []

    section = lines[start:end] if end else lines[start:]
    segments = []

    i = 0
    while i < len(section):
        if _is_flight_number(section[i]):
            dep_airport = dep_date = dep_time = None
            arr_airport = arr_date = arr_time = None

            j = i + 1
            while j < len(section):
                line = section[j]

                if not line:
                    j += 1
                    continue

                # Skip "operated by" notes
                if 'is operated by' in line.lower():
                    j += 1
                    continue

                # Next flight number — stop collecting for current segment
                if _is_flight_number(line):
                    break

                # Skip table header/junk words
                if line in _SKIP_LINES or re.match(r'^Flex', line):
                    j += 1
                    continue

                # Collect fields in arrival order
                if _is_airport_line(line):
                    if dep_airport is None:
                        dep_airport = line
                    elif arr_airport is None:
                        arr_airport = line

                elif _is_date_line(line):
                    if dep_date is None:
                        dep_date = line
                    elif arr_date is None:
                        arr_date = line

                elif _is_time_line(line):
                    if dep_time is None:
                        dep_time = line
                    elif arr_time is None:
                        arr_time = line

                # All 6 fields collected — save segment
                if all([dep_airport, dep_date, dep_time,
                        arr_airport, arr_date, arr_time]):
                    segments.append({
                        'dep_airport': dep_airport,
                        'dep_date':    dep_date,
                        'dep_time':    dep_time,
                        'arr_airport': arr_airport,
                        'arr_date':    arr_date,
                        'arr_time':    arr_time,
                    })
                    i = j
                    break

                j += 1

        i += 1

    return segments


# ── Segment Classification ────────────────────────────────────────────────

def _classify_segments(segments: list) -> tuple:
    """
    Split segments into:
      inbound  → legs arriving IN Montreal  → ReturnSegments  (col G)
      outbound → legs departing FROM Montreal → OutboundSegments (col F)

    Multi-stop logic:
      All legs up to and including the Montreal-arrival leg = inbound
      All legs from the Montreal-departure leg onwards = outbound
    """
    mtl_arrival_idx   = None
    mtl_departure_idx = None

    for i, seg in enumerate(segments):
        if _is_montreal(seg['arr_airport']) and mtl_arrival_idx is None:
            mtl_arrival_idx = i
        if _is_montreal(seg['dep_airport']) and mtl_departure_idx is None:
            mtl_departure_idx = i

    inbound  = segments[:mtl_arrival_idx + 1]  if mtl_arrival_idx   is not None else []
    outbound = segments[mtl_departure_idx:]     if mtl_departure_idx is not None else []

    mtl_arrival_time   = (segments[mtl_arrival_idx]['arr_time']
                          if mtl_arrival_idx   is not None else '')
    mtl_departure_time = (segments[mtl_departure_idx]['dep_time']
                          if mtl_departure_idx is not None else '')

    return inbound, outbound, mtl_arrival_time, mtl_departure_time


# ── Passenger Extraction ──────────────────────────────────────────────────

def _extract_passengers(lines: list) -> list:
    """
    Return list of {name, aeroplan} dicts.
    Passenger header format: '1: Betty Freeburn  : Ticket Number:  0142...'
    Aeroplan format: 'Air Canada Aeroplan:' then number on next non-blank line.
    """
    start = None
    for i, line in enumerate(lines):
        if line == 'Passenger Information':
            start = i
            break

    if start is None:
        return []

    section  = lines[start:]
    passengers = []

    for i, line in enumerate(section):
        m = re.match(r'^\d+\s*:\s*(.*?)\s*:\s*Ticket Number:', line)
        if not m:
            continue

        name     = m.group(1).strip()
        aeroplan = ''

        for j in range(i + 1, min(i + 60, len(section))):
            # Stop if we've hit the next passenger's block
            if re.match(r'^\d+\s*:\s*.*:\s*Ticket Number:', section[j]):
                break
            if 'Air Canada Aeroplan:' not in section[j]:
                continue
            # Number may be on the same line or the next non-blank line
            candidates = [section[j].split('Air Canada Aeroplan:')[-1].strip()]
            candidates += [section[k] for k in range(j + 1, min(j + 10, len(section)))]
            for c in candidates:
                digits = re.sub(r'\D', '', c)
                if len(digits) == 9:
                    aeroplan = digits
                    break
            break

        passengers.append({'name': name, 'aeroplan': aeroplan})

    return passengers


# ── Flight Credit Summary ─────────────────────────────────────────────────

def _extract_credit_info(lines: list, pax_count: int) -> tuple:
    """Returns (product_name, credits_per_pax_str)."""
    for i, line in enumerate(lines):
        if line != 'Flight Credit Summary':
            continue

        product       = ''
        total_credits = 0

        for j in range(i + 1, min(i + 10, len(lines))):
            candidate = lines[j]
            if not candidate:
                continue

            m = re.match(r'^(\d+)\s+Flight Credits?$', candidate)
            if m:
                total_credits = int(m.group(1))
            elif not product:
                product = candidate

            if product and total_credits:
                break

        pax = max(pax_count, 1)
        per_pax = (total_credits // pax) if total_credits else ''
        return product, per_pax

    return '', ''


# ── Public API ────────────────────────────────────────────────────────────

def get_email_type(subject: str) -> str:
    s = subject.lower()
    if config.SUBJECT_PAID.lower() in s:
        return 'paidTickets'
    if config.SUBJECT_FLIGHT_PASS.lower() in s:
        return 'flightPass'
    return ''


def parse_flight_pass_email(body: str, entry_id: str) -> list:
    lines = _normalise(body)

    pnr      = _extract_pnr(lines)
    segments = _extract_segments(lines)

    inbound, outbound, mtl_arrival, mtl_departure = _classify_segments(segments)

    # inbound  = legs heading TO   Montreal = outbound to the conference
    # outbound = legs leaving FROM Montreal = returning home
    outbound_text = '\n'.join(_fmt_segment(s) for s in inbound)
    return_text   = '\n'.join(_fmt_segment(s) for s in outbound)
    first_dep     = _simplify_airport(segments[0]['dep_airport']) if segments else ''

    passengers = _extract_passengers(lines)
    if not passengers:
        passengers = [{'name': '', 'aeroplan': ''}]

    pax_count = len(passengers)
    fp_product, credits_per_pax = _extract_credit_info(lines, pax_count)

    rows = []
    for pax in passengers:
        rows.append({
            'EntryID':               entry_id,
            'PNR':                   pnr,
            'PassengerName':         pax['name'],
            'AeroplanNumber':        pax['aeroplan'],
            'FirstDepartureAirport': first_dep,
            'OutboundSegments':      outbound_text,
            'ReturnSegments':        return_text,
            'MontrealArrivalTime':   mtl_arrival,
            'MontrealDepartureTime': mtl_departure,
            'FlightPassProduct':     fp_product,
            'CreditsPerPassenger':   credits_per_pax,
            'Cost':                  '',
            'Type':                  'Flight Pass',
        })

    return rows
