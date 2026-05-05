"""
Quick sanity-check for the paid-tickets parser.

Enter 0 to use the hardcoded test body below.
Enter -1 or X to list all paid-ticket emails with their number and subject.
Enter 1, 2, 3 ... to pull that email (1-indexed) from the Outlook folder.
"""
import sys
from parse_paid_tickets import (
    _normalise_body,
    extract_paid_pnr,
    extract_paid_segments,
    extract_paid_segments_code_line_format,
    extract_paid_segments_marker_format,
    extract_paid_passenger_names,
    extract_paid_ticket_cost,
    extract_montreal_times_paid,
    extract_trip_segment_groups_paid,
)

BODY = """\
"""   # paste a sample paid-ticket email body here if needed


def _get_outlook_items():
    import config
    from outlook_connector import get_outlook_folder, get_folder_items
    from parse_flight_pass import get_email_type
    folder = get_outlook_folder(config.FOLDER_PATH)
    items = [m for m in get_folder_items(folder)
             if get_email_type(m.Subject or "") == "paidTickets"]
    return items


def _list_all_emails():
    """List every email in the folder with detected type, so UNKNOWN subjects are visible."""
    import config
    from outlook_connector import get_outlook_folder, get_folder_items
    from parse_flight_pass import get_email_type
    folder = get_outlook_folder(config.FOLDER_PATH)
    all_items = list(get_folder_items(folder))
    if not all_items:
        print("No emails found in folder.")
        return
    print(f"{'#':<5}  {'Type':<14}  Subject")
    print("-" * 85)
    for i, m in enumerate(all_items, 1):
        subj  = m.Subject or "(no subject)"
        etype = get_email_type(subj) or "UNKNOWN"
        print(f"{i:<5}  {etype:<14}  {subj[:60]}")


def _body_from_outlook(index: int = 0) -> tuple:
    items = _get_outlook_items()
    if not items:
        print("No paid-ticket emails found in folder.")
        sys.exit(1)
    mail = items[index]
    print(f"Using email [{index + 1}]: {mail.Subject[:70]}\n")
    return mail.Body or "", mail.Subject or ""


try:
    choice = input("Enter 0 for test body, -1 or X to list emails, or 1/2/3... for email from Outlook: ").strip()
    if choice == "0" or choice == "":
        body, subject = BODY, ""
    elif choice == "-1" or choice.upper() == "X":
        _list_all_emails()
        input("\nPress Enter to close...")
        sys.exit(0)
    else:
        body, subject = _body_from_outlook(int(choice) - 1)

    _, lines = _normalise_body(body)

    # ── dump Flights section ────────────────────────────────────────────────
    flights_start = passengers_start = None
    for i, l in enumerate(lines):
        if l == "Flights" and flights_start is None:
            flights_start = i
        if l == "Passengers" and passengers_start is None:
            passengers_start = i

    if flights_start is not None:
        end = passengers_start if passengers_start else min(flights_start + 80, len(lines))
        print(f"[Flights] lines {flights_start}-{end}:")
        for j in range(flights_start, end):
            print(f"  {j:>4}: {lines[j]!r}")
        print()
    else:
        print("[Flights] NOT FOUND — searching for nearby lines:")
        for i, l in enumerate(lines):
            if "flight" in l.lower() or "departure" in l.lower():
                print(f"  {i:>4}: {l!r}")
        print()

    if passengers_start is not None:
        print(f"[Passengers] lines {passengers_start}-{min(passengers_start + 30, len(lines))}:")
        for j in range(passengers_start, min(passengers_start + 30, len(lines))):
            print(f"  {j:>4}: {lines[j]!r}")
        print()
    else:
        print("[Passengers] NOT FOUND")
        print()

    # ── parse results ───────────────────────────────────────────────────────
    pnr = extract_paid_pnr(body, subject)
    print(f"PNR            : {pnr!r}")

    segs_a = extract_paid_segments_code_line_format(body)
    segs_b = extract_paid_segments_marker_format(body)
    segs   = extract_paid_segments(body)
    print(f"\nFormat A segments : {len(segs_a)}")
    for s in segs_a:
        print(f"  {s['date']!r:40}  {s['dep']!r} {s['dep_time']} -> {s['arr']!r} {s['arr_time']}")
    print(f"Format B segments : {len(segs_b)}")
    for s in segs_b:
        print(f"  {s['date']!r:40}  {s['dep']!r} {s['dep_time']} -> {s['arr']!r} {s['arr_time']}")
    print(f"Used             : {len(segs)} segment(s)")

    outbound, ret = extract_trip_segment_groups_paid(body)
    print(f"\nOutbound segments:\n  {outbound!r}")
    print(f"Return segments:\n  {ret!r}")

    mtl_arr, mtl_dep = extract_montreal_times_paid(body)
    print(f"\nMontreal arrival time   : {mtl_arr!r}")
    print(f"Montreal departure time : {mtl_dep!r}")

    names = extract_paid_passenger_names(body)
    print(f"\nPassenger names ({len(names)}):")
    for n in names:
        print(f"  {n!r}")

    cost = extract_paid_ticket_cost(body)
    print(f"\nCost : {cost!r}")

except Exception:
    import traceback
    traceback.print_exc()

input("\nPress Enter to close...")
