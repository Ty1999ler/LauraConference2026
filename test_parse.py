"""
Quick sanity-check for the flight pass parser.

Enter 0 to use the hardcoded test body below.
Enter 1, 2, 3 ... to pull that email (1-indexed) from the Outlook folder.
"""
import sys
from parse_flight_pass import (
    _normalise, _extract_pnr, _extract_segments,
    _classify_segments, _extract_passengers, _extract_credit_info,
    _fmt_segment,
)

BODY = """\
From: Air Canada <fpconfirmation@aircanada.ca>
Date: May 1, 2026 at 10:16:31 AM EDT
To: laura.mcalear@hotmail.com
Subject: Air Canada - Electronic Ticket Itinerary/Receipt





****** PLEASE DO NOT REPLY TO THIS E-MAIL ******

AC logo

Itinerary/Receipt
Electronic Ticketing confirmed. This is your official itinerary/receipt. We thank you for choosing Air Canada and look forward to welcoming you on board.




Seats have been pre-selected for you.
Click on the button below to see all seat options
and change your seat(s) to adjacent seats, if available.

Choose your seat(s)


Booking Information


Booking Reference:

BF6HVB

Main Contact:

lmcale@aseq.com
Mobile:

Manage My Booking (change, cancel, upgrade).
Change Seats get more seating options for my flight.
Purchase Maple Leaf Lounge Access / Meal Vouchers
Receive Flight Status Notifications directly to my email or mobile phone.
Check Flight Arrivals and Departures.
Check in online and print my boarding pass.
Customer Care
Air Canada
1-888-247-2262

Flight Arrivals and Departures
1-888-422-7533

Flight Itinerary


Flight

From

To

Stops

Fare Type

Meal

AC322

Calgary (AB)
Mon 06-Jul 2026
09:25

Montreal Trudeau (PQ)
Mon 06-Jul 2026
15:33

0

Flex,
M

AC323

Montreal Trudeau (PQ)
Thu 09-Jul 2026
13:30

Calgary (AB)
Thu 09-Jul 2026
16:08

0

Flex,
M

Passenger Information


1: Bennett Boyd  : Ticket Number:  0142326569908

Air Canada Aeroplan:

333900405

Meal Preference:

Normal

Seat Selection:

AC322 : 31D, AC323 : 31D

Special Needs:

None

2: Jeniffer onyedikachi Orajekwe  : Ticket Number:  0142326569909

Air Canada Aeroplan:

152910857

Meal Preference:

Normal

Seat Selection:

AC322 : No Seat Preferences, AC323 : No Seat Preferences

Special Needs:

None

3: Kayana Robinson  : Ticket Number:  0142326569910

Air Canada Aeroplan:

335665642

Meal Preference:

Vegetarian meal (lacto-ovo)

Seat Selection:

AC322 : No Seat Preferences, AC323 : No Seat Preferences

Special Needs:

None

Flight Credit Summary


Flexible Benefits East West Connector Fl

6 Flight Credits

Taxes, fees, and charges included
"""

def _all_outlook_items():
    import config
    from outlook_connector import get_outlook_folder, get_folder_items
    folder = get_outlook_folder(config.FOLDER_PATH)
    return list(get_folder_items(folder))


def _list_all_emails():
    from parse_flight_pass import get_email_type
    items = _all_outlook_items()
    if not items:
        print("No emails found in folder.")
        return
    print(f"{'#':<5}  {'Type':<14}  Subject")
    print("-" * 85)
    for i, m in enumerate(items, 1):
        subj  = m.Subject or "(no subject)"
        etype = get_email_type(subj) or "UNKNOWN"
        print(f"{i:<5}  {etype:<14}  {subj[:60]}")


def _mail_from_outlook(index: int):
    items = _all_outlook_items()
    if not items:
        print("No emails found in folder.")
        sys.exit(1)
    mail = items[index]
    print(f"Using email [{index + 1}]: {(mail.Subject or '')[:70]}\n")
    return mail


try:
    choice = input("Enter 0 for test body, -1 or X to list all emails, or 1/2/3... for email from Outlook: ").strip()

    if choice == "-1" or choice.upper() == "X":
        _list_all_emails()
        input("\nPress Enter to close...")
        sys.exit(0)

    from parse_flight_pass import get_email_type

    if choice == "0" or choice == "":
        body    = BODY
        subject = ""
        email_type = "flightPass"
    else:
        mail       = _mail_from_outlook(int(choice) - 1)
        body       = mail.Body or ""
        subject    = mail.Subject or ""
        email_type = get_email_type(subject) or "UNKNOWN"

    print(f"Email type: {email_type}\n")

    # ── always dump ALL normalised lines so we can see what the parser sees ──
    lines = _normalise(body)
    print(f"=== ALL LINES ({len(lines)} total) ===")
    for j, l in enumerate(lines):
        print(f"  {j:>4}: {l!r}")
    print()

    if email_type == "flightPass":
        # ── Flight Itinerary section ─────────────────────────────────────────
        fi_start = pi_start = None
        for i, l in enumerate(lines):
            if l == 'Flight Itinerary' and fi_start is None:
                fi_start = i
            if l == 'Passenger Information' and pi_start is None:
                pi_start = i

        if fi_start is not None:
            end = pi_start if pi_start else min(fi_start + 80, len(lines))
            print(f"[Flight Itinerary] lines {fi_start}-{end}:")
            for j in range(fi_start, end):
                print(f"  {j:>4}: {lines[j]!r}")
            print()
        else:
            print("[Flight Itinerary] NOT FOUND")
            for i, l in enumerate(lines):
                if 'flight itinerary' in l.lower():
                    print(f"  similar at {i}: {l!r}")
            print()

        if pi_start is not None:
            print(f"[Passenger Information] lines {pi_start}-{min(pi_start+40, len(lines))}:")
            for j in range(pi_start, min(pi_start + 40, len(lines))):
                print(f"  {j:>4}: {lines[j]!r}")
            print()
        else:
            print("[Passenger Information] NOT FOUND")
            print()

        pnr = _extract_pnr(lines)
        print(f"PNR            : {pnr!r}")

        segments = _extract_segments(lines)
        print(f"\nSegments found : {len(segments)}")
        for s in segments:
            print(f"  {_fmt_segment(s)}")

        inbound, outbound, mtl_arr, mtl_dep = _classify_segments(segments)
        print(f"\nOutbound (to MTL) : {len(inbound)} leg(s)")
        for s in inbound:
            print(f"  {_fmt_segment(s)}")
        print(f"\nReturn (from MTL) : {len(outbound)} leg(s)")
        for s in outbound:
            print(f"  {_fmt_segment(s)}")
        print(f"\nMontreal arrival time   : {mtl_arr!r}")
        print(f"Montreal departure time : {mtl_dep!r}")

        passengers = _extract_passengers(lines)
        print(f"\nPassengers found : {len(passengers)}")
        for p in passengers:
            print(f"  {p['name']:<40} Aeroplan: {p['aeroplan']!r}")

        product, credits = _extract_credit_info(lines, len(passengers))
        print(f"\nFlight Pass product  : {product!r}")
        print(f"Credits per passenger: {credits!r}")

    elif email_type == "paidTickets":
        from parse_paid_tickets import (
            extract_paid_pnr, extract_paid_segments,
            extract_paid_segments_code_line_format,
            extract_paid_segments_marker_format,
            extract_paid_passenger_names, extract_paid_ticket_cost,
            extract_montreal_times_paid, extract_trip_segment_groups_paid,
        )
        pnr = extract_paid_pnr(body, subject)
        print(f"PNR : {pnr!r}")

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
        print(f"\nOutbound:\n  {outbound!r}")
        print(f"Return:\n  {ret!r}")

        mtl_arr, mtl_dep = extract_montreal_times_paid(body)
        print(f"\nMontreal arrival time   : {mtl_arr!r}")
        print(f"Montreal departure time : {mtl_dep!r}")

        names = extract_paid_passenger_names(body)
        print(f"\nPassenger names ({len(names)}):")
        for n in names:
            print(f"  {n!r}")

        cost = extract_paid_ticket_cost(body)
        print(f"\nCost : {cost!r}")

    else:
        print(f"[UNKNOWN type — showing all lines above for manual inspection]")

except Exception:
    import traceback
    traceback.print_exc()

input("\nPress Enter to close...")
