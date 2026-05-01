"""
Quick sanity-check: paste an email body below and run this script.
Usage:  py test_parse.py
"""
import json
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

lines = _normalise(BODY)

# ── quick line dump around key markers ───────────────────────────────────────
for marker in ('Flight Itinerary', 'Passenger Information'):
    for i, l in enumerate(lines):
        if l == marker:
            print(f"[{marker}] found at line {i}")
            for j in range(i, min(i + 20, len(lines))):
                print(f"  {j:>4}: {lines[j]!r}")
            print()
            break
    else:
        print(f"[{marker}] NOT FOUND")
        # show lines that contain it as a substring
        for i, l in enumerate(lines):
            if marker.lower() in l.lower():
                print(f"  similar at {i}: {l!r}")
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
