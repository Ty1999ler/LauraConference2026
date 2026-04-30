# Outlook folder path
MAILBOX_EMAIL = "lmcale@aseq.com"
FOLDER_PATH   = ["Inbox", "Alumno Conference", "AC Scan"]
# NOTE: confirm with Laura whether path is Inbox > Alumno Conference > AC Scan
# or whether VBA's olInbox.Folders("Laura") points somewhere different

# Excel sheet names
SHEET_PASSENGER = "PassengerData"
SHEET_DEBUG     = "Debug"

# Email type identifiers matched against subject line
SUBJECT_PAID        = "Numéro de réservation"
SUBJECT_FLIGHT_PASS = "Electronic Ticket Itinerary/Receipt"

# Formatting colours as hex strings for openpyxl
COLOR_HEADER   = "A0A0A0"   # darker grey header row  (RGB 160,160,160)
COLOR_ROW_BAND = "C8C8C8"   # alternating light-grey  (RGB 200,200,200)

# Column layout — 1-indexed (A=1 … M=13)
COL_ENTRY_ID        = 1   # A
COL_PNR             = 2   # B
COL_PASSENGER_NAME  = 3   # C
COL_AEROPLAN        = 4   # D
COL_FIRST_DEP       = 5   # E
COL_OUTBOUND_SEG    = 6   # F
COL_RETURN_SEG      = 7   # G
COL_MTL_ARRIVAL     = 8   # H
COL_MTL_DEPARTURE   = 9   # I
COL_FP_PRODUCT      = 10  # J
COL_CREDITS_PER_PAX = 11  # K
COL_COST            = 12  # L
COL_TYPE            = 13  # M

HEADERS = [
    "EntryID", "PNR", "PassengerName", "AeroplanNumber",
    "FirstDepartureAirport", "OutboundSegments", "ReturnSegments",
    "MontrealArrivalTime", "MontrealDepartureTime",
    "FlightPassProduct", "CreditsPerPassenger", "Cost", "Type"
]

# Default row height for non-wrapped rows
DEFAULT_ROW_HEIGHT = 15

# Preview cap for email opener
MAX_PREVIEW_EMAILS = 10

import os as _os

# Machine-specific settings keyed by Windows username (os.getenv('USERNAME'))
_MACHINE = _os.getenv("USERNAME", "").lower()

_CONFIGS = {
    "atp2txw": {
        "EXCEL_FILE":  r"C:\Users\atp2txw\path\to\test_workbook.xlsx",
        "FOLDER_PATH": ["Inbox", "Alumno Conference", "AC Scan"],
    },
    # TODO: replace LAURAS_USERNAME with her actual Windows username
    # (the name in C:\Users\______ on her PC)
    "lauras_username": {
        "EXCEL_FILE":  r"C:\Users\lauras_username\path\to\Laura_Workbook.xlsx",
        "FOLDER_PATH": ["Inbox", "Alumno Conference", "AC Scan"],
    },
}

# Fall back to the Laura entry if the machine isn't recognised
_cfg = _CONFIGS.get(_MACHINE, _CONFIGS["lauras_username"])

EXCEL_FILE  = _cfg["EXCEL_FILE"]
FOLDER_PATH = _cfg["FOLDER_PATH"]
