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

# Column layout — 1-indexed (A=1 … P=16)
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
COL_EMAIL_STATUS    = 14  # N
COL_DETAILS_SHEET   = 15  # O — which details sheet this row was sent to
COL_MATCH_STATUS    = 16  # P — "Student", "Staff", or "Error"

HEADERS = [
    "EntryID", "PNR", "PassengerName", "AeroplanNumber",
    "FirstDepartureAirport", "OutboundSegments", "ReturnSegments",
    "MontrealArrivalTime", "MontrealDepartureTime",
    "FlightPassProduct", "CreditsPerPassenger", "Cost", "Type",
    "EmailStatus", "DetailsSheet", "MatchStatus"
]

# Sheet names
SHEET_BUTTONS         = "Buttons"
SHEET_STUDENT         = "Student"
SHEET_STAFF           = "Staff"
SHEET_STUDENT_DETAILS = "Student Plane Details"
SHEET_STAFF_DETAILS   = "Staff Plane Details"
SHEET_ERROR           = "Error"

# Default row height for non-wrapped rows
DEFAULT_ROW_HEIGHT = 15

# Preview cap for email opener
MAX_PREVIEW_EMAILS = 10

import os as _os

# Machine-specific settings keyed by Windows username (os.getenv('USERNAME'))
_MACHINE = _os.getenv("USERNAME", "").lower()

_CONFIGS = {
    "atp2txw": {
        "EXCEL_FILE":  r"C:\Users\atp2txw\OneDrive - ATPCO\Documents\Laura\Alumo Summit - Master - Copy.xlsm",
        "FOLDER_PATH": ["Inbox", "Laura"],
    },
    "lmcale": {
        "EXCEL_FILE":  r"C:\Users\lmcale\OneDrive - aseq.com\Desktop\Conferences\Alumo Conference 2026\Alumo Summit - Master - Test.xlsx",
        "FOLDER_PATH": ["Inbox", "Alumo Summit 2026", "AC Flight Pass"],
    },
}

# Fall back to the Laura entry if the machine isn't recognised
_cfg = _CONFIGS.get(_MACHINE, _CONFIGS["lmcale"])

EXCEL_FILE  = _cfg["EXCEL_FILE"]
FOLDER_PATH = _cfg["FOLDER_PATH"]
