import sys
import traceback
import openpyxl

import config
from outlook_connector import get_outlook_folder, get_folder_items
from parse_flight_pass import get_email_type, parse_flight_pass_email
from parse_paid_tickets import parse_paid_email
from excel_writer import (
    ensure_headers, write_row, format_passenger_sheet,
    get_all_entry_ids, get_next_row, log_debug
)


# ── Set this to the full path of Laura's workbook ──────────────────────────
EXCEL_FILE = r"C:\path\to\Laura_Workbook.xlsx"
# ───────────────────────────────────────────────────────────────────────────


def run_everything(excel_path: str):
    print("Opening workbook…")
    wb = openpyxl.load_workbook(excel_path)

    # Ensure PassengerData sheet exists
    if config.SHEET_PASSENGER not in wb.sheetnames:
        wb.create_sheet(config.SHEET_PASSENGER)
    ws = wb[config.SHEET_PASSENGER]

    # Ensure Debug sheet exists
    if config.SHEET_DEBUG not in wb.sheetnames:
        debug_ws = wb.create_sheet(config.SHEET_DEBUG)
        debug_ws.append(["EntryID", "Subject", "Error", "RowNum"])
    else:
        debug_ws = wb[config.SHEET_DEBUG]

    ensure_headers(ws)
    processed_ids = get_all_entry_ids(ws)
    print(f"Already processed: {len(processed_ids)} email(s)")

    print("Connecting to Outlook…")
    folder = get_outlook_folder(config.FOLDER_PATH)
    items  = get_folder_items(folder)
    print(f"Found {len(items)} email(s) in folder")

    next_row    = get_next_row(ws)
    added_count = 0

    for mail in items:
        subject  = mail.Subject or ""
        entry_id = mail.EntryID

        if entry_id in processed_ids:
            continue

        email_type = get_email_type(subject)
        if not email_type:
            continue

        try:
            body = mail.Body or ""

            if email_type == "flightPass":
                rows = parse_flight_pass_email(body, entry_id)
            else:
                rows = parse_paid_email(body, entry_id, subject)

            for row_data in rows:
                write_row(ws, next_row, row_data)
                next_row  += 1
                added_count += 1

            processed_ids.add(entry_id)

        except Exception as exc:
            err_msg = traceback.format_exc()
            print(f"  ERROR processing '{subject[:60]}': {exc}")
            log_debug(debug_ws, entry_id, subject, err_msg, next_row)

    print(f"Added {added_count} new row(s). Applying formatting…")
    format_passenger_sheet(ws)

    wb.save(excel_path)
    print("Done — workbook saved.")


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else EXCEL_FILE
    run_everything(path)
