"""
Batch email preview: opens Outlook forward drafts for unpreviewd passengers.

Order: Staff rows first, then Student rows.
Cap: 10 per run (config.MAX_PREVIEW_EMAILS).
Forward-to address: TEST_FORWARD_EMAIL (testing) — swap for real passenger emails later.

Never calls .Send() — always opens as a draft for human review.

On each run, scans Outlook Sent Items first: any Previewed row whose original
email's ConversationID matches a sent forward → automatically marked "Sent".

Uses openpyxl read-only for all reads, win32com for cell updates and save
so openpyxl never writes the file (prevents xlsm data corruption).
"""
import os
import sys
import traceback
from datetime import datetime, timedelta
import openpyxl
import win32com.client
import pythoncom

import config

TEST_FORWARD_EMAIL = "tigerrock1999@gmail.com"

# Column positions in the Plane Details sheets (1-indexed, matches DETAILS_HEADERS)
_DETAILS_COL_PREFERRED_NAME = 2   # B — Preferred Name
_DETAILS_COL_EMAIL          = 3   # C — Email
_DETAILS_COL_AEROPLAN       = 7   # G — AeroplanNumber
_DETAILS_COL_EMAIL_STATUS   = 18  # R — EmailStatus


def _get_outlook():
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def _open_forward_draft(namespace, entry_id: str, preferred_name: str, to_address: str):
    """Open a forward draft addressed to to_address. Never calls .Send()."""
    import re as _re
    item   = namespace.GetItemFromID(entry_id)
    fwd    = item.Forward()
    fwd.To = to_address

    name_part = preferred_name if preferred_name else ""
    intro = (
        f"<p>Hi {name_part},</p>"
        f"<p>I'm very excited to welcome you to the inaugural Alumo Summit!</p>"
        f"<p>Please find your travel booking below!</p>"
        f"<p>I'll be sharing additional information about the conference in June so stay tuned! This will include:</p>"
        f"<ul><li>Summit agenda</li><li>Accommodation details</li><li>Shuttle schedule</li>"
        f"<li>Meal options</li><li>App details</li><li>And much more!!</li></ul>"
        f"<p>In the meantime, if you have any questions, please don't hesitate to reach out.</p>"
        f"<p>The Alumo team is excited to welcome you to Tremblant this July!</p>"
        f"<p>Looking forward to seeing you soon,</p>"
    )

    # Inject intro right after the opening <body> tag so we don't stack HTML documents
    html  = fwd.HTMLBody
    match = _re.search(r'<body[^>]*>', html, _re.IGNORECASE)
    if match:
        pos          = match.end()
        fwd.HTMLBody = html[:pos] + intro + html[pos:]
    else:
        fwd.HTMLBody = intro + html

    fwd.Subject = "Alumo Summit – Travel Booking"
    fwd.Display()  # preview only — NEVER .Send()


def _build_details_maps(wb_ro) -> tuple:
    """
    Returns ({aeroplan_str: preferred_name}, {aeroplan_str: email})
    from both Plane Details sheets.
    """
    names  = {}
    emails = {}
    for sheet_name in [config.SHEET_STUDENT_DETAILS, config.SHEET_STAFF_DETAILS]:
        if sheet_name not in wb_ro.sheetnames:
            continue
        ws = wb_ro[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < _DETAILS_COL_AEROPLAN:
                continue
            aeroplan = row[_DETAILS_COL_AEROPLAN - 1]
            if not aeroplan:
                continue
            key = str(aeroplan).replace(' ', '')
            pref  = row[_DETAILS_COL_PREFERRED_NAME - 1]
            email = row[_DETAILS_COL_EMAIL - 1]
            names[key]  = str(pref).strip()  if pref  else ''
            emails[key] = str(email).strip() if email else ''
    return names, emails


def _find_sent_entry_ids(namespace, previewed_entry_ids: list, forward_addresses: set) -> set:
    """
    Returns the subset of previewed_entry_ids whose forward has been sent.
    Matches via ConversationID within the last SENT_SCAN_DAYS days.
    forward_addresses: set of lowercase email addresses to match against.
    """
    if not previewed_entry_ids:
        return set()

    print(f"  Scanning Sent Items (last {config.SENT_SCAN_DAYS} days) for completed forwards...")

    cutoff     = datetime.now() - timedelta(days=config.SENT_SCAN_DAYS)
    cutoff_str = cutoff.strftime("%m/%d/%Y %H:%M %p")

    sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail
    restricted  = sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff_str}'")

    sent_conv_ids = set()
    for item in restricted:
        try:
            to_field = (item.To or '').lower()
            if any(addr in to_field for addr in forward_addresses):
                sent_conv_ids.add(item.ConversationID)
        except Exception:
            continue

    if not sent_conv_ids:
        print("  No matching sent forwards found.")
        return set()

    matched = set()
    for entry_id in previewed_entry_ids:
        try:
            original = namespace.GetItemFromID(entry_id)
            if original.ConversationID in sent_conv_ids:
                matched.add(entry_id)
        except Exception:
            continue

    print(f"  Found {len(matched)} sent forward(s).")
    return matched


def _update_details_sheet(wb_com, sheet_name: str, aeroplan_str: str, status: str):
    """Set EmailStatus (col R) in the Plane Details sheet for the matching Aeroplan row."""
    try:
        ws = wb_com.Sheets(sheet_name)
    except Exception:
        return
    last_row = ws.Cells(ws.Rows.Count, _DETAILS_COL_AEROPLAN).End(-4162).Row  # xlUp
    for row_num in range(2, last_row + 1):
        cell_val = ws.Cells(row_num, _DETAILS_COL_AEROPLAN).Value
        if not cell_val:
            continue
        # COM returns numbers as floats — normalise to plain digit string
        if isinstance(cell_val, float):
            cell_str = str(int(cell_val))
        else:
            cell_str = str(cell_val).replace(' ', '')
        if cell_str == aeroplan_str:
            ws.Cells(row_num, _DETAILS_COL_EMAIL_STATUS).Value = status
            return


def run_preview(excel_path: str):
    pythoncom.CoInitialize()

    abs_path = os.path.abspath(excel_path)

    # ── Read workbook (read-only — never writes) ──────────────────────────
    print("Reading workbook...")
    wb_ro = openpyxl.load_workbook(abs_path, read_only=True, data_only=True)

    if config.SHEET_PASSENGER not in wb_ro.sheetnames:
        print("PassengerData sheet not found — nothing to do.")
        wb_ro.close()
        return

    preferred_name_map, email_map = _build_details_maps(wb_ro)

    ws_ro        = wb_ro[config.SHEET_PASSENGER]
    staff_rows   = []   # unpreviewd
    student_rows = []   # unpreviewd
    previewed_rows = [] # EmailStatus == "Previewed" → check if now sent

    for row in ws_ro.iter_rows(min_row=2, values_only=True):
        if len(row) < config.COL_MATCH_STATUS:
            continue

        entry_id     = row[config.COL_ENTRY_ID      - 1]
        email_status = row[config.COL_EMAIL_STATUS   - 1]
        match_status = row[config.COL_MATCH_STATUS   - 1]
        name         = row[config.COL_PASSENGER_NAME - 1] or ""
        aeroplan     = row[config.COL_AEROPLAN       - 1]

        if not entry_id:
            continue

        aeroplan_str   = str(aeroplan).replace(' ', '') if aeroplan else ''
        preferred_name = preferred_name_map.get(aeroplan_str, '') or str(name)
        to_email       = email_map.get(aeroplan_str, '') or TEST_FORWARD_EMAIL
        match_str      = str(match_status or '')
        # record: (entry_id, preferred_name, aeroplan_str, match_status, to_email)
        record = (str(entry_id), preferred_name, aeroplan_str, match_str, to_email)

        if email_status == "Previewed":
            previewed_rows.append(record)
        elif not email_status:
            if match_status == "Staff":
                staff_rows.append(record)
            elif match_status == "Student":
                student_rows.append(record)

    wb_ro.close()

    # ── Connect to Outlook ────────────────────────────────────────────────
    outlook   = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")

    # ── Step 1: Check Sent Items for previously previewed rows ────────────
    previewed_entry_ids = [r[0] for r in previewed_rows]
    # Collect all forward addresses used so the scan covers all of them
    forward_addresses   = {r[4].lower() for r in previewed_rows if r[4]}
    forward_addresses.add(TEST_FORWARD_EMAIL.lower())
    newly_sent_ids      = _find_sent_entry_ids(namespace, previewed_entry_ids,
                                               forward_addresses)
    newly_sent_map      = {r[0]: r for r in previewed_rows if r[0] in newly_sent_ids}

    if newly_sent_ids:
        print(f"  Marking {len(newly_sent_ids)} row(s) as Sent.")
    print()

    # ── Step 2: Open forward drafts for unpreviewd rows ───────────────────
    to_preview = staff_rows + student_rows

    new_previewed = []  # (entry_id, aeroplan_str, match_status)
    new_errored   = []  # (entry_id, aeroplan_str, match_status, error_msg)

    if not to_preview:
        if not newly_sent_ids:
            print("Nothing to do — all passengers already previewed or sent.")
    else:
        total     = len(to_preview)
        cap       = config.MAX_PREVIEW_EMAILS
        batch     = to_preview[:cap]
        remaining = total - len(batch)

        print(f"Found {total} unpreviewd passenger(s) — opening {len(batch)} draft(s)...")
        print(f"  ({len(staff_rows)} Staff, {len(student_rows)} Student pending)")
        print()

        for entry_id, preferred_name, aeroplan_str, match_status, to_email in batch:
            try:
                _open_forward_draft(namespace, entry_id, preferred_name, to_email)
                new_previewed.append((entry_id, aeroplan_str, match_status))
                print(f"  [OK ] {preferred_name or entry_id[:12]} → {to_email}")
            except Exception as exc:
                new_errored.append((entry_id, aeroplan_str, match_status, str(exc)))
                print(f"  [ERR] {preferred_name or entry_id[:12]} — {exc}")

        if remaining:
            print()
            print(f"{remaining} more unpreviewd — run again for next batch of {cap}.")

    # ── Step 3: Write all status updates via win32com ─────────────────────
    if not newly_sent_ids and not new_previewed and not new_errored:
        return

    print()
    print("Saving updates via Excel...")
    try:
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            xl = win32com.client.Dispatch("Excel.Application")

        xl.Visible = False
        wb_com = None
        for w in xl.Workbooks:
            if w.FullName.lower() == abs_path.lower():
                wb_com = w
                break
        if wb_com is None:
            wb_com = xl.Workbooks.Open(abs_path)

        ws_com   = wb_com.Sheets(config.SHEET_PASSENGER)
        last_row = ws_com.Cells(ws_com.Rows.Count, config.COL_ENTRY_ID).End(-4162).Row

        sent_map      = newly_sent_map                                        # entry_id → record
        previewed_map = {eid: (ap, ms) for eid, ap, ms in new_previewed}
        errored_map   = {eid: (ap, ms, msg) for eid, ap, ms, msg in new_errored}

        for row_num in range(2, last_row + 1):
            eid = ws_com.Cells(row_num, config.COL_ENTRY_ID).Value
            if not eid:
                continue
            eid_str = str(eid)

            if eid_str in sent_map:
                _, _, aeroplan_str, match_status = sent_map[eid_str]
                ws_com.Cells(row_num, config.COL_EMAIL_STATUS).Value = "Sent"
                details = (config.SHEET_STAFF_DETAILS if match_status == "Staff"
                           else config.SHEET_STUDENT_DETAILS)
                if aeroplan_str:
                    _update_details_sheet(wb_com, details, aeroplan_str, "Sent")

            elif eid_str in previewed_map:
                aeroplan_str, match_status = previewed_map[eid_str]
                ws_com.Cells(row_num, config.COL_EMAIL_STATUS).Value = "Previewed"
                details = (config.SHEET_STAFF_DETAILS if match_status == "Staff"
                           else config.SHEET_STUDENT_DETAILS)
                if aeroplan_str:
                    _update_details_sheet(wb_com, details, aeroplan_str, "Previewed")

            elif eid_str in errored_map:
                _, _, msg = errored_map[eid_str]
                ws_com.Cells(row_num, config.COL_EMAIL_STATUS).Value = f"Error: {msg}"

        wb_com.Save()
        xl.Visible = True
        print("Saved.")

    except Exception as exc:
        print(f"  [WARNING] Could not save updates — {exc}")
        print("  Open the workbook manually and re-run to retry.")

    print()
    summary = []
    if newly_sent_ids:
        summary.append(f"{len(newly_sent_ids)} marked Sent")
    if new_previewed:
        summary.append(f"{len(new_previewed)} draft(s) opened")
    if new_errored:
        summary.append(f"{len(new_errored)} error(s)")
    print("Done — " + ", ".join(summary) + ".")


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else config.EXCEL_FILE
    try:
        run_preview(path)
    except Exception:
        traceback.print_exc()
    input("\nPress Enter to close...")
