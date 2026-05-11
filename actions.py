import sys
import xlwings as xw
import win32com.client
import pythoncom
import config

TEST_FORWARD_EMAIL = "tigerrock1999@gmail.com"


def _get_outlook():
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def _get_selected_row_data():
    """Read EntryID and passenger name from the selected row in the active Excel workbook."""
    try:
        app = xw.apps.active
    except Exception:
        raise RuntimeError("Excel is not open.")

    wb  = app.books.active
    ws  = wb.sheets[config.SHEET_PASSENGER]
    row = app.selection.row

    if row < 2:
        raise ValueError("Click a passenger data row first (not the header).")

    entry_id       = ws.cells(row, config.COL_ENTRY_ID).value
    passenger_name = ws.cells(row, config.COL_PASSENGER_NAME).value or ""

    if not entry_id:
        raise ValueError("No EntryID in the selected row.")

    return wb, ws, row, str(entry_id), passenger_name


def _open_forward_draft(namespace, entry_id: str, passenger_name: str):
    """Build and display a forward draft. Never calls .Send()."""
    item         = namespace.GetItemFromID(entry_id)
    fwd          = item.Forward()
    fwd.To       = TEST_FORWARD_EMAIL
    greeting     = "<p>Hi,</p><br>"
    fwd.HTMLBody = greeting + fwd.HTMLBody
    fwd.Display()  # NEVER .Send() — always preview only


# ---------------------------------------------------------------------------
# Desktop shortcut actions (standalone — read active Excel selection)
# ---------------------------------------------------------------------------

def open_email():
    try:
        _, _, _, entry_id, _ = _get_selected_row_data()
        outlook   = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
        item      = namespace.GetItemFromID(entry_id)
        item.Display()
    except Exception as exc:
        print(f"Error: {exc}")
        input("Press Enter to close...")


def preview_forward():
    try:
        wb, ws, row, entry_id, passenger_name = _get_selected_row_data()
        outlook   = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
        _open_forward_draft(namespace, entry_id, passenger_name)
        ws.cells(row, config.COL_EMAIL_STATUS).value = "Previewed"
        wb.save()
    except Exception as exc:
        print(f"Error: {exc}")
        input("Press Enter to close...")


# ---------------------------------------------------------------------------
# Buttons sheet action (called from Excel via xlwings RunPython)
# ---------------------------------------------------------------------------

def preview_all_unsent():
    """Open forward drafts for every PassengerData row without an EmailStatus value."""
    wb = xw.Book.caller()
    ws = wb.sheets[config.SHEET_PASSENGER]

    # Find last row with data in column A
    last_row = ws.range(
        f"A{ws.api.Rows.Count}"
    ).end("up").row

    if last_row < 2:
        wb.app.alert("No passenger data found.")
        return

    # Collect rows that haven't been previewed yet
    to_preview = []
    for row in range(2, last_row + 1):
        entry_id = ws.cells(row, config.COL_ENTRY_ID).value
        status   = ws.cells(row, config.COL_EMAIL_STATUS).value
        if entry_id and not status:
            name = ws.cells(row, config.COL_PASSENGER_NAME).value or ""
            to_preview.append((row, str(entry_id), name))

    if not to_preview:
        wb.app.alert("All emails have already been previewed.")
        return

    total   = len(to_preview)
    cap     = config.MAX_PREVIEW_EMAILS
    batch   = to_preview[:cap]
    remaining = total - len(batch)

    outlook   = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    opened    = 0
    errors    = 0

    for row, entry_id, passenger_name in batch:
        try:
            _open_forward_draft(namespace, entry_id, passenger_name)
            ws.cells(row, config.COL_EMAIL_STATUS).value = "Previewed"
            opened += 1
        except Exception as exc:
            ws.cells(row, config.COL_EMAIL_STATUS).value = f"Error: {exc}"
            errors += 1

    wb.save()

    msg = f"Opened {opened} preview(s)."
    if errors:
        msg += f"\n{errors} error(s) — check EmailStatus column."
    if remaining:
        msg += f"\n{remaining} more unsent — run again for next batch of {cap}."
    wb.app.alert(msg)


# ---------------------------------------------------------------------------
# Buttons sheet action — check forwards only
# ---------------------------------------------------------------------------

def check_forwards():
    """Scan Sent Items for Alumo Summit forwards and mark Previewed rows as Sent."""
    import openpyxl
    pythoncom.CoInitialize()

    wb_xw    = xw.Book.caller()
    abs_path = wb_xw.fullname

    from preview_emails import _build_details_maps, _find_sent_entry_ids, _update_details_sheet

    wb_ro = openpyxl.load_workbook(abs_path, read_only=True, data_only=True)
    preferred_name_map, email_map = _build_details_maps(wb_ro)

    ws_ro          = wb_ro[config.SHEET_PASSENGER]
    previewed_rows = []

    for row in ws_ro.iter_rows(min_row=2, values_only=True):
        if len(row) < config.COL_MATCH_STATUS:
            continue
        entry_id     = row[config.COL_ENTRY_ID     - 1]
        email_status = row[config.COL_EMAIL_STATUS  - 1]
        match_status = row[config.COL_MATCH_STATUS  - 1]
        aeroplan     = row[config.COL_AEROPLAN      - 1]
        name         = str(row[config.COL_PASSENGER_NAME - 1] or "")

        if not entry_id or email_status != "Previewed":
            continue
        if str(match_status or '') not in ("Staff", "Student"):
            continue

        if aeroplan:
            ap = int(aeroplan) if isinstance(aeroplan, float) else aeroplan
            aeroplan_str = str(ap).replace(' ', '')
        else:
            aeroplan_str = ''

        to_email = email_map.get(aeroplan_str, '')
        if not to_email:
            continue

        preferred_name = preferred_name_map.get(aeroplan_str, '') or name
        previewed_rows.append((str(entry_id), preferred_name, aeroplan_str,
                               str(match_status), to_email))

    wb_ro.close()

    if not previewed_rows:
        wb_xw.app.alert("No Previewed rows to check.")
        return

    outlook        = _get_outlook()
    namespace      = outlook.GetNamespace("MAPI")
    newly_sent_pairs = _find_sent_entry_ids(namespace, previewed_rows)

    if not newly_sent_pairs:
        wb_xw.app.alert("No newly sent forwards found.")
        return

    sent_map = {(r[0], r[2]): r for r in previewed_rows if (r[0], r[2]) in newly_sent_pairs}

    try:
        xl     = win32com.client.GetActiveObject("Excel.Application")
        wb_com = None
        for w in xl.Workbooks:
            if w.FullName.lower() == abs_path.lower():
                wb_com = w
                break
        if wb_com is None:
            wb_com = xl.Workbooks.Open(abs_path)

        ws_com   = wb_com.Sheets(config.SHEET_PASSENGER)
        last_row = ws_com.Cells(ws_com.Rows.Count, config.COL_ENTRY_ID).End(-4162).Row

        for row_num in range(2, last_row + 1):
            eid = ws_com.Cells(row_num, config.COL_ENTRY_ID).Value
            if not eid:
                continue
            ap_val = ws_com.Cells(row_num, config.COL_AEROPLAN).Value
            if isinstance(ap_val, float):
                ap_str = str(int(ap_val))
            else:
                ap_str = str(ap_val or '').replace(' ', '')
            key = (str(eid), ap_str)
            if key not in sent_map:
                continue
            ws_com.Cells(row_num, config.COL_EMAIL_STATUS).Value = "Sent"
            _, _, aeroplan_str, match_status, _ = sent_map[key]
            details = (config.SHEET_STAFF_DETAILS if match_status == "Staff"
                       else config.SHEET_STUDENT_DETAILS)
            if aeroplan_str:
                _update_details_sheet(wb_com, details, aeroplan_str, "Sent")

        wb_com.Save()

    except Exception as exc:
        wb_xw.app.alert(f"Error saving updates: {exc}")
        return

    wb_xw.app.alert(f"Marked {len(newly_sent_pairs)} row(s) as Sent.")


# ---------------------------------------------------------------------------
# Entry point for desktop shortcuts
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "open":
        open_email()
    elif cmd == "forward":
        preview_forward()
    else:
        print("Usage: python actions.py open|forward")
        input("Press Enter to close...")
