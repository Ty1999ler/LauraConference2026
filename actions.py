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
