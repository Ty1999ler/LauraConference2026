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


def open_email():
    """Open the original email for the selected row in Outlook."""
    wb = xw.Book.caller()
    ws = wb.sheets[config.SHEET_PASSENGER]
    row = wb.app.selection.row

    if row < 2:
        wb.app.alert("Click a passenger data row first (not the header).")
        return

    entry_id = ws.cells(row, config.COL_ENTRY_ID).value
    if not entry_id:
        wb.app.alert("No EntryID found in the selected row.")
        return

    try:
        outlook   = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
        item      = namespace.GetItemFromID(str(entry_id))
        item.Display()
    except Exception as exc:
        wb.app.alert(f"Could not open email:\n{exc}")


def preview_forward():
    """
    Open a forward draft for the selected row.
    NEVER calls .Send() — always previews for manual review.
    """
    wb = xw.Book.caller()
    ws = wb.sheets[config.SHEET_PASSENGER]
    row = wb.app.selection.row

    if row < 2:
        wb.app.alert("Click a passenger data row first (not the header).")
        return

    entry_id       = ws.cells(row, config.COL_ENTRY_ID).value
    passenger_name = ws.cells(row, config.COL_PASSENGER_NAME).value or ""

    if not entry_id:
        wb.app.alert("No EntryID found in the selected row.")
        return

    try:
        outlook   = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
        item      = namespace.GetItemFromID(str(entry_id))

        fwd    = item.Forward()
        fwd.To = TEST_FORWARD_EMAIL

        greeting     = "<p>Hi,</p><br>"
        fwd.HTMLBody = greeting + fwd.HTMLBody

        fwd.Display()  # NEVER .Send() — always preview only

    except Exception as exc:
        wb.app.alert(f"Could not open forward preview:\n{exc}")
