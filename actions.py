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

    return str(entry_id), passenger_name


def open_email():
    try:
        entry_id, _ = _get_selected_row_data()
        outlook     = _get_outlook()
        namespace   = outlook.GetNamespace("MAPI")
        item        = namespace.GetItemFromID(entry_id)
        item.Display()
    except Exception as exc:
        print(f"Error: {exc}")
        input("Press Enter to close...")


def preview_forward():
    try:
        entry_id, passenger_name = _get_selected_row_data()
        outlook   = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
        item      = namespace.GetItemFromID(entry_id)

        fwd          = item.Forward()
        fwd.To       = TEST_FORWARD_EMAIL
        greeting     = "<p>Hi,</p><br>"
        fwd.HTMLBody = greeting + fwd.HTMLBody

        fwd.Display()  # NEVER .Send() — always preview only

    except Exception as exc:
        print(f"Error: {exc}")
        input("Press Enter to close...")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "open":
        open_email()
    elif cmd == "forward":
        preview_forward()
    else:
        print("Usage: python actions.py open|forward")
        input("Press Enter to close...")
