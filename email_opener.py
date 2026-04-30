import win32com.client
import pythoncom
import config


def _get_outlook():
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def open_email_from_entry_id(entry_id: str):
    """Open a single email in Outlook by its EntryID."""
    try:
        outlook   = _get_outlook()
        namespace = outlook.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(entry_id)
        mail_item.Display()
    except Exception as exc:
        print(f"Email not found for ID: {entry_id}  ({exc})")


def open_email_from_active_cell(workbook_path: str, sheet_name: str):
    """
    Read the EntryID from the active cell in the given sheet (via xlwings)
    and open that email in Outlook.
    """
    try:
        import xlwings as xw
        wb   = xw.Book(workbook_path)
        ws   = wb.sheets[sheet_name]
        cell = ws.range(ws.api.Application.ActiveCell.Address)
        entry_id = cell.value
        if entry_id:
            open_email_from_entry_id(str(entry_id))
        else:
            print("Active cell is empty — no EntryID to open.")
    except ImportError:
        print("xlwings is not installed. Run: pip install xlwings")
    except Exception as exc:
        print(f"Could not open email from active cell: {exc}")


def preview_emails_from_sheet(ws_forward, passenger_ws,
                               max_count: int = config.MAX_PREVIEW_EMAILS,
                               debug_ws=None):
    """
    Loop rows in ws_forward.
    - Skip rows where column U already = "YES"
    - Get recipient from column F
    - Get EntryID from PassengerData at the row number stored in column T
    - Forward the email with a greeting prepended to HTMLBody
    - Mark column U = "Yes"
    - Stop after max_count previews

    Errors are logged to debug_ws instead of raising.
    """
    outlook   = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    previewed = 0

    for row_num in range(2, ws_forward.max_row + 1):
        if previewed >= max_count:
            break

        # Column U = 21
        already_done = ws_forward.cell(row=row_num, column=21).value
        if str(already_done).upper() == "YES":
            continue

        recipient_email = ws_forward.cell(row=row_num, column=6).value  # col F
        pdata_row       = ws_forward.cell(row=row_num, column=20).value  # col T

        if not recipient_email or not pdata_row:
            continue

        try:
            pdata_row = int(pdata_row)
            entry_id  = passenger_ws.cell(row=pdata_row,
                                           column=1).value  # col A = EntryID
            if not entry_id:
                raise ValueError("EntryID is blank at PassengerData row "
                                  f"{pdata_row}")

            mail_item = namespace.GetItemFromID(str(entry_id))
            fwd       = mail_item.Forward()
            fwd.To    = recipient_email

            passenger_name = passenger_ws.cell(row=pdata_row, column=3).value or ""
            greeting = (
                f"<p>Bonjour {passenger_name},</p>"
                f"<p>Veuillez trouver ci-joint votre confirmation de voyage.</p>"
                f"<br>"
            )
            fwd.HTMLBody = greeting + fwd.HTMLBody
            fwd.Display()   # show for review — do NOT call .Send() automatically

            ws_forward.cell(row=row_num, column=21).value = "Yes"
            previewed += 1

        except Exception as exc:
            msg = f"Row {row_num}: {exc}"
            print(f"Preview error — {msg}")
            if debug_ws is not None:
                debug_ws.append(["preview_error", str(recipient_email), msg, row_num])
