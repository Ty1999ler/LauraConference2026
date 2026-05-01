"""
Batch email preview: opens Outlook forward drafts for unpreviewd passengers.

Order: Staff rows first, then Student rows.
Cap: 10 per run (config.MAX_PREVIEW_EMAILS).
Forward-to address: TEST_FORWARD_EMAIL (testing) — update when ready for real sends.

Never calls .Send() — always opens as a draft for human review.
"""
import os
import sys
import traceback
import openpyxl
import win32com.client
import pythoncom

import config

TEST_FORWARD_EMAIL = "tigerrock1999@gmail.com"


def _close_workbook_if_open(excel_path: str):
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return
    target = os.path.abspath(excel_path).lower()
    for wb in excel.Workbooks:
        if wb.FullName.lower() == target:
            wb.Save()
            wb.Close(False)
            print("Closed workbook before run.")
            return


def _get_outlook():
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def _open_forward_draft(namespace, entry_id: str, passenger_name: str, to_address: str):
    """Open a forward draft addressed to to_address. Never calls .Send()."""
    item = namespace.GetItemFromID(entry_id)
    fwd = item.Forward()
    fwd.To = to_address
    greeting = f"<p>Hi{(' ' + passenger_name) if passenger_name else ''},</p><br>"
    fwd.HTMLBody = greeting + fwd.HTMLBody
    fwd.Display()  # preview only — NEVER .Send()


def run_preview(excel_path: str):
    _close_workbook_if_open(excel_path)

    print("Opening workbook...")
    wb = openpyxl.load_workbook(excel_path)

    if config.SHEET_PASSENGER not in wb.sheetnames:
        print("PassengerData sheet not found — nothing to do.")
        wb.close()
        return

    ws = wb[config.SHEET_PASSENGER]

    # Collect unpreviewd rows grouped by match status
    staff_rows   = []
    student_rows = []

    for row_num in range(2, ws.max_row + 1):
        entry_id     = ws.cell(row=row_num, column=config.COL_ENTRY_ID).value
        email_status = ws.cell(row=row_num, column=config.COL_EMAIL_STATUS).value
        match_status = ws.cell(row=row_num, column=config.COL_MATCH_STATUS).value

        if not entry_id:
            continue
        if email_status:  # already previewed or errored
            continue

        name = ws.cell(row=row_num, column=config.COL_PASSENGER_NAME).value or ""

        if match_status == "Staff":
            staff_rows.append((row_num, str(entry_id), name))
        elif match_status == "Student":
            student_rows.append((row_num, str(entry_id), name))
        # Skip "Error" rows — no valid email to forward

    # Staff first, then Student
    to_preview = staff_rows + student_rows

    if not to_preview:
        print("All passengers already previewed — nothing to do.")
        wb.save(excel_path)
        os.startfile(excel_path)
        return

    total   = len(to_preview)
    cap     = config.MAX_PREVIEW_EMAILS
    batch   = to_preview[:cap]
    remaining = total - len(batch)

    print(f"Found {total} unpreviewd passenger(s) — opening {len(batch)} draft(s)...")
    print(f"  ({len(staff_rows)} Staff, {len(student_rows)} Student pending)")
    print()

    outlook   = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")

    opened = 0
    errors = 0

    for row_num, entry_id, name in batch:
        try:
            _open_forward_draft(namespace, entry_id, name, TEST_FORWARD_EMAIL)
            ws.cell(row=row_num, column=config.COL_EMAIL_STATUS).value = "Previewed"
            print(f"  [OK ] {name or entry_id[:12]}")
            opened += 1
        except Exception as exc:
            ws.cell(row=row_num, column=config.COL_EMAIL_STATUS).value = f"Error: {exc}"
            print(f"  [ERR] {name or entry_id[:12]} — {exc}")
            errors += 1

    print()
    wb.save(excel_path)
    print(f"Saved. Opened {opened} draft(s)." +
          (f" {errors} error(s) — check EmailStatus column." if errors else ""))
    if remaining:
        print(f"{remaining} more unpreviewd — run again for next batch of {cap}.")

    os.startfile(excel_path)


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else config.EXCEL_FILE
    try:
        run_preview(path)
    except Exception:
        traceback.print_exc()
    input("\nPress Enter to close...")
