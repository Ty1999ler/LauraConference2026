"""
Batch email preview: opens Outlook forward drafts for unpreviewd passengers.

Order: Staff rows first, then Student rows.
Cap: 10 per run (config.MAX_PREVIEW_EMAILS).
Forward-to address: TEST_FORWARD_EMAIL (testing) — update when ready for real sends.

Never calls .Send() — always opens as a draft for human review.

Uses openpyxl read-only for all reads, win32com for cell updates and save
so openpyxl never writes the file (prevents xlsm data corruption).
"""
import os
import sys
import traceback
import openpyxl
import win32com.client
import pythoncom

import config

TEST_FORWARD_EMAIL = "tigerrock1999@gmail.com"

# Column positions in the Plane Details sheets (1-indexed)
_DETAILS_COL_PREFERRED_NAME = 2   # B
_DETAILS_COL_AEROPLAN       = 7   # G


def _get_outlook():
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def _open_forward_draft(namespace, entry_id: str, preferred_name: str, to_address: str):
    """Open a forward draft addressed to to_address. Never calls .Send()."""
    item = namespace.GetItemFromID(entry_id)
    fwd  = item.Forward()
    fwd.To = to_address
    greeting     = f"<p>Hi{(' ' + preferred_name) if preferred_name else ''},</p><br>"
    fwd.HTMLBody = greeting + fwd.HTMLBody
    fwd.Display()  # preview only — NEVER .Send()


def _build_preferred_name_map(wb_ro) -> dict:
    """
    Returns {aeroplan_str: preferred_name} from both Plane Details sheets.
    Looks up col G (AeroplanNumber) → col B (Preferred Name).
    """
    mapping = {}
    for sheet_name in [config.SHEET_STUDENT_DETAILS, config.SHEET_STAFF_DETAILS]:
        if sheet_name not in wb_ro.sheetnames:
            continue
        ws = wb_ro[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < _DETAILS_COL_AEROPLAN:
                continue
            aeroplan = row[_DETAILS_COL_AEROPLAN - 1]
            pref     = row[_DETAILS_COL_PREFERRED_NAME - 1]
            if aeroplan:
                mapping[str(aeroplan).replace(' ', '')] = str(pref).strip() if pref else ''
    return mapping


def run_preview(excel_path: str):
    pythoncom.CoInitialize()

    abs_path = os.path.abspath(excel_path)

    # Read everything with openpyxl read-only — never writes the file
    print("Reading workbook...")
    wb_ro = openpyxl.load_workbook(abs_path, read_only=True, data_only=True)

    if config.SHEET_PASSENGER not in wb_ro.sheetnames:
        print("PassengerData sheet not found — nothing to do.")
        wb_ro.close()
        return

    preferred_name_map = _build_preferred_name_map(wb_ro)

    ws_ro = wb_ro[config.SHEET_PASSENGER]
    staff_rows   = []
    student_rows = []

    for row in ws_ro.iter_rows(min_row=2, values_only=True):
        if len(row) < config.COL_MATCH_STATUS:
            continue

        entry_id     = row[config.COL_ENTRY_ID     - 1]
        email_status = row[config.COL_EMAIL_STATUS  - 1]
        match_status = row[config.COL_MATCH_STATUS  - 1]
        name         = row[config.COL_PASSENGER_NAME - 1] or ""
        aeroplan     = row[config.COL_AEROPLAN       - 1]

        if not entry_id:
            continue
        if email_status:
            continue

        aeroplan_str   = str(aeroplan).replace(' ', '') if aeroplan else ''
        preferred_name = preferred_name_map.get(aeroplan_str, '') or str(name)

        record = (str(entry_id), preferred_name)
        if match_status == "Staff":
            staff_rows.append(record)
        elif match_status == "Student":
            student_rows.append(record)

    wb_ro.close()

    to_preview = staff_rows + student_rows

    if not to_preview:
        print("All passengers already previewed — nothing to do.")
        return

    total     = len(to_preview)
    cap       = config.MAX_PREVIEW_EMAILS
    batch     = to_preview[:cap]
    remaining = total - len(batch)

    print(f"Found {total} unpreviewd passenger(s) — opening {len(batch)} draft(s)...")
    print(f"  ({len(staff_rows)} Staff, {len(student_rows)} Student pending)")
    print()

    outlook   = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")

    # Track which entry_ids succeeded so we can update EmailStatus via COM
    previewed_ids = []
    errored       = []

    for entry_id, preferred_name in batch:
        try:
            _open_forward_draft(namespace, entry_id, preferred_name, TEST_FORWARD_EMAIL)
            previewed_ids.append(entry_id)
            print(f"  [OK ] {preferred_name or entry_id[:12]}")
        except Exception as exc:
            errored.append((entry_id, str(exc)))
            print(f"  [ERR] {preferred_name or entry_id[:12]} — {exc}")

    # Update EmailStatus via win32com — never touches other sheets
    print()
    print("Saving EmailStatus updates via Excel...")
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

        ws_com = wb_com.Sheets(config.SHEET_PASSENGER)
        last_row = ws_com.Cells(ws_com.Rows.Count, config.COL_ENTRY_ID).End(-4162).Row

        previewed_set = set(previewed_ids)
        errored_map   = {eid: msg for eid, msg in errored}

        for row_num in range(2, last_row + 1):
            eid = ws_com.Cells(row_num, config.COL_ENTRY_ID).Value
            if not eid:
                continue
            eid_str = str(eid)
            if eid_str in previewed_set:
                ws_com.Cells(row_num, config.COL_EMAIL_STATUS).Value = "Previewed"
            elif eid_str in errored_map:
                ws_com.Cells(row_num, config.COL_EMAIL_STATUS).Value = f"Error: {errored_map[eid_str]}"

        wb_com.Save()
        xl.Visible = True

    except Exception as exc:
        print(f"  [WARNING] Could not save EmailStatus — {exc}")
        print("  Open the workbook manually and re-run to retry.")

    print(f"Done. Opened {len(previewed_ids)} draft(s)." +
          (f" {len(errored)} error(s) — check EmailStatus column." if errored else ""))
    if remaining:
        print(f"{remaining} more unpreviewd — run again for next batch of {cap}.")


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else config.EXCEL_FILE
    try:
        run_preview(path)
    except Exception:
        traceback.print_exc()
    input("\nPress Enter to close...")
