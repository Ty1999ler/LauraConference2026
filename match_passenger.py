import config


def _find_col_by_header(ws, header_text: str) -> int:
    """Return 1-based column index of the first header cell containing header_text."""
    for cell in ws[1]:
        if cell.value and header_text.lower() in str(cell.value).lower():
            return cell.column
    return -1


def _build_header_map(ws) -> dict:
    """Return {col_index: header_name} for row 1."""
    return {
        cell.column: str(cell.value)
        for cell in ws[1]
        if cell.value
    }


def lookup_aeroplan(wb, aeroplan_number) -> tuple:
    """
    Search Student sheet then Staff sheet for aeroplan_number.
    Returns (source_sheet_name, registration_data_dict) or (None, None).
    """
    if not aeroplan_number or not str(aeroplan_number).strip():
        return None, None

    aeroplan_str = str(aeroplan_number).strip()

    for sheet_name in [config.SHEET_STUDENT, config.SHEET_STAFF]:
        if sheet_name not in wb.sheetnames:
            continue

        ws           = wb[sheet_name]
        aeroplan_col = _find_col_by_header(ws, "Aeroplan")

        if aeroplan_col == -1:
            continue

        headers = _build_header_map(ws)

        for row in ws.iter_rows(min_row=2):
            cell_val = row[aeroplan_col - 1].value
            if cell_val and str(cell_val).strip() == aeroplan_str:
                reg_data = {
                    headers[cell.column]: cell.value
                    for cell in row
                    if cell.column in headers
                }
                return sheet_name, reg_data

    return None, None


def details_sheet_for(source_sheet: str) -> str:
    if source_sheet == config.SHEET_STUDENT:
        return config.SHEET_STUDENT_DETAILS
    if source_sheet == config.SHEET_STAFF:
        return config.SHEET_STAFF_DETAILS
    return config.SHEET_ERROR
