import os
import sys
import win32com.client
import config

VBA_MODULE_NAME = "AlumoPython"

VBA_CODE = """\
Sub PreviewAllUnsent()
    RunPython "import actions; actions.preview_all_unsent()"
End Sub
"""


def setup_workbook(excel_path: str):
    try:
        excel        = win32com.client.GetActiveObject("Excel.Application")
        we_launched  = False
    except Exception:
        excel        = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        we_launched  = True

    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))

        # .xlsx cannot hold macros — save as .xlsm
        if excel_path.lower().endswith(".xlsx"):
            xlsm_path = excel_path[:-5] + ".xlsm"
            print(f"Converting to .xlsm: {os.path.basename(xlsm_path)}")
            wb.SaveAs(os.path.abspath(xlsm_path), FileFormat=52)
            wb.Close()
            wb         = excel.Workbooks.Open(os.path.abspath(xlsm_path))
            excel_path = xlsm_path
            print("Update EXCEL_FILE in config.py to use the .xlsm filename.")

        # Inject VBA module
        try:
            vbp = wb.VBProject
        except Exception:
            print()
            print("ERROR: Excel blocked access to the VBA project.")
            print("Enable this once in Excel:")
            print("  File > Options > Trust Center > Trust Center Settings")
            print("  > Macro Settings > check 'Trust access to the VBA project object model'")
            wb.Close(False)
            return

        for comp in vbp.VBComponents:
            if comp.Name == VBA_MODULE_NAME:
                vbp.VBComponents.Remove(comp)
                break

        module = vbp.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        module.Name = VBA_MODULE_NAME
        module.CodeModule.AddFromString(VBA_CODE)

        # Add or clear the Buttons sheet
        sheet_names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
        if config.SHEET_BUTTONS in sheet_names:
            btn_ws = wb.Worksheets(config.SHEET_BUTTONS)
            for shape in btn_ws.Shapes:
                shape.Delete()
        else:
            btn_ws = wb.Worksheets.Add()
            btn_ws.Name = config.SHEET_BUTTONS

        # Clean background
        btn_ws.Cells.Interior.Color = 0xF2F2F2

        # Title label
        title_cell = btn_ws.Cells(2, 2)
        title_cell.Value = "Email Actions"
        title_cell.Font.Bold  = True
        title_cell.Font.Size  = 16
        title_cell.Font.Color = 0x333333

        # "Preview All Unsent" button — large, centered
        btn = btn_ws.Buttons().Add(80, 60, 200, 40)
        btn.Name     = "btnPreviewAllUnsent"
        btn.Caption  = "Preview All Unsent Emails"
        btn.OnAction = "PreviewAllUnsent"
        btn.Font.Size = 11

        wb.Save()
        wb.Close()
        print("Done — Buttons sheet added to workbook.")

    except Exception as exc:
        print(f"ERROR: {exc}")
    finally:
        if we_launched:
            excel.Quit()


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else config.EXCEL_FILE
    setup_workbook(path)
