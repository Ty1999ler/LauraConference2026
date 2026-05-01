import os
import sys
import win32com.client
import config

VBA_MODULE_NAME = "AlumoPython"

VBA_CODE = """\
Sub OpenEmail()
    RunPython "import actions; actions.open_email()"
End Sub

Sub PreviewForward()
    RunPython "import actions; actions.preview_forward()"
End Sub
"""


def setup_workbook(excel_path: str):
    # Attach to existing Excel if open, otherwise launch a new hidden instance
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

        # .xlsx cannot hold macros — save as .xlsm first
        if excel_path.lower().endswith(".xlsx"):
            xlsm_path = excel_path[:-5] + ".xlsm"
            print(f"Converting to .xlsm: {os.path.basename(xlsm_path)}")
            wb.SaveAs(os.path.abspath(xlsm_path), FileFormat=52)  # 52 = xlsm
            wb.Close()
            wb          = excel.Workbooks.Open(os.path.abspath(xlsm_path))
            excel_path  = xlsm_path
            print(f"Update EXCEL_FILE in config.py to point to the new .xlsm file.")

        # Inject VBA module
        try:
            vbp = wb.VBProject
        except Exception:
            print()
            print("ERROR: Excel blocked access to the VBA project.")
            print("Enable this once in Excel:")
            print("  File → Options → Trust Center → Trust Center Settings")
            print("  → Macro Settings → check 'Trust access to the VBA project object model'")
            wb.Close(False)
            return

        # Remove old version of module if it exists
        for comp in vbp.VBComponents:
            if comp.Name == VBA_MODULE_NAME:
                vbp.VBComponents.Remove(comp)
                break

        module = vbp.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
        module.Name = VBA_MODULE_NAME
        module.CodeModule.AddFromString(VBA_CODE)

        # Add buttons to PassengerData sheet
        ws = wb.Worksheets(config.SHEET_PASSENGER)

        # Remove old buttons if re-running setup
        for shape in ws.Shapes:
            if shape.Name in ("btnOpenEmail", "btnPreviewForward"):
                shape.Delete()

        # Button positions: just above the header row area, top-right of sheet
        btn1          = ws.Buttons().Add(5, 5, 110, 22)
        btn1.Name     = "btnOpenEmail"
        btn1.Caption  = "Open Email"
        btn1.OnAction = "OpenEmail"

        btn2          = ws.Buttons().Add(125, 5, 140, 22)
        btn2.Name     = "btnPreviewForward"
        btn2.Caption  = "Preview Forward"
        btn2.OnAction = "PreviewForward"

        wb.Save()
        wb.Close()
        print("Done — buttons added to PassengerData sheet.")

    except Exception as exc:
        print(f"ERROR: {exc}")
    finally:
        # Only quit Excel if we launched it — never close someone else's session
        if we_launched:
            excel.Quit()


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else config.EXCEL_FILE
    setup_workbook(path)
