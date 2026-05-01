import os
import sys
import win32com.client

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DESKTOP    = os.path.join(os.path.expanduser("~"), "Desktop")

SHORTCUTS = [
    {
        "name":        "1_Open Email",
        "bat":         os.path.join(SCRIPT_DIR, "1_OpenEmail.bat"),
        "hotkey":      "Ctrl+Alt+O",
        "description": "Open the Outlook email for the selected Excel row",
    },
    {
        "name":        "2_Preview Forward",
        "bat":         os.path.join(SCRIPT_DIR, "2_PreviewForward.bat"),
        "hotkey":      "Ctrl+Alt+F",
        "description": "Preview a forward of the email for the selected Excel row",
    },
]


def create_shortcuts():
    shell = win32com.client.Dispatch("WScript.Shell")

    for s in SHORTCUTS:
        path             = os.path.join(DESKTOP, f"{s['name']}.lnk")
        shortcut         = shell.CreateShortcut(path)
        shortcut.TargetPath      = s["bat"]
        shortcut.WorkingDirectory = SCRIPT_DIR
        shortcut.Description     = s["description"]
        shortcut.Hotkey          = s["hotkey"]
        shortcut.WindowStyle     = 7   # 7 = minimised — no cmd window flash
        shortcut.Save()
        print(f"Created: {s['name']}.lnk  (Hotkey: {s['hotkey']})")

    print()
    print("Shortcuts created on your Desktop.")
    print("Hotkeys are active from anywhere once the shortcuts exist on the Desktop.")


if __name__ == "__main__":
    create_shortcuts()
    input("Press Enter to close...")
