# Alumo Conference 2026 — How To Guide

---

## What This Does

This tool reads Air Canada booking confirmation emails from Outlook and automatically:
- Extracts passenger flight details (PNR, segments, Montreal times, etc.)
- Matches each passenger to their registration record by Aeroplan number
- Writes matched passengers to **Student Plane Details** or **Staff Plane Details**
- Flags anyone not found in either sheet into an **Error** sheet
- Lets you preview and send individual or batch forwarded emails

---

## First Time Setup (Do This Once Per Computer)

### Step 1 — Install Python
1. Go to **python.org/downloads**
2. Download Python 3.13
3. Run the installer — on the first screen check **"Add Python to PATH"**
4. Click Install

### Step 2 — Install Git
1. Go to **git-scm.com/download/win**
2. Run the installer — all default options are fine

### Step 3 — Get the project
1. Open File Explorer, navigate to where you want the folder (e.g. Desktop)
2. Right-click in an empty area → **Open in Terminal**
3. Type this and press Enter:
```
git clone https://github.com/Ty1999ler/LauraConference2026.git
```
4. A folder called `LauraConference2026` will appear

### Step 4 — Install packages
1. Open the `LauraConference2026` folder
2. Double-click **`install.bat`**
3. Wait for it to finish — it will say "Installation complete"

### Step 5 — Allow Excel to work with Python
1. Open Excel
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Click **Macro Settings**
4. Check **"Trust access to the VBA project object model"**
5. Click OK and close Options

### Step 6 — Set up the Excel workbook
1. Double-click **`setup_workbook.bat`**
2. This adds the **Buttons** sheet with the Preview All Unsent button
3. Only needs to be run once (or again if the workbook is replaced)

### Step 7 — Create desktop shortcuts
1. Double-click **`setup_shortcuts.bat`**
2. Two shortcuts appear on your Desktop:
   - **Open Email** — opens the original Outlook email for a selected row
   - **Preview Forward** — opens a forward draft for a selected row
3. Keyboard shortcuts are also set up: **Ctrl+Alt+O** and **Ctrl+Alt+F**

---

## Every Day — Running the Tool

### Step 1 — Get latest updates (optional but recommended)
Double-click **`update.bat`**
This pulls any fixes or improvements that have been pushed.

### Step 2 — Run the import
1. Make sure Excel is **closed**
2. Double-click **`run.bat`**
3. A window will show progress — you will see:
   - Each email being processed
   - Each passenger found, with any missing fields flagged
   - Whether each passenger was matched to Student or Staff
   - A summary at the end
4. Open Excel when it finishes

### What you will see in Excel after running:
| Sheet | What's in it |
|---|---|
| **PassengerData** | Every passenger from every email — the full staging log |
| **Student Plane Details** | Students matched by Aeroplan number |
| **Staff Plane Details** | Staff matched by Aeroplan number |
| **Error** | Anyone whose Aeroplan wasn't found in Student or Staff |
| **Debug** | Emails that caused errors during processing |
| **Buttons** | The Preview All Unsent button |

---

## Sending Emails

### Preview a single passenger's email
1. Click the passenger's row in **Student Plane Details** or **Staff Plane Details**
2. Press **Ctrl+Alt+F** (or double-click the **Preview Forward** shortcut on your Desktop)
3. Outlook opens a draft — review it, then click Send manually
4. The row's **EmailStatus** column updates to "Previewed"

### Open the original booking email
1. Click the passenger's row
2. Press **Ctrl+Alt+O** (or double-click **Open Email** on your Desktop)
3. The original Air Canada email opens in Outlook

### Preview all unsent emails at once
1. Open the Excel workbook
2. Click the **Buttons** sheet (tab at the bottom)
3. Click **Preview All Unsent Emails**
4. Outlook opens up to 10 draft forwards at a time
5. Review and send each one manually
6. Run again for the next batch if there are more than 10

---

## If Something Goes Wrong

| Problem | What to do |
|---|---|
| "Could not find folder: Alumo Summit 2026" | The Outlook folder path is wrong — contact Ty |
| Passenger shows in Error sheet | Their Aeroplan number in the email doesn't match the Student or Staff sheet — check for typos in either place |
| run.bat says the file is locked | Excel is open — close it first, then run again |
| Missing fields in the output | The email format may be unusual — check the Debug sheet for details |
| py is not recognized | Re-run `install.bat` |

---

## Getting Updates

Whenever Ty pushes a fix or new feature:
1. Double-click **`update.bat`**
2. Run **`run.bat`** as normal

No other steps needed — the update is automatic.

---

## File Summary

| File | What it does |
|---|---|
| `run.bat` | Runs the import — use this every day |
| `update.bat` | Pulls latest code from GitHub |
| `install.bat` | Installs Python packages — run once |
| `setup_workbook.bat` | Adds the Buttons sheet to Excel — run once |
| `setup_shortcuts.bat` | Creates desktop shortcuts — run once |
