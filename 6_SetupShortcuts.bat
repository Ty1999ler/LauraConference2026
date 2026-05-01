@echo off
echo Creating desktop shortcuts...
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0setup_shortcuts.py"
) else (
    python "%~dp0setup_shortcuts.py"
)
