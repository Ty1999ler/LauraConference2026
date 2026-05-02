@echo off
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0preview_emails.py"
) else (
    python "%~dp0preview_emails.py"
)
pause
