@echo off
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0update_flight_info.py"
) else (
    python "%~dp0update_flight_info.py"
)
pause
