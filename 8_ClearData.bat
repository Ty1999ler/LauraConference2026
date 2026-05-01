@echo off
echo This will clear ALL passenger data from PassengerData, Student Plane Details, Staff Plane Details, and Error sheets.
echo.
set /p confirm=Type YES to confirm:
if /i not "%confirm%"=="YES" (
    echo Cancelled.
    pause
    exit /b
)

where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0clear_data.py"
) else (
    python "%~dp0clear_data.py"
)
pause
