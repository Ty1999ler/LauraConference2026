@echo off
echo This will clear all rows from the following worksheets:
echo   - PassengerData
echo   - Student Plane Details
echo   - Staff Plane Details
echo   - Error
echo Headers will be kept. This cannot be undone.
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
