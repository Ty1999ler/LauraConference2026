@echo off
echo Setting up Excel workbook...
echo.
echo IMPORTANT - if this fails, do this once in Excel first:
echo   File ^> Options ^> Trust Center ^> Trust Center Settings
echo   ^> Macro Settings ^> check "Trust access to the VBA project object model"
echo.
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0setup_workbook.py"
) else (
    python "%~dp0setup_workbook.py"
)
pause
