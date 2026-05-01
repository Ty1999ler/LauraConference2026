@echo off
echo Installing required packages...
where py >nul 2>&1
if %errorlevel% == 0 (
    py -m pip install pywin32 openpyxl xlwings
    py -m pywin32_postinstall -install
) else (
    python -m pip install pywin32 openpyxl xlwings
    python -m pywin32_postinstall -install
)
echo.
echo Installation complete. You can close this window.
pause
