@echo off
echo Installing required packages...
py -m pip install pywin32 openpyxl xlwings
py -m pywin32_postinstall -install
echo.
echo Installation complete. You can close this window.
pause
