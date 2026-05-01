@echo off
echo Installing required packages...
where py >nul 2>&1
if %errorlevel% == 0 (
    py -m pip install pywin32 openpyxl xlwings fpdf2
    py -m pywin32_postinstall -install
    py -m xlwings addin install
    py "%~dp0generate_howto_pdf.py"
) else (
    python -m pip install pywin32 openpyxl xlwings fpdf2
    python -m pywin32_postinstall -install
    python -m xlwings addin install
    python "%~dp0generate_howto_pdf.py"
)
echo.
echo Installation complete. You can close this window.
pause
