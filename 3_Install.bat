@echo off
echo Installing required packages...

where py >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON=py
) else (
    set PYTHON=python
)

%PYTHON% -m pip install pywin32 openpyxl xlwings fpdf2

echo.
echo Running pywin32 post-install (may not apply on all systems)...
%PYTHON% -m pywin32_postinstall -install 2>nul
if %errorlevel% neq 0 echo   Skipped - not needed on this Python installation.

echo.
echo Installing xlwings Excel addin (may not apply on all systems)...
%PYTHON% -m xlwings addin install 2>nul
if %errorlevel% neq 0 echo   Skipped - not needed on this Python installation.

echo.
echo Generating PDF how-to guide...
%PYTHON% "%~dp0generate_howto_pdf.py"

echo.
echo Installation complete. You can close this window.
pause
