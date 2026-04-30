@echo off
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0main.py"
) else (
    python "%~dp0main.py"
)
pause
