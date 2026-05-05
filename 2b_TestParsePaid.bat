@echo off
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0test_parse_paid.py"
) else (
    python "%~dp0test_parse_paid.py"
)
pause
