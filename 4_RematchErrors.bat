@echo off
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0rematch_errors.py"
) else (
    python "%~dp0rematch_errors.py"
)
pause
