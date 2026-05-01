@echo off
where py >nul 2>&1
if %errorlevel% == 0 (
    py "%~dp0actions.py" open
) else (
    python "%~dp0actions.py" open
)
