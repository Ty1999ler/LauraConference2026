@echo off
echo Checking for updates...
cd /d "%~dp0"
git pull
echo.
echo Update complete. You can close this window.
pause
