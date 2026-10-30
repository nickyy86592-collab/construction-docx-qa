@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-ConstructionRunnerUpgrade.ps1"
echo.
pause
