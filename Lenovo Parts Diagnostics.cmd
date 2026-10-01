@echo off
REM Double-click launcher for the part lookup diagnostic.
REM Asks for a serial, runs Lookup-Part.ps1 -Diagnose, and leaves the report
REM next to this file as "Lenovo parts diagnostics <date>.txt".
setlocal
set "SERIAL=%~1"
if "%SERIAL%"=="" set /p SERIAL=Serial number: 
if "%SERIAL%"=="" goto :eof
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Lookup-Part.ps1" "%SERIAL%" -Diagnose
echo.
pause
