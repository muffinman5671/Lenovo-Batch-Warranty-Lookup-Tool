@echo off
REM Double-click launcher for the Lenovo Batch Warranty Lookup Tool.
REM Starts the GUI without a console window hanging around behind it.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0LenovoWarrantyLookup.ps1"
