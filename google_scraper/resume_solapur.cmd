@echo off
REM Double-click to resume the Solapur taluka queue.
REM Skips talukas that already have a finished output file.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_solapur_remaining.ps1"
pause
