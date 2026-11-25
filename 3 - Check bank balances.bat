@echo off
rem Compares the tracker's Balance / Savings with the latest Maybank / CIMB statement balances.
cd /d "%~dp0"
python tools\statement_import.py check
echo.
pause
