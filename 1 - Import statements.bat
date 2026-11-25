@echo off
rem Reads every new PDF in the "statements" folder into the Import Review sheet.
cd /d "%~dp0"
python tools\statement_import.py import
echo.
pause
