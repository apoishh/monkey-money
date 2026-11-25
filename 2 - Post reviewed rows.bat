@echo off
rem Moves rows marked Post? = Yes from Import Review into Budget Tracking.
cd /d "%~dp0"
python tools\statement_import.py post
echo.
pause
