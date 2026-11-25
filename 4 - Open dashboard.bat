@echo off
rem Starts the local dashboard (http://localhost:8765) and opens it in your browser.
rem Keep this window open while you use the dashboard; close it to stop.
cd /d "%~dp0"
python tools\dashboard\server.py
pause
