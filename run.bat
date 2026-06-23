@echo off
title POB Collection Analytics (Flask - Port 8060)
cd /d "%~dp0"

echo Installing dependencies...
py -m pip install -r requirements.txt -q

echo.
echo  POB Collection Dashboard running at: http://localhost:8060
echo  Press Ctrl+C to stop.
echo.
start "" http://localhost:8060
py app.py
pause
