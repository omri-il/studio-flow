@echo off
cd /d "%~dp0"
echo Installing dependencies...
pip install -r requirements.txt --quiet
echo.
echo Starting Mic Volume Tracker...
python tracker.py
pause
