@echo off
cd /d "%~dp0"
echo Installing required Python packages...
python -m pip install --upgrade -r requirements.txt
echo.
echo Setup finished.
pause
