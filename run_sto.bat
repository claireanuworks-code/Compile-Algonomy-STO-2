@echo off
cd /d "%~dp0"
rem Install any missing packages (only runs pip if something is missing)
python -c "import pandas, numpy, python_calamine, xlsxwriter, win32com.client" 2>nul || (
  echo Installing missing Python packages...
  python -m pip install pandas numpy python-calamine xlsxwriter pywin32
)
echo.
python sto_allocation.py
echo.
pause
