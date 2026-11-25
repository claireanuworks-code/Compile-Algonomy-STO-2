@echo off
cd /d "%~dp0"
rem Install any missing packages (only runs pip if something is missing)
python -c "import pandas, numpy, python_calamine, xlsxwriter" 2>nul || (
  echo Installing missing Python packages...
  python -m pip install pandas numpy python-calamine xlsxwriter
)
echo.
python sto_allocation.py
echo.
pause
