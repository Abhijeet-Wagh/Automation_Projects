@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python from https://www.python.org/downloads/
  echo Make sure to check "Add python.exe to PATH" during setup.
  pause
  exit /b 1
)

echo Installing / updating UI dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Failed to install dependencies.
  pause
  exit /b 1
)

echo.
echo Starting iPhone File Sorter UI...
echo A browser window should open automatically.
echo Close this window to stop the app.
echo.
python -m streamlit run app.py
pause
