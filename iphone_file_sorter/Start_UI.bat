@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PYEXE="

REM 1) python on PATH
where python >nul 2>nul
if not errorlevel 1 (
  for /f "delims=" %%I in ('where python') do (
    echo %%I | find /i "WindowsApps" >nul
    if errorlevel 1 (
      set "PYEXE=%%I"
      goto :found_python
    )
  )
)

REM 2) Windows py launcher
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -c "import sys; print(sys.executable)" > "%TEMP%\iphone_sorter_py.txt" 2>nul
  if not errorlevel 1 (
    set /p PYEXE=<"%TEMP%\iphone_sorter_py.txt"
    if defined PYEXE goto :found_python
  )
)

REM 3) Common Anaconda / Miniconda locations
for %%P in (
  "%USERPROFILE%\anaconda3\python.exe"
  "%USERPROFILE%\Anaconda3\python.exe"
  "%USERPROFILE%\miniconda3\python.exe"
  "%USERPROFILE%\Miniconda3\python.exe"
  "%LOCALAPPDATA%\anaconda3\python.exe"
  "%LOCALAPPDATA%\Anaconda3\python.exe"
  "%LOCALAPPDATA%\miniconda3\python.exe"
  "%LOCALAPPDATA%\Miniconda3\python.exe"
  "C:\ProgramData\anaconda3\python.exe"
  "C:\ProgramData\Anaconda3\python.exe"
  "C:\ProgramData\miniconda3\python.exe"
  "C:\ProgramData\Miniconda3\python.exe"
) do (
  if exist %%~P (
    set "PYEXE=%%~P"
    goto :found_python
  )
)

REM 4) conda on PATH -> base python
where conda >nul 2>nul
if not errorlevel 1 (
  for /f "delims=" %%I in ('conda info --base 2^>nul') do (
    if exist "%%I\python.exe" (
      set "PYEXE=%%I\python.exe"
      goto :found_python
    )
  )
)

echo.
echo Could not find Python/Anaconda on PATH.
echo.
echo Do this instead:
echo   1. Open "Anaconda Prompt" from the Start menu
echo   2. Run:
echo      cd /d "%CD%"
echo      python -m pip install -r requirements.txt
echo      python -m streamlit run app.py
echo.
echo Or tell me where Anaconda is installed, e.g.:
echo   where python
echo   where conda
echo.
pause
exit /b 1

:found_python
echo Using Python: %PYEXE%
"%PYEXE%" --version
if errorlevel 1 (
  echo Found Python path but it failed to run: %PYEXE%
  pause
  exit /b 1
)

echo.
echo Installing / updating UI dependencies...
echo (Includes pymobiledevice3 for real iPhone file copy via Apple AFC)
"%PYEXE%" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Failed to install dependencies.
  pause
  exit /b 1
)

echo.
echo Starting iPhone File Sorter UI...
echo A browser window should open automatically.
echo If not, open: http://localhost:8501
echo Close this window to stop the app.
echo.
"%PYEXE%" -m streamlit run app.py
echo.
pause
