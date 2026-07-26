@echo off
setlocal EnableExtensions

:: Restart Apple Mobile Device Service (needs Administrator).
:: Right-click this file -> Run as administrator, or just double-click
:: (it will re-launch itself elevated).

net session >nul 2>&1
if errorlevel 1 (
  echo Requesting Administrator permission...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)

echo ============================================================
echo  Restarting Apple Mobile Device Service
echo ============================================================
echo.

echo Stopping service...
net stop "Apple Mobile Device Service"
echo.
echo Starting service...
net start "Apple Mobile Device Service"
if errorlevel 1 (
  echo.
  echo Failed to start the service.
  echo Install/update "Apple Devices" from the Microsoft Store ^(or iTunes^), then try again.
  echo.
  pause
  exit /b 1
)

echo.
echo Done. Next steps:
echo   1. Unlock iPhone, reconnect USB, tap Trust
echo   2. In Anaconda Prompt:
echo        conda activate iphone_copy
echo        pymobiledevice3 usbmux list
echo   3. If the phone is listed, run Start_UI_Py312.bat
echo.
pause
exit /b 0
